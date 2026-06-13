"""
External integrations for the EULER Deal Desk app — all stdlib, no pip installs.

  * Microsoft Entra ID (Azure AD) SSO  — OpenID Connect Authorization-Code flow + PKCE
  * Salesforce                          — Username-Password OAuth -> insert custom object
  * Microsoft 365 / Graph               — app-only Mail.Send (reuses the Entra app)

Everything is driven by environment variables. If a block is not configured the
corresponding function returns {"ok": False, "skipped": True} and the app keeps
working (SQLite stays the source of truth). Nothing here blocks a submission.
"""
import os, json, base64, hashlib, secrets, urllib.parse, urllib.request, urllib.error

# ----------------------------------------------------------------- env
def _e(k, d=""): return os.environ.get(k, d).strip()

# Azure / Entra (SSO + Graph share one app registration by default)
AZ_TENANT   = _e("AZURE_TENANT_ID")
AZ_CLIENT   = _e("AZURE_CLIENT_ID")
AZ_SECRET   = _e("AZURE_CLIENT_SECRET")
AZ_REDIRECT = _e("AZURE_REDIRECT_URI", "http://localhost:4610/api/auth/callback")
AZ_GROUP    = _e("AZURE_ALLOWED_GROUP")          # optional: restrict to one security group (object id)

# Salesforce (username-password OAuth)
SF_LOGIN    = _e("SF_LOGIN_URL", "https://login.salesforce.com")   # sandbox: https://test.salesforce.com
SF_CLIENT   = _e("SF_CLIENT_ID")                 # connected-app consumer key
SF_SECRET   = _e("SF_CLIENT_SECRET")             # connected-app consumer secret
SF_USER     = _e("SF_USERNAME")
SF_PASS     = _e("SF_PASSWORD")                  # password + security token concatenated
SF_OBJECT   = _e("SF_OBJECT", "Deal_Registration__c")
SF_APIVER   = _e("SF_API_VERSION", "v60.0")
# Optional override: {"<SF field API name>": "<payload key>"} — defaults below otherwise.
SF_FIELD_MAP = _e("SF_FIELD_MAP")

# Microsoft Graph email (defaults to the same Entra app as SSO)
GR_TENANT   = _e("GRAPH_TENANT_ID")  or AZ_TENANT
GR_CLIENT   = _e("GRAPH_CLIENT_ID")  or AZ_CLIENT
GR_SECRET   = _e("GRAPH_CLIENT_SECRET") or AZ_SECRET
MAIL_SENDER = _e("MAIL_SENDER")                  # the from-mailbox UPN, e.g. dealdesk@hcltech.com
MAIL_TO     = [a.strip() for a in _e("MAIL_TO").split(",") if a.strip()]


def config_status():
    return {
        "sso_mode": "azure" if (AZ_TENANT and AZ_CLIENT and AZ_SECRET) else "mock",
        "salesforce": bool(SF_CLIENT and SF_USER and SF_PASS),
        "email": bool(GR_TENANT and GR_CLIENT and GR_SECRET and MAIL_SENDER and MAIL_TO),
    }

# ----------------------------------------------------------------- tiny HTTP
def _post_form(url, fields, headers=None):
    data = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(url, data=data, method="POST",
                                 headers={"Content-Type": "application/x-www-form-urlencoded", **(headers or {})})
    with urllib.request.urlopen(req, timeout=12) as r:
        return json.loads(r.read() or b"{}")

def _post_json(url, obj, headers=None):
    data = json.dumps(obj).encode()
    req = urllib.request.Request(url, data=data, method="POST",
                                 headers={"Content-Type": "application/json", **(headers or {})})
    with urllib.request.urlopen(req, timeout=12) as r:
        body = r.read()
        return r.status, (json.loads(body) if body else {})

# ================================================================= AZURE SSO
def azure_authorize_url(state, nonce, verifier):
    """Build the Microsoft authorize redirect URL (Auth-Code flow + PKCE)."""
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    q = {
        "client_id": AZ_CLIENT, "response_type": "code", "redirect_uri": AZ_REDIRECT,
        "response_mode": "query", "scope": "openid profile email User.Read",
        "state": state, "nonce": nonce,
        "code_challenge": challenge, "code_challenge_method": "S256",
    }
    return f"https://login.microsoftonline.com/{AZ_TENANT}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(q)

def azure_exchange_code(code, verifier):
    """Exchange the auth code for tokens; return validated user claims."""
    tok = _post_form(f"https://login.microsoftonline.com/{AZ_TENANT}/oauth2/v2.0/token", {
        "client_id": AZ_CLIENT, "client_secret": AZ_SECRET, "grant_type": "authorization_code",
        "code": code, "redirect_uri": AZ_REDIRECT, "code_verifier": verifier,
        "scope": "openid profile email User.Read",
    })
    idt = tok.get("id_token", "")
    claims = _decode_jwt_payload(idt) if idt else {}
    email = claims.get("preferred_username") or claims.get("email") or ""
    name = claims.get("name") or (email.split("@")[0].replace(".", " ").title() if email else "User")
    return {"email": email.lower(), "name": name, "oid": claims.get("oid"),
            "tenant": claims.get("tid"), "access_token": tok.get("access_token")}

def _decode_jwt_payload(jwt):
    # Token comes straight from Microsoft's token endpoint over TLS (server-to-server),
    # so we read claims directly. For defence-in-depth, verify the signature against the
    # tenant JWKS (https://login.microsoftonline.com/<tid>/discovery/v2.0/keys) in prod.
    try:
        p = jwt.split(".")[1]
        p += "=" * (-len(p) % 4)
        return json.loads(base64.urlsafe_b64decode(p))
    except Exception:
        return {}

# ================================================================= SALESFORCE
def _sf_default_map(p, project_code):
    svcs = p.get("services") or []
    return {
        "Name": (p.get("dealName") or "Deal Registration")[:80],
        "Project_Code__c": project_code,
        "Customer_Company__c": p.get("company"),
        "Customer_Website__c": p.get("website"),
        "Customer_Business_Unit__c": p.get("businessUnit"),
        "TCV__c": float(p.get("tcv") or 0),
        "Currency_Code__c": p.get("currency"),
        "Status__c": p.get("status"),
        "Delivery_Model__c": p.get("deliveryModel"),
        "Services__c": "; ".join(svcs),
        "Primary_Use_Case__c": p.get("customerUseCase"),
        "POC_Name__c": p.get("pocName"),
        "POC_Email__c": p.get("pocEmail"),
        "POC_Title__c": p.get("pocTitle"),
        "POC_Phone__c": p.get("pocPhone"),
        "Delivery_Lead_Name__c": p.get("leadName"),
        "Delivery_Lead_Email__c": p.get("leadEmail"),
        "Start_Date__c": p.get("startDate") or None,
        "End_Date__c": p.get("endDate") or None,
        "Linked_Deal_Id__c": p.get("linkedDeal"),
        "Registered_By__c": p.get("createdBy"),
        "Description__c": p.get("projectDescription") or p.get("description"),
        "Rollout_Plan__c": p.get("rolloutPlan"),
    }

def salesforce_create(payload, project_code):
    if not (SF_CLIENT and SF_USER and SF_PASS):
        return {"ok": False, "skipped": True, "reason": "Salesforce not configured"}
    try:
        tok = _post_form(f"{SF_LOGIN}/services/oauth2/token", {
            "grant_type": "password", "client_id": SF_CLIENT, "client_secret": SF_SECRET,
            "username": SF_USER, "password": SF_PASS,
        })
        access, instance = tok.get("access_token"), tok.get("instance_url")
        if not access:
            return {"ok": False, "error": "SF auth failed", "detail": tok}
        fields = _sf_default_map(payload, project_code)
        if SF_FIELD_MAP:  # optional remap: {"<SF field>": "<payload key>"}
            try:
                remap = json.loads(SF_FIELD_MAP)
                fields = {sf: payload.get(pk) for sf, pk in remap.items()}
                fields.setdefault("Name", (payload.get("dealName") or "Deal Registration")[:80])
                fields["Project_Code__c"] = project_code
            except Exception:
                pass
        fields = {k: v for k, v in fields.items() if v not in (None, "")}
        status, body = _post_json(
            f"{instance}/services/data/{SF_APIVER}/sobjects/{SF_OBJECT}/", fields,
            headers={"Authorization": f"Bearer {access}"})
        if status in (200, 201) and body.get("id"):
            return {"ok": True, "id": body["id"], "object": SF_OBJECT}
        return {"ok": False, "error": "SF insert failed", "detail": body}
    except urllib.error.HTTPError as e:
        return {"ok": False, "error": f"SF HTTP {e.code}", "detail": e.read().decode(errors="ignore")[:400]}
    except Exception as e:
        return {"ok": False, "error": str(e)}

# ================================================================= GRAPH EMAIL
def _graph_token():
    tok = _post_form(f"https://login.microsoftonline.com/{GR_TENANT}/oauth2/v2.0/token", {
        "client_id": GR_CLIENT, "client_secret": GR_SECRET,
        "grant_type": "client_credentials", "scope": "https://graph.microsoft.com/.default",
    })
    return tok.get("access_token")

def _mail_html(p, project_code, sf_result):
    money = f"{p.get('currency','USD')} {float(p.get('tcv') or 0):,.0f}"
    svcs = "".join(f"<li>{s}</li>" for s in (p.get("services") or [])) or "<li>—</li>"
    sf = (f"Salesforce {SF_OBJECT}: <b>{sf_result.get('id')}</b>" if sf_result.get("ok")
          else "Salesforce: not synced")
    return f"""<div style="font-family:Segoe UI,Arial,sans-serif;color:#191915">
      <h2 style="font-family:Georgia,serif">New Deal Desk Registration</h2>
      <p><b>{p.get('dealName','—')}</b> &nbsp;·&nbsp; <code>{project_code}</code></p>
      <table cellpadding="6" style="border-collapse:collapse;font-size:14px">
        <tr><td style="color:#6E6E68">Customer</td><td><b>{p.get('company','—')}</b> ({p.get('website','—')})</td></tr>
        <tr><td style="color:#6E6E68">TCV</td><td><b>{money}</b></td></tr>
        <tr><td style="color:#6E6E68">Status</td><td>{p.get('status','—')} · {p.get('deliveryModel','—')}</td></tr>
        <tr><td style="color:#6E6E68">Primary use case</td><td>{p.get('customerUseCase','—')}</td></tr>
        <tr><td style="color:#6E6E68">Customer POC</td><td>{p.get('pocName','—')} — {p.get('pocEmail','—')}</td></tr>
        <tr><td style="color:#6E6E68">Delivery lead</td><td>{p.get('leadName','—')} — {p.get('leadEmail','—')}</td></tr>
        <tr><td style="color:#6E6E68">Registered by</td><td>{p.get('createdBy','—')}</td></tr>
      </table>
      <p style="color:#6E6E68">Services:</p><ul>{svcs}</ul>
      <p style="color:#888;font-size:12px">{sf}<br/>Sent by EULER · Claude Partner Network Deal Desk</p>
    </div>"""

def send_email(payload, project_code, sf_result=None):
    if not (GR_TENANT and GR_CLIENT and GR_SECRET and MAIL_SENDER and MAIL_TO):
        return {"ok": False, "skipped": True, "reason": "Email (Graph) not configured"}
    try:
        access = _graph_token()
        if not access:
            return {"ok": False, "error": "Graph auth failed"}
        msg = {
            "message": {
                "subject": f"New Deal Desk Registration — {payload.get('dealName','')} ({project_code})",
                "body": {"contentType": "HTML", "content": _mail_html(payload, project_code, sf_result or {})},
                "toRecipients": [{"emailAddress": {"address": a}} for a in MAIL_TO],
            },
            "saveToSentItems": True,
        }
        status, _ = _post_json(f"https://graph.microsoft.com/v1.0/users/{MAIL_SENDER}/sendMail",
                               msg, headers={"Authorization": f"Bearer {access}"})
        return {"ok": status in (200, 202), "status": status}
    except urllib.error.HTTPError as e:
        return {"ok": False, "error": f"Graph HTTP {e.code}", "detail": e.read().decode(errors="ignore")[:400]}
    except Exception as e:
        return {"ok": False, "error": str(e)}
