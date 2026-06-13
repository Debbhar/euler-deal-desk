#!/usr/bin/env python3
"""
EULER - Claude Partner Network · Deal Desk
Zero-dependency full-stack app: Python stdlib HTTP server + embedded SQLite.

  python3 server.py            # serves on http://localhost:4610
  PORT=5000 python3 server.py

Optional integrations (set as env vars):
  DEALDESK_WEBHOOK=https://...  -> POST every submission to Airtable/ServiceNow/Zapier/etc.
  ALLOWED_DOMAINS=hcltech.com   -> comma list of email domains allowed to sign in (SSO gate)
"""
import os, json, sqlite3, secrets, string, datetime, urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
import integrations as intg

# in-memory stores (auth sessions + pending OIDC handshakes)
SESSIONS = {}      # session_token -> user dict
OIDC_WAIT = {}     # state -> {nonce, verifier}

ROOT = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(ROOT, "dealdesk.db"))
PORT = int(os.environ.get("PORT", "4610"))
WEBHOOK = os.environ.get("DEALDESK_WEBHOOK", "").strip()
ALLOWED_DOMAINS = [d.strip().lower() for d in
                   os.environ.get("ALLOWED_DOMAINS", "hcltech.com").split(",") if d.strip()]

# ---------------------------------------------------------------- DB layer
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS projects(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_code TEXT UNIQUE,
            deal_name TEXT, company TEXT, status TEXT DEFAULT 'New',
            tcv REAL DEFAULT 0, currency TEXT DEFAULT 'USD',
            services_json TEXT, payload_json TEXT,
            created_by TEXT, created_at TEXT, updated_at TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS drafts(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner TEXT, label TEXT, payload_json TEXT, updated_at TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS activity(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor TEXT, action TEXT, detail TEXT, at TEXT)""")

def now():
    return datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

def code():
    return "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(12))

def log(actor, action, detail=""):
    with db() as c:
        c.execute("INSERT INTO activity(actor,action,detail,at) VALUES(?,?,?,?)",
                  (actor, action, detail, now()))

def _make_csv(headers, rows):
    import csv, io
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(headers)
    for r in rows:
        w.writerow(r)
    return "﻿" + buf.getvalue()  # BOM so Excel reads UTF-8 cleanly

def _xl_esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))

def _make_xlsx(headers, rows):
    """Minimal valid .xlsx (OOXML) using only stdlib zipfile — no openpyxl."""
    import io, zipfile
    def col(n):
        s = ""
        n += 1
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    def cell(ci, ri, val):
        ref = f"{col(ci)}{ri}"
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return f'<c r="{ref}"><v>{val}</v></c>'
        return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{_xl_esc(val)}</t></is></c>'
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    sheet.append("<row r=\"1\">" + "".join(cell(i, 1, h) for i, h in enumerate(headers)) + "</row>")
    for ri, row in enumerate(rows, start=2):
        sheet.append(f'<row r="{ri}">' + "".join(cell(i, ri, v) for i, v in enumerate(row)) + "</row>")
    sheet.append("</sheetData></worksheet>")
    sheet_xml = "".join(sheet)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheets><sheet name="Deal Desk" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wbr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbr)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return buf.getvalue()

def forward_webhook(payload):
    if not WEBHOOK:
        return None
    try:
        req = urllib.request.Request(WEBHOOK, data=json.dumps(payload).encode(),
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=8) as r:
            return r.status
    except Exception as e:
        log("system", "webhook_error", str(e))
        return None

# ---------------------------------------------------------------- HTTP
class H(BaseHTTPRequestHandler):
    def _send(self, code_, body=b"", ctype="application/json"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode()
        if isinstance(body, str):
            body = body.encode()
        self.send_response(code_)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _redirect(self, url, cookie=None):
        self.send_response(302)
        self.send_header("Location", url)
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()

    def _cookie(self, name):
        raw = self.headers.get("Cookie", "")
        for part in raw.split(";"):
            if "=" in part:
                k, v = part.strip().split("=", 1)
                if k == name:
                    return v
        return None

    def _current_user(self):
        tok = self._cookie("euler_sess")
        return SESSIONS.get(tok) if tok else None

    def _body(self):
        n = int(self.headers.get("Content-Length", 0) or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except Exception:
            return {}

    def log_message(self, *a):
        pass  # quiet

    # ---- routes
    def do_HEAD(self): self.do_GET()

    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html"):
            return self._file("app.html", "text/html; charset=utf-8")
        if p == "/api/health":
            return self._send(200, {"ok": True})
        if p == "/api/config":
            return self._send(200, intg.config_status())
        if p == "/api/me":
            u = self._current_user()
            return self._send(200, {"user": u}) if u else self._send(401, {"error": "not signed in"})
        if p == "/api/logout":
            tok = self._cookie("euler_sess")
            SESSIONS.pop(tok, None)
            return self._redirect("/", "euler_sess=; Path=/; Max-Age=0")
        if p == "/api/auth/login":
            # begin Microsoft Entra OIDC handshake
            state, nonce = secrets.token_urlsafe(16), secrets.token_urlsafe(16)
            verifier = secrets.token_urlsafe(48)
            OIDC_WAIT[state] = {"nonce": nonce, "verifier": verifier}
            return self._redirect(intg.azure_authorize_url(state, nonce, verifier))
        if p == "/api/auth/callback":
            q = parse_qs(urlparse(self.path).query)
            state = (q.get("state") or [""])[0]
            code = (q.get("code") or [""])[0]
            wait = OIDC_WAIT.pop(state, None)
            if not wait or not code:
                return self._redirect("/?sso_error=state")
            try:
                claims = intg.azure_exchange_code(code, wait["verifier"])
            except Exception as e:
                log("system", "sso_error", str(e))
                return self._redirect("/?sso_error=exchange")
            dom = (claims["email"].split("@")[-1] if "@" in claims["email"] else "")
            if ALLOWED_DOMAINS and dom not in ALLOWED_DOMAINS:
                return self._redirect("/?sso_error=domain")
            tok = secrets.token_hex(24)
            SESSIONS[tok] = {"email": claims["email"], "name": claims["name"], "via": "azure"}
            log(claims["email"], "login", "azure-sso")
            return self._redirect("/", f"euler_sess={tok}; Path=/; HttpOnly; SameSite=Lax")
        if p == "/api/stats":
            with db() as c:
                rows = c.execute("SELECT status, COUNT(*) n, COALESCE(SUM(tcv),0) tcv FROM projects GROUP BY status").fetchall()
            total = sum(r["n"] for r in rows)
            tcv = sum(r["tcv"] for r in rows)
            by = {r["status"]: r["n"] for r in rows}
            done = by.get("Complete", 0) + by.get("Completed", 0) + by.get("Committed / Won", 0)
            prog = by.get("In Progress", 0) + by.get("Qualified", 0) + by.get("Negotiation", 0) + by.get("Proposal", 0)
            return self._send(200, {"total": total, "complete": done, "in_progress": prog,
                                    "not_started": by.get("New", 0) + by.get("Not Started", 0), "tcv": tcv})
        if p == "/api/projects":
            with db() as c:
                rows = c.execute("SELECT * FROM projects ORDER BY id DESC").fetchall()
            return self._send(200, [self._proj(r) for r in rows])
        if p in ("/api/export/projects.csv", "/api/export/projects.xlsx"):
            headers, rows = self._export_rows()
            if p.endswith(".csv"):
                return self._download(_make_csv(headers, rows), "deal-desk-projects.csv", "text/csv")
            return self._download(_make_xlsx(headers, rows), "deal-desk-projects.xlsx",
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if p.startswith("/api/projects/"):
            pid = p.rsplit("/", 1)[-1]
            with db() as c:
                r = c.execute("SELECT * FROM projects WHERE project_code=? OR id=?", (pid, pid)).fetchone()
            return self._send(200, self._proj(r)) if r else self._send(404, {"error": "not found"})
        if p == "/api/drafts":
            owner = self._q("owner")
            with db() as c:
                rows = c.execute("SELECT * FROM drafts WHERE owner=? ORDER BY id DESC", (owner,)).fetchall()
            return self._send(200, [{"id": r["id"], "label": r["label"],
                                     "updated_at": r["updated_at"], "payload": json.loads(r["payload_json"] or "{}")} for r in rows])
        if p == "/api/activity":
            with db() as c:
                rows = c.execute("SELECT * FROM activity ORDER BY id DESC LIMIT 25").fetchall()
            return self._send(200, [dict(r) for r in rows])
        # static fallback
        return self._file(p.lstrip("/"), None)

    def do_POST(self):
        p = self.path.split("?")[0]
        b = self._body()
        if p == "/api/login":
            email = (b.get("email") or "").strip().lower()
            if "@" not in email:
                return self._send(400, {"error": "Enter a valid email"})
            dom = email.split("@")[-1]
            if ALLOWED_DOMAINS and dom not in ALLOWED_DOMAINS:
                return self._send(403, {"error": f"SSO restricted to: {', '.join(ALLOWED_DOMAINS)}"})
            name = email.split("@")[0].replace(".", " ").title()
            user = {"email": email, "name": name, "domain": dom, "via": "mock"}
            tok = secrets.token_hex(24)
            SESSIONS[tok] = user
            log(email, "login", "mock")
            body = json.dumps({"ok": True, "user": user}).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Set-Cookie", f"euler_sess={tok}; Path=/; HttpOnly; SameSite=Lax")
            self.end_headers()
            return self.wfile.write(body)
        if p == "/api/projects":
            pc = code()
            ts = now()
            services = b.get("services") or []
            with db() as c:
                c.execute("""INSERT INTO projects(project_code,deal_name,company,status,tcv,currency,
                    services_json,payload_json,created_by,created_at,updated_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (pc, b.get("dealName", ""), b.get("company", ""), b.get("status", "New"),
                     float(b.get("tcv") or 0), b.get("currency", "USD"),
                     json.dumps(services), json.dumps(b), b.get("createdBy", ""), ts, ts))
            log(b.get("createdBy", "user"), "project_created", f'{b.get("dealName","")} ({pc})')
            # ---- system of record: Salesforce custom object + Microsoft Graph email
            sf = intg.salesforce_create(b, pc)
            if sf.get("ok"):
                log("system", "salesforce_synced", f'{pc} -> {sf.get("id")}')
            elif not sf.get("skipped"):
                log("system", "salesforce_error", json.dumps(sf)[:300])
            mail = intg.send_email(b, pc, sf)
            if mail.get("ok"):
                log("system", "email_sent", pc)
            elif not mail.get("skipped"):
                log("system", "email_error", json.dumps(mail)[:300])
            wh = forward_webhook({"event": "deal_desk.project_created", "project_code": pc, "data": b})
            return self._send(201, {"ok": True, "project_code": pc, "created_at": ts,
                                    "salesforce": sf, "email": mail, "webhook_status": wh})
        if p == "/api/drafts":
            ts = now()
            owner = b.get("owner", "")
            label = b.get("label") or (b.get("payload", {}).get("dealName") or "Untitled draft")
            did = b.get("id")
            with db() as c:
                if did:
                    c.execute("UPDATE drafts SET label=?,payload_json=?,updated_at=? WHERE id=? AND owner=?",
                              (label, json.dumps(b.get("payload", {})), ts, did, owner))
                else:
                    cur = c.execute("INSERT INTO drafts(owner,label,payload_json,updated_at) VALUES(?,?,?,?)",
                                    (owner, label, json.dumps(b.get("payload", {})), ts))
                    did = cur.lastrowid
            return self._send(200, {"ok": True, "id": did, "updated_at": ts})
        if p == "/api/export/crm":
            # Returns a normalized CRM record + optionally forwards to webhook
            data = b.get("data", b)
            rec = {
                "object": "Opportunity",
                "name": data.get("dealName"),
                "account": data.get("company"),
                "amount": data.get("tcv"),
                "currency": data.get("currency"),
                "stage": data.get("status"),
                "services": data.get("services"),
                "use_cases": data.get("useCases"),
                "primary_contact": {"name": data.get("pocName"), "email": data.get("pocEmail"),
                                    "title": data.get("pocTitle"), "phone": data.get("pocPhone")},
                "delivery_lead": {"name": data.get("leadName"), "email": data.get("leadEmail")},
                "source": "EULER Deal Desk",
            }
            wh = forward_webhook({"event": "deal_desk.crm_export", "data": rec})
            return self._send(200, {"ok": True, "crm_record": rec, "webhook_status": wh})
        return self._send(404, {"error": "no route"})

    def do_PUT(self):
        p = self.path.split("?")[0]
        b = self._body()
        if p.startswith("/api/projects/"):
            pid = p.rsplit("/", 1)[-1]
            ts = now()
            with db() as c:
                r = c.execute("SELECT * FROM projects WHERE project_code=? OR id=?", (pid, pid)).fetchone()
                if not r:
                    return self._send(404, {"error": "not found"})
                merged = json.loads(r["payload_json"] or "{}")
                merged.update(b)  # patch
                c.execute("""UPDATE projects SET deal_name=?,company=?,status=?,tcv=?,currency=?,
                    services_json=?,payload_json=?,updated_at=? WHERE id=?""",
                    (merged.get("dealName", ""), merged.get("company", ""), merged.get("status", "New"),
                     float(merged.get("tcv") or 0), merged.get("currency", "USD"),
                     json.dumps(merged.get("services") or []), json.dumps(merged), ts, r["id"]))
            log(b.get("updatedBy", "user"), "project_updated", f'{merged.get("dealName","")} ({r["project_code"]})')
            return self._send(200, {"ok": True, "project_code": r["project_code"], "updated_at": ts})
        return self._send(404, {"error": "no route"})

    def do_DELETE(self):
        p = self.path.split("?")[0]
        if p.startswith("/api/projects/"):
            pid = p.rsplit("/", 1)[-1]
            with db() as c:
                c.execute("DELETE FROM projects WHERE project_code=? OR id=?", (pid, pid))
            log("user", "project_deleted", pid)
            return self._send(200, {"ok": True})
        if p.startswith("/api/drafts/"):
            did = p.rsplit("/", 1)[-1]
            with db() as c:
                c.execute("DELETE FROM drafts WHERE id=?", (did,))
            return self._send(200, {"ok": True})
        return self._send(404, {"error": "no route"})

    # ---- helpers
    def _q(self, k):
        from urllib.parse import urlparse, parse_qs
        return (parse_qs(urlparse(self.path).query).get(k) or [""])[0]

    def _proj(self, r):
        if not r:
            return None
        d = dict(r)
        d["services"] = json.loads(d.pop("services_json") or "[]")
        d["payload"] = json.loads(d.pop("payload_json") or "{}")
        return d

    def _download(self, data, filename, ctype):
        if isinstance(data, str):
            data = data.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def _export_rows(self):
        cols = [("Project ID", "project_code"), ("Deal Name", "dealName"), ("Excalibur ID", "excaliburId"),
                ("SFDC Opportunity ID", "sfdcId"), ("Linked Deal", "linkedDeal"), ("Customer", "company"),
                ("Website", "website"), ("Business Unit", "businessUnit"), ("Status", "status"),
                ("Currency", "currency"), ("TCV", "tcv"), ("Services", "services"),
                ("Primary Use Case", "customerUseCase"), ("Delivery Model", "deliveryModel"),
                ("Anthropic Products", "products"), ("Start Date", "startDate"), ("End Date", "endDate"),
                ("POC Name", "pocName"), ("POC Title", "pocTitle"),
                ("POC Email", "pocEmail"), ("POC Phone Code", "pocPhoneCountry"), ("POC Phone", "pocPhone"),
                ("Delivery Lead", "leadName"), ("Delivery Lead Email", "leadEmail"),
                ("Project Description", "projectDescription"), ("Rollout Plan", "rolloutPlan"),
                ("Registered By", "createdBy")]
        with db() as c:
            recs = c.execute("SELECT * FROM projects ORDER BY id DESC").fetchall()
        headers = [h for h, _ in cols] + ["Created", "Updated"]
        rows = []
        for r in recs:
            pl = json.loads(r["payload_json"] or "{}")
            row = []
            for _, k in cols:
                v = pl.get(k, "")
                if k == "project_code": v = r["project_code"]
                if k == "tcv": v = r["tcv"]
                if isinstance(v, list): v = "; ".join(str(x) for x in v)
                row.append("" if v is None else v)
            row += [r["created_at"], r["updated_at"]]
            rows.append(row)
        return headers, rows

    def _file(self, rel, ctype):
        fp = os.path.normpath(os.path.join(ROOT, rel))
        if not fp.startswith(ROOT) or not os.path.isfile(fp):
            return self._send(404, {"error": "not found"})
        if ctype is None:
            ext = os.path.splitext(fp)[1]
            ctype = {".html": "text/html; charset=utf-8", ".js": "text/javascript",
                     ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png"}.get(ext, "application/octet-stream")
        with open(fp, "rb") as f:
            self._send(200, f.read(), ctype)


if __name__ == "__main__":
    init_db()
    print(f"EULER · Deal Desk  →  http://localhost:{PORT}")
    print(f"  DB: {DB_PATH}")
    print(f"  SSO domains: {', '.join(ALLOWED_DOMAINS) or 'any'}")
    cs = intg.config_status()
    print(f"  SSO mode:   {cs['sso_mode']}   (azure = Entra ID configured)")
    print(f"  Salesforce: {'configured (' + intg.SF_OBJECT + ')' if cs['salesforce'] else 'not configured'}")
    print(f"  Email:      {'configured (Graph -> ' + ', '.join(intg.MAIL_TO) + ')' if cs['email'] else 'not configured'}")
    print(f"  Webhook:    {WEBHOOK or '(none)'}")
    ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()
