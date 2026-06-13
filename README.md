# EULER · Claude Partner Network — Deal Desk

Internal Deal Desk registration app for the HCLTech sales team, replicating the EULER
Partner Portal flow (Services → New Project → Details → Service Type → Fields → Project detail).

**Zero dependencies** — Python standard library only, with an **embedded SQLite** database.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Debbhar/euler-deal-desk)

> One click → a permanent public login link (`https://<name>.onrender.com`) to share with the sales team.

## Run

```bash
cd dealdesk-euler
python3 server.py
# → http://localhost:4610
```

With nothing configured it runs in **mock mode**: email-only sign-in, SQLite-only storage —
ideal for local demos. Set the env vars below to switch on the real integrations. Each block
activates independently; anything left unset is skipped gracefully (a submission never fails
because an integration is down — SQLite stays the source of truth).

### Core

| Var | Purpose | Default |
|-----|---------|---------|
| `PORT` | server port | `4610` |
| `ALLOWED_DOMAINS` | comma list of email domains allowed through the gate | `hcltech.com` |

### 1 · HCLTech Azure / Entra ID SSO (OpenID Connect, Auth-Code + PKCE)

Your IT registers an app in the HCLTech Entra tenant, sets the redirect URI to
`https://<host>/api/auth/callback`, then provides:

| Var | Purpose |
|-----|---------|
| `AZURE_TENANT_ID` | HCLTech Entra tenant (directory) ID |
| `AZURE_CLIENT_ID` | the registered app's client ID |
| `AZURE_CLIENT_SECRET` | client secret |
| `AZURE_REDIRECT_URI` | must match the app registration (default `http://localhost:4610/api/auth/callback`) |
| `AZURE_ALLOWED_GROUP` | *(optional)* restrict to one security-group object ID |

When set, the login screen becomes **"Sign in with HCLTech SSO"** → Microsoft redirect.
App-registration scopes: `openid profile email User.Read`. Sign-in is also constrained to `ALLOWED_DOMAINS`.

### 2 · Salesforce — custom object `Deal_Registration__c` (Username-Password OAuth)

Create a Connected App (OAuth enabled) in Salesforce and provide:

| Var | Purpose |
|-----|---------|
| `SF_CLIENT_ID` / `SF_CLIENT_SECRET` | Connected App consumer key / secret |
| `SF_USERNAME` | integration user |
| `SF_PASSWORD` | password **+ security token** concatenated |
| `SF_LOGIN_URL` | `https://login.salesforce.com` (prod) or `https://test.salesforce.com` (sandbox) |
| `SF_OBJECT` | target object (default `Deal_Registration__c`) |
| `SF_FIELD_MAP` | *(optional)* JSON `{"<SF field>":"<payload key>"}` to override the default field mapping |

Default field mapping (edit `integrations.py` `_sf_default_map` or set `SF_FIELD_MAP` to match your schema):
`Name, Project_Code__c, Customer_Company__c, Customer_Website__c, Customer_Business_Unit__c, TCV__c,
Currency_Code__c, Status__c, Delivery_Model__c, Services__c, Primary_Use_Case__c, POC_Name__c,
POC_Email__c, POC_Title__c, POC_Phone__c, Delivery_Lead_Name__c, Delivery_Lead_Email__c,
Start_Date__c, End_Date__c, Linked_Deal_Id__c, Registered_By__c, Description__c, Rollout_Plan__c`.

### 3 · Email — Microsoft 365 / Graph (app-only `Mail.Send`)

Reuses the same Entra app as SSO (or its own). Grant **application** permission `Mail.Send`
(admin consent) and provide:

| Var | Purpose |
|-----|---------|
| `GRAPH_TENANT_ID` / `GRAPH_CLIENT_ID` / `GRAPH_CLIENT_SECRET` | defaults to the `AZURE_*` values |
| `MAIL_SENDER` | the from-mailbox UPN, e.g. `dealdesk@hcltech.com` |
| `MAIL_TO` | comma list of Deal Desk recipients |

### Example

```bash
AZURE_TENANT_ID=... AZURE_CLIENT_ID=... AZURE_CLIENT_SECRET=... \
SF_CLIENT_ID=... SF_CLIENT_SECRET=... SF_USERNAME=svc@hcltech.com SF_PASSWORD='pwd+token' \
MAIL_SENDER=dealdesk@hcltech.com MAIL_TO='dealdesk@hcltech.com,apac.dealdesk@hcltech.com' \
python3 server.py
```

On submit each registration is: stored in SQLite → inserted into Salesforce `Deal_Registration__c`
→ emailed to the Deal Desk via Graph. The submit toast and the `/api/projects` response report
`salesforce ✓/✕` and `email ✓/✕` per submission.

## Deploy a shareable login link

The app is a single self-contained web service. To give HCLTech users a public URL they can
sign into and edit deals:

### Render (recommended — free tier, public HTTPS URL)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Debbhar/euler-deal-desk)

1. Click the button above (or Render → **New → Blueprint** pointed at this repo). `render.yaml`
   provisions a Docker web service and gives you `https://euler-deal-desk.onrender.com`.
2. Sign in with GitHub and authorize Render for `Debbhar/euler-deal-desk`, then **Apply**.
3. Add the secret env vars (Azure / Salesforce / Graph) in the service dashboard.
4. Set `AZURE_REDIRECT_URI` to `https://<your-render-url>/api/auth/callback` and add that same
   URL as a redirect URI in the Entra app registration.

That public URL is the permanent shareable login link to put in emails.

> Free-tier note: the filesystem is ephemeral, so SQLite resets on each redeploy. For a pilot
> that's fine. For persistence, attach a Render **Disk** at `/data` and set `DB_PATH=/data/dealdesk.db`
> (uncomment the block in `render.yaml`), or migrate to Postgres — ask and I'll wire it.

### Anywhere else
A `Dockerfile` and `Procfile` are included, so it also runs as-is on **Railway**, **Fly.io**,
**Azure App Service / Container Apps** (natural fit for an HCLTech Azure tenant), or any box with
Python 3 (`python3 server.py`).

## What's included

- **SSO gate** — login restricted to allowed email domains (mock IdP; swap `/api/login` for real SAML/OIDC).
- **Embedded SQLite** (`dealdesk.db`) — `projects`, `drafts`, `activity` tables. Created automatically.
- **6-step New Project wizard** — Details → Service Type (14-service multi-select) → Commercials →
  Scope → Contacts → Review. Each step validates its required fields; **Next** advances until the
  final **Review** page, which has the **Submit** button. Steps are also clickable in the stepper.
- **Reference IDs** — **Excalibur ID** and **SFDC Opportunity ID** (plus optional Linked Deal ID).
- **Per-product detail** — each selected Anthropic product gets its own scope/seats/notes field.
- **International phone** — Customer POC phone has a global country-code selector.
- **14 service offerings** with descriptions (Rollout & Activation … Platform Migration to Claude).
- **Commercials** — TCV with currency + live formatting; delivery model; dates; status.
- **Use cases** — primary Customer Use Case + repeatable use-case portfolio (category + KPI).
- **Contacts** — Customer POC + HCLTech Delivery Lead.
- **Save draft / resume** — drafts persisted per user in SQLite.
- **Project detail view** — lifecycle bar, project info, services & scope, contacts, activity.
- **Edit / update** any registration — the detail view's **Edit** button reopens the wizard
  prefilled; saving issues a `PUT` and logs a `project_updated` activity entry.
- **Download all registrations** as **`.xlsx`** or **`.csv`** (buttons on the Services list and on
  each project) — generated server-side, stdlib only (no openpyxl).
- **Export PDF** (print) and **Export to CRM** (normalized JSON + optional webhook push).

## API

| Method | Path | Purpose |
|--------|------|---------|
| POST | `/api/login` | SSO domain check |
| GET | `/api/stats` | dashboard counters |
| GET/POST | `/api/projects` | list / create |
| GET/PUT/DELETE | `/api/projects/{code}` | detail / update / delete |
| GET | `/api/export/projects.csv` · `.xlsx` | download all registrations |
| GET/POST/DELETE | `/api/drafts` | draft persistence |
| POST | `/api/export/crm` | normalized CRM record (+ webhook) |

## Notes / next steps

- **HCLTech logo:** the app looks for `hcltech-logo.svg` in this folder and uses it automatically
  (top bar + login). If it's absent, it falls back to a typographic `HCLTech™` wordmark. Drop the
  official asset in as `hcltech-logo.svg` — no code change needed. The **Claude** mark is an SVG recreation.
- SSO, Salesforce, and email are real and stdlib-only — no SDKs. They activate purely from env vars.
- For defence-in-depth, verify the Entra `id_token` signature against the tenant JWKS in production
  (`integrations._decode_jwt_payload` currently trusts the token-endpoint response over TLS).
- Confirm the `Deal_Registration__c` field API names match your org, then adjust `_sf_default_map`
  (or set `SF_FIELD_MAP`) accordingly.
- `DEALDESK_WEBHOOK` remains available as an extra fan-out target (Zapier/Make/etc.) alongside Salesforce + email.
