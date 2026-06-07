# Konu — Odoo ⇄ Google Workspace integration

The canonical "everything Google Workspace ⇄ Odoo" product for Konu: search,
create, and **cross-link** Odoo tasks / tickets / leads with emails, Google
Docs and Sheets — from Gmail, Docs, Sheets and Google Chat.

> Ticket: **KOTASK-053**. Odoo target: **19.0**.

## Components

| Path | What it is | Runs where |
| --- | --- | --- |
| `gmail_addon/` | Google **Apps Script** add-on (`.gs`). The UI surface: Gmail card, Docs/Sheets sidebar, Google Chat. | Google cloud (Apps Script) |
| `gmail_addon_search/` | Companion **Odoo module**. HTTP endpoint (`controllers/main.py`) + link models + smart buttons/views. | Odoo (Python) |
| `chrome-ext/` | Thin **Chrome MV3** extension (phase 1, read/inline-augment). Reuses the same Odoo endpoint. | Desktop Chrome |
| `outlook_addin/` | Outlook task-pane add-in. *Out of scope this session.* Served copy lives under `gmail_addon_search/static/outlook_addin/`. | Outlook |

The add-on is the **auth / write / mobile backbone**; the Chrome extension only
adds what Apps Script structurally cannot (inline DOM in Gmail). See
`chrome-ext/README.md`.

## Architecture: the request path

```
Apps Script (.gs)  ──Bearer API key──►  /gmail_addon/*  (controllers/main.py)
   apiTaskSearch_()        ───────────►  /gmail_addon/task/search
   apiEmailLinkRecord_()   ───────────►  /gmail_addon/email/link_record
   apiRecordLinks_()       ───────────►  /gmail_addon/record/links
   apiDocumentLinkRecord_()───────────►  /gmail_addon/document/link_record
```

Auth is the standard Odoo `mail_plugin` Bearer-token (API key) flow
(`auth='outlook'`), so every endpoint runs as the real user and Odoo
record-rules/ACLs apply.

### The cross-link spine

`gmail.email.link` and `gmail.document.link` store email/doc → record links.
Linked emails & docs are surfaced **back inside Odoo** via smart buttons on the
task / ticket / lead form (`gmail.linked.records.mixin`) and browsable under
**Settings ▸ Technical ▸ Workspace Add-on Links**.

## The `.gs` files share one global namespace

Apps Script concatenates all `.gs` files into a single global scope.
`ChatCode.gs` and `DocsSheetsCode.gs` call helpers defined in `Code.gs`
(`odooPost_`, `isConfigured_`, `apiTaskSearch_`, `buildLoginCard_`, …). Treat
`Code.gs` as the shared library; don't duplicate its helpers.

## Deploy (Apps Script)

Three deployed Apps Script projects back this repo (push with **clasp**, do not
hand-copy):

| Script | Purpose |
| --- | --- |
| Konu Workspace add-on | Gmail + Docs/Sheets + Chat add-on (`gmail_addon/`) |
| KOnu - Odoo Docs… | Docs/Sheets smart-link variant |
| Docs ⇄ GitHub | Docs ↔ GitHub sync (separate utility) |

```bash
cd gmail_addon
clasp push          # uses the local .clasp.json (per-developer, gitignored)
```

`.clasp.json` / `.clasprc.json` hold per-developer credentials and are
**gitignored** — never commit them.

## Odoo module

```bash
# install / upgrade (helpdesk required; crm optional)
docker compose exec odoo odoo -u gmail_addon_search -d <db> --stop-after-init
```

Depends on `mail_plugin`, `project`, `helpdesk`. CRM (leads) is optional and
degrades gracefully.

## License

LGPL-3.0-or-later — see `LICENSE`.
