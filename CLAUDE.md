# CLAUDE.md — Odoo ⇄ Google Workspace integration

Repo-specific guidance for AI agents. Konu's global rules (Odoo coding
standards, branch naming, commit conventions, MCP write-safety) still apply —
see `~/claude-rules/docs/rules/`.

## What this repo is

Canonical "everything Google Workspace ⇄ Odoo" product. Three components that
must stay in sync (see `README.md`):

- `gmail_addon/` — Apps Script (`.gs`), the UI. **One shared global namespace**:
  `ChatCode.gs` / `DocsSheetsCode.gs` depend on helpers in `Code.gs`.
- `gmail_addon_search/` — Odoo 19 module: the HTTP endpoint + link models.
- `chrome-ext/` — thin MV3 extension reusing the same endpoint.

## Contract to preserve

The `.gs` `api*_()` wrappers map 1:1 to `controllers/main.py` routes
(`/gmail_addon/...`). If you change a route's name, params, or response shape,
update **both sides** (and `chrome-ext/lib/odoo.js`) in the same change.

## Hard rules

- **Auth:** endpoints use `auth='outlook'` (mail_plugin Bearer/API-key). They
  run as the real user — rely on ACL/record rules, justify every `.sudo()`.
- **`helpdesk` is a hard dependency** (the helpdesk views inherit helpdesk
  forms). `crm` is optional — guard crm model imports and don't add a crm view
  inherit to the always-loaded manifest data.
- **Never edit the served Outlook copy by hand.** `outlook_addin/` and
  `gmail_addon_search/static/outlook_addin/` are byte-identical duplicates;
  treat the `static/` copy as generated.
- **Don't commit clasp creds.** `.clasp.json` / `.clasprc.json` are gitignored.
- User-facing Python strings go through `_()`; description/HTML fields are HTML,
  not markdown.

## Verify before "done"

- Python: `python3 -m py_compile gmail_addon_search/**/*.py`
- XML well-formedness on every `views/*.xml`, `data/*.xml`.
- Apps Script: `node --check` each `.gs` (ES5-compatible syntax).
- Odoo install/upgrade: `-u gmail_addon_search` with no ORM/XML errors
  (needs a running Odoo with helpdesk installed).
- Run `/konu-skills:connie-review` on the diff before pushing.
