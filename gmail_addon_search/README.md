# gmail_addon_search — Odoo Companion Module

Odoo 19+ module that provides the backend API endpoints used by the **Gmail Add-on** (`gmail_addon/`).

Without this module, the Gmail add-on cannot search tasks, create records, or log emails. The Docs/Sheets and Chat add-ons do **not** require this module.

---

## What this module provides

| Endpoint | Description |
|----------|-------------|
| `/gmail_addon/suggest_context` | Detect Odoo partner from sender email; suggest project/team |
| `/gmail_addon/search_tasks` | Search `project.task` with domain filters |
| `/gmail_addon/search_tickets` | Search `helpdesk.ticket` with domain filters |
| `/gmail_addon/create_task` | Create a new task pre-filled from email data |
| `/gmail_addon/create_ticket` | Create a new helpdesk ticket |
| `/gmail_addon/log_email` | Post email as internal note on a task or ticket |
| `/gmail_addon/autocomplete` | Autocomplete partners, projects, stages, teams |

Authentication uses the `mail_plugin` Bearer token flow (Odoo API key passed as `Authorization: Bearer <api_key>`).

---

## Requirements

- Odoo **19.0** or later
- Odoo modules: `mail_plugin`, `project` (installed as dependencies)
- Optional: `helpdesk` module (Odoo Enterprise) — enables ticket endpoints; gracefully skipped if not installed

---

## Installation

### 1. Copy the module to your addons path

```bash
cp -r gmail_addon_search/ /path/to/your/odoo/addons/
```

Or add the parent folder to your `addons_path` in `odoo.conf`.

### 2. Update the module list

```bash
./odoo-bin -d <your_db> --update=base
```

Or in Odoo: **Settings > Technical > Update Apps List**.

### 3. Install the module

**Option A — from the Odoo Apps menu:**

1. Go to **Apps**.
2. Search for `Gmail Add-on Search & Create`.
3. Click **Install**.

**Option B — from the command line:**

```bash
./odoo-bin -d <your_db> -i gmail_addon_search
```

### 4. Verify dependencies

The module automatically installs `mail_plugin` and `project`. If `helpdesk` is installed, ticket endpoints activate automatically.

---

## Authentication

All endpoints use the `mail_plugin` Bearer token flow:

```
Authorization: Bearer <odoo_api_key>
```

The API key must belong to a user with appropriate access rights to the records being searched or created.

Generate API keys in Odoo: **Settings > Users > (your user) > API Keys > New**.

---

## Module structure

```
gmail_addon_search/
├── __manifest__.py          ← module metadata and dependencies
├── __init__.py
├── controllers/
│   └── main.py              ← HTTP controllers (all API endpoints)
├── models/
│   ├── gmail_email_link.py  ← Gmail ↔ Odoo record link model
│   ├── project_task_type.py ← extends project.task.type
│   └── helpdesk_stage.py    ← extends helpdesk.stage (loaded if helpdesk is available)
└── views/
    ├── project_task_type_views.xml  ← task stage form extension
    └── helpdesk_stage_views.xml     ← helpdesk stage form extension
```

---

## Upgrading

When updating the module:

```bash
./odoo-bin -d <your_db> -u gmail_addon_search
```

Or in Odoo: **Apps > Gmail Add-on Search & Create > Upgrade**.

---

## Uninstalling

Uninstalling this module disables all Gmail add-on API endpoints. The Gmail add-on will show connection errors until the module is reinstalled.

To uninstall: **Apps > Gmail Add-on Search & Create > Uninstall**.

---

## Changelog

### 19.0.1.4.0 — 2026-03-04
- `email_linked_records` endpoint now returns `task_number` and `user_name` for tasks, and `ticket_ref` and `user_name` for tickets, so the add-on can display reference IDs and compose draft messages
- `ticket_ref` and `user_id` (→ `user_name`) added to `_TICKET_FIELDS` and `_format_ticket_dict`, making them available in ticket search results and `suggest_context` recent tickets

### 19.0.1.3.0 — 2026-03-03
- Add `gmail_hide_in_search` boolean field to `project.task.type` and `helpdesk.stage`
- Tasks/tickets whose stage has this flag enabled are excluded from all search results, `suggest_context` recent items, and stage filter dropdowns
- View extensions add the checkbox to both stage form views in Odoo
- Helpdesk model extension is loaded conditionally so the module remains installable without the Helpdesk module
- Fix: correct task stage form view XML ID (`project.task_type_edit`) and helpdesk stage form view XML ID (`helpdesk.helpdesk_stage_view_form`) for Odoo 19
- Fix: replace deprecated `type='json'` with `type='jsonrpc'` on all routes (Odoo 19)
- Fix: add missing `author` key to module manifest
- Fix: use `//sheet` xpath so the boolean checkbox renders correctly in stage form views

### 19.0.1.2.0 — 2026-03-03
- Fix CC follower resolution: use RFC 2822–compliant parsing (`email.utils.getaddresses`) to correctly handle `"Name <email>"` address strings
- Partner creation now uses the display name from the CC header instead of the raw address string
- Per-address exception handling: a single unparseable CC address no longer aborts follower subscription for the rest

### 19.0.1.0.0 — 2026-02-01
- Initial release
- Partner detection from sender email with commercial partner resolution
- Project/team suggestion based on partner task/ticket history
- Task search with domain builder (project, stage, freetext)
- Ticket search with domain builder (team, stage, freetext)
- Task creation from email context (subject, sender, CC, body)
- Ticket creation from email context
- Email logging as internal note on tasks and tickets
- Autocomplete for partners, projects, stages, helpdesk teams
- Graceful degradation when `helpdesk` module is not installed
- License: LGPL-3

---

## Roadmap


