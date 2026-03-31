# gmail_addon_search — Odoo Companion Module

Odoo 19+ companion module for the Google Workspace add-on in this repository. It exposes the JSON-RPC endpoints used by Gmail, Docs, Sheets, and Chat, and it adds Odoo settings for Gmail-specific field configuration.

## What it provides

### Record APIs

- Task search, create, and email logging
- Ticket search, create, and email logging
- CRM lead/opportunity search, create, and email logging
- Linked-record lookup for Gmail threads and linked Docs/Sheets files
- Sender-context suggestion for recent and recommended records

### Odoo configuration

In **Settings**, admins can configure per model:

- which field is used as the displayed/searchable reference
- which extra fields should appear on Gmail create forms
- a direct Outlook manifest download link generated from the current Odoo base URL

Supported model groups:

- `project.task`
- `helpdesk.ticket`
- `crm.lead`

### Odoo model additions

- `gmail.email.link` for Gmail thread/message to Odoo record links
- `gmail.document.link` for Docs/Sheets links
- task and helpdesk stage flags to hide stages from add-on search

## Important routes

| Route | Purpose |
|------|---------|
| `/gmail_addon/suggest_context` | Sender matching, suggested records, recent items |
| `/gmail_addon/task/search` | Search tasks |
| `/gmail_addon/task/create` | Create task |
| `/gmail_addon/ticket/search` | Search tickets |
| `/gmail_addon/ticket/create` | Create ticket |
| `/gmail_addon/lead/search` | Search leads and opportunities |
| `/gmail_addon/lead/create` | Create lead or opportunity |
| `/gmail_addon/log_email` | Post current email as note on a record |
| `/gmail_addon/form/schema` | Return Gmail form schema for configured extra fields |

Authentication uses the `mail_plugin` Bearer-token flow:

`Authorization: Bearer <odoo_api_key>`

## Requirements

- Odoo 19.0+
- Required modules: `mail_plugin`, `project`
- Optional:
  - `helpdesk` for ticket support
  - `crm` for lead/opportunity support

## Installation

```bash
cp -r gmail_addon_search/ /path/to/odoo/addons/
./odoo-bin -d <db> --update=base
./odoo-bin -d <db> -i gmail_addon_search
```

To upgrade after changes:

```bash
./odoo-bin -d <db> -u gmail_addon_search
```

## Module structure

```text
gmail_addon_search/
├── controllers/main.py              # JSON-RPC endpoints
├── models/gmail_addon_settings.py   # Settings + field schema/config logic
├── models/gmail_email_link.py       # Gmail thread/message links
├── models/gmail_document_link.py    # Docs/Sheets links
├── models/project_task_type.py      # Task stage visibility flag
├── models/helpdesk_stage.py         # Helpdesk stage visibility flag
└── views/res_config_settings_views.xml
```

## Notes

- Ticket features return safe fallback errors if Helpdesk is not installed.
- CRM features return safe fallback errors if CRM is not installed.
- Docs/Sheets and Chat currently use only the task/ticket parts of this module.
