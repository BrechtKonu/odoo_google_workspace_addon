# gmail_addon_search

Odoo 19+ backend for the Google Workspace add-on (Gmail, Docs, Sheets, Chat). Serves the JSON-RPC endpoints the add-on calls and adds a Settings page for Gmail field config.

## What it does

- Search, create, and log email to tasks, tickets, and CRM leads/opportunities.
- Link Gmail threads and Docs/Sheets files to Odoo records.
- Suggest records from the email sender.
- Per-model settings (task, ticket, lead): choose the reference field, add extra fields to the Gmail create form, get the Outlook manifest link.

## Endpoints

| Route | Purpose |
|------|---------|
| `/gmail_addon/suggest_context` | Sender match, suggested and recent records |
| `/gmail_addon/task/search` · `/task/create` | Tasks |
| `/gmail_addon/ticket/search` · `/ticket/create` | Tickets |
| `/gmail_addon/lead/search` · `/lead/create` | Leads and opportunities |
| `/gmail_addon/log_email` | Log the current email as a note on a record |
| `/gmail_addon/form/schema` | Gmail form schema for the configured fields |

Auth is the `mail_plugin` bearer flow: `Authorization: Bearer <odoo_api_key>`.

## Requirements

Odoo 19.0+, `mail_plugin`, `project`. Optional: `helpdesk` (tickets), `crm` (leads). Ticket and CRM features fall back safely when their module is absent.

## Install

```bash
./odoo-bin -d <db> -i gmail_addon_search
./odoo-bin -d <db> -u gmail_addon_search   # upgrade after changes
```

## Changelog

### 19.0.1.8.0
- Task/ticket search and record preview now return `priority`, `deadline`, `tags`, and `kanban_state` for the Google Chat cards.
