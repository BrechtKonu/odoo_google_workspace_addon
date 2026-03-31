# Gmail Add-on — Odoo Tasks, Tickets, and Leads

Google Apps Script add-on for Gmail that lets users search, create, and link Odoo records from the current email. The Gmail surface supports:

- Project tasks
- Helpdesk tickets
- CRM leads and opportunities

Google Docs, Google Sheets, and Google Chat use the same Apps Script project, but CRM support is Gmail-only in the current version.

## Key features

- Context-aware home card with sender matching, linked records, and recent records
- Search cards for tasks, tickets, and CRM leads/opportunities
- Create cards pre-filled from the current email
- Log email to any linked record as an internal note
- Insert a record reference into a reply draft
- Optional Drive upload for inline images and attachments
- Odoo-configurable reference field per model
- Odoo-configurable extra create fields rendered dynamically in Gmail

## Requirements

- Odoo 19+
- `gmail_addon_search` installed in Odoo
- Google Workspace account with Gmail
- Optional:
  - `helpdesk` for ticket features
  - `crm` for lead/opportunity features

See [`gmail_addon_search/README.md`](../gmail_addon_search/README.md) for backend installation.

## Main files

| File | Purpose |
|------|---------|
| `Code.gs` | Gmail logic, cards, Odoo API calls, create/search/log flows |
| `ChatCode.gs` | Google Chat add-on logic |
| `DocsSheetsCode.gs` | Google Docs / Sheets add-on logic |
| `EmailProcessor.gs` | Inline image and attachment handling for Drive logging |
| `appsscript.json` | Manifest, scopes, triggers, and add-on metadata |

## Developer setup

1. Install the Odoo companion module.
2. Create an Apps Script project at `script.google.com`.
3. Enable **Show "appsscript.json" in editor**.
4. Copy the files from `gmail_addon/` into the Apps Script project.
5. If needed, extend `urlFetchWhitelist` in `appsscript.json` for your Odoo domain.
6. Deploy as a **Gmail Add-on** using **Deploy > Test deployments**.

After any OAuth scope change, create a new test deployment and re-authorize.

## Gmail usage overview

### Home card

When an email opens, the add-on:

- detects the sender
- fetches related recent records from Odoo
- shows linked tasks, tickets, and leads for the current thread
- offers quick actions for search and create

### Search

Each search card supports pagination and opens results directly in Odoo.

- **Tasks**: filter by project, stage, assignee
- **Tickets**: filter by team, stage, assignee
- **Leads**: filter by type, sales team, stage, assignee

Reference display and search use the configured Odoo field for that model when one is set.

### Create

Create cards pre-fill the name and description from the current email. Admin-configured extra fields from Odoo settings are rendered automatically in Gmail for tasks, tickets, and leads.

### Log email

Any result or linked record can receive the current email as an internal note. If a Drive folder is configured, images and attachments can be uploaded first and then referenced from the note body.

## Notes

- Ticket features degrade gracefully when Helpdesk is missing.
- Lead/opportunity features degrade gracefully when CRM is missing.
- Chat and Docs/Sheets still cover tasks and tickets only.
