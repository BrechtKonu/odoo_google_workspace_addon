# Gmail Add-on — Odoo Tasks & Tickets

Contextual Gmail add-on for searching, creating, and logging Odoo project tasks and helpdesk tickets directly from email context.

When you open an email, the add-on auto-detects the sender as an Odoo partner, suggests related projects or teams, and lets you act without leaving Gmail.

> **Beta** — the add-on is functional but under active development. Expect rough edges.

---

## Features

| Feature | Description |
|---------|-------------|
| **Home card** | Auto-detects email sender as an Odoo partner; shows linked tasks/tickets for this email; suggests related project/team based on history |
| **Search Tasks** | Filter by project and stage; paginated results; open in Odoo or log email |
| **Search Tickets** | Filter by team and stage; graceful fallback if Helpdesk module is not installed |
| **New Task** | Create a task pre-filled with email subject, sender, CC recipients, and body snippet |
| **New Ticket** | Create a helpdesk ticket pre-filled from email context |
| **Log Email** | Log the current email as an internal note (chatter message) on any task or ticket |
| **Linked records** | Re-opening an email shows which tasks/tickets were already created or logged from it; each record shows its task number or ticket reference |
| **Add to email** | Inserts a reference line (`KOTASK-001 - Task Name - Assignee` + URL) into a reply draft |
| **Add to subject** | Prepends the record reference (`[KOTASK-001]`) to the reply subject in a new draft |

---

## User setup

### 1. Install the add-on in Gmail

1. In Gmail, click the **+** icon at the bottom of the right sidebar (or go to **Settings > Get add-ons**).
2. Search for **Konu** and select **Konu Gmail Add-on**.
3. Click **Install** and authorize all requested Google permissions.

### 2. Connect to your Odoo instance

1. Open any email in Gmail — the Konu icon appears in the right sidebar.
2. Click the icon to open the add-on, then click **Settings**.
3. Fill in the two fields:
   - **Odoo URL** — copy the base URL of your Odoo instance (e.g. `https://yourcompany.odoo.com`)
   - **API Key** — in Odoo, go to your avatar (top-right) > **My Profile** > **API Keys** tab, generate a new key, and paste it here
4. Click **Save & Connect**.

That's it. You can now create tasks, tickets, log emails, and more directly from Gmail.

---

## Requirements

- **Odoo 19+** with the `gmail_addon_search` companion module installed
- An Odoo API key (generated in Odoo: avatar > **My Profile** > **API Keys**)
- Google Workspace account with Gmail

> The `gmail_addon_search` Odoo module is in the `gmail_addon_search/` folder of this repository. See [`gmail_addon_search/README.md`](../gmail_addon_search/README.md) for installation.

---

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | Gmail add-on logic: cards, handlers, Odoo API client, email parsing |
| `ChatCode.gs` | Google Chat add-on logic: slash commands, dialogs, space memory |
| `DocsSheetsCode.gs` | Google Docs & Sheets add-on logic: document context, insert link, per-document memory |
| `EmailProcessor.gs` | Email body processing: inline image resolution, Drive upload, HTML sanitization |
| `appsscript.json` | Apps Script manifest: scopes, triggers, add-on metadata |

---

## Developer installation (from source)

### 1. Install the Odoo module

Before setting up the Gmail add-on, install the companion Odoo module. See [`gmail_addon_search/README.md`](../gmail_addon_search/README.md).

### 2. Create the Apps Script project

1. Go to [https://script.google.com](https://script.google.com) and create a new project.
2. In **Project Settings** (gear icon), enable **Show "appsscript.json" in editor**.
3. Replace `Code.gs` content with the contents of `Code.gs` from this folder.
4. Replace `appsscript.json` with the contents of `appsscript.json` from this folder.
5. If your Odoo domain is not `*.odoo.com`, `*.konu.be`, or `*.konu.care`, add it to `urlFetchWhitelist` in `appsscript.json`.
6. Save the project.

### 3. Deploy as Gmail Add-on

1. **Deploy > Test deployments > Gmail Add-on**.
2. Click **Install** and authorize the requested scopes.

> After any scope change in `appsscript.json`, create a **new** test deployment (not just update) and re-authorize.

---

## Usage

### Home card

Opens automatically when you open an email. Shows:
- Detected Odoo partner (matched by sender email)
- **Linked to this email** — tasks/tickets already created or logged from this email; each row shows the task number or ticket reference, is clickable to open in Odoo, and has **Add to email** / **Add to subject** compose buttons
- Quick action buttons: Search Tasks, Search Tickets, New Task, New Ticket
- Recent tasks/tickets for the detected partner

### Search Tasks / Search Tickets

1. Click **Search Tasks** (or **Search Tickets**) from the home card.
2. Enter a search query (task name, ticket number, etc.).
3. Optionally filter by project/team and stage.
4. Results are paginated. Click the open icon to open the record in Odoo, **Log Email** to attach the current email as a note, **Add to email** to insert the reference into a reply draft, or **Add to subject** to prepend the reference to the reply subject.

### New Task / New Ticket

1. Click **New Task** or **New Ticket**.
2. Fields are pre-filled from the email:
   - Subject → task/ticket name
   - Sender → partner
   - CC recipients → followers
   - Body snippet → description
   - Suggested project/team based on partner history
3. Adjust as needed and click **Create**.

### Log Email

From any task or ticket result, click **Log Email** to post the current email as an internal note (visible in the Odoo chatter).

---

## Troubleshooting

### "HTTP 401" error

- Check that your API key is correct.
- In Odoo: avatar > **My Profile** > **API Keys** — regenerate if needed.
- Confirm the `gmail_addon_search` module is installed and the user has access.

### "Helpdesk module not installed"

- Ticket search and create require the Odoo Helpdesk module (Enterprise edition).
- All task features work without it.

### Scope authorization errors

- After any `appsscript.json` change, create a **new** test deployment.
- Revoke old authorization: **Google Account > Security > Third-party apps**.
- Reinstall and re-authorize.

### Partner not found for sender email

- Partner auto-detection matches by normalized email against `res.partner` records.
- If the sender is not in Odoo, suggestion and pre-fill are skipped.
- You can manually type the partner email in create forms.

### Add-on does not appear in Gmail

- Confirm the deployment type is **Gmail Add-on** (not Editor Add-on or Web App).
- Refresh Gmail completely after installation.

---

## Odoo prerequisites

- Odoo 19+ reachable via HTTPS from Google servers
- `gmail_addon_search` module installed and up to date
- `mail_plugin` and `project` modules installed (dependencies of `gmail_addon_search`)
- Optional: `helpdesk` module for ticket features (Odoo Enterprise)
- API keys enabled in Odoo settings
- API key generated per user

---

## Changelog

### v1.5.0 — 2026-03-13
- **Google Chat add-on** — search tasks/tickets, create task/ticket with selected message pre-filled, per-space project/team memory, "Post to Chat" after creation, `/recent` slash command
- **Google Docs & Sheets add-on** — search tasks/tickets, create task/ticket with selected text pre-filled, per-document project/team memory, insert task/ticket hyperlink at cursor or selected cell, linked records panel per document
- SSL / connection error recovery: `onGmailMessageOpen` and `onGmailCompose` now catch startup exceptions and show the Settings card with the error message instead of crashing with a runtime error dialog
- "Clear Credentials" button on the Settings card: wipes stored URL, API key, and Drive folder ID so users can re-enter credentials after a server migration
- Fix: `buildTaskSearchCard_` and `buildTicketSearchCard_` name collision between `Code.gs` and `ChatCode.gs` (all GAS files share one global scope); Chat versions renamed to `buildChatTaskSearchCard_` / `buildChatTicketSearchCard_`

### v1.4.0 — 2026-03-04
- Linked records section now shows the task number (e.g. `KOTASK-001`) or ticket reference in the row label instead of just "Task" / "Ticket"
- **Add to email** button on linked records and search results: opens a reply-all draft with the record reference and URL inserted in the body
- **Add to subject** button on linked records and search results: opens a reply-all draft with `[KOTASK-001]` (or ticket ref) prepended to the subject
- Ticket search results now show the ticket reference (e.g. `HD-001`) in the row label instead of the bare `#id`
- New `gmail.compose` OAuth scope added — requires a new test deployment and re-authorization

### v1.3.0 — 2026-03-03
- Stage visibility: task stages and helpdesk stages can now be marked "Hide in Gmail Add-on" in Odoo; hidden stages are excluded from search results, recent items, and stage filter dropdowns — no add-on script change required

### v1.2.0 — 2026-03-03
- Resolve CID inline images (`src="cid:..."`) by fetching the full Gmail MIME structure and uploading each image to Drive; inline images are shared publicly and served via the `lh3.googleusercontent.com` CDN so Odoo chatter renders them without auth redirects
- Attachments uploaded to Drive keep domain-only sharing with the standard Drive viewer URL
- Fix CC address parsing: strip display names (`"Name <email>"` → `email`) in both the add-on and the Odoo backend using RFC 2822–compliant parsing
- Login card shows current Google OAuth permission status and a "Re-authorize Google Permissions" button
- Drive folder ID is validated on save; an error notification is shown if the folder is inaccessible

### v1.1.0 — 2026-03-02
- Linked records section on home card: re-opening an email shows which tasks/tickets were already created or logged from it (cross-user visibility via RFC Message-ID)
- Linked record rows are fully clickable and open the record in Odoo

### v1.0.1 — 2026-02-XX
- Person icon next to sender when matched to an Odoo contact
- Prefer partner with sales/invoices when email matches multiple contacts
- Recent records sorted by date, clickable rows

### v1.0.0 — 2026-02-01
- Initial release
- Home card with partner auto-detection and project/team suggestion
- Task search with project and stage filters
- Ticket search with team and stage filters (graceful fallback when Helpdesk not installed)
- New Task creation pre-filled from email context
- New Ticket creation pre-filled from email context
- Log email as internal note on any task or ticket
- Paginated results
- Per-user credential storage

---

## Roadmap

### Gmail add-on

- [x] Images from Gmail visible in Odoo (CID and URL-based inline images uploaded to Drive; served via lh3 CDN for chatter rendering)
- [x] Add images and attachments to project/team folder
- [x] Performance and speed improvements
- [x] Odoo settings to hide certain stages from search results (per-stage "Hide in Gmail Add-on" checkbox)
- [x] Sender icon: green when matched to an Odoo contact, grey when no match found
- [x] SSL / connection error recovery: config card shown with error message instead of crashing
- [x] "Clear Credentials" button on Settings card for credential reset after server migration
- [ ] Odoo catchall email in Settings — auto-CC the catchall address when inserting a reference into a reply, so Odoo's mail gateway threads the reply onto the record
- [ ] AI summary — "Summarise" button on New Task / New Ticket pre-fills the description with an AI-generated summary of the email body
- [ ] Followers selector — explicit multi-select of Odoo users as followers on task/ticket create (in addition to automatic CC followers)
- [ ] filter 'my tasks/tickets' in recent for this contact
- [ ] dropdown to select user in task/ticket search
- [ ] checkbox to show all tasks/tickets of parent company
- [ ] Publish on Odoo App Store and Google Workspace Marketplace

### Google Chat add-on

- [x] Search tasks and tickets from a Chat space or direct message
- [x] New task / new ticket with selected message text pre-filled in description
- [x] Propose project based on last used project in that space (per-space memory)
- [x] Remember pre-filled project/team filter per space
- [x] Show recent tasks/tickets for the current space
- [x] Post task/ticket link to the conversation after creation
- [ ] Create task / ticket from a Chat message via message action (right-click → Create task)
- [ ] Stage filter memory per space (remember last-used stage alongside project/team)

### Google Docs & Sheets add-on

- [x] Search tasks and tickets from any open document or sheet
- [x] New task / new ticket with selected text pre-filled in description
- [x] Propose project based on last used project in that document (per-document memory)
- [x] Remember pre-filled project/team filter per document
- [x] Show recent tasks/tickets linked to the current document
- [x] Insert task/ticket hyperlink at cursor (Docs) or selected cell (Sheets)
- [x] SSL / connection error recovery: config card shown with error message instead of crashing
- [ ] Followers selector — same as Gmail add-on

### General

- [ ] Publish on Odoo App Store and Google Workspace Marketplace
