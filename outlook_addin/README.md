# Outlook Add-in

Outlook taskpane add-in for the same Odoo workflows already exposed to Gmail:

- search tasks, tickets, and leads/opportunities
- create tasks, tickets, and leads/opportunities
- log the current email to Odoo
- open linked records in Odoo
- insert the configured reference into a draft body or subject

## Files

- `manifest.xml`: Outlook add-in manifest for read and compose surfaces
- `taskpane.html`: taskpane entrypoint
- `taskpane.js`: Office.js client, Odoo API wrapper, and UI logic
- `taskpane.css`: taskpane styling
- `commands.html`, `commands.js`: ribbon command host files

## Hosting

There are now two supported hosting modes:

1. **Recommended:** let Odoo serve the Outlook files from the addon itself
2. **Development only:** host `outlook_addin/` separately and use the local `manifest.xml`

### Recommended: host through Odoo

When `gmail_addon_search` is installed, Odoo serves:

- taskpane files from `/gmail_addon_search/static/outlook_addin/...`
- a ready-to-install manifest from `/gmail_addon/outlook/manifest.xml`

This is the easiest path because users only need the Odoo URL. No separate static hosting is required.

### Alternative: standalone hosting

The local `manifest.xml` in this folder still assumes:

`https://localhost:3000/outlook_addin/`

Use that only if you intentionally want to host the add-in outside Odoo.

## Install in Outlook

Install works in two stages: make the manifest reachable, then sideload it into Outlook.

### 1. Get the manifest

- Recommended: download the manifest from your Odoo instance:
  - `https://your-odoo.example.com/gmail_addon/outlook/manifest.xml`
- Development alternative: use the local [manifest.xml](/home/brecht/odoo_google_workspace_addon/outlook_addin/manifest.xml) and host this folder yourself over HTTPS.

### 2. Sideload in Outlook Web

1. Open Outlook on the web.
2. Go to `Get Add-ins`.
3. Open `My add-ins`.
4. Choose `Add a custom add-in` > `Add from file`.
5. Select the manifest you downloaded from Odoo, or the local [manifest.xml](/home/brecht/odoo_google_workspace_addon/outlook_addin/manifest.xml) if you are using standalone hosting.
6. Accept the sideload prompt.

After installation, open any email or draft and use the `Odoo` ribbon button to open the taskpane.

### 3. Sideload in New Outlook desktop

1. Open New Outlook for Windows or macOS.
2. Go to `Get Add-ins`.
3. Open `My add-ins`.
4. Choose `Add a custom add-in` > `Add from file`.
5. Select the manifest you downloaded from Odoo, or the local [manifest.xml](/home/brecht/odoo_google_workspace_addon/outlook_addin/manifest.xml) if you are using standalone hosting.

If your tenant blocks custom add-ins, an Exchange or Microsoft 365 admin must deploy the manifest centrally instead of sideloading it per user.

### 4. First run

1. Open the taskpane from an email or compose window.
2. Enter your Odoo base URL and API key from `mail_plugin`.
3. Save the connection.
4. Test `Home`, `Search`, and `Create` against a real message.

## Odoo backend requirements

The add-in reuses the `gmail_addon_search` Odoo module and its JSON-RPC routes. It expects:

- `mail_plugin` API-key auth
- `/gmail_addon/ping`
- task, ticket, and lead search/create routes
- `/gmail_addon/suggest_context`
- `/gmail_addon/email/linked_records`
- `/gmail_addon/log_email`
- dynamic form schema and dropdown endpoints

## Notes

- Outlook item IDs and conversation IDs are sent to Odoo for record linking.
- Logging currently sends the email body HTML only. Attachment and inline-image parity is intentionally deferred.
- Reference insertion requires Outlook compose mode.
