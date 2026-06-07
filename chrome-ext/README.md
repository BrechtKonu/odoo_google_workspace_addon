# Konu Odoo Workspace — Chrome extension (phase 1)

A **thin** Chrome (MV3) companion to the Apps Script add-on. It does **only**
the things an Apps Script add-on structurally *cannot* do, and reuses the exact
same Odoo HTTP endpoint (`gmail_addon_search`, Bearer/API-key auth).

> **Status: scaffold.** Phase 1 = read / inline-augment. Writes, creation,
> attachments, Docs/Sheets insertion, Chat and **mobile** all stay in the Apps
> Script add-on, which remains the auth/write backbone and the only
> cross-platform path. See the deep-review recommendation in the repo root
> `README.md`.

## What it does now

- **Inline reference detection in the Gmail read view** — scans the open
  message body for Odoo references (`KOTASK-053`, `LATR.PS-002`,
  `LATR.HT-2095`) and turns them into clickable chips that resolve the live
  record and open it in Odoo. (Apps Script can't touch the message DOM.)
- **Popup** — configure the Odoo URL + API key, test the connection, and look
  up a reference.

## Architecture

```
manifest.json        MV3 manifest (content script on mail.google.com, SW, popup)
background.js        Service worker — ALL Odoo network calls happen here
lib/odoo.js          Shared JSON-RPC client (Bearer token, same routes as the add-on)
content/gmail.js     Read-view ref detection; selectors isolated for Gmail-DOM churn
content/inline.css   Chip styling
popup.html/js        Connection config + reference lookup
```

**Why all network calls go through the service worker:** content scripts run in
the page origin (`mail.google.com`) and are subject to CORS; the service worker
uses the manifest `host_permissions` and is not. So the content script and popup
only ever `chrome.runtime.sendMessage` — they never `fetch` Odoo directly. This
means **no CORS changes are needed on the Odoo controller** for phase 1.

## Backend prerequisites for later phases

- **Batch linked-records endpoint.** Painting "linked" badges on a 50-row inbox
  with the current single-id `/gmail_addon/email/linked_records` would be 50
  calls. A batch variant is needed first.
- **CORS** only becomes relevant if a future phase fetches Odoo directly from a
  page context (not planned).

## Install (unpacked, for testing)

1. `chrome://extensions` → enable Developer mode → **Load unpacked** → select
   this `chrome-ext/` folder.
2. Click the extension icon, enter your Odoo URL + API key, **Test connection**.
3. Open a Gmail message containing a reference like `KOTASK-053`.

## Not included

Icons are intentionally omitted from the scaffold (no binaries committed). Add
`icons/` + an `"icons"` manifest key before publishing.
