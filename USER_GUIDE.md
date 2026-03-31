# Konu Tasks & Tickets User Guide

Konu lets you work with Odoo directly from Google Workspace and Outlook. In Gmail and Outlook, you can search, create, and link:

- Tasks
- Tickets
- Leads and opportunities

Google Docs, Google Sheets, and Google Chat currently support tasks and tickets only.

## Table of Contents

1. [Before you start](#1-before-you-start)
2. [First-time setup](#2-first-time-setup)
3. [Using the add-on in Gmail](#3-using-the-add-on-in-gmail)
4. [Using the add-in in Outlook](#4-using-the-add-in-in-outlook)
5. [Drive attachments and inline images](#5-drive-attachments-and-inline-images)
6. [Using the add-on in Google Docs and Google Sheets](#6-using-the-add-on-in-google-docs-and-google-sheets)
7. [Using the add-on in Google Chat](#7-using-the-add-on-in-google-chat)
8. [Troubleshooting](#8-troubleshooting)

## 1. Before you start

You need:

- an Odoo user account with access to the records you want to use
- your own Odoo API key
- the `gmail_addon_search` module installed in Odoo
- for Outlook: the Outlook add-in hosted over HTTPS and installed from its manifest
- optional: a Google Drive folder ID for image and attachment upload

Optional Odoo apps:

- `helpdesk` for ticket features
- `crm` for lead and opportunity features in Gmail

## 2. First-time setup

### Generate an Odoo API key

1. Open Odoo.
2. Click your avatar and open **My Profile**.
3. Open **API Keys**.
4. Click **New API Key**.
5. Copy the key immediately.

### Connect the add-on

1. Open any email in Gmail.
2. Click the **Konu** icon in the right sidebar.
3. On first use, open **Settings**.
4. Enter:
   - **Odoo URL**
   - **API Key**
   - optional **Drive Folder ID**
5. Click **Save & Connect**.

The same URL and API key are reused across Gmail, Docs, Sheets, and Chat.

### Install the Outlook add-in

If you use Outlook, the easiest setup is to install the manifest from your Odoo server.

1. Open this URL in your browser and download the manifest:
   - `https://your-odoo.example.com/gmail_addon/outlook/manifest.xml`
   - or open Odoo **Settings** and use the **Outlook manifest download** link
2. In Outlook Web or New Outlook desktop, open **Get Add-ins**.
3. Open **My add-ins**.
4. Choose **Add a custom add-in** > **Add from file**.
5. Select the downloaded manifest file.
6. Open an email or draft and launch the **Odoo** taskpane.
7. Enter your **Odoo URL** and **API Key** on first use.

### Odoo admin settings

In Odoo **Settings**, an administrator can configure:

- the reference field shown and searched for tasks, tickets, and leads
- extra fields that appear on Gmail create forms

This means your Gmail forms may contain additional company-specific fields.

## 3. Using the add-on in Gmail

### Home card

When you open an email, the home card shows:

- the matched Odoo contact, when found
- linked records already associated with the thread
- quick actions for search and create
- recent records for the sender

Linked records are clickable and can also offer:

- **Log Email**
- **Add to email**
- **Add to subject**

### Search

Gmail provides separate search cards for:

- **Tasks**
- **Tickets**
- **Leads**

You can search by name or by the configured reference field. Depending on the record type, filters can include project, team, stage, type, and assignee.

### Create

Gmail provides create cards for:

- **New Task**
- **New Ticket**
- **New Lead**

The cards pre-fill values from the current email, including subject, sender, and body text. If your administrator configured extra create fields in Odoo, those fields appear automatically in Gmail.

For leads, you can create either:

- a **Lead**
- an **Opportunity**

### Email logging

From search results, linked records, or recent records, you can log the current email as an internal note on the selected Odoo record.

## 4. Using the add-in in Outlook

Outlook supports the same core email workflows as Gmail for:

- Tasks
- Tickets
- Leads and opportunities

### Home view

When you open an email, Outlook shows:

- linked records found from the message or conversation
- recent records for the sender
- quick actions for search and create

### Search and create

The Outlook taskpane supports:

- task search and create
- ticket search and create
- lead and opportunity search and create

Like Gmail, search uses the configured reference field when one is set in Odoo, and create forms can include admin-configured extra fields.

### Log email and insert reference

From Outlook you can:

- log the current email to a task, ticket, or lead
- open the record in Odoo
- insert the record reference into the draft body
- prepend the record reference to the draft subject

Reference insertion works only in compose mode.

### Current Outlook limitation

Outlook currently logs the email body only. Attachment and inline-image upload parity with Gmail is not included yet.

## 5. Drive attachments and inline images

If a **Drive Folder ID** is configured, the add-on can upload inline images and attachments to Google Drive before logging the email. This helps Odoo chatter render images correctly.

Without a Drive folder, the email still logs, but inline images may not display correctly.

To find a folder ID, open the folder in Google Drive and copy the last part of:

`https://drive.google.com/drive/folders/<folder-id>`

## 6. Using the add-on in Google Docs and Google Sheets

Docs and Sheets support task and ticket workflows.

- selected text can pre-fill the description
- linked records are shown for the current document
- **Insert link** adds a task or ticket reference at the cursor or into the selected cell

Lead/opportunity support is not available here yet.

## 7. Using the add-on in Google Chat

Google Chat supports task and ticket workflows in spaces and direct messages.

Available commands include:

- `/task`
- `/ticket`
- `/config`
- `/recent`
- `/help`

The add-on remembers the last-used project or team per space. Lead/opportunity support is not available in Chat yet.

## 8. Troubleshooting

### The add-on opens with an error

Your saved Odoo URL or API key may be outdated. Open **Settings**, update the values, and save again.

### HTTP 401 Unauthorized

Your API key is invalid or revoked. Generate a new one in Odoo and save it again in the add-on.

### HTTP 403 Forbidden

Your Odoo user does not have access to the requested records or models.

### Ticket features are missing

The `helpdesk` module is not installed in Odoo.

### Lead features are missing

The `crm` module is not installed in Odoo.

### Images do not display in Odoo chatter

Configure a valid **Drive Folder ID** and log the email again.

### Outlook add-in does not appear in Outlook

Check that:

- Odoo is reachable over HTTPS
- `/gmail_addon/outlook/manifest.xml` opens correctly in the browser
- Outlook accepted the manifest during sideloading
- your Microsoft 365 tenant allows custom add-ins

If custom add-ins are blocked, ask your Microsoft 365 or Exchange administrator to deploy the manifest centrally.

### Google permissions error

Re-authorize the add-on, or revoke it under **Google Account > Security > Third-party apps** and open it again.
