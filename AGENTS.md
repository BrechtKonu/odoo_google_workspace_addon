# Repository Guidelines

## Project Structure & Module Organization
`gmail_addon/` contains the Google Apps Script add-on code: `Code.gs` for Gmail flows, `DocsSheetsCode.gs` for Docs/Sheets, `ChatCode.gs` for Google Chat, `EmailProcessor.gs` for HTML and attachment handling, and `appsscript.json` for scopes and deployment metadata. `gmail_addon_search/` is the Odoo 19 companion module, split into `controllers/`, `models/`, `views/`, `data/`, and `security/`. Keep user-facing documentation in [`USER_GUIDE.md`](/home/brecht/odoo_google_workspace_addon/USER_GUIDE.md) and module-specific notes in each folder’s `README.md`.

## Build, Test, and Development Commands
There is no local build pipeline in this repo; development is driven by Odoo install/upgrade and Apps Script deployment.

```bash
cp -r gmail_addon_search/ /path/to/odoo/addons/
./odoo-bin -d <db> -i gmail_addon_search
./odoo-bin -d <db> -u gmail_addon_search
./odoo-bin -d <db> --update=base
```

Use install for first setup, upgrade after backend changes, and `--update=base` when refreshing the apps list. For the add-on, copy the `.gs` files and `appsscript.json` into a Google Apps Script project, then create a new Gmail Add-on test deployment after any scope change.

## Coding Style & Naming Conventions
Match the existing file style instead of reformatting broadly. In Apps Script, use `var`, 2-space indentation, section banners, and helper names ending in `_` such as `getOdooUrl_()`. In the Odoo module, use 4-space indentation, `snake_case`, concise controller helpers, and standard manifest keys in `__manifest__.py`. Keep XML IDs descriptive and module-scoped.

## Testing Guidelines
No automated test suite or lint config is checked in today. Validate backend changes by upgrading `gmail_addon_search` in an Odoo 19 database and exercising the affected JSON-RPC endpoints. Validate add-on changes manually in Gmail, and when relevant in Docs/Sheets and Chat, especially auth flow, search, create, and log-email paths. If you add automated tests, place them under `gmail_addon_search/tests/` with `test_*.py` names.

## Commit & Pull Request Guidelines
The history is minimal (`first commit`, `initial commit`), so use clearer commit subjects going forward: short, imperative, and scoped when useful, for example `gmail_addon: fix linked record compose action`. PRs should state which surface changed (`gmail_addon` or `gmail_addon_search`), summarize user-visible behavior, note any required Odoo upgrade or manifest/version bump, and include screenshots for card UI changes.

## Security & Configuration Tips
Never commit Odoo API keys, deployment IDs, or customer URLs. When adding new endpoints or external calls, keep `urlFetchWhitelist` in [`gmail_addon/appsscript.json`](/home/brecht/odoo_google_workspace_addon/gmail_addon/appsscript.json) aligned with the required domains.
