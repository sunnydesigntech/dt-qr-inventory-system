# D&T QR Inventory System

Google Apps Script inventory web app backed by Google Sheets.

## Architecture

The runtime app now uses the prototype UI as Apps Script `HtmlService` templates:

- `code.gs` contains backend logic, routing payloads, sheet access, admin tools, import helpers, QR helpers, and save validation.
- `index.html` is the web app shell.
- `app_styles.html` contains the prototype-derived design tokens and UI CSS.
- `app_script.html` contains the prototype-derived plain JavaScript UI wired to live Apps Script data.
- `appsscript.json` is the Apps Script manifest.

Repo support files are allowed:

- `.clasp.json` links this folder to the Apps Script project.
- `.claspignore` keeps repo-only files, including `prototype/`, out of Apps Script pushes.
- `prototype/` contains the static design prototype reference.
- `README.md` and `LICENSE` are local repository files.

## Requirements

- Node.js/npm available locally.
- `clasp` available through `npx --yes @google/clasp ...` or installed globally.
- A valid clasp login for the VSA Google account:

```sh
npx --yes @google/clasp login
```

## Apps Script Project

Script ID:

```text
1p0WyeTnFlWpjcQiDrkEamHDv3T3VgbGD73ENc0X6ix7OP2NcoeLuVxNr
```

Current web app deployment ID:

```text
AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ
```

Current deployed version:

```text
@12 - 419A location-code database workflow
```

Current web app URL:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec
```

## Script Properties

Required:

- `SPREADSHEET_ID`

Recommended:

- `WEB_APP_BASE_URL`

Optional:

- `INVENTORY_SHEET_NAME`

Set these from the spreadsheet menu: `D&T Inventory` -> `Set App Config`, or use Apps Script project settings.

## Local Checks

Syntax:

```sh
node --check --input-type=commonjs < code.gs
```

Local browser preview:

```sh
node scripts/build-preview.mjs
```

Then open `preview.html` in a browser. Do not open `app_script.html` directly; it is an Apps Script include fragment and needs the `index.html` shell plus bootstrap data.

Confirm only intended Apps Script runtime files are tracked for push:

```sh
npx --yes @google/clasp status
```

Expected tracked files:

- `app_script.html`
- `app_styles.html`
- `appsscript.json`
- `code.gs`
- `index.html`

## Push And Deploy

Push source to Apps Script:

```sh
npx --yes @google/clasp push --force
```

Create/update a deployment only after testing the pushed code. To update the current web app URL:

```sh
npx --yes @google/clasp deploy \
  --deploymentId AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ \
  --description "D&T QR Inventory deployment"
```

Do not redeploy during development unless the implementation pass is complete and tested.

## Test URLs

The app uses one Apps Script deployment URL. Individual storage pages are in-app routes on that same deployment, using query parameters after `/exec`.

Base app / landing page:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec
```

419A view mode:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec?room=419A&loc=<storage-id-or-location>
```

419A Update Mode:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec?room=419A&loc=<storage-id-or-location>&mode=tech
```

`mode=tech` is retained as the internal/backward-compatible route parameter, but the UI labels this workflow as `Update`.

QR labels should normally encode the View URL, not the Update URL. The printed QR opens the correct storage page in read-only View Mode; authorised users can enter Update Mode from inside the app.

## Admin Menu

The spreadsheet menu exposes:

- Refresh QR Links
- Refresh QR Images (Optional)
- Prepare App Columns
- Build Storage Master
- Build QR Label Sheet
- Create Readiness Report
- 419A Readiness Summary
- Import 419A Storage Master
- Import 419A App Load Ready
- Open Web App
- Set App Config
- Set WEB_APP_BASE_URL
- Config Status / Diagnostics

## 419A Rollout Workflow

For role-specific operating instructions and the go-live gate, use [ROLLOUT_CHECKLIST.md](/Users/wcchun/Documents/inventory-system/ROLLOUT_CHECKLIST.md).

Recommended setup order:

1. Run `Prepare App Columns` to add the optional rollout columns if they are missing.
2. Run `Import 419A Storage Master` using a Google Sheet converted from the authoritative 419A workbook. This creates placeholder storage rows so empty storage pages can still render.
3. Run `Import 419A App Load Ready` to append non-duplicate item rows.
4. Run `Build Storage Master` to generate the room/storage operating map from `Inventory`.
5. Run `Create Readiness Report` and fix any errors or warnings.
6. Run `Refresh QR Links`.
7. Run `Build QR Label Sheet` for printable labels.
8. Run `Refresh QR Images (Optional)` only if the inventory sheet has `QR Code Image`.

## Workshop Operations Columns

`Prepare App Columns` keeps the legacy 8-column sheet compatible and appends optional rollout/operations columns when missing:

- `Unit`
- `Remarks`
- `Location Code`
- `Storage ID`
- `Storage Label`
- `QR Code Image`
- `Storage Type`
- `Last Updated`
- `Updated By`
- `Is Placeholder`
- `Safety Note`
- `Reorder Level`
- `Supplier`
- `Purchase Link`
- `Asset Value`
- `Maintenance Due`
- `SDS Link`

The app remains compatible with the original 8-column sheet. The extended columns enable storage mapping, audit metadata, chemical safety notes, reorder planning, maintenance tracking, and future HoD budget/compliance views.

## Generated Operations Sheets

- `Storage_Master`: one row per unique storage/location, generated from `Inventory`.
- `QR_Labels`: printable QR label rows with View URL, Update URL, QR formula, and hazard label text where relevant.
- `Inventory_Readiness_Report`: critical errors, warnings, QR readiness, 419A rollout status, duplicate checks, chemical safety note checks, reorder-level checks, and maintenance-detail checks.
- `Audit_Log`: appended automatically when Update Mode saves quantity/status changes.

## Bound Sheet Admin Script

The standalone web app project is authoritative for the deployed web runtime. The live dashboard Google Sheet may also need a small bound Apps Script project so the spreadsheet menu shows the same `D&T Inventory` admin tools.

The bound script source is:

- `sheet_admin/InventoryAdmin.gs`

This file is not pushed by the standalone `.claspignore` rules. To install or refresh the live Sheet menu:

1. Open the live dashboard Google Sheet while signed in with an authorised VSA account.
2. Open `Extensions` -> `Apps Script`.
3. If it opens an unrelated project such as `ReadyLoop`, replace/remove the obsolete menu code or create the correct bound script for the dashboard Sheet.
4. Paste the contents of `sheet_admin/InventoryAdmin.gs` into the bound script project.
5. Save, reload the Sheet, and confirm the `D&T Inventory` menu appears.
6. Run non-destructive checks first: `Config Status / Diagnostics` and `419A Readiness Summary`.

Do not run import actions until the source workbooks have been converted to Google Sheets and the source tabs are confirmed.

The storage master import looks for:

- `419A_Storage_Master`
- `419A Storage Master`
- `Room_QR_Label_Plan`
- `RM 419A 2026`

The item import looks for:

- `419A_App_Load_Ready`
- `419A App Load Ready`

Both imports append rows only; they do not delete existing inventory.

## Release Procedure

Pre-release checks:

```bash
node --check --input-type=commonjs < code.gs
npx --yes @google/clasp status
```

Before pushing, confirm `clasp status` lists only these tracked runtime files:

- `app_script.html`
- `app_styles.html`
- `appsscript.json`
- `code.gs`
- `index.html`

Recommended spreadsheet validation before release:

1. Open the configured Google Sheet.
2. Run `D&T Inventory` -> `Config Status / Diagnostics`.
3. Run `D&T Inventory` -> `Create Readiness Report`.
4. Fix all `ERROR` rows in `Inventory_Readiness_Report`.
5. Review `WARN` rows before rollout.

Push source:

```bash
npx --yes @google/clasp push --force
```

Deploy only after the push succeeds and the Apps Script editor shows the expected latest source:

```bash
npx --yes @google/clasp deploy \
  --deploymentId AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ \
  --description "D&T QR Inventory release"
```

Post-deploy smoke tests:

1. Open the landing page and confirm grouped storage cards render.
2. Open a 419A storage page by Storage ID.
3. Open the same 419A storage in Update Mode.
4. Confirm item search and status filter work on a populated location.
5. Confirm a valid empty storage page shows the friendly empty message.
6. Run `Refresh QR Links` and spot-check that `V++` and Storage ID URLs are encoded correctly.
7. Run `Build QR Label Sheet` and spot-check a generated QR image/link.

## Release Log

### 2026-05-01 14:50 HKT

- Version: `@12`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: 419A authoritative Location Code workflow: Room 419A now routes QR/storage pages by the latest `Location Code` values such as `419A-FCU-01` and `419A-CAB-01`, while preserving workbook Storage IDs as metadata. The live Inventory was appended with 53 storage placeholder rows and 40 matched current item rows, including 21 latest Chemical Cabinet 01 chemical rows. Storage_Master, QR links, QR image formulas, QR_Labels, and readiness outputs were regenerated from the live Sheet.
- Rollback note: previous deployment/version `@11` retained in Apps Script version history as fallback.
- Remaining gate: review the remaining readiness warnings, manually reassign unmatched old 419A rows, and sample-scan pilot QR labels before full physical rollout.

### 2026-05-01 12:44 HKT

- Version: `@11`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: Single-deployment QR workflow polish: documented one Apps Script `/exec` URL with route parameters, clarified that QR labels scan to View Mode by default, kept Update Mode as the authorised in-app stock-check workflow, and added the routing explanation to the QR labels UI.
- Rollback note: previous deployment/version `@10` retained in Apps Script version history as fallback.
- Remaining gate: authenticated VSA web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 23:42 HKT

- Version: `@10`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: UI demo operations screens wired into the active template runtime: landing operations tiles, Storage Master view, Audit Log view, Low Stock/Reorder view, Maintenance/Safety view, and clarified single-deployment QR routing.
- Rollback note: previous deployment/version `@9` retained in Apps Script version history as fallback.
- Remaining gate: authenticated VSA web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 22:51 HKT

- Version: `@9`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: Workshop operations workflow hardening: Update wording, category filter, extended operations columns, Storage_Master generation, Audit_Log save logging, stronger QR/readiness checks, and updated matched database workbook.
- Rollback note: previous deployment/version `@8` retained in Apps Script version history as fallback.
- Remaining gate: authenticated VSA web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 16:03 HKT

- Version: `@8`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: Operational rollout workflow polish and QR labels route fix.
- Rollback note: previous deployment/version `@7` retained in Apps Script version history as fallback.

### 2026-04-30 15:41 HKT

- Version: `@7`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Deployment URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`
- Summary: Production hardening diagnostics and UI polish.
- Rollback note: previous deployment/version `@6` retained in Apps Script version history as fallback.

## Rollback

Use Apps Script deployment history to redeploy an earlier version if a rollout has an issue. The previous live deployment before this pass was version `@10`.
