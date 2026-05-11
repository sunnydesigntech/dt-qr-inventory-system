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

- `.claspignore` keeps repo-only files out of Apps Script pushes.
- `README.md` and `LICENSE` are local repository files.

Local/private files are intentionally not committed:

- `.clasp.json` because it contains the Apps Script project ID.
- downloaded `.xlsx` workbooks.
- generated database-match outputs.
- prototype screenshots or local UI export artifacts.

## Privacy And Public Repository Notes

This repository is safe to publish as source code and operating documentation. It does not include the live Google Sheet, private student/staff records, downloaded workbook data, Apps Script project ID, live deployment ID, or Google Workspace account details.

Before deploying your own copy, configure these values in Apps Script `Script Properties` or through the spreadsheet menu:

- `SPREADSHEET_ID`
- `WEB_APP_BASE_URL`
- `EXTERNAL_SCANNER_URL`
- `INVENTORY_SHEET_NAME` if your inventory tab is not named `Inventory`

The source code deliberately leaves `DEFAULT_SPREADSHEET_ID` and `DEFAULT_WEB_APP_BASE_URL` blank so public GitHub code does not expose a production database or deployment URL.

## Requirements

- Node.js/npm available locally.
- `clasp` available through `npx --yes @google/clasp ...` or installed globally.
- A valid clasp login for an authorised Google Workspace account:

```sh
npx --yes @google/clasp login
```

## Apps Script Project

Script ID:

```text
<YOUR_SCRIPT_ID>
```

Current web app deployment ID:

```text
<YOUR_DEPLOYMENT_ID>
```

Current deployed version:

```text
@43 - Scanner top-level link handoff
```

Current web app URL:

```text
https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec
```

## Script Properties

Required:

- `SPREADSHEET_ID`

Recommended:

- `WEB_APP_BASE_URL`

Optional:

- `INVENTORY_SHEET_NAME`
- `EXTERNAL_SCANNER_URL` if you host the camera scanner somewhere other than the default GitHub Pages path.

Update Mode authorization:

- `UPDATE_AUTH_ALLOWED_EMAILS`: comma-separated staff emails allowed to change inventory.
- `UPDATE_AUTH_ALLOWED_DOMAINS`: comma-separated domains allowed to change inventory.
- `UPDATE_MODE_PIN_SHA256`: SHA-256 hash of the Update Mode PIN.
- `UPDATE_MODE_PIN_SALT`: salt used when hashing the PIN.
- `UPDATE_AUTH_DISABLED`: explicit development-only escape hatch. Leave unset in production.

Set these from the spreadsheet menu: `D&T Inventory` -> `Set App Config`, or use Apps Script project settings.

If `WEB_APP_BASE_URL` is not configured, the deployed web app now tries `ScriptApp.getService().getUrl()` as a runtime fallback before disabling QR/external shortcuts. Keep the Script Property set in production anyway so generated QR links and Sheet-side tools are explicit and stable.

Update Mode fails closed if no authorization properties are configured. A direct `?mode=tech` URL may show the Update workflow screen, but save/add/remove calls are rejected server-side until the active Google account is allowed or a valid PIN unlock token is supplied.

To create a PIN hash without storing the raw PIN in source, generate a random salt and hash `salt + pin`, for example:

```sh
node -e "const crypto=require('node:crypto'); const salt='replace-with-random-salt'; const pin='replace-with-pin'; console.log(crypto.createHash('sha256').update(salt+pin).digest('hex'))"
```

Store the salt in `UPDATE_MODE_PIN_SALT` and the printed hash in `UPDATE_MODE_PIN_SHA256`. Do not commit the raw PIN, salt used by production, or hash values to Git.

Apps Script web apps may not always expose `Session.getActiveUser().getEmail()` depending on deployment and Workspace settings. Use `UPDATE_AUTH_ALLOWED_EMAILS` / `UPDATE_AUTH_ALLOWED_DOMAINS` where active-user email is available, and configure PIN unlock as the fallback for authorised workshop staff.

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

Local smoke tests:

```sh
node scripts/smoke-tests.mjs
```

These tests do not use live Google Sheets. They cover safe external URL handling, View-first QR URL generation, V++ encoding, placeholder exclusion, and the local fail-closed authorization model.

## Push And Deploy

Push source to Apps Script:

```sh
npx --yes @google/clasp push --force
```

Create/update a deployment only after testing the pushed code. To update the current web app URL:

```sh
npx --yes @google/clasp deploy \
  --deploymentId <YOUR_DEPLOYMENT_ID> \
  --description "D&T QR Inventory deployment"
```

Do not redeploy during development unless the implementation pass is complete and tested.

## Test URLs

The app uses one Apps Script deployment URL. Individual storage pages are in-app routes on that same deployment, using query parameters after `/exec`.

Base app / landing page:

```text
https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec
```

419A view mode:

```text
https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec?room=419A&loc=<storage-id-or-location>
```

419A Update Mode:

```text
https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec?room=419A&loc=<storage-id-or-location>&mode=tech
```

`mode=tech` is retained as the internal/backward-compatible route parameter, but the UI labels this workflow as `Update`.

QR labels should normally encode the View URL, not the Update URL. The printed QR opens the correct storage page in read-only View Mode; authorised users can enter Update Mode from inside the app.

Manual authorization QA:

1. Open a storage View URL and confirm no edit/add/remove controls are visible.
2. Open the same route with `&mode=tech`.
3. If Update authorization is not configured, confirm save/add/remove are blocked with a configuration message.
4. If an allowed Google account is configured, confirm Update Mode unlocks for that account and mutation calls succeed only after server authorization.
5. If PIN is configured, confirm an incorrect PIN fails and a correct PIN unlocks the session.
6. Confirm QR labels still open View Mode by default.

## Brother QL-1110 Label Printing

The QR Labels page includes browser-print presets for the Brother QL-1110 / QL-1110NWB 102mm direct thermal printer. Apps Script cannot silently print to USB, Bluetooth, Ethernet, or Wi-Fi printers from HtmlService; the app generates print-ready layouts and the user prints through the normal Brother driver, AirPrint, or OS print dialog.

Use:

```text
/exec?admin=labels&printer=brother-ql1110
```

Available presets:

- `A4 labels`: two-column office/PDF preview.
- `Brother 90x29`: slim 90mm x 29mm continuous-roll labels for short normal-storage labels.
- `Brother 102x50`: compact 102mm x 50mm continuous-roll labels for normal storage.
- `Brother 102x70 safety`: 102mm x 70mm continuous-roll labels with enough space for hazard wording; recommended for mixed batches and chemical storage.

Each preview label has a `Print this label` button. The page also includes a `Choose one label to print` selector. Either method switches the QR Labels page into single-label mode and opens the print dialog so only that one label is printed. Use `Show all labels` to return to the full batch.

Recommended Brother driver settings:

- Select the Brother QL-1110 / QL-1110NWB printer.
- Select the matching continuous roll size, such as `90mm x 29mm`, `102mm x 50mm`, or `102mm x 70mm`.
- Set scaling to `100%` / actual size.
- Set margins to `None`; Chrome's default margins can crop the QR code on small labels.
- Disable browser headers and footers.
- Print a small sample before a batch.
- QR labels still open View Mode by default; Update remains an explicit in-app action.

## In-App QR Scanner

The mobile landing page and top bars include a `Scan QR` action. When `EXTERNAL_SCANNER_URL` is configured, visible Scan QR controls render as real links with `target="_top"` to the standalone HTTPS scanner, so camera permission is requested outside the Apps Script iframe without relying on a popup/new-tab handoff. The in-app scanner modal remains as a fallback for manual entry, QR photo upload, and pasted QR URLs when no external scanner is configured. The scanner uses the browser camera over HTTPS and tries the native `BarcodeDetector` API first. Where native QR detection is unavailable, it loads `jsQR` from the jsDelivr CDN (`https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js`) as a lightweight fallback.

Camera access is not hosted by Apps Script. Google Apps Script HtmlService runs inside a browser frame, and some browsers or device policies block `getUserMedia()` even on the deployed HTTPS `/exec` URL. The app therefore opens a small top-level HTTPS scanner page first. The default scanner URL is:

```text
https://sunnydesigntech.github.io/dt-qr-inventory-system/scanner/
```

The scanner receives the active `/exec` URL as a `target` query parameter, scans the QR code, strips any `mode=tech`, and redirects back to the inventory app in View Mode. If you fork or self-host, set `EXTERNAL_SCANNER_URL` to your own HTTPS scanner page.

For no-camera QA, open the scanner with `selftest=1` and a `target` URL. This does not auto-redirect; it renders the parsed room, location, whether `mode=tech` was stripped, and the final View URL. Use this before staff training or rollout replay to prove that pasted/scanned QR inputs resolve safely even when browser camera permission is blocked:

```text
https://sunnydesigntech.github.io/dt-qr-inventory-system/scanner/?target=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2FYOUR_DEPLOYMENT_ID%2Fexec&selftest=1
```

If camera access is still blocked, the scanner remains usable through fallbacks:

1. Open the camera scanner, tap `Start camera`, and allow camera permission.
2. Scan the printed QR label with the phone's native Camera app and open the `/exec?room=...&loc=...` link directly.
3. Paste a copied QR URL into the scanner modal.
4. Enter the room and Location Code/storage route manually.
5. Upload or take a QR image for local `jsQR` decoding where supported.

Scans, uploaded QR images, and pasted QR URLs open View Mode by default. If a QR URL contains `mode=tech`, the scanner strips that route mode; Update Mode still requires an explicit in-app action plus the configured server-side Update authorization.

Browser tips:

- Use the deployed HTTPS `/exec` URL, not an editor preview.
- Grant camera permission if prompted.
- On iPhone/iPad, test Safari and Chrome, but expect Apps Script iframe limitations.
- If the first scanner modal reports the Apps Script frame is blocking camera, use `Open full-screen camera scanner` before trying `Start camera`.
- If the camera remains blocked, use native Camera scanning, pasted QR URLs, or manual room/location entry.

The static scanner source lives in `scanner/index.html`. The public scanner is published from the repository's `gh-pages` branch at the GitHub Pages URL above. The scanner does not contain Spreadsheet IDs, Apps Script deployment IDs, PINs, or private data.

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

## Standalone Runtime vs Bound Sheet Script

`code.gs` is the standalone web app runtime pushed by clasp with `index.html`, `app_styles.html`, `app_script.html`, and `appsscript.json`.

`sheet_admin/InventoryAdmin.gs` is a separate bound spreadsheet control-plane script for the live Google Sheet menu. Changes in one project do not automatically update the other. Keep shared workflow behavior aligned deliberately, and do not assume a bound-script menu update changes the deployed web runtime.

Sheet-provided external links such as Purchase Link and SDS Link are rendered clickable only when they use `http://` or `https://`. Unsupported or unsafe schemes are shown as blocked/non-clickable text.

## 419A Rollout Workflow

For role-specific operating instructions and the go-live gate, use [ROLLOUT_CHECKLIST.md](ROLLOUT_CHECKLIST.md).

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
- `Storage_Master` includes clickable `Open View` and `Open Update` links, QR link/image columns, item/chemical/attention counts, and placeholder-only status so it can be used as an operational navigation sheet.
- `QR_Labels`: printable QR label rows with View URL, Update URL, QR formula, hazard label text where relevant, and Brother QL-1110 label-size guidance.
- `Inventory_Readiness_Report`: critical errors, warnings, QR readiness, 419A rollout status, duplicate checks, chemical safety note checks, reorder-level checks, and maintenance-detail checks.
- `Audit_Log`: appended automatically when Update Mode saves quantity/status changes, adds a new item, or removes an item. New rows include the active route/location code where available.

## Bound Sheet Admin Script

The standalone web app project is authoritative for the deployed web runtime. The live dashboard Google Sheet may also need a small bound Apps Script project so the spreadsheet menu shows the same `D&T Inventory` admin tools.

The bound script source is:

- `sheet_admin/InventoryAdmin.gs`

This file is not pushed by the standalone `.claspignore` rules. To install or refresh the live Sheet menu:

1. Open the live dashboard Google Sheet while signed in with an authorised Google Workspace account.
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
  --deploymentId <YOUR_DEPLOYMENT_ID> \
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

### 2026-05-11 08:27 HKT

- Version: `@43`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: changed visible Scan QR controls to render as real `target="_top"` links to the standalone HTTPS scanner when `EXTERNAL_SCANNER_URL` is configured. This gives the browser a user-activated top-level navigation path out of the Apps Script iframe before camera permission is requested, while retaining the in-app scanner modal as fallback when no external scanner is configured.
- Rollback note: version `@42` remains the same-tab JavaScript scanner handoff fallback.

### 2026-05-11 08:21 HKT

- Version: `@42`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: changed the Scan QR handoff from popup/new-tab opening to same-tab navigation to the standalone HTTPS scanner when `EXTERNAL_SCANNER_URL` is configured. This removes popup-blocker/new-tab friction before camera permission is requested and keeps QR scans View-first.
- Rollback note: version `@41` remains the external scanner first fallback.

### 2026-05-11 08:16 HKT

- Version: `@41`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: changed the Scan QR action to open the standalone HTTPS scanner first when `EXTERNAL_SCANNER_URL` is configured, so camera permission is requested outside the Apps Script iframe. The in-app scanner remains the fallback for popup-blocked cases, QR photo upload, pasted QR URLs, and manual room/location entry.
- Rollback note: version `@40` remains the runtime web app URL fallback for scanner.

### 2026-05-07 10:24 HKT

- Version: `@40`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: added a runtime `ScriptApp.getService().getUrl()` fallback for `WEB_APP_BASE_URL` so the external scanner receives a proper `/exec` target even when the Script Property is missing. Production should still set `WEB_APP_BASE_URL` explicitly before label printing.
- Rollback note: version `@39` remains the external HTTPS scanner handoff fallback.

### 2026-05-07 10:06 HKT

- Version: `@39`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: changed live camera scanning from an Apps Script-frame feature into an external top-level HTTPS scanner handoff. The app now passes the active `/exec` URL as a safe redirect target to the static scanner page. The scanner strips Update routes and redirects only to View Mode.
- Rollback note: version `@38` remains the full-screen Apps Script-frame scanner fallback.

### 2026-05-06 14:35 HKT

- Version: `@38`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: added a tappable fallback link for the full-screen scanner in case Safari/Chrome blocks the scanner popup. The scanner still strips `mode=tech` from scanned/pasted URLs and keeps QR labels View-first.
- Rollback note: version `@37` remains the initial full-screen scanner launch fallback.

### 2026-05-06 14:32 HKT

- Version: `@37`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: added a full-screen scanner launch path for Apps Script camera blocking. When the scanner detects it is running inside the Apps Script frame, it now offers `Open full-screen camera scanner`; the opened top-level scanner page then asks for camera permission. View-first QR routing, manual room/location entry, pasted QR URL handling, QR image decoding, and Update Mode authorization remain unchanged.
- Rollback note: version `@36` remains the scanner photo-capture fallback.

### 2026-05-05 09:57 HKT

- Version: `@28`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: tightened the Brother `90mm x 29mm` preset after Chrome/Brother preview still showed the label overflowing and producing a second blank label. The slim preset now uses a 76mm x 22mm border-box content area inside the physical media, smaller QR/text sizing, no forced per-label page break, and zero-margin print CSS.
- Rollback note: version `@27` remains the previous 90x29 label fit fallback.

### 2026-05-05 09:14 HKT

- Version: `@27`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: fixed the 90mm x 29mm print layout after Brother/Chrome preview showed cropped QR output and an extra blank sheet. The slim preset now uses a safer 80mm x 24mm printable area inside the 90mm x 29mm page, smaller QR/text sizing, no forced trailing blank page, clearer print-dialog guidance for `Margins: None`, and a top-level single-label selector.
- Rollback note: version `@26` remains the initial single-label/90x29 print fallback.

### 2026-05-05 09:04 HKT

- Version: `@26`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: refined QR label printing so each preview label has a `Print this label` action for single-label replacement/sample printing, added a `Brother 90x29` slim thermal preset, and kept `Brother 102x50` / `Brother 102x70 safety` presets for larger labels. Sheet-generated `QR_Labels` guidance now notes 90mm x 29mm as an option for short non-chemical labels while keeping chemical labels on the taller safety preset.
- Rollback note: version `@25` remains the Brother 102mm print-preset fallback.

### 2026-05-04 20:39 HKT

- Version: `@25`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: added Brother QL-1110 / QL-1110NWB print support to the QR Labels web page with `A4`, `Brother 102x50`, and `Brother 102x70 safety` presets. The Brother presets generate 102mm thermal-label layouts, inject matching print CSS for browser print, keep QR labels View-first, and document driver settings for 100% scale/sample printing. `QR_Labels` sheet generation now includes Brother label-size guidance and printer notes.
- Rollback note: version `@24` remains the phone camera permission and large-text fallback.

### 2026-05-04 19:45 HKT

- Version: `@24`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: phone usability fix after live testing found camera permission was not opening reliably and mobile text still read too small. Scanner now opens with an explicit Start camera action so the browser permission prompt is tied to a user tap, includes clearer Safari/Chrome and manual Location Code fallback guidance, and the mobile typography floor was raised again across landing, storage cards, location pages, Update Mode, scanner, QR labels, and admin pages.
- Rollback note: version `@23` remains the Figma dashboard hierarchy fallback.

### 2026-05-03 19:10 HKT

- Version: `@23`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: refined the dashboard hierarchy using the referenced Figma SaaS/financial dashboard patterns. Added icon-led metric cards, a desktop workshop command panel for Scan QR / Storage Master / QR Labels / Readiness, and tighter desktop/tablet grouping while preserving the mobile scanner/search-first layout and large phone-readable text.
- Rollback note: version `@22` remains the Figma-inspired dashboard card fallback.

### 2026-05-03 14:25 HKT

- Version: `@22`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: Figma-inspired dashboard UI polish based on the referenced SaaS and financial dashboard community files. Applied a softer light dashboard canvas, elevated white cards, primary/cyan accent tokens, clearer metric/card shadows, rounded route identity blocks, stronger action styling, and desktop hover polish while preserving the mobile scanner-first workflow and large phone-readable typography from `@21`.
- Rollback note: version `@21` remains the mobile large-text readability fallback.

### 2026-05-03 13:19 HKT

- Version: `@21`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: mobile large-text readability polish after the first readability pass was still too small on a real phone. Added a final mobile typography floor for operational text, larger Location Code / route badges, larger storage and item card titles, larger scanner/admin/QR label text, and reduced dashboard density where needed instead of shrinking content.
- Rollback note: version `@20` remains the previous mobile readability fallback.

### 2026-05-03 00:03 HKT

- Version: `@20`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: mobile readability polish. Added mobile typography tokens and responsive overrides so the phone UI uses readable title, body, metadata, badge, button, form, route-code, scanner, item card, QR label, and admin text sizes without pinch zoom. The scan/search-first layout remains compact by reducing density instead of shrinking operational text.
- Rollback note: version `@19` remains the mobile scanner and UX redesign fallback.

### 2026-05-02 23:37 HKT

- Version: `@19`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: mobile scanner and UX redesign. The phone landing page is now scan/search/storage-first, workshop tools are collapsed on mobile, storage cards are larger and Location Code-led, top bars include a scanner shortcut, and the scanner supports camera QR detection with manual Location Code / QR URL fallback.
- Rollback note: version `@18` remains the mobile UI fit fallback.

### 2026-05-02 17:32 HKT

- Version: `@18`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: mobile UI fit polish for the live template runtime. Phone layouts now stack Update controls, Add Item, save bar, item cards, storage cards, QR labels, and Storage Master rows more predictably at 320-390px widths, with safer tap targets and reduced horizontal overflow risk.
- Rollback note: version `@17` remains the fallback for the real-use add/remove and clickable Storage_Master workflow.

### 2026-05-01 20:53 HKT

- Version: `@17`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: real-use inventory workflow polish. `Storage_Master` now acts as a clickable operations map with `Open View` and `Open Update` links, QR image support, item/chemical/attention counts, and placeholder-only status. Update Mode add/remove flow now shows current storage identity, safer remove confirmation, Test category support, and route/location-code metadata in new `Audit_Log` rows.
- Rollback note: version `@16` remains the fallback for the first add/remove implementation.
- Remaining gate: complete physical sample QR label phone scans and keep full 419A rollout paused until warning decisions and unmatched-row governance are accepted by HoD/technician.

### 2026-05-01 18:36 HKT

- Version: `@15`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: applied the approved Figma UI/UX direction to the live dashboard: clearer one-deployment/many-storage route model, stronger 419A Location Code prominence, View-first QR label copy, clearer Update save/audit language, chemical storage callouts, readiness gate copy, and wider desktop dashboard layout.
- Rollback note: version `@13` remains the stable pilot-readiness fallback; version `@14` was superseded by `@15` after tightening legacy 419A route wording.
- Remaining gate: physical sample QR label scan evidence and final warning/unmatched-row decisions are still required before full 419A rollout.

### 2026-05-01 14:50 HKT

- Version: `@12`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: 419A authoritative Location Code workflow: Room 419A now routes QR/storage pages by the latest `Location Code` values such as `419A-FCU-01` and `419A-CAB-01`, while preserving workbook Storage IDs as metadata. The live Inventory was appended with 53 storage placeholder rows and 40 matched current item rows, including 21 latest Chemical Cabinet 01 chemical rows. Storage_Master, QR links, QR image formulas, QR_Labels, and readiness outputs were regenerated from the live Sheet.
- Rollback note: previous deployment/version `@11` retained in Apps Script version history as fallback.
- Remaining gate: review the remaining readiness warnings, manually reassign unmatched old 419A rows, and sample-scan pilot QR labels before full physical rollout.

### 2026-05-01 12:44 HKT

- Version: `@11`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: Single-deployment QR workflow polish: documented one Apps Script `/exec` URL with route parameters, clarified that QR labels scan to View Mode by default, kept Update Mode as the authorised in-app stock-check workflow, and added the routing explanation to the QR labels UI.
- Rollback note: previous deployment/version `@10` retained in Apps Script version history as fallback.
- Remaining gate: authenticated Google Workspace web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 23:42 HKT

- Version: `@10`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: UI demo operations screens wired into the active template runtime: landing operations tiles, Storage Master view, Audit Log view, Low Stock/Reorder view, Maintenance/Safety view, and clarified single-deployment QR routing.
- Rollback note: previous deployment/version `@9` retained in Apps Script version history as fallback.
- Remaining gate: authenticated Google Workspace web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 22:51 HKT

- Version: `@9`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: Workshop operations workflow hardening: Update wording, category filter, extended operations columns, Storage_Master generation, Audit_Log save logging, stronger QR/readiness checks, and updated matched database workbook.
- Rollback note: previous deployment/version `@8` retained in Apps Script version history as fallback.
- Remaining gate: authenticated Google Workspace web app QA and live Sheet-bound admin script installation/testing still required before physical QR rollout.

### 2026-04-30 16:03 HKT

- Version: `@8`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: Operational rollout workflow polish and QR labels route fix.
- Rollback note: previous deployment/version `@7` retained in Apps Script version history as fallback.

### 2026-04-30 15:41 HKT

- Version: `@7`
- Deployment ID: `<YOUR_DEPLOYMENT_ID>`
- Deployment URL: `https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec`
- Summary: Production hardening diagnostics and UI polish.
- Rollback note: previous deployment/version `@6` retained in Apps Script version history as fallback.

## Rollback

Use Apps Script deployment history to redeploy an earlier version if a rollout has an issue. The previous live deployment before this pass was version `@10`.
