# D&T QR Inventory System

Version: V2.5 Progress PRD and Workshop Workflow Plan
Date: 2026-05-01
Status: Template-based Apps Script runtime deployed through `@18`; live 419A Location Code routing, QR generation, clickable Storage_Master, QR_Labels, Audit_Log add/remove/update recording, mobile UI fit polish, and readiness outputs have been regenerated from the dashboard Sheet workflow. Full physical QR rollout remains gated by warning review, manual reassignment of unmatched old rows, and sample scans.

## 1. Executive Summary

The D&T QR Inventory System is a Google Apps Script web app for QR-based room and storage inventory management across Design & Technology spaces. It uses a Google Sheet as the database and is designed for mobile use by technicians, teaching staff, students, and department administrators.

The system is intended to support a real rollout starting with Room 419A, using authoritative Storage IDs, then expand to other rooms and areas including V++.

The runtime app now uses Apps Script `HtmlService` templates so the approved prototype UI can run against live Google Sheets data:

- `code.gs` for backend logic, admin tools, imports, QR generation, diagnostics, and save validation.
- `index.html` for the app shell.
- `app_styles.html` for the prototype-derived design system CSS.
- `app_script.html` for the prototype-derived plain JavaScript UI.

Repository support files are allowed and now exist:

- `.clasp.json`
- `.claspignore`
- `appsscript.json`
- `README.md`
- `LICENSE`
- `PRD_PROGRESS.md`
- `prototype/`

## 2. Current Release State

### Live Apps Script Deployment

Current documented production deployment:

- Deployment version: `@18`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Web app URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`

### Local Repo State

The local repo contains the template-based runtime, UI demo operations screens, documentation, matched-workbook updates, 419A Location Code routing polish, Figma-approved dashboard UI refinements, clickable Storage_Master generation, real-use Add/Remove item workflow hardening, and mobile UI fit polish. The standalone web runtime was pushed and deployed as Apps Script version `@18`.

Important note:

- `clasp status` now shows the Apps Script runtime files: `app_script.html`, `app_styles.html`, `appsscript.json`, `code.gs`, and `index.html`.
- Forced push should be used only after database consolidation and validation:

```bash
npx --yes @google/clasp push --force
```

## 3. Product Vision

The system should let a user standing in front of a storage unit:

1. Scan a QR code.
2. Open the exact storage page.
3. View expected items.
4. See stock quantity, status, remarks, and hazard indicators.
5. Enter Update Mode when authorized.
6. Update quantity and status safely.
7. Keep the live dashboard database accurate.

The system must work for both mature inventory sheets and incomplete rollout sheets where some storage units exist before item-level data is entered.

## 4. Primary Users

### Technician

Needs to:

- Open storage by QR.
- Update stock quantity.
- Update status.
- identify low stock, missing items, maintenance needs, and hazards.

### Teaching Staff / Students

Needs to:

- Open read-only storage pages.
- Confirm expected equipment or materials.
- Identify missing or hazardous items.

### Head of Department / Admin

Needs to:

- Maintain central inventory visibility.
- Import and consolidate source records.
- Generate QR links and QR label sheets.
- Run diagnostics and readiness reports.
- Prepare rollout by room.

## 5. Architecture

### Platform

- Google Apps Script standalone web app.
- Google Sheets database.
- clasp deployment from the repo.
- Tailwind CSS via CDN.
- Server-built bootstrap payload from Google Sheets.
- Client-rendered mobile-first prototype UI without runtime JSX/Babel compilation.

### Runtime Rule

All deployed runtime behavior is split across:

- `code.gs`
- `index.html`
- `app_styles.html`
- `app_script.html`

The earlier single-file runtime rule has been superseded by the prototype UI integration decision.

### Spreadsheet Access

Runtime access uses:

```javascript
SpreadsheetApp.openById(SPREADSHEET_ID)
```

The web app should not depend on `getActiveSpreadsheet()` for runtime browsing.

## 6. Data Model

### Required Columns

The app must work with the original 8-column format:

1. `Item ID`
2. `Item Name`
3. `Room`
4. `Specific Location`
5. `Qty`
6. `Category`
7. `Status`
8. `QR Code Link (Auto-Generated)`

### Optional Recommended Columns

The app now supports and recommends:

9. `Unit`
10. `Remarks`
11. `Location Code`
12. `Storage ID`
13. `Storage Label`
14. `QR Code Image`

### Optional Operations Columns

The current local pass extends `Prepare App Columns` with operations fields for a full D&T workshop workflow:

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

These remain optional so the original 8-column sheet still works. When present, saves update timestamp/user metadata and readiness checks can flag chemical safety, reorder, and maintenance issues.

### Status Values

Allowed values:

- `Good`
- `Low Stock`
- `Missing`
- `Needs Maintenance`

### Hazard Detection

Chemical styling is triggered when `Category` is:

- `Chemical`
- `Chemicals`

Matching is case-insensitive.

## 7. Current Implemented Capabilities

### Core Web App

Implemented in `code.gs`:

- `doGet(e)` web app routing.
- Landing page.
- Grouped room/location directory.
- Client-side location search.
- View Mode.
- Update Mode.
- Async saves using `google.script.run`.
- Save validation.
- Storage-aware routing and matching.
- Empty storage states.
- Hazard styling.
- Status badges.
- Mobile-friendly layout.
- Visible configuration and runtime errors.

### Routing

Supported routes:

- `/exec`
- `/exec?room=...&loc=...`
- `/exec?room=...&loc=...&mode=tech`

The app has one Apps Script deployment URL. These are routes on the same `/exec` web app, not separate Apps Script apps. QR labels should scan to the View route by default; authorised users can enter Update Mode from inside the app.

Storage lookup supports:

- `Specific Location`
- `Storage ID`
- `Storage Label`
- `Location Code`

### Navigation

Implemented:

- `All Locations`
- `View Mode`
- `Update Mode`
- Relative in-app links that do not require `WEB_APP_BASE_URL`

`WEB_APP_BASE_URL` is required only for:

- QR link generation.
- QR label sheets.
- external absolute links.

### Update Mode

Implemented:

- Editable quantity input.
- Non-negative numeric validation.
- Decimal quantity support.
- Editable status dropdown.
- Save button loading/disabled state.
- Success/failure notices.
- Post-save re-render.
- Bridge warning outside deployed Apps Script context.
- `Last Updated` and `Updated By` metadata updates when those optional columns exist.
- `Audit_Log` append for actual quantity/status changes.

### Location Page Filters

Implemented locally:

- Item search.
- Status filter.
- Category filter.
- Item count.
- Chemical count.
- Attention count.

### Admin Menu

Current menu tools in local `code.gs`:

- `Refresh QR Links`
- `Refresh QR Images (Optional)`
- `Prepare App Columns`
- `Build Storage Master`
- `Build QR Label Sheet`
- `Create Readiness Report`
- `419A Readiness Summary`
- `Import 419A Storage Master`
- `Import 419A App Load Ready`
- `Open Web App`
- `Set App Config`
- `Set WEB_APP_BASE_URL`
- `Config Status / Diagnostics`

Menu entrypoints have been changed to directly callable public function names.

## 8. Admin / Data Operations

### Prepare App Columns

Purpose:

- Add missing required/recommended columns.
- Avoid duplicate columns.
- Preserve existing sheet data.

Expected result:

- all required columns exist.
- optional rollout columns exist.
- existing rows remain intact.

### Import 419A Storage Master

Purpose:

- Import authoritative 419A storage units as placeholder rows.
- Support empty QR/location pages before item data exists.

Supported source tabs:

- `419A_Storage_Master`
- `419A Storage Master`
- `419AStorageMaster`
- `Room_QR_Label_Plan`
- `RM 419A 2026`

Important merge behavior:

- appends only missing storage placeholders.
- skips duplicates.
- checks all storage identities:
  - Specific Location
  - Storage ID
  - Storage Label
  - Location Code

Placeholder rows are intended to be clearly distinguishable:

- Category: `Storage`
- Status: `Good`
- Qty: `0`
- Remarks: `Placeholder row for QR/location page`

Current implementation now writes this placeholder convention and sets `Is Placeholder` when the column exists.

### Import 419A App Load Ready

Purpose:

- Import cleaned item rows from 419A app-load-ready tabs.
- Enrich rows using 419A storage master data where possible.
- Skip duplicate item identities.

Supported source tabs:

- `419A_App_Load_Ready`
- `419A App Load Ready`
- `419AAppLoadReady`

### Create Readiness Report

Purpose:

- Generate an `Inventory_Readiness_Report` sheet.
- Surface critical data issues before QR rollout or deployment.

Checks include:

- missing required fields.
- invalid quantity.
- invalid status.
- missing room.
- missing specific location.
- missing 419A Storage IDs.
- duplicate item identities.
- duplicate storage placeholders.
- QR link issues.
- V++ encoding risks.
- chemical rows without remarks or safety notes.
- reorder threshold warnings.
- maintenance rows without maintenance detail.

Current report generation exists locally, but should be validated against the consolidated live database before deployment.

### Build Storage Master

Purpose:

- Generate a `Storage_Master` sheet from the live `Inventory` tab.
- Provide the operational room/storage map for rollout checks.

Expected fields:

- Storage ID
- Room
- Storage Label
- Specific Location
- Location Code
- Storage Type
- QR Link
- QR Image
- Status
- Notes

The generated sheet is derived from `Inventory`; it should not replace `Inventory` as the source of truth.

### Build QR Label Sheet

Purpose:

- Generate a `QR_Labels` sheet.
- One row per unique storage/location.
- Include QR links and image formulas.

Expected fields:

- Room
- Specific Location
- Storage ID
- Storage Label
- Location Code
- View URL
- Update URL
- QR Image Formula
- Print Label Text

Current implementation creates QR label rows locally and labels the edit route as `Update URL` for clearer staff-facing language.

## 9. Source Data Landscape

Two source workbooks must be consolidated into the live dashboard database.

### Source 1

File:

- `dt_inventory_419A_authoritative_storage_ids-2.xlsx`

Converted Google Sheet found:

- Name: `dt_inventory_419A_authoritative_storage_ids`
- ID: `1Oz6iojJbRXG1b7KVV-t721Gcc_fqCA8erOcKJvZmGlM`

Important tabs:

- `419A_Storage_Master`
- `419A_Current_Items_Remapped`
- `419A_Unmatched_Old_Locations`
- `419A_Item_Input_Template`
- `419A_App_Load_Ready`

Known extracted facts:

- 53 authoritative 419A storage IDs.
- 35 app-load-ready item rows.

### Source 2

File:

- `DT items in room 419 V++-2.xlsx`

Converted Google Sheet found:

- Name: `DT items in room 419 V++`
- ID: `1-3frb9txIDaSVzWJFoMc8hiKoWi2l3-fwkvWocNQQgM`

Important tabs:

- `Room_QR_Label_Plan`
- `RM 419A 2026`
- `Items in room V++`
- `Sheet11`
- `Chemical in Design Department`
- `items in room 419A_Curtis`
- other legacy room/item sheets

Known extracted facts:

- `Room_QR_Label_Plan` and `RM 419A 2026` contain 53 Room 419A storage IDs.
- V++ data exists and must preserve display room name `V++`.

## 10. Live Dashboard Database Status

Likely live dashboard database found:

- Name: `D&T QR Inventory Database - 2026-03-25 13:30:33`
- ID: `1GqK9XsPdTiPREhVXLeexreZ7cCfZJJ7FNueotpSZpqM`

Expected target tab:

- `Inventory`

Current blocker:

- The available OAuth access can list the file in Drive but cannot read/export its contents.
- Attempted Drive export returned `403 appNotAuthorizedToFile`.
- Attempted Sheets API access returned `403 SERVICE_DISABLED`.

Therefore, live database consolidation has not yet been performed.

## 11. Data Consolidation Plan

### Priority

The dashboard database is now the priority.

Before any deployment:

1. Confirm live spreadsheet access.
2. Confirm or prepare `Inventory` headers.
3. Merge source workbooks into the live database.
4. Generate readiness report.
5. Clear critical errors.
6. Generate QR links and labels.
7. Verify web app against consolidated data.

### Merge Order

1. Import 419A authoritative storage master.
2. Import cleaned 419A app-load-ready rows.
3. Import or map relevant Room 419 / V++ item data.
4. Enrich 419A rows with Storage ID, Location Code, and Storage Label.
5. Preserve V++ display room exactly as `V++`.

### Merge Rules

- Do not overwrite existing live rows unless explicitly safe.
- Skip exact duplicates.
- Detect duplicates using:
  - Room
  - Specific Location
  - Storage ID
  - Storage Label
  - Location Code
  - Item ID
  - Item Name
- Treat 419A Storage ID as authoritative when present.
- Preserve legacy Specific Location for reference.
- Use placeholder rows only for storage pages that need QR/location access.
- Avoid duplicate placeholders when legacy location rows already exist.

## 12. Deployment Plan

Deployment is paused until data consolidation is complete.

When ready:

```bash
node --check --input-type=commonjs < code.gs
npx --yes @google/clasp status
npx --yes @google/clasp push --force
```

Only if forced push actually uploads:

```bash
npx --yes @google/clasp version "Dashboard database + 419A/V++ rollout"
npx --yes @google/clasp deploy --versionNumber <VERSION_NUMBER> --description "Dashboard database + 419A/V++ rollout"
```

Do not delete deployment/version `@2`; keep it as fallback.

## 13. Current Validation Status

Passed locally:

- Apps Script syntax check.
- `clasp status`.
- menu function reference scan.
- duplicate function scan.
- no runtime references to local `.xlsx` paths.
- mocked admin flow tests.
- UI HTML smoke tests.
- 8-column compatibility smoke tests.

Not yet complete:

- live database read/write verification.
- source workbook consolidation into live database.
- readiness report against consolidated live database.
- QR label generation against consolidated live database.
- deployed app smoke tests against consolidated data.
- final push/version/deploy.

## 14. Known Risks / Gaps

### Data Access Blocker

The live Google Sheet could not be exported or read with current OAuth access.

Needed:

- grant content access to the clasp OAuth app/user, or
- provide an accessible exported live dashboard workbook, or
- enable/use a Sheets API-capable credential, or
- run the import/admin tools directly inside the bound Google Sheet UI.

### Source Mapping Complexity

The second workbook contains legacy, semi-structured sheets. Some item rows require interpretation and cleanup before import, especially:

- `Items in room V++`
- `Chemical in Design Department`
- `items in room 419A_Curtis`
- older mixed room tabs

### QR Label Details

The QR label sheet should be verified to include:

- View URL
- Update URL
- QR image formula
- print label text

If the current local implementation does not include all fields exactly, adjust before release.

## 15. Roadmap

### Immediate Next Steps

1. Resolve live database access.
2. Consolidate both source workbooks into the live `Inventory` tab.
3. Run readiness report.
4. Fix critical errors.
5. Generate QR links and QR labels.
6. Verify 419A, V++, chemical, and empty storage workflows.
7. Push with `clasp push --force`.
8. Version and deploy.

### Short-Term Enhancements

- PIN or authorised role protection for Update Mode.
- Add/edit item workflow inside the app.
- QR label print layout improvements.
- Import preview/dry-run summary before committing rows.
- Room-by-room rollout controls.

### Longer-Term Enhancements

- Audit/history log.
- SDS/PDF links for chemicals.
- Dashboard analytics.
- Low-stock alerts.
- Per-room admin pages.
- Barcode scanning support.
- QR label batch export/print workflow.

## 16. Go / No-Go Summary

Current state:

- Code readiness: mostly go.
- Deployment readiness: no-go until database consolidation is done.
- Data readiness: blocked by live database access.

Final release should not proceed until the live dashboard database is consolidated, readiness report critical errors are cleared, and app workflows are verified against the consolidated data.
