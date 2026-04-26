# D&T QR Inventory System

Version: V2.1 Progress PRD and Rollout Plan  
Date: 2026-04-27  
Status: Local production-readiness build complete; live deployment still on earlier version `@2`

## 1. Executive Summary

The D&T QR Inventory System is a Google Apps Script web app for QR-based room and storage inventory management across Design & Technology spaces. It uses a Google Sheet as the database and is designed for mobile use by technicians, teaching staff, students, and department administrators.

The system is intended to support a real rollout starting with Room 419A, using authoritative Storage IDs, then expand to other rooms and areas including V++.

The runtime app remains a single Apps Script file:

- `code.gs`

Repository support files are allowed and now exist:

- `.clasp.json`
- `.claspignore`
- `appsscript.json`
- `README.md`
- `LICENSE`
- `PRD_PROGRESS.md`

## 2. Current Release State

### Live Apps Script Deployment

Current live deployment is still the earlier version:

- Deployment version: `@2`
- Deployment ID: `AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ`
- Web app URL: `https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec`

### Local Repo State

The local repo contains substantial production-readiness improvements that have not yet been pushed or deployed after the latest validation/data-consolidation work.

Important note:

- `clasp status` shows only `appsscript.json` and `code.gs` as Apps Script push files.
- Standard `clasp push` previously returned `Skipping push`.
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
5. Enter Technician Mode when authorized.
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
- Server-rendered HTML plus embedded vanilla JavaScript.

### Runtime Rule

All deployed runtime logic remains in:

- `code.gs`

Do not split runtime HTML, CSS, or JS into separate Apps Script files unless the architecture decision changes later.

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
- Technician Mode.
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

Storage lookup supports:

- `Specific Location`
- `Storage ID`
- `Storage Label`
- `Location Code`

### Navigation

Implemented:

- `All Locations`
- `View Mode`
- `Technician Mode`
- Relative in-app links that do not require `WEB_APP_BASE_URL`

`WEB_APP_BASE_URL` is required only for:

- QR link generation.
- QR label sheets.
- external absolute links.

### Technician Mode

Implemented:

- Editable quantity input.
- Non-negative numeric validation.
- Decimal quantity support.
- Editable status dropdown.
- Save button loading/disabled state.
- Success/failure notices.
- Post-save re-render.
- Bridge warning outside deployed Apps Script context.

### Location Page Filters

Implemented locally:

- Item search.
- Status filter.
- Item count.
- Chemical count.
- Attention count.

### Admin Menu

Current menu tools in local `code.gs`:

- `Refresh QR Links`
- `Refresh QR Images (Optional)`
- `Prepare App Columns`
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

Note: the current implementation should be reviewed to confirm placeholder output fields exactly match this convention before final deployment.

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

Current report generation exists locally, but should be validated against the consolidated live database before deployment.

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
- Tech URL
- QR Image Formula
- Print Label Text

Current implementation creates QR label rows locally. Before final rollout, verify that it includes both View and Tech URL fields exactly as required.

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
- Tech URL
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

- PIN-protected Technician Mode.
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
