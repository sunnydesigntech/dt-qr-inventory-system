# D&T QR Inventory Rollout Checklist

This checklist is for the live Google Apps Script deployment and the live dashboard spreadsheet. The current production deployment is version `@40`; version `@39` is retained in Apps Script version history as the external HTTPS scanner handoff fallback.

Production URL:

```text
https://script.google.com/macros/s/<YOUR_DEPLOYMENT_ID>/exec
```

## Student / Staff Workflow

1. Scan the QR label on a cupboard, trolley, cabinet, tray area, or storage zone.
2. Confirm the page shows the correct room and storage identity.
3. Read the item cards to check the expected contents, quantity, category, and status.
4. Look for Low Stock, Missing, or Needs Maintenance badges before using the storage.
5. Treat Chemical badges and chemical warning panels as safety-critical.
6. Use All Locations only when browsing or checking another storage area.
7. Report missing, low-stock, damaged, or unexpected items to the technician or teacher.

Students and teaching staff should use View Mode only. They should not use Update Mode unless specifically authorised.

## Technician / Update Workflow

1. Open the storage QR page.
2. Use the in-app Update button, or add `&mode=tech` to the storage URL for the same storage route.
3. Unlock Update Mode with an authorised Google account or configured PIN. The URL alone must not be treated as authorisation.
4. Search or filter items inside the storage page.
5. Update quantity using a non-negative number. Decimal quantities are allowed where appropriate.
6. Set status to one of:
   - Good
   - Low Stock
   - Missing
   - Needs Maintenance
7. Check the changed-row indicator and unsaved change count.
8. Press Save updates.
9. Confirm the green saved message appears and the unsaved count clears.
10. If save fails, check the validation message, correct the row, and retry.
11. Use Add item only for items physically confirmed in the current storage.
12. Use Remove only after checking the confirmation details: item name, item ID, storage route, and current quantity.

Placeholder storage rows exist only to make empty QR/location pages routable. Do not treat placeholder rows as stock items.

## Admin / HoD Workflow

Use the web admin pages for quick checks:

- `/exec?admin=diagnostics`
- `/exec?admin=readiness`
- `/exec?admin=labels`

Use the Google Sheet menu for data-changing admin actions:

Before using the menu, confirm the live Sheet-bound Apps Script is the D&T Inventory admin script from `sheet_admin/InventoryAdmin.gs`, not an unrelated older project/menu.

1. D&T Inventory -> Config Status / Diagnostics
2. D&T Inventory -> Prepare App Columns
3. D&T Inventory -> Import 419A Storage Master, only when converted Google Sheets source tabs are ready
4. D&T Inventory -> Build Storage Master
5. D&T Inventory -> Create Readiness Report
6. Fix critical errors in the `Inventory` tab
7. D&T Inventory -> Refresh QR Links
8. D&T Inventory -> Build QR Label Sheet
9. Print from `QR_Labels`
10. Sample-scan printed QR labels before rollout

Do not run QR printing until the readiness report has no critical errors.

## Data Workflow

The live `Inventory` tab is the source of truth for the app.

Source Excel workbooks are not read by the deployed app at runtime. They must be converted to Google Sheets before import. Generated sheets such as `QR_Labels` and `Inventory_Readiness_Report` are outputs from the live `Inventory` data.

For Room 419A, the latest operational storage numbering comes from the `Location Code` column in `419A_Storage_Master`. QR labels and app routes should use those codes, for example `419A-FCU-01` or `419A-CAB-01`. Workbook `Storage ID` values are preserved as metadata, but the live 419A stocktake workflow should use Location Code as the code printed on storage labels.

`Storage_Master` is generated from `Inventory` and is the operational storage map for rollout checks. It includes clickable `Open View` and `Open Update` links for each storage route. `Audit_Log` is append-only and records Update Mode quantity/status changes, add-item actions, and remove-item actions when the save flow is active.

Do not directly edit:

- Apps Script deployment files outside this repo
- generated QR formula cells unless rebuilding labels manually
- placeholder row identities unless replacing them with real storage/item data
- `Item ID`, `Room`, `Specific Location`, `Storage ID`, `Storage Label`, or `Location Code` without checking QR routes afterward

## QR Label Workflow

1. Confirm `WEB_APP_BASE_URL` points to the active `/exec` deployment.
   - The web app has a runtime fallback, but the Script Property should still be set before label printing so QR links are stable and explicit.
2. Run Refresh QR Links.
3. Run Build QR Label Sheet.
4. Open `QR_Labels`.
5. Confirm one label appears per storage/location.
6. Confirm label text includes room, storage identity, and "Scan to view inventory".
7. Confirm QR images render.
8. Print a small sample first.
9. Scan one 419A label and one V++ label on a phone.
10. Confirm each scan opens View Mode for the correct storage.

The system has one Apps Script web app URL. QR labels do not point to separate apps or separate pages; they point to the same `/exec` URL with route parameters such as `?room=419A&loc=419A-CHEM-001`. The QR code should scan to View Mode by default. Update Mode is available from the in-app Update button for authorised stock checks.

Brother QL-1110 / QL-1110NWB printing is supported through browser print presets on the QR Labels page:

- `/exec?admin=labels&printer=brother-ql1110`
- `Brother 90x29` for slim short-label printing.
- `Brother 102x50` for compact normal-storage labels.
- `Brother 102x70 safety` for mixed/chemical labels with hazard text.

Each label preview includes `Print this label`, and the QR Labels page includes a `Choose one label to print` selector for single-label replacement/sample printing. Use the Brother driver or AirPrint print dialog, choose the matching continuous roll size, set margins to `None`, set scale to 100%, disable browser headers/footers, and print a small sample before any batch. Chrome default margins can crop the QR code on 90mm x 29mm labels.

## Go-Live Gate

419A rollout is ready only when all of these pass:

- diagnostics has no critical configuration error
- readiness report has no critical errors
- 419A storage pages load from direct QR-style URLs
- V++ direct links open correctly with URL encoding
- QR links are populated
- QR labels are generated and sample-scanned
- scanner self-test is run with the active `/exec` target to prove QR parsing, View-only routing, V++ encoding, `mode=tech` stripping, and unsafe URL rejection without relying on camera permission
- tapping `Scan QR` follows a real `target="_top"` link to the standalone HTTPS scanner when `EXTERNAL_SCANNER_URL` is configured; the workflow must not depend on popup/new-tab behaviour or an iframe camera request
- scanner workflow works: full-screen scanner can request camera where browser policy allows, and fallback works when camera remains blocked: native phone Camera scan, pasted QR URL, manual room/location entry, or QR image upload
- external scanner page is online over HTTPS and redirects scanned QR labels back to the same `/exec?room=...&loc=...` View route
- Update Mode authorization is configured and direct `?mode=tech` access alone cannot save/add/remove inventory
- concurrent Update Mode saves are serialized; all changed rows validate before the first write, and no-op saves do not change metadata or append Audit_Log rows
- Update Mode save works on a safe test row
- at least one chemical cabinet page displays hazard styling
- View Mode is understandable to students and teaching staff
- technicians know how to save and recover from failed validation

If any critical item fails, pause rollout and fix it before printing/applying labels.
