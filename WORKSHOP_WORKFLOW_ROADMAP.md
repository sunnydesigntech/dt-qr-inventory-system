# D&T QR Inventory System - Workshop Workflow Roadmap

Date: 2026-05-01
Status: Active roadmap for turning the QR inventory app into the D&T workshop operating system. The standalone web app is deployed at `@40`; Room 419A uses the authoritative `Location Code` list for QR/storage routes, with live clickable Storage_Master, QR_Labels, Audit_Log add/remove/update recording, readiness outputs, Figma-inspired dashboard hierarchy polish, phone-fit layout refinements, stronger phone-readable typography, an external top-level HTTPS QR scanner for reliable camera access outside the Apps Script frame plus tappable/manual route fallback, and Brother QL-1110 / QL-1110NWB print presets including single-label and safer 90mm x 29mm slim label support.

## Target Operating Model

Every physical storage point has a Storage ID and QR code. Every inventory item belongs to a storage point. Every scan opens the correct live storage page. Every authorised update writes back to the live Google Sheet and records enough metadata for audit, reorder, maintenance, and safety planning.

The app uses one Apps Script deployment URL. Storage, admin, and update views are in-app routes on that same `/exec` URL, using query parameters such as `?room=419A&loc=419A-FCU-01`. Printed QR labels should scan to View Mode first; Update Mode is entered from inside the app by authorised staff.

## Core Workflow

1. Storage mapping
   - Define every cupboard, tray, cabinet, trolley, rack, chemical cabinet, machine zone, and material storage point.
   - Required identity: `Room`, `Storage ID`, `Storage Label`, `Specific Location`, `Location Code`, `Storage Type`.
   - For Room 419A, `Location Code` is the operational numbering used on phone stocktake pages and printed QR labels, for example `419A-FCU-01` and `419A-CAB-01`.
   - Display room names remain human-readable, for example `V++`; ID-safe names may use forms such as `VPP-ELEC-001`.

2. Item entry
   - Add tools, machines, materials, chemicals, electronics, and consumables under the mapped storage point.
   - Required app fields remain compatible with the original 8-column sheet.
   - Optional operations fields support `Safety Note`, `Reorder Level`, `Supplier`, `Purchase Link`, `Asset Value`, `Maintenance Due`, and `SDS Link`.

3. QR generation and print
   - Run `Prepare App Columns`.
   - Run `Build Storage Master`.
   - Run `Refresh QR Links`.
   - Run `Build QR Label Sheet`.
   - Print and sample-scan labels before physical rollout.

4. Student/staff View Mode
   - Scan QR code.
   - See storage identity, expected items, quantity, status, remarks, and chemical hazard warnings.
   - No edit controls are shown in View Mode.

5. Technician Update Mode
   - Open Update Mode from the storage page.
   - Unlock with an authorised account or configured PIN before changing live data.
   - Search/filter items by text, status, and category.
   - Update quantity and status.
   - Add items physically confirmed in the current storage.
   - Remove items only after confirmation.
   - Save/add/remove writes back to the live Sheet, updates `Last Updated` / `Updated By` where those columns exist, and appends `Audit_Log` rows for actual changes.

6. Admin and HoD workflow
   - Use diagnostics to confirm configuration and columns.
   - Use readiness to clear critical data errors before rollout.
   - Use Storage_Master, QR_Labels, and Inventory_Readiness_Report for rollout, audit, purchasing, maintenance, and safety review.

## Current Implementation Status

Implemented:

- Template-based Apps Script web app.
- Landing dashboard metrics and room filters.
- Storage-aware routing by Specific Location, Storage ID, Storage Label, and Location Code.
- Relative in-app navigation.
- View Mode and Update Mode.
- Sticky save bar, unsaved count, changed-row highlight, decimal quantity support, and save validation.
- Item search and status filter.
- Category filter added in this pass.
- QR links and QR label sheet generation.
- Admin diagnostics, readiness, and QR labels web routes.
- Google Sheet menu support for QR links/images, readiness, imports, and 419A summary.
- Extended operations columns.
- Generated clickable `Storage_Master` with View/Update links and counts.
- Save/add/remove audit logging to `Audit_Log`.
- Update Mode Add item and Remove item workflow with placeholder-route preservation.
- Server-enforced Update Mode authorization for save/add/remove mutations.
- Stale-row item identity checks before update/remove writes.
- Chemical safety note, reorder-level, and maintenance-detail readiness checks.

Partially implemented:

- Full-room consolidation beyond 419A and V++ depends on cleaned source data.
- Low-stock/reorder dashboard exists as readiness signals but not a dedicated purchasing dashboard yet.
- Machine maintenance tracking uses optional columns and readiness warnings, not a full service workflow yet.
- Chemical safety uses hazard badges, safety notes, SDS links, and readiness warnings; expiry/PPE/restricted access fields are future enhancements.

Not yet implemented:

- Edit/archive item from the web UI beyond the current add/remove workflow.
- Dedicated HoD purchasing/budget dashboard.
- Email alerts.
- SDS file attachment workflow.
- Detailed machine service log.
- Stocktake history beyond save audit rows.

## Version Roadmap

### V1 - Working QR Inventory

Stable landing page, storage pages, View Mode, Update Mode, QR link generation, save updates, and chemical warnings.

### V2 - 419A Rollout Ready

Authoritative Storage IDs, placeholder storage rows, readiness report, QR labels, Storage_Master, better mobile UI, and admin diagnostics.

### V3 - Full Workshop Rollout

Scale the same data model to V++, 415B, 415C, machine zones, electronics trays, material racks, and room-level QR labels.

### V4 - Advanced Operations

Audit log, reorder levels, maintenance planning, chemical safety tracking, supplier/purchase metadata, and HoD readiness views.

### V5 - Smart Assistant / Analytics

Low-stock alerts, missing-item trends, breakage history, predicted reorder lists, termly audit reports, and optional AI explanation support. Human staff remain responsible for purchasing, safety, and workshop decisions.

## Rollout Gate

Do not physically roll out labels until:

- diagnostics show no critical configuration error;
- readiness report has no critical errors;
- QR links are populated and sample-scanned;
- QR labels render;
- 419A pages load by Storage ID;
- V++ URLs encode correctly;
- Update Mode authorization is configured and direct `?mode=tech` access alone cannot mutate inventory;
- Update Mode save works on a safe row;
- chemical storage displays hazard styling;
- staff/student View Mode is understandable;
- technicians understand validation and failed-save recovery.
