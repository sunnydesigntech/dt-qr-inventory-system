# D&T QR Inventory System

Single-file Google Apps Script inventory web app backed by Google Sheets.

## Architecture

Runtime app code must remain in `code.gs`. Do not split the Apps Script runtime into separate HTML, CSS, or JS files.

Repo support files are allowed:

- `appsscript.json` is the Apps Script manifest.
- `.clasp.json` links this folder to the Apps Script project.
- `.claspignore` keeps repo-only files out of Apps Script pushes.
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

Confirm only intended Apps Script runtime files are tracked for push:

```sh
npx --yes @google/clasp status
```

Expected tracked files:

- `appsscript.json`
- `code.gs`

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

Landing page:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec
```

419A view mode:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec?room=419A&loc=<storage-id-or-location>
```

419A technician mode:

```text
https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec?room=419A&loc=<storage-id-or-location>&mode=tech
```

## Admin Menu

The spreadsheet menu exposes:

- Refresh QR Links
- Refresh QR Images (Optional)
- Prepare App Columns
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

Recommended setup order:

1. Run `Prepare App Columns` to add the optional rollout columns if they are missing.
2. Run `Import 419A Storage Master` using a Google Sheet converted from the authoritative 419A workbook. This creates placeholder storage rows so empty storage pages can still render.
3. Run `Import 419A App Load Ready` to append non-duplicate item rows.
4. Run `Create Readiness Report` and fix any errors or warnings.
5. Run `Refresh QR Links`.
6. Run `Build QR Label Sheet` for printable labels.
7. Run `Refresh QR Images (Optional)` only if the inventory sheet has `QR Code Image`.

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

Before pushing, confirm `clasp status` lists only these tracked files:

- `appsscript.json`
- `code.gs`

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
3. Open the same 419A storage in Technician Mode.
4. Confirm item search and status filter work on a populated location.
5. Confirm a valid empty storage page shows the friendly empty message.
6. Run `Refresh QR Links` and spot-check that `V++` and Storage ID URLs are encoded correctly.
7. Run `Build QR Label Sheet` and spot-check a generated QR image/link.

## Rollback

Use Apps Script deployment history to redeploy an earlier version if a rollout has an issue. The previous live deployment before this pass was version `@2`.
