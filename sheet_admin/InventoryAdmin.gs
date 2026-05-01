/**
 * D&T QR Inventory - bound Google Sheet admin menu.
 *
 * This file is intended for the Apps Script project bound to the live
 * dashboard spreadsheet. It does not serve the web app UI; the authoritative
 * web runtime remains the standalone clasp project in this repository.
 */

var DTINV_CONFIG = {
  MENU_NAME: 'D&T Inventory',
  WEB_APP_BASE_URL_PROPERTY: 'WEB_APP_BASE_URL',
  DEFAULT_WEB_APP_BASE_URL: 'https://script.google.com/a/macros/vsa.edu.hk/s/AKfycbyB3esZWpSm0WDydyoJMHw3EtXkag0Qg0WpClSgBcxzaAwUcQk8m-MGJw-uCyfKcptFzQ/exec',
  INVENTORY_SHEET_NAME: 'Inventory',
  STORAGE_MASTER_SHEET_NAME: 'Storage_Master',
  QR_LABEL_SHEET_NAME: 'QR_Labels',
  READINESS_SHEET_NAME: 'Inventory_Readiness_Report',
  ROLLOUT_ROOM: '419A',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemical', 'chemicals'],
  STORAGE_MASTER_SHEET_NAMES: ['419A_Storage_Master', '419A Storage Master', '419AStorageMaster', 'Room_QR_Label_Plan', 'RM 419A 2026'],
  IMPORT_READY_SHEET_NAMES: ['419A_App_Load_Ready', '419A App Load Ready', '419AAppLoadReady'],
  REQUIRED_COLUMNS: [
    ['itemId', 'Item ID'],
    ['itemName', 'Item Name'],
    ['room', 'Room'],
    ['location', 'Specific Location'],
    ['qty', 'Qty'],
    ['category', 'Category'],
    ['status', 'Status'],
    ['qrLink', 'QR Code Link (Auto-Generated)']
  ],
  OPTIONAL_COLUMNS: [
    ['unit', 'Unit'],
    ['remarks', 'Remarks'],
    ['locationCode', 'Location Code'],
    ['storageId', 'Storage ID'],
    ['storageLabel', 'Storage Label'],
    ['qrImage', 'QR Code Image'],
    ['storageType', 'Storage Type'],
    ['lastUpdated', 'Last Updated'],
    ['updatedBy', 'Updated By'],
    ['isPlaceholder', 'Is Placeholder'],
    ['safetyNote', 'Safety Note'],
    ['reorderLevel', 'Reorder Level'],
    ['supplier', 'Supplier'],
    ['purchaseLink', 'Purchase Link'],
    ['assetValue', 'Asset Value'],
    ['maintenanceDue', 'Maintenance Due'],
    ['sdsLink', 'SDS Link']
  ],
  HEADER_ALIASES: {
    itemId: ['item id', 'itemid', 'id', 'asset id'],
    itemName: ['item name', 'itemname', 'items', 'item', 'chemical name', 'machine name'],
    room: ['room', 'room no', 'room number'],
    location: ['specific location', 'location', 'display location', 'location in v++'],
    qty: ['qty', 'quantity', 'number of items', 'unit qty'],
    category: ['category', 'type'],
    status: ['status', 'condition'],
    qrLink: ['qr code link (auto-generated)', 'qr code link', 'qr link', 'view url'],
    unit: ['unit'],
    remarks: ['remarks', 'remark', 'notes', 'note'],
    locationCode: ['location code', 'locationcode'],
    storageId: ['storage id', 'storageid', 'new storage id'],
    storageLabel: ['storage label', 'storagelabel', 'new display location'],
    qrImage: ['qr code image', 'qr image'],
    storageType: ['storage type', 'storagetype', 'storage_type'],
    lastUpdated: ['last updated', 'lastupdated', 'updated at', 'updated date'],
    updatedBy: ['updated by', 'updatedby'],
    isPlaceholder: ['is placeholder', 'isplaceholder', 'placeholder'],
    safetyNote: ['safety note', 'safetynote', 'safety notes', 'hazard note'],
    reorderLevel: ['reorder level', 'reorderlevel', 'minimum stock', 'minimum qty', 'min qty'],
    supplier: ['supplier', 'vendor'],
    purchaseLink: ['purchase link', 'purchaselink', 'purchase url', 'supplier link'],
    assetValue: ['asset value', 'assetvalue', 'unit cost', 'value'],
    maintenanceDue: ['maintenance due', 'maintenancedue', 'service due', 'next service'],
    sdsLink: ['sds link', 'sds', 'safety data sheet', 'safety data sheet link']
  }
};

function onOpen(e) {
  DTInv_onOpen_(e);
}

function onInstall(e) {
  DTInv_onOpen_(e);
}

function DTInv_onOpen_(e) {
  SpreadsheetApp.getUi()
    .createMenu(DTINV_CONFIG.MENU_NAME)
    .addItem('Config Status / Diagnostics', 'DTInv_menuConfigStatus')
    .addSeparator()
    .addItem('Prepare App Columns', 'DTInv_menuPrepareAppColumns')
    .addItem('Build Storage Master', 'DTInv_menuBuildStorageMasterSheet')
    .addItem('Create Readiness Report', 'DTInv_menuCreateReadinessReport')
    .addItem('Refresh QR Links', 'DTInv_menuRefreshQrLinks')
    .addItem('Build QR Label Sheet', 'DTInv_menuBuildQrLabelSheet')
    .addItem('Refresh QR Images (Optional)', 'DTInv_menuRefreshQrImages')
    .addSeparator()
    .addItem('Import 419A Storage Master', 'DTInv_promptImport419AStorageMaster')
    .addItem('Import 419A App Load Ready', 'DTInv_promptImport419AReady')
    .addItem('419A Readiness Summary', 'DTInv_menu419AReadinessSummary')
    .addSeparator()
    .addItem('Open Web App', 'DTInv_menuOpenWebApp')
    .addItem('Set WEB_APP_BASE_URL', 'DTInv_promptSetWebAppBaseUrl')
    .addToUi();
}

function DTInv_menuConfigStatus() {
  var ui = SpreadsheetApp.getUi();
  try {
    var status = DTInv_getDiagnostics_();
    var lines = [
      'Spreadsheet: ' + status.title,
      'Spreadsheet ID: ' + DTInv_maskId_(status.id),
      'Inventory tab: ' + (status.inventorySheetName || 'not found'),
      'Data rows: ' + status.dataRows,
      'WEB_APP_BASE_URL: ' + (status.webAppBaseUrlSource + ' - ' + status.webAppBaseUrl),
      'Missing required columns: ' + (status.missingRequired.length ? status.missingRequired.join(', ') : 'none'),
      'Missing optional columns: ' + (status.missingOptional.length ? status.missingOptional.join(', ') : 'none'),
      'Available sheets: ' + status.sheetNames.join(', ')
    ];
    ui.alert('D&T Inventory Diagnostics', lines.join('\n'), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Diagnostics Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuPrepareAppColumns() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_prepareAppColumns_();
    ui.alert(
      'App Columns Prepared',
      result.initialized
        ? 'Initialized the Inventory header row.'
        : 'Added columns: ' + (result.addedColumns.length ? result.addedColumns.join(', ') : 'none needed'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Prepare App Columns Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuBuildStorageMasterSheet() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_buildStorageMasterSheet_();
    ui.alert('Storage Master Ready', 'Created/updated "' + result.sheetName + '" with ' + result.storageCount + ' storage row(s).', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Build Storage Master Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuCreateReadinessReport() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_createReadinessReport_();
    ui.alert(
      'Readiness Report Ready',
      'Sheet: ' + result.sheetName +
        '\nCritical errors: ' + result.errorCount +
        '\nWarnings: ' + result.warningCount +
        '\n419A storages: ' + result.rolloutStorageCount +
        '\nQR-ready locations: ' + result.qrReadyCount + '/' + result.locationCount,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Readiness Report Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuRefreshQrLinks() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_refreshQrLinks_();
    ui.alert(
      'QR Links Refreshed',
      'Updated rows: ' + result.updatedRows +
        '\nSkipped rows: ' + result.skippedRows +
        '\n419A QR-ready: ' + result.rolloutQrReady + '/' + result.rolloutLocationCount +
        '\nBase URL: ' + result.baseUrl,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Refresh QR Links Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuBuildQrLabelSheet() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_buildQrLabelSheet_();
    ui.alert('QR Label Sheet Ready', 'Created/updated "' + result.sheetName + '" with ' + result.labelCount + ' label row(s).', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Build QR Label Sheet Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuRefreshQrImages() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = DTInv_refreshQrImages_();
    ui.alert('QR Images Refreshed', 'Updated formulas: ' + result.updatedRows + '\nSkipped rows: ' + result.skippedRows, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Refresh QR Images Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menu419AReadinessSummary() {
  var ui = SpreadsheetApp.getUi();
  try {
    var summary = DTInv_get419AReadinessSummary_();
    var lines = [
      '419A storage locations: ' + summary.locationCount,
      'With Storage ID: ' + summary.withStorageId,
      'Missing Storage ID: ' + summary.missingStorageId,
      'QR-ready: ' + summary.qrReady,
      'Item rows: ' + summary.itemRows,
      'Placeholder rows: ' + summary.placeholderRows,
      'Chemical rows: ' + summary.chemicalRows,
      'Attention rows: ' + summary.attentionRows
    ];
    ui.alert('419A Readiness Summary', lines.join('\n'), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('419A Readiness Summary Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_menuOpenWebApp() {
  SpreadsheetApp.getUi().alert('Open Web App', DTInv_getWebAppBaseUrl_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function DTInv_promptSetWebAppBaseUrl() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Set WEB_APP_BASE_URL', 'Paste the active deployed /exec URL:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var url = DTInv_clean_(response.getResponseText()).replace(/[?#].*$/, '');
  if (!/^https:\/\/script\.google\.com\//.test(url)) {
    ui.alert('Invalid URL', 'Paste the deployed Apps Script /exec URL.', ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty(DTINV_CONFIG.WEB_APP_BASE_URL_PROPERTY, url);
  ui.alert('WEB_APP_BASE_URL Saved', url, ui.ButtonSet.OK);
}

function DTInv_promptImport419AStorageMaster() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Import 419A Storage Master',
    'Paste the converted Google Sheet URL or ID containing 419A_Storage_Master, Room_QR_Label_Plan, or RM 419A 2026:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var sourceId = DTInv_extractSpreadsheetId_(response.getResponseText());
  if (!sourceId) {
    ui.alert('Invalid Source', 'Paste a valid Google Sheet URL or spreadsheet ID.', ui.ButtonSet.OK);
    return;
  }
  var confirm = ui.alert(
    'Confirm Import',
    'This appends missing 419A placeholder storage rows and fills blank storage metadata on matching existing rows. Existing Inventory rows are not deleted. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  try {
    var result = DTInv_import419AStorageMasterFromSource_(sourceId);
    ui.alert(
      '419A Storage Master Imported',
      'Source tab: ' + result.sourceSheet +
        '\nImported rows: ' + result.importedRows +
        '\nExisting rows enriched: ' + result.enrichedRows +
        '\nSkipped existing: ' + result.skippedExisting +
        '\nSkipped invalid: ' + result.skippedInvalid,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('419A Storage Import Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_promptImport419AReady() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Import 419A App Load Ready',
    'Paste the converted Google Sheet URL or ID containing 419A_App_Load_Ready / 419A App Load Ready:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var sourceId = DTInv_extractSpreadsheetId_(response.getResponseText());
  if (!sourceId) {
    ui.alert('Invalid Source', 'Paste a valid Google Sheet URL or spreadsheet ID.', ui.ButtonSet.OK);
    return;
  }
  var confirm = ui.alert(
    'Confirm Import',
    'This appends non-duplicate app-load-ready item rows only. Existing Inventory rows are not deleted. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  try {
    var result = DTInv_import419AReadyFromSource_(sourceId);
    ui.alert(
      '419A App Load Ready Imported',
      'Source tab: ' + result.sourceSheet +
        '\nImported rows: ' + result.importedRows +
        '\nSkipped duplicates: ' + result.skippedDuplicates +
        '\nSkipped invalid: ' + result.skippedInvalid,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('419A App Load Ready Import Failed', DTInv_errorMessage_(err), ui.ButtonSet.OK);
  }
}

function DTInv_prepareAppColumns_() {
  var sheet = DTInv_getInventorySheet_({ allowMissingHeaders: true });
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var initialized = false;
  if (sheet.getLastRow() === 0 || !DTInv_rowHasContent_(sheet.getRange(1, 1, 1, lastColumn).getValues()[0])) {
    var initialHeaders = DTINV_CONFIG.REQUIRED_COLUMNS.concat(DTINV_CONFIG.OPTIONAL_COLUMNS).map(function (pair) { return pair[1]; });
    sheet.getRange(1, 1, 1, initialHeaders.length).setValues([initialHeaders]);
    sheet.setFrozenRows(1);
    return { initialized: true, addedColumns: initialHeaders, headers: initialHeaders };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = DTInv_buildColumnMap_(headers);
  var added = [];
  DTINV_CONFIG.REQUIRED_COLUMNS.concat(DTINV_CONFIG.OPTIONAL_COLUMNS).forEach(function (pair) {
    if (map[pair[0]] === -1) added.push(pair[1]);
  });
  if (added.length) {
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, added.length).setValues([added]);
  }
  sheet.setFrozenRows(1);
  return { initialized: initialized, addedColumns: added, headers: sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] };
}

function DTInv_refreshQrLinks_() {
  var sheet = DTInv_getInventorySheet_();
  var values = sheet.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  var baseUrl = DTInv_getWebAppBaseUrl_();
  var output = [];
  var updated = 0;
  var skipped = 0;
  var rolloutLocations = {};
  var rolloutReady = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var room = DTInv_clean_(row[map.room]);
    var loc = DTInv_clean_(row[map.location]);
    var storageId = DTInv_optional_(row, map.storageId);
    if (!room || !loc) {
      output.push(['']);
      skipped += 1;
      continue;
    }
    var storageLabel = DTInv_optional_(row, map.storageLabel);
    var locationCode = DTInv_optional_(row, map.locationCode);
    var routeLoc = storageId || storageLabel || locationCode || loc;
    var url = DTInv_buildLocationUrl_(baseUrl, room, routeLoc);
    output.push([url]);
    updated += 1;
    if (room.toLowerCase() === DTINV_CONFIG.ROLLOUT_ROOM.toLowerCase()) {
      var key = DTInv_locationKey_(row, map);
      rolloutLocations[key] = true;
      rolloutReady[key] = true;
    }
  }
  if (output.length) sheet.getRange(2, map.qrLink + 1, output.length, 1).setValues(output);
  return {
    updatedRows: updated,
    skippedRows: skipped,
    baseUrl: baseUrl,
    rolloutLocationCount: Object.keys(rolloutLocations).length,
    rolloutQrReady: Object.keys(rolloutReady).length
  };
}

function DTInv_refreshQrImages_() {
  var sheet = DTInv_getInventorySheet_();
  var values = sheet.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  if (map.qrImage === -1) throw new Error('QR Code Image column is missing. Run Prepare App Columns first.');
  var linkCol = DTInv_columnLetter_(map.qrLink + 1);
  var formulas = [];
  var updated = 0;
  var skipped = 0;
  for (var i = 1; i < values.length; i++) {
    var rowNumber = i + 1;
    if (DTInv_clean_(values[i][map.qrLink])) {
      formulas.push(['=IMAGE("https://quickchart.io/qr?text="&ENCODEURL($' + linkCol + rowNumber + ')&"&size=180")']);
      updated += 1;
    } else {
      formulas.push(['']);
      skipped += 1;
    }
  }
  if (formulas.length) sheet.getRange(2, map.qrImage + 1, formulas.length, 1).setFormulas(formulas);
  return { updatedRows: updated, skippedRows: skipped };
}

function DTInv_buildStorageMasterSheet_() {
  var ss = DTInv_getSpreadsheet_();
  var inventory = DTInv_getInventorySheet_();
  var values = inventory.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  var baseUrl = DTInv_getWebAppBaseUrl_();
  var locations = DTInv_collectLocations_(values, map, baseUrl);
  var sheet = ss.getSheetByName(DTINV_CONFIG.STORAGE_MASTER_SHEET_NAME) || ss.insertSheet(DTINV_CONFIG.STORAGE_MASTER_SHEET_NAME);
  var headers = ['Storage ID', 'Room', 'Storage Label', 'Specific Location', 'Location Code', 'Storage Type', 'QR Link', 'QR Image', 'Status', 'Notes'];
  var rows = locations.map(function (loc, index) {
    var rowNumber = index + 2;
    var storageType = loc.storageType || DTInv_inferStorageType_(loc);
    var status = loc.attentionCount ? 'Needs Attention' : 'Good';
    return [
      loc.storageId,
      loc.room,
      loc.storageLabel || loc.specificLocation,
      loc.specificLocation,
      loc.locationCode,
      storageType,
      loc.viewUrl,
      '=IMAGE("https://quickchart.io/qr?text="&ENCODEURL(G' + rowNumber + ')&"&size=180")',
      status,
      loc.itemCount + ' item row(s); ' + loc.chemicalCount + ' chemical row(s); ' + loc.attentionCount + ' attention row(s)'
    ];
  });
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 8, rows.length, 1).setFormulas(rows.map(function (row) { return [row[7]]; }));
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return { sheetName: sheet.getName(), storageCount: rows.length };
}

function DTInv_buildQrLabelSheet_() {
  var ss = DTInv_getSpreadsheet_();
  var inventory = DTInv_getInventorySheet_();
  var values = inventory.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  var baseUrl = DTInv_getWebAppBaseUrl_();
  var labels = DTInv_collectLocations_(values, map, baseUrl);
  var sheet = ss.getSheetByName(DTINV_CONFIG.QR_LABEL_SHEET_NAME) || ss.insertSheet(DTINV_CONFIG.QR_LABEL_SHEET_NAME);
  sheet.clear();
  var headers = ['Room', 'Specific Location', 'Storage ID', 'Storage Label', 'Location Code', 'View URL', 'Update URL', 'QR Image Formula', 'Print Label Text'];
  var rows = labels.map(function (loc, index) {
    var rowNumber = index + 2;
    var text = [
      loc.room,
      loc.storageId || loc.specificLocation,
      loc.storageLabel || '',
      'Scan to view inventory'
    ].filter(Boolean).join('\n');
    return [
      loc.room,
      loc.specificLocation,
      loc.storageId,
      loc.storageLabel,
      loc.locationCode,
      loc.viewUrl,
      loc.techUrl,
      '=IMAGE("https://quickchart.io/qr?text="&ENCODEURL(F' + rowNumber + ')&"&size=180")',
      loc.chemicalCount ? text + '\nHAZARD STORAGE - CHECK SAFETY FIRST' : text
    ];
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 8, rows.length, 1).setFormulas(rows.map(function (r) { return [r[7]]; }));
    sheet.getRange(2, 9, rows.length, 1).setWrap(true);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return { sheetName: sheet.getName(), labelCount: rows.length };
}

function DTInv_createReadinessReport_() {
  var ss = DTInv_getSpreadsheet_();
  var inventory = DTInv_getInventorySheet_({ allowMissingHeaders: true });
  var values = inventory.getDataRange().getValues();
  var headers = values.length ? values[0] : [];
  var map = DTInv_buildColumnMap_(headers);
  var missingRequired = DTInv_missingColumns_(map, DTINV_CONFIG.REQUIRED_COLUMNS);
  var missingOptional = DTInv_missingColumns_(map, DTINV_CONFIG.OPTIONAL_COLUMNS);
  var issues = [];
  var roomCounts = {};
  var statusCounts = {};
  var categoryCounts = {};
  var locationKeys = {};
  var duplicateItemKeys = {};
  var qrReadyLocations = {};
  var rolloutLocations = {};
  var rolloutStorageIds = {};
  var placeholderRows = 0;

  missingRequired.forEach(function (label) {
    DTInv_addIssue_(issues, 'ERROR', '', 'MISSING_REQUIRED_COLUMN', label + ' is missing.', '', '', '', '', 'Run Prepare App Columns and map existing data if needed.');
  });
  missingOptional.forEach(function (label) {
    DTInv_addIssue_(issues, 'WARN', '', 'MISSING_OPTIONAL_COLUMN', label + ' is missing.', '', '', '', '', 'Run Prepare App Columns.');
  });

  if (!missingRequired.length) {
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowNumber = i + 1;
      var room = DTInv_clean_(row[map.room]);
      var loc = DTInv_clean_(row[map.location]);
      var itemId = DTInv_clean_(row[map.itemId]);
      var itemName = DTInv_clean_(row[map.itemName]);
      var qtyRaw = row[map.qty];
      var category = DTInv_clean_(row[map.category]);
      var status = DTInv_clean_(row[map.status]);
      var storageId = DTInv_optional_(row, map.storageId);
      var storageLabel = DTInv_optional_(row, map.storageLabel);
      var locationCode = DTInv_optional_(row, map.locationCode);
      var qrLink = DTInv_optional_(row, map.qrLink);
      var qrImage = DTInv_optional_(row, map.qrImage);
      var remarks = DTInv_optional_(row, map.remarks);
      var safetyNote = DTInv_optional_(row, map.safetyNote);
      var reorderLevelText = DTInv_optional_(row, map.reorderLevel);
      var maintenanceDue = DTInv_optional_(row, map.maintenanceDue);
      var isPlaceholder = DTInv_isPlaceholderRow_(row, map);
      var isItem = !isPlaceholder && !!(itemId || itemName);

      if (!DTInv_rowHasContent_(row)) continue;
      if (room) roomCounts[room] = (roomCounts[room] || 0) + 1;
      if (status) statusCounts[status] = (statusCounts[status] || 0) + 1;
      if (category) categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      if (isPlaceholder) placeholderRows += 1;

      if (!room) DTInv_addIssue_(issues, 'ERROR', rowNumber, 'MISSING_ROOM', 'Room is required.', room, loc, itemId, itemName, 'Enter a room.');
      if (!loc) DTInv_addIssue_(issues, 'ERROR', rowNumber, 'MISSING_LOCATION', 'Specific Location is required.', room, loc, itemId, itemName, 'Enter a storage/location.');

      if (isItem || isPlaceholder) {
        var qty = Number(qtyRaw);
        if (qtyRaw === '' || qtyRaw === null || qtyRaw === undefined || !isFinite(qty) || qty < 0) {
          DTInv_addIssue_(issues, 'ERROR', rowNumber, 'INVALID_QTY', 'Quantity must be a non-negative number.', room, loc, itemId, itemName, 'Use 0 or greater.');
        }
        if (DTINV_CONFIG.STATUS_OPTIONS.map(function (s) { return s.toLowerCase(); }).indexOf(status.toLowerCase()) === -1) {
          DTInv_addIssue_(issues, 'ERROR', rowNumber, 'INVALID_STATUS', 'Status is not one of the approved values.', room, loc, itemId, itemName, 'Use Good, Low Stock, Missing, or Needs Maintenance.');
        }
        if (DTINV_CONFIG.HAZARD_CATEGORIES.indexOf(category.toLowerCase()) !== -1 && !remarks && !safetyNote) {
          DTInv_addIssue_(issues, 'WARN', rowNumber, 'CHEMICAL_SAFETY_NOTE_MISSING', 'Chemical rows should include a remark or Safety Note.', room, loc, itemId, itemName, 'Add safety handling notes or SDS reference.');
        }
        if (DTINV_CONFIG.HAZARD_CATEGORIES.indexOf(category.toLowerCase()) !== -1 && map.sdsLink !== -1 && !DTInv_optional_(row, map.sdsLink)) {
          DTInv_addIssue_(issues, 'WARN', rowNumber, 'CHEMICAL_SDS_LINK_MISSING', 'Chemical rows should include an SDS Link where available.', room, loc, itemId, itemName, 'Add an SDS link or confirm it is unavailable.');
        }
        if (status === 'Low Stock' && map.reorderLevel !== -1 && !reorderLevelText) {
          DTInv_addIssue_(issues, 'WARN', rowNumber, 'LOW_STOCK_REORDER_LEVEL_MISSING', 'Low Stock rows should include a Reorder Level for purchasing review.', room, loc, itemId, itemName, 'Enter a reorder threshold.');
        }
        if (reorderLevelText) {
          var reorderLevel = Number(reorderLevelText);
          var qtyForReorder = Number(qtyRaw);
          if (!isFinite(reorderLevel) || reorderLevel < 0) {
            DTInv_addIssue_(issues, 'WARN', rowNumber, 'INVALID_REORDER_LEVEL', 'Reorder Level should be a non-negative number.', room, loc, itemId, itemName, 'Use 0 or greater.');
          } else if (isFinite(qtyForReorder) && qtyForReorder <= reorderLevel && status === 'Good') {
            DTInv_addIssue_(issues, 'WARN', rowNumber, 'REORDER_THRESHOLD_REACHED', 'Quantity is at or below reorder level but status is still Good.', room, loc, itemId, itemName, 'Consider Low Stock status.');
          }
        }
        if (status === 'Needs Maintenance' && !maintenanceDue && !remarks) {
          DTInv_addIssue_(issues, 'WARN', rowNumber, 'MAINTENANCE_DETAIL_MISSING', 'Maintenance rows should include a remark or Maintenance Due date.', room, loc, itemId, itemName, 'Add maintenance details.');
        }
      }

      if (room && loc) {
        var locKey = DTInv_locationKeyFromValues_(room, loc, storageId, storageLabel, locationCode);
        locationKeys[locKey] = true;
        if (qrLink && DTInv_isValidQrLink_(qrLink, room)) qrReadyLocations[locKey] = true;
        if (!qrLink) DTInv_addIssue_(issues, 'WARN', rowNumber, 'MISSING_QR_LINK', 'QR link is missing.', room, loc, itemId, itemName, 'Run Refresh QR Links.');
        else if (!DTInv_isValidQrLink_(qrLink, room)) DTInv_addIssue_(issues, 'WARN', rowNumber, 'QR_LINK_INVALID', 'QR link is not a valid deployed /exec storage URL.', room, loc, itemId, itemName, 'Run Refresh QR Links.');
        if (room.toLowerCase() === DTINV_CONFIG.ROLLOUT_ROOM.toLowerCase()) {
          rolloutLocations[locKey] = true;
          if (storageId) rolloutStorageIds[storageId] = true;
          if (!storageId) DTInv_addIssue_(issues, 'WARN', rowNumber, '419A_MISSING_STORAGE_ID', '419A rows should use Storage ID where possible.', room, loc, itemId, itemName, 'Import/enrich from the 419A storage master.');
        }
      }

      if (map.qrImage !== -1 && qrLink && !qrImage) {
        DTInv_addIssue_(issues, 'WARN', rowNumber, 'MISSING_QR_IMAGE', 'QR image formula is missing.', room, loc, itemId, itemName, 'Run Refresh QR Images.');
      }
      if (room === 'V++' && qrLink && qrLink.indexOf('room=V%2B%2B') === -1) {
        DTInv_addIssue_(issues, 'WARN', rowNumber, 'VPP_URL_ENCODING_RISK', 'V++ link may not be URL-encoded correctly.', room, loc, itemId, itemName, 'Run Refresh QR Links.');
      }

      if (isItem) {
        var itemKey = [room, loc, storageId, storageLabel, locationCode, itemId, itemName].map(DTInv_keyPart_).join('||');
        if (duplicateItemKeys[itemKey]) {
          DTInv_addIssue_(issues, 'WARN', rowNumber, 'POSSIBLE_DUPLICATE_ITEM', 'Possible duplicate item identity.', room, loc, itemId, itemName, 'Check whether this is the same item row as row ' + duplicateItemKeys[itemKey] + '.');
        } else {
          duplicateItemKeys[itemKey] = rowNumber;
        }
      }
    }
  }

  var report = ss.getSheetByName(DTINV_CONFIG.READINESS_SHEET_NAME) || ss.insertSheet(DTINV_CONFIG.READINESS_SHEET_NAME);
  report.clear();
  var errorCount = issues.filter(function (issue) { return issue.severity === 'ERROR'; }).length;
  var warningCount = issues.filter(function (issue) { return issue.severity !== 'ERROR'; }).length;
  var locationCount = Object.keys(locationKeys).length;
  var qrReadyCount = Object.keys(qrReadyLocations).length;
  var rolloutStorageCount = Object.keys(rolloutLocations).length;

  var summaryRows = [
    ['Generated At', new Date()],
    ['Spreadsheet', ss.getName()],
    ['Inventory Sheet', inventory.getName()],
    ['Critical Errors', errorCount],
    ['Warnings', warningCount],
    ['Unique Locations', locationCount],
    ['QR-ready Locations', qrReadyCount],
    ['419A Storage Count', rolloutStorageCount],
    ['419A Storage IDs', Object.keys(rolloutStorageIds).length],
    ['Placeholder Rows', placeholderRows],
    [''],
    ['Summary By Room'],
  ];
  Object.keys(roomCounts).sort().forEach(function (key) { summaryRows.push([key, roomCounts[key]]); });
  summaryRows.push([''], ['Summary By Status']);
  Object.keys(statusCounts).sort().forEach(function (key) { summaryRows.push([key, statusCounts[key]]); });
  summaryRows.push([''], ['Summary By Category']);
  Object.keys(categoryCounts).sort().forEach(function (key) { summaryRows.push([key, categoryCounts[key]]); });

  report.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows.map(function (row) {
    return [row[0] || '', row.length > 1 ? row[1] : ''];
  }));
  var issueHeaderRow = summaryRows.length + 2;
  var issueHeaders = ['Severity', 'Row', 'Issue', 'Detail', 'Room', 'Specific Location', 'Item ID', 'Item Name', 'Suggested Fix'];
  report.getRange(issueHeaderRow, 1, 1, issueHeaders.length).setValues([issueHeaders]);
  if (issues.length) {
    report.getRange(issueHeaderRow + 1, 1, issues.length, issueHeaders.length).setValues(issues.map(function (issue) {
      return [issue.severity, issue.row, issue.issue, issue.detail, issue.room, issue.location, issue.itemId, issue.itemName, issue.suggestedFix];
    }));
  }
  report.setFrozenRows(issueHeaderRow);
  report.autoResizeColumns(1, issueHeaders.length);
  return {
    sheetName: report.getName(),
    errorCount: errorCount,
    warningCount: warningCount,
    issueCount: issues.length,
    locationCount: locationCount,
    qrReadyCount: qrReadyCount,
    rolloutStorageCount: rolloutStorageCount,
    roomCounts: roomCounts,
    statusCounts: statusCounts,
    categoryCounts: categoryCounts
  };
}

function DTInv_import419AStorageMasterFromSource_(sourceSpreadsheetId) {
  DTInv_prepareAppColumns_();
  var sourceId = DTInv_extractSpreadsheetId_(sourceSpreadsheetId);
  var sourceSs = SpreadsheetApp.openById(sourceId);
  var source = DTInv_getStorageMasterRowsFromSource_(sourceSs);
  if (!source.sheetName) throw new Error('Could not find a supported 419A storage master tab.');

  var sheet = DTInv_getInventorySheet_();
  var values = sheet.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  var existing = {};
  var existingRowsByKey = {};
  for (var i = 1; i < values.length; i++) {
    DTInv_locationMatchKeys_(values[i], map).forEach(function (key) {
      existing[key] = true;
      if (!existingRowsByKey[key]) existingRowsByKey[key] = [];
      existingRowsByKey[key].push(i);
    });
  }

  var rows = [];
  var skippedExisting = 0;
  var skippedInvalid = 0;
  var enrichedRows = 0;
  var baseUrl = DTInv_getWebAppBaseUrl_();
  source.rows.forEach(function (meta) {
    if (meta.room.toLowerCase() !== DTINV_CONFIG.ROLLOUT_ROOM.toLowerCase()) return;
    if (!meta.storageId || !meta.displayLocation) {
      skippedInvalid += 1;
      return;
    }
    var keys = DTInv_locationMatchKeysFromValues_(meta.room, meta.displayLocation, meta.storageId, meta.storageLabel, meta.locationCode);
    if (keys.some(function (key) { return existing[key]; })) {
      skippedExisting += 1;
      var matchedRows = {};
      keys.forEach(function (key) {
        (existingRowsByKey[key] || []).forEach(function (rowIndex) { matchedRows[rowIndex] = true; });
      });
      enrichedRows += DTInv_enrichExistingStorageRows_(sheet, values, map, Object.keys(matchedRows), meta, baseUrl);
      return;
    }
    var row = new Array(values[0].length).fill('');
    row[map.room] = meta.room;
    row[map.location] = meta.displayLocation;
    row[map.qty] = 0;
    row[map.category] = 'Storage';
    row[map.status] = 'Good';
    if (map.remarks !== -1) row[map.remarks] = 'Placeholder row for QR/location page';
    if (map.isPlaceholder !== -1) row[map.isPlaceholder] = true;
    if (map.storageType !== -1) row[map.storageType] = DTInv_inferStorageType_({
      storageId: meta.storageId,
      storageLabel: meta.storageLabel,
      specificLocation: meta.displayLocation,
      locationCode: meta.locationCode,
      categories: {}
    });
    if (map.locationCode !== -1) row[map.locationCode] = meta.locationCode;
    if (map.storageId !== -1) row[map.storageId] = meta.storageId;
    if (map.storageLabel !== -1) row[map.storageLabel] = meta.storageLabel;
    if (map.qrLink !== -1) row[map.qrLink] = DTInv_buildLocationUrl_(baseUrl, meta.room, meta.storageId);
    rows.push(row);
    keys.forEach(function (key) { existing[key] = true; });
  });

  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, values[0].length).setValues(rows);
  return { importedRows: rows.length, enrichedRows: enrichedRows, skippedExisting: skippedExisting, skippedInvalid: skippedInvalid, sourceSheet: source.sheetName };
}

function DTInv_enrichExistingStorageRows_(sheet, values, map, rowIndexes, meta, baseUrl) {
  var enriched = 0;
  rowIndexes.forEach(function (rowIndexText) {
    var rowIndex = Number(rowIndexText);
    var row = values[rowIndex];
    var changed = false;
    changed = DTInv_setBlankCell_(sheet, row, rowIndex, map.locationCode, meta.locationCode) || changed;
    changed = DTInv_setBlankCell_(sheet, row, rowIndex, map.storageId, meta.storageId) || changed;
    changed = DTInv_setBlankCell_(sheet, row, rowIndex, map.storageLabel, meta.storageLabel) || changed;
    if (map.qrLink !== -1 && !DTInv_clean_(row[map.qrLink])) {
      var routeLoc = meta.storageId || DTInv_clean_(row[map.location]) || meta.displayLocation;
      changed = DTInv_setBlankCell_(sheet, row, rowIndex, map.qrLink, DTInv_buildLocationUrl_(baseUrl, meta.room, routeLoc)) || changed;
    }
    if (changed) enriched += 1;
  });
  return enriched;
}

function DTInv_setBlankCell_(sheet, row, rowIndex, columnIndex, value) {
  var nextValue = DTInv_clean_(value);
  if (columnIndex === -1 || !nextValue || DTInv_clean_(row[columnIndex])) return false;
  sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(nextValue);
  row[columnIndex] = nextValue;
  return true;
}

function DTInv_import419AReadyFromSource_(sourceSpreadsheetId) {
  DTInv_prepareAppColumns_();
  var sourceId = DTInv_extractSpreadsheetId_(sourceSpreadsheetId);
  var sourceSs = SpreadsheetApp.openById(sourceId);
  var sourceSheet = DTInv_findSheetByNames_(sourceSs, DTINV_CONFIG.IMPORT_READY_SHEET_NAMES);
  if (!sourceSheet) throw new Error('Could not find 419A_App_Load_Ready / 419A App Load Ready.');

  var sourceValues = sourceSheet.getDataRange().getValues();
  var sourceHeaderRow = DTInv_findInventoryHeaderRow_(sourceValues);
  if (sourceHeaderRow === -1) throw new Error('Source App Load Ready tab does not contain inventory headers.');
  var sourceMap = DTInv_getRequiredMap_(sourceValues[sourceHeaderRow], { allowMissingQrLink: true });
  var storageLookup = DTInv_buildStorageLookupFromSource_(sourceSs);

  var target = DTInv_getInventorySheet_();
  var targetValues = target.getDataRange().getValues();
  var targetMap = DTInv_getRequiredMap_(targetValues[0]);
  var existing = DTInv_existingImportKeys_(targetValues, targetMap);
  var baseUrl = DTInv_getWebAppBaseUrl_();
  var rows = [];
  var skippedDuplicates = 0;
  var skippedInvalid = 0;

  for (var i = sourceHeaderRow + 1; i < sourceValues.length; i++) {
    var src = sourceValues[i];
    var itemId = DTInv_clean_(src[sourceMap.itemId]);
    var itemName = DTInv_clean_(src[sourceMap.itemName]);
    var room = DTInv_clean_(src[sourceMap.room]) || DTINV_CONFIG.ROLLOUT_ROOM;
    var loc = DTInv_clean_(src[sourceMap.location]);
    if (!room || !loc || (!itemId && !itemName)) {
      skippedInvalid += 1;
      continue;
    }
    var storageId = DTInv_optional_(src, sourceMap.storageId);
    var meta = storageLookup[DTInv_keyPart_(room) + '||' + DTInv_keyPart_(storageId)] ||
      storageLookup[DTInv_keyPart_(room) + '||' + DTInv_keyPart_(loc)] || {};
    storageId = storageId || meta.storageId || '';
    var storageLabel = DTInv_optional_(src, sourceMap.storageLabel) || meta.storageLabel || '';
    var locationCode = DTInv_optional_(src, sourceMap.locationCode) || meta.locationCode || '';
    var keys = DTInv_importKeysFromValues_(room, loc, storageId, storageLabel, locationCode, itemId, itemName);
    if (keys.some(function (key) { return existing[key]; })) {
      skippedDuplicates += 1;
      continue;
    }
    var row = new Array(targetValues[0].length).fill('');
    row[targetMap.itemId] = itemId;
    row[targetMap.itemName] = itemName;
    row[targetMap.room] = room;
    row[targetMap.location] = loc;
    row[targetMap.qty] = DTInv_toNonNegativeNumber_(src[sourceMap.qty]);
    row[targetMap.category] = DTInv_clean_(src[sourceMap.category]);
    row[targetMap.status] = DTInv_normalizeStatus_(src[sourceMap.status]);
    if (targetMap.unit !== -1) row[targetMap.unit] = DTInv_optional_(src, sourceMap.unit);
    if (targetMap.remarks !== -1) row[targetMap.remarks] = DTInv_optional_(src, sourceMap.remarks);
    if (targetMap.locationCode !== -1) row[targetMap.locationCode] = locationCode;
    if (targetMap.storageId !== -1) row[targetMap.storageId] = storageId;
    if (targetMap.storageLabel !== -1) row[targetMap.storageLabel] = storageLabel;
    if (targetMap.qrLink !== -1) row[targetMap.qrLink] = DTInv_buildLocationUrl_(baseUrl, room, storageId || storageLabel || locationCode || loc);
    rows.push(row);
    keys.forEach(function (key) { existing[key] = true; });
  }

  if (rows.length) target.getRange(target.getLastRow() + 1, 1, rows.length, targetValues[0].length).setValues(rows);
  return { importedRows: rows.length, skippedDuplicates: skippedDuplicates, skippedInvalid: skippedInvalid, sourceSheet: sourceSheet.getName() };
}

function DTInv_getDiagnostics_() {
  var ss = DTInv_getSpreadsheet_();
  var sheet = DTInv_getInventorySheet_({ allowMissingHeaders: true });
  var headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  var map = DTInv_buildColumnMap_(headers);
  var webUrl = DTInv_getWebAppBaseUrl_();
  return {
    id: ss.getId(),
    title: ss.getName(),
    sheetNames: ss.getSheets().map(function (s) { return s.getName(); }),
    inventorySheetName: sheet.getName(),
    dataRows: Math.max(sheet.getLastRow() - 1, 0),
    webAppBaseUrl: webUrl,
    webAppBaseUrlSource: PropertiesService.getScriptProperties().getProperty(DTINV_CONFIG.WEB_APP_BASE_URL_PROPERTY) ? 'Script Property' : 'default @8',
    missingRequired: DTInv_missingColumns_(map, DTINV_CONFIG.REQUIRED_COLUMNS),
    missingOptional: DTInv_missingColumns_(map, DTINV_CONFIG.OPTIONAL_COLUMNS)
  };
}

function DTInv_get419AReadinessSummary_() {
  var sheet = DTInv_getInventorySheet_();
  var values = sheet.getDataRange().getValues();
  var map = DTInv_getRequiredMap_(values[0]);
  var locations = {};
  var withStorage = {};
  var qrReady = {};
  var itemRows = 0;
  var placeholders = 0;
  var chemicals = 0;
  var attention = 0;
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var room = DTInv_clean_(row[map.room]);
    if (room.toLowerCase() !== DTINV_CONFIG.ROLLOUT_ROOM.toLowerCase()) continue;
    var locKey = DTInv_locationKey_(row, map);
    locations[locKey] = true;
    if (DTInv_optional_(row, map.storageId)) withStorage[locKey] = true;
    if (DTInv_isValidQrLink_(DTInv_optional_(row, map.qrLink), room)) qrReady[locKey] = true;
    var isItem = !DTInv_isPlaceholderRow_(row, map) && !!(DTInv_clean_(row[map.itemId]) || DTInv_clean_(row[map.itemName]));
    if (isItem) itemRows += 1;
    if (!isItem && DTInv_clean_(row[map.category]).toLowerCase() === 'storage') placeholders += 1;
    if (DTINV_CONFIG.HAZARD_CATEGORIES.indexOf(DTInv_clean_(row[map.category]).toLowerCase()) !== -1) chemicals += 1;
    if (DTInv_clean_(row[map.status]) && DTInv_clean_(row[map.status]) !== 'Good') attention += 1;
  }
  return {
    locationCount: Object.keys(locations).length,
    withStorageId: Object.keys(withStorage).length,
    missingStorageId: Math.max(Object.keys(locations).length - Object.keys(withStorage).length, 0),
    qrReady: Object.keys(qrReady).length,
    itemRows: itemRows,
    placeholderRows: placeholders,
    chemicalRows: chemicals,
    attentionRows: attention
  };
}

function DTInv_getStorageMasterRowsFromSource_(sourceSs) {
  var sheet = DTInv_findSheetByNames_(sourceSs, DTINV_CONFIG.STORAGE_MASTER_SHEET_NAMES);
  var result = { sheetName: '', rows: [] };
  if (!sheet) return result;
  var values = sheet.getDataRange().getValues();
  var headerRow = DTInv_findStorageHeaderRow_(values);
  if (headerRow === -1) return result;
  var headers = values[headerRow].map(DTInv_normalizeHeader_);
  var map = {
    room: DTInv_findHeaderIndex_(headers, ['room']),
    storageId: DTInv_findHeaderIndex_(headers, ['storage id', 'storageid']),
    displayLocation: DTInv_findHeaderIndex_(headers, ['display location', 'specific location', 'location']),
    locationCode: DTInv_findHeaderIndex_(headers, ['location code', 'locationcode']),
    storageLabel: DTInv_findHeaderIndex_(headers, ['storage label', 'storagelabel'])
  };
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    var room = DTInv_optional_(row, map.room) || DTINV_CONFIG.ROLLOUT_ROOM;
    var storageId = DTInv_optional_(row, map.storageId);
    var display = DTInv_optional_(row, map.displayLocation);
    if (!storageId && !display) continue;
    result.rows.push({
      room: room,
      storageId: storageId,
      displayLocation: display || storageId,
      locationCode: DTInv_optional_(row, map.locationCode),
      storageLabel: DTInv_optional_(row, map.storageLabel)
    });
  }
  result.sheetName = sheet.getName();
  return result;
}

function DTInv_buildStorageLookupFromSource_(sourceSs) {
  var source = DTInv_getStorageMasterRowsFromSource_(sourceSs);
  var lookup = {};
  source.rows.forEach(function (meta) {
    var roomKey = DTInv_keyPart_(meta.room);
    if (meta.storageId) lookup[roomKey + '||' + DTInv_keyPart_(meta.storageId)] = meta;
    if (meta.displayLocation) lookup[roomKey + '||' + DTInv_keyPart_(meta.displayLocation)] = meta;
    if (meta.storageLabel) lookup[roomKey + '||' + DTInv_keyPart_(meta.storageLabel)] = meta;
    if (meta.locationCode) lookup[roomKey + '||' + DTInv_keyPart_(meta.locationCode)] = meta;
  });
  return lookup;
}

function DTInv_collectLocations_(values, map, baseUrl) {
  var byKey = {};
  var labels = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var room = DTInv_clean_(row[map.room]);
    var loc = DTInv_clean_(row[map.location]);
    if (!room || !loc) continue;
    var storageId = DTInv_optional_(row, map.storageId);
    var storageLabel = DTInv_optional_(row, map.storageLabel);
    var locationCode = DTInv_optional_(row, map.locationCode);
    var storageType = DTInv_optional_(row, map.storageType);
    var key = DTInv_locationKeyFromValues_(room, loc, storageId, storageLabel, locationCode);
    var existing = byKey[key];
    if (!existing) {
      var routeLoc = storageId || storageLabel || locationCode || loc;
      var viewUrl = DTInv_buildLocationUrl_(baseUrl, room, routeLoc);
      existing = {
        room: room,
        specificLocation: loc,
        storageId: storageId,
        storageLabel: storageLabel,
        locationCode: locationCode,
        storageType: storageType,
        viewUrl: viewUrl,
        techUrl: viewUrl + '&mode=tech',
        itemCount: 0,
        chemicalCount: 0,
        attentionCount: 0,
        categories: {}
      };
      byKey[key] = existing;
      labels.push(existing);
    }
    if (!DTInv_isPlaceholderRow_(row, map) && (DTInv_clean_(row[map.itemId]) || DTInv_clean_(row[map.itemName]))) {
      var category = DTInv_clean_(row[map.category]);
      var status = DTInv_clean_(row[map.status]);
      existing.itemCount += 1;
      if (category) existing.categories[category.toLowerCase()] = true;
      if (DTINV_CONFIG.HAZARD_CATEGORIES.indexOf(category.toLowerCase()) !== -1) existing.chemicalCount += 1;
      if (status && status !== 'Good') existing.attentionCount += 1;
    }
  }
  labels.sort(function (a, b) {
    return [a.room, a.storageId || a.specificLocation].join('|').localeCompare([b.room, b.storageId || b.specificLocation].join('|'));
  });
  return labels;
}

function DTInv_inferStorageType_(loc) {
  var text = [loc.storageLabel, loc.specificLocation, loc.locationCode, loc.storageId].filter(Boolean).join(' ').toLowerCase();
  var categories = Object.keys(loc.categories || {}).join(' ');
  if (loc.chemicalCount || text.indexOf('chem') !== -1 || categories.indexOf('chemical') !== -1) return 'Chemical Storage';
  if (text.indexOf('machine') !== -1 || categories.indexOf('machine') !== -1) return 'Machine Zone';
  if (text.indexOf('elect') !== -1 || text.indexOf('arduino') !== -1 || categories.indexOf('electronics') !== -1) return 'Electronics Storage';
  if (text.indexOf('tool') !== -1 || categories.indexOf('tool') !== -1) return 'Tool Storage';
  if (text.indexOf('rack') !== -1 || text.indexOf('material') !== -1 || categories.indexOf('material') !== -1) return 'Material Storage';
  if (text.indexOf('tray') !== -1) return 'Tray Storage';
  if (text.indexOf('trolley') !== -1) return 'Trolley';
  if (text.indexOf('cupboard') !== -1 || text.indexOf('cabinet') !== -1) return 'Cupboard / Cabinet';
  return 'Storage';
}

function DTInv_getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet. Open the live dashboard Google Sheet first.');
  return ss;
}

function DTInv_getInventorySheet_(options) {
  var opts = options || {};
  var ss = DTInv_getSpreadsheet_();
  var sheet = ss.getSheetByName(DTINV_CONFIG.INVENTORY_SHEET_NAME);
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var values = sheets[i].getDataRange().getValues();
      if (values.length && DTInv_missingColumns_(DTInv_buildColumnMap_(values[0]), DTINV_CONFIG.REQUIRED_COLUMNS).length === 0) {
        sheet = sheets[i];
        break;
      }
    }
  }
  if (!sheet) throw new Error('Could not find an Inventory tab with required headers.');
  if (!opts.allowMissingHeaders) DTInv_getRequiredMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  return sheet;
}

function DTInv_getRequiredMap_(headers, options) {
  var map = DTInv_buildColumnMap_(headers || []);
  var required = DTINV_CONFIG.REQUIRED_COLUMNS.filter(function (pair) {
    return !(options && options.allowMissingQrLink && pair[0] === 'qrLink');
  });
  var missing = DTInv_missingColumns_(map, required);
  if (missing.length) throw new Error('Missing required Inventory columns: ' + missing.join(', '));
  return map;
}

function DTInv_buildColumnMap_(headers) {
  var normalized = (headers || []).map(DTInv_normalizeHeader_);
  var map = {};
  DTINV_CONFIG.REQUIRED_COLUMNS.concat(DTINV_CONFIG.OPTIONAL_COLUMNS).forEach(function (pair) {
    map[pair[0]] = DTInv_findHeaderIndex_(normalized, DTINV_CONFIG.HEADER_ALIASES[pair[0]] || [pair[1]]);
  });
  return map;
}

function DTInv_missingColumns_(map, pairs) {
  return pairs.filter(function (pair) { return map[pair[0]] === -1; }).map(function (pair) { return pair[1]; });
}

function DTInv_findInventoryHeaderRow_(values) {
  for (var i = 0; i < Math.min(values.length, 10); i++) {
    var map = DTInv_buildColumnMap_(values[i]);
    var missing = DTInv_missingColumns_(map, DTINV_CONFIG.REQUIRED_COLUMNS.filter(function (pair) { return pair[0] !== 'qrLink'; }));
    if (!missing.length) return i;
  }
  return -1;
}

function DTInv_findStorageHeaderRow_(values) {
  for (var i = 0; i < Math.min(values.length, 10); i++) {
    var headers = values[i].map(DTInv_normalizeHeader_);
    if (DTInv_findHeaderIndex_(headers, ['storage id', 'storageid']) !== -1 &&
        DTInv_findHeaderIndex_(headers, ['display location', 'specific location', 'location']) !== -1) {
      return i;
    }
  }
  return -1;
}

function DTInv_findHeaderIndex_(normalizedHeaders, aliases) {
  for (var i = 0; i < normalizedHeaders.length; i++) {
    if (aliases.indexOf(normalizedHeaders[i]) !== -1) return i;
  }
  return -1;
}

function DTInv_findSheetByNames_(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  return null;
}

function DTInv_getWebAppBaseUrl_() {
  return DTInv_clean_(PropertiesService.getScriptProperties().getProperty(DTINV_CONFIG.WEB_APP_BASE_URL_PROPERTY)) || DTINV_CONFIG.DEFAULT_WEB_APP_BASE_URL;
}

function DTInv_buildLocationUrl_(baseUrl, room, loc) {
  return String(baseUrl).replace(/[?#].*$/, '') + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function DTInv_isValidQrLink_(url, room) {
  var value = DTInv_clean_(url);
  if (!value) return false;
  if (!/^https:\/\/script\.google\.com\//.test(value)) return false;
  if (DTInv_clean_(room) === 'V++' && value.indexOf('room=V%2B%2B') === -1) return false;
  return value.indexOf('/exec?') !== -1 && value.indexOf('room=') !== -1 && value.indexOf('loc=') !== -1;
}

function DTInv_extractSpreadsheetId_(value) {
  var raw = DTInv_clean_(value);
  var match = raw.match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function DTInv_existingImportKeys_(values, map) {
  var keys = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var room = DTInv_clean_(row[map.room]);
    var loc = DTInv_clean_(row[map.location]);
    var itemId = DTInv_clean_(row[map.itemId]);
    var itemName = DTInv_clean_(row[map.itemName]);
    if (DTInv_isPlaceholderRow_(row, map) || (!itemId && !itemName)) continue;
    var storageId = DTInv_optional_(row, map.storageId);
    var storageLabel = DTInv_optional_(row, map.storageLabel);
    var locationCode = DTInv_optional_(row, map.locationCode);
    DTInv_importKeysFromValues_(room, loc, storageId, storageLabel, locationCode, itemId, itemName).forEach(function (key) { keys[key] = true; });
  }
  return keys;
}

function DTInv_importKeysFromValues_(room, loc, storageId, storageLabel, locationCode, itemId, itemName) {
  return [
    [room, loc, storageId, storageLabel, locationCode, itemId, itemName],
    [room, loc, itemId, itemName],
    [room, storageId, itemId, itemName],
    [room, locationCode, itemId, itemName],
    [room, storageLabel, itemId, itemName]
  ].map(function (parts) { return parts.map(DTInv_keyPart_).join('||'); }).filter(function (key) {
    return key.replace(/\|/g, '') !== '';
  });
}

function DTInv_locationKey_(row, map) {
  return DTInv_locationKeyFromValues_(
    DTInv_clean_(row[map.room]),
    DTInv_clean_(row[map.location]),
    DTInv_optional_(row, map.storageId),
    DTInv_optional_(row, map.storageLabel),
    DTInv_optional_(row, map.locationCode)
  );
}

function DTInv_locationKeyFromValues_(room, loc, storageId, storageLabel, locationCode) {
  return [room, storageId || loc, storageLabel, locationCode].map(DTInv_keyPart_).join('||');
}

function DTInv_locationMatchKeys_(row, map) {
  return DTInv_locationMatchKeysFromValues_(
    DTInv_clean_(row[map.room]),
    DTInv_clean_(row[map.location]),
    DTInv_optional_(row, map.storageId),
    DTInv_optional_(row, map.storageLabel),
    DTInv_optional_(row, map.locationCode)
  );
}

function DTInv_locationMatchKeysFromValues_(room, loc, storageId, storageLabel, locationCode) {
  return [loc, storageId, storageLabel, locationCode].filter(Boolean).map(function (value) {
    return DTInv_keyPart_(room) + '||' + DTInv_keyPart_(value);
  });
}

function DTInv_isPlaceholderRow_(row, map) {
  var explicitPlaceholder = DTInv_optional_(row, map.isPlaceholder).toLowerCase();
  if (['true', 'yes', 'y', '1', 'placeholder'].indexOf(explicitPlaceholder) !== -1) return true;
  var category = DTInv_optional_(row, map.category).toLowerCase();
  var remarks = DTInv_optional_(row, map.remarks).toLowerCase();
  var itemId = DTInv_optional_(row, map.itemId);
  var itemName = DTInv_optional_(row, map.itemName);
  return category === 'storage' &&
    remarks.indexOf('placeholder row for qr/location page') !== -1 &&
    !itemId &&
    !itemName;
}

function DTInv_addIssue_(issues, severity, row, issue, detail, room, location, itemId, itemName, suggestedFix) {
  issues.push({
    severity: severity,
    row: row,
    issue: issue,
    detail: detail,
    room: room,
    location: location,
    itemId: itemId,
    itemName: itemName,
    suggestedFix: suggestedFix
  });
}

function DTInv_columnLetter_(indexOneBased) {
  var n = indexOneBased;
  var result = '';
  while (n > 0) {
    var mod = (n - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    n = Math.floor((n - mod) / 26);
  }
  return result;
}

function DTInv_normalizeStatus_(value) {
  var raw = DTInv_clean_(value);
  for (var i = 0; i < DTINV_CONFIG.STATUS_OPTIONS.length; i++) {
    if (DTINV_CONFIG.STATUS_OPTIONS[i].toLowerCase() === raw.toLowerCase()) return DTINV_CONFIG.STATUS_OPTIONS[i];
  }
  return 'Good';
}

function DTInv_toNonNegativeNumber_(value) {
  var num = Number(value);
  return isFinite(num) && num >= 0 ? num : 0;
}

function DTInv_optional_(row, index) {
  return typeof index === 'number' && index >= 0 ? DTInv_clean_(row[index]) : '';
}

function DTInv_clean_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function DTInv_normalizeHeader_(value) {
  return DTInv_clean_(value).replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').toLowerCase();
}

function DTInv_rowHasContent_(row) {
  return (row || []).some(function (cell) { return DTInv_clean_(cell) !== ''; });
}

function DTInv_keyPart_(value) {
  return DTInv_clean_(value).toLowerCase();
}

function DTInv_maskId_(id) {
  var value = DTInv_clean_(id);
  return value.length <= 12 ? value : value.slice(0, 6) + '...' + value.slice(-6);
}

function DTInv_errorMessage_(err) {
  return err && err.message ? err.message : String(err);
}
