/**
 * D&T QR Inventory System
 * Standalone single-file Google Apps Script web app.
 *
 * Required Script Properties:
 * - SPREADSHEET_ID
 * - WEB_APP_BASE_URL (recommended; app can still read inventory without it)
 * - INVENTORY_SHEET_NAME (optional)
 */

const CONFIG = Object.freeze({
  APP_TITLE: 'D&T QR Inventory System',
  HEADER_ROW: 1,
  DEFAULT_SHEET_NAME: 'Inventory',
  SPREADSHEET_ID_PROPERTY: 'SPREADSHEET_ID',
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  INVENTORY_SHEET_NAME_PROPERTY: 'INVENTORY_SHEET_NAME',
  ROLLOUT_ROOM: '419A',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemicals', 'chemical'],
  QUICKCHART_QR_BASE: 'https://quickchart.io/qr?size=220&text=',
  IMPORT_READY_SHEET_NAMES: ['419A_App_Load_Ready', '419A App Load Ready', '419AAppLoadReady'],
  STORAGE_MASTER_SHEET_NAMES: ['419A_Storage_Master', '419A Storage Master', '419AStorageMaster', 'Room_QR_Label_Plan', 'RM 419A 2026'],
  QR_LABEL_SHEET_NAME: 'QR_Labels',
  READINESS_REPORT_SHEET_NAME: 'Inventory_Readiness_Report',
  DEBUG_PANEL: false,
  ALIASES: {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'itemname', 'name'],
    room: ['room'],
    location: ['specific location', 'location', 'specificlocation'],
    qty: ['qty', 'quantity'],
    category: ['category'],
    status: ['status'],
    qrLink: [
      'qr code link (auto-generated)',
      'qr code link',
      'auto-generated qr link',
      'qr link',
      'qr url'
    ],
    qrImage: ['qr code image', 'qr image'],
    unit: ['unit', 'uom'],
    remarks: ['remarks', 'remark', 'notes', 'note'],
    locationCode: ['location code', 'locationcode', 'location id'],
    storageId: ['storage id', 'storageid', 'storage_id'],
    storageLabel: ['storage label', 'storagelabel', 'storage_label']
  },
  REQUIRED_COLUMNS: ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'],
  OPTIONAL_COLUMNS: ['qrLink', 'qrImage', 'unit', 'remarks', 'locationCode', 'storageId', 'storageLabel'],
  COLUMN_LABELS: {
    itemId: 'Item ID',
    itemName: 'Item Name',
    room: 'Room',
    location: 'Specific Location',
    qty: 'Qty',
    category: 'Category',
    status: 'Status',
    qrLink: 'QR Code Link (Auto-Generated)',
    unit: 'Unit',
    remarks: 'Remarks',
    locationCode: 'Location Code',
    storageId: 'Storage ID',
    storageLabel: 'Storage Label',
    qrImage: 'QR Code Image'
  }
});

function doGet(e) {
  const params = getRequestParams_(e);
  let bootstrap;

  try {
    bootstrap = buildBootstrapData_(params);
  } catch (err) {
    bootstrap = buildErrorBootstrap_(params, err);
  }

  return HtmlService
    .createHtmlOutput(buildPageHtml_(params, bootstrap))
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('D&T Inventory')
    .addItem('Refresh QR Links', 'menuRefreshQrLinks')
    .addItem('Refresh QR Images (Optional)', 'menuRefreshQrImages')
    .addSeparator()
    .addItem('Prepare App Columns', 'menuPrepareAppColumns')
    .addItem('Build QR Label Sheet', 'menuBuildQrLabelSheet')
    .addItem('Create Readiness Report', 'menuCreateReadinessReport')
    .addSeparator()
    .addItem('419A Readiness Summary', 'show419AReadiness')
    .addItem('Import 419A Storage Master', 'promptImport419AStorageMaster')
    .addItem('Import 419A App Load Ready', 'promptImport419AReady')
    .addSeparator()
    .addItem('Open Web App', 'openWebApp')
    .addItem('Set App Config', 'promptSetAppConfig')
    .addItem('Set WEB_APP_BASE_URL', 'promptSetWebAppBaseUrl')
    .addSeparator()
    .addItem('Config Status / Diagnostics', 'showConfigStatus')
    .addToUi();
}

function openWebApp() {
  const ui = SpreadsheetApp.getUi();
  const url = getWebAppBaseUrl_({ silent: true });
  if (!url) {
    ui.alert('WEB_APP_BASE_URL is not configured yet.');
    return;
  }
  ui.alert('Open this URL in your browser:\n\n' + url);
}

function menuRefreshQrLinks() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = refreshQrLinks();
    ui.alert(
      'QR Links Refreshed',
      'Updated ' + result.updatedRows + ' row(s).\nSkipped ' + result.skippedRows + ' row(s) without room/location.\nBase URL: ' + result.baseUrl,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('QR Link Refresh Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuRefreshQrImages() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = refreshQrImages();
    ui.alert('QR Images Refreshed', 'Updated ' + result.updatedRows + ' QR image formula(s).', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('QR Image Refresh Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuPrepareAppColumns() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = ensureAppColumns_();
    ui.alert(
      'App Columns Prepared',
      result.initialized
        ? 'Initialized header row with all required and rollout columns.'
        : 'Added columns: ' + (result.addedColumns.length ? result.addedColumns.join(', ') : 'none needed'),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Prepare Columns Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuBuildQrLabelSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = buildQrLabelSheet();
    ui.alert(
      'QR Label Sheet Ready',
      'Created/updated "' + result.sheetName + '" with ' + result.labelCount + ' storage label row(s).',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('QR Label Sheet Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuCreateReadinessReport() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = createReadinessReport();
    ui.alert(
      'Readiness Report Ready',
      'Created/updated "' + result.sheetName + '".\nIssues found: ' + result.issueCount + '\nWarnings: ' + result.warningCount + '\nErrors: ' + result.errorCount,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Readiness Report Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function show419AReadiness() {
  const ui = SpreadsheetApp.getUi();
  try {
    const summary = get419AReadinessSummary(CONFIG.ROLLOUT_ROOM);
    const lines = [
      'Room: ' + summary.room,
      'Storage locations: ' + summary.locationCount,
      'Locations with Storage ID: ' + summary.locationsWithStorageId,
      'Locations missing Storage ID: ' + summary.locationsMissingStorageId,
      'Inventory item rows: ' + summary.itemRows,
      'Empty storage placeholder rows: ' + summary.emptyStorageRows,
      'Chemical item rows: ' + summary.chemicalRows,
      'Legacy-only locations: ' + summary.legacyOnlyLocations,
      'Status counts: ' + JSON.stringify(summary.statusCounts)
    ];
    if (summary.sampleMissingStorageIds.length) {
      lines.push('Sample missing Storage ID: ' + summary.sampleMissingStorageIds.join(', '));
    }
    ui.alert('419A Readiness Summary', lines.join('\n'), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('419A Readiness Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function promptImport419AStorageMaster() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Import 419A Storage Master',
    'Paste the Google Sheet URL or ID that contains 419A_Storage_Master, Room_QR_Label_Plan, or RM 419A 2026:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const sourceId = extractSpreadsheetId_(response.getResponseText());
  if (!sourceId) {
    ui.alert('Import cancelled', 'No valid Google Sheet ID was provided.', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    'Confirm Storage Import',
    'This will append missing 419A storage placeholder rows only. Existing inventory rows will not be removed. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const result = import419AStorageMasterFromSource_(sourceId);
    ui.alert(
      '419A Storage Import Complete',
      'Imported storage placeholders: ' + result.importedRows +
        '\nSkipped existing: ' + result.skippedExisting +
        '\nSkipped invalid/blank: ' + result.skippedInvalid +
        '\nSource sheet: ' + result.sourceSheet,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('419A Storage Import Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function promptImport419AReady() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Import 419A App Load Ready',
    'Paste the Google Sheet URL or ID that contains 419A_App_Load_Ready / 419A App Load Ready:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const sourceId = extractSpreadsheetId_(response.getResponseText());
  if (!sourceId) {
    ui.alert('Import cancelled', 'No valid Google Sheet ID was provided.', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    'Confirm Import',
    'This will append non-duplicate 419A App Load Ready rows to the configured inventory sheet. It will not remove existing rows. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const result = import419AReadyFromSource_(sourceId);
    ui.alert(
      '419A Import Complete',
      'Imported: ' + result.importedRows +
        '\nSkipped duplicates: ' + result.skippedDuplicates +
        '\nSkipped invalid/blank rows: ' + result.skippedInvalid +
        '\nSource sheet: ' + result.sourceSheet,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('419A Import Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function promptSetAppConfig() {
  const ui = SpreadsheetApp.getUi();

  const spreadsheetId = ui.prompt(
    'Set SPREADSHEET_ID',
    'Paste the Google Sheet ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (spreadsheetId.getSelectedButton() !== ui.Button.OK) return;

  const webAppUrl = ui.prompt(
    'Set WEB_APP_BASE_URL',
    'Paste the deployed /exec URL. Leave blank to clear it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (webAppUrl.getSelectedButton() !== ui.Button.OK) return;

  const sheetName = ui.prompt(
    'Set INVENTORY_SHEET_NAME',
    'Paste the inventory sheet tab name. Leave blank to use the default fallback.',
    ui.ButtonSet.OK_CANCEL
  );
  if (sheetName.getSelectedButton() !== ui.Button.OK) return;

  setAppConfig(
    spreadsheetId.getResponseText(),
    webAppUrl.getResponseText(),
    sheetName.getResponseText()
  );

  ui.alert('App configuration saved.');
}

function promptSetWebAppBaseUrl() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Set WEB_APP_BASE_URL',
    'Paste your deployed /exec web app URL:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  setWebAppBaseUrl(result.getResponseText());
  ui.alert('WEB_APP_BASE_URL saved.');
}

function showConfigStatus() {
  const status = getConfigStatus();
  const ui = SpreadsheetApp.getUi();
  const missingRequired = Object.keys(status.requiredColumns || {}).filter(function (key) {
    return status.requiredColumns[key] === 'MISSING';
  });
  const optionalFound = Object.keys(status.optionalColumns || {}).filter(function (key) {
    return String(status.optionalColumns[key]).indexOf('found') === 0;
  });

  const lines = [
    'SPREADSHEET_ID: ' + (status.scriptProperties.spreadsheetIdConfigured ? 'configured' : 'missing'),
    'WEB_APP_BASE_URL: ' + (status.scriptProperties.webAppBaseUrlConfigured ? 'configured' : 'missing'),
    'INVENTORY_SHEET_NAME: ' + (status.scriptProperties.inventorySheetNameConfigured ? status.scriptProperties.inventorySheetName : 'default/fallback'),
    'Sheet in use: ' + (status.sheetInUse || '-'),
    'Data rows: ' + (typeof status.dataRows === 'number' ? status.dataRows : '-'),
    'Missing required columns: ' + (missingRequired.length ? missingRequired.join(', ') : 'none'),
    'Optional columns found: ' + (optionalFound.length ? optionalFound.join(', ') : 'none'),
    'Header preview: ' + ((status.headerPreview || []).slice(0, 12).join(', ') || '-')
  ];

  if (status.rollout419A) {
    lines.push('419A locations/items: ' + status.rollout419A.locationCount + ' location(s), ' + status.rollout419A.itemRows + ' item row(s)');
  }
  if (status.sheetError) lines.push('Sheet/config error: ' + status.sheetError);
  ui.alert('D&T Inventory Diagnostics', lines.join('\n'), ui.ButtonSet.OK);
}

function getRequestParams_(e) {
  return {
    room: cleanString_(e && e.parameter && e.parameter.room),
    loc: cleanString_(e && e.parameter && e.parameter.loc),
    mode: normalizeMode_(e && e.parameter && e.parameter.mode)
  };
}

function buildBootstrapData_(params) {
  const appConfig = getAppConfig_();
  const warnings = [];
  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  const diagnostics = getDiagnostics_();

  if (!webAppBaseUrl) {
    warnings.push('WEB_APP_BASE_URL is not configured. In-app browsing still works, but QR link generation and external shortcuts are disabled until this is set.');
  }

  const hasLocation = !!(params.room && params.loc);
  if (hasLocation) {
    const locationResult = getInventoryData({ room: params.room, loc: params.loc });
    return {
      pageType: 'location',
      appTitle: CONFIG.APP_TITLE,
      room: locationResult.room,
      loc: locationResult.loc,
      routeLoc: locationResult.routeLoc,
      displayLoc: locationResult.displayLoc,
      specificLocation: locationResult.specificLocation,
      storageId: locationResult.storageId,
      storageLabel: locationResult.storageLabel,
      locationCode: locationResult.locationCode,
      mode: params.mode,
      rows: locationResult.rows,
      locations: [],
      message: locationResult.message || '',
      error: '',
      warnings: warnings,
      config: appConfig,
      webAppBaseUrl: webAppBaseUrl,
      diagnostics: diagnostics
    };
  }

  const locationDirectory = getAllLocations();
  return {
    pageType: 'landing',
    appTitle: CONFIG.APP_TITLE,
    room: '',
    loc: '',
    routeLoc: '',
    displayLoc: '',
    specificLocation: '',
    storageId: '',
    storageLabel: '',
    locationCode: '',
    mode: 'view',
    rows: [],
    locations: locationDirectory.locations,
    message: locationDirectory.locations.length ? '' : 'No inventory locations found yet.',
    error: '',
    warnings: warnings,
    config: appConfig,
    webAppBaseUrl: webAppBaseUrl,
    diagnostics: diagnostics
  };
}

function buildErrorBootstrap_(params, err) {
  const message = err && err.message ? err.message : 'Unexpected application error.';
  return {
    pageType: 'error',
    appTitle: CONFIG.APP_TITLE,
    room: params.room || '',
    loc: params.loc || '',
    routeLoc: params.loc || '',
    displayLoc: params.loc || '',
    specificLocation: '',
    storageId: '',
    storageLabel: '',
    locationCode: '',
    mode: params.mode || 'view',
    rows: [],
    locations: [],
    message: '',
    error: 'Configuration error: ' + message,
    warnings: [],
    config: getAppConfigSafe_(),
    webAppBaseUrl: '',
    diagnostics: getDiagnosticsSafe_()
  };
}

function getInventoryData(params) {
  const room = cleanString_(params && params.room);
  const loc = cleanString_(params && params.loc);

  if (!room || !loc) {
    return {
      success: true,
      rows: [],
      room: room,
      loc: loc,
      routeLoc: loc,
      displayLoc: loc,
      specificLocation: '',
      storageId: '',
      storageLabel: '',
      locationCode: '',
      message: 'Please scan a valid QR code for a storage location.'
    };
  }

  const inventory = getLocationInventory_(room, loc);
  const context = inventory.context;

  if (!context) {
    return {
      success: true,
      rows: [],
      room: room,
      loc: loc,
      routeLoc: loc,
      displayLoc: loc,
      specificLocation: '',
      storageId: '',
      storageLabel: '',
      locationCode: '',
      message: 'No matching storage location was found for this room/location.'
    };
  }

  return {
    success: true,
    rows: inventory.rows,
    room: context.room,
    loc: context.routeLoc,
    routeLoc: context.routeLoc,
    displayLoc: context.displayLoc,
    specificLocation: context.loc,
    storageId: context.storageId,
    storageLabel: context.storageLabel,
    locationCode: context.locationCode,
    message: inventory.rows.length ? '' : 'No inventory items have been entered for this storage yet.'
  };
}

function saveInventoryUpdates(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid save payload.');
  }

  const room = cleanString_(payload.room);
  const loc = cleanString_(payload.loc);
  if (!room || !loc) {
    throw new Error('Room and location are required to save updates.');
  }

  if (!Array.isArray(payload.updates) || !payload.updates.length) {
    throw new Error('No updates provided.');
  }

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('The inventory sheet is empty.');
  }

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const roomNeedle = room.toLowerCase();
  const locNeedle = loc.toLowerCase();
  const lastRow = values.length;
  const seen = {};

  payload.updates.forEach(function (update) {
    const rowNum = Number(update.sheetRow);
    if (!Number.isInteger(rowNum) || rowNum <= CONFIG.HEADER_ROW || rowNum > lastRow) {
      throw new Error('Invalid row number: ' + update.sheetRow);
    }
    if (seen[rowNum]) {
      throw new Error('Duplicate row in payload: ' + rowNum);
    }
    seen[rowNum] = true;

    const row = values[rowNum - 1];
    if (!rowMatchesRoomLoc_(row, map, roomNeedle, locNeedle)) {
      throw new Error('Row ' + rowNum + ' does not belong to the selected room/location.');
    }

    if (update.qty === '' || update.qty == null) {
      throw new Error('Quantity is required for row ' + rowNum + '.');
    }

    const qty = Number(update.qty);
    if (!Number.isFinite(qty) || qty < 0) {
      throw new Error('Quantity must be a non-negative number for row ' + rowNum + '.');
    }

    const status = matchStatus_(update.status);
    if (!status) {
      throw new Error('Invalid status for row ' + rowNum + ': ' + cleanString_(update.status));
    }

    sheet.getRange(rowNum, map.qty + 1).setValue(qty);
    sheet.getRange(rowNum, map.status + 1).setValue(status);
  });

  const refreshedRows = getInventoryRowsForLocation_(room, loc);
  return {
    success: true,
    updatedCount: payload.updates.length,
    room: room,
    loc: loc,
    rows: refreshedRows,
    html: renderInventoryHtml_(refreshedRows, 'tech'),
    timestamp: new Date().toISOString()
  };
}

function getAllLocations() {
  return { success: true, locations: getAllLocations_() };
}

function getAllLocations_() {
  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const seen = {};
  const locations = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) continue;

    const entry = buildLocationEntry_(row, map, room, loc);
    const dedupeKey = entry.canonicalKey;

    if (seen[dedupeKey]) continue;
    seen[dedupeKey] = true;

    locations.push(entry);
  }

  locations.sort(function (a, b) {
    return a.sortKey.localeCompare(b.sortKey);
  });

  return locations;
}

function getLocationInventory_(room, loc) {
  const selectedRoom = cleanString_(room);
  const selectedLoc = cleanString_(loc);
  if (!selectedRoom || !selectedLoc) return { rows: [], context: null };

  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const roomNeedle = selectedRoom.toLowerCase();
  const locNeedle = selectedLoc.toLowerCase();
  const rows = [];
  let context = null;

  for (let i = 1; i < values.length; i++) {
    const sourceRow = values[i];
    const roomValue = cleanString_(sourceRow[map.room]);
    const locValue = cleanString_(sourceRow[map.location]);
    if (!roomValue) continue;
    if (!rowMatchesRoomLoc_(sourceRow, map, roomNeedle, locNeedle)) continue;
    if (!context) {
      context = buildLocationEntry_(sourceRow, map, roomValue, locValue);
    }

    if (!isInventoryItemRow_(sourceRow, map)) continue;

    rows.push(buildInventoryRowView_(
      sourceRow,
      map,
      i + 1,
      cleanString_(sourceRow[map.category]),
      roomValue,
      locValue
    ));
  }

  rows.sort(function (a, b) {
    return a.itemName.localeCompare(b.itemName);
  });

  return {
    rows: rows,
    context: context
  };
}

function getInventoryRowsForLocation_(room, loc) {
  return getLocationInventory_(room, loc).rows;
}

function buildLocationEntry_(row, map, room, loc) {
  const storageId = getOptionalValue_(row, map.storageId);
  const storageLabel = getOptionalValue_(row, map.storageLabel);
  const locationCode = getOptionalValue_(row, map.locationCode);
  const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);

  return {
    room: room,
    loc: loc,
    displayLoc: storageLabel || loc,
    locationCode: locationCode,
    storageId: storageId,
    storageLabel: storageLabel,
    routeLoc: identity.routeLoc,
    canonicalKey: identity.key,
    searchText: [room, loc, storageLabel, storageId, locationCode].filter(Boolean).join(' ').toLowerCase(),
    sortKey: [room, storageId || '', storageLabel || '', locationCode || '', loc].join(' | ').toLowerCase()
  };
}

function isInventoryItemRow_(row, map) {
  return !!(cleanString_(row[map.itemId]) || cleanString_(row[map.itemName]));
}

function rowMatchesRoomLoc_(rowValues, map, roomNeedle, locNeedle) {
  const rowRoom = cleanString_(rowValues[map.room]).toLowerCase();
  if (rowRoom !== roomNeedle) return false;

  const rowLoc = cleanString_(rowValues[map.location]).toLowerCase();
  if (rowLoc === locNeedle) return true;

  const rowStorageId = getOptionalValue_(rowValues, map.storageId).toLowerCase();
  if (rowStorageId && rowStorageId === locNeedle) return true;

  const rowStorageLabel = getOptionalValue_(rowValues, map.storageLabel).toLowerCase();
  if (rowStorageLabel && rowStorageLabel === locNeedle) return true;

  const rowLocationCode = getOptionalValue_(rowValues, map.locationCode).toLowerCase();
  if (rowLocationCode && rowLocationCode === locNeedle) return true;

  return false;
}

function buildInventoryRowView_(sourceRow, map, sheetRow, category, roomVal, locVal) {
  const status = normalizeStatus_(sourceRow[map.status]);
  const storageId = getOptionalValue_(sourceRow, map.storageId);
  const storageLabel = getOptionalValue_(sourceRow, map.storageLabel);
  const locationCode = getOptionalValue_(sourceRow, map.locationCode);
  return {
    sheetRow: sheetRow,
    itemId: cleanString_(sourceRow[map.itemId]),
    itemName: cleanString_(sourceRow[map.itemName]),
    room: roomVal,
    specificLocation: locVal,
    qty: toNonNegativeNumber_(sourceRow[map.qty]),
    category: category,
    status: status,
    unit: getOptionalValue_(sourceRow, map.unit),
    remarks: getOptionalValue_(sourceRow, map.remarks),
    locationCode: locationCode,
    storageId: storageId,
    storageLabel: storageLabel,
    isHazard: isHazardCategory_(category),
    statusClass: statusClassServer_(status),
    displayLocation: storageLabel || locVal,
    routeLoc: storageId || storageLabel || locationCode || locVal
  };
}

function refreshQrLinks() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) {
    return { success: true, updatedRows: 0, skippedRows: 0, baseUrl: getWebAppBaseUrl_() };
  }

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true });
  const baseUrl = getWebAppBaseUrl_();
  const rows = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, sheet.getLastColumn()).getValues();
  let updatedRows = 0;
  let skippedRows = 0;

  const output = rows.map(function (row) {
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) {
      skippedRows += 1;
      return [''];
    }

    const storageId = getOptionalValue_(row, map.storageId);
    const storageLabel = getOptionalValue_(row, map.storageLabel);
    const locationCode = getOptionalValue_(row, map.locationCode);
    const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
    updatedRows += 1;
    return [buildLocationUrl_(baseUrl, room, identity.routeLoc)];
  });

  sheet.getRange(CONFIG.HEADER_ROW + 1, map.qrLink + 1, output.length, 1).setValues(output);
  return { success: true, updatedRows: updatedRows, skippedRows: skippedRows, baseUrl: baseUrl };
}

function refreshQrImages() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return { success: true, updatedRows: 0 };

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true });
  if (map.qrImage === -1) {
    throw new Error('QR Code Image column is missing. Add an optional "QR Code Image" column before refreshing QR images.');
  }
  const qrLinkColA1 = columnLetter_(map.qrLink + 1);
  const qrImageCol = map.qrImage + 1;
  let updatedRows = 0;

  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const formula = '=IF(' + qrLinkColA1 + row + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(' + qrLinkColA1 + row + ')))';
    sheet.getRange(row, qrImageCol).setFormula(formula);
    updatedRows += 1;
  }

  return { success: true, updatedRows: updatedRows };
}

function ensureAppColumns_() {
  const sheet = getInventorySheet_();
  const desiredKeys = CONFIG.REQUIRED_COLUMNS.concat(['qrLink', 'unit', 'remarks', 'locationCode', 'storageId', 'storageLabel', 'qrImage']);
  const desiredLabels = desiredKeys.map(function (key) { return CONFIG.COLUMN_LABELS[key]; });
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headerRange = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn);
  const header = headerRange.getValues()[0];
  const hasAnyHeader = header.some(function (value) { return !!cleanString_(value); });

  if (!hasAnyHeader) {
    sheet.getRange(CONFIG.HEADER_ROW, 1, 1, desiredLabels.length).setValues([desiredLabels]);
    sheet.setFrozenRows(1);
    return { success: true, initialized: true, addedColumns: desiredLabels };
  }

  const normalizedHeaders = header.map(normalizeHeader_);
  const added = [];
  desiredKeys.forEach(function (key) {
    if (findHeaderIndex_(normalizedHeaders, CONFIG.ALIASES[key]) !== -1) return;
    added.push(CONFIG.COLUMN_LABELS[key]);
  });

  if (added.length) {
    sheet.getRange(CONFIG.HEADER_ROW, sheet.getLastColumn() + 1, 1, added.length).setValues([added]);
  }

  sheet.setFrozenRows(1);
  return { success: true, initialized: false, addedColumns: added };
}

function buildQrLabelSheet() {
  const baseUrl = getWebAppBaseUrl_();
  const locations = getAllLocations_();
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.QR_LABEL_SHEET_NAME);
  const headers = ['Room', 'Storage ID', 'Location Code', 'Display Location', 'Specific Location', 'QR Label', 'Web App Link', 'QR Image'];
  const output = locations.map(function (entry) {
    const label = entry.room + ' · ' + (entry.displayLoc || entry.loc);
    return [
      entry.room,
      entry.storageId,
      entry.locationCode,
      entry.displayLoc || entry.loc,
      entry.loc,
      label,
      buildLocationUrl_(baseUrl, entry.room, entry.routeLoc),
      ''
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (output.length) {
    sheet.getRange(2, 1, output.length, headers.length).setValues(output);
    for (let i = 0; i < output.length; i++) {
      const row = i + 2;
      sheet.getRange(row, 8).setFormula('=IF(G' + row + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(G' + row + ')))');
    }
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.setColumnWidths(1, 1, 80);
  sheet.setColumnWidths(2, 2, 140);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 240);
  sheet.setColumnWidth(7, 520);
  sheet.setColumnWidth(8, 120);

  return { success: true, sheetName: sheet.getName(), labelCount: output.length };
}

function createReadinessReport() {
  const report = validateInventoryData_();
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.READINESS_REPORT_SHEET_NAME);
  const now = new Date();
  const summaryRows = [
    ['Generated At', now],
    ['Inventory Sheet', report.sheetName],
    ['Data Rows', report.dataRows],
    ['Location Count', report.locationCount],
    ['Item Rows', report.itemRows],
    ['Empty Storage Rows', report.emptyStorageRows],
    ['Issue Count', report.issues.length],
    ['Error Count', report.errorCount],
    ['Warning Count', report.warningCount]
  ];
  const issueHeaders = ['Severity', 'Row', 'Room', 'Location', 'Storage ID', 'Item ID', 'Issue', 'Detail'];
  const issueRows = report.issues.map(function (issue) {
    return [issue.severity, issue.row, issue.room, issue.location, issue.storageId, issue.itemId, issue.issue, issue.detail];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Readiness Summary', 'Value']]);
  sheet.getRange(2, 1, summaryRows.length, 2).setValues(summaryRows);
  const issueStartRow = summaryRows.length + 4;
  sheet.getRange(issueStartRow, 1, 1, issueHeaders.length).setValues([issueHeaders]);
  if (issueRows.length) {
    sheet.getRange(issueStartRow + 1, 1, issueRows.length, issueHeaders.length).setValues(issueRows);
  } else {
    sheet.getRange(issueStartRow + 1, 1, 1, 1).setValue('No readiness issues found.');
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#dcfce7');
  sheet.getRange(issueStartRow, 1, 1, issueHeaders.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.setColumnWidths(1, 8, 150);
  sheet.setColumnWidth(8, 420);

  return {
    success: true,
    sheetName: sheet.getName(),
    issueCount: report.issues.length,
    warningCount: report.warningCount,
    errorCount: report.errorCount
  };
}

function get419AReadinessSummary(room) {
  const targetRoom = cleanString_(room || CONFIG.ROLLOUT_ROOM);
  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const seenLocations = {};
  const statusCounts = {};
  const sampleMissingStorageIds = [];
  let locationCount = 0;
  let locationsWithStorageId = 0;
  let locationsMissingStorageId = 0;
  let legacyOnlyLocations = 0;
  let itemRows = 0;
  let emptyStorageRows = 0;
  let chemicalRows = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowRoom = cleanString_(row[map.room]);
    if (rowRoom.toLowerCase() !== targetRoom.toLowerCase()) continue;

    const loc = cleanString_(row[map.location]);
    if (!loc) continue;

    const entry = buildLocationEntry_(row, map, rowRoom, loc);
    if (!seenLocations[entry.canonicalKey]) {
      seenLocations[entry.canonicalKey] = true;
      locationCount += 1;
      if (entry.storageId) {
        locationsWithStorageId += 1;
      } else {
        locationsMissingStorageId += 1;
        if (sampleMissingStorageIds.length < 5) sampleMissingStorageIds.push(entry.displayLoc || entry.loc);
      }
      if (!entry.storageId && !entry.storageLabel && !entry.locationCode) legacyOnlyLocations += 1;
    }

    if (isInventoryItemRow_(row, map)) {
      itemRows += 1;
      const status = normalizeStatus_(row[map.status]);
      statusCounts[status] = (statusCounts[status] || 0) + 1;
      if (isHazardCategory_(row[map.category])) chemicalRows += 1;
    } else {
      emptyStorageRows += 1;
    }
  }

  return {
    room: targetRoom,
    locationCount: locationCount,
    locationsWithStorageId: locationsWithStorageId,
    locationsMissingStorageId: locationsMissingStorageId,
    legacyOnlyLocations: legacyOnlyLocations,
    itemRows: itemRows,
    emptyStorageRows: emptyStorageRows,
    chemicalRows: chemicalRows,
    statusCounts: statusCounts,
    sampleMissingStorageIds: sampleMissingStorageIds
  };
}

function import419AReadyFromSource_(sourceSpreadsheetId) {
  const sourceId = extractSpreadsheetId_(sourceSpreadsheetId);
  if (!sourceId) throw new Error('A valid source spreadsheet ID is required.');

  ensureAppColumns_();
  const sourceSs = SpreadsheetApp.openById(sourceId);
  const sourceSheet = findSheetByNames_(sourceSs, CONFIG.IMPORT_READY_SHEET_NAMES);
  if (!sourceSheet) {
    throw new Error('Could not find 419A_App_Load_Ready / 419A App Load Ready in the source spreadsheet.');
  }

  const sourceValues = sourceSheet.getDataRange().getValues();
  const sourceHeaderRow = findHeaderRowIndex_(sourceValues, { requireQrLink: false });
  if (sourceHeaderRow === -1) {
    throw new Error('The source App Load Ready sheet does not contain the required inventory headers.');
  }

  const sourceMap = getColumnMap_(sourceValues[sourceHeaderRow], { requireQrLink: false });
  const storageLookup = buildStorageLookupFromSource_(sourceSs);
  const targetSheet = getInventorySheet_();
  const targetValues = targetSheet.getDataRange().getValues();
  if (!targetValues.length) throw new Error('The target inventory sheet is empty.');

  const targetHeader = targetValues[0];
  const targetMap = getColumnMap_(targetHeader, { requireQrLink: false });
  const existingKeys = buildExistingImportKeys_(targetValues, targetMap);
  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  const output = [];
  let skippedDuplicates = 0;
  let skippedInvalid = 0;

  for (let i = sourceHeaderRow + 1; i < sourceValues.length; i++) {
    const sourceRow = sourceValues[i];
    const itemId = cleanString_(sourceRow[sourceMap.itemId]);
    const itemName = cleanString_(sourceRow[sourceMap.itemName]);
    const room = cleanString_(sourceRow[sourceMap.room]) || CONFIG.ROLLOUT_ROOM;
    const loc = cleanString_(sourceRow[sourceMap.location]);
    if (!room || !loc || (!itemId && !itemName)) {
      skippedInvalid += 1;
      continue;
    }

    const sourceStorageId = getOptionalValue_(sourceRow, sourceMap.storageId);
    const storageMeta = storageLookup.byStorageId[buildLookupKey_(room, sourceStorageId)] ||
      storageLookup.byDisplayLocation[buildLookupKey_(room, loc)] ||
      {};
    const storageId = sourceStorageId || storageMeta.storageId || '';
    const locationCode = getOptionalValue_(sourceRow, sourceMap.locationCode) || storageMeta.locationCode || '';
    const storageLabel = getOptionalValue_(sourceRow, sourceMap.storageLabel) || storageMeta.storageLabel || '';
    const duplicateKeys = buildInventoryImportKeys_(itemId, itemName, room, loc, storageId, storageLabel, locationCode);
    const hasDuplicate = duplicateKeys.some(function (key) { return !!existingKeys[key]; });

    if (hasDuplicate) {
      skippedDuplicates += 1;
      continue;
    }

    const status = normalizeStatus_(sourceRow[sourceMap.status]);
    const qty = toNonNegativeNumber_(sourceRow[sourceMap.qty]);
    const rowOut = new Array(targetHeader.length).fill('');
    rowOut[targetMap.itemId] = itemId;
    rowOut[targetMap.itemName] = itemName;
    rowOut[targetMap.room] = room;
    rowOut[targetMap.location] = loc;
    rowOut[targetMap.qty] = qty;
    rowOut[targetMap.category] = cleanString_(sourceRow[sourceMap.category]);
    rowOut[targetMap.status] = status;

    setOptionalOutputValue_(rowOut, targetMap.unit, getOptionalValue_(sourceRow, sourceMap.unit));
    setOptionalOutputValue_(rowOut, targetMap.remarks, getOptionalValue_(sourceRow, sourceMap.remarks));
    setOptionalOutputValue_(rowOut, targetMap.locationCode, locationCode);
    setOptionalOutputValue_(rowOut, targetMap.storageId, storageId);
    setOptionalOutputValue_(rowOut, targetMap.storageLabel, storageLabel);
    if (targetMap.qrLink !== -1 && webAppBaseUrl) {
      rowOut[targetMap.qrLink] = buildLocationUrl_(webAppBaseUrl, room, buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode).routeLoc);
    }

    output.push(rowOut);
    duplicateKeys.forEach(function (key) { existingKeys[key] = true; });
  }

  if (output.length) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, output.length, targetHeader.length).setValues(output);
  }

  return {
    success: true,
    importedRows: output.length,
    skippedDuplicates: skippedDuplicates,
    skippedInvalid: skippedInvalid,
    sourceSheet: sourceSheet.getName()
  };
}

function import419AStorageMasterFromSource_(sourceSpreadsheetId) {
  const sourceId = extractSpreadsheetId_(sourceSpreadsheetId);
  if (!sourceId) throw new Error('A valid source spreadsheet ID is required.');

  ensureAppColumns_();
  const sourceSs = SpreadsheetApp.openById(sourceId);
  const source = getStorageMasterRowsFromSource_(sourceSs);
  if (!source.sheetName) {
    throw new Error('Could not find a supported 419A storage master sheet in the source spreadsheet.');
  }

  const targetSheet = getInventorySheet_();
  const targetValues = targetSheet.getDataRange().getValues();
  if (!targetValues.length) throw new Error('The target inventory sheet is empty.');

  const targetHeader = targetValues[0];
  const targetMap = getColumnMap_(targetHeader, { requireQrLink: false });
  const existingKeys = buildExistingLocationKeys_(targetValues, targetMap);
  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  const output = [];
  let skippedExisting = 0;
  let skippedInvalid = 0;

  source.rows.forEach(function (meta) {
    if (meta.room.toLowerCase() !== CONFIG.ROLLOUT_ROOM.toLowerCase()) return;
    if (!meta.displayLocation || !meta.storageId) {
      skippedInvalid += 1;
      return;
    }

    const identity = buildLocationIdentity_(meta.room, meta.displayLocation, meta.storageId, meta.storageLabel, meta.locationCode);
    const locationKeys = buildLocationMatchKeys_(meta.room, meta.displayLocation, meta.storageId, meta.storageLabel, meta.locationCode);
    const hasExisting = locationKeys.some(function (key) { return !!existingKeys[key]; });
    if (hasExisting) {
      skippedExisting += 1;
      return;
    }

    const rowOut = new Array(targetHeader.length).fill('');
    rowOut[targetMap.room] = meta.room;
    rowOut[targetMap.location] = meta.displayLocation;
    rowOut[targetMap.qty] = '';
    rowOut[targetMap.category] = '';
    rowOut[targetMap.status] = '';
    setOptionalOutputValue_(rowOut, targetMap.locationCode, meta.locationCode);
    setOptionalOutputValue_(rowOut, targetMap.storageId, meta.storageId);
    setOptionalOutputValue_(rowOut, targetMap.storageLabel, meta.storageLabel);
    if (targetMap.qrLink !== -1 && webAppBaseUrl) {
      rowOut[targetMap.qrLink] = buildLocationUrl_(webAppBaseUrl, meta.room, identity.routeLoc);
    }

    output.push(rowOut);
    locationKeys.forEach(function (key) { existingKeys[key] = true; });
  });

  if (output.length) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, output.length, targetHeader.length).setValues(output);
  }

  return {
    success: true,
    importedRows: output.length,
    skippedExisting: skippedExisting,
    skippedInvalid: skippedInvalid,
    sourceSheet: source.sheetName
  };
}

function getStorageMasterRowsFromSource_(sourceSs) {
  const sheet = findSheetByNames_(sourceSs, CONFIG.STORAGE_MASTER_SHEET_NAMES);
  const result = { sheetName: '', rows: [] };
  if (!sheet) return result;

  const values = sheet.getDataRange().getValues();
  const headerRow = findStorageHeaderRowIndex_(values);
  if (headerRow === -1) return result;

  const headers = values[headerRow].map(normalizeHeader_);
  const map = {
    room: findHeaderIndex_(headers, ['room']),
    storageId: findHeaderIndex_(headers, ['storage id', 'storageid']),
    displayLocation: findHeaderIndex_(headers, ['display location', 'specific location', 'location']),
    locationCode: findHeaderIndex_(headers, ['location code', 'locationcode']),
    storageLabel: findHeaderIndex_(headers, ['storage label', 'storagelabel'])
  };

  for (let i = headerRow + 1; i < values.length; i++) {
    const row = values[i];
    const room = getOptionalValue_(row, map.room) || CONFIG.ROLLOUT_ROOM;
    const displayLocation = getOptionalValue_(row, map.displayLocation);
    const storageId = getOptionalValue_(row, map.storageId);
    if (!room || (!displayLocation && !storageId)) continue;
    result.rows.push({
      room: room,
      displayLocation: displayLocation,
      storageId: storageId,
      locationCode: getOptionalValue_(row, map.locationCode),
      storageLabel: getOptionalValue_(row, map.storageLabel)
    });
  }

  result.sheetName = sheet.getName();
  return result;
}

function buildStorageLookupFromSource_(sourceSs) {
  const source = getStorageMasterRowsFromSource_(sourceSs);
  const lookup = { byDisplayLocation: {}, byStorageId: {} };
  source.rows.forEach(function (meta) {
    if (meta.displayLocation) lookup.byDisplayLocation[buildLookupKey_(meta.room, meta.displayLocation)] = meta;
    if (meta.storageId) lookup.byStorageId[buildLookupKey_(meta.room, meta.storageId)] = meta;
  });

  return lookup;
}

function buildExistingImportKeys_(values, map) {
  const existing = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!isInventoryItemRow_(row, map)) continue;
    const itemId = cleanString_(row[map.itemId]);
    const itemName = cleanString_(row[map.itemName]);
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc || (!itemId && !itemName)) continue;
    buildInventoryImportKeys_(
      itemId,
      itemName,
      room,
      loc,
      getOptionalValue_(row, map.storageId),
      getOptionalValue_(row, map.storageLabel),
      getOptionalValue_(row, map.locationCode)
    ).forEach(function (key) {
      existing[key] = true;
    });
  }
  return existing;
}

function buildExistingLocationKeys_(values, map) {
  const existing = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) continue;
    buildLocationMatchKeys_(
      room,
      loc,
      getOptionalValue_(row, map.storageId),
      getOptionalValue_(row, map.storageLabel),
      getOptionalValue_(row, map.locationCode)
    ).forEach(function (key) {
      existing[key] = true;
    });
  }
  return existing;
}

function validateInventoryData_() {
  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('The inventory sheet is empty.');
  const map = getLooseColumnMap_(values[0]);
  const issues = [];
  const locationCounts = {};
  const itemCounts = {};
  let itemRows = 0;
  let emptyStorageRows = 0;
  let errorCount = 0;
  let warningCount = 0;

  function addIssue(severity, rowIndex, row, issue, detail) {
    if (severity === 'ERROR') errorCount += 1;
    else warningCount += 1;
    issues.push({
      severity: severity,
      row: rowIndex,
      room: row ? cleanString_(row[map.room]) : '',
      location: row ? cleanString_(row[map.location]) : '',
      storageId: row ? getOptionalValue_(row, map.storageId) : '',
      itemId: row ? cleanString_(row[map.itemId]) : '',
      issue: issue,
      detail: detail
    });
  }

  CONFIG.REQUIRED_COLUMNS.forEach(function (key) {
    if (map[key] === -1) {
      addIssue('ERROR', '', null, 'MISSING_COLUMN', CONFIG.COLUMN_LABELS[key] + ' is missing.');
    }
  });

  const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetRow = i + 1;
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    const itemId = cleanString_(row[map.itemId]);
    const itemName = cleanString_(row[map.itemName]);
    const storageId = getOptionalValue_(row, map.storageId);
    const storageLabel = getOptionalValue_(row, map.storageLabel);
    const locationCode = getOptionalValue_(row, map.locationCode);
    const isItem = isInventoryItemRow_(row, map);

    if (!room && !loc && !itemId && !itemName) continue;
    if (!room) addIssue('ERROR', sheetRow, row, 'MISSING_ROOM', 'Room is required for browsing and QR links.');
    if (!loc) addIssue('ERROR', sheetRow, row, 'MISSING_LOCATION', 'Specific Location is required for browsing and QR links.');

    if (room && loc) {
      const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
      locationCounts[identity.key] = (locationCounts[identity.key] || 0) + 1;
      if (room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase() && !storageId) {
        addIssue('WARN', sheetRow, row, '419A_MISSING_STORAGE_ID', '419A rollout rows should use Storage ID where possible.');
      }
      if (map.qrLink !== -1 && webAppBaseUrl) {
        const expected = buildLocationUrl_(webAppBaseUrl, room, identity.routeLoc);
        const current = getOptionalValue_(row, map.qrLink);
        if (!current) addIssue('WARN', sheetRow, row, 'QR_LINK_MISSING', 'QR link is blank.');
        else if (current !== expected) addIssue('WARN', sheetRow, row, 'QR_LINK_OUTDATED', 'Expected ' + expected);
      }
    }

    if (isItem) {
      itemRows += 1;
      const qtyText = cleanString_(row[map.qty]);
      const qty = Number(row[map.qty]);
      if (qtyText === '' || !Number.isFinite(qty) || qty < 0) {
        addIssue('ERROR', sheetRow, row, 'INVALID_QTY', 'Quantity must be numeric and non-negative.');
      }
      if (!matchStatus_(row[map.status])) {
        addIssue('ERROR', sheetRow, row, 'INVALID_STATUS', 'Allowed values: ' + CONFIG.STATUS_OPTIONS.join(', '));
      }
      if (!itemName) addIssue('WARN', sheetRow, row, 'MISSING_ITEM_NAME', 'Item rows should include Item Name.');
      const itemKey = buildInventoryImportKey_(itemId, itemName, room, loc, storageId, storageLabel, locationCode);
      itemCounts[itemKey] = (itemCounts[itemKey] || 0) + 1;
      if (itemCounts[itemKey] > 1) {
        addIssue('WARN', sheetRow, row, 'POSSIBLE_DUPLICATE_ITEM', 'Same item identity appears more than once in this storage.');
      }
    } else if (room && loc) {
      emptyStorageRows += 1;
    }
  }

  Object.keys(locationCounts).forEach(function (key) {
    if (locationCounts[key] > 20) {
      issues.push({
        severity: 'INFO',
        row: '',
        room: '',
        location: '',
        storageId: '',
        itemId: '',
        issue: 'HIGH_ROW_COUNT_LOCATION',
        detail: key + ' has ' + locationCounts[key] + ' row(s).'
      });
    }
  });

  return {
    sheetName: sheet.getName(),
    dataRows: Math.max(values.length - CONFIG.HEADER_ROW, 0),
    locationCount: Object.keys(locationCounts).length,
    itemRows: itemRows,
    emptyStorageRows: emptyStorageRows,
    issues: issues,
    errorCount: errorCount,
    warningCount: warningCount
  };
}

function buildInventoryImportKey_(itemId, itemName, room, loc, storageId, storageLabel, locationCode) {
  const itemKey = cleanString_(itemId || itemName).toLowerCase();
  const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
  return [cleanString_(room).toLowerCase(), identity.routeLoc.toLowerCase(), itemKey].join('||');
}

function buildInventoryImportKeys_(itemId, itemName, room, loc, storageId, storageLabel, locationCode) {
  const itemKey = cleanString_(itemId || itemName).toLowerCase();
  return buildLocationMatchKeys_(room, loc, storageId, storageLabel, locationCode).map(function (locationKey) {
    return locationKey + '||item||' + itemKey;
  });
}

function buildLocationMatchKeys_(room, loc, storageId, storageLabel, locationCode) {
  const roomKey = cleanString_(room).toLowerCase();
  const keys = {};
  [
    ['location', loc],
    ['storageId', storageId],
    ['storageLabel', storageLabel],
    ['locationCode', locationCode]
  ].forEach(function (pair) {
    const value = cleanString_(pair[1]);
    if (value) keys[[roomKey, pair[0], value.toLowerCase()].join('||')] = true;
  });
  keys[buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode).key] = true;
  return Object.keys(keys);
}

function findHeaderRowIndex_(values, options) {
  for (let i = 0; i < values.length; i++) {
    try {
      getColumnMap_(values[i], options || {});
      return i;
    } catch (err) {
      // Keep scanning: import workbooks often include title/instruction rows above headers.
    }
  }
  return -1;
}

function findStorageHeaderRowIndex_(values) {
  for (let i = 0; i < values.length; i++) {
    const headers = values[i].map(normalizeHeader_);
    const hasStorageId = findHeaderIndex_(headers, ['storage id', 'storageid']) !== -1;
    const hasLocation = findHeaderIndex_(headers, ['display location', 'specific location', 'location']) !== -1;
    if (hasStorageId && hasLocation) return i;
  }
  return -1;
}

function findSheetByNames_(spreadsheet, names) {
  const normalizedNames = {};
  names.forEach(function (name) {
    normalizedNames[normalizeHeader_(name)] = true;
  });

  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (normalizedNames[normalizeHeader_(sheets[i].getName())]) return sheets[i];
  }
  return null;
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function setOptionalOutputValue_(rowOut, index, value) {
  if (typeof index === 'number' && index >= 0) rowOut[index] = value;
}

function buildLookupKey_(room, value) {
  return [cleanString_(room).toLowerCase(), cleanString_(value).toLowerCase()].join('||');
}

function extractSpreadsheetId_(input) {
  const value = cleanString_(input);
  if (!value) return '';
  const match = value.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match && match[1] ? match[1] : value;
}

function setAppConfig(spreadsheetId, webAppBaseUrl, inventorySheetName) {
  const props = PropertiesService.getScriptProperties();
  const nextSpreadsheetId = cleanString_(spreadsheetId);
  const nextWebAppBaseUrl = normalizeWebAppBaseUrl_(webAppBaseUrl);
  const nextSheetName = cleanString_(inventorySheetName);

  if (!nextSpreadsheetId) {
    throw new Error('Please provide a non-empty spreadsheet ID.');
  }

  props.setProperty(CONFIG.SPREADSHEET_ID_PROPERTY, nextSpreadsheetId);
  if (nextWebAppBaseUrl) props.setProperty(CONFIG.WEB_APP_URL_PROPERTY, nextWebAppBaseUrl);
  else props.deleteProperty(CONFIG.WEB_APP_URL_PROPERTY);

  if (nextSheetName) props.setProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY, nextSheetName);
  else props.deleteProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY);

  return getAppConfig_();
}

function setWebAppBaseUrl(url) {
  const value = normalizeWebAppBaseUrl_(url);
  if (!value) throw new Error('Please provide a non-empty URL.');
  PropertiesService.getScriptProperties().setProperty(CONFIG.WEB_APP_URL_PROPERTY, value);
}

function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(CONFIG.SPREADSHEET_ID_PROPERTY);
  if (!spreadsheetId || !spreadsheetId.trim()) {
    throw new Error('SPREADSHEET_ID is not set.');
  }
  return SpreadsheetApp.openById(spreadsheetId.trim());
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const configuredName = PropertiesService.getScriptProperties().getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY);
  const preferredNames = [];

  if (configuredName && configuredName.trim()) preferredNames.push(configuredName.trim());
  preferredNames.push(CONFIG.DEFAULT_SHEET_NAME);

  for (var i = 0; i < preferredNames.length; i++) {
    const sheet = ss.getSheetByName(preferredNames[i]);
    if (sheet) return sheet;
  }

  const first = ss.getSheets()[0];
  if (!first) throw new Error('No sheets found in the configured spreadsheet.');
  return first;
}

function getInventoryDataset_() {
  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    throw new Error('The inventory sheet is empty.');
  }
  return {
    sheet: sheet,
    values: values,
    map: getColumnMap_(values[0], { requireQrLink: false })
  };
}

function getColumnMap_(headerRow, options) {
  const opts = options || {};
  const map = getLooseColumnMap_(headerRow);

  const required = CONFIG.REQUIRED_COLUMNS.slice();
  if (opts.requireQrLink) required.push('qrLink');
  if (opts.requireQrImage) required.push('qrImage');

  const missing = required.filter(function (key) {
    return map[key] === -1;
  });

  if (missing.length) {
    throw new Error('Required column missing: ' + missing.map(function (key) {
      return CONFIG.ALIASES[key][0];
    }).join(', '));
  }

  return map;
}

function getLooseColumnMap_(headerRow) {
  const headers = (headerRow || []).map(normalizeHeader_);
  const map = {};

  Object.keys(CONFIG.ALIASES).forEach(function (key) {
    map[key] = findHeaderIndex_(headers, CONFIG.ALIASES[key]);
  });

  return map;
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .trim();
}

function findHeaderIndex_(normalizedHeaders, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const alias = normalizeHeader_(aliases[i]);
    const idx = normalizedHeaders.indexOf(alias);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getWebAppBaseUrl_(options) {
  const opts = options || {};
  const propertyUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.WEB_APP_URL_PROPERTY);
  if (propertyUrl && propertyUrl.trim()) return normalizeWebAppBaseUrl_(propertyUrl);

  if (opts.silent) return '';
  throw new Error('WEB_APP_BASE_URL is not configured.');
}

function getAppConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: cleanString_(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY)),
    inventorySheetName: cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY)),
    webAppBaseUrl: cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY))
  };
}

function getAppConfigSafe_() {
  try {
    return getAppConfig_();
  } catch (err) {
    return {
      spreadsheetId: '',
      inventorySheetName: '',
      webAppBaseUrl: ''
    };
  }
}

function getDiagnostics_() {
  const status = getConfigStatus();
  return status;
}

function getDiagnosticsSafe_() {
  try {
    return getDiagnostics_();
  } catch (err) {
    return { error: err.message || String(err) };
  }
}

function getConfigStatus() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = cleanString_(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY));
  const webAppBaseUrl = cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY));
  const inventorySheetName = cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));

  const status = {
    scriptProperties: {
      spreadsheetIdConfigured: !!spreadsheetId,
      webAppBaseUrlConfigured: !!webAppBaseUrl,
      inventorySheetNameConfigured: !!inventorySheetName,
      inventorySheetName: inventorySheetName || CONFIG.DEFAULT_SHEET_NAME
    }
  };

  try {
    const sheet = getInventorySheet_();
    const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const normalizedHeaders = header.map(normalizeHeader_);
    const map = {};

    Object.keys(CONFIG.ALIASES).forEach(function (key) {
      map[key] = findHeaderIndex_(normalizedHeaders, CONFIG.ALIASES[key]);
    });

    status.sheetInUse = sheet.getName();
    status.dataRows = Math.max(sheet.getLastRow() - CONFIG.HEADER_ROW, 0);
    status.requiredColumns = {};
    status.optionalColumns = {};

    CONFIG.REQUIRED_COLUMNS.forEach(function (key) {
      status.requiredColumns[key] = map[key] !== -1 ? 'found (col ' + (map[key] + 1) + ')' : 'MISSING';
    });

    CONFIG.OPTIONAL_COLUMNS.forEach(function (key) {
      status.optionalColumns[key] = map[key] !== -1 ? 'found (col ' + (map[key] + 1) + ')' : 'not present';
    });

    status.headerPreview = normalizedHeaders;
    status.rollout419A = get419AReadinessSummary(CONFIG.ROLLOUT_ROOM);
  } catch (err) {
    status.sheetError = err.message || String(err);
  }

  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode) {
  const routeLoc = storageId || storageLabel || locationCode || loc;
  const identityType = storageId ? 'storageId' : (storageLabel ? 'storageLabel' : (locationCode ? 'locationCode' : 'location'));
  const identityValue = storageId || storageLabel || locationCode || loc;
  const key = [room, identityType, identityValue].join('||').toLowerCase();
  return {
    routeLoc: routeLoc,
    key: key
  };
}

function buildLocationUrl_(baseUrl, room, loc) {
  return normalizeWebAppBaseUrl_(baseUrl) + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function buildLocationHref_(baseUrl, room, loc, mode) {
  const route = '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
  const href = baseUrl ? buildLocationUrl_(baseUrl, room, loc) : route;
  return mode === 'tech' ? href + '&mode=tech' : href;
}

function normalizeMode_(value) {
  return cleanString_(value).toLowerCase() === 'tech' ? 'tech' : 'view';
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function normalizeWebAppBaseUrl_(value) {
  return cleanString_(value).replace(/[?#].*$/, '').replace(/\?+$/, '');
}

function getOptionalValue_(row, index) {
  if (typeof index !== 'number' || index < 0) return '';
  return cleanString_(row[index]);
}

function normalizeStatus_(value) {
  return matchStatus_(value) || 'Good';
}

function matchStatus_(value) {
  const raw = cleanString_(value);
  return CONFIG.STATUS_OPTIONS.filter(function (status) {
    return status.toLowerCase() === raw.toLowerCase();
  })[0] || '';
}

function toNonNegativeNumber_(value) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : 0;
}

function isHazardCategory_(category) {
  return CONFIG.HAZARD_CATEGORIES.indexOf(cleanString_(category).toLowerCase()) !== -1;
}

function statusClassServer_(status) {
  switch (cleanString_(status)) {
    case 'Good':
      return 'text-emerald-700 bg-emerald-100';
    case 'Low Stock':
      return 'text-amber-700 bg-amber-100';
    case 'Missing':
      return 'text-red-700 bg-red-100';
    case 'Needs Maintenance':
      return 'text-orange-800 bg-orange-100';
    default:
      return 'text-stone-700 bg-stone-100';
  }
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function columnLetter_(indexOneBased) {
  let n = indexOneBased;
  let result = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    n = Math.floor((n - mod) / 26);
  }
  return result;
}

function renderInitialItemsHtml_(bootstrap) {
  if (bootstrap.pageType === 'landing') {
    return renderLandingHtml_(bootstrap.locations || [], bootstrap.webAppBaseUrl);
  }
  if (bootstrap.pageType === 'error') {
    return renderErrorStateHtml_(bootstrap.error);
  }
  return renderInventoryHtml_(bootstrap.rows || [], bootstrap.mode);
}

function renderErrorStateHtml_(message) {
  return '<div class="p-5"><div class="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-800"><p class="font-semibold">Setup needs attention</p><p class="mt-1">' + escapeHtml_(message) + '</p></div></div>';
}

function buildLocationStats_(locations) {
  const rooms = {};
  let storageIdCount = 0;
  let rolloutRoomCount = 0;

  locations.forEach(function (entry) {
    rooms[entry.room] = true;
    if (entry.storageId) storageIdCount += 1;
    if (entry.room === CONFIG.ROLLOUT_ROOM) rolloutRoomCount += 1;
  });

  return {
    roomCount: Object.keys(rooms).length,
    locationCount: locations.length,
    storageIdCount: storageIdCount,
    rolloutRoomCount: rolloutRoomCount
  };
}

function renderLandingHtml_(locations, webAppBaseUrl) {
  const stats = buildLocationStats_(locations);
  let html = '';
  html += '<div class="border-b border-stone-200 bg-stone-50 px-4 py-4 sm:px-5">';
  html += '<div class="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">';
  html += '<div><h2 class="text-base font-semibold text-stone-900">All storage locations</h2>';
  html += '<p class="mt-1 text-sm text-stone-600">Search by room, display location, storage ID, storage label, or location code.</p></div>';
  html += '<div class="grid grid-cols-2 gap-2 text-right sm:grid-cols-4">';
  html += '<span class="rounded-xl border border-stone-200 bg-white px-3 py-2"><span class="block text-[11px] uppercase tracking-wide text-stone-500">Rooms</span><span class="text-sm font-semibold text-stone-900">' + stats.roomCount + '</span></span>';
  html += '<span class="rounded-xl border border-stone-200 bg-white px-3 py-2"><span class="block text-[11px] uppercase tracking-wide text-stone-500">Locations</span><span class="text-sm font-semibold text-stone-900">' + stats.locationCount + '</span></span>';
  html += '<span class="rounded-xl border border-stone-200 bg-white px-3 py-2"><span class="block text-[11px] uppercase tracking-wide text-stone-500">Storage IDs</span><span class="text-sm font-semibold text-stone-900">' + stats.storageIdCount + '</span></span>';
  html += '<span class="rounded-xl border border-teal-200 bg-teal-50 px-3 py-2"><span class="block text-[11px] uppercase tracking-wide text-teal-700">419A</span><span class="text-sm font-semibold text-teal-900">' + stats.rolloutRoomCount + '</span></span>';
  html += '</div></div>';
  html += '<div class="mt-4"><label class="sr-only" for="locationSearch">Search room, location, storage ID, or code</label><input id="locationSearch" type="search" placeholder="Search 419A, storage ID, label, location code..." class="w-full rounded-xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-900 shadow-sm outline-none transition focus:border-sky-500 focus:ring-2 focus:ring-sky-200" /></div>';
  html += '</div>';

  if (!locations.length) {
    html += '<div class="p-5"><div class="rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-5 text-sm text-stone-600"><p class="font-semibold text-stone-900">No storage locations found yet.</p><p class="mt-1">Add rows with Room and Specific Location, then refresh this page.</p></div></div>';
    return html;
  }

  html += '<div id="locationResults">';
  let currentRoom = '';
  locations.forEach(function (entry) {
    if (entry.room !== currentRoom) {
      currentRoom = entry.room;
      html += '<div class="room-group-header border-y border-stone-200 bg-stone-100 px-4 py-2 text-[11px] font-semibold uppercase tracking-[0.18em] text-stone-600" data-room-group="' + escapeHtml_(entry.room) + '">Room ' + escapeHtml_(entry.room) + '</div>';
    }

    const viewUrl = buildLocationHref_(webAppBaseUrl, entry.room, entry.routeLoc, 'view');
    const techUrl = buildLocationHref_(webAppBaseUrl, entry.room, entry.routeLoc, 'tech');
    const displayName = entry.displayLoc || entry.loc;

    html += '<article class="location-card border-b border-stone-200 px-4 py-4 last:border-b-0 ' + (entry.room === CONFIG.ROLLOUT_ROOM ? 'bg-teal-50/40' : 'bg-white') + '" data-search="' + escapeHtml_(entry.searchText) + '" data-room="' + escapeHtml_(entry.room) + '">';
    html += '<div class="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">';
    html += '<div class="min-w-0">';
    html += '<p class="text-base font-semibold text-stone-900">' + escapeHtml_(displayName) + '</p>';
    if (entry.displayLoc !== entry.loc) {
      html += '<p class="mt-0.5 text-xs text-stone-400 italic">Location: ' + escapeHtml_(entry.loc) + '</p>';
    }
    html += '<p class="mt-1 text-xs text-stone-500">Room: ' + escapeHtml_(entry.room) + '</p>';
    if (entry.storageId || entry.locationCode) {
      html += '<div class="mt-2 flex flex-wrap gap-1.5">';
      if (entry.storageId) {
        html += '<span class="inline-flex rounded-full bg-teal-100 px-2.5 py-1 text-[11px] font-semibold text-teal-800">ID: ' + escapeHtml_(entry.storageId) + '</span>';
      }
      if (entry.locationCode) {
        html += '<span class="inline-flex rounded-full bg-sky-100 px-2.5 py-1 text-[11px] font-medium text-sky-800">Code: ' + escapeHtml_(entry.locationCode) + '</span>';
      }
      html += '</div>';
    }
    html += '</div>';
    html += '<div class="flex shrink-0 gap-2 sm:flex-row">';
    html += '<a class="rounded-xl bg-stone-200 px-3 py-2 text-xs font-medium text-stone-800 transition hover:bg-stone-300" href="' + escapeHtml_(viewUrl) + '">View items</a>';
    html += '<a class="rounded-xl bg-sky-700 px-3 py-2 text-xs font-medium text-white transition hover:bg-sky-800" href="' + escapeHtml_(techUrl) + '">Tech update</a>';
    html += '</div></div></article>';
  });
  html += '<p id="locationNoResults" class="hidden p-5 text-sm text-stone-500">No matching storage locations found.</p>';
  html += '</div>';
  return html;
}

function renderInventoryHtml_(rows, mode) {
  if (!rows.length) {
    return '<div class="p-5"><div class="rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-5 text-sm text-stone-600"><p class="font-semibold text-stone-900">No items entered yet</p><p class="mt-1">No inventory items have been entered for this storage yet.</p></div></div>';
  }

  let html = '';
  const itemStats = buildItemStats_(rows);
  html += '<div class="border-b border-stone-200 bg-stone-50 px-4 py-3">';
  html += '<div class="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">';
  html += '<div class="flex flex-wrap gap-2 text-xs">';
  html += '<span class="rounded-full bg-white px-2.5 py-1 font-medium text-stone-700 ring-1 ring-stone-200">' + itemStats.total + ' item(s)</span>';
  html += '<span class="rounded-full bg-red-50 px-2.5 py-1 font-medium text-red-700 ring-1 ring-red-200">' + itemStats.hazard + ' chemical</span>';
  html += '<span class="rounded-full bg-amber-50 px-2.5 py-1 font-medium text-amber-800 ring-1 ring-amber-200">' + itemStats.attention + ' need attention</span>';
  html += '</div>';
  html += '<div class="grid grid-cols-1 gap-2 sm:grid-cols-[minmax(0,1fr)_160px] lg:min-w-[520px]">';
  html += '<label class="sr-only" for="itemSearch">Search items</label><input id="itemSearch" type="search" placeholder="Search item, ID, category, remarks..." class="rounded-xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200" />';
  html += '<label class="sr-only" for="statusFilter">Filter status</label><select id="statusFilter" class="rounded-xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200"><option value="">All statuses</option>' + CONFIG.STATUS_OPTIONS.map(function (status) { return '<option value="' + escapeHtml_(status.toLowerCase()) + '">' + escapeHtml_(status) + '</option>'; }).join('') + '</select>';
  html += '</div></div></div>';
  rows.forEach(function (row) {
    const hazardBadge = row.isHazard
      ? '<span class="inline-flex rounded-full bg-red-200 px-2.5 py-1 text-[11px] font-semibold uppercase tracking-wide text-red-900">Hazard: Chemical</span>'
      : '';
    const hazardNote = row.isHazard
      ? '<p class="mt-3 rounded-xl border border-red-200 bg-red-100 px-3 py-2 text-xs font-medium text-red-900">Chemical item: handle according to D&T chemical storage and safety procedures.</p>'
      : '';

    const meta = [
      'ID: ' + escapeHtml_(row.itemId),
      'Category: ' + escapeHtml_(row.category)
    ];
    if (row.unit) meta.push('Unit: ' + escapeHtml_(row.unit));
    if (row.storageId) meta.push('Storage ID: ' + escapeHtml_(row.storageId));
    if (row.locationCode) meta.push('Code: ' + escapeHtml_(row.locationCode));

    const searchText = [
      row.itemId,
      row.itemName,
      row.category,
      row.status,
      row.unit,
      row.remarks,
      row.storageId,
      row.storageLabel,
      row.locationCode
    ].filter(Boolean).join(' ').toLowerCase();

    html += '<article class="inventory-card border-b border-stone-200 p-4 last:border-b-0 ' + (row.isHazard ? 'border-l-4 border-l-red-400 bg-red-50' : 'bg-white') + '" data-item-search="' + escapeHtml_(searchText) + '" data-status="' + escapeHtml_(row.status.toLowerCase()) + '">';
    html += '<div class="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">';
    html += '<div class="min-w-0 flex-1">';

    const storageBadge = row.storageLabel
      ? '<span class="ml-1 inline-flex rounded-full bg-teal-100 px-2 py-0.5 text-[10px] font-semibold text-teal-800">' + escapeHtml_(row.storageLabel) + '</span>'
      : '';

    html += '<div class="flex flex-wrap items-center gap-2"><h3 class="text-base font-semibold text-stone-900">' + escapeHtml_(row.itemName || '(Unnamed item)') + '</h3>' + hazardBadge + storageBadge + '</div>';
    html += '<p class="mt-1 text-xs text-stone-500">' + meta.join(' · ') + '</p>';

    if (row.remarks) {
      html += '<p class="mt-2 text-sm text-stone-700"><span class="font-medium text-stone-900">Remarks:</span> ' + escapeHtml_(row.remarks) + '</p>';
    }

    html += hazardNote;
    html += '</div>';

    if (mode === 'tech') {
      const opts = CONFIG.STATUS_OPTIONS.map(function (status) {
        return '<option value="' + escapeHtml_(status) + '" ' + (status === row.status ? 'selected' : '') + '>' + escapeHtml_(status) + '</option>';
      }).join('');

      html += '<div class="grid grid-cols-1 gap-3 sm:min-w-[240px]">';
      if (row.storageId) {
        html += '<p class="text-xs text-stone-500"><span class="font-medium text-stone-700">Storage ID:</span> ' + escapeHtml_(row.storageId) + '</p>';
      }
      if (row.unit) {
        html += '<p class="text-xs text-stone-500"><span class="font-medium text-stone-700">Unit:</span> ' + escapeHtml_(row.unit) + '</p>';
      }
      html += '<label class="block"><span class="text-xs font-medium uppercase tracking-wide text-stone-500">Quantity</span><input type="number" min="0" step="any" inputmode="decimal" class="qty-input mt-1 w-full rounded-xl border border-stone-300 px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200" data-row="' + escapeHtml_(String(row.sheetRow)) + '" value="' + escapeHtml_(String(row.qty)) + '" /></label>';
      html += '<label class="block"><span class="text-xs font-medium uppercase tracking-wide text-stone-500">Status</span><select class="status-input mt-1 w-full rounded-xl border border-stone-300 px-3 py-2 text-sm text-stone-900 outline-none focus:border-sky-500 focus:ring-2 focus:ring-sky-200" data-row="' + escapeHtml_(String(row.sheetRow)) + '">' + opts + '</select></label>';
      html += '<p class="inline-flex w-fit rounded-full px-2.5 py-1 text-xs font-medium ' + row.statusClass + '">' + escapeHtml_(row.status) + '</p>';
      html += '</div>';
    } else {
      html += '<div class="rounded-2xl bg-stone-100 px-4 py-3 text-right sm:min-w-[170px]">';
      html += '<p class="text-[11px] font-medium uppercase tracking-wide text-stone-500">Expected Qty</p>';
      html += '<p class="mt-1 text-3xl font-bold text-stone-900">' + escapeHtml_(String(row.qty)) + '</p>';
      if (row.unit) {
        html += '<p class="mt-1 text-xs text-stone-500">' + escapeHtml_(row.unit) + '</p>';
      }
      html += '<p class="mt-2 inline-flex rounded-full px-2.5 py-1 text-xs font-medium ' + row.statusClass + '">' + escapeHtml_(row.status) + '</p>';
      html += '</div>';
    }

    html += '</div></article>';
  });
  html += '<p id="itemNoResults" class="hidden p-5 text-sm text-stone-500">No matching inventory items found.</p>';
  return html;
}

function buildItemStats_(rows) {
  const stats = { total: rows.length, hazard: 0, attention: 0 };
  rows.forEach(function (row) {
    if (row.isHazard) stats.hazard += 1;
    if (row.status !== 'Good') stats.attention += 1;
  });
  return stats;
}

function buildPageHtml_(params, bootstrap) {
  const room = bootstrap.room || params.room || '';
  const routeLoc = bootstrap.routeLoc || bootstrap.loc || params.loc || '';
  const displayLoc = bootstrap.displayLoc || bootstrap.loc || params.loc || '';
  const mode = bootstrap.mode === 'tech' ? 'tech' : 'view';
  const initialHtml = renderInitialItemsHtml_(bootstrap);
  const isLocationPage = bootstrap.pageType === 'location';
  const hasEditableRows = isLocationPage && bootstrap.rows && bootstrap.rows.length > 0;
  const modeBadgeLabel = bootstrap.pageType === 'landing'
    ? 'Landing'
    : (bootstrap.pageType === 'error' ? 'Configuration' : (mode === 'tech' ? 'Technician Mode' : 'View Mode'));

  const firstRow = (bootstrap.rows && bootstrap.rows.length > 0) ? bootstrap.rows[0] : null;
  const storageIdDisplay = bootstrap.storageId || (firstRow && firstRow.storageId ? firstRow.storageId : '');
  const storageLabelDisplay = bootstrap.storageLabel || (firstRow && firstRow.storageLabel ? firstRow.storageLabel : '');
  const locationCodeDisplay = bootstrap.locationCode || (firstRow && firstRow.locationCode ? firstRow.locationCode : '');
  const specificLocationDisplay = bootstrap.specificLocation || (firstRow && firstRow.specificLocation ? firstRow.specificLocation : '');
  const diagnostics = bootstrap.diagnostics || {};

  const storageHeaderBadges = [];
  if (storageIdDisplay) {
    storageHeaderBadges.push('<span class="inline-flex rounded-full bg-teal-500/30 px-2.5 py-1 text-[11px] font-semibold text-teal-100">Storage ID: ' + escapeHtml_(storageIdDisplay) + '</span>');
  }
  if (storageLabelDisplay) {
    storageHeaderBadges.push('<span class="inline-flex rounded-full bg-white/15 px-2.5 py-1 text-[11px] font-medium text-slate-200">' + escapeHtml_(storageLabelDisplay) + '</span>');
  }
  if (locationCodeDisplay) {
    storageHeaderBadges.push('<span class="inline-flex rounded-full bg-sky-500/25 px-2.5 py-1 text-[11px] font-medium text-sky-100">Code: ' + escapeHtml_(locationCodeDisplay) + '</span>');
  }
  if (specificLocationDisplay && specificLocationDisplay !== displayLoc) {
    storageHeaderBadges.push('<span class="inline-flex rounded-full bg-white/10 px-2.5 py-1 text-[11px] font-medium text-slate-200">Location: ' + escapeHtml_(specificLocationDisplay) + '</span>');
  }
  const storageHeaderHtml = storageHeaderBadges.length
    ? '<p class="mt-1.5 flex flex-wrap items-center gap-1.5">' + storageHeaderBadges.join('') + '</p>'
    : '';

  return '<!doctype html>' +
    '<html><head>' +
    '<meta charset="utf-8" />' +
    '<meta name="viewport" content="width=device-width,initial-scale=1" />' +
    '<title>' + escapeHtml_(CONFIG.APP_TITLE) + '</title>' +
    '<script src="https://cdn.tailwindcss.com"></script>' +
    '<script>tailwind.config={theme:{extend:{fontFamily:{sans:[\'Instrument Sans\',\'Segoe UI\',\'sans-serif\']},boxShadow:{panel:\'0 18px 45px rgba(15,23,42,0.10)\'}}}};</script>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
    '<link href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">' +
    '</head><body class="min-h-screen bg-[radial-gradient(circle_at_top,_rgba(14,165,233,0.14),_transparent_28%),linear-gradient(180deg,_#f8fafc_0%,_#f5f5f4_100%)] text-stone-900">' +
    '<main class="mx-auto max-w-5xl px-4 py-5 sm:px-6 sm:py-8">' +
    '<section class="overflow-hidden rounded-[28px] border border-white/70 bg-white/90 shadow-panel backdrop-blur">' +
    '<header class="border-b border-stone-200 bg-[linear-gradient(135deg,_#0f172a_0%,_#1e293b_55%,_#0f766e_100%)] px-5 py-5 text-white sm:px-6 sm:py-6">' +
    '<div class="flex flex-col gap-5 sm:flex-row sm:items-end sm:justify-between">' +
    '<div>' +
    '<p class="text-[11px] font-semibold uppercase tracking-[0.22em] text-sky-200">Internal Inventory</p>' +
    '<h1 class="mt-2 text-2xl font-semibold sm:text-3xl">' + escapeHtml_(CONFIG.APP_TITLE) + '</h1>' +
    '<p class="mt-2 max-w-2xl text-sm text-slate-200">Room: <span id="roomLabel" class="font-semibold text-white">' + (escapeHtml_(room) || '-') + '</span><span class="mx-2 text-slate-400">/</span>Location: <span id="locLabel" class="font-semibold text-white">' + (escapeHtml_(displayLoc) || '-') + '</span></p>' +
    storageHeaderHtml +
    '</div>' +
    '<div class="flex flex-wrap gap-2">' +
    '<a id="allLocationsLink" class="rounded-xl bg-white/15 px-3.5 py-2 text-xs font-medium text-white transition hover:bg-white/25" href="#">All Locations</a>' +
    '<a id="viewModeLink" class="rounded-xl bg-white/15 px-3.5 py-2 text-xs font-medium text-white transition hover:bg-white/25" href="#">View Mode</a>' +
    '<a id="techModeLink" class="rounded-xl bg-sky-400 px-3.5 py-2 text-xs font-medium text-slate-950 transition hover:bg-sky-300" href="#">Technician Mode</a>' +
    '</div></div></header>' +
    '<section class="px-4 pt-4 sm:px-6">' +
    '<section id="notice" class="hidden rounded-2xl px-4 py-3 text-sm"></section>' +
    '<section id="bridgeWarning" class="hidden mt-3 rounded-2xl bg-amber-100 px-4 py-3 text-sm text-amber-900"></section>' +
    '<section id="warningList" class="' + ((bootstrap.warnings && bootstrap.warnings.length) ? '' : 'hidden') + ' mt-3 space-y-2">' +
    (bootstrap.warnings || []).map(function (warning) {
      return '<div class="rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">' + escapeHtml_(warning) + '</div>';
    }).join('') +
    '</section>' +
    '<section id="errorBanner" class="' + (bootstrap.error ? '' : 'hidden') + ' mt-3 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800">' + escapeHtml_(bootstrap.error || '') + '</section>' +
    '<section id="messageBanner" class="' + (bootstrap.message ? '' : 'hidden') + ' mt-3 rounded-2xl border border-sky-200 bg-sky-50 px-4 py-3 text-sm text-sky-900">' + escapeHtml_(bootstrap.message || '') + '</section>' +
    (CONFIG.DEBUG_PANEL ? '<section id="debugPanel" class="mt-3 rounded-2xl bg-slate-900 px-4 py-3 text-xs text-slate-100"></section>' : '') +
    '</section>' +
    '<section class="px-4 pb-4 pt-4 sm:px-6 sm:pb-6">' +
    '<div class="overflow-hidden rounded-[24px] border border-stone-200 bg-white">' +
    '<div class="flex items-center justify-between border-b border-stone-200 bg-stone-50 px-4 py-3">' +
    '<div><h2 class="text-sm font-semibold text-stone-900">' + (isLocationPage ? 'Inventory Items' : 'Storage Locations') + '</h2><p class="mt-0.5 text-xs text-stone-500">' + (isLocationPage ? 'Review expected stock and update technician records.' : 'Select a location to view or update inventory.') + '</p></div>' +
    '<span id="modeBadge" class="rounded-full px-2.5 py-1 text-xs font-medium">' + escapeHtml_(modeBadgeLabel) + '</span>' +
    '</div>' +
    '<div id="loading" class="hidden p-4 text-sm text-stone-500">Saving updates...</div>' +
    '<div id="items">' + initialHtml + '</div>' +
    '</div>' +
    '<footer class="mt-4 flex justify-end">' +
    '<button id="saveBtn" class="' + ((mode === 'tech' && hasEditableRows) ? '' : 'hidden ') + 'rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-50" type="button">Save stock updates</button>' +
    '</footer>' +
    '</section>' +
    '</section>' +
    '</main>' +
    '<script>' +
    'const APP={room:' + JSON.stringify(room) + ',loc:' + JSON.stringify(routeLoc) + ',displayLoc:' + JSON.stringify(displayLoc) + ',mode:' + JSON.stringify(mode) + ',bootstrap:' + JSON.stringify(bootstrap) + ',diagnostics:' + JSON.stringify(diagnostics) + '};' +
    'document.addEventListener("DOMContentLoaded",init);' +
    'function init(){buildModeLinks();renderModeBadge();setBridgeWarning();bindLandingSearch();bindItemFilters();if(' + (CONFIG.DEBUG_PANEL ? 'true' : 'false') + '){renderDebug();}var saveBtn=document.getElementById("saveBtn");if(saveBtn){saveBtn.addEventListener("click",saveUpdates);}}' +
    'function buildModeLinks(){var all=document.getElementById("allLocationsLink");var view=document.getElementById("viewModeLink");var tech=document.getElementById("techModeLink");if(all){all.href=window.location.pathname;}if(!APP.room||!APP.loc){view.classList.add("opacity-50","pointer-events-none");tech.classList.add("opacity-50","pointer-events-none");view.title="Choose a location below first";tech.title="Choose a location below first";view.href="#";tech.href="#";return;}var base=window.location.pathname+"?room="+encodeURIComponent(APP.room)+"&loc="+encodeURIComponent(APP.loc);view.href=base;tech.href=base+"&mode=tech";}' +
    'function renderModeBadge(){var badge=document.getElementById("modeBadge");var saveBtn=document.getElementById("saveBtn");var hasRows=!!(APP.bootstrap&&APP.bootstrap.rows&&APP.bootstrap.rows.length);if(APP.bootstrap.pageType==="landing"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Landing";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.bootstrap.pageType==="error"){badge.className="rounded-full bg-red-100 px-2.5 py-1 text-xs font-medium text-red-800";badge.textContent="Configuration";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.mode==="tech"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Technician Mode";if(saveBtn)saveBtn.classList.toggle("hidden",!hasRows);}else{badge.className="rounded-full bg-stone-200 px-2.5 py-1 text-xs font-medium text-stone-700";badge.textContent="View Mode";if(saveBtn)saveBtn.classList.add("hidden");}}' +
    'function hasBridge(){return !!(window.google&&google.script&&google.script.run);}' +
    'function setBridgeWarning(){if(APP.mode!=="tech")return;if(hasBridge())return;var warn=document.getElementById("bridgeWarning");warn.textContent="Interactive save is unavailable in this context. Open the deployed /exec web app URL to use full functionality.";warn.classList.remove("hidden");}' +
    'function bindLandingSearch(){var input=document.getElementById("locationSearch");if(!input)return;input.addEventListener("input",filterLocations);}' +
    'function filterLocations(){var input=document.getElementById("locationSearch");var query=(input.value||"").trim().toLowerCase();var cards=Array.prototype.slice.call(document.querySelectorAll(".location-card"));var visibleCount=0;cards.forEach(function(card){var haystack=card.getAttribute("data-search")||"";var matches=!query||haystack.indexOf(query)!==-1;card.classList.toggle("hidden",!matches);if(matches)visibleCount+=1;});var headers=Array.prototype.slice.call(document.querySelectorAll(".room-group-header"));headers.forEach(function(header){var roomName=header.getAttribute("data-room-group");var roomCards=cards.filter(function(card){return card.getAttribute("data-room")===roomName&&!card.classList.contains("hidden");});header.classList.toggle("hidden",roomCards.length===0);});var empty=document.getElementById("locationNoResults");if(empty)empty.classList.toggle("hidden",visibleCount!==0);}' +
    'function bindItemFilters(){var search=document.getElementById("itemSearch");var status=document.getElementById("statusFilter");if(search)search.oninput=filterItems;if(status)status.onchange=filterItems;}' +
    'function filterItems(){var search=document.getElementById("itemSearch");var status=document.getElementById("statusFilter");var query=(search&&search.value||"").trim().toLowerCase();var selected=(status&&status.value||"").trim().toLowerCase();var cards=Array.prototype.slice.call(document.querySelectorAll(".inventory-card"));var visible=0;cards.forEach(function(card){var haystack=card.getAttribute("data-item-search")||"";var cardStatus=card.getAttribute("data-status")||"";var matchesText=!query||haystack.indexOf(query)!==-1;var matchesStatus=!selected||cardStatus===selected;var show=matchesText&&matchesStatus;card.classList.toggle("hidden",!show);if(show)visible+=1;});var empty=document.getElementById("itemNoResults");if(empty)empty.classList.toggle("hidden",visible!==0);}' +
    'function renderDebug(){var el=document.getElementById("debugPanel");if(!el)return;var rowCount=APP.bootstrap&&APP.bootstrap.rows?APP.bootstrap.rows.length:0;var locCount=APP.bootstrap&&APP.bootstrap.locations?APP.bootstrap.locations.length:0;var urlConfigured=APP.bootstrap&&APP.bootstrap.webAppBaseUrl?"yes":"no";var sheetName=(APP.diagnostics&&APP.diagnostics.sheetInUse)||"-";el.innerHTML="mode="+esc(APP.mode)+" | room="+esc(APP.room||"-")+" | loc="+esc(APP.loc||"-")+" | bridge="+(hasBridge()?"yes":"no")+" | urlConfigured="+urlConfigured+" | sheet="+esc(sheetName)+" | rows="+rowCount+" | locations="+locCount;}' +
    'function saveUpdates(){if(!hasBridge()){showNotice("Interactive save is unavailable in this context. Open the deployed /exec web app URL.",true);return;}var saveBtn=document.getElementById("saveBtn");var loading=document.getElementById("loading");var originalText=saveBtn.textContent;saveBtn.disabled=true;saveBtn.textContent="Saving...";loading.classList.remove("hidden");var byRow={};var invalidQty=false;document.querySelectorAll(".qty-input").forEach(function(input){var row=input.getAttribute("data-row");var qty=Number(input.value);if(!Number.isFinite(qty)||qty<0)invalidQty=true;byRow[row]=byRow[row]||{sheetRow:Number(row)};byRow[row].qty=qty;});if(invalidQty){saveBtn.disabled=false;saveBtn.textContent=originalText;loading.classList.add("hidden");showNotice("Enter a non-negative number for every quantity before saving.",true);return;}document.querySelectorAll(".status-input").forEach(function(select){var row=select.getAttribute("data-row");byRow[row]=byRow[row]||{sheetRow:Number(row)};byRow[row].status=select.value;});var rowKeys=Object.keys(byRow);if(!rowKeys.length){saveBtn.disabled=false;saveBtn.textContent=originalText;loading.classList.add("hidden");showNotice("There are no inventory items to save.",true);return;}google.script.run.withSuccessHandler(function(res){saveBtn.disabled=false;saveBtn.textContent=originalText;loading.classList.add("hidden");APP.bootstrap.rows=res.rows||[];document.getElementById("items").innerHTML=res.html||"";bindItemFilters();renderModeBadge();showNotice("Save successful. Updated "+res.updatedCount+" item(s).",false);}).withFailureHandler(function(err){saveBtn.disabled=false;saveBtn.textContent=originalText;loading.classList.add("hidden");showNotice((err&&err.message)||"Save failed. Check the values and try again.",true);}).saveInventoryUpdates({room:APP.room,loc:APP.loc,updates:rowKeys.map(function(key){return byRow[key];})});}' +
    'function showNotice(message,isError){if(!message)return;var el=document.getElementById("notice");el.textContent=message;el.className="rounded-2xl px-4 py-3 text-sm mt-0";if(isError){el.classList.add("bg-red-100","text-red-800","border","border-red-200");}else{el.classList.add("bg-emerald-100","text-emerald-800","border","border-emerald-200");}el.scrollIntoView({behavior:"smooth",block:"nearest"});}' +
    'function esc(value){return String(value==null?"":value).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/\'/g,"&#39;");}' +
    '</script></body></html>';
}
