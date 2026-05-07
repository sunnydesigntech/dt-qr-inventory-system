/**
 * D&T QR Inventory System
 * Standalone Google Apps Script web app using HtmlService templates.
 *
 * Runtime files:
 * - code.gs
 * - index.html
 * - app_styles.html
 * - app_script.html
 *
 * Required Script Properties:
 * - SPREADSHEET_ID
 * - WEB_APP_BASE_URL (recommended; app can still read inventory without it)
 * - INVENTORY_SHEET_NAME (optional)
 * - EXTERNAL_SCANNER_URL (optional; top-level HTTPS scanner page)
 */

const CONFIG = Object.freeze({
  APP_TITLE: 'D&T QR Inventory System',
  HEADER_ROW: 1,
  DEFAULT_SHEET_NAME: 'Inventory',
  DEFAULT_SPREADSHEET_ID: '',
  DEFAULT_WEB_APP_BASE_URL: '',
  DEFAULT_EXTERNAL_SCANNER_URL: 'https://sunnydesigntech.github.io/dt-qr-inventory-system/scanner/',
  SPREADSHEET_ID_PROPERTY: 'SPREADSHEET_ID',
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  EXTERNAL_SCANNER_URL_PROPERTY: 'EXTERNAL_SCANNER_URL',
  INVENTORY_SHEET_NAME_PROPERTY: 'INVENTORY_SHEET_NAME',
  UPDATE_AUTH_ALLOWED_EMAILS_PROPERTY: 'UPDATE_AUTH_ALLOWED_EMAILS',
  UPDATE_AUTH_ALLOWED_DOMAINS_PROPERTY: 'UPDATE_AUTH_ALLOWED_DOMAINS',
  UPDATE_MODE_PIN_SHA256_PROPERTY: 'UPDATE_MODE_PIN_SHA256',
  UPDATE_MODE_PIN_SALT_PROPERTY: 'UPDATE_MODE_PIN_SALT',
  UPDATE_AUTH_DISABLED_PROPERTY: 'UPDATE_AUTH_DISABLED',
  UPDATE_AUTH_TOKEN_PREFIX: 'dt-inventory-update-auth:',
  UPDATE_AUTH_TOKEN_TTL_SECONDS: 3600,
  ROLLOUT_ROOM: '419A',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemicals', 'chemical'],
  QUICKCHART_QR_BASE: 'https://quickchart.io/qr?size=220&text=',
  IMPORT_READY_SHEET_NAMES: ['419A_App_Load_Ready', '419A App Load Ready', '419AAppLoadReady'],
  STORAGE_MASTER_SHEET_NAMES: ['419A_Storage_Master', '419A Storage Master', '419AStorageMaster', 'Room_QR_Label_Plan', 'RM 419A 2026'],
  STORAGE_MASTER_SHEET_NAME: 'Storage_Master',
  AUDIT_LOG_SHEET_NAME: 'Audit_Log',
  QR_LABEL_SHEET_NAME: 'QR_Labels',
  READINESS_REPORT_SHEET_NAME: 'Inventory_Readiness_Report',
  WARNING_TRIAGE_SHEET_NAME: 'Readiness_Warning_Triage',
  UNMATCHED_REVIEW_SHEET_NAME: '419A_Unmatched_Review',
  PILOT_TEST_LOG_SHEET_NAME: 'PILOT_TEST_LOG',
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
    storageLabel: ['storage label', 'storagelabel', 'storage_label'],
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
  },
  REQUIRED_COLUMNS: ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'],
  OPTIONAL_COLUMNS: [
    'qrLink',
    'qrImage',
    'unit',
    'remarks',
    'locationCode',
    'storageId',
    'storageLabel',
    'storageType',
    'lastUpdated',
    'updatedBy',
    'isPlaceholder',
    'safetyNote',
    'reorderLevel',
    'supplier',
    'purchaseLink',
    'assetValue',
    'maintenanceDue',
    'sdsLink'
  ],
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
    qrImage: 'QR Code Image',
    storageType: 'Storage Type',
    lastUpdated: 'Last Updated',
    updatedBy: 'Updated By',
    isPlaceholder: 'Is Placeholder',
    safetyNote: 'Safety Note',
    reorderLevel: 'Reorder Level',
    supplier: 'Supplier',
    purchaseLink: 'Purchase Link',
    assetValue: 'Asset Value',
    maintenanceDue: 'Maintenance Due',
    sdsLink: 'SDS Link'
  }
});

function doGet(e) {
  const params = getRequestParams_(e);
  const payload = buildClientPayload_(params);

  const template = HtmlService.createTemplateFromFile('index');
  template.bootstrapJson = JSON.stringify(payload);

  return template.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('D&T Inventory')
    .addItem('Refresh QR Links', 'menuRefreshQrLinks')
    .addItem('Refresh QR Images (Optional)', 'menuRefreshQrImages')
    .addSeparator()
    .addItem('Prepare App Columns', 'menuPrepareAppColumns')
    .addItem('Build Storage Master', 'menuBuildStorageMasterSheet')
    .addItem('Build QR Label Sheet', 'menuBuildQrLabelSheet')
    .addItem('Create Readiness Report', 'menuCreateReadinessReport')
    .addItem('Create Warning Triage Board', 'menuCreateWarningTriageBoard')
    .addSeparator()
    .addItem('419A Readiness Summary', 'show419AReadiness')
    .addItem('Prepare 419A Unmatched Review', 'menuPrepareUnmatchedReview')
    .addItem('Create Pilot Test Log', 'menuCreatePilotTestLog')
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

function menuBuildStorageMasterSheet() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = buildStorageMasterSheet();
    ui.alert(
      'Storage Master Ready',
      'Created/updated "' + result.sheetName + '" with ' + result.storageCount + ' storage row(s).',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Storage Master Failed', err.message || String(err), ui.ButtonSet.OK);
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

function menuCreateWarningTriageBoard() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = createWarningTriageBoard();
    ui.alert(
      'Warning Triage Board Ready',
      'Created/updated "' + result.sheetName + '" with ' + result.warningCount + ' warning row(s).\n' +
        'Warnings are now grouped, assigned, and ready for pilot decisions.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Warning Triage Board Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuPrepareUnmatchedReview() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = prepare419AUnmatchedReviewSheet();
    ui.alert(
      '419A Unmatched Review Ready',
      'Prepared "' + result.sheetName + '" with ' + result.reviewRows + ' review row(s).\n' +
        'No rows were imported into Inventory.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('419A Unmatched Review Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function menuCreatePilotTestLog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = createPilotTestLog();
    ui.alert(
      'Pilot Test Log Ready',
      'Created/updated "' + result.sheetName + '" with ' + result.templateRows + ' starter test row(s).',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Pilot Test Log Failed', err.message || String(err), ui.ButtonSet.OK);
  }
}

function show419AReadiness() {
  const ui = SpreadsheetApp.getUi();
  try {
    const summary = get419AReadinessSummary(CONFIG.ROLLOUT_ROOM);
    const lines = [
      'Room: ' + summary.room,
      'Storage locations: ' + summary.locationCount,
      'Locations with 419A code: ' + summary.locationsWithStorageId,
      'Locations missing 419A code: ' + summary.locationsMissingStorageId,
      'Inventory item rows: ' + summary.itemRows,
      'Empty storage placeholder rows: ' + summary.emptyStorageRows,
      'Chemical item rows: ' + summary.chemicalRows,
      'Legacy-only locations: ' + summary.legacyOnlyLocations,
      'Status counts: ' + JSON.stringify(summary.statusCounts)
    ];
    if (summary.sampleMissingStorageIds.length) {
      lines.push('Sample missing 419A code: ' + summary.sampleMissingStorageIds.join(', '));
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
    'SPREADSHEET_ID: ' + (status.scriptProperties.spreadsheetIdConfigured ? (status.scriptProperties.spreadsheetIdFromFallback ? 'fallback ' : 'configured ') + (status.maskedSpreadsheetId || '') : 'missing'),
    'WEB_APP_BASE_URL: ' + (status.scriptProperties.webAppBaseUrlConfigured ? (status.scriptProperties.webAppBaseUrlFromRuntime ? 'runtime fallback' : (status.scriptProperties.webAppBaseUrlFromFallback ? 'fallback/default' : 'configured')) : 'missing'),
    'INVENTORY_SHEET_NAME: ' + (status.scriptProperties.inventorySheetNameConfigured ? status.scriptProperties.inventorySheetName : 'default/fallback'),
    'Sheet in use: ' + (status.sheetInUse || '-'),
    'Data rows: ' + (typeof status.dataRows === 'number' ? status.dataRows : '-'),
    'Missing required columns: ' + (missingRequired.length ? missingRequired.join(', ') : 'none'),
    'Optional columns found: ' + (optionalFound.length ? optionalFound.join(', ') : 'none'),
    'Available sheets: ' + ((status.availableSheetNames || []).join(', ') || '-'),
    'Header preview: ' + ((status.headerPreview || []).slice(0, 12).join(', ') || '-')
  ];

  if (status.rollout419A) {
    lines.push('419A locations/items: ' + status.rollout419A.locationCount + ' location(s), ' + status.rollout419A.itemRows + ' item row(s)');
  }
  if (status.sheetError) lines.push('Sheet/config error: ' + status.sheetError);
  if (status.nextAction) lines.push('Next action: ' + status.nextAction);
  ui.alert('D&T Inventory Diagnostics', lines.join('\n'), ui.ButtonSet.OK);
}

function getRequestParams_(e) {
  const admin = cleanString_(e && e.parameter && e.parameter.admin).toLowerCase();
  return {
    room: cleanString_(e && e.parameter && e.parameter.room),
    loc: cleanString_(e && e.parameter && e.parameter.loc),
    mode: normalizeMode_(e && e.parameter && e.parameter.mode),
    admin: ['diagnostics', 'readiness', 'labels', 'storage', 'audit', 'lowstock', 'maintenance'].indexOf(admin) !== -1 ? admin : ''
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
    updateAuth: getUpdateAuthClientStateSafe_(),
    webAppBaseUrl: '',
    diagnostics: getDiagnosticsSafe_()
  };
}

function buildClientPayload_(params) {
  const safeParams = params || {};
  try {
    const webAppBaseUrl = getWebAppBaseUrl_({ silent: true });
    const needsDiagnostics = safeParams.admin === 'diagnostics' || safeParams.admin === 'readiness';
    const dataset = getInventoryDataset_();
    const validation = safeParams.admin === 'readiness' ? validateInventoryData_() : null;
    const diagnostics = needsDiagnostics ? getDiagnostics_() : getLightDiagnostics_();
    const appData = buildClientAppData_(dataset, webAppBaseUrl, {
      validation: validation,
      audit: readRecentAuditEvents_(25)
    });
    const route = buildClientRoute_(safeParams, appData.locations);
    return {
      appTitle: CONFIG.APP_TITLE,
      webAppBaseUrl: webAppBaseUrl,
      externalScannerUrl: getExternalScannerUrl_({ silent: true }),
      route: route,
      error: '',
      warnings: webAppBaseUrl ? [] : ['WEB_APP_BASE_URL is not configured. In-app browsing still works; QR generation and external links are disabled until it is set.'],
      updateAuth: getUpdateAuthClientState_(),
      diagnostics: buildClientDiagnostics_(diagnostics),
      appData: appData
    };
  } catch (err) {
    return buildClientErrorPayload_(safeParams, err);
  }
}

function buildClientErrorPayload_(params, err) {
  const message = err && err.message ? err.message : 'Unexpected application error.';
  return {
    appTitle: CONFIG.APP_TITLE,
    webAppBaseUrl: getWebAppBaseUrl_({ silent: true }),
    externalScannerUrl: getExternalScannerUrl_({ silent: true }),
    route: { name: 'error', kind: 'generic', message: message },
    error: message,
    warnings: [],
    updateAuth: getUpdateAuthClientStateSafe_(),
    diagnostics: buildClientDiagnostics_(getDiagnosticsSafe_()),
    appData: buildEmptyClientAppData_()
  };
}

function buildEmptyClientAppData_() {
  return {
    metrics: {
      totalRooms: 0,
      totalLocations: 0,
      totalItems: 0,
      totalChemicals: 0,
      totalAttention: 0
    },
    rooms: [],
    locations: [],
    items: {},
    readiness: {
      score: 0,
      errors: 1,
      warnings: 0,
      missingQr: 0,
      missingStorageId: 0,
      invalidQty: 0,
      invalidStatus: 0,
      duplicates: 0,
      rollout419A: {
        ready: 0,
        total: 0,
        pct: 0,
        itemCount: 0,
        chemicalCount: 0
      }
    },
    audit: []
  };
}

function buildClientAppData_(dataset, webAppBaseUrl, options) {
  const opts = options || {};
  const map = dataset.map;
  const values = dataset.values;
  const locationsByKey = {};
  const locations = [];
  const itemsByLocation = {};
  const roomsByCode = {};
  const roomOrder = [];
  const metrics = {
    totalRooms: 0,
    totalLocations: 0,
    totalItems: 0,
    totalChemicals: 0,
    totalAttention: 0
  };

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) continue;

    const entry = buildLocationEntry_(row, map, room, loc);
    let location = locationsByKey[entry.canonicalKey];
    if (!location) {
      location = buildClientLocation_(entry, row, map, webAppBaseUrl);
      locationsByKey[entry.canonicalKey] = location;
      locations.push(location);
      itemsByLocation[entry.canonicalKey] = [];
      ensureClientRoom_(roomsByCode, roomOrder, room);
      roomsByCode[room].locationCount += 1;
    }

    if (map.qrLink !== -1 && getOptionalValue_(row, map.qrLink)) {
      location.qrReady = isValidQrLink_(getOptionalValue_(row, map.qrLink), room);
    }

    if (!isInventoryItemRow_(row, map)) continue;

    const item = buildClientItem_(row, map, i + 1, room, loc);
    itemsByLocation[entry.canonicalKey].push(item);
    location.items += 1;
    location.itemCount = location.items;
    if (item.status !== 'Good') {
      location.attention += 1;
      location.attentionCount = location.attention;
    }
    if (item.hazard) {
      location.hazard = true;
      location.chemicalCount += 1;
    }

    ensureClientRoom_(roomsByCode, roomOrder, room);
    roomsByCode[room].itemCount += 1;
    if (item.hazard) roomsByCode[room].chemicalCount += 1;
    if (item.status !== 'Good') {
      roomsByCode[room].attention += 1;
      roomsByCode[room].attentionCount = roomsByCode[room].attention;
    }
  }

  locations.sort(function (a, b) {
    return a.sortKey.localeCompare(b.sortKey);
  });

  Object.keys(itemsByLocation).forEach(function (key) {
    itemsByLocation[key].sort(function (a, b) {
      return (a.name || '').localeCompare(b.name || '');
    });
  });

  roomOrder.sort(function (a, b) {
    if (a === CONFIG.ROLLOUT_ROOM) return -1;
    if (b === CONFIG.ROLLOUT_ROOM) return 1;
    return a.localeCompare(b);
  });

  metrics.totalRooms = roomOrder.length;
  metrics.totalLocations = locations.length;
  metrics.totalItems = locations.reduce(function (sum, location) { return sum + (location.itemCount || location.items || 0); }, 0);
  metrics.totalChemicals = locations.reduce(function (sum, location) { return sum + (location.chemicalCount || 0); }, 0);
  metrics.totalAttention = locations.reduce(function (sum, location) { return sum + (location.attentionCount || location.attention || 0); }, 0);

  return {
    metrics: metrics,
    rooms: roomOrder.map(function (room) { return roomsByCode[room]; }),
    locations: locations,
    items: itemsByLocation,
    readiness: buildClientReadiness_(locations, itemsByLocation, opts.validation),
    audit: opts.audit || []
  };
}

function ensureClientRoom_(roomsByCode, roomOrder, room) {
  if (roomsByCode[room]) return;
  roomsByCode[room] = {
    code: room,
    name: getClientRoomName_(room),
    rollout: room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase(),
    locationCount: 0,
    itemCount: 0,
    chemicalCount: 0,
    attention: 0,
    attentionCount: 0
  };
  roomOrder.push(room);
}

function getClientRoomName_(room) {
  if (room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase()) return 'D&T Workshop 419A';
  if (room.toLowerCase() === 'v++') return 'V++ Maker Studio';
  return 'Room ' + room;
}

function buildClientLocation_(entry, row, map, webAppBaseUrl) {
  const displayId = entry.routeLoc || entry.storageId || entry.loc;
  const relativeViewUrl = buildLocationHref_('', entry.room, entry.routeLoc, 'view');
  const absoluteViewUrl = webAppBaseUrl ? buildLocationUrl_(webAppBaseUrl, entry.room, entry.routeLoc) : '';
  const viewUrl = relativeViewUrl;
  return {
    key: entry.canonicalKey,
    id: displayId,
    room: entry.room,
    label: entry.displayLoc || entry.loc,
    specific: entry.loc,
    code: entry.locationCode,
    storageId: entry.storageId,
    storageLabel: entry.storageLabel,
    storageType: entry.storageType,
    lastChecked: getOptionalValue_(row, map.lastUpdated),
    notes: getOptionalValue_(row, map.remarks),
    routeLoc: entry.routeLoc,
    searchText: entry.searchText,
    sortKey: entry.sortKey,
    items: 0,
    itemCount: 0,
    attention: 0,
    attentionCount: 0,
    chemicalCount: 0,
    hazard: false,
    qrReady: map.qrLink !== -1 && isValidQrLink_(getOptionalValue_(row, map.qrLink), entry.room),
    viewUrl: viewUrl,
    techUrl: buildLocationHref_('', entry.room, entry.routeLoc, 'tech'),
    absoluteViewUrl: absoluteViewUrl,
    absoluteTechUrl: absoluteViewUrl ? absoluteViewUrl + '&mode=tech' : '',
    qrImageUrl: absoluteViewUrl ? CONFIG.QUICKCHART_QR_BASE + encodeURIComponent(absoluteViewUrl) : ''
  };
}

function buildClientItem_(row, map, sheetRow, room, loc) {
  const category = cleanString_(row[map.category]);
  const status = normalizeStatus_(row[map.status]);
  return {
    sheetRow: sheetRow,
    id: cleanString_(row[map.itemId]),
    name: cleanString_(row[map.itemName]) || '(Unnamed item)',
    room: room,
    specificLocation: loc,
    category: category,
    qty: toNonNegativeNumber_(row[map.qty]),
    unit: getOptionalValue_(row, map.unit),
    status: status,
    remarks: getOptionalValue_(row, map.remarks),
    safetyNote: getOptionalValue_(row, map.safetyNote),
    reorderLevel: getOptionalValue_(row, map.reorderLevel),
    supplier: getOptionalValue_(row, map.supplier),
    purchaseLink: getOptionalValue_(row, map.purchaseLink),
    assetValue: getOptionalValue_(row, map.assetValue),
    maintenanceDue: getOptionalValue_(row, map.maintenanceDue),
    sdsLink: getOptionalValue_(row, map.sdsLink),
    lastUpdated: getOptionalValue_(row, map.lastUpdated),
    updatedBy: getOptionalValue_(row, map.updatedBy),
    storageId: getOptionalValue_(row, map.storageId),
    storageLabel: getOptionalValue_(row, map.storageLabel),
    locationCode: getOptionalValue_(row, map.locationCode),
    hazard: isHazardCategory_(category)
  };
}

function buildClientRoute_(params, locations) {
  if (params && params.admin === 'labels') {
    return { name: 'labels', admin: 'labels' };
  }
  if (params && params.admin === 'storage') {
    return { name: 'storage', admin: 'storage' };
  }
  if (params && params.admin === 'audit') {
    return { name: 'audit', admin: 'audit' };
  }
  if (params && params.admin === 'lowstock') {
    return { name: 'lowstock', admin: 'lowstock' };
  }
  if (params && params.admin === 'maintenance') {
    return { name: 'maintenance', admin: 'maintenance' };
  }
  if (params && (params.admin === 'diagnostics' || params.admin === 'readiness')) {
    return { name: 'admin', admin: params.admin };
  }
  if (!(params && params.room && params.loc)) {
    return { name: 'landing' };
  }

  const location = findClientLocation_(locations, params.room, params.loc);
  if (!location) {
    return {
      name: 'error',
      kind: 'no-match',
      message: 'No matching storage location was found for this room/location.'
    };
  }

  return {
    name: 'location',
    locKey: location.key,
    mode: params.mode === 'tech' ? 'tech' : 'view'
  };
}

function findClientLocation_(locations, room, loc) {
  const roomNeedle = cleanString_(room).toLowerCase();
  const locNeedle = cleanString_(loc).toLowerCase();
  return locations.filter(function (entry) {
    if (entry.room.toLowerCase() !== roomNeedle) return false;
    return [
      entry.routeLoc,
      entry.id,
      entry.storageId,
      entry.storageLabel,
      entry.code,
      entry.specific,
      entry.label
    ].filter(Boolean).map(function (value) {
      return cleanString_(value).toLowerCase();
    }).indexOf(locNeedle) !== -1;
  })[0] || null;
}

function buildClientReadiness_(locations, itemsByLocation, validation) {
  const hasValidation = !!validation;
  const validationResult = validation || { issues: [], errorCount: 0, warningCount: 0 };

  const issues = validationResult.issues || [];
  const rolloutLocations = locations.filter(function (location) {
    return location.room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase();
  });
  const ready419A = rolloutLocations.filter(function (location) {
    return !!((location.code || location.storageId) && location.qrReady);
  }).length;
  const rolloutItemCount = rolloutLocations.reduce(function (sum, location) {
    return sum + location.items;
  }, 0);
  const rolloutChemicalCount = rolloutLocations.reduce(function (sum, location) {
    return sum + (itemsByLocation[location.key] || []).filter(function (item) { return item.hazard; }).length;
  }, 0);
  const warningCount = hasValidation ? (validationResult.warningCount || issues.filter(function (issue) { return issue.severity !== 'ERROR'; }).length) : 0;
  const errorCount = hasValidation ? (validationResult.errorCount || issues.filter(function (issue) { return issue.severity === 'ERROR'; }).length) : 0;

  return {
    score: Math.max(0, Math.min(100, 100 - (errorCount * 18) - (warningCount * 3))),
    errors: errorCount,
    warnings: warningCount,
    missingQr: locations.filter(function (location) { return !location.qrReady; }).length,
    missingStorageId: rolloutLocations.filter(function (location) { return !(location.code || location.storageId); }).length,
    invalidQty: countClientIssues_(issues, 'INVALID_QTY'),
    invalidStatus: countClientIssues_(issues, 'INVALID_STATUS'),
    duplicates: countClientIssues_(issues, 'POSSIBLE_DUPLICATE_ITEM'),
    validationLoaded: hasValidation,
    issueSample: issues.slice(0, 20),
    rollout419A: {
      ready: ready419A,
      total: rolloutLocations.length,
      pct: rolloutLocations.length ? Math.round((ready419A / rolloutLocations.length) * 100) : 0,
      itemCount: rolloutItemCount,
      chemicalCount: rolloutChemicalCount
    }
  };
}

function countClientIssues_(issues, issueName) {
  return (issues || []).filter(function (issue) {
    return issue.issue === issueName;
  }).length;
}

function buildClientDiagnostics_(status) {
  const safeStatus = status || {};
  const props = safeStatus.scriptProperties || {};
  const required = safeStatus.requiredColumns || {};
  const optional = safeStatus.optionalColumns || {};
  const missingRequired = safeStatus.missingRequiredColumns || CONFIG.REQUIRED_COLUMNS.filter(function (key) {
    return String(required[key] || '').indexOf('found') !== 0;
  }).map(function (key) { return CONFIG.COLUMN_LABELS[key]; });

  return {
    spreadsheetIdSet: !!props.spreadsheetIdConfigured,
    spreadsheetIdSource: props.spreadsheetIdFromFallback ? 'fallback' : (props.spreadsheetIdPropertySet ? 'Script Property' : 'missing'),
    spreadsheetMaskedId: safeStatus.maskedSpreadsheetId || '',
    webAppBaseUrlSet: !!props.webAppBaseUrlConfigured,
    webAppBaseUrlSource: props.webAppBaseUrlFromRuntime ? 'runtime' : (props.webAppBaseUrlFromFallback ? 'fallback' : (props.webAppBaseUrlPropertySet ? 'Script Property' : 'missing')),
    inventorySheetNameSet: !!props.inventorySheetNameConfigured,
    sheetFound: !!safeStatus.sheetInUse && !safeStatus.sheetError,
    sheetInUse: safeStatus.sheetInUse ? safeStatus.sheetInUse + ' · ' + (safeStatus.dataRows || 0) + ' rows' : '',
    sheetName: safeStatus.sheetInUse || '',
    dataRows: safeStatus.dataRows || 0,
    requiredColumnsOk: missingRequired.length === 0 && !safeStatus.sheetError,
    missingRequiredColumns: missingRequired,
    qrImageColumn: String(optional.qrImage || '').indexOf('found') === 0,
    requiredColumns: required,
    optionalColumns: optional,
    headerPreview: safeStatus.headerPreview || [],
    availableSheets: safeStatus.availableSheets || [],
    sheetError: safeStatus.sheetError || '',
    nextAction: safeStatus.nextAction || '',
    raw: safeStatus
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
  const auth = requireUpdateAuthorization_(payload);

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
  const actor = auth.user || getActiveUserEmail_();
  const timestamp = new Date();
  const auditEvents = [];

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
    assertExpectedInventoryRowIdentity_(row, map, update, rowNum);

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

    const oldQty = row[map.qty];
    const oldStatus = normalizeStatus_(row[map.status]);
    const oldQtyText = cleanString_(oldQty);
    const newQtyText = cleanString_(qty);
    const changed = oldQtyText !== newQtyText || oldStatus !== status;

    sheet.getRange(rowNum, map.qty + 1).setValue(qty);
    sheet.getRange(rowNum, map.status + 1).setValue(status);
    if (map.lastUpdated !== -1) sheet.getRange(rowNum, map.lastUpdated + 1).setValue(timestamp);
    if (map.updatedBy !== -1) sheet.getRange(rowNum, map.updatedBy + 1).setValue(actor);

    if (changed) {
      auditEvents.push({
        timestamp: timestamp,
        user: actor,
        action: 'Update stock',
        room: cleanString_(row[map.room]),
        storageId: getOptionalValue_(row, map.storageId),
        location: cleanString_(row[map.location]),
        routeLoc: loc,
        itemId: cleanString_(row[map.itemId]),
        itemName: cleanString_(row[map.itemName]),
        oldQty: oldQty,
        newQty: qty,
        oldStatus: oldStatus,
        newStatus: status,
        notes: 'Update Mode save; route=' + loc
      });
    }
  });

  appendAuditEvents_(auditEvents);

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

function addInventoryItemToLocation(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid add-item payload.');
  }
  const auth = requireUpdateAuthorization_(payload);

  const room = cleanString_(payload.room);
  const loc = cleanString_(payload.loc);
  const item = payload.item || {};
  if (!room || !loc) {
    throw new Error('Room and location are required to add an item.');
  }

  const itemName = cleanString_(item.itemName);
  if (!itemName) {
    throw new Error('Item name is required.');
  }

  const qty = Number(item.qty);
  if (!Number.isFinite(qty) || qty < 0) {
    throw new Error('Quantity must be a non-negative number.');
  }

  const status = matchStatus_(item.status);
  if (!status) {
    throw new Error('Invalid status: ' + cleanString_(item.status));
  }

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('The inventory sheet is empty.');

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const context = findLocationContextRow_(values, map, room, loc);
  if (!context) {
    throw new Error('Storage location not found. Add items only from a valid storage page.');
  }

  const actor = auth.user || getActiveUserEmail_();
  const timestamp = new Date();
  const newRow = buildNewInventoryRow_(values[0].length, map, context.row, {
    itemId: cleanString_(item.itemId) || generateItemId_(room, loc, itemName),
    itemName: itemName,
    qty: qty,
    unit: cleanString_(item.unit),
    category: cleanString_(item.category) || 'Tools',
    status: status,
    remarks: cleanString_(item.remarks),
    actor: actor,
    timestamp: timestamp
  });

  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, newRow.length).setValues([newRow]);
  setQrImageFormulaForRow_(sheet, map, rowNumber);

  appendAuditEvents_([{
    timestamp: timestamp,
    user: actor,
    action: 'Add item',
    room: room,
    storageId: getOptionalValue_(context.row, map.storageId),
    location: cleanString_(context.row[map.location]),
    routeLoc: loc,
    itemId: newRow[map.itemId],
    itemName: itemName,
    oldQty: '',
    newQty: qty,
    oldStatus: '',
    newStatus: status,
    notes: 'Update Mode add item; route=' + loc
  }]);

  return {
    success: true,
    addedRow: rowNumber,
    rows: getInventoryRowsForLocation_(room, loc),
    timestamp: new Date().toISOString()
  };
}

function removeInventoryItemFromLocation(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid remove-item payload.');
  }
  const auth = requireUpdateAuthorization_(payload);

  const room = cleanString_(payload.room);
  const loc = cleanString_(payload.loc);
  const rowNum = Number(payload.sheetRow);
  if (!room || !loc) throw new Error('Room and location are required to remove an item.');
  if (!Number.isInteger(rowNum) || rowNum <= CONFIG.HEADER_ROW) {
    throw new Error('Invalid row number for removal.');
  }

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (rowNum > values.length) throw new Error('The selected row no longer exists. Refresh and try again.');

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const row = values[rowNum - 1];
  if (!rowMatchesRoomLoc_(row, map, room.toLowerCase(), loc.toLowerCase())) {
    throw new Error('The selected item does not belong to this storage.');
  }
  assertExpectedInventoryRowIdentity_(row, map, payload, rowNum);
  if (!isInventoryItemRow_(row, map)) {
    throw new Error('Only real inventory item rows can be removed.');
  }

  const actor = auth.user || getActiveUserEmail_();
  const timestamp = new Date();
  const matching = countLocationRows_(values, map, room, loc);
  if (matching.total <= 1 && !matching.hasPlaceholder) {
    const placeholder = buildPlaceholderRowFromContext_(values[0].length, map, row, actor, timestamp);
    const placeholderRow = sheet.getLastRow() + 1;
    sheet.getRange(placeholderRow, 1, 1, placeholder.length).setValues([placeholder]);
    setQrImageFormulaForRow_(sheet, map, placeholderRow);
  }

  appendAuditEvents_([{
    timestamp: timestamp,
    user: actor,
    action: 'Remove item',
    room: cleanString_(row[map.room]),
    storageId: getOptionalValue_(row, map.storageId),
    location: cleanString_(row[map.location]),
    routeLoc: loc,
    itemId: cleanString_(row[map.itemId]),
    itemName: cleanString_(row[map.itemName]),
    oldQty: row[map.qty],
    newQty: '',
    oldStatus: normalizeStatus_(row[map.status]),
    newStatus: '',
    notes: 'Update Mode remove item; route=' + loc
  }]);

  sheet.deleteRow(rowNum);

  return {
    success: true,
    removedRow: rowNum,
    rows: getInventoryRowsForLocation_(room, loc),
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
  const storageType = getOptionalValue_(row, map.storageType);
  const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);

  return {
    room: room,
    loc: loc,
    displayLoc: storageLabel || loc,
    locationCode: locationCode,
    storageId: storageId,
    storageLabel: storageLabel,
    storageType: storageType,
    routeLoc: identity.routeLoc,
    canonicalKey: identity.key,
    searchText: [room, loc, storageLabel, storageId, locationCode, storageType].filter(Boolean).join(' ').toLowerCase(),
    sortKey: [room, storageId || '', storageLabel || '', locationCode || '', loc].join(' | ').toLowerCase()
  };
}

function isInventoryItemRow_(row, map) {
  if (isPlaceholderRow_(row, map)) return false;
  return !!(cleanString_(row[map.itemId]) || cleanString_(row[map.itemName]));
}

function isPlaceholderRow_(row, map) {
  const explicitPlaceholder = getOptionalValue_(row, map.isPlaceholder).toLowerCase();
  if (['true', 'yes', 'y', '1', 'placeholder'].indexOf(explicitPlaceholder) !== -1) return true;

  const category = getOptionalValue_(row, map.category).toLowerCase();
  const remarks = getOptionalValue_(row, map.remarks).toLowerCase();
  const itemId = getOptionalValue_(row, map.itemId);
  const itemName = getOptionalValue_(row, map.itemName);
  return category === 'storage' &&
    remarks.indexOf('placeholder row for qr/location page') !== -1 &&
    !itemId &&
    !itemName;
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

function assertExpectedInventoryRowIdentity_(row, map, expected, rowNum) {
  const payload = expected || {};
  const expectedItemId = cleanString_(payload.itemId || payload.id);
  const expectedItemName = cleanString_(payload.itemName || payload.name);
  const actualItemId = cleanString_(row[map.itemId]);
  const actualItemName = cleanString_(row[map.itemName]);

  if (!expectedItemId && !expectedItemName) {
    throw new Error('Missing expected item identity for row ' + rowNum + '. Refresh the page before changing inventory.');
  }

  if (expectedItemId || actualItemId) {
    if (!expectedItemId || !actualItemId || normalizeIdentity_(expectedItemId) !== normalizeIdentity_(actualItemId)) {
      throw new Error('Row ' + rowNum + ' no longer matches the item loaded on this page. Refresh before saving.');
    }
  } else if (normalizeIdentity_(expectedItemName) !== normalizeIdentity_(actualItemName)) {
    throw new Error('Row ' + rowNum + ' no longer matches the item loaded on this page. Refresh before saving.');
  }

  const expectedCategory = cleanString_(payload.category || payload.expectedCategory);
  if (expectedCategory && normalizeIdentity_(expectedCategory) !== normalizeIdentity_(row[map.category])) {
    throw new Error('Row ' + rowNum + ' no longer matches the item loaded on this page. Refresh before saving.');
  }

  const expectedQty = cleanString_(payload.expectedQty);
  if (expectedQty && expectedQty !== cleanString_(row[map.qty])) {
    throw new Error('Row ' + rowNum + ' was changed by another session. Refresh before saving.');
  }

  const expectedStatus = cleanString_(payload.expectedStatus);
  if (expectedStatus && normalizeStatus_(expectedStatus) !== normalizeStatus_(row[map.status])) {
    throw new Error('Row ' + rowNum + ' was changed by another session. Refresh before saving.');
  }

  const expectedLocation = cleanString_(payload.specificLocation || payload.location);
  if (expectedLocation && normalizeIdentity_(expectedLocation) !== normalizeIdentity_(row[map.location])) {
    throw new Error('Row ' + rowNum + ' storage identity changed. Refresh before saving.');
  }

  const storageChecks = [
    { key: 'storageId', index: map.storageId },
    { key: 'storageLabel', index: map.storageLabel },
    { key: 'locationCode', index: map.locationCode }
  ];
  storageChecks.forEach(function (check) {
    const expectedValue = cleanString_(payload[check.key]);
    if (!expectedValue || check.index === -1) return;
    const actualValue = getOptionalValue_(row, check.index);
    if (actualValue && normalizeIdentity_(expectedValue) !== normalizeIdentity_(actualValue)) {
      throw new Error('Row ' + rowNum + ' storage identity changed. Refresh before saving.');
    }
  });
}

function normalizeIdentity_(value) {
  return cleanString_(value).toLowerCase().replace(/\s+/g, ' ');
}

function findLocationContextRow_(values, map, room, loc) {
  const roomNeedle = cleanString_(room).toLowerCase();
  const locNeedle = cleanString_(loc).toLowerCase();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (rowMatchesRoomLoc_(row, map, roomNeedle, locNeedle)) {
      return { row: row, rowNumber: i + 1 };
    }
  }
  return null;
}

function countLocationRows_(values, map, room, loc) {
  const roomNeedle = cleanString_(room).toLowerCase();
  const locNeedle = cleanString_(loc).toLowerCase();
  let total = 0;
  let hasPlaceholder = false;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!rowMatchesRoomLoc_(row, map, roomNeedle, locNeedle)) continue;
    total += 1;
    if (isPlaceholderRow_(row, map)) hasPlaceholder = true;
  }
  return { total: total, hasPlaceholder: hasPlaceholder };
}

function buildNewInventoryRow_(width, map, contextRow, item) {
  const row = new Array(width).fill('');
  const room = cleanString_(contextRow[map.room]);
  const loc = cleanString_(contextRow[map.location]);
  const storageId = getOptionalValue_(contextRow, map.storageId);
  const storageLabel = getOptionalValue_(contextRow, map.storageLabel);
  const locationCode = getOptionalValue_(contextRow, map.locationCode);
  const storageType = getOptionalValue_(contextRow, map.storageType);
  const baseUrl = getWebAppBaseUrl_({ silent: true });
  const routeLoc = getPreferredRouteLocation_(room, loc, storageId, storageLabel, locationCode);

  row[map.itemId] = item.itemId;
  row[map.itemName] = item.itemName;
  row[map.room] = room;
  row[map.location] = loc;
  row[map.qty] = item.qty;
  row[map.category] = item.category;
  row[map.status] = item.status;
  setOptionalOutputValue_(row, map.unit, item.unit);
  setOptionalOutputValue_(row, map.remarks, item.remarks);
  setOptionalOutputValue_(row, map.locationCode, locationCode);
  setOptionalOutputValue_(row, map.storageId, storageId);
  setOptionalOutputValue_(row, map.storageLabel, storageLabel);
  setOptionalOutputValue_(row, map.storageType, storageType);
  setOptionalOutputValue_(row, map.isPlaceholder, 'FALSE');
  setOptionalOutputValue_(row, map.lastUpdated, item.timestamp);
  setOptionalOutputValue_(row, map.updatedBy, item.actor);
  if (map.qrLink !== -1 && baseUrl) row[map.qrLink] = buildLocationUrl_(baseUrl, room, routeLoc);
  return row;
}

function buildPlaceholderRowFromContext_(width, map, contextRow, actor, timestamp) {
  const row = new Array(width).fill('');
  const room = cleanString_(contextRow[map.room]);
  const loc = cleanString_(contextRow[map.location]);
  const storageId = getOptionalValue_(contextRow, map.storageId);
  const storageLabel = getOptionalValue_(contextRow, map.storageLabel);
  const locationCode = getOptionalValue_(contextRow, map.locationCode);
  const storageType = getOptionalValue_(contextRow, map.storageType);
  const baseUrl = getWebAppBaseUrl_({ silent: true });
  const routeLoc = getPreferredRouteLocation_(room, loc, storageId, storageLabel, locationCode);

  row[map.room] = room;
  row[map.location] = loc;
  row[map.qty] = 0;
  row[map.category] = 'Storage';
  row[map.status] = 'Good';
  setOptionalOutputValue_(row, map.remarks, 'Placeholder row for QR/location page');
  setOptionalOutputValue_(row, map.locationCode, locationCode);
  setOptionalOutputValue_(row, map.storageId, storageId);
  setOptionalOutputValue_(row, map.storageLabel, storageLabel);
  setOptionalOutputValue_(row, map.storageType, storageType);
  setOptionalOutputValue_(row, map.isPlaceholder, 'TRUE');
  setOptionalOutputValue_(row, map.lastUpdated, timestamp);
  setOptionalOutputValue_(row, map.updatedBy, actor);
  if (map.qrLink !== -1 && baseUrl) row[map.qrLink] = buildLocationUrl_(baseUrl, room, routeLoc);
  return row;
}

function setQrImageFormulaForRow_(sheet, map, rowNumber) {
  if (map.qrImage === -1 || map.qrLink === -1) return;
  const qrLinkColA1 = columnLetter_(map.qrLink + 1);
  sheet.getRange(rowNumber, map.qrImage + 1)
    .setFormula('=IF(' + qrLinkColA1 + rowNumber + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(' + qrLinkColA1 + rowNumber + ')))');
}

function generateItemId_(room, loc, itemName) {
  const base = [room, loc, itemName].map(function (part) {
    return cleanString_(part).toUpperCase().replace(/[^A-Z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }).filter(Boolean).join('-');
  const suffix = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Etc/GMT', 'yyyyMMddHHmmss');
  return (base || 'ITEM') + '-' + suffix;
}

function isValidQrLink_(url, room) {
  const value = cleanString_(url);
  if (!value) return false;
  if (!/^https:\/\/script\.google\.com\//.test(value)) return false;
  if (cleanString_(room) === 'V++' && value.indexOf('room=V%2B%2B') === -1) return false;
  return value.indexOf('/exec?') !== -1 && value.indexOf('room=') !== -1 && value.indexOf('loc=') !== -1;
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
    safetyNote: getOptionalValue_(sourceRow, map.safetyNote),
    reorderLevel: getOptionalValue_(sourceRow, map.reorderLevel),
    supplier: getOptionalValue_(sourceRow, map.supplier),
    purchaseLink: getOptionalValue_(sourceRow, map.purchaseLink),
    assetValue: getOptionalValue_(sourceRow, map.assetValue),
    maintenanceDue: getOptionalValue_(sourceRow, map.maintenanceDue),
    sdsLink: getOptionalValue_(sourceRow, map.sdsLink),
    locationCode: locationCode,
    storageId: storageId,
    storageLabel: storageLabel,
    isHazard: isHazardCategory_(category),
    statusClass: statusClassServer_(status),
    displayLocation: storageLabel || locVal,
    routeLoc: getPreferredRouteLocation_(roomVal, locVal, storageId, storageLabel, locationCode)
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

function appendAuditEvents_(events) {
  const safeEvents = (events || []).filter(function (event) { return !!event; });
  if (!safeEvents.length) return { success: true, appendedRows: 0 };

  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.AUDIT_LOG_SHEET_NAME);
  const desiredHeaders = [
    'Timestamp',
    'User',
    'Action',
    'Room',
    'Storage ID',
    'Specific Location',
    'Route / Location Code',
    'Item ID',
    'Item Name',
    'Old Qty',
    'New Qty',
    'Old Status',
    'New Status',
    'Notes'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    sheet.setFrozenRows(1);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(cleanString_);
    const normalizedExisting = existingHeaders.map(normalizeHeader_);
    desiredHeaders.forEach(function (header) {
      if (normalizedExisting.indexOf(normalizeHeader_(header)) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        normalizedExisting.push(normalizeHeader_(header));
      }
    });
  }

  const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(cleanString_);
  const headerIndex = {};
  actualHeaders.forEach(function (header, index) {
    headerIndex[normalizeHeader_(header)] = index;
  });
  const setValue = function (row, header, value) {
    const idx = headerIndex[normalizeHeader_(header)];
    if (idx !== undefined) row[idx] = value;
  };
  const rows = safeEvents.map(function (event) {
    const row = new Array(actualHeaders.length).fill('');
    setValue(row, 'Timestamp', event.timestamp);
    setValue(row, 'User', event.user);
    setValue(row, 'Action', event.action);
    setValue(row, 'Room', event.room);
    setValue(row, 'Storage ID', event.storageId);
    setValue(row, 'Specific Location', event.location);
    setValue(row, 'Route / Location Code', event.routeLoc || event.locationCode || '');
    setValue(row, 'Item ID', event.itemId);
    setValue(row, 'Item Name', event.itemName);
    setValue(row, 'Old Qty', event.oldQty);
    setValue(row, 'New Qty', event.newQty);
    setValue(row, 'Old Status', event.oldStatus);
    setValue(row, 'New Status', event.newStatus);
    setValue(row, 'Notes', event.notes);
    return row;
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, actualHeaders.length).setValues(rows);
  return { success: true, appendedRows: rows.length };
}

function readRecentAuditEvents_(limit) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.AUDIT_LOG_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const maxRows = Math.max(1, Number(limit) || 25);
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(normalizeHeader_);
    const index = function (names) {
      const aliases = Array.isArray(names) ? names : [names];
      for (let i = 0; i < aliases.length; i++) {
        const needle = normalizeHeader_(aliases[i]);
        const found = headers.indexOf(needle);
        if (found !== -1) return found;
      }
      return -1;
    };
    const map = {
      timestamp: index(['timestamp', 'time']),
      user: index(['user', 'updated by']),
      action: index('action'),
      room: index('room'),
      storageId: index(['storage id', 'storage']),
      location: index(['specific location', 'location']),
      routeLoc: index(['route / location code', 'location code', 'route']),
      itemId: index('item id'),
      itemName: index('item name'),
      oldQty: index('old qty'),
      newQty: index('new qty'),
      oldStatus: index('old status'),
      newStatus: index('new status'),
      notes: index(['notes', 'note'])
    };

    return values.slice(1).filter(function (row) {
      return row.some(function (cell) { return cleanString_(cell); });
    }).slice(-maxRows).reverse().map(function (row) {
      const storage = getOptionalValue_(row, map.storageId) || getOptionalValue_(row, map.location);
      const itemName = getOptionalValue_(row, map.itemName);
      const itemId = getOptionalValue_(row, map.itemId);
      return {
        ts: getOptionalValue_(row, map.timestamp),
        user: getOptionalValue_(row, map.user) || 'unknown user',
        action: getOptionalValue_(row, map.action) || 'Update',
      room: getOptionalValue_(row, map.room),
        storage: storage,
        storageId: getOptionalValue_(row, map.storageId),
        location: getOptionalValue_(row, map.location),
        routeLoc: getOptionalValue_(row, map.routeLoc),
      item: itemName || itemId || '-',
        itemId: itemId,
        itemName: itemName,
        oldQty: getOptionalValue_(row, map.oldQty),
        newQty: getOptionalValue_(row, map.newQty),
        oldStatus: getOptionalValue_(row, map.oldStatus),
        newStatus: getOptionalValue_(row, map.newStatus),
        notes: getOptionalValue_(row, map.notes)
      };
    });
  } catch (err) {
    return [];
  }
}

function getActiveUserEmail_() {
  return getActiveUserEmailSafe_() || 'unknown user';
}

function getActiveUserEmailSafe_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return cleanString_(email).toLowerCase();
  } catch (err) {
    return '';
  }
}

function getUpdateAuthConfig_() {
  const props = PropertiesService.getScriptProperties();
  const allowedEmails = splitPropertyList_(props.getProperty(CONFIG.UPDATE_AUTH_ALLOWED_EMAILS_PROPERTY)).map(function (email) {
    return email.toLowerCase();
  });
  const allowedDomains = splitPropertyList_(props.getProperty(CONFIG.UPDATE_AUTH_ALLOWED_DOMAINS_PROPERTY)).map(function (domain) {
    return domain.replace(/^@+/, '').toLowerCase();
  });
  const pinHash = cleanString_(props.getProperty(CONFIG.UPDATE_MODE_PIN_SHA256_PROPERTY)).toLowerCase();
  const pinSalt = cleanString_(props.getProperty(CONFIG.UPDATE_MODE_PIN_SALT_PROPERTY));
  const disabled = isTruthyProperty_(props.getProperty(CONFIG.UPDATE_AUTH_DISABLED_PROPERTY));
  const pinEnabled = !!(pinHash && pinSalt);
  return {
    allowedEmails: allowedEmails,
    allowedDomains: allowedDomains,
    pinHash: pinHash,
    pinSalt: pinSalt,
    pinEnabled: pinEnabled,
    disabled: disabled,
    configured: disabled || allowedEmails.length > 0 || allowedDomains.length > 0 || pinEnabled
  };
}

function getUpdateAuthClientState_() {
  const config = getUpdateAuthConfig_();
  const email = getActiveUserEmailSafe_();
  const activeUserAllowed = isActiveUserAllowedForUpdate_(config, email);
  return {
    configured: config.configured,
    pinEnabled: config.pinEnabled,
    activeUserAvailable: !!email,
    activeUserAllowed: activeUserAllowed || config.disabled,
    authDisabled: config.disabled,
    tokenTtlSeconds: CONFIG.UPDATE_AUTH_TOKEN_TTL_SECONDS,
    message: config.configured
      ? (activeUserAllowed || config.disabled ? 'Update Mode is available for this session.' : 'Unlock Update Mode before changing live inventory.')
      : 'Update Mode is not configured. Ask an administrator to configure update authorization.'
  };
}

function getUpdateAuthClientStateSafe_() {
  try {
    return getUpdateAuthClientState_();
  } catch (err) {
    return {
      configured: false,
      pinEnabled: false,
      activeUserAvailable: false,
      activeUserAllowed: false,
      authDisabled: false,
      tokenTtlSeconds: CONFIG.UPDATE_AUTH_TOKEN_TTL_SECONDS,
      message: 'Update Mode authorization state is unavailable.'
    };
  }
}

function authorizeUpdateMode(pin) {
  const config = getUpdateAuthConfig_();
  if (!config.configured) {
    throw new Error('Update Mode is not configured. Ask an administrator to configure update authorization.');
  }

  const email = getActiveUserEmailSafe_();
  if (config.disabled) {
    return buildUpdateAuthSuccess_('disabled', getActiveUserEmail_());
  }
  if (isActiveUserAllowedForUpdate_(config, email)) {
    return buildUpdateAuthSuccess_('account', email);
  }

  if (!config.pinEnabled) {
    throw new Error('This account is not authorised for Update Mode, and PIN unlock is not configured.');
  }
  const providedPin = cleanString_(pin);
  if (!providedPin) {
    throw new Error('Enter the Update Mode PIN.');
  }
  const computedHash = computeSha256Hex_(config.pinSalt + providedPin);
  if (!constantTimeStringEquals_(computedHash, config.pinHash)) {
    throw new Error('Update Mode PIN is incorrect.');
  }

  return buildUpdateAuthSuccess_('pin', email || 'pin unlock');
}

function buildUpdateAuthSuccess_(method, user) {
  const token = Utilities.getUuid();
  const cachePayload = JSON.stringify({
    method: method,
    user: user || 'unknown user',
    createdAt: new Date().toISOString()
  });
  CacheService.getScriptCache().put(CONFIG.UPDATE_AUTH_TOKEN_PREFIX + token, cachePayload, CONFIG.UPDATE_AUTH_TOKEN_TTL_SECONDS);
  return {
    success: true,
    token: token,
    method: method,
    expiresInSeconds: CONFIG.UPDATE_AUTH_TOKEN_TTL_SECONDS,
    updateAuth: getUpdateAuthClientState_(),
    message: 'Update Mode unlocked for this browser session.'
  };
}

function requireUpdateAuthorization_(payload) {
  const config = getUpdateAuthConfig_();
  if (!config.configured) {
    throw new Error('Update Mode is not configured. Ask an administrator to configure update authorization.');
  }
  if (config.disabled) {
    return { method: 'disabled', user: getActiveUserEmail_() };
  }

  const email = getActiveUserEmailSafe_();
  if (isActiveUserAllowedForUpdate_(config, email)) {
    return { method: 'account', user: email };
  }

  const token = cleanString_(payload && (payload.updateAuthToken || payload.authToken));
  const tokenInfo = getUpdateAuthTokenInfo_(token);
  if (tokenInfo) {
    return { method: tokenInfo.method || 'token', user: email || tokenInfo.user || 'pin unlock' };
  }

  throw new Error('Update Mode is locked. Unlock with an authorised account or PIN before changing live inventory.');
}

function getUpdateAuthTokenInfo_(token) {
  if (!token) return false;
  try {
    const cached = CacheService.getScriptCache().get(CONFIG.UPDATE_AUTH_TOKEN_PREFIX + token);
    if (!cached) return false;
    const parsed = JSON.parse(cached);
    return parsed && parsed.createdAt ? parsed : null;
  } catch (err) {
    return null;
  }
}

function isActiveUserAllowedForUpdate_(config, email) {
  const value = cleanString_(email).toLowerCase();
  if (!value) return false;
  if (config.allowedEmails.indexOf(value) !== -1) return true;
  const domain = value.indexOf('@') !== -1 ? value.split('@').pop().toLowerCase() : '';
  return !!(domain && config.allowedDomains.indexOf(domain) !== -1);
}

function splitPropertyList_(value) {
  return cleanString_(value).split(',').map(function (part) {
    return cleanString_(part);
  }).filter(Boolean);
}

function isTruthyProperty_(value) {
  return ['1', 'true', 'yes', 'y', 'on'].indexOf(cleanString_(value).toLowerCase()) !== -1;
}

function computeSha256Hex_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(function (byte) {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function constantTimeStringEquals_(left, right) {
  const a = cleanString_(left);
  const b = cleanString_(right);
  let diff = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function ensureAppColumns_() {
  const sheet = getInventorySheetForPreparation_();
  const desiredKeys = CONFIG.REQUIRED_COLUMNS.concat(CONFIG.OPTIONAL_COLUMNS);
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

function getInventorySheetForPreparation_() {
  const ss = getSpreadsheet_();
  const configuredName = cleanString_(PropertiesService.getScriptProperties().getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));
  const targetName = configuredName || CONFIG.DEFAULT_SHEET_NAME;
  return ss.getSheetByName(targetName) || ss.insertSheet(targetName);
}

function buildStorageMasterSheet() {
  const baseUrl = getWebAppBaseUrl_({ silent: true });
  const locations = getAllLocations_();
  const statsByKey = getLocationStatsByKey_();
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.STORAGE_MASTER_SHEET_NAME);
  const headers = [
    'Room',
    'Location Code',
    'Specific Location',
    'Storage ID',
    'Storage Label',
    'Storage Type',
    'Open View',
    'Open Update',
    'QR Link',
    'QR Image',
    'Item Count',
    'Chemical Count',
    'Attention Count',
    'Placeholder Only',
    'Status / Notes'
  ];
  const rows = locations.map(function (entry, index) {
    const stats = statsByKey[entry.canonicalKey] || {};
    const viewUrl = baseUrl ? buildLocationUrl_(baseUrl, entry.room, entry.routeLoc) : '';
    const updateUrl = viewUrl ? viewUrl + '&mode=tech' : '';
    const status = (stats.attentionCount || 0) > 0 ? 'Needs Attention' : 'Good';
    const storageType = entry.storageType || inferStorageType_(entry, stats);
    const itemCount = stats.itemCount || 0;
    const chemicalCount = stats.chemicalCount || 0;
    const attentionCount = stats.attentionCount || 0;
    const placeholderOnly = itemCount === 0;
    const notes = [
      status,
      itemCount + ' item row(s)',
      chemicalCount + ' chemical row(s)',
      attentionCount + ' attention row(s)',
      placeholderOnly ? 'placeholder-only route' : ''
    ].filter(Boolean).join('; ');
    const rowNumber = index + 2;
    return [
      entry.room,
      entry.locationCode,
      entry.loc,
      entry.storageId,
      entry.storageLabel || entry.displayLoc || entry.loc,
      storageType,
      viewUrl ? buildSpreadsheetHyperlinkFormula_(viewUrl, 'Open View') : '',
      updateUrl ? buildSpreadsheetHyperlinkFormula_(updateUrl, 'Open Update') : '',
      viewUrl,
      viewUrl ? '=IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(I' + rowNumber + '))' : '',
      itemCount,
      chemicalCount,
      attentionCount,
      placeholderOnly ? 'Yes' : 'No',
      notes
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 7, rows.length, 1).setFormulas(rows.map(function (row) { return [row[6]]; }));
    sheet.getRange(2, 8, rows.length, 1).setFormulas(rows.map(function (row) { return [row[7]]; }));
    sheet.getRange(2, 10, rows.length, 1).setFormulas(rows.map(function (row) { return [row[9]]; }));
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(1, 1, Math.max(rows.length + 1, 2), headers.length).createFilter();
  } catch (err) {
    // Filters are a convenience; generation should not fail if the sheet cannot create one.
  }
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 520);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(15, 340);

  return { success: true, sheetName: sheet.getName(), storageCount: rows.length };
}

function buildSpreadsheetHyperlinkFormula_(url, label) {
  if (!url) return '';
  return '=HYPERLINK("' + escapeSpreadsheetFormulaString_(url) + '","' + escapeSpreadsheetFormulaString_(label || url) + '")';
}

function escapeSpreadsheetFormulaString_(value) {
  return cleanString_(value).replace(/"/g, '""');
}

function getLocationStatsByKey_() {
  const dataset = getInventoryDataset_();
  const map = dataset.map;
  const values = dataset.values;
  const statsByKey = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = cleanString_(row[map.room]);
    const loc = cleanString_(row[map.location]);
    if (!room || !loc) continue;

    const entry = buildLocationEntry_(row, map, room, loc);
    if (!statsByKey[entry.canonicalKey]) {
      statsByKey[entry.canonicalKey] = {
        itemCount: 0,
        chemicalCount: 0,
        attentionCount: 0,
        categories: {}
      };
    }
    if (!isInventoryItemRow_(row, map)) continue;

    const category = cleanString_(row[map.category]);
    const status = normalizeStatus_(row[map.status]);
    statsByKey[entry.canonicalKey].itemCount += 1;
    if (category) {
      statsByKey[entry.canonicalKey].categories[category.toLowerCase()] = true;
    }
    if (isHazardCategory_(category)) statsByKey[entry.canonicalKey].chemicalCount += 1;
    if (status !== 'Good') statsByKey[entry.canonicalKey].attentionCount += 1;
  }

  return statsByKey;
}

function inferStorageType_(entry, stats) {
  const text = [entry.storageLabel, entry.loc, entry.locationCode, entry.storageId].filter(Boolean).join(' ').toLowerCase();
  const categories = stats && stats.categories ? Object.keys(stats.categories).join(' ') : '';
  if ((stats && stats.chemicalCount) || text.indexOf('chem') !== -1 || categories.indexOf('chemical') !== -1) return 'Chemical Storage';
  if (text.indexOf('machine') !== -1 || categories.indexOf('machine') !== -1) return 'Machine Zone';
  if (text.indexOf('elect') !== -1 || text.indexOf('arduino') !== -1 || categories.indexOf('electronics') !== -1) return 'Electronics Storage';
  if (text.indexOf('tool') !== -1 || categories.indexOf('tool') !== -1) return 'Tool Storage';
  if (text.indexOf('rack') !== -1 || text.indexOf('material') !== -1 || categories.indexOf('material') !== -1) return 'Material Storage';
  if (text.indexOf('tray') !== -1) return 'Tray Storage';
  if (text.indexOf('trolley') !== -1) return 'Trolley';
  if (text.indexOf('cupboard') !== -1 || text.indexOf('cabinet') !== -1) return 'Cupboard / Cabinet';
  return 'Storage';
}

function buildQrLabelSheet() {
  const baseUrl = getWebAppBaseUrl_();
  const locations = getAllLocations_();
  const statsByKey = getLocationStatsByKey_();
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.QR_LABEL_SHEET_NAME);
  const headers = [
    'Room',
    'Specific Location',
    'Storage ID',
    'Storage Label',
    'Location Code',
    'View URL',
    'Update URL',
    'QR Image Formula',
    'Print Label Text',
    'Brother QL-1110 Label Size',
    'Printer Notes'
  ];
  const output = locations.map(function (entry) {
    const stats = statsByKey[entry.canonicalKey] || {};
    const displayName = entry.displayLoc || entry.loc;
    const viewUrl = buildLocationUrl_(baseUrl, entry.room, entry.routeLoc);
    const techUrl = viewUrl + '&mode=tech';
    const labelLines = [
      'D&T Inventory',
      'Room: ' + entry.room,
      'Storage Code: ' + (entry.routeLoc || entry.locationCode || entry.storageId),
      entry.storageLabel || displayName,
      'Scan to view inventory'
    ];
    if (stats.chemicalCount) labelLines.push('HAZARD STORAGE - CHECK SAFETY FIRST');
    const brotherSize = stats.chemicalCount ? '102mm x 70mm safety' : '102mm x 50mm or 90mm x 29mm';
    const printerNotes = stats.chemicalCount
      ? 'Use Brother 102mm safety preset; 90mm x 29mm is not recommended for hazard labels; print sample first.'
      : 'Use Brother 102mm compact/safety preset, or 90mm x 29mm slim preset for short labels; print sample first; scale 100%.';
    return [
      entry.room,
      entry.loc,
      entry.storageId,
      entry.storageLabel,
      entry.locationCode,
      viewUrl,
      techUrl,
      '',
      labelLines.join('\n'),
      brotherSize,
      printerNotes
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (output.length) {
    sheet.getRange(2, 1, output.length, headers.length).setValues(output);
    for (let i = 0; i < output.length; i++) {
      const row = i + 2;
      sheet.getRange(row, 8).setFormula('=IF(F' + row + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(F' + row + ')))');
    }
    sheet.getRange(2, 9, output.length, 3).setWrap(true);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.setColumnWidths(1, 1, 80);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 520);
  sheet.setColumnWidth(7, 520);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 260);
  sheet.setColumnWidth(10, 160);
  sheet.setColumnWidth(11, 320);

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
    ['419A Storage ID Row Count', report.storageId419ACount],
    ['QR Link Issues', report.qrLinkIssueCount],
    ['QR Image Missing', report.qrImageMissingCount],
    ['Rows by Room', JSON.stringify(report.summaryByRoom || {})],
    ['Rows by Status', JSON.stringify(report.summaryByStatus || {})],
    ['Rows by Category', JSON.stringify(report.summaryByCategory || {})],
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

function createWarningTriageBoard() {
  const ss = getSpreadsheet_();
  let reportSheet = ss.getSheetByName(CONFIG.READINESS_REPORT_SHEET_NAME);
  if (!reportSheet) {
    createReadinessReport();
    reportSheet = ss.getSheetByName(CONFIG.READINESS_REPORT_SHEET_NAME);
  }
  if (!reportSheet) throw new Error('Inventory_Readiness_Report could not be found or created.');

  const values = reportSheet.getDataRange().getValues();
  const headerRowIndex = findReviewHeaderRow_(values, ['severity', 'row', 'issue', 'detail']);
  if (headerRowIndex === -1) {
    throw new Error('Could not find the readiness issue table. Run Create Readiness Report first.');
  }

  const normalizedHeaders = values[headerRowIndex].map(normalizeHeader_);
  const output = [];
  for (let i = headerRowIndex + 1; i < values.length; i++) {
    const source = rowByHeaders_(values[i], normalizedHeaders);
    const severity = cleanString_(source.severity);
    if (!severity || ['warn', 'warning', 'error', 'critical'].indexOf(severity.toLowerCase()) === -1) continue;

    const issueType = cleanString_(source.issue);
    const triage = triageForIssue_(issueType);
    output.push([
      'W-' + padNumber_(output.length + 1, 3),
      i + 1,
      cleanString_(source.room),
      cleanString_(source.location),
      cleanString_(source.itemId),
      cleanString_(source.itemName),
      issueType,
      cleanString_(source.detail),
      severity.toUpperCase() === 'ERROR' ? 'Critical' : triage.severity,
      triage.owner,
      triage.decision,
      triage.action,
      '',
      triage.status,
      ''
    ]);
  }

  const sheet = getOrCreateSheet_(ss, CONFIG.WARNING_TRIAGE_SHEET_NAME);
  const headers = [
    'Warning ID',
    'Source Report Row',
    'Room',
    'Location Code',
    'Item ID',
    'Item Name',
    'Warning Type',
    'Current Message',
    'Severity',
    'Owner',
    'Decision',
    'Action Needed',
    'Due / Review Date',
    'Status',
    'Notes'
  ];

  sheet.clear();
  sheet.getRange(1, 1).setValue('Readiness Warning Triage');
  sheet.getRange(2, 1).setValue('Use this board to decide which warnings block the pilot, which can be deferred, and who owns the follow-up.');
  sheet.getRange(3, 1).setValue('Keep critical errors at zero. QR sample printing may proceed only when warnings are assigned, accepted, or scheduled.');
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  if (output.length) {
    sheet.getRange(6, 1, output.length, headers.length).setValues(output);
  }

  sheet.setFrozenRows(5);
  sheet.getRange(1, 1, 1, headers.length).merge().setFontSize(14).setFontWeight('bold').setBackground('#fef3c7');
  sheet.getRange(2, 1, 2, headers.length).mergeAcross().setWrap(true).setBackground('#fff7ed');
  sheet.getRange(5, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.getRange(6, 9, Math.max(1, output.length), 1).setDataValidation(listValidation_(['High', 'Medium', 'Low', 'Critical']));
  sheet.getRange(6, 10, Math.max(1, output.length), 1).setDataValidation(listValidation_(['Technician', 'HoD', 'Teacher', 'Admin', 'Data Owner']));
  sheet.getRange(6, 11, Math.max(1, output.length), 1).setDataValidation(listValidation_(['Fix Now', 'Defer to Pilot', 'Needs HoD Decision', 'Needs Technician Check', 'Accepted Risk', 'Resolved']));
  sheet.getRange(6, 14, Math.max(1, output.length), 1).setDataValidation(listValidation_(['Fix Now', 'Defer to Pilot', 'Needs HoD Decision', 'Needs Technician Check', 'Accepted Risk', 'Resolved']));
  sheet.setColumnWidths(1, headers.length, 140);
  sheet.setColumnWidth(8, 420);
  sheet.setColumnWidth(12, 360);
  sheet.setColumnWidth(15, 280);
  sheet.getRange(1, 1, Math.max(6, output.length + 5), headers.length).setWrap(true);

  return { success: true, sheetName: sheet.getName(), warningCount: output.length };
}

function prepare419AUnmatchedReviewSheet() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.UNMATCHED_REVIEW_SHEET_NAME);
  const existing = sheet.getDataRange().getValues();
  const headerRowIndex = findReviewHeaderRow_(existing, ['old item name']);
  const existingHeaders = headerRowIndex === -1 ? [] : existing[headerRowIndex].map(normalizeHeader_);
  const rows = [];

  if (headerRowIndex !== -1) {
    for (let i = headerRowIndex + 1; i < existing.length; i++) {
      const source = rowByHeaders_(existing[i], existingHeaders);
      const oldItemName = cleanString_(source.oldItemName || source.itemName);
      if (!oldItemName) continue;
      const decision = cleanString_(source.reviewerDecision);
      rows.push([
        oldItemName,
        cleanString_(source.oldRoom || source.room),
        cleanString_(source.oldLocation || source.location),
        cleanString_(source.oldCategory || source.category),
        cleanString_(source.oldQty || source.qty),
        cleanString_(source.possibleMatchSuggestedLocationCode || source.suggestedLocationCode),
        cleanString_(source.confidence) || 'Low',
        cleanString_(source.reasonUnmatched) || 'Not matched automatically to authoritative 419A Location Code list.',
        decision && decision !== 'Pending Review' ? decision : 'Keep for Later Review',
        cleanString_(source.finalLocationCode),
        cleanString_(source.newStorageLabel),
        cleanString_(source.action) || 'Needs Physical Check',
        cleanString_(source.notes)
      ]);
    }
  }

  const headers = [
    'Old Item Name',
    'Old Room',
    'Old Location',
    'Old Category',
    'Old Qty',
    'Possible Match / Suggested Location Code',
    'Confidence',
    'Reason Unmatched',
    'Reviewer Decision',
    'Final Location Code',
    'New Storage Label',
    'Action',
    'Notes'
  ];

  sheet.clear();
  sheet.getRange(1, 1).setValue('419A Unmatched Review');
  sheet.getRange(2, 1).setValue('Physically check the item before assigning it to a Location Code. Do not append rows to Inventory until Final Location Code is confirmed.');
  sheet.getRange(3, 1).setValue('Location Code is the operational 419A route identity. Only reviewed rows should be imported or appended.');
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);

  sheet.setFrozenRows(5);
  sheet.getRange(1, 1, 1, headers.length).merge().setFontSize(14).setFontWeight('bold').setBackground('#fef3c7');
  sheet.getRange(2, 1, 2, headers.length).mergeAcross().setWrap(true).setBackground('#fff7ed');
  sheet.getRange(5, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.getRange(6, 7, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['High', 'Medium', 'Low']));
  sheet.getRange(6, 9, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['Assign to Location Code', 'Archive / Remove', 'Duplicate of Existing', 'Needs Physical Check', 'Not 419A', 'Keep for Later Review']));
  sheet.getRange(6, 12, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['Needs Physical Check', 'Append to Inventory', 'Archive / Remove', 'Do Not Import', 'Merge with Existing']));
  sheet.setColumnWidths(1, headers.length, 160);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(6, 220);
  sheet.setColumnWidth(8, 320);
  sheet.setColumnWidth(13, 300);
  sheet.getRange(1, 1, Math.max(6, rows.length + 5), headers.length).setWrap(true);

  return { success: true, sheetName: sheet.getName(), reviewRows: rows.length };
}

function createPilotTestLog() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CONFIG.PILOT_TEST_LOG_SHEET_NAME);
  const samples = pilotSampleRows_();
  const headers = [
    'Timestamp',
    'Tester',
    'Role',
    'Storage Code',
    'Test Type',
    'Expected Result',
    'Actual Result',
    'Pass/Fail',
    'Issue',
    'Follow-up Owner',
    'Notes'
  ];
  const rows = [];
  samples.forEach(function (sample) {
    rows.push(['', '', 'Student/Staff', sample.code, 'View scan', 'QR opens the correct View Mode page for ' + sample.code + '.', '', '', '', '', sample.note]);
    rows.push(['', '', 'Technician', sample.code, 'Update save', 'Authorised staff can enter Update Mode, save a safe change, and restore it.', '', '', '', '', sample.note]);
    rows.push(['', '', 'Admin/HoD', sample.code, 'QR label check', 'Printed label is readable and scans to View Mode by default.', '', '', '', '', sample.note]);
  });

  sheet.clear();
  sheet.getRange(1, 1).setValue('419A Controlled Pilot Test Log');
  sheet.getRange(2, 1).setValue('Pilot scope: 419A-CAB-01, 419A-FCU-01, and one additional non-chemical 419A storage. Record actual results before wider rollout.');
  sheet.getRange(3, 1).setValue('Do not run full QR printing until sample labels scan correctly and readiness warnings are accepted, assigned, or scheduled.');
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);

  sheet.setFrozenRows(5);
  sheet.getRange(1, 1, 1, headers.length).merge().setFontSize(14).setFontWeight('bold').setBackground('#dcfce7');
  sheet.getRange(2, 1, 2, headers.length).mergeAcross().setWrap(true).setBackground('#f0fdf4');
  sheet.getRange(5, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0f2fe');
  sheet.getRange(6, 3, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['Student/Staff', 'Technician', 'Teacher', 'Admin/HoD']));
  sheet.getRange(6, 5, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['View scan', 'Update save', 'Audit log', 'QR label check', 'Hazard warning check', 'Restore test value']));
  sheet.getRange(6, 8, Math.max(1, rows.length), 1).setDataValidation(listValidation_(['Pass', 'Fail', 'Partial', 'Not Tested']));
  sheet.setColumnWidths(1, headers.length, 150);
  sheet.setColumnWidth(6, 360);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(9, 260);
  sheet.setColumnWidth(11, 280);
  sheet.getRange(1, 1, Math.max(6, rows.length + 5), headers.length).setWrap(true);

  return { success: true, sheetName: sheet.getName(), templateRows: rows.length };
}

function triageForIssue_(issueType) {
  const issue = cleanString_(issueType).toUpperCase();
  if (issue.indexOf('SDS') !== -1) {
    return {
      severity: 'High',
      owner: 'Technician',
      decision: 'Needs Technician Check',
      status: 'Needs Technician Check',
      action: 'Add SDS Link for chemical item or record accepted pilot risk with HoD approval.'
    };
  }
  if (issue.indexOf('SAFETY') !== -1 || issue.indexOf('CHEMICAL') !== -1) {
    return {
      severity: 'High',
      owner: 'Technician',
      decision: 'Fix Now',
      status: 'Fix Now',
      action: 'Add Safety Note for chemical storage/item before physical rollout.'
    };
  }
  if (issue.indexOf('REORDER') !== -1 || issue.indexOf('LOW_STOCK') !== -1) {
    return {
      severity: 'Medium',
      owner: 'HoD',
      decision: 'Needs HoD Decision',
      status: 'Needs HoD Decision',
      action: 'Set Reorder Level or defer purchasing metadata until after pilot.'
    };
  }
  if (issue.indexOf('MAINTENANCE') !== -1) {
    return {
      severity: 'Medium',
      owner: 'Technician',
      decision: 'Needs Technician Check',
      status: 'Needs Technician Check',
      action: 'Add Maintenance Due date or maintenance detail in Remarks.'
    };
  }
  if (issue.indexOf('STORAGE') !== -1 || issue.indexOf('LOCATION') !== -1) {
    return {
      severity: 'Medium',
      owner: 'Data Owner',
      decision: 'Needs Technician Check',
      status: 'Needs Technician Check',
      action: 'Confirm Location Code / storage identity against the physical room.'
    };
  }
  return {
    severity: 'Low',
    owner: 'Data Owner',
    decision: 'Defer to Pilot',
    status: 'Defer to Pilot',
    action: 'Review during pilot data cleanup.'
  };
}

function pilotSampleRows_() {
  const samples = [
    { code: '419A-CAB-01', note: 'Chemical storage sample; check hazard wording and View Mode default.' },
    { code: '419A-FCU-01', note: 'Normal storage sample; already used for safe Update Mode/Audit_Log test.' }
  ];
  const fallback = findAdditionalPilotStorage_();
  if (fallback) samples.push({ code: fallback, note: 'Additional non-chemical 419A storage sample.' });
  return samples;
}

function findAdditionalPilotStorage_() {
  const locations = getAllLocations_();
  const statsByKey = getLocationStatsByKey_();
  for (let i = 0; i < locations.length; i++) {
    const entry = locations[i];
    const code = entry.routeLoc || entry.locationCode || entry.storageId || entry.loc;
    if (entry.room !== CONFIG.ROLLOUT_ROOM) continue;
    if (code === '419A-CAB-01' || code === '419A-FCU-01') continue;
    const stats = statsByKey[entry.canonicalKey] || {};
    if ((stats.chemicalCount || 0) > 0) continue;
    return code;
  }
  return '';
}

function findReviewHeaderRow_(values, requiredHeaders) {
  for (let i = 0; i < values.length; i++) {
    const normalized = values[i].map(normalizeHeader_);
    const hasAll = requiredHeaders.every(function (header) {
      return normalized.indexOf(normalizeHeader_(header)) !== -1;
    });
    if (hasAll) return i;
  }
  return -1;
}

function rowByHeaders_(row, normalizedHeaders) {
  const output = {};
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (!normalizedHeaders[i]) continue;
    output[normalizedHeaders[i]] = row[i];
    output[camelHeaderKey_(normalizedHeaders[i])] = row[i];
  }
  return output;
}

function camelHeaderKey_(header) {
  const parts = String(header || '').split(/[^a-z0-9]+/).filter(Boolean);
  if (!parts.length) return '';
  return parts[0] + parts.slice(1).map(function (part) {
    return part.charAt(0).toUpperCase() + part.slice(1);
  }).join('');
}

function listValidation_(items) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(false)
    .build();
}

function padNumber_(value, width) {
  const text = String(value);
  return text.length >= width ? text : new Array(width - text.length + 1).join('0') + text;
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
      if (entry.locationCode || entry.storageId) {
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
    rowOut[targetMap.qty] = 0;
    rowOut[targetMap.category] = 'Storage';
    rowOut[targetMap.status] = 'Good';
    setOptionalOutputValue_(rowOut, targetMap.remarks, 'Placeholder row for QR/location page');
    setOptionalOutputValue_(rowOut, targetMap.isPlaceholder, true);
    setOptionalOutputValue_(rowOut, targetMap.storageType, inferStorageType_({
      storageId: meta.storageId,
      storageLabel: meta.storageLabel,
      loc: meta.displayLocation,
      locationCode: meta.locationCode
    }, {}));
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
  const placeholderCounts = {};
  const roomCounts = {};
  const statusCounts = {};
  const categoryCounts = {};
  let itemRows = 0;
  let emptyStorageRows = 0;
  let errorCount = 0;
  let warningCount = 0;
  let qrLinkIssueCount = 0;
  let qrImageMissingCount = 0;
  let storageId419ACount = 0;

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
  CONFIG.OPTIONAL_COLUMNS.forEach(function (key) {
    if (map[key] === -1) {
      addIssue('WARN', '', null, 'MISSING_OPTIONAL_COLUMN', CONFIG.COLUMN_LABELS[key] + ' is missing. Run Prepare App Columns before full workshop rollout.');
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
    const remarks = getOptionalValue_(row, map.remarks);
    const safetyNote = getOptionalValue_(row, map.safetyNote);
    const reorderLevelText = getOptionalValue_(row, map.reorderLevel);
    const maintenanceDue = getOptionalValue_(row, map.maintenanceDue);
    const isItem = isInventoryItemRow_(row, map);
    const status = normalizeStatus_(row[map.status]);
    const category = cleanString_(row[map.category]) || '(blank)';

    if (!room && !loc && !itemId && !itemName) continue;
    if (room) roomCounts[room] = (roomCounts[room] || 0) + 1;
    if (!room) addIssue('ERROR', sheetRow, row, 'MISSING_ROOM', 'Room is required for browsing and QR links.');
    if (!loc) addIssue('ERROR', sheetRow, row, 'MISSING_LOCATION', 'Specific Location is required for browsing and QR links.');

    if (room && loc) {
      const identity = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode);
      locationCounts[identity.key] = (locationCounts[identity.key] || 0) + 1;
      if (room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase() && !locationCode && !storageId) {
        addIssue('WARN', sheetRow, row, '419A_MISSING_STORAGE_CODE', '419A rollout rows should use Location Code where possible.');
      }
      if (room.toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase() && (locationCode || storageId)) storageId419ACount += 1;
      if (map.qrLink !== -1 && webAppBaseUrl) {
        const expected = buildLocationUrl_(webAppBaseUrl, room, identity.routeLoc);
        const current = getOptionalValue_(row, map.qrLink);
        if (!current) {
          qrLinkIssueCount += 1;
          addIssue('WARN', sheetRow, row, 'QR_LINK_MISSING', 'QR link is blank.');
        } else if (!isValidQrLink_(current, room)) {
          qrLinkIssueCount += 1;
          addIssue('WARN', sheetRow, row, 'QR_LINK_INVALID', 'QR link is not a valid deployed /exec storage URL.');
        } else if (current !== expected) {
          qrLinkIssueCount += 1;
          addIssue('WARN', sheetRow, row, 'QR_LINK_OUTDATED', 'Expected ' + expected);
        }
        if (room === 'V++' && current && current.indexOf('room=V%2B%2B') === -1 && current.indexOf('room=V++') !== -1) {
          addIssue('WARN', sheetRow, row, 'VPP_ENCODING_RISK', 'V++ should be URL-encoded as V%2B%2B in QR links.');
        }
      }
      if (map.qrImage !== -1 && map.qrLink !== -1 && getOptionalValue_(row, map.qrLink)) {
        const qrImageValue = cleanString_(row[map.qrImage]);
        if (!qrImageValue) {
          qrImageMissingCount += 1;
          addIssue('WARN', sheetRow, row, 'QR_IMAGE_MISSING', 'QR Code Image formula/value is blank.');
        }
      }
    }

    if (isItem) {
      itemRows += 1;
      statusCounts[status] = (statusCounts[status] || 0) + 1;
      categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      const qtyText = cleanString_(row[map.qty]);
      const qty = Number(row[map.qty]);
      if (qtyText === '' || !Number.isFinite(qty) || qty < 0) {
        addIssue('ERROR', sheetRow, row, 'INVALID_QTY', 'Quantity must be numeric and non-negative.');
      }
      if (!matchStatus_(row[map.status])) {
        addIssue('ERROR', sheetRow, row, 'INVALID_STATUS', 'Allowed values: ' + CONFIG.STATUS_OPTIONS.join(', '));
      }
      if (!itemName) addIssue('WARN', sheetRow, row, 'MISSING_ITEM_NAME', 'Item rows should include Item Name.');
      if (isHazardCategory_(category) && !remarks && !safetyNote) {
        addIssue('WARN', sheetRow, row, 'CHEMICAL_SAFETY_NOTE_MISSING', 'Chemical rows should include a remark or Safety Note.');
      }
      if (isHazardCategory_(category) && map.sdsLink !== -1 && !getOptionalValue_(row, map.sdsLink)) {
        addIssue('WARN', sheetRow, row, 'CHEMICAL_SDS_LINK_MISSING', 'Chemical rows should include an SDS Link where available.');
      }
      if (status === 'Low Stock' && map.reorderLevel !== -1 && !reorderLevelText) {
        addIssue('WARN', sheetRow, row, 'LOW_STOCK_REORDER_LEVEL_MISSING', 'Low Stock rows should include a Reorder Level for purchasing review.');
      }
      if (reorderLevelText) {
        const reorderLevel = Number(reorderLevelText);
        const qty = Number(row[map.qty]);
        if (!Number.isFinite(reorderLevel) || reorderLevel < 0) {
          addIssue('WARN', sheetRow, row, 'INVALID_REORDER_LEVEL', 'Reorder Level should be a non-negative number.');
        } else if (Number.isFinite(qty) && qty <= reorderLevel && status === 'Good') {
          addIssue('WARN', sheetRow, row, 'REORDER_THRESHOLD_REACHED', 'Quantity is at or below reorder level; consider Low Stock status.');
        }
      }
      if (status === 'Needs Maintenance' && !maintenanceDue && !remarks) {
        addIssue('WARN', sheetRow, row, 'MAINTENANCE_DETAIL_MISSING', 'Maintenance rows should include a remark or Maintenance Due date.');
      }
      const itemKey = buildInventoryImportKey_(itemId, itemName, room, loc, storageId, storageLabel, locationCode);
      itemCounts[itemKey] = (itemCounts[itemKey] || 0) + 1;
      if (itemCounts[itemKey] > 1) {
        addIssue('WARN', sheetRow, row, 'POSSIBLE_DUPLICATE_ITEM', 'Same item identity appears more than once in this storage.');
      }
    } else if (room && loc) {
      emptyStorageRows += 1;
      const placeholderKey = buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode).key;
      placeholderCounts[placeholderKey] = (placeholderCounts[placeholderKey] || 0) + 1;
      if (placeholderCounts[placeholderKey] > 1) {
        addIssue('WARN', sheetRow, row, 'DUPLICATE_STORAGE_PLACEHOLDER', 'More than one placeholder row exists for this storage identity.');
      }
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
    summaryByRoom: roomCounts,
    summaryByStatus: statusCounts,
    summaryByCategory: categoryCounts,
    qrLinkIssueCount: qrLinkIssueCount,
    qrImageMissingCount: qrImageMissingCount,
    storageId419ACount: storageId419ACount,
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
  const info = getSpreadsheetIdInfo_();
  if (!info.id) {
    throw new Error('SPREADSHEET_ID is not set.');
  }
  try {
    return SpreadsheetApp.openById(info.id);
  } catch (err) {
    throw new Error('Could not open inventory spreadsheet from ' + info.source + ' (' + info.maskedId + '). Check sharing and Apps Script execution permissions. ' + (err.message || String(err)));
  }
}

function getInventorySheet_() {
  const ss = getSpreadsheet_();
  const configuredName = cleanString_(PropertiesService.getScriptProperties().getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));
  const sheets = ss.getSheets();
  if (!sheets.length) throw new Error('No sheets found in the configured spreadsheet.');

  if (configuredName) {
    const configuredSheet = ss.getSheetByName(configuredName);
    if (!configuredSheet) {
      throw new Error('Configured INVENTORY_SHEET_NAME "' + configuredName + '" was not found. Available sheets: ' + sheets.map(function (sheet) { return sheet.getName(); }).join(', ') + '.');
    }
    const configuredInfo = getHeaderInfoForSheet_(configuredSheet);
    if (!configuredInfo.hasRequiredHeaders) {
      throw new Error('Configured inventory sheet "' + configuredName + '" is missing required headers: ' + formatMissingHeaderLabels_(configuredInfo.missingRequired) + '. Header preview: ' + configuredInfo.headerPreview.join(', ') + '.');
    }
    return configuredSheet;
  }

  const inventorySheet = ss.getSheetByName(CONFIG.DEFAULT_SHEET_NAME);
  if (inventorySheet) {
    const inventoryInfo = getHeaderInfoForSheet_(inventorySheet);
    if (inventoryInfo.hasRequiredHeaders) return inventorySheet;
  }

  for (let i = 0; i < sheets.length; i++) {
    const info = getHeaderInfoForSheet_(sheets[i]);
    if (info.hasRequiredHeaders) return sheets[i];
  }

  throw new Error(buildNoInventorySheetMessage_(ss));
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

function getSpreadsheetIdInfo_() {
  const propValue = cleanString_(PropertiesService.getScriptProperties().getProperty(CONFIG.SPREADSHEET_ID_PROPERTY));
  const fallbackValue = cleanString_(CONFIG.DEFAULT_SPREADSHEET_ID);
  const id = propValue || fallbackValue;
  return {
    id: id,
    maskedId: maskSpreadsheetId_(id),
    source: propValue ? 'Script Property SPREADSHEET_ID' : (fallbackValue ? 'fallback spreadsheet ID' : 'missing configuration'),
    fromFallback: !propValue && !!fallbackValue,
    propertySet: !!propValue
  };
}

function maskSpreadsheetId_(id) {
  const value = cleanString_(id);
  if (!value) return '';
  if (value.length <= 12) return value.slice(0, 3) + '...' + value.slice(-3);
  return value.slice(0, 6) + '...' + value.slice(-6);
}

function getHeaderInfoForSheet_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const lastRow = sheet.getLastRow();
  const header = lastRow >= CONFIG.HEADER_ROW
    ? sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getValues()[0]
    : [];
  const normalizedHeaders = header.map(normalizeHeader_);
  const map = getLooseColumnMap_(header);
  const requiredColumns = {};
  const optionalColumns = {};
  const missingRequired = [];
  const foundRequired = [];
  const foundOptional = [];

  CONFIG.REQUIRED_COLUMNS.forEach(function (key) {
    if (map[key] !== -1) {
      requiredColumns[key] = 'found (col ' + (map[key] + 1) + ')';
      foundRequired.push(key);
    } else {
      requiredColumns[key] = 'MISSING';
      missingRequired.push(key);
    }
  });

  CONFIG.OPTIONAL_COLUMNS.forEach(function (key) {
    if (map[key] !== -1) {
      optionalColumns[key] = 'found (col ' + (map[key] + 1) + ')';
      foundOptional.push(key);
    } else {
      optionalColumns[key] = 'not present';
    }
  });

  return {
    name: sheet.getName(),
    rowCount: Math.max(lastRow - CONFIG.HEADER_ROW, 0),
    columnCount: sheet.getLastColumn(),
    headerPreview: normalizedHeaders.filter(function (value) { return !!value; }).slice(0, 24),
    requiredColumns: requiredColumns,
    optionalColumns: optionalColumns,
    foundRequired: foundRequired,
    foundOptional: foundOptional,
    missingRequired: missingRequired,
    hasRequiredHeaders: missingRequired.length === 0,
    score: foundRequired.length
  };
}

function sheetHasRequiredHeaders_(sheet) {
  return getHeaderInfoForSheet_(sheet).hasRequiredHeaders;
}

function describeAvailableSheets_(spreadsheet) {
  return spreadsheet.getSheets().map(function (sheet) {
    const info = getHeaderInfoForSheet_(sheet);
    return {
      name: info.name,
      rowCount: info.rowCount,
      columnCount: info.columnCount,
      hasRequiredHeaders: info.hasRequiredHeaders,
      missingRequired: info.missingRequired.map(function (key) { return CONFIG.COLUMN_LABELS[key]; }),
      foundRequired: info.foundRequired.map(function (key) { return CONFIG.COLUMN_LABELS[key]; }),
      foundOptional: info.foundOptional.map(function (key) { return CONFIG.COLUMN_LABELS[key]; }),
      headerPreview: info.headerPreview,
      score: info.score
    };
  });
}

function formatMissingHeaderLabels_(keys) {
  return (keys || []).map(function (key) {
    return CONFIG.COLUMN_LABELS[key] || key;
  }).join(', ');
}

function buildNoInventorySheetMessage_(spreadsheet) {
  const available = describeAvailableSheets_(spreadsheet);
  const best = available.slice().sort(function (a, b) {
    return (b.score || 0) - (a.score || 0);
  })[0];
  const parts = [
    'No valid inventory sheet was found.',
    'Required headers: ' + CONFIG.REQUIRED_COLUMNS.map(function (key) { return CONFIG.COLUMN_LABELS[key]; }).join(', ') + '.',
    'Available sheets: ' + (available.length ? available.map(function (sheet) { return sheet.name; }).join(', ') : 'none') + '.'
  ];
  if (best) {
    parts.push('Best candidate "' + best.name + '" is missing: ' + (best.missingRequired.length ? best.missingRequired.join(', ') : 'none') + '. Header preview: ' + (best.headerPreview.join(', ') || '-') + '.');
  }
  parts.push('Fix: rename the inventory tab to "Inventory" or set INVENTORY_SHEET_NAME to a sheet that contains all required headers.');
  return parts.join(' ');
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
  if (CONFIG.DEFAULT_WEB_APP_BASE_URL) return normalizeWebAppBaseUrl_(CONFIG.DEFAULT_WEB_APP_BASE_URL);
  const runtimeUrl = getRuntimeWebAppBaseUrl_();
  if (runtimeUrl) return runtimeUrl;

  if (opts.silent) return '';
  throw new Error('WEB_APP_BASE_URL is not configured.');
}

function getRuntimeWebAppBaseUrl_() {
  try {
    if (typeof ScriptApp === 'undefined' || !ScriptApp.getService) return '';
    const url = ScriptApp.getService().getUrl();
    return url ? normalizeWebAppBaseUrl_(url) : '';
  } catch (err) {
    return '';
  }
}

function getExternalScannerUrl_(options) {
  const opts = options || {};
  const propertyUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.EXTERNAL_SCANNER_URL_PROPERTY);
  const raw = cleanString_(propertyUrl) || cleanString_(CONFIG.DEFAULT_EXTERNAL_SCANNER_URL);
  if (!raw) {
    if (opts.silent) return '';
    throw new Error('EXTERNAL_SCANNER_URL is not configured.');
  }
  const normalized = normalizeWebAppBaseUrl_(raw);
  if (!/^https?:\/\//i.test(normalized)) {
    if (opts.silent) return '';
    throw new Error('EXTERNAL_SCANNER_URL must start with http:// or https://.');
  }
  return normalized.replace(/\/?$/, '/');
}

function getAppConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: cleanString_(props.getProperty(CONFIG.SPREADSHEET_ID_PROPERTY)),
    inventorySheetName: cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY)),
    webAppBaseUrl: getWebAppBaseUrl_({ silent: true }),
    externalScannerUrl: getExternalScannerUrl_({ silent: true })
  };
}

function getAppConfigSafe_() {
  try {
    return getAppConfig_();
  } catch (err) {
    return {
      spreadsheetId: '',
      inventorySheetName: '',
      webAppBaseUrl: '',
      externalScannerUrl: ''
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

function getLightDiagnostics_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetInfo = getSpreadsheetIdInfo_();
  const webAppBaseUrl = cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY));
  const inventorySheetName = cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));
  const runtimeWebAppBaseUrl = getRuntimeWebAppBaseUrl_();
  const effectiveWebAppBaseUrl = webAppBaseUrl || CONFIG.DEFAULT_WEB_APP_BASE_URL || runtimeWebAppBaseUrl;
  return {
    light: true,
    maskedSpreadsheetId: spreadsheetInfo.maskedId,
    scriptProperties: {
      spreadsheetIdConfigured: !!spreadsheetInfo.id,
      spreadsheetIdPropertySet: spreadsheetInfo.propertySet,
      spreadsheetIdFromFallback: spreadsheetInfo.fromFallback,
      webAppBaseUrlConfigured: !!effectiveWebAppBaseUrl,
      webAppBaseUrlPropertySet: !!webAppBaseUrl,
      webAppBaseUrlFromFallback: !webAppBaseUrl && !!CONFIG.DEFAULT_WEB_APP_BASE_URL,
      webAppBaseUrlFromRuntime: !webAppBaseUrl && !CONFIG.DEFAULT_WEB_APP_BASE_URL && !!runtimeWebAppBaseUrl,
      inventorySheetNameConfigured: !!inventorySheetName,
      inventorySheetName: inventorySheetName || CONFIG.DEFAULT_SHEET_NAME
    },
    nextAction: 'Open ?admin=diagnostics for full sheet and column diagnostics.'
  };
}

function getConfigStatus() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetInfo = getSpreadsheetIdInfo_();
  const webAppBaseUrl = cleanString_(props.getProperty(CONFIG.WEB_APP_URL_PROPERTY));
  const inventorySheetName = cleanString_(props.getProperty(CONFIG.INVENTORY_SHEET_NAME_PROPERTY));
  const runtimeWebAppBaseUrl = getRuntimeWebAppBaseUrl_();
  const effectiveWebAppBaseUrl = webAppBaseUrl || CONFIG.DEFAULT_WEB_APP_BASE_URL || runtimeWebAppBaseUrl;

  const status = {
    maskedSpreadsheetId: spreadsheetInfo.maskedId,
    spreadsheetSource: spreadsheetInfo.source,
    scriptProperties: {
      spreadsheetIdConfigured: !!spreadsheetInfo.id,
      spreadsheetIdPropertySet: spreadsheetInfo.propertySet,
      spreadsheetIdFromFallback: spreadsheetInfo.fromFallback,
      webAppBaseUrlConfigured: !!effectiveWebAppBaseUrl,
      webAppBaseUrlPropertySet: !!webAppBaseUrl,
      webAppBaseUrlFromFallback: !webAppBaseUrl && !!CONFIG.DEFAULT_WEB_APP_BASE_URL,
      webAppBaseUrlFromRuntime: !webAppBaseUrl && !CONFIG.DEFAULT_WEB_APP_BASE_URL && !!runtimeWebAppBaseUrl,
      inventorySheetNameConfigured: !!inventorySheetName,
      inventorySheetName: inventorySheetName || CONFIG.DEFAULT_SHEET_NAME
    }
  };

  try {
    const ss = getSpreadsheet_();
    status.availableSheets = describeAvailableSheets_(ss);
    status.availableSheetNames = status.availableSheets.map(function (sheet) { return sheet.name; });
    const sheet = getInventorySheet_();
    const info = getHeaderInfoForSheet_(sheet);

    status.sheetInUse = sheet.getName();
    status.dataRows = info.rowCount;
    status.requiredColumns = info.requiredColumns;
    status.optionalColumns = info.optionalColumns;
    status.missingRequiredColumns = info.missingRequired.map(function (key) { return CONFIG.COLUMN_LABELS[key]; });
    status.headerPreview = info.headerPreview;
    status.rollout419A = get419AReadinessSummary(CONFIG.ROLLOUT_ROOM);
    status.nextAction = status.missingRequiredColumns.length
      ? 'Run D&T Inventory -> Prepare App Columns or fix the header row in "' + sheet.getName() + '".'
      : 'Configuration is readable. Run Create Readiness Report before QR printing.';
  } catch (err) {
    status.sheetError = err.message || String(err);
    status.nextAction = 'Open the Google Sheet, verify sharing/access, then run D&T Inventory -> Config Status / Diagnostics.';
  }

  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function getPreferredRouteLocation_(room, loc, storageId, storageLabel, locationCode) {
  if (cleanString_(room).toLowerCase() === CONFIG.ROLLOUT_ROOM.toLowerCase() && locationCode) {
    return locationCode;
  }
  return storageId || storageLabel || locationCode || loc;
}

function buildLocationIdentity_(room, loc, storageId, storageLabel, locationCode) {
  const routeLoc = getPreferredRouteLocation_(room, loc, storageId, storageLabel, locationCode);
  const identityType = routeLoc === locationCode ? 'locationCode' : (routeLoc === storageId ? 'storageId' : (routeLoc === storageLabel ? 'storageLabel' : 'location'));
  const identityValue = routeLoc;
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

// Legacy server-rendered UI helpers. The active @7+ runtime renders
// index.html with app_styles.html and app_script.html; keep these for
// compatibility with older save responses and earlier deployed versions.
function renderInitialItemsHtml_(bootstrap) {
  if (bootstrap.pageType === 'landing') {
    return renderLandingHtml_(bootstrap.locations || []);
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

    const viewUrl = buildLocationHref_('', entry.room, entry.routeLoc, 'view');
    const techUrl = buildLocationHref_('', entry.room, entry.routeLoc, 'tech');
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
    html += '<a class="rounded-xl bg-sky-700 px-3 py-2 text-xs font-medium text-white transition hover:bg-sky-800" href="' + escapeHtml_(techUrl) + '">Update</a>';
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
    : (bootstrap.pageType === 'error' ? 'Configuration' : (mode === 'tech' ? 'Update Mode' : 'View Mode'));

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
    '<a id="techModeLink" class="rounded-xl bg-sky-400 px-3.5 py-2 text-xs font-medium text-slate-950 transition hover:bg-sky-300" href="#">Update</a>' +
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
    'function renderModeBadge(){var badge=document.getElementById("modeBadge");var saveBtn=document.getElementById("saveBtn");var hasRows=!!(APP.bootstrap&&APP.bootstrap.rows&&APP.bootstrap.rows.length);if(APP.bootstrap.pageType==="landing"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Landing";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.bootstrap.pageType==="error"){badge.className="rounded-full bg-red-100 px-2.5 py-1 text-xs font-medium text-red-800";badge.textContent="Configuration";if(saveBtn)saveBtn.classList.add("hidden");return;}if(APP.mode==="tech"){badge.className="rounded-full bg-sky-100 px-2.5 py-1 text-xs font-medium text-sky-800";badge.textContent="Update Mode";if(saveBtn)saveBtn.classList.toggle("hidden",!hasRows);}else{badge.className="rounded-full bg-stone-200 px-2.5 py-1 text-xs font-medium text-stone-700";badge.textContent="View Mode";if(saveBtn)saveBtn.classList.add("hidden");}}' +
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
