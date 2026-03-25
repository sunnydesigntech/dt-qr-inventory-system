/**
 * D&T QR Inventory System - Single-file Google Apps Script implementation.
 */
const CONFIG = {
  SHEET_NAME: 'Inventory',
  HEADER_ROW: 1,
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemicals', 'chemical'],
  QUICKCHART_QR_BASE: 'https://quickchart.io/qr?size=220&text=',
  DEBUG_PANEL: false,
  ALIASES: {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'name'],
    room: ['room'],
    location: ['specific location', 'location', 'specificlocation'],
    qty: ['qty', 'quantity'],
    category: ['category'],
    status: ['status'],
    qrLink: ['qr code link (auto-generated)', 'auto-generated qr link', 'qr link', 'qr url'],
    qrImage: ['qr code image', 'qr image']
  }
};

function doGet(e) {
  const params = {
    room: ((e && e.parameter && e.parameter.room) || '').trim(),
    loc: ((e && e.parameter && e.parameter.loc) || '').trim(),
    mode: (((e && e.parameter && e.parameter.mode) || 'view').toLowerCase() === 'tech') ? 'tech' : 'view'
  };

  const bootstrap = buildBootstrapData_(params);

  return HtmlService
    .createHtmlOutput(buildPageHtml_(params, bootstrap))
    .setTitle('D&T Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buildBootstrapData_(params) {
  const hasLocation = !!(params.room && params.loc);
  if (hasLocation) {
    const result = getInventoryData({ room: params.room, loc: params.loc });
    return {
      pageType: 'location',
      room: params.room,
      loc: params.loc,
      mode: params.mode,
      rows: result.rows || [],
      locations: [],
      message: result.message || '',
      error: ''
    };
  }

  const locResult = getAllLocations();
  return {
    pageType: 'landing',
    room: '',
    loc: '',
    mode: params.mode,
    rows: [],
    locations: locResult.locations || [],
    message: '',
    error: ''
  };
}

function getInventoryData(params) {
  const room = String((params && params.room) || '').trim();
  const loc = String((params && params.loc) || '').trim();

  if (!room || !loc) {
    return {
      success: true,
      rows: [],
      room: room,
      loc: loc,
      message: 'Please scan a valid QR code for a storage location.'
    };
  }

  const rows = getInventoryRowsForLocation_(room, loc);

  return {
    success: true,
    rows: rows,
    room: room,
    loc: loc,
    message: rows.length ? '' : 'No inventory records found for this location.'
  };
}

function getInventoryRowsForLocation_(room, loc) {
  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const roomNeedle = String(room || '').toLowerCase();
  const locNeedle = String(loc || '').toLowerCase();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const roomVal = String(r[map.room] || '').trim();
    const locVal = String(r[map.location] || '').trim();
    if (!roomVal || !locVal) continue;

    if (roomVal.toLowerCase() !== roomNeedle || locVal.toLowerCase() !== locNeedle) continue;

    const category = String(r[map.category] || '').trim();
    rows.push({
      sheetRow: i + 1,
      itemId: String(r[map.itemId] || '').trim(),
      itemName: String(r[map.itemName] || '').trim(),
      room: roomVal,
      specificLocation: locVal,
      qty: toNonNegativeNumber_(r[map.qty]),
      category: category,
      status: normalizeStatus_(r[map.status]),
      isHazard: isHazardCategory_(category)
    });
  }

  return rows;
}

function saveInventoryUpdates(payload) {
  if (!payload || !Array.isArray(payload.updates) || !payload.updates.length) {
    throw new Error('No updates provided.');
  }

  const sheet = getInventorySheet_();
  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: false });

  const qtyCol = map.qty + 1;
  const statusCol = map.status + 1;
  const maxRow = sheet.getLastRow();

  payload.updates.forEach(function (u) {
    const rowNum = Number(u.sheetRow);
    if (!Number.isInteger(rowNum) || rowNum <= CONFIG.HEADER_ROW || rowNum > maxRow) {
      throw new Error('Invalid row number: ' + u.sheetRow);
    }

    const qty = toNonNegativeNumber_(u.qty);
    const status = normalizeStatus_(u.status);
    if (CONFIG.STATUS_OPTIONS.indexOf(status) === -1) {
      throw new Error('Invalid status for row ' + rowNum + ': ' + status);
    }

    sheet.getRange(rowNum, qtyCol).setValue(qty);
    sheet.getRange(rowNum, statusCol).setValue(status);
  });

  return {
    success: true,
    updatedCount: payload.updates.length,
    timestamp: new Date().toISOString()
  };
}

function getAllLocations() {
  return { success: true, locations: getAllLocations_() };
}

function getAllLocations_() {
  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const seen = {};
  const locations = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const room = String(row[map.room] || '').trim();
    const loc = String(row[map.location] || '').trim();
    if (!room || !loc) continue;

    const key = room.toLowerCase() + '||' + loc.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    locations.push({ room: room, loc: loc });
  }

  locations.sort(function (a, b) {
    if (a.room === b.room) return a.loc.localeCompare(b.loc);
    return a.room.localeCompare(b.room);
  });

  return locations;
}

function refreshQrLinks() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true });
  const baseUrl = getWebAppBaseUrl_();

  const range = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, sheet.getLastColumn());
  const rows = range.getValues();

  const output = rows.map(function (r) {
    const room = String(r[map.room] || '').trim();
    const loc = String(r[map.location] || '').trim();
    if (!room || !loc) return [''];

    const url = baseUrl + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
    return [url];
  });

  sheet.getRange(CONFIG.HEADER_ROW + 1, map.qrLink + 1, output.length, 1).setValues(output);
}

/**
 * Optional helper:
 * If a QR image column exists (e.g. "QR Code Image" / "QR Image"), this will
 * populate it with QuickChart IMAGE formulas based on the QR link column.
 * This function is safe to skip when your sheet only has 8 columns.
 */
function refreshQrImages() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = getColumnMap_(header, { requireQrLink: true, requireQrImage: true });
  const qrLinkColA1 = columnLetter_(map.qrLink + 1);
  const qrImageCol = map.qrImage + 1;

  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const formula = '=IF(' + qrLinkColA1 + row + '="","",IMAGE("' + CONFIG.QUICKCHART_QR_BASE + '"&ENCODEURL(' + qrLinkColA1 + row + ')))';
    sheet.getRange(row, qrImageCol).setFormula(formula);
  }
}

/**
 * Optional convenience setter.
 * Run once to set Script Property WEB_APP_BASE_URL.
 */
function setWebAppBaseUrl(url) {
  const value = String(url || '').trim();
  if (!value) throw new Error('Please provide a non-empty URL.');
  PropertiesService.getScriptProperties().setProperty(CONFIG.WEB_APP_URL_PROPERTY, value);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('D&T Inventory')
    .addItem('Refresh QR Links', 'refreshQrLinks')
    .addItem('Open Web App', 'openWebApp_')
    .addItem('Refresh QR Images (Optional)', 'refreshQrImages')
    .addSeparator()
    .addItem('Set WEB_APP_BASE_URL', 'promptSetWebAppBaseUrl_')
    .addToUi();
}

function openWebApp_() {
  const ui = SpreadsheetApp.getUi();
  const url = getWebAppBaseUrl_();
  ui.alert('Open this URL in your browser:\n\n' + url);
}

function promptSetWebAppBaseUrl_() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Set WEB_APP_BASE_URL', 'Paste your deployed /exec web app URL:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  setWebAppBaseUrl(result.getResponseText());
  ui.alert('WEB_APP_BASE_URL saved.');
}

function getInventorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventory = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (inventory) return inventory;

  const first = ss.getSheets()[0];
  if (!first) throw new Error('No sheets found in the active spreadsheet.');
  return first;
}

function getColumnMap_(headerRow, options) {
  const opts = options || {};
  const headers = headerRow.map(normalizeHeader_);

  const map = {
    itemId: findHeaderIndex_(headers, CONFIG.ALIASES.itemId),
    itemName: findHeaderIndex_(headers, CONFIG.ALIASES.itemName),
    room: findHeaderIndex_(headers, CONFIG.ALIASES.room),
    location: findHeaderIndex_(headers, CONFIG.ALIASES.location),
    qty: findHeaderIndex_(headers, CONFIG.ALIASES.qty),
    category: findHeaderIndex_(headers, CONFIG.ALIASES.category),
    status: findHeaderIndex_(headers, CONFIG.ALIASES.status),
    qrLink: findHeaderIndex_(headers, CONFIG.ALIASES.qrLink),
    qrImage: findHeaderIndex_(headers, CONFIG.ALIASES.qrImage)
  };

  const required = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'];
  if (opts.requireQrLink) required.push('qrLink');
  if (opts.requireQrImage) required.push('qrImage');

  const missing = required.filter(function (k) { return map[k] === -1; });
  if (missing.length) {
    const label = missing.map(function (k) {
      return CONFIG.ALIASES[k][0];
    }).join(', ');
    throw new Error('Required column missing: ' + label);
  }

  return map;
}

function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function findHeaderIndex_(normalizedHeaders, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalizedHeaders.indexOf(normalizeHeader_(aliases[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function columnLetter_(indexOneBased) {
  let n = indexOneBased;
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}

function getWebAppBaseUrl_() {
  const propertyUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.WEB_APP_URL_PROPERTY);
  if (propertyUrl && propertyUrl.trim()) return propertyUrl.trim();

  const deployedUrl = ScriptApp.getService().getUrl();
  if (deployedUrl) return deployedUrl;

  throw new Error(
    'Web App URL not configured. Set Script Property ' + CONFIG.WEB_APP_URL_PROPERTY + ' or deploy the app first.'
  );
}

function toNonNegativeNumber_(value) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? num : 0;
}

function normalizeStatus_(value) {
  const status = String(value || '').trim();
  return status || 'Good';
}

function isHazardCategory_(category) {
  return CONFIG.HAZARD_CATEGORIES.indexOf(String(category || '').trim().toLowerCase()) !== -1;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function statusClassServer_(status) {
  switch (String(status || '')) {
    case 'Good': return 'text-emerald-700';
    case 'Low Stock': return 'text-amber-700';
    case 'Missing': return 'text-red-700';
    case 'Needs Maintenance': return 'text-orange-700';
    default: return 'text-slate-700';
  }
}

function renderInitialItemsHtml_(bootstrap) {
  if (bootstrap.pageType === 'landing') {
    return renderLandingHtml_(bootstrap.locations || []);
  }
  return renderInventoryHtml_(bootstrap.rows || [], bootstrap.mode);
}

function renderLandingHtml_(locations) {
  let html = '';
  html += '<div class="p-4 border-b border-slate-100 bg-blue-50">';
  html += '<h3 class="font-semibold text-slate-900">Choose a location</h3>';
  html += '<p class="text-sm text-slate-700 mt-1">Scan a QR code or select a room/location below.</p>';
  html += '</div>';

  if (!locations.length) {
    html += '<p class="p-4 text-sm text-slate-500">No locations found in the inventory sheet.</p>';
    return html;
  }

  let currentRoom = '';
  locations.forEach(function (entry) {
    if (entry.room !== currentRoom) {
      currentRoom = entry.room;
      html += '<div class="px-4 py-2 bg-slate-50 border-y border-slate-100 text-xs font-semibold text-slate-600 tracking-wide">ROOM ' + escapeHtml_(currentRoom) + '</div>';
    }

    const viewUrl = '?room=' + encodeURIComponent(entry.room) + '&loc=' + encodeURIComponent(entry.loc);
    const techUrl = viewUrl + '&mode=tech';
    html += '<article class="p-4 border-b border-slate-100 last:border-b-0">';
    html += '<p class="font-medium text-slate-900">' + escapeHtml_(entry.loc) + '</p>';
    html += '<p class="text-xs text-slate-500 mt-1">Room: ' + escapeHtml_(entry.room) + '</p>';
    html += '<div class="mt-3 flex gap-2">';
    html += '<a class="text-xs px-3 py-2 rounded-md bg-slate-200 text-slate-800" href="' + escapeHtml_(viewUrl) + '">Open View</a>';
    html += '<a class="text-xs px-3 py-2 rounded-md bg-indigo-600 text-white" href="' + escapeHtml_(techUrl) + '">Open Tech</a>';
    html += '</div></article>';
  });

  return html;
}

function renderInventoryHtml_(rows, mode) {
  if (!rows.length) {
    return '<p class="p-4 text-sm text-slate-500">No inventory records found for this location.</p>';
  }

  let html = '';
  rows.forEach(function (row) {
    const hazard = row.isHazard ? '<span class="inline-block mt-2 text-[10px] font-bold tracking-wide text-red-800 bg-red-200 rounded px-2 py-1">HAZARD - CHEMICAL</span>' : '';
    const statusClass = statusClassServer_(row.status);

    if (mode === 'tech') {
      const opts = CONFIG.STATUS_OPTIONS.map(function (s) {
        return '<option value="' + escapeHtml_(s) + '" ' + (s === row.status ? 'selected' : '') + '>' + escapeHtml_(s) + '</option>';
      }).join('');
      html += '<article class="p-4 border-b border-slate-100 last:border-b-0 ' + (row.isHazard ? 'bg-red-50' : 'bg-white') + '">';
      html += '<div class="grid grid-cols-1 sm:grid-cols-[1fr_auto_auto] gap-3 items-end">';
      html += '<div><h3 class="font-semibold text-slate-900">' + escapeHtml_(row.itemName) + '</h3><p class="text-xs text-slate-500 mt-1">ID: ' + escapeHtml_(row.itemId) + ' · Category: ' + escapeHtml_(row.category) + '</p>' + hazard + '</div>';
      html += '<label class="block"><span class="text-xs text-slate-500">Qty</span><input type="number" min="0" class="qty-input mt-1 w-24 border border-slate-300 rounded px-2 py-1" data-row="' + escapeHtml_(String(row.sheetRow)) + '" value="' + escapeHtml_(String(row.qty)) + '" /></label>';
      html += '<label class="block"><span class="text-xs text-slate-500">Status</span><select class="status-input mt-1 border border-slate-300 rounded px-2 py-1 ' + statusClass + '" data-row="' + escapeHtml_(String(row.sheetRow)) + '">' + opts + '</select></label>';
      html += '</div></article>';
    } else {
      html += '<article class="p-4 border-b border-slate-100 last:border-b-0 ' + (row.isHazard ? 'bg-red-50' : 'bg-white') + '">';
      html += '<div class="flex items-start justify-between gap-3"><div><h3 class="font-semibold text-slate-900">' + escapeHtml_(row.itemName) + '</h3><p class="text-xs text-slate-500 mt-1">ID: ' + escapeHtml_(row.itemId) + ' · Category: ' + escapeHtml_(row.category) + '</p>' + hazard + '</div>';
      html += '<div class="text-right"><p class="text-sm text-slate-500">Expected Qty</p><p class="text-2xl font-bold text-slate-900">' + escapeHtml_(String(row.qty)) + '</p><p class="text-xs font-medium mt-1 ' + statusClass + '">' + escapeHtml_(row.status) + '</p></div></div>';
      html += '</article>';
    }
  });
  return html;
}

function buildPageHtml_(params, bootstrap) {
  const room = bootstrap.room || params.room || '';
  const loc = bootstrap.loc || params.loc || '';
  const mode = bootstrap.mode === 'tech' ? 'tech' : 'view';
  const initialHtml = renderInitialItemsHtml_(bootstrap);

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>D&T Inventory</title>
  <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-100 min-h-screen">
  <main class="max-w-3xl mx-auto p-4 sm:p-6">
    <header class="mb-4">
      <h1 class="text-2xl font-bold text-slate-900">D&T Inventory</h1>
      <p class="text-sm text-slate-600 mt-1">
        Room: <span id="roomLabel" class="font-semibold">${escapeHtml_(room) || '-'}</span>
        · Location: <span id="locLabel" class="font-semibold">${escapeHtml_(loc) || '-'}</span>
      </p>
      <div class="mt-3 flex flex-wrap gap-2">
        <a id="viewModeLink" class="text-xs px-3 py-2 rounded-lg bg-slate-200 text-slate-800" href="#">View Mode</a>
        <a id="techModeLink" class="text-xs px-3 py-2 rounded-lg bg-indigo-600 text-white" href="#">Technician Access</a>
      </div>
    </header>

    <section id="notice" class="hidden mb-4 p-3 rounded-lg text-sm"></section>
    <section id="bridgeWarning" class="hidden mb-4 p-3 rounded-lg text-sm bg-amber-100 text-amber-800"></section>

    ${CONFIG.DEBUG_PANEL ? '<section id="debugPanel" class="mb-4 p-3 rounded-lg text-xs bg-slate-900 text-slate-100"></section>' : ''}

    <section class="bg-white shadow-sm border border-slate-200 rounded-xl overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-200 flex items-center justify-between">
        <h2 class="font-semibold text-slate-800">Inventory Items</h2>
        <span id="modeBadge" class="text-xs px-2 py-1 rounded-full font-medium"></span>
      </div>
      <div id="loading" class="hidden p-4 text-slate-500 text-sm">Loading inventory...</div>
      <div id="items">${initialHtml}</div>
    </section>

    <footer class="mt-4 flex justify-end">
      <button id="saveBtn" class="hidden px-4 py-2 rounded-lg bg-emerald-600 text-white font-medium disabled:opacity-50" type="button">Save Updates</button>
    </footer>
  </main>

  <script>
    const APP = {
      room: ${JSON.stringify(room)},
      loc: ${JSON.stringify(loc)},
      mode: ${JSON.stringify(mode)},
      bootstrap: ${JSON.stringify(bootstrap)}
    };

    document.addEventListener('DOMContentLoaded', init);

    function init() {
      buildModeLinks();
      renderModeBadge();
      setBridgeWarning();
      if (${CONFIG.DEBUG_PANEL ? 'true' : 'false'}) renderDebug();

      const saveBtn = document.getElementById('saveBtn');
      if (saveBtn) saveBtn.addEventListener('click', saveUpdates);
    }

    function buildModeLinks() {
      const view = document.getElementById('viewModeLink');
      const tech = document.getElementById('techModeLink');
      if (!APP.room || !APP.loc) {
        view.classList.add('opacity-50', 'pointer-events-none');
        tech.classList.add('opacity-50', 'pointer-events-none');
        view.setAttribute('title', 'Choose a location below first');
        tech.setAttribute('title', 'Choose a location below first');
        view.href = '#';
        tech.href = '#';
        return;
      }
      const base = window.location.pathname + '?room=' + encodeURIComponent(APP.room) + '&loc=' + encodeURIComponent(APP.loc);
      view.href = base;
      tech.href = base + '&mode=tech';
    }

    function renderModeBadge() {
      const badge = document.getElementById('modeBadge');
      const saveBtn = document.getElementById('saveBtn');
      if (!APP.room || !APP.loc) {
        badge.textContent = 'Landing';
        badge.className = 'text-xs px-2 py-1 rounded-full font-medium bg-blue-100 text-blue-700';
        saveBtn.classList.add('hidden');
        return;
      }
      if (APP.mode === 'tech') {
        badge.textContent = 'Technician Mode';
        badge.className = 'text-xs px-2 py-1 rounded-full font-medium bg-indigo-100 text-indigo-700';
        saveBtn.classList.remove('hidden');
      } else {
        badge.textContent = 'View Mode';
        badge.className = 'text-xs px-2 py-1 rounded-full font-medium bg-slate-100 text-slate-700';
        saveBtn.classList.add('hidden');
      }
    }

    function hasBridge() {
      return !!(window.google && google.script && google.script.run);
    }

    function setBridgeWarning() {
      if (hasBridge()) return;
      const warn = document.getElementById('bridgeWarning');
      warn.textContent = 'Interactive save is unavailable in this preview. Open the deployed /exec web app URL to use full functionality.';
      warn.classList.remove('hidden');
    }

    function renderDebug() {
      const el = document.getElementById('debugPanel');
      if (!el) return;
      const rowCount = APP.bootstrap && APP.bootstrap.rows ? APP.bootstrap.rows.length : 0;
      const locCount = APP.bootstrap && APP.bootstrap.locations ? APP.bootstrap.locations.length : 0;
      el.innerHTML = 'mode=' + esc(APP.mode) + ' | room=' + esc(APP.room || '-') + ' | loc=' + esc(APP.loc || '-') + ' | bridge=' + (hasBridge() ? 'yes' : 'no') + ' | rows=' + rowCount + ' | locations=' + locCount;
    }

    function saveUpdates() {
      if (!hasBridge()) {
        showNotice('Cannot save in this context. Open deployed /exec URL.', true);
        return;
      }
      const saveBtn = document.getElementById('saveBtn');
      saveBtn.disabled = true;

      const byRow = {};
      let invalidQty = false;
      document.querySelectorAll('.qty-input').forEach(function (input) {
        const row = input.getAttribute('data-row');
        const qty = Number(input.value);
        if (!Number.isFinite(qty) || qty < 0) invalidQty = true;
        byRow[row] = byRow[row] || { sheetRow: Number(row) };
        byRow[row].qty = qty;
      });

      if (invalidQty) {
        saveBtn.disabled = false;
        showNotice('Please enter valid non-negative quantities before saving.', true);
        return;
      }

      document.querySelectorAll('.status-input').forEach(function (select) {
        const row = select.getAttribute('data-row');
        byRow[row] = byRow[row] || { sheetRow: Number(row) };
        byRow[row].status = select.value;
      });

      google.script.run
        .withSuccessHandler(function (res) {
          saveBtn.disabled = false;
          showNotice('Saved ' + res.updatedCount + ' item(s) successfully.', false);
          window.location.reload();
        })
        .withFailureHandler(function (err) {
          saveBtn.disabled = false;
          showNotice((err && err.message) || 'Save failed. Please try again.', true);
        })
        .saveInventoryUpdates({ updates: Object.keys(byRow).map(function (k) { return byRow[k]; }) });
    }

    function showNotice(message, isError) {
      if (!message) return;
      const el = document.getElementById('notice');
      el.textContent = message;
      el.classList.remove('hidden', 'bg-red-100', 'text-red-700', 'bg-blue-100', 'text-blue-700');
      if (isError) el.classList.add('bg-red-100', 'text-red-700');
      else el.classList.add('bg-blue-100', 'text-blue-700');
    }

    function esc(v) {
      return String(v == null ? '' : v)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }
  </script>
</body>
</html>`;
}
