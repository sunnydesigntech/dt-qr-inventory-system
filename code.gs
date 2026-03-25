/**
 * D&T QR Inventory System - Single-file Google Apps Script implementation.
 */
const CONFIG = {
  SHEET_NAME: 'Inventory',
  HEADER_ROW: 1,
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_CATEGORIES: ['chemicals', 'chemical'],
  ALIASES: {
    itemId: ['item id', 'itemid', 'id'],
    itemName: ['item name', 'name'],
    room: ['room'],
    location: ['specific location', 'location', 'specificlocation'],
    qty: ['qty', 'quantity'],
    category: ['category'],
    status: ['status'],
    qrLink: ['qr code link (auto-generated)', 'auto-generated qr link', 'qr link', 'qr url']
  }
};

function doGet(e) {
  const params = {
    room: ((e && e.parameter && e.parameter.room) || '').trim(),
    loc: ((e && e.parameter && e.parameter.loc) || '').trim(),
    mode: (((e && e.parameter && e.parameter.mode) || 'view').toLowerCase() === 'tech') ? 'tech' : 'view'
  };

  return HtmlService
    .createHtmlOutput(buildPageHtml_(params))
    .setTitle('D&T Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return {
      success: true,
      rows: [],
      room: room,
      loc: loc,
      message: 'No inventory data found in this sheet.'
    };
  }

  const map = getColumnMap_(values[0], { requireQrLink: false });
  const roomNeedle = room.toLowerCase();
  const locNeedle = loc.toLowerCase();
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

  return {
    success: true,
    rows: rows,
    room: room,
    loc: loc,
    message: rows.length ? '' : 'No inventory records found for this location.'
  };
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
    qrLink: findHeaderIndex_(headers, CONFIG.ALIASES.qrLink)
  };

  const required = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'];
  if (opts.requireQrLink) required.push('qrLink');

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

function buildPageHtml_(params) {
  const room = params.room || '';
  const loc = params.loc || '';
  const mode = params.mode === 'tech' ? 'tech' : 'view';

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

    <section class="bg-white shadow-sm border border-slate-200 rounded-xl overflow-hidden">
      <div class="px-4 py-3 border-b border-slate-200 flex items-center justify-between">
        <h2 class="font-semibold text-slate-800">Inventory Items</h2>
        <span id="modeBadge" class="text-xs px-2 py-1 rounded-full font-medium"></span>
      </div>
      <div id="loading" class="p-4 text-slate-500 text-sm">Loading inventory...</div>
      <div id="items"></div>
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
      rows: []
    };

    const STATUS_CLASS = {
      'Good': 'text-emerald-700',
      'Low Stock': 'text-amber-700',
      'Missing': 'text-red-700',
      'Needs Maintenance': 'text-orange-700'
    };

    document.addEventListener('DOMContentLoaded', init);

    function init() {
      buildModeLinks();
      renderModeBadge();
      document.getElementById('saveBtn').addEventListener('click', saveUpdates);

      if (!APP.room || !APP.loc) {
        hideLoading();
        showNotice('Welcome. Scan a location QR code to load inventory for a room and storage location.', false);
        return;
      }

      loadInventory();
    }

    function buildModeLinks() {
      const qs = new URLSearchParams(window.location.search);
      const room = qs.get('room') || APP.room || '';
      const loc = qs.get('loc') || APP.loc || '';
      const base = window.location.pathname + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
      document.getElementById('viewModeLink').href = base;
      document.getElementById('techModeLink').href = base + '&mode=tech';
    }

    function renderModeBadge() {
      const badge = document.getElementById('modeBadge');
      const saveBtn = document.getElementById('saveBtn');
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

    function loadInventory() {
      google.script.run
        .withSuccessHandler(renderRows)
        .withFailureHandler(function (err) {
          hideLoading();
          const msg = (err && err.message) || 'This inventory sheet is missing required columns. Please contact the technician.';
          showNotice(msg, true);
        })
        .getInventoryData({ room: APP.room, loc: APP.loc });
    }

    function renderRows(res) {
      hideLoading();
      const items = document.getElementById('items');
      items.innerHTML = '';

      APP.rows = (res && res.rows) ? res.rows : [];
      if (res && res.room) document.getElementById('roomLabel').textContent = res.room;
      if (res && res.loc) document.getElementById('locLabel').textContent = res.loc;

      if (res && res.message) {
        showNotice(res.message, false);
      }

      if (!APP.rows.length) {
        items.innerHTML = '<p class="p-4 text-sm text-slate-500">No inventory records found for this location.</p>';
        return;
      }

      APP.rows.forEach(function (row) {
        const statusClass = STATUS_CLASS[row.status] || 'text-slate-700';
        const card = document.createElement('article');
        card.className = 'p-4 border-b border-slate-100 last:border-b-0 ' + (row.isHazard ? 'bg-red-50' : 'bg-white');
        card.innerHTML = APP.mode === 'tech' ? techCardHtml(row, statusClass) : viewCardHtml(row, statusClass);
        items.appendChild(card);
      });
    }

    function viewCardHtml(row, statusClass) {
      const hazard = row.isHazard
        ? '<span class="inline-block mt-2 text-[10px] font-bold tracking-wide text-red-800 bg-red-200 rounded px-2 py-1">HAZARD - CHEMICAL</span>'
        : '';

      return '' +
        '<div class="flex items-start justify-between gap-3">' +
          '<div>' +
            '<h3 class="font-semibold text-slate-900">' + esc(row.itemName) + '</h3>' +
            '<p class="text-xs text-slate-500 mt-1">ID: ' + esc(row.itemId) + ' · Category: ' + esc(row.category) + '</p>' +
            hazard +
          '</div>' +
          '<div class="text-right">' +
            '<p class="text-sm text-slate-500">Expected Qty</p>' +
            '<p class="text-2xl font-bold text-slate-900">' + esc(String(row.qty)) + '</p>' +
            '<p class="text-xs font-medium mt-1 ' + statusClass + '">' + esc(row.status) + '</p>' +
          '</div>' +
        '</div>';
    }

    function techCardHtml(row, statusClass) {
      const hazard = row.isHazard
        ? '<span class="inline-block mt-2 text-[10px] font-bold tracking-wide text-red-800 bg-red-200 rounded px-2 py-1">HAZARD - CHEMICAL</span>'
        : '';

      const options = ['Good', 'Low Stock', 'Missing', 'Needs Maintenance']
        .map(function (s) {
          return '<option value="' + esc(s) + '" ' + (s === row.status ? 'selected' : '') + '>' + esc(s) + '</option>';
        })
        .join('');

      return '' +
        '<div class="grid grid-cols-1 sm:grid-cols-[1fr_auto_auto] gap-3 items-end">' +
          '<div>' +
            '<h3 class="font-semibold text-slate-900">' + esc(row.itemName) + '</h3>' +
            '<p class="text-xs text-slate-500 mt-1">ID: ' + esc(row.itemId) + ' · Category: ' + esc(row.category) + '</p>' +
            hazard +
          '</div>' +
          '<label class="block">' +
            '<span class="text-xs text-slate-500">Qty</span>' +
            '<input type="number" min="0" class="qty-input mt-1 w-24 border border-slate-300 rounded px-2 py-1" data-row="' + esc(String(row.sheetRow)) + '" value="' + esc(String(row.qty)) + '" />' +
          '</label>' +
          '<label class="block">' +
            '<span class="text-xs text-slate-500">Status</span>' +
            '<select class="status-input mt-1 border border-slate-300 rounded px-2 py-1 ' + statusClass + '" data-row="' + esc(String(row.sheetRow)) + '">' + options + '</select>' +
          '</label>' +
        '</div>';
    }

    function saveUpdates() {
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

      const updates = Object.keys(byRow).map(function (k) { return byRow[k]; });

      google.script.run
        .withSuccessHandler(function (res) {
          saveBtn.disabled = false;
          showNotice('Saved ' + res.updatedCount + ' item(s) successfully.', false);
          loadInventory();
        })
        .withFailureHandler(function (err) {
          saveBtn.disabled = false;
          showNotice((err && err.message) || 'Save failed. Please try again.', true);
        })
        .saveInventoryUpdates({ updates: updates });
    }

    function hideLoading() {
      document.getElementById('loading').classList.add('hidden');
    }

    function showNotice(message, isError) {
      if (!message) return;
      const el = document.getElementById('notice');
      el.textContent = message;
      el.classList.remove('hidden', 'bg-red-100', 'text-red-700', 'bg-blue-100', 'text-blue-700');
      if (isError) {
        el.classList.add('bg-red-100', 'text-red-700');
      } else {
        el.classList.add('bg-blue-100', 'text-blue-700');
      }
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
