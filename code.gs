/**
 * D&T QR Inventory System (V1)
 * Single-file Google Apps Script implementation.
 */

const CONFIG = {
  SHEET_NAME: 'Inventory',
  HEADER_ROW: 1,
  WEB_APP_URL_PROPERTY: 'WEB_APP_BASE_URL',
  STATUS_OPTIONS: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  HAZARD_KEYWORDS: ['chemical', 'chemicals', 'hazardous', 'flammable', 'solvent', 'acetone', 'spray paint'],
  COLUMN_ALIASES: {
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
    mode: (((e && e.parameter && e.parameter.mode) || 'view').trim().toLowerCase() === 'tech') ? 'tech' : 'view'
  };

  return HtmlService.createHtmlOutput(buildPageHtml_(params))
    .setTitle('D&T Inventory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInventoryData(filters) {
  const roomFilter = ((filters && filters.room) || '').toString().trim().toLowerCase();
  const locFilter = ((filters && filters.loc) || '').toString().trim().toLowerCase();

  if (!roomFilter || !locFilter) {
    return {
      success: true,
      rows: [],
      message: 'Scan a valid QR code with both room and location parameters.'
    };
  }

  const sheet = getInventorySheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= CONFIG.HEADER_ROW) {
    return { success: true, rows: [], message: 'No inventory rows found in sheet.' };
  }

  const cols = getColumnMap_(values[0], { requireQrLink: false, requireQrImage: false });
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const roomVal = String(row[cols.room] || '').trim();
    const locVal = String(row[cols.location] || '').trim();
    if (!roomVal || !locVal) continue;

    const roomMatches = roomVal.toLowerCase() === roomFilter;
    const locMatches = locVal.toLowerCase() === locFilter;
    if (!roomMatches || !locMatches) continue;

    const categoryVal = String(row[cols.category] || '').trim();
    rows.push({
      sheetRow: i + 1,
      itemId: String(row[cols.itemId] || '').trim(),
      itemName: String(row[cols.itemName] || '').trim(),
      room: roomVal,
      specificLocation: locVal,
      qty: normalizeQty_(row[cols.qty]),
      category: categoryVal,
      status: normalizeStatus_(row[cols.status]),
      isHazard: isHazardousCategory_(categoryVal)
    });
  }

  return {
    success: true,
    rows: rows,
    room: roomFilter,
    loc: locFilter,
    message: rows.length ? '' : 'No inventory records found for this location.'
  };
}

function saveInventoryUpdates(payload) {
  if (!payload || !Array.isArray(payload.updates) || !payload.updates.length) {
    throw new Error('No updates provided.');
  }

  const sheet = getInventorySheet_();
  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = getColumnMap_(header, { requireQrLink: false, requireQrImage: false });

  const qtyCol = cols.qty + 1;
  const statusCol = cols.status + 1;
  const timestampCol = cols.lastUpdated !== undefined ? cols.lastUpdated + 1 : null;
  const maxRow = sheet.getLastRow();

  payload.updates.forEach(function (update) {
    const rowNum = Number(update.sheetRow);
    if (!rowNum || rowNum <= CONFIG.HEADER_ROW || rowNum > maxRow) {
      throw new Error('Invalid row number in update payload: ' + update.sheetRow);
    }

    const qty = normalizeQty_(update.qty);
    const status = normalizeStatus_(update.status);
    if (CONFIG.STATUS_OPTIONS.indexOf(status) === -1) {
      throw new Error('Invalid status value for row ' + rowNum + ': ' + status);
    }

    sheet.getRange(rowNum, qtyCol).setValue(qty);
    sheet.getRange(rowNum, statusCol).setValue(status);
    if (timestampCol) {
      sheet.getRange(rowNum, timestampCol).setValue(new Date());
    }
  });

  return { success: true, updatedCount: payload.updates.length, timestamp: new Date().toISOString() };
}

function refreshQrLinks() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = getColumnMap_(header, { requireQrLink: true, requireQrImage: false });
  const baseUrl = getWebAppBaseUrl_();

  const data = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, sheet.getLastColumn()).getValues();
  const output = [];

  data.forEach(function (row) {
    const roomVal = String(row[cols.room] || '').trim();
    const locVal = String(row[cols.location] || '').trim();
    if (!roomVal || !locVal) {
      output.push(['']);
      return;
    }
    const url = baseUrl + '?room=' + encodeURIComponent(roomVal) + '&loc=' + encodeURIComponent(locVal);
    output.push([url]);
  });

  sheet.getRange(CONFIG.HEADER_ROW + 1, cols.qrLink + 1, output.length, 1).setValues(output);
}

function refreshQrImages() {
  const sheet = getInventorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  const header = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const cols = getColumnMap_(header, { requireQrLink: true, requireQrImage: true });

  const linkColA1 = columnToLetter_(cols.qrLink + 1);
  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const formula = '=IF(' + linkColA1 + row + '="","",IMAGE("https://quickchart.io/qr?size=220&text="&ENCODEURL(' + linkColA1 + row + ')))';
    sheet.getRange(row, cols.qrImage + 1).setFormula(formula);
  }
}

function getInventorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const named = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (named) return named;

  const first = ss.getSheets()[0];
  if (!first) throw new Error('No sheets found in this spreadsheet.');
  return first;
}

function getColumnMap_(headerRow, options) {
  const opts = options || {};
  const normalizedHeaders = headerRow.map(normalizeHeader_);

  const map = {
    itemId: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.itemId),
    itemName: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.itemName),
    room: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.room),
    location: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.location),
    qty: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.qty),
    category: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.category),
    status: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.status),
    qrLink: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.qrLink),
    qrImage: findHeaderIndex_(normalizedHeaders, CONFIG.COLUMN_ALIASES.qrImage),
    lastUpdated: findHeaderIndex_(normalizedHeaders, ['last updated', 'updated at', 'timestamp'])
  };

  const requiredKeys = ['itemId', 'itemName', 'room', 'location', 'qty', 'category', 'status'];
  if (opts.requireQrLink) requiredKeys.push('qrLink');
  if (opts.requireQrImage) requiredKeys.push('qrImage');

  const missing = requiredKeys.filter(function (k) { return map[k] === -1; });
  if (missing.length) {
    const missingLabel = missing.map(function (k) {
      return (CONFIG.COLUMN_ALIASES[k] || [k])[0];
    });
    throw new Error('Required column missing: ' + missingLabel.join(', '));
  }

  Object.keys(map).forEach(function (k) {
    if (map[k] === -1) delete map[k];
  });

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
  const prop = PropertiesService.getScriptProperties().getProperty(CONFIG.WEB_APP_URL_PROPERTY);
  if (prop) return prop;

  const deployedUrl = ScriptApp.getService().getUrl();
  if (deployedUrl) return deployedUrl;

  throw new Error('Web app URL not configured. Set script property ' + CONFIG.WEB_APP_URL_PROPERTY + ' or deploy as web app.');
}

function normalizeQty_(value) {
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : 0;
}

function normalizeStatus_(status) {
  const normalized = String(status || '').trim();
  return normalized || 'Good';
}

function isHazardousCategory_(category) {
  const c = String(category || '').toLowerCase();
  return CONFIG.HAZARD_KEYWORDS.some(function (k) {
    return c.indexOf(k) !== -1;
  });
}

function columnToLetter_(column) {
  let temp;
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function buildPageHtml_(params) {
  const safeRoom = escapeHtml_(params.room || '');
  const safeLoc = escapeHtml_(params.loc || '');
  const mode = params.mode === 'tech' ? 'tech' : 'view';

  return '<!doctype html>' +
    '<html><head>' +
    '<meta charset="utf-8" />' +
    '<meta name="viewport" content="width=device-width,initial-scale=1" />' +
    '<title>D&T Inventory</title>' +
    '<script src="https://cdn.tailwindcss.com"></script>' +
    '</head>' +
    '<body class="bg-slate-100 min-h-screen">' +
      '<main class="max-w-3xl mx-auto p-4 sm:p-6">' +
        '<header class="mb-4">' +
          '<h1 class="text-2xl font-bold text-slate-900">D&T Inventory</h1>' +
          '<p class="text-sm text-slate-600 mt-1">Room: <span class="font-semibold">' + safeRoom + '</span> · Location: <span class="font-semibold">' + safeLoc + '</span></p>' +
          '<div class="mt-3 flex flex-wrap gap-2">' +
            '<a id="viewModeLink" class="text-xs px-3 py-2 rounded-lg bg-slate-200 text-slate-800" href="#">View Mode</a>' +
            '<a id="techModeLink" class="text-xs px-3 py-2 rounded-lg bg-indigo-600 text-white" href="#">Technician Access</a>' +
          '</div>' +
        '</header>' +

        '<section id="notice" class="hidden mb-4 p-3 rounded-lg text-sm"></section>' +

        '<section class="bg-white shadow-sm border border-slate-200 rounded-xl overflow-hidden">' +
          '<div class="px-4 py-3 border-b border-slate-200 flex items-center justify-between">' +
            '<h2 class="font-semibold text-slate-800">Inventory Items</h2>' +
            '<span id="modeBadge" class="text-xs px-2 py-1 rounded-full font-medium"></span>' +
          '</div>' +
          '<div id="loading" class="p-4 text-slate-500 text-sm">Loading inventory...</div>' +
          '<div id="items"></div>' +
        '</section>' +

        '<footer class="mt-4 flex justify-end">' +
          '<button id="saveBtn" class="hidden px-4 py-2 rounded-lg bg-emerald-600 text-white font-medium disabled:opacity-50" type="button">Save Updates</button>' +
        '</footer>' +
      '</main>' +

      '<script>' +
        'const APP = {' +
          'room: ' + JSON.stringify(params.room || '') + ',' +
          'loc: ' + JSON.stringify(params.loc || '') + ',' +
          'mode: ' + JSON.stringify(mode) + ',' +
          'rows: []' +
        '};' +

        'const STATUS_CLASS = {' +
          '"Good": "text-emerald-700",' +
          '"Low Stock": "text-amber-700",' +
          '"Missing": "text-rose-700",' +
          '"Needs Maintenance": "text-red-700"' +
        '};' +

        'document.addEventListener("DOMContentLoaded", function(){ bindModeLinks(); setModeUI(); validateParamsAndFetch(); document.getElementById("saveBtn").addEventListener("click", saveUpdates); });' +

        'function bindModeLinks(){' +
          'const qs = new URLSearchParams(window.location.search);' +
          'const room = qs.get("room") || APP.room || "";' +
          'const loc = qs.get("loc") || APP.loc || "";' +
          'const base = window.location.origin + window.location.pathname + "?room=" + encodeURIComponent(room) + "&loc=" + encodeURIComponent(loc);' +
          'document.getElementById("viewModeLink").href = base;' +
          'document.getElementById("techModeLink").href = base + "&mode=tech";' +
        '}' +

        'function setModeUI(){' +
          'const badge = document.getElementById("modeBadge");' +
          'const saveBtn = document.getElementById("saveBtn");' +
          'if(APP.mode === "tech"){ badge.textContent = "Technician Mode"; badge.className = "text-xs px-2 py-1 rounded-full font-medium bg-indigo-100 text-indigo-700"; saveBtn.classList.remove("hidden"); }' +
          'else { badge.textContent = "View Mode"; badge.className = "text-xs px-2 py-1 rounded-full font-medium bg-slate-100 text-slate-700"; saveBtn.classList.add("hidden"); }' +
        '}' +

        'function validateParamsAndFetch(){ if(!APP.room || !APP.loc){ document.getElementById("loading").classList.add("hidden"); showNotice("Missing room/location in URL. Please scan a valid location QR code.", true); return; } fetchInventory(); }' +

        'function fetchInventory(){ google.script.run.withSuccessHandler(renderRows).withFailureHandler(function(err){ showNotice((err && err.message) || "This inventory sheet is missing required columns. Please contact the technician.", true); }).getInventoryData({ room: APP.room, loc: APP.loc }); }' +

        'function renderRows(res){ document.getElementById("loading").classList.add("hidden"); const items = document.getElementById("items"); items.innerHTML = ""; APP.rows = (res && res.rows) ? res.rows : []; if(res && res.message){ showNotice(res.message, false); } if(!APP.rows.length){ items.innerHTML = "<p class=\"p-4 text-sm text-slate-500\">No inventory records found for this location.</p>"; return; } APP.rows.forEach(function(row){ const wrap = document.createElement("article"); wrap.className = "p-4 border-b border-slate-100 last:border-b-0 " + (row.isHazard ? "bg-red-50" : "bg-white"); const statusClass = STATUS_CLASS[row.status] || "text-slate-700"; wrap.innerHTML = APP.mode === "tech" ? techRowHtml(row, statusClass) : viewRowHtml(row, statusClass); items.appendChild(wrap); }); }' +

        'function viewRowHtml(row, statusClass){ return "<div class=\"flex items-start justify-between gap-3\"><div><h3 class=\"font-semibold text-slate-900\">" + esc(row.itemName) + "</h3><p class=\"text-xs text-slate-500 mt-1\">ID: " + esc(row.itemId) + " · Category: " + esc(row.category) + "</p>" + (row.isHazard ? "<span class=\"inline-block mt-2 text-[10px] font-bold tracking-wide text-red-800 bg-red-200 rounded px-2 py-1\">HAZARD</span>" : "") + "</div><div class=\"text-right\"><p class=\"text-sm text-slate-500\">Expected Qty</p><p class=\"text-2xl font-bold text-slate-900\">" + esc(String(row.qty)) + "</p><p class=\"text-xs font-medium mt-1 " + statusClass + "\">" + esc(row.status) + "</p></div></div>"; }' +

        'function techRowHtml(row, statusClass){ const options = ["Good","Low Stock","Missing","Needs Maintenance"].map(function(s){ return "<option value=\"" + esc(s) + "\" " + (s === row.status ? "selected" : "") + ">" + esc(s) + "</option>"; }).join(""); return "<div class=\"grid grid-cols-1 sm:grid-cols-[1fr_auto_auto] gap-3 items-end\"><div><h3 class=\"font-semibold text-slate-900\">" + esc(row.itemName) + "</h3><p class=\"text-xs text-slate-500 mt-1\">ID: " + esc(row.itemId) + " · Category: " + esc(row.category) + "</p>" + (row.isHazard ? "<span class=\"inline-block mt-2 text-[10px] font-bold tracking-wide text-red-800 bg-red-200 rounded px-2 py-1\">HAZARD</span>" : "") + "</div><label class=\"block\"><span class=\"text-xs text-slate-500\">Qty</span><input type=\"number\" min=\"0\" class=\"qty-input mt-1 w-24 border border-slate-300 rounded px-2 py-1\" data-row=\"" + esc(String(row.sheetRow)) + "\" value=\"" + esc(String(row.qty)) + "\" /></label><label class=\"block\"><span class=\"text-xs text-slate-500\">Status</span><select class=\"status-input mt-1 border border-slate-300 rounded px-2 py-1 " + statusClass + "\" data-row=\"" + esc(String(row.sheetRow)) + "\">" + options + "</select></label></div>"; }' +

        'function saveUpdates(){ const saveBtn = document.getElementById("saveBtn"); saveBtn.disabled = true; const byRow = {}; let invalidQty = false; document.querySelectorAll(".qty-input").forEach(function(el){ const row = el.getAttribute("data-row"); const val = Number(el.value); if(!Number.isFinite(val) || val < 0){ invalidQty = true; } byRow[row] = byRow[row] || { sheetRow: Number(row) }; byRow[row].qty = val; }); if(invalidQty){ saveBtn.disabled = false; showNotice("Please enter valid non-negative quantities.", true); return; } document.querySelectorAll(".status-input").forEach(function(el){ const row = el.getAttribute("data-row"); byRow[row] = byRow[row] || { sheetRow: Number(row) }; byRow[row].status = el.value; }); const updates = Object.keys(byRow).map(function(k){ return byRow[k]; }); google.script.run.withSuccessHandler(function(res){ saveBtn.disabled = false; showNotice("Saved " + res.updatedCount + " item(s) successfully.", false); fetchInventory(); }).withFailureHandler(function(err){ saveBtn.disabled = false; showNotice((err && err.message) || "Save failed.", true); }).saveInventoryUpdates({ updates: updates }); }' +

        'function showNotice(msg, isError){ if(!msg) return; const el = document.getElementById("notice"); el.textContent = msg; el.classList.remove("hidden", "bg-red-100", "text-red-700", "bg-blue-100", "text-blue-700"); el.classList.add(isError ? "bg-red-100" : "bg-blue-100", isError ? "text-red-700" : "text-blue-700"); }' +

        'function esc(v){ return String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;").replace(/\'/g, "&#39;"); }' +
      '</script>' +
    '</body></html>';
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
