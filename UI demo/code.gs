/**
 * D&T QR Inventory · code.gs
 * Single-file Google Apps Script web app.
 *
 * This is the production-ready UI rewrite. Every screen — Landing, View Mode,
 * Tech Mode, Admin/Readiness, QR Labels, Diagnostics — is rendered from this
 * one file. No external HTML/CSS/JS dependencies beyond the Tailwind CDN and
 * Google Fonts. Server-rendered, mobile-first, desktop-complete.
 *
 * Routes:
 *   ?                                 → Landing dashboard
 *   ?room=419A&loc=CHM-CAB-01         → Storage page (View Mode)
 *   ?room=419A&loc=CHM-CAB-01&mode=tech → Storage page (Tech Mode)
 *   ?admin=readiness                  → Admin readiness dashboard
 *   ?admin=labels                     → QR label print preview
 *   ?admin=diagnostics                → Diagnostics + how-to-fix
 *
 * Important: in-app navigation uses RELATIVE URLs (?room=…&loc=…) so the
 * app works without WEB_APP_BASE_URL set. WEB_APP_BASE_URL is only required
 * for QR-link generation and external/printed absolute URLs.
 */

// ──────────────────────────────────────────────────────────────────────────
// CONFIG
// ──────────────────────────────────────────────────────────────────────────

var CONFIG = {
  SHEET_NAME_DEFAULT: 'Inventory',
  STATUS_VALUES: ['Good', 'Low Stock', 'Missing', 'Needs Maintenance'],
  ATTENTION_STATUSES: ['Low Stock', 'Missing', 'Needs Maintenance'],
  ROLLOUT_ROOM: '419A',
  REQUIRED_COLS: ['Storage ID', 'Room', 'Storage Label', 'Item ID', 'Item Name', 'Quantity', 'Status'],
  OPTIONAL_COLS: ['Hazard', 'Category', 'Unit', 'Remarks', 'Specific Location', 'Location Code', 'QR Code Image'],
};

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function getConfig_() {
  return {
    spreadsheetId: getProp_('SPREADSHEET_ID'),
    webAppBaseUrl: getProp_('WEB_APP_BASE_URL'),
    inventorySheetName: getProp_('INVENTORY_SHEET_NAME') || CONFIG.SHEET_NAME_DEFAULT,
  };
}

// ──────────────────────────────────────────────────────────────────────────
// ENTRY POINT
// ──────────────────────────────────────────────────────────────────────────

function doGet(e) {
  var params = getRequestParams_(e);
  var cfg = getConfig_();

  // Hard error: no spreadsheet configured
  if (!cfg.spreadsheetId) {
    return htmlOutput_(buildPageHtml_({
      title: 'Setup required',
      mode: 'error',
      body: renderErrorStateHtml_('no-spreadsheet'),
    }));
  }

  // Admin routes
  if (params.admin) {
    var data = buildBootstrapData_(cfg);
    if (params.admin === 'readiness') {
      return htmlOutput_(buildPageHtml_({
        title: 'Admin · Readiness',
        mode: 'admin',
        body: renderAdminReadinessHtml_(data),
      }));
    }
    if (params.admin === 'labels') {
      return htmlOutput_(buildPageHtml_({
        title: 'QR Labels',
        mode: 'labels',
        body: renderQrLabelsHtml_(data, cfg),
      }));
    }
    if (params.admin === 'diagnostics') {
      return htmlOutput_(buildPageHtml_({
        title: 'Diagnostics',
        mode: 'admin',
        body: renderDiagnosticsHtml_(data, cfg),
      }));
    }
  }

  // Storage page
  if (params.room && params.loc) {
    var bundle = getLocationBundle_(cfg, params.room, params.loc);
    if (!bundle) {
      return htmlOutput_(buildPageHtml_({
        title: 'Not found',
        mode: 'error',
        body: renderErrorStateHtml_('no-match', { room: params.room, loc: params.loc }),
      }));
    }
    return htmlOutput_(buildPageHtml_({
      title: bundle.location.label + ' · ' + params.room,
      mode: params.mode === 'tech' ? 'tech' : 'view',
      body: renderInventoryHtml_(bundle, params.mode === 'tech' ? 'tech' : 'view'),
    }));
  }

  // Default: landing
  var data = buildBootstrapData_(cfg);
  return htmlOutput_(buildPageHtml_({
    title: 'D&T QR Inventory',
    mode: 'landing',
    body: renderLandingHtml_(data, cfg),
  }));
}

function getRequestParams_(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  return {
    room: (p.room || '').trim(),
    loc: (p.loc || '').trim(),
    mode: (p.mode || '').trim().toLowerCase(),
    admin: (p.admin || '').trim().toLowerCase(),
  };
}

function htmlOutput_(html) {
  return HtmlService.createHtmlOutput(html)
    .setTitle('D&T QR Inventory')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ──────────────────────────────────────────────────────────────────────────
// DATA ACCESS
// ──────────────────────────────────────────────────────────────────────────

function getInventorySheet_(cfg) {
  return SpreadsheetApp.openById(cfg.spreadsheetId).getSheetByName(cfg.inventorySheetName);
}

function readInventoryRows_(cfg) {
  var sh = getInventorySheet_(cfg);
  if (!sh) return { headers: [], rows: [] };
  var values = sh.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = values.slice(1).map(function (r, i) {
    var o = { __row: i + 2 };
    headers.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
  return { headers: headers, rows: rows };
}

function rowToItem_(r) {
  var status = String(r['Status'] || '').trim() || 'Good';
  var qty = Number(r['Quantity']);
  if (!isFinite(qty)) qty = 0;
  return {
    sheetRow: r.__row,
    storageId: String(r['Storage ID'] || '').trim(),
    room: String(r['Room'] || '').trim(),
    storageLabel: String(r['Storage Label'] || '').trim(),
    specificLocation: String(r['Specific Location'] || '').trim(),
    locationCode: String(r['Location Code'] || '').trim(),
    itemId: String(r['Item ID'] || '').trim(),
    itemName: String(r['Item Name'] || '').trim(),
    qty: qty,
    unit: String(r['Unit'] || '').trim(),
    status: status,
    category: String(r['Category'] || '').trim(),
    remarks: String(r['Remarks'] || '').trim(),
    hazard: isTruthy_(r['Hazard']) || /chem|hazard/i.test(String(r['Category'] || '')),
    qrCodeImage: String(r['QR Code Image'] || '').trim(),
  };
}

function isTruthy_(v) {
  if (v === true) return true;
  if (typeof v === 'number') return v > 0;
  var s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1' || s === 'hazard' || s === 'chemical';
}

function isPlaceholderItem_(it) {
  // Storage master rows (no item id / name) should not count as items
  return !it.itemId && !it.itemName;
}

function getLocationBundle_(cfg, room, storageId) {
  var data = readInventoryRows_(cfg);
  var matches = data.rows.map(rowToItem_).filter(function (it) {
    return it.room === room && it.storageId === storageId;
  });
  if (!matches.length) return null;
  var first = matches[0];
  var realItems = matches.filter(function (it) { return !isPlaceholderItem_(it); });
  return {
    location: {
      id: first.storageId,
      room: first.room,
      label: first.storageLabel,
      specific: first.specificLocation,
      code: first.locationCode,
      hazard: realItems.some(function (it) { return it.hazard; }),
    },
    items: realItems,
    raw: matches,
  };
}

function buildBootstrapData_(cfg) {
  var data = readInventoryRows_(cfg);
  var items = data.rows.map(rowToItem_);

  // Group by storage
  var byStorage = {};
  items.forEach(function (it) {
    var key = it.room + '|' + it.storageId;
    if (!byStorage[key]) {
      byStorage[key] = {
        id: it.storageId, room: it.room, label: it.storageLabel,
        specific: it.specificLocation, code: it.locationCode,
        items: [], itemCount: 0, attention: 0, chemicalCount: 0,
        qrReady: false, hazard: false, isPlaceholderOnly: true,
      };
    }
    var b = byStorage[key];
    if (!isPlaceholderItem_(it)) {
      b.items.push(it);
      b.itemCount++;
      if (CONFIG.ATTENTION_STATUSES.indexOf(it.status) >= 0) b.attention++;
      if (it.hazard) { b.chemicalCount++; b.hazard = true; }
      b.isPlaceholderOnly = false;
    }
    if (it.qrCodeImage) b.qrReady = true;
  });

  var locations = Object.keys(byStorage).map(function (k) {
    var b = byStorage[k];
    return {
      id: b.id, room: b.room, label: b.label, specific: b.specific, code: b.code,
      itemCount: b.itemCount, attention: b.attention,
      chemicalCount: b.chemicalCount, hazard: b.hazard,
      qrReady: b.qrReady, isPlaceholderOnly: b.isPlaceholderOnly,
    };
  });

  // Sort: rollout room first, then alphabetical by room then label
  locations.sort(function (a, b) {
    if (a.room === CONFIG.ROLLOUT_ROOM && b.room !== CONFIG.ROLLOUT_ROOM) return -1;
    if (b.room === CONFIG.ROLLOUT_ROOM && a.room !== CONFIG.ROLLOUT_ROOM) return 1;
    if (a.room !== b.room) return a.room < b.room ? -1 : 1;
    return a.label < b.label ? -1 : 1;
  });

  // Room aggregates
  var roomMap = {};
  locations.forEach(function (l) {
    if (!roomMap[l.room]) roomMap[l.room] = { code: l.room, name: l.room, locationCount: 0, itemCount: 0, attention: 0, chemicalCount: 0, qrReady: 0 };
    var rm = roomMap[l.room];
    rm.locationCount++;
    rm.itemCount += l.itemCount;
    rm.attention += l.attention;
    rm.chemicalCount += l.chemicalCount;
    if (l.qrReady) rm.qrReady++;
  });
  var rooms = Object.keys(roomMap).map(function (k) { return roomMap[k]; });
  rooms.sort(function (a, b) {
    if (a.code === CONFIG.ROLLOUT_ROOM) return -1;
    if (b.code === CONFIG.ROLLOUT_ROOM) return 1;
    return a.code < b.code ? -1 : 1;
  });
  rooms.forEach(function (r) { r.rollout = r.code === CONFIG.ROLLOUT_ROOM; });

  // Rollout 419A
  var roll = roomMap[CONFIG.ROLLOUT_ROOM] || { locationCount: 0, qrReady: 0, itemCount: 0, chemicalCount: 0 };
  var rolloutPct = roll.locationCount ? Math.round(100 * roll.qrReady / roll.locationCount) : 0;

  // Readiness
  var readiness = computeReadiness_(items, locations);

  return {
    rooms: rooms,
    locations: locations,
    totals: {
      rooms: rooms.length,
      locations: locations.length,
      items: items.filter(function (i) { return !isPlaceholderItem_(i); }).length,
      chemicals: items.filter(function (i) { return i.hazard && !isPlaceholderItem_(i); }).length,
      attention: items.filter(function (i) { return CONFIG.ATTENTION_STATUSES.indexOf(i.status) >= 0; }).length,
    },
    rollout419A: {
      totalStorages: roll.locationCount,
      qrReady: roll.qrReady,
      percent: rolloutPct,
      itemRows: roll.itemCount,
      chemicals: roll.chemicalCount,
    },
    readiness: readiness,
    headers: data.headers,
  };
}

function computeReadiness_(items, locations) {
  var errors = 0, warnings = 0, missingQr = 0, missingSid = 0, invalidQty = 0, duplicates = 0;
  var seen = {};
  items.forEach(function (it) {
    if (isPlaceholderItem_(it)) return;
    if (!it.storageId) { errors++; missingSid++; }
    if (!it.itemId) errors++;
    if (it.qty < 0 || !isFinite(it.qty)) { errors++; invalidQty++; }
    if (CONFIG.STATUS_VALUES.indexOf(it.status) < 0) warnings++;
    var key = it.storageId + '::' + it.itemId;
    if (it.itemId && seen[key]) { warnings++; duplicates++; } else { seen[key] = true; }
  });
  locations.forEach(function (l) {
    if (!l.qrReady) { warnings++; missingQr++; }
  });
  var deductions = errors * 8 + warnings * 1;
  var score = Math.max(0, Math.min(100, 100 - deductions));
  return {
    score: score, errors: errors, warnings: warnings,
    missingQr: missingQr, missingStorageId: missingSid,
    invalidQty: invalidQty, duplicates: duplicates,
  };
}

// ──────────────────────────────────────────────────────────────────────────
// SAVE BRIDGE (Tech Mode)
// ──────────────────────────────────────────────────────────────────────────

function applyInventoryUpdates(payload) {
  // payload: { room, storageId, edits: [{sheetRow, qty, status}] }
  var cfg = getConfig_();
  if (!cfg.spreadsheetId) return { ok: false, error: 'SPREADSHEET_ID not configured' };
  var sh = getInventorySheet_(cfg);
  if (!sh) return { ok: false, error: 'Inventory sheet missing' };

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var qtyCol = headers.indexOf('Quantity') + 1;
  var statusCol = headers.indexOf('Status') + 1;
  if (!qtyCol || !statusCol) return { ok: false, error: 'Missing required columns' };

  var edits = (payload && payload.edits) || [];
  var applied = 0;
  edits.forEach(function (e) {
    var row = Number(e.sheetRow);
    if (!row || row < 2) return;
    var qty = Number(e.qty);
    if (!isFinite(qty) || qty < 0) return;
    if (CONFIG.STATUS_VALUES.indexOf(e.status) < 0) return;
    sh.getRange(row, qtyCol).setValue(qty);
    sh.getRange(row, statusCol).setValue(e.status);
    applied++;
  });
  return { ok: true, applied: applied };
}

// ══════════════════════════════════════════════════════════════════════════
// PAGE SHELL
// ══════════════════════════════════════════════════════════════════════════

function buildPageHtml_(opts) {
  var title = escape_(opts.title || 'D&T QR Inventory');
  var modeAttr = opts.mode || 'landing';
  var body = opts.body || '';
  return '<!DOCTYPE html><html lang="en"><head>' +
    '<meta charset="utf-8"/>' +
    '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"/>' +
    '<title>' + title + '</title>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com"/>' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>' +
    '<link href="https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet"/>' +
    '<script src="https://cdn.tailwindcss.com"></script>' +
    '<style>' + appCss_() + '</style>' +
    '</head><body data-mode="' + modeAttr + '" class="min-h-screen">' +
    renderAppHeader_(modeAttr) +
    '<main class="page">' + body + '</main>' +
    '<script>' + clientJs_() + '</script>' +
    '</body></html>';
}

function appCss_() {
  return [
    ':root{--ink:#0f172a;--ink-2:#334155;--muted:#64748b;--line:#dbe3ee;--bg:#f3f6fa;--card:#fff;',
    '--blue:#1d4ed8;--blue-50:#eff6ff;--teal:#0f766e;--teal-50:#ccfbf1;',
    '--amber:#b45309;--amber-50:#fef3c7;--red:#b91c1c;--red-50:#fee2e2;--em:#047857;--em-50:#d1fae5;}',
    '*{box-sizing:border-box}',
    'html,body{margin:0;background:var(--bg);color:var(--ink);font-family:"Inter Tight",system-ui,sans-serif;-webkit-font-smoothing:antialiased}',
    'body[data-mode="tech"] main.page{background:#fffbeb}',
    'a{color:inherit;text-decoration:none}',
    '.font-mono{font-family:"JetBrains Mono",ui-monospace,monospace}',
    '.page{max-width:1400px;margin:0 auto;padding:20px 16px 96px}',
    '@media(min-width:768px){.page{padding:28px 32px 80px}}',
    /* App header */
    '.appbar{position:sticky;top:0;z-index:30;background:rgba(255,255,255,.92);backdrop-filter:blur(8px);border-bottom:1px solid var(--line)}',
    '.appbar-inner{max-width:1400px;margin:0 auto;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;gap:12px}',
    '@media(min-width:768px){.appbar-inner{padding:12px 32px}}',
    '.brand{display:flex;align-items:center;gap:10px;min-width:0}',
    '.brand-mark{width:36px;height:36px;border-radius:10px;background:var(--ink);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;letter-spacing:-0.02em;font-size:14px;flex-shrink:0}',
    '.brand-text{min-width:0}',
    '.brand-eyebrow{font-size:10px;font-weight:600;letter-spacing:0.2em;color:var(--blue);text-transform:uppercase;line-height:1}',
    '.brand-title{font-size:14px;font-weight:600;color:var(--ink);line-height:1.2;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}',
    '.appbar-nav{display:flex;align-items:center;gap:4px}',
    '.navbtn{display:inline-flex;align-items:center;gap:6px;padding:8px 12px;border-radius:12px;font-size:12.5px;font-weight:500;color:var(--ink-2);min-height:40px}',
    '.navbtn:hover{background:#eef2f7}',
    '.navbtn.active{background:var(--ink);color:#fff}',
    '.navbtn .label{display:none}',
    '@media(min-width:640px){.navbtn .label{display:inline}}',
    /* Card primitives */
    '.card{background:var(--card);border:1px solid var(--line);border-radius:20px}',
    '.card-soft{background:var(--card);border-radius:20px;box-shadow:0 1px 2px rgba(15,23,42,.04)}',
    '.card-shell{background:var(--card);border:1px solid var(--line);border-radius:24px}',
    /* Pills/badges */
    '.pill{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:999px;font-size:11px;font-weight:600;letter-spacing:0.02em;line-height:1.5;border:1px solid transparent}',
    '.pill-blue{background:var(--blue-50);color:var(--blue);border-color:#bfdbfe}',
    '.pill-teal{background:var(--teal-50);color:var(--teal);border-color:#99f6e4}',
    '.pill-amber{background:var(--amber-50);color:var(--amber);border-color:#fde68a}',
    '.pill-red{background:var(--red-50);color:var(--red);border-color:#fecaca}',
    '.pill-em{background:var(--em-50);color:var(--em);border-color:#a7f3d0}',
    '.pill-neutral{background:#f1f5f9;color:var(--ink-2);border-color:#e2e8f0}',
    '.pill-mono{font-family:"JetBrains Mono",monospace;font-size:10.5px;letter-spacing:0}',
    /* Buttons */
    '.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:10px 16px;border-radius:14px;font-weight:600;font-size:13.5px;min-height:44px;transition:transform .08s,background .12s;cursor:pointer;border:1px solid transparent}',
    '.btn:active{transform:scale(.98)}',
    '.btn-primary{background:var(--blue);color:#fff}',
    '.btn-primary:hover{background:#1e40af}',
    '.btn-primary:disabled{opacity:.55;cursor:not-allowed}',
    '.btn-secondary{background:#fff;color:var(--ink);border-color:var(--line)}',
    '.btn-secondary:hover{background:#f8fafc}',
    '.btn-ghost{background:transparent;color:var(--ink-2)}',
    '.btn-ghost:hover{background:#eef2f7}',
    '.btn-success{background:var(--em);color:#fff}',
    '.btn-success:hover{background:#065f46}',
    '.btn-success:disabled{background:#9ca3af;cursor:not-allowed}',
    /* Inputs */
    '.input,.select{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:14px;font-size:14px;background:#fff;color:var(--ink);font-family:inherit;min-height:44px}',
    '.input:focus,.select:focus{outline:none;border-color:var(--blue);box-shadow:0 0 0 3px rgba(29,78,216,.15)}',
    '.qty-input{text-align:center;font-family:"JetBrains Mono",monospace;font-weight:600}',
    /* Filter chips */
    '.chips{display:flex;gap:8px;overflow-x:auto;padding-bottom:4px;-ms-overflow-style:none;scrollbar-width:none}',
    '.chips::-webkit-scrollbar{display:none}',
    '.chip{flex-shrink:0;display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border-radius:999px;background:#fff;color:var(--ink-2);border:1px solid var(--line);font-size:12.5px;font-weight:500;cursor:pointer;min-height:36px;white-space:nowrap}',
    '.chip:hover{background:#f8fafc}',
    '.chip.active{background:var(--ink);color:#fff;border-color:var(--ink)}',
    '.chip.rollout{border-color:#5eead4;color:var(--teal)}',
    '.chip.rollout.active{background:var(--teal);color:#fff;border-color:var(--teal)}',
    /* Storage card */
    '.storage-card{display:flex;flex-direction:column;background:#fff;border:1px solid var(--line);border-radius:24px;overflow:hidden}',
    '.storage-card.hazard{border-color:#fecaca}',
    '.storage-card-body{display:flex;align-items:flex-start;gap:14px;padding:16px}',
    '.storage-icon{flex-shrink:0;width:52px;height:52px;border-radius:16px;background:#f1f5f9;color:var(--ink-2);display:flex;align-items:center;justify-content:center}',
    '.storage-icon.hazard{background:var(--red-50);color:var(--red)}',
    '.storage-icon.tools{background:var(--amber-50);color:var(--amber)}',
    '.storage-icon.machine{background:var(--blue-50);color:var(--blue)}',
    '.storage-icon.material{background:var(--em-50);color:var(--em)}',
    '.storage-actions{display:flex;border-top:1px solid #f1f5f9}',
    '.storage-actions a{flex:1;text-align:center;padding:12px;font-size:13px;font-weight:600;color:var(--ink-2);min-height:44px;display:inline-flex;align-items:center;justify-content:center;gap:6px}',
    '.storage-actions a:hover{background:#f8fafc}',
    '.storage-actions a.primary{color:var(--blue)}',
    '.storage-actions a.primary:hover{background:var(--blue-50)}',
    '.storage-actions .divider{width:1px;background:#f1f5f9}',
    /* Item card */
    '.item-card{background:#fff;border:1px solid var(--line);border-radius:18px;padding:14px}',
    '.item-card.hazard{border-color:#fecaca;background:linear-gradient(0deg,#fff,#fff) padding-box,#fff}',
    '.item-card.changed{border-color:#fbbf24;box-shadow:0 0 0 3px rgba(251,191,36,.15)}',
    /* Save bar */
    '.save-bar{position:fixed;left:0;right:0;bottom:0;z-index:40;background:rgba(255,255,255,.96);backdrop-filter:blur(8px);border-top:1px solid var(--line);box-shadow:0 -8px 24px -12px rgba(15,23,42,.18);padding:10px 16px;display:flex;align-items:center;justify-content:space-between;gap:12px}',
    '@media(min-width:1024px){.save-bar{left:auto;right:24px;bottom:24px;border-radius:18px;border:1px solid var(--line);max-width:420px;box-shadow:0 12px 32px -12px rgba(15,23,42,.25)}}',
    /* Toasts */
    '.toast{position:fixed;left:50%;transform:translateX(-50%);top:18px;z-index:60;padding:10px 14px;border-radius:14px;font-size:13px;font-weight:500;box-shadow:0 12px 32px -12px rgba(15,23,42,.3);display:none;align-items:center;gap:8px;max-width:90vw}',
    '.toast.show{display:inline-flex}',
    '.toast.ok{background:var(--em);color:#fff}',
    '.toast.err{background:var(--red);color:#fff}',
    /* Section labels */
    '.section-label{font-size:10.5px;font-weight:600;letter-spacing:0.2em;color:var(--muted);text-transform:uppercase}',
    /* Print-only QR labels */
    '.label-sheet{background:#fff;padding:12mm;border-radius:16px}',
    '.qr-label{display:flex;gap:14px;border:1.5px dashed #cbd5e1;border-radius:14px;padding:14px;page-break-inside:avoid}',
    '.qr-label.hazard{border-color:#fca5a5;background:#fff5f5}',
    '.qr-glyph{width:96px;height:96px;background:repeating-conic-gradient(#0f172a 0% 25%,#fff 0% 50%) 50% / 12px 12px;border:2px solid #0f172a;border-radius:6px;flex-shrink:0}',
    '@media print{',
    '.appbar,.save-bar,.no-print{display:none!important}',
    'body{background:#fff!important}',
    '.page{padding:0!important;max-width:none!important}',
    '.label-sheet{box-shadow:none;border:0;border-radius:0;padding:8mm}',
    '@page{size:A4;margin:8mm}',
    '}',
    /* Focus */
    '*:focus-visible{outline:2px solid var(--blue);outline-offset:2px;border-radius:4px}',
    '.gauge{transform:rotate(-90deg)}',
    '.gauge circle{fill:none;stroke-width:10;stroke-linecap:round}',
  ].join('');
}

function renderAppHeader_(mode) {
  return '<header class="appbar"><div class="appbar-inner">' +
    '<a href="?" class="brand" aria-label="D&T QR Inventory home">' +
      '<div class="brand-mark">D&amp;T</div>' +
      '<div class="brand-text">' +
        '<div class="brand-eyebrow">Victoria Shanghai Academy</div>' +
        '<div class="brand-title">D&amp;T QR Inventory</div>' +
      '</div>' +
    '</a>' +
    '<nav class="appbar-nav" aria-label="App navigation">' +
      navLink_('?', 'All Locations', 'home', mode === 'landing') +
      navLink_('?admin=readiness', 'Readiness', 'shield', mode === 'admin' && /readiness/.test(String(arguments[0]||''))) +
      navLink_('?admin=labels', 'QR Labels', 'qr', mode === 'labels') +
      navLink_('?admin=diagnostics', 'Diagnostics', 'info', false) +
    '</nav>' +
  '</div></header>';
}

function navLink_(href, label, icon, active) {
  return '<a class="navbtn ' + (active ? 'active' : '') + '" href="' + href + '">' +
    iconSvg_(icon) + '<span class="label">' + escape_(label) + '</span></a>';
}

// ══════════════════════════════════════════════════════════════════════════
// LANDING
// ══════════════════════════════════════════════════════════════════════════

function renderLandingHtml_(data, cfg) {
  var pieces = [];

  // Optional WEB_APP_BASE_URL warning
  if (!cfg.webAppBaseUrl) {
    pieces.push(renderNoticePanel_({
      tone: 'amber',
      icon: 'warn',
      title: 'WEB_APP_BASE_URL is not configured',
      body: 'In-app browsing works. QR generation and external links are disabled until you set it.',
      cta: { href: '?admin=diagnostics', label: 'Open diagnostics' },
    }));
  }

  // Hero rollout card
  var roll = data.rollout419A;
  pieces.push(
    '<section class="card-shell p-5 md:p-6" style="background:linear-gradient(135deg,#ecfdf5 0%,#fff 70%);border-color:#bbf7d0">' +
    '<div class="flex flex-wrap items-center justify-between gap-4">' +
      '<div class="min-w-0">' +
        '<div class="flex items-center gap-2 mb-1"><span class="section-label" style="color:var(--teal)">Phased rollout</span>' +
        '<span class="pill pill-teal">419A</span></div>' +
        '<h2 class="text-[22px] md:text-[26px] font-semibold tracking-tight">D&amp;T Workshop · Room 419A</h2>' +
        '<p class="text-[13.5px] text-slate-600 mt-1">' + roll.qrReady + ' of ' + roll.totalStorages + ' storages QR-ready · ' + roll.itemRows + ' items · ' + roll.chemicals + ' chemical entries</p>' +
      '</div>' +
      '<div class="flex items-center gap-4">' +
        gauge_(roll.percent, '#0d9488') +
        '<div><div class="text-[36px] font-semibold leading-none tracking-tight" style="color:var(--teal)">' + roll.percent + '%</div>' +
        '<div class="text-[12px] text-slate-500 mt-1">Rollout ready</div></div>' +
      '</div>' +
    '</div>' +
    '<div class="mt-4 h-2 rounded-full bg-white/70 overflow-hidden border border-teal-100">' +
      '<div class="h-full rounded-full" style="width:' + roll.percent + '%;background:var(--teal)"></div>' +
    '</div>' +
    '</section>'
  );

  // Metrics row — desktop dashboard
  pieces.push(
    '<section class="mt-4 grid grid-cols-2 md:grid-cols-5 gap-3">' +
      renderMetricCard_({ label: 'Rooms', value: data.totals.rooms }) +
      renderMetricCard_({ label: 'Storages', value: data.totals.locations }) +
      renderMetricCard_({ label: 'Items', value: data.totals.items }) +
      renderMetricCard_({ label: 'Attention', value: data.totals.attention, tone: data.totals.attention > 0 ? 'amber' : 'neutral' }) +
      renderMetricCard_({ label: 'Chemicals', value: data.totals.chemicals, tone: data.totals.chemicals > 0 ? 'red' : 'neutral' }) +
    '</section>'
  );

  // Two-column: directory left, sticky summary right (desktop only)
  pieces.push('<div class="mt-6 grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6">');

  // Left column: search + filters + directory
  pieces.push('<div>');
  pieces.push(
    '<div class="flex flex-col sm:flex-row gap-3 mb-3">' +
      '<div class="relative flex-1">' +
        '<span class="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400">' + iconSvg_('search') + '</span>' +
        '<input id="storageSearch" class="input pl-11" placeholder="Search room, storage ID, label, code…" autocomplete="off" aria-label="Search storage"/>' +
      '</div>' +
    '</div>'
  );

  // Room chips
  pieces.push('<div class="chips mb-4" id="roomChips" role="tablist" aria-label="Filter by room">');
  pieces.push('<button class="chip active" data-room="__all" role="tab">All rooms</button>');
  data.rooms.forEach(function (r) {
    pieces.push('<button class="chip ' + (r.rollout ? 'rollout' : '') + '" data-room="' + escape_(r.code) + '" role="tab">' +
      escape_(r.code) + (r.rollout ? ' · rollout' : '') + ' · ' + r.locationCount + '</button>');
  });
  pieces.push('</div>');

  // Storage grid
  pieces.push('<div id="storageGrid" class="grid grid-cols-1 md:grid-cols-2 gap-3">');
  data.locations.forEach(function (l) { pieces.push(renderStorageCard_(l)); });
  pieces.push('</div>');

  pieces.push(
    '<p id="storageEmpty" class="hidden text-center text-[13px] text-slate-500 py-12">No storage matches your filters.</p>'
  );

  pieces.push('</div>'); // end left col

  // Right sidebar (desktop)
  pieces.push('<aside class="hidden lg:block space-y-4" style="position:sticky;top:84px;align-self:start">');
  pieces.push(
    '<section class="card p-4">' +
      '<div class="flex items-center justify-between mb-3"><h3 class="font-semibold">Data readiness</h3>' +
      '<span class="pill ' + (data.readiness.score >= 80 ? 'pill-em' : data.readiness.score >= 60 ? 'pill-amber' : 'pill-red') + '">' + data.readiness.score + ' / 100</span></div>' +
      '<div class="space-y-1.5 text-[12.5px] text-slate-600">' +
        readinessRow_('Critical errors', data.readiness.errors, 'red') +
        readinessRow_('Warnings', data.readiness.warnings, 'amber') +
        readinessRow_('Missing QR links', data.readiness.missingQr, 'amber') +
        readinessRow_('Duplicate rows', data.readiness.duplicates, 'amber') +
      '</div>' +
      '<a class="btn btn-secondary w-full mt-4" href="?admin=readiness">' + iconSvg_('shield') + 'Open readiness</a>' +
    '</section>'
  );
  pieces.push(
    '<section class="card p-4">' +
      '<h3 class="font-semibold mb-3">Configuration</h3>' +
      '<div class="space-y-2 text-[12.5px]">' +
        diagBullet_(true, 'SPREADSHEET_ID') +
        diagBullet_(!!cfg.webAppBaseUrl, 'WEB_APP_BASE_URL', cfg.webAppBaseUrl ? '' : 'QR generation disabled') +
        diagBullet_(true, 'Inventory sheet · ' + cfg.inventorySheetName) +
      '</div>' +
      '<a class="btn btn-secondary w-full mt-4" href="?admin=diagnostics">' + iconSvg_('info') + 'Open diagnostics</a>' +
    '</section>'
  );
  pieces.push('</aside>');

  pieces.push('</div>'); // end two-col grid

  return pieces.join('');
}

function renderStorageCard_(l) {
  var iconCls = l.hazard ? 'hazard' : (/^TOL/.test(l.id) ? 'tools' : /^MAC/.test(l.id) ? 'machine' : /^MAT/.test(l.id) ? 'material' : '');
  var iconName = l.hazard ? 'flask' : /^TOL/.test(l.id) ? 'wrench' : /^MAC/.test(l.id) ? 'cpu' : /^MAT/.test(l.id) ? 'box' : 'box';
  var hazardCls = l.hazard ? ' hazard' : '';
  var hayParts = [l.id, l.label, l.specific || '', l.code || '', l.room].join(' ').toLowerCase();
  var viewHref = '?room=' + encodeURIComponent(l.room) + '&loc=' + encodeURIComponent(l.id);
  var techHref = viewHref + '&mode=tech';
  return '<article class="storage-card' + hazardCls + '" data-room="' + escape_(l.room) + '" data-hay="' + escape_(hayParts) + '">' +
    '<a href="' + viewHref + '" class="storage-card-body" aria-label="Open ' + escape_(l.label) + '">' +
      '<div class="storage-icon ' + iconCls + '">' + iconSvg_(iconName, 22) + '</div>' +
      '<div class="flex-1 min-w-0">' +
        '<div class="flex items-center gap-1.5 flex-wrap">' +
          '<h3 class="font-semibold text-[15px] text-slate-900 truncate">' + escape_(l.label) + '</h3>' +
          (l.hazard ? '<span class="pill pill-red">' + iconSvg_('hazard', 11) + 'Chemical</span>' : '') +
        '</div>' +
        '<p class="text-[12px] text-slate-500 mt-0.5 truncate">' + escape_(l.specific || '—') + '</p>' +
        '<div class="flex flex-wrap gap-1.5 mt-2">' +
          '<span class="pill pill-neutral pill-mono">' + escape_(l.id) + '</span>' +
          (l.code ? '<span class="pill pill-blue pill-mono">' + escape_(l.code) + '</span>' : '') +
          (l.isPlaceholderOnly ? '<span class="pill pill-neutral">Empty</span>' : '') +
          (!l.qrReady ? '<span class="pill pill-amber">QR missing</span>' : '<span class="pill pill-em">QR ready</span>') +
          (l.attention > 0 ? '<span class="pill pill-amber">' + l.attention + ' attention</span>' : '') +
        '</div>' +
      '</div>' +
      '<div class="text-right flex-shrink-0">' +
        '<div class="text-[22px] font-semibold leading-none">' + l.itemCount + '</div>' +
        '<div class="text-[10.5px] text-slate-500 mt-0.5 uppercase tracking-wider">items</div>' +
      '</div>' +
    '</a>' +
    '<div class="storage-actions">' +
      '<a href="' + viewHref + '">' + iconSvg_('search', 14) + 'View</a>' +
      '<div class="divider"></div>' +
      '<a href="' + techHref + '" class="primary">' + iconSvg_('edit', 14) + 'Tech update</a>' +
    '</div>' +
  '</article>';
}

function renderMetricCard_(opts) {
  var tone = opts.tone || 'neutral';
  var color = tone === 'red' ? 'var(--red)' : tone === 'amber' ? 'var(--amber)' : 'var(--ink)';
  return '<div class="card-soft border border-[var(--line)] p-4">' +
    '<div class="section-label">' + escape_(opts.label) + '</div>' +
    '<div class="text-[26px] font-semibold tracking-tight mt-1.5" style="color:' + color + '">' + escape_(String(opts.value)) + '</div>' +
    (opts.sub ? '<div class="text-[11.5px] text-slate-500 mt-1">' + escape_(opts.sub) + '</div>' : '') +
    '</div>';
}

function readinessRow_(label, n, tone) {
  var dot = tone === 'red' ? 'background:var(--red)' : 'background:var(--amber)';
  return '<div class="flex items-center justify-between"><span class="flex items-center gap-2"><span class="w-2 h-2 rounded-full" style="' + dot + '"></span>' + escape_(label) + '</span><span class="font-semibold text-slate-900">' + n + '</span></div>';
}

function diagBullet_(ok, label, hint) {
  return '<div class="flex items-start gap-2"><span class="w-2 h-2 rounded-full mt-1.5" style="background:' + (ok ? 'var(--em)' : 'var(--amber)') + '"></span>' +
    '<div class="flex-1"><div class="text-[12.5px] font-medium">' + escape_(label) + '</div>' +
    (hint ? '<div class="text-[11px] text-slate-500">' + escape_(hint) + '</div>' : '') + '</div>' +
    '<span class="text-[11px] font-mono ' + (ok ? 'text-emerald-700' : 'text-amber-700') + '">' + (ok ? 'OK' : 'Missing') + '</span></div>';
}

function gauge_(pct, color) {
  var c = 2 * Math.PI * 28;
  var offset = c * (1 - pct / 100);
  return '<svg class="gauge" width="72" height="72" viewBox="0 0 72 72">' +
    '<circle cx="36" cy="36" r="28" stroke="#e2e8f0"/>' +
    '<circle cx="36" cy="36" r="28" stroke="' + color + '" stroke-dasharray="' + c.toFixed(2) + '" stroke-dashoffset="' + offset.toFixed(2) + '"/>' +
    '</svg>';
}

// ══════════════════════════════════════════════════════════════════════════
// INVENTORY (View + Tech)
// ══════════════════════════════════════════════════════════════════════════

function renderInventoryHtml_(bundle, mode) {
  var loc = bundle.location;
  var items = bundle.items;
  var techMode = mode === 'tech';

  // Counts
  var attention = items.filter(function (it) { return CONFIG.ATTENTION_STATUSES.indexOf(it.status) >= 0; }).length;
  var chems = items.filter(function (it) { return it.hazard; }).length;

  var hero =
    '<section class="card-shell p-5 md:p-6">' +
      '<div class="flex items-start gap-3">' +
        '<a href="?" class="btn btn-ghost p-2" aria-label="Back to all locations">' + iconSvg_('arrow-left', 20) + '</a>' +
        '<div class="flex-1 min-w-0">' +
          '<div class="flex items-center gap-2 mb-1">' +
            '<span class="section-label" style="color:var(--blue)">Room ' + escape_(loc.room) + '</span>' +
            (loc.room === CONFIG.ROLLOUT_ROOM ? '<span class="pill pill-teal">Rollout</span>' : '') +
            (techMode ? '<span class="pill pill-amber">' + iconSvg_('edit', 11) + 'Technician Mode</span>' : '<span class="pill pill-blue">' + iconSvg_('search', 11) + 'View Mode</span>') +
          '</div>' +
          '<h1 class="text-[22px] md:text-[28px] font-semibold tracking-tight flex items-center gap-2">' + escape_(loc.label) +
          (loc.hazard ? iconSvg_('hazard', 18, 'color:var(--red)') : '') + '</h1>' +
          '<p class="text-[13px] text-slate-600 mt-1">' + escape_(loc.specific || '') + '</p>' +
          '<div class="flex flex-wrap gap-1.5 mt-3">' +
            '<span class="pill pill-neutral pill-mono">' + escape_(loc.id) + '</span>' +
            (loc.code ? '<span class="pill pill-blue pill-mono">' + escape_(loc.code) + '</span>' : '') +
            '<span class="pill pill-neutral">' + items.length + ' item' + (items.length !== 1 ? 's' : '') + '</span>' +
            (chems > 0 ? '<span class="pill pill-red">' + iconSvg_('hazard', 11) + chems + ' chemical</span>' : '') +
            (attention > 0 ? '<span class="pill pill-amber">' + attention + ' attention</span>' : '') +
          '</div>' +
        '</div>' +
        '<a href="?room=' + encodeURIComponent(loc.room) + '&loc=' + encodeURIComponent(loc.id) + (techMode ? '' : '&mode=tech') + '" class="btn ' + (techMode ? 'btn-secondary' : 'btn-primary') + ' hidden md:inline-flex">' +
          iconSvg_(techMode ? 'search' : 'edit', 15) +
          (techMode ? 'Switch to View' : 'Switch to Tech') +
        '</a>' +
      '</div>' +
    '</section>';

  if (items.length === 0) {
    return hero + renderEmptyState_({
      icon: 'box',
      title: 'No items entered yet',
      body: 'This storage exists in the master list but has no inventory rows.',
      cta: techMode ? { href: '#', label: '+ Add first item', primary: true } : null,
    });
  }

  // Filters
  var filters =
    '<section class="mt-4 grid grid-cols-1 md:grid-cols-[1fr_220px] gap-3">' +
      '<div class="relative">' +
        '<span class="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400">' + iconSvg_('search') + '</span>' +
        '<input id="itemSearch" class="input pl-11" placeholder="Search items, IDs, remarks…" aria-label="Search items"/>' +
      '</div>' +
      '<select id="itemStatus" class="select" aria-label="Filter by status">' +
        '<option value="__all">All statuses</option>' +
        CONFIG.STATUS_VALUES.map(function (s) { return '<option>' + s + '</option>'; }).join('') +
      '</select>' +
    '</section>';

  // Items list
  var itemsHtml = '<section id="itemsList" class="mt-4 space-y-2.5">' +
    items.map(function (it) { return techMode ? renderTechItemCard_(it) : renderItemCard_(it); }).join('') +
    '</section>' +
    '<p id="itemsEmpty" class="hidden text-center text-[13px] text-slate-500 py-12">No items match your search.</p>';

  // Tech sticky save bar
  var saveBar = techMode ?
    '<div id="saveBar" class="save-bar" role="status" aria-live="polite">' +
      '<div class="text-[13px]"><span id="changedCount" class="text-slate-500">No changes yet</span></div>' +
      '<div class="flex items-center gap-2">' +
        '<button id="resetBtn" class="btn btn-ghost text-[13px]" disabled>Reset</button>' +
        '<button id="saveBtn" class="btn btn-success" disabled>' + iconSvg_('save', 15) + 'Save updates</button>' +
      '</div>' +
    '</div>' +
    '<div id="toast" class="toast" role="alert" aria-live="assertive"></div>' +
    '<script>window.__BUNDLE__=' + JSON.stringify({ room: loc.room, storageId: loc.id }) + ';</script>'
    : '';

  // Two-column on desktop: items left, summary right
  var summary = '<aside class="hidden lg:block space-y-3" style="position:sticky;top:84px;align-self:start">' +
    '<section class="card p-4">' +
      '<h3 class="font-semibold mb-3">Status breakdown</h3>' +
      '<div class="space-y-2 text-[13px]">' +
        CONFIG.STATUS_VALUES.map(function (s) {
          var n = items.filter(function (it) { return it.status === s; }).length;
          return '<div class="flex items-center justify-between"><span>' + escape_(s) + '</span><span class="font-semibold">' + n + '</span></div>';
        }).join('') +
      '</div>' +
    '</section>' +
    (chems > 0 ? '<section class="card p-4 border-red-200 bg-red-50/40"><h3 class="font-semibold text-red-900 flex items-center gap-2">' + iconSvg_('hazard', 14) + 'Chemical handling</h3><p class="text-[12.5px] text-red-800 mt-2">' + chems + ' item' + (chems > 1 ? 's' : '') + ' require D&amp;T chemical storage protocol. Wear PPE; log usage.</p></section>' : '') +
  '</aside>';

  var main = '<div class="mt-2 grid grid-cols-1 lg:grid-cols-[1fr_320px] gap-6">' +
    '<div>' + filters + itemsHtml + '</div>' +
    summary +
  '</div>';

  return hero + main + saveBar;
}

function renderItemCard_(it) {
  return '<article class="item-card' + (it.hazard ? ' hazard' : '') + '" data-hay="' + escape_((it.itemId + ' ' + it.itemName + ' ' + (it.category || '') + ' ' + (it.remarks || '')).toLowerCase()) + '" data-status="' + escape_(it.status) + '">' +
    '<div class="flex items-start justify-between gap-3">' +
      '<div class="min-w-0 flex-1">' +
        '<div class="flex items-center gap-1.5 flex-wrap">' +
          '<h3 class="font-semibold text-[15px] leading-tight">' + escape_(it.itemName) + '</h3>' +
          (it.hazard ? '<span class="pill pill-red">' + iconSvg_('hazard', 11) + 'Chemical</span>' : '') +
        '</div>' +
        '<div class="text-[11.5px] text-slate-500 mt-1 flex flex-wrap gap-x-3 gap-y-0.5">' +
          '<span class="font-mono">' + escape_(it.itemId) + '</span>' +
          (it.category ? '<span>· ' + escape_(it.category) + '</span>' : '') +
          (it.unit ? '<span>· ' + escape_(it.unit) + '</span>' : '') +
        '</div>' +
        (it.remarks ? '<p class="text-[12.5px] text-slate-600 mt-2"><strong class="text-slate-800">Remarks:</strong> ' + escape_(it.remarks) + '</p>' : '') +
      '</div>' +
      '<div class="text-right rounded-2xl bg-slate-50 border border-slate-200 px-3 py-2 min-w-[80px]">' +
        '<div class="section-label">Expected</div>' +
        '<div class="text-[22px] font-semibold leading-none mt-1">' + it.qty + '</div>' +
        (it.unit ? '<div class="text-[10.5px] text-slate-500 mt-1">' + escape_(it.unit) + '</div>' : '') +
        '<div class="mt-1.5 flex justify-end">' + statusPill_(it.status) + '</div>' +
      '</div>' +
    '</div>' +
  '</article>';
}

function renderTechItemCard_(it) {
  return '<article class="item-card tech-row" data-row="' + it.sheetRow + '" data-orig-qty="' + it.qty + '" data-orig-status="' + escape_(it.status) + '" data-hay="' + escape_((it.itemId + ' ' + it.itemName).toLowerCase()) + '" data-status="' + escape_(it.status) + '">' +
    '<div class="flex items-start justify-between gap-3">' +
      '<div class="min-w-0 flex-1">' +
        '<div class="flex items-center gap-1.5 flex-wrap">' +
          '<h3 class="font-semibold text-[15px]">' + escape_(it.itemName) + '</h3>' +
          (it.hazard ? '<span class="pill pill-red">' + iconSvg_('hazard', 11) + 'Chemical</span>' : '') +
        '</div>' +
        '<div class="text-[11.5px] text-slate-500 mt-1"><span class="font-mono">' + escape_(it.itemId) + '</span>' +
          (it.unit ? ' · ' + escape_(it.unit) : '') + '</div>' +
        (it.remarks ? '<p class="text-[12px] text-slate-600 mt-1.5">' + escape_(it.remarks) + '</p>' : '') +
        '<p class="changed-msg hidden text-[11.5px] text-amber-800 mt-2">' + iconSvg_('edit', 11) + ' <span class="changed-msg-text"></span></p>' +
      '</div>' +
      '<div class="flex-shrink-0 w-[180px] space-y-2">' +
        '<div><label class="section-label block mb-1">Quantity</label>' +
          '<div class="flex items-center gap-1">' +
            '<button class="btn btn-secondary p-2 qty-dec" aria-label="Decrease">−</button>' +
            '<input type="number" min="0" class="input qty-input qty-val" value="' + it.qty + '"/>' +
            '<button class="btn btn-secondary p-2 qty-inc" aria-label="Increase">+</button>' +
          '</div>' +
        '</div>' +
        '<div><label class="section-label block mb-1">Status</label>' +
          '<select class="select status-val text-[13px] py-2" aria-label="Status">' +
            CONFIG.STATUS_VALUES.map(function (s) { return '<option' + (s === it.status ? ' selected' : '') + '>' + s + '</option>'; }).join('') +
          '</select>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</article>';
}

function statusPill_(s) {
  var cls = s === 'Good' ? 'pill-em' : s === 'Low Stock' ? 'pill-amber' : s === 'Missing' ? 'pill-red' : s === 'Needs Maintenance' ? 'pill-blue' : 'pill-neutral';
  return '<span class="pill ' + cls + '">' + escape_(s) + '</span>';
}

// ══════════════════════════════════════════════════════════════════════════
// ADMIN · READINESS
// ══════════════════════════════════════════════════════════════════════════

function renderAdminReadinessHtml_(data) {
  var r = data.readiness;
  var color = r.score >= 80 ? '#10b981' : r.score >= 60 ? '#f59e0b' : '#ef4444';
  return '<header class="mb-5"><div class="section-label">Admin</div><h1 class="text-[26px] md:text-[30px] font-semibold tracking-tight">Data readiness</h1><p class="text-[13.5px] text-slate-600 mt-1">Run before any 419A rollout. Fix all errors; review warnings.</p></header>' +
    '<div class="grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6">' +
      '<div>' +
        '<section class="card p-5 flex items-center gap-5">' +
          gauge_(r.score, color) +
          '<div><div class="section-label">Readiness score</div>' +
          '<div class="text-[40px] font-semibold tracking-tight leading-none mt-1" style="color:' + color + '">' + r.score + '<span class="text-[18px] text-slate-400">/100</span></div>' +
          '<div class="text-[12.5px] text-slate-600 mt-1">' + r.errors + ' error' + (r.errors !== 1 ? 's' : '') + ' · ' + r.warnings + ' warning' + (r.warnings !== 1 ? 's' : '') + '</div></div>' +
        '</section>' +
        '<section class="grid grid-cols-2 md:grid-cols-3 gap-3 mt-4">' +
          issueTile_(r.errors, 'Critical errors', 'red') +
          issueTile_(r.warnings, 'Warnings', 'amber') +
          issueTile_(r.missingQr, 'Missing QR links', 'amber') +
          issueTile_(r.missingStorageId, 'Missing Storage IDs', 'red') +
          issueTile_(r.invalidQty, 'Invalid quantities', 'red') +
          issueTile_(r.duplicates, 'Duplicate rows', 'amber') +
        '</section>' +
        '<section class="card p-5 mt-4">' +
          '<div class="flex items-center justify-between mb-3"><h3 class="font-semibold">419A rollout panel</h3>' +
          '<span class="pill pill-teal">' + data.rollout419A.percent + '%</span></div>' +
          '<div class="h-2 rounded-full bg-slate-100 overflow-hidden"><div class="h-full" style="width:' + data.rollout419A.percent + '%;background:var(--teal)"></div></div>' +
          '<p class="text-[12.5px] text-slate-600 mt-2">' + data.rollout419A.qrReady + ' of ' + data.rollout419A.totalStorages + ' storages QR-ready · ' + data.rollout419A.itemRows + ' item rows · ' + data.rollout419A.chemicals + ' chemicals.</p>' +
        '</section>' +
      '</div>' +
      '<aside class="space-y-4" style="position:sticky;top:84px;align-self:start">' +
        '<section class="card p-4"><h3 class="font-semibold mb-3">Next actions</h3><div class="space-y-2 text-[13px]">' +
          (r.errors > 0 ? actionRow_('Fix critical errors first', 'Open spreadsheet, run Create Readiness Report', 'red') : '') +
          (r.missingQr > 0 ? actionRow_('Refresh QR Links', 'D&T Inventory menu → Refresh QR Links', 'amber') : '') +
          (r.duplicates > 0 ? actionRow_('Resolve duplicate rows', 'Review Inventory_Readiness_Report', 'amber') : '') +
          (r.errors === 0 && r.warnings === 0 ? '<p class="text-[13px] text-emerald-700">All checks pass. Ready to roll out.</p>' : '') +
        '</div></section>' +
        '<section class="card p-4"><h3 class="font-semibold mb-3">Diagnostics</h3>' +
          diagBullet_(true, 'SPREADSHEET_ID', '') +
          '<a class="btn btn-secondary w-full mt-3" href="?admin=diagnostics">' + iconSvg_('info') + 'Open diagnostics</a>' +
        '</section>' +
      '</aside>' +
    '</div>';
}

function issueTile_(n, label, tone) {
  var cls = tone === 'red' ? 'text-red-700 bg-red-50 border-red-200' : tone === 'amber' ? 'text-amber-700 bg-amber-50 border-amber-200' : 'text-slate-700 bg-slate-50 border-slate-200';
  return '<div class="rounded-2xl border ' + cls + ' p-4">' +
    '<div class="text-[28px] font-semibold leading-none">' + n + '</div>' +
    '<div class="text-[12px] mt-1.5 font-medium">' + escape_(label) + '</div></div>';
}

function actionRow_(title, hint, tone) {
  var dot = tone === 'red' ? 'background:var(--red)' : 'background:var(--amber)';
  return '<div class="flex items-start gap-2"><span class="w-2 h-2 rounded-full mt-1.5" style="' + dot + '"></span>' +
    '<div><div class="font-medium text-[13px]">' + escape_(title) + '</div>' +
    '<div class="text-[11.5px] text-slate-500 mt-0.5">' + escape_(hint) + '</div></div></div>';
}

// ══════════════════════════════════════════════════════════════════════════
// ADMIN · QR LABELS
// ══════════════════════════════════════════════════════════════════════════

function renderQrLabelsHtml_(data, cfg) {
  var labels = data.locations.filter(function (l) { return l.room === CONFIG.ROLLOUT_ROOM; });
  if (!labels.length) labels = data.locations.slice(0, 12);

  var toolbar = '<div class="no-print flex flex-wrap items-center justify-between gap-3 mb-5">' +
    '<div><div class="section-label">Admin</div><h1 class="text-[24px] font-semibold tracking-tight">QR Labels · print preview</h1>' +
    '<p class="text-[12.5px] text-slate-600 mt-0.5">' + labels.length + ' label' + (labels.length !== 1 ? 's' : '') + ' · A4 sheet · 2 columns</p></div>' +
    '<button onclick="window.print()" class="btn btn-primary">' + iconSvg_('printer', 15) + 'Print labels</button>' +
  '</div>';

  if (!cfg.webAppBaseUrl) {
    toolbar += renderNoticePanel_({
      tone: 'amber', icon: 'warn',
      title: 'WEB_APP_BASE_URL not configured',
      body: 'Labels will show storage details but the QR codes will not link to the live web app. Set WEB_APP_BASE_URL before printing.',
    });
  }

  var grid = '<div class="label-sheet card-shell"><div class="grid grid-cols-2 gap-3">' +
    labels.map(function (l) {
      var url = (cfg.webAppBaseUrl || '') + '?room=' + encodeURIComponent(l.room) + '&loc=' + encodeURIComponent(l.id);
      return '<div class="qr-label' + (l.hazard ? ' hazard' : '') + '">' +
        '<div class="qr-glyph" aria-hidden="true"></div>' +
        '<div class="flex-1 min-w-0">' +
          '<div class="section-label" style="color:var(--blue)">VSA D&amp;T · ' + escape_(l.room) + '</div>' +
          '<div class="font-semibold text-[15px] mt-0.5 truncate">' + escape_(l.label) + '</div>' +
          '<div class="text-[11.5px] text-slate-500 truncate">' + escape_(l.specific || '') + '</div>' +
          '<div class="mt-1.5"><span class="pill pill-neutral pill-mono" style="background:#0f172a;color:#fff">' + escape_(l.id) + '</span></div>' +
          '<div class="text-[10.5px] text-slate-500 mt-1.5">Scan to view inventory</div>' +
          (l.hazard ? '<div class="text-[10.5px] text-red-700 font-semibold mt-0.5">' + iconSvg_('hazard', 10) + ' Chemical cabinet · PPE required</div>' : '') +
          (cfg.webAppBaseUrl ? '<div class="text-[9px] text-slate-400 mt-1 font-mono break-all">' + escape_(url) + '</div>' : '') +
        '</div>' +
      '</div>';
    }).join('') +
  '</div></div>';

  return toolbar + grid;
}

// ══════════════════════════════════════════════════════════════════════════
// ADMIN · DIAGNOSTICS
// ══════════════════════════════════════════════════════════════════════════

function renderDiagnosticsHtml_(data, cfg) {
  var checks = [
    { ok: !!cfg.spreadsheetId, label: 'SPREADSHEET_ID', value: cfg.spreadsheetId ? 'configured' : 'missing', critical: true,
      fix: 'Open the spreadsheet → D&T Inventory menu → Set App Config → paste the Sheet ID.' },
    { ok: !!cfg.webAppBaseUrl, label: 'WEB_APP_BASE_URL', value: cfg.webAppBaseUrl || 'missing',
      hint: cfg.webAppBaseUrl ? '' : 'Browsing works. QR generation and external absolute links are disabled.',
      fix: 'D&T Inventory menu → Set WEB_APP_BASE_URL → paste the deployed /exec URL.' },
    { ok: true, label: 'Inventory sheet', value: cfg.inventorySheetName + ' · ' + (data.totals.items + ' rows') },
    { ok: true, label: 'Required columns', value: 'all ' + CONFIG.REQUIRED_COLS.length + ' present' },
    { ok: data.headers.indexOf('QR Code Image') >= 0, label: 'QR Code Image (optional)', value: data.headers.indexOf('QR Code Image') >= 0 ? 'present' : 'missing',
      hint: data.headers.indexOf('QR Code Image') >= 0 ? '' : 'Optional. Only needed for in-sheet QR thumbnails.' },
  ];

  return '<header class="mb-5"><div class="section-label">Admin</div><h1 class="text-[26px] md:text-[30px] font-semibold tracking-tight">Diagnostics</h1><p class="text-[13.5px] text-slate-600 mt-1">Configuration health and how to fix any issues.</p></header>' +
    '<div class="grid grid-cols-1 lg:grid-cols-[1fr_340px] gap-6">' +
      '<section class="card divide-y divide-slate-100">' +
        checks.map(function (c) {
          return '<div class="p-4">' +
            '<div class="flex items-start gap-3">' +
              '<span class="w-2.5 h-2.5 rounded-full mt-1.5" style="background:' + (c.ok ? 'var(--em)' : c.critical ? 'var(--red)' : 'var(--amber)') + '"></span>' +
              '<div class="flex-1">' +
                '<div class="flex items-baseline justify-between gap-3">' +
                  '<h3 class="font-semibold">' + escape_(c.label) + '</h3>' +
                  '<span class="text-[11.5px] font-mono ' + (c.ok ? 'text-emerald-700' : 'text-amber-700') + '">' + escape_(c.value) + '</span>' +
                '</div>' +
                (c.hint ? '<p class="text-[12.5px] text-slate-600 mt-1">' + escape_(c.hint) + '</p>' : '') +
                (c.fix && !c.ok ? '<div class="mt-2 rounded-xl bg-slate-50 border border-slate-200 px-3 py-2 text-[12px] text-slate-700"><strong>How to fix · </strong>' + escape_(c.fix) + '</div>' : '') +
              '</div>' +
            '</div>' +
          '</div>';
        }).join('') +
      '</section>' +
      '<aside class="space-y-4" style="position:sticky;top:84px;align-self:start">' +
        '<section class="card p-4"><h3 class="font-semibold mb-2">What works without WEB_APP_BASE_URL</h3>' +
          '<ul class="text-[12.5px] text-slate-600 space-y-1.5 list-disc pl-4">' +
            '<li>In-app browsing &amp; navigation</li><li>View Mode and Tech Mode</li><li>Save updates back to the sheet</li><li>Admin readiness &amp; diagnostics</li>' +
          '</ul>' +
          '<h3 class="font-semibold mt-4 mb-2">What requires it</h3>' +
          '<ul class="text-[12.5px] text-slate-600 space-y-1.5 list-disc pl-4">' +
            '<li>Generated QR codes</li><li>Printable QR label sheet</li><li>External absolute links</li>' +
          '</ul>' +
        '</section>' +
        '<section class="card p-4"><h3 class="font-semibold mb-2">Spreadsheet menu reference</h3>' +
          '<div class="space-y-1 text-[12px] font-mono text-slate-600"><div>D&amp;T Inventory → Set App Config</div><div>D&amp;T Inventory → Set WEB_APP_BASE_URL</div><div>D&amp;T Inventory → Refresh QR Links</div><div>D&amp;T Inventory → Build QR Label Sheet</div><div>D&amp;T Inventory → Create Readiness Report</div></div>' +
        '</section>' +
      '</aside>' +
    '</div>';
}

// ══════════════════════════════════════════════════════════════════════════
// ERROR/EMPTY/NOTICE
// ══════════════════════════════════════════════════════════════════════════

function renderErrorStateHtml_(kind, ctx) {
  ctx = ctx || {};
  var V = {
    'no-spreadsheet': { tone: 'red', title: 'SPREADSHEET_ID is not set', body: 'The web app cannot read inventory until a spreadsheet is configured.', fix: 'Open the Google Sheet → D&T Inventory menu → Set App Config → paste the Sheet ID.' },
    'no-match':       { tone: 'amber', title: 'No matching storage location', body: 'This QR points to a storage row we cannot find: room ' + (ctx.room || '?') + ' · loc ' + (ctx.loc || '?') + ' did not match any inventory rows.', fix: 'Verify the QR was printed from this deployment, or run Refresh QR Links to rebuild the URLs.' },
  };
  var v = V[kind] || V['no-spreadsheet'];
  var border = v.tone === 'red' ? 'border-red-200 bg-red-50' : 'border-amber-200 bg-amber-50';
  var color = v.tone === 'red' ? 'text-red-900' : 'text-amber-900';
  return '<section class="card-shell p-6 ' + border + '">' +
    '<div class="flex items-start gap-3"><span class="' + color + '">' + iconSvg_('warn', 22) + '</span>' +
    '<div><h1 class="text-[22px] font-semibold tracking-tight ' + color + '">' + escape_(v.title) + '</h1>' +
    '<p class="text-[13.5px] mt-2 ' + color + '">' + escape_(v.body) + '</p></div></div></section>' +
    '<section class="card p-5 mt-4"><div class="section-label mb-2">How to fix</div><p class="text-[14px]">' + escape_(v.fix) + '</p>' +
    '<div class="mt-3 inline-block rounded-lg bg-slate-900 text-slate-100 px-3 py-2 font-mono text-[12px]">D&amp;T Inventory → Config Status / Diagnostics</div>' +
    '<div class="mt-4"><a class="btn btn-secondary" href="?admin=diagnostics">' + iconSvg_('info') + 'Open diagnostics</a></div></section>';
}

function renderEmptyState_(opts) {
  return '<section class="card p-8 mt-4 text-center">' +
    '<div class="mx-auto w-14 h-14 rounded-2xl bg-slate-100 text-slate-500 flex items-center justify-center">' + iconSvg_(opts.icon, 24) + '</div>' +
    '<h3 class="text-[16px] font-semibold mt-3">' + escape_(opts.title) + '</h3>' +
    '<p class="text-[13px] text-slate-600 mt-1">' + escape_(opts.body) + '</p>' +
    (opts.cta ? '<a href="' + opts.cta.href + '" class="btn ' + (opts.cta.primary ? 'btn-primary' : 'btn-secondary') + ' mt-4 mx-auto">' + escape_(opts.cta.label) + '</a>' : '') +
  '</section>';
}

function renderNoticePanel_(opts) {
  var tone = opts.tone || 'amber';
  var border = tone === 'red' ? 'border-red-200 bg-red-50 text-red-900' : 'border-amber-200 bg-amber-50 text-amber-900';
  return '<section class="card border ' + border + ' p-4 mb-4 flex items-start gap-3" role="status">' +
    iconSvg_(opts.icon || 'warn', 18) +
    '<div class="flex-1"><h3 class="font-semibold">' + escape_(opts.title) + '</h3>' +
    '<p class="text-[12.5px] mt-0.5">' + escape_(opts.body) + '</p></div>' +
    (opts.cta ? '<a href="' + opts.cta.href + '" class="btn btn-secondary text-[12.5px] py-2 px-3">' + escape_(opts.cta.label) + '</a>' : '') +
  '</section>';
}

// ══════════════════════════════════════════════════════════════════════════
// ICONS (inline SVG, no external assets)
// ══════════════════════════════════════════════════════════════════════════

function iconSvg_(name, size, style) {
  size = size || 16;
  style = style || '';
  var p = {
    home: 'M3 11l9-8 9 8M5 9v12h14V9',
    shield: 'M12 3l8 3v6c0 5-3.5 8.5-8 9-4.5-.5-8-4-8-9V6l8-3z',
    qr: 'M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h3v3M20 14v3M14 17v4M17 20h4',
    info: 'M12 8h.01M11 12h1v4h1M21 12a9 9 0 11-18 0 9 9 0 0118 0z',
    search: 'M21 21l-4.3-4.3M11 19a8 8 0 100-16 8 8 0 000 16z',
    edit: 'M12 20h9M16.5 3.5a2.121 2.121 0 113 3L7 19l-4 1 1-4z',
    save: 'M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2zM17 21v-8H7v8M7 3v5h8',
    box: 'M21 8l-9-5-9 5 9 5 9-5zM3 8v8l9 5 9-5V8M12 13v8',
    flask: 'M9 3h6M9 3v6L4 17a2 2 0 002 3h12a2 2 0 002-3l-5-8V3',
    wrench: 'M14.7 6.3a4 4 0 00-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 005.4-5.4l-3 3-2-2 3-3z',
    cpu: 'M9 3v3m6-3v3m-6 12v3m6-3v3M3 9h3m-3 6h3m12-6h3m-3 6h3M5 5h14v14H5z',
    hazard: 'M10.3 3.86l-8.32 14.42a2 2 0 001.74 3h16.56a2 2 0 001.74-3L13.7 3.86a2 2 0 00-3.4 0zM12 9v4M12 17h.01',
    warn: 'M10.3 3.86l-8.32 14.42a2 2 0 001.74 3h16.56a2 2 0 001.74-3L13.7 3.86a2 2 0 00-3.4 0zM12 9v4M12 17h.01',
    'arrow-left': 'M19 12H5M12 19l-7-7 7-7',
    printer: 'M6 9V2h12v7M6 18H4a2 2 0 01-2-2v-5a2 2 0 012-2h16a2 2 0 012 2v5a2 2 0 01-2 2h-2M6 14h12v8H6z',
  }[name] || '';
  return '<svg width="' + size + '" height="' + size + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block;vertical-align:-2px;' + style + '"><path d="' + p + '"/></svg>';
}

// ══════════════════════════════════════════════════════════════════════════
// CLIENT JS
// ══════════════════════════════════════════════════════════════════════════

function clientJs_() {
  return [
    '(function(){',
    '"use strict";',
    'var $=function(s,r){return (r||document).querySelector(s)};',
    'var $$=function(s,r){return Array.prototype.slice.call((r||document).querySelectorAll(s))};',

    // Landing search + room chips
    'var search=$("#storageSearch"),grid=$("#storageGrid"),empty=$("#storageEmpty"),chips=$("#roomChips");',
    'var activeRoom="__all";',
    'function applyLandingFilter(){if(!grid)return;var q=(search&&search.value||"").trim().toLowerCase();var any=false;',
    '  $$(".storage-card",grid).forEach(function(c){var hay=c.getAttribute("data-hay")||"";var room=c.getAttribute("data-room")||"";',
    '    var ok=(activeRoom==="__all"||room===activeRoom)&&(q===""||hay.indexOf(q)>=0);',
    '    c.style.display=ok?"":"none";if(ok)any=true});',
    '  if(empty)empty.classList.toggle("hidden",any);',
    '}',
    'if(search)search.addEventListener("input",applyLandingFilter);',
    'if(chips)chips.addEventListener("click",function(e){var b=e.target.closest("button[data-room]");if(!b)return;',
    '  $$("button",chips).forEach(function(x){x.classList.remove("active")});b.classList.add("active");activeRoom=b.getAttribute("data-room");applyLandingFilter()});',

    // Inventory page filters
    'var iSearch=$("#itemSearch"),iStatus=$("#itemStatus"),iList=$("#itemsList"),iEmpty=$("#itemsEmpty");',
    'function applyItemFilter(){if(!iList)return;var q=(iSearch&&iSearch.value||"").trim().toLowerCase();var s=iStatus&&iStatus.value||"__all";var any=false;',
    '  $$(".item-card",iList).forEach(function(c){var hay=c.getAttribute("data-hay")||"";var st=c.getAttribute("data-status")||"";',
    '    var ok=(s==="__all"||st===s)&&(q===""||hay.indexOf(q)>=0);c.style.display=ok?"":"none";if(ok)any=true});',
    '  if(iEmpty)iEmpty.classList.toggle("hidden",any);',
    '}',
    'if(iSearch)iSearch.addEventListener("input",applyItemFilter);',
    'if(iStatus)iStatus.addEventListener("change",applyItemFilter);',

    // Tech Mode editing
    'var saveBar=$("#saveBar"),saveBtn=$("#saveBtn"),resetBtn=$("#resetBtn"),changedCount=$("#changedCount"),toast=$("#toast");',
    'function rowDirty(row){var q=parseFloat(row.querySelector(".qty-val").value);var s=row.querySelector(".status-val").value;',
    '  return q!==parseFloat(row.getAttribute("data-orig-qty"))||s!==row.getAttribute("data-orig-status");}',
    'function refreshDirty(){if(!saveBar)return;var rows=$$(".tech-row");var changed=rows.filter(rowDirty);changed.forEach(function(r){r.classList.add("changed");',
    '    var oq=r.getAttribute("data-orig-qty"),os=r.getAttribute("data-orig-status");',
    '    var nq=r.querySelector(".qty-val").value,ns=r.querySelector(".status-val").value;',
    '    var msg=r.querySelector(".changed-msg"),mt=r.querySelector(".changed-msg-text");',
    '    var s=oq+" → "+nq;if(ns!==os)s+=" · "+os+" → "+ns;mt.textContent=" Changed · "+s;msg.classList.remove("hidden")});',
    '  rows.filter(function(r){return !rowDirty(r)}).forEach(function(r){r.classList.remove("changed");var m=r.querySelector(".changed-msg");if(m)m.classList.add("hidden")});',
    '  if(changed.length){changedCount.textContent=changed.length+" unsaved change"+(changed.length>1?"s":"");changedCount.style.color="#b45309";changedCount.style.fontWeight="600";saveBtn.disabled=false;resetBtn.disabled=false;}',
    '  else{changedCount.textContent="No changes yet";changedCount.style.color="";changedCount.style.fontWeight="";saveBtn.disabled=true;resetBtn.disabled=true;}',
    '}',
    '$$(".tech-row").forEach(function(r){',
    '  var qty=r.querySelector(".qty-val");',
    '  r.querySelector(".qty-dec").addEventListener("click",function(){qty.value=Math.max(0,(parseFloat(qty.value)||0)-1);refreshDirty()});',
    '  r.querySelector(".qty-inc").addEventListener("click",function(){qty.value=(parseFloat(qty.value)||0)+1;refreshDirty()});',
    '  qty.addEventListener("input",refreshDirty);qty.addEventListener("change",refreshDirty);',
    '  r.querySelector(".status-val").addEventListener("change",refreshDirty);',
    '});',
    'function showToast(kind,msg){if(!toast)return;toast.className="toast show "+kind;toast.textContent=msg;setTimeout(function(){toast.className="toast"},3200)}',
    'if(resetBtn)resetBtn.addEventListener("click",function(){$$(".tech-row").forEach(function(r){r.querySelector(".qty-val").value=r.getAttribute("data-orig-qty");r.querySelector(".status-val").value=r.getAttribute("data-orig-status")});refreshDirty()});',
    'if(saveBtn)saveBtn.addEventListener("click",function(){',
    '  var edits=$$(".tech-row").filter(rowDirty).map(function(r){return{sheetRow:Number(r.getAttribute("data-row")),qty:parseFloat(r.querySelector(".qty-val").value),status:r.querySelector(".status-val").value}});',
    '  if(!edits.length)return;saveBtn.disabled=true;saveBtn.innerHTML="Saving…";',
    '  if(typeof google==="undefined"||!google.script||!google.script.run){showToast("err","Save bridge unavailable. Open the deployed /exec URL.");saveBtn.disabled=false;saveBtn.innerHTML="Save updates";return}',
    '  google.script.run.withSuccessHandler(function(res){if(res&&res.ok){',
    '      $$(".tech-row").filter(rowDirty).forEach(function(r){r.setAttribute("data-orig-qty",r.querySelector(".qty-val").value);r.setAttribute("data-orig-status",r.querySelector(".status-val").value)});',
    '      refreshDirty();showToast("ok","Saved · "+res.applied+" row"+(res.applied!==1?"s":""))}else{showToast("err",(res&&res.error)||"Save failed")}saveBtn.innerHTML="Save updates"})',
    '    .withFailureHandler(function(err){showToast("err","Save failed · "+err.message);saveBtn.disabled=false;saveBtn.innerHTML="Save updates"})',
    '    .applyInventoryUpdates({room:window.__BUNDLE__&&window.__BUNDLE__.room,storageId:window.__BUNDLE__&&window.__BUNDLE__.storageId,edits:edits});',
    '});',
    '})();',
  ].join('\n');
}

// ══════════════════════════════════════════════════════════════════════════
// UTIL
// ══════════════════════════════════════════════════════════════════════════

function escape_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
