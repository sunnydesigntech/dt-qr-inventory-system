import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

// Local smoke tests only. These mirror small pure helper decisions without
// requiring Apps Script services, live Google Sheets, or deployment access.

function safeExternalUrl(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw || !/^https?:\/\//i.test(raw)) return '';
  try {
    const parsed = new URL(raw);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:' ? parsed.href : '';
  } catch {
    return '';
  }
}

function buildLocationUrl(baseUrl, room, loc) {
  const cleanBase = String(baseUrl || '').trim().replace(/[?#].*$/, '').replace(/\?+$/, '');
  return cleanBase + '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function resolveWebAppBaseUrl({ propertyUrl = '', defaultUrl = '', runtimeUrl = '' } = {}) {
  const clean = (value) => String(value || '').trim().replace(/[?#].*$/, '').replace(/\?+$/, '');
  return clean(propertyUrl) || clean(defaultUrl) || clean(runtimeUrl);
}

function buildLocationHref(baseUrl, room, loc, mode) {
  const href = baseUrl ? buildLocationUrl(baseUrl, room, loc) : '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
  return mode === 'tech' ? href + '&mode=tech' : href;
}

const DEPLOYED_EXEC = 'https://script.google.com/macros/s/DEPLOYMENT/exec';

function routeUrlKey(url) {
  return (url.protocol + '//' + url.host + url.pathname).replace(/\/+$/, '').toLowerCase();
}

function parseInventoryQrUrl(input, allowedBase = DEPLOYED_EXEC) {
  const text = String(input == null ? '' : input).trim();
  if (!text) return { room: '', loc: '', modeStripped: false, error: '' };
  let params = null;
  let parsed = null;
  const configuredKey = routeUrlKey(new URL(allowedBase));
  if (text.startsWith('?')) {
    params = new URLSearchParams(text.slice(1));
  } else if (/^(room|loc|mode)=/i.test(text)) {
    params = new URLSearchParams(text);
  } else if (/^\/?exec\?/i.test(text)) {
    const rel = text.charAt(0) === '/' ? text : '/' + text;
    parsed = new URL(rel, allowedBase);
    params = parsed.searchParams;
  } else if (/^[a-z][a-z0-9+.-]*:/i.test(text) || /^\/\//.test(text)) {
    if (/^\/\//.test(text)) return { room: '', loc: '', modeStripped: false, error: 'external' };
    try {
      parsed = new URL(text);
    } catch {
      return { room: '', loc: '', modeStripped: false, error: 'malformed' };
    }
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      return { room: '', loc: '', modeStripped: false, error: 'scheme' };
    }
    if (routeUrlKey(parsed) !== configuredKey) {
      return { room: '', loc: '', modeStripped: false, error: 'external' };
    }
    params = parsed.searchParams;
  }
  if (!params) return { room: '', loc: '', modeStripped: false, error: '' };
  const room = String(params.get('room') || '').trim();
  const loc = String(params.get('loc') || '').trim();
  return { room, loc, modeStripped: !!params.get('mode'), error: room && loc ? '' : 'missing' };
}

function fallbackRouteHref(room, loc) {
  return '?room=' + encodeURIComponent(room) + '&loc=' + encodeURIComponent(loc);
}

function scannerSelfTestViewUrl(input, allowedBase = DEPLOYED_EXEC) {
  const parsed = parseInventoryQrUrl(input, allowedBase);
  if (parsed.error) return { href: '', error: parsed.error, modeStripped: parsed.modeStripped };
  return {
    href: buildLocationUrl(allowedBase, parsed.room, parsed.loc),
    error: '',
    modeStripped: parsed.modeStripped
  };
}

function buildStandaloneScannerUrl(currentUrl) {
  const url = new URL(currentUrl);
  ['room', 'loc', 'mode', 'admin', 'printer', 'label'].forEach((key) => url.searchParams.delete(key));
  url.searchParams.set('scanner', '1');
  url.searchParams.set('topscanner', '1');
  return url.href;
}

function buildExternalScannerLaunchUrl(scannerUrl, targetUrl) {
  const url = new URL(scannerUrl);
  url.searchParams.set('target', targetUrl);
  url.searchParams.set('source', 'apps-script');
  return url.href;
}

function stripScannerParams(currentUrl, room, loc) {
  const url = new URL(currentUrl);
  ['room', 'loc', 'mode', 'admin', 'printer', 'scanner', 'scan', 'topscanner'].forEach((key) => url.searchParams.delete(key));
  url.searchParams.set('room', room);
  url.searchParams.set('loc', loc);
  return url.href;
}

function isPlaceholderRow(row) {
  const value = (key) => String(row[key] == null ? '' : row[key]).trim();
  const explicit = value('isPlaceholder').toLowerCase();
  if (['true', 'yes', 'y', '1', 'placeholder'].includes(explicit)) return true;
  return value('category').toLowerCase() === 'storage' &&
    value('remarks').toLowerCase().includes('placeholder row for qr/location page') &&
    !value('itemId') &&
    !value('itemName');
}

function isInventoryItemRow(row) {
  if (isPlaceholderRow(row)) return false;
  return !!(String(row.itemId || '').trim() || String(row.itemName || '').trim());
}

function mutationAllowed(config, token) {
  const cfg = config || {};
  if (!cfg.configured) return false;
  if (cfg.disabled) return true;
  return !!token;
}

function normalizeIdentity(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

function rowIdentityMatches(row, expected) {
  const actualId = String(row.itemId || '').trim();
  const expectedId = String(expected.itemId || '').trim();
  if (actualId || expectedId) return !!actualId && !!expectedId && normalizeIdentity(actualId) === normalizeIdentity(expectedId);
  return normalizeIdentity(row.itemName) === normalizeIdentity(expected.itemName) &&
    normalizeIdentity(row.category) === normalizeIdentity(expected.category) &&
    String(row.qty).trim() === String(expected.expectedQty).trim() &&
    normalizeIdentity(row.status) === normalizeIdentity(expected.expectedStatus);
}

assert.equal(safeExternalUrl('https://example.com/sds.pdf'), 'https://example.com/sds.pdf');
assert.equal(safeExternalUrl('http://example.com/order'), 'http://example.com/order');
assert.equal(safeExternalUrl('javascript:alert(1)'), '');
assert.equal(safeExternalUrl('data:text/html,hello'), '');
assert.equal(safeExternalUrl('vbscript:msgbox(1)'), '');
assert.equal(safeExternalUrl('ftp://example.com/file'), '');
assert.equal(safeExternalUrl('//example.com/path'), '');
assert.equal(safeExternalUrl('/relative/path'), '');
assert.equal(safeExternalUrl('https://'), '');
assert.equal(resolveWebAppBaseUrl({ runtimeUrl: 'https://script.google.com/macros/s/DEPLOYMENT/exec?scanner=1' }), 'https://script.google.com/macros/s/DEPLOYMENT/exec');
assert.equal(resolveWebAppBaseUrl({ propertyUrl: 'https://configured.example/exec', runtimeUrl: 'https://runtime.example/exec' }), 'https://configured.example/exec');

const vppViewUrl = buildLocationUrl('https://script.google.com/macros/s/DEPLOYMENT/exec', 'V++', 'Maker Bench 1');
assert.match(vppViewUrl, /room=V%2B%2B/);
assert.doesNotMatch(vppViewUrl, /mode=tech/);
assert.match(buildLocationHref('', '419A', '419A-FCU-01', 'view'), /^\?room=419A&loc=419A-FCU-01$/);
assert.match(buildLocationHref('', '419A', '419A-FCU-01', 'tech'), /mode=tech$/);

assert.deepEqual(parseInventoryQrUrl(`${DEPLOYED_EXEC}?room=419A&loc=419A-FCU-01`), {
  room: '419A',
  loc: '419A-FCU-01',
  modeStripped: false,
  error: ''
});
assert.deepEqual(parseInventoryQrUrl('/exec?room=419A&loc=419A-CAB-01'), {
  room: '419A',
  loc: '419A-CAB-01',
  modeStripped: false,
  error: ''
});
assert.deepEqual(parseInventoryQrUrl('room=V%2B%2B&loc=Rack%20trolley'), {
  room: 'V++',
  loc: 'Rack trolley',
  modeStripped: false,
  error: ''
});
const pastedTech = parseInventoryQrUrl(`${DEPLOYED_EXEC}?room=419A&loc=419A-FCU-01&mode=tech`);
assert.equal(pastedTech.room, '419A');
assert.equal(pastedTech.loc, '419A-FCU-01');
assert.equal(pastedTech.modeStripped, true);
assert.doesNotMatch(fallbackRouteHref(pastedTech.room, pastedTech.loc), /mode=tech/);
assert.equal(parseInventoryQrUrl('javascript:alert(1)').error, 'scheme');
assert.equal(parseInventoryQrUrl('data:text/html,hello').error, 'scheme');
assert.equal(parseInventoryQrUrl('https://example.com/exec?room=419A&loc=419A-FCU-01').error, 'external');
assert.equal(parseInventoryQrUrl('//example.com/exec?room=419A&loc=419A-FCU-01').error, 'external');
assert.equal(
  scannerSelfTestViewUrl(`${DEPLOYED_EXEC}?room=419A&loc=419A-FCU-01&mode=tech`).href,
  `${DEPLOYED_EXEC}?room=419A&loc=419A-FCU-01`
);
assert.equal(scannerSelfTestViewUrl(`${DEPLOYED_EXEC}?room=419A&loc=419A-FCU-01&mode=tech`).modeStripped, true);
assert.equal(scannerSelfTestViewUrl('room=V%2B%2B&loc=Maker%20Bench%201').href, `${DEPLOYED_EXEC}?room=V%2B%2B&loc=Maker%20Bench%201`);
assert.equal(scannerSelfTestViewUrl('https://example.com/exec?room=419A&loc=419A-FCU-01').error, 'external');

const topScanner = buildStandaloneScannerUrl('https://n-abc-script.googleusercontent.com/userCodeAppPanel?room=419A&loc=419A-FCU-01&mode=tech&scanner=1');
assert.match(topScanner, /scanner=1/);
assert.match(topScanner, /topscanner=1/);
assert.doesNotMatch(topScanner, /mode=tech/);
assert.doesNotMatch(topScanner, /loc=419A-FCU-01/);
assert.equal(
  stripScannerParams('https://n-abc-script.googleusercontent.com/userCodeAppPanel?scanner=1&topscanner=1', 'V++', 'Maker Bench 1'),
  'https://n-abc-script.googleusercontent.com/userCodeAppPanel?room=V%2B%2B&loc=Maker+Bench+1'
);
const externalScannerLaunch = buildExternalScannerLaunchUrl(
  'https://sunnydesigntech.github.io/dt-qr-inventory-system/scanner/',
  DEPLOYED_EXEC
);
assert.match(externalScannerLaunch, /^https:\/\/sunnydesigntech\.github\.io\/dt-qr-inventory-system\/scanner\/?\?/);
assert.match(externalScannerLaunch, /target=https%3A%2F%2Fscript\.google\.com%2Fmacros%2Fs%2FDEPLOYMENT%2Fexec/);
assert.doesNotMatch(externalScannerLaunch, /mode=tech/);

const scannerHtml = readFileSync(new URL('../scanner/index.html', import.meta.url), 'utf8');
assert.match(scannerHtml, /Camera scanner/);
assert.match(scannerHtml, /View Mode/);
assert.match(scannerHtml, /mode=tech|Update Mode is never opened by this scanner/);
assert.match(scannerHtml, /No-camera route self-test/);
assert.match(scannerHtml, /data-selftest-index/);
assert.match(scannerHtml, /Passed: resolves to View Mode/);
assert.doesNotMatch(scannerHtml, /SPREADSHEET_ID|AKfyc|1GqK9/);

const appScript = readFileSync(new URL('../app_script.html', import.meta.url), 'utf8');
assert.match(appScript, /function openScanner\(\)[\s\S]*configuredExternalScannerUrl\(\)[\s\S]*openStandaloneScanner\(\)/);
assert.match(appScript, /function openStandaloneScanner\(\)[\s\S]*return true;[\s\S]*return false;/);

const rows = [
  { itemId: '', itemName: '', category: 'Storage', remarks: 'Placeholder row for QR/location page', isPlaceholder: 'TRUE' },
  { itemId: 'ITEM-1', itemName: 'Safety Goggles', category: 'Tools', remarks: '', isPlaceholder: '' }
];
assert.equal(rows.filter(isInventoryItemRow).length, 1);

assert.equal(mutationAllowed({ configured: false }, ''), false);
assert.equal(mutationAllowed({ configured: true, disabled: false }, ''), false);
assert.equal(mutationAllowed({ configured: true, disabled: false }, 'token'), true);
assert.equal(mutationAllowed({ configured: true, disabled: true }, ''), true);

assert.equal(rowIdentityMatches(
  { itemId: '', itemName: 'Glue Stick', category: 'Consumables', qty: 1, status: 'Good' },
  { itemName: 'Glue Stick', category: 'Consumables', expectedQty: 1, expectedStatus: 'Good' }
), true);
assert.equal(rowIdentityMatches(
  { itemId: '', itemName: 'Glue Stick', category: 'Consumables', qty: 2, status: 'Good' },
  { itemName: 'Glue Stick', category: 'Consumables', expectedQty: 1, expectedStatus: 'Good' }
), false);

console.log('Smoke tests passed');
