#!/usr/bin/env node
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const SCANNER_URL = process.env.SCANNER_URL || 'https://sunnydesigntech.github.io/dt-qr-inventory-system/scanner/';
const TARGET_URL = process.env.SCANNER_TARGET_URL || 'https://script.google.com/macros/s/DEPLOYMENT/exec';
const origin = new URL(SCANNER_URL).origin;
const testUrl = new URL(SCANNER_URL);
testUrl.searchParams.set('target', TARGET_URL);
testUrl.searchParams.set('source', 'camera-smoke');

async function loadPlaywright() {
  try {
    return await import('playwright');
  } catch (err) {
    try {
      const require = createRequire(import.meta.url);
      return require('playwright');
    } catch (requireErr) {}

    const searchRoots = [
      String(process.env.PLAYWRIGHT_NODE_MODULES || ''),
      ...String(process.env.NODE_PATH || '').split(':').filter(Boolean),
    ].filter(Boolean);
    for (const root of searchRoots) {
      const candidate = join(root, 'playwright', 'index.js');
      if (existsSync(candidate)) {
        const mod = await import(pathToFileURL(candidate).href);
        return mod.default || mod;
      }
    }

    console.error('Playwright is not available. Install Playwright locally or run with PLAYWRIGHT_NODE_MODULES=/path/to/node_modules or NODE_PATH=/path/to/node_modules.');
    process.exit(2);
  }
}

async function launchChromium(chromium) {
  const launchOptions = {
    headless: true,
    args: [
      '--use-fake-ui-for-media-stream',
      '--use-fake-device-for-media-stream',
      '--no-sandbox'
    ]
  };
  try {
    return await chromium.launch({ ...launchOptions, channel: 'chrome' });
  } catch (err) {
    return chromium.launch(launchOptions);
  }
}

const { chromium } = await loadPlaywright();
const browser = await launchChromium(chromium);
const context = await browser.newContext({
  permissions: ['camera'],
  viewport: { width: 390, height: 844 },
  isMobile: true,
  hasTouch: true
});
await context.grantPermissions(['camera'], { origin });
const page = await context.newPage();

try {
  await page.goto(testUrl.toString(), { waitUntil: 'domcontentloaded', timeout: 30000 });
  await assert.doesNotReject(async () => page.locator('#startBtn').waitFor({ state: 'visible', timeout: 10000 }));
  await page.locator('#startBtn').click();
  await page.waitForFunction(() => {
    const status = document.querySelector('#status');
    const video = document.querySelector('#video');
    const text = status ? status.textContent || '' : '';
    return /Scanning/i.test(text) || !!(video && video.srcObject);
  }, null, { timeout: 15000 });

  const status = await page.locator('#status').textContent();
  const href = await page.evaluate(() => location.href);
  assert.match(status || '', /Scanning|QR/i);
  assert.doesNotMatch(href, /mode=tech/i);
  console.log(JSON.stringify({
    ok: true,
    scannerUrl: SCANNER_URL,
    targetUrl: TARGET_URL,
    status: status || '',
    note: 'Top-level scanner started camera stream with fake camera permission.'
  }, null, 2));
} finally {
  await browser.close();
}
