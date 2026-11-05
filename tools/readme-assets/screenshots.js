#!/usr/bin/env node
/*
 * Takes the README screenshots of viewer.html into assets/readme/.
 * Office.js is replaced by a small stub, and https://dashboard.example.com
 * is served from demo-dashboard.html, so no network access is needed.
 *
 * Needs Playwright with Chromium:  node tools/readme-assets/screenshots.js
 */
'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.join(__dirname, '..', '..');
const OUT = path.join(ROOT, 'assets', 'readme');
const DEMO = fs.readFileSync(path.join(__dirname, 'demo-dashboard.html'), 'utf8');
const DEMO_URL = 'https://dashboard.example.com/q3';

function officeStub(savedUrl) {
  return `window.Office = {
    PlatformType: { OfficeOnline: 'OfficeOnline' },
    AsyncResultStatus: { Failed: 'failed', Succeeded: 'succeeded' },
    EventType: { DialogEventReceived: 'dialogEvent' },
    context: {
      platform: 'PC',
      ui: { displayDialogAsync: function (u, o, cb) { cb({ status: 'failed', error: { code: 12007 } }); } },
      document: { settings: { get: function () { return ${JSON.stringify(savedUrl)}; }, set: function () {}, saveAsync: function (cb) { if (cb) cb({ status: 'succeeded' }); } } }
    },
    onReady: function (cb) { setTimeout(cb, 30); }
  };`;
}

function serve() {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      const file = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]));
      if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
        res.writeHead(404); res.end(); return;
      }
      res.writeHead(200, { 'content-type': file.endsWith('.js') ? 'text/javascript' : 'text/html' });
      res.end(fs.readFileSync(file));
    }).listen(0, () => resolve(server));
  });
}

(async () => {
  const server = await serve();
  const base = `http://localhost:${server.address().port}/viewer.html`;
  const browser = await chromium.launch();

  async function open(savedUrl) {
    const page = await browser.newPage({ viewport: { width: 800, height: 450 }, deviceScaleFactor: 2 });
    await page.route('**/office.js', (r) => r.fulfill({ contentType: 'text/javascript', body: officeStub(savedUrl) }));
    await page.route('https://dashboard.example.com/**', (r) => r.fulfill({ contentType: 'text/html', body: DEMO }));
    await page.goto(base);
    await page.waitForTimeout(600);
    return page;
  }
  const shot = (page, name) => page.screenshot({ path: path.join(OUT, name) }).then(() => console.log('  ' + name));

  let page = await open(null);
  await shot(page, 'viewer-welcome.png');

  page = await open(DEMO_URL);
  await shot(page, 'viewer-loaded.png');
  await page.waitForTimeout(2400);
  await shot(page, 'viewer-hint.png');
  await page.click('#config .btn-hide');
  await page.waitForTimeout(500);
  await shot(page, 'viewer-presenting.png');

  page = await open("javascript:alert('hi')");
  await shot(page, 'viewer-safety.png');

  page = await open(DEMO_URL);
  await page.click('#btnOpenWindow');
  await page.waitForTimeout(300);
  await shot(page, 'viewer-open-link.png');

  await browser.close();
  server.close();
})();
