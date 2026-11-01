// Server uji lokal: menjalankan gas/Code.gs apa adanya dengan tiruan SpreadsheetApp di memori.
// Pemakaian: node tools/gas-dev-server.js [port=8787] [--tanpa-seed]
//   --tanpa-seed: data/seed.json disembunyikan, meniru aplikasi di hosting publik (jadwal dari Sheet)
//   Aplikasi: http://localhost:8787/   URL sinkron: http://localhost:8787/exec   Kode akses: DEV
// Data hanya di memori (hilang saat server dimatikan) – khusus untuk uji, bukan produksi.
'use strict';
const fs = require('fs');
const http = require('http');
const path = require('path');
const vm = require('vm');

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  getLastRow() { return this.data.length; }
  getMaxRows() { return 1000; }
  setFrozenRows() { return this; }
  appendRow(r) { this.data.push(r.slice()); return this; }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this;
    const range = {
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (sh.data[r - 1 + i] || [])[c - 1 + j] ?? '')),
      setValues: v => { v.forEach((row, i) => { const t = sh.data[r - 1 + i] ||= []; row.forEach((x, j) => { t[c - 1 + j] = x; }); }); return range; },
      setNumberFormat: () => range, setFontWeight: () => range,
    };
    return range;
  }
}
const sheets = {};
const props = {};
const ctx = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)) }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = String(v); } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) },
  Utilities: { getUuid: () => 'dev00000-0000' },
  Logger: { log: m => console.log(m) },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.gs'), 'utf8'), ctx);
props.TEAM_KEY = 'DEV';
ctx.setup();

const port = +process.argv[2] || 8787;
const noSeed = process.argv.includes('--tanpa-seed');
http.createServer((req, res) => {
  const cors = { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' };
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/dump') { res.writeHead(200, cors); return res.end(JSON.stringify(Object.fromEntries(Object.entries(sheets).map(([k, s]) => [k, s.data])), null, 1)); }
  if (req.method === 'GET' && url.pathname !== '/exec') { // sajikan juga aplikasinya: http://localhost:PORT/
    const root = path.join(__dirname, '..', 'app');
    const f = path.join(root, path.normalize(decodeURIComponent(url.pathname)).replace(/^[\\/]+/, '') || 'index.html');
    if (!f.startsWith(root) || (noSeed && f.endsWith('seed.json')) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end(); }
    const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.webmanifest': 'application/manifest+json', '.png': 'image/png' };
    res.writeHead(200, { 'Content-Type': (types[path.extname(f)] || 'application/octet-stream') + '; charset=utf-8' });
    return fs.createReadStream(f).pipe(res);
  }
  if (req.method === 'GET') {
    const out = ctx.doGet({ parameter: Object.fromEntries(url.searchParams) });
    res.writeHead(200, cors); return res.end(out.s);
  }
  let body = '';
  req.on('data', c => { body += c; });
  req.on('end', () => {
    // Apps Script menolak preflight; klien harus kirim "simple request" (text/plain). Tiru dengan menolak content-type lain.
    if (!/^text\/plain/.test(req.headers['content-type'] || '')) { res.writeHead(415, cors); return res.end('{"ok":false,"error":"harus text/plain"}'); }
    res.writeHead(200, cors);
    res.end(ctx.doPost({ postData: { contents: body } }).s);
  });
}).listen(port, () => console.log(`GAS dev server http://localhost:${port}/exec  (kode: DEV)`));
