/* xpath-honmono-de-osu.mjs — ★XPath を ★本物の Chrome★ で 押す★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    FILTERXML の XPath は ★ブラウザの 物★を そのまま 使っています。
 *    ⇒ 試験は jsdom で 押していますが、★jsdom は 本物より 出来る事が 少ない★。
 *    ⇒ 実測 … `//*[name()='b']` は
 *        ・実Excel …… ★取れる★
 *        ・jsdom ……… ★投げる★（name() を 知らない）
 *        ・本物の Chrome … ★ここで 数えます★
 *    ⇒★★『試験が 赤い＝お客さんの 所でも 動かない』とは 限らない★★
 *    ⇒★でも ★確かめずに 「ブラウザなら 動く」と 書くのは 嘘★★
 *
 *  ★外へは 1回も 出ません★（手元の 中だけ）
 *
 *  使い方: node docs/measured/kansuu46/xpath-honmono-de-osu.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import http from 'node:http';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const CHROME = 'C:/Program Files/Google/Chrome/Application/chrome.exe';

const XML = "<r><b>1</b><b>2</b><b>3</b><c id='x'>A</c><c id='y'>B</c><d><e>ne</e></d></r>";
const 式たち = [
  '//b', '//c', '//c/@id', '/r/d/e', '//d/*',
  "//c[@id='y']", '//b[1]', '//b[last()]', '//b[position()>1]',
  "//*[name()='b']", "//*[local-name()='b']", '//e/text()', '//zzz', 'count(//b)', '//[',
];

const 画面 = `<!doctype html><meta charset="utf-8"><title>xpath</title><body>
<script src="/__spy.js"></script>`;
const SPY = `
var XML = ${JSON.stringify(XML)};
var 式 = ${JSON.stringify(式たち)};
var 出 = [];
var doc = new DOMParser().parseFromString(XML, 'text/xml');
for (var i = 0; i < 式.length; i++) {
  var xp = 式[i];
  try {
    var r = doc.evaluate(xp, doc, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
    var v = [];
    for (var j = 0; j < r.snapshotLength; j++) {
      var n = r.snapshotItem(j);
      v.push({ 型: n.nodeType, 字: (n.nodeType === 2 ? n.value : n.textContent) });
    }
    出.push({ 式: xp, 結: '取れた', 件: v.length, 中: v.slice(0, 3) });
  } catch (e) {
    出.push({ 式: xp, 結: '投げた', 訳: String(e && e.message).slice(0, 60) });
  }
}
var x = new XMLHttpRequest();
x.open('POST', '/report', true);
x.setRequestHeader('Content-Type', 'text/plain');
x.send(JSON.stringify(出));
`;

let 報告 = null;
const さば = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/report') {
    let b = ''; req.on('data', (c) => { b += c; });
    req.on('end', () => { try { 報告 = JSON.parse(b); } catch (e) { /* 読めない */ } res.end('ok'); });
    return;
  }
  if (req.url.split('?')[0] === '/__spy.js') {
    res.setHeader('Content-Type', 'application/javascript'); res.end(SPY); return;
  }
  res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(画面);
});
await new Promise((ok) => さば.listen(8821, '127.0.0.1', ok));

if (!fs.existsSync(CHROME)) { console.log('★Chrome が 無い★'); process.exit(1); }
const 仮 = fs.mkdtempSync(path.join(os.tmpdir(), 'xpath-'));
const ch = spawn(CHROME, [
  '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
  '--disable-background-networking', '--disable-component-update', '--disable-extensions',
  '--user-data-dir=' + 仮, '--virtual-time-budget=8000', '--dump-dom',
  'http://127.0.0.1:8821/',
], { stdio: ['ignore', 'ignore', 'ignore'] });
await new Promise((ok) => { ch.on('exit', ok); setTimeout(() => { try { ch.kill(); } catch (e) { /* もう 死んでいる */ } ok(); }, 30000); });
await new Promise((ok) => setTimeout(ok, 500));
さば.close();

const 行 = [];
const 言う = (s) => { 行.push(s); console.log(s); };
言う('# ★XPath を 本物の Chrome で 押した★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★外へは 1回も 出ていません（127.0.0.1 だけ）★');
言う('');
if (!報告) { 言う('★返事が 来ませんでした★'); process.exit(1); }
for (const r of 報告) {
  言う('  ' + r.式.padEnd(24) + ' … ' + (r.結 === '取れた'
    ? '取れた ' + r.件 + '件  ' + JSON.stringify(r.中.map((x) => x.字))
    : '★投げた★ ' + r.訳));
}
言う('');
const n = 報告.find((r) => r.式 === "//*[name()='b']");
言う('★jsdom が 投げる `name()` は … ' + (n && n.結 === '取れた' ? '★本物の Chrome では 取れる★' : '本物でも 投げる') + '★');
fs.writeFileSync(path.join(ここ, 'xpath-honmono-de-osu.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/kansuu46/xpath-honmono-de-osu.txt★');
