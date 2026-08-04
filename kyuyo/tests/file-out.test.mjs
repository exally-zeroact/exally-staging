/* file-out.test.mjs — ★ファイルの渡し方（種類・共有シート・落とす）を固定する★
 *
 * なぜ必要か（2026-08-04・司さんの実機）:
 *   iPhone に Excel が入っているのに、落としたファイルを開けなかった。
 *   原因は端末ではなく、★種類を application/octet-stream で落としていた★こと。
 *   さらに ★共有シート(navigator.share)を1箇所も使っていなかった★＝
 *   iPhoneでファイルを渡す普通のやり方をしていなかった。
 *
 * ここで固定すること:
 *   ① 拡張子から種類が必ず決まる（xlsx/csv/txt…）。★分からない拡張子は落とさない★
 *   ② 指で触る端末(pointer: coarse)では共有シートに渡す＝「Excelで開く」が並ぶ
 *   ③ ★PCでは今までどおり落ちる★（共有シートに行かない）
 *      — 実測でここを一度壊した: デスクトップChromeも canShare は true を返すので、
 *        機能だけで判定すると PCでファイルが落ちなくなる（退行）。だから pointer で分ける。
 *   ④ 客が共有シートを閉じただけの時は、エラーにしない
 *
 * 使い方: node tests/file-out.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = fs.readFileSync(path.join(ROOT, '../js/file-out.js'), 'utf8');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。'); process.exit(1); }

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const TA = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };

/* 端末のふりをして FileOut を読み込む。coarse=指で触る端末 / share=共有シートが使えるか */
function load({ coarse = false, share = true } = {}) {
  const dom = new JSDOM('<!doctype html><body></body>', { runScripts: 'dangerously', url: 'http://localhost/' });
  const win = dom.window;
  win.matchMedia = (q) => ({ matches: q === '(pointer: coarse)' ? coarse : false });
  const shared = [], downloaded = [];
  if (share) {
    win.navigator.canShare = (d) => !!(d && d.files && d.files.length);
    win.navigator.share = (d) => { shared.push({ name: d.files[0].name, type: d.files[0].type }); return Promise.resolve(); };
  } else { win.navigator.canShare = undefined; win.navigator.share = undefined; }
  win.URL.createObjectURL = () => 'blob:test';
  win.URL.revokeObjectURL = () => { };
  const origCreate = win.document.createElement.bind(win.document);
  win.document.createElement = (t) => {
    const el = origCreate(t);
    if (t === 'a') el.click = () => downloaded.push({ name: el.download });
    return el;
  };
  const s = win.document.createElement('script');
  s.textContent = SRC;
  win.document.body.appendChild(s);
  return { win, FileOut: win.FileOut, shared, downloaded };
}

console.log('\n[file-out] ファイルの渡し方（種類・共有シート・落とす）');

T('① 拡張子から種類が決まる', () => {
  const { FileOut } = load();
  eq(FileOut.mimeOf('a.xlsx'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx');
  eq(FileOut.mimeOf('a.csv'), 'text/csv', 'csv');
  eq(FileOut.mimeOf('furikomi_2026-08.txt'), 'text/plain', 'txt（全銀ファイル）');
  eq(FileOut.mimeOf('a.pdf'), 'application/pdf', 'pdf');
});

T('★① 分からない拡張子は種類を作らない（octet-stream に落とさない）', () => {
  const { FileOut } = load();
  eq(FileOut.mimeOf('nazo.bin'), null, '知らない拡張子');
  eq(FileOut.mimeOf('拡張子なし'), null, '拡張子なし');
});

await TA('★① 分からない拡張子は【渡さずに止める】（iPhoneで開けないファイルを作らない）', async () => {
  const { FileOut, downloaded, shared } = load();
  let msg = null;
  try { await FileOut.deliver(new Uint8Array([1]), 'nazo.bin'); } catch (e) { msg = e.message; }
  ok(msg && /種類が分かりません/.test(msg), '止めた理由を言っている: ' + msg);
  eq(downloaded.length, 0, 'ファイルを作っていない');
  eq(shared.length, 0, '共有もしていない');
});

await TA('★② 指で触る端末では共有シートに渡す（「Excelで開く」が並ぶ）', async () => {
  const { FileOut, shared, downloaded } = load({ coarse: true, share: true });
  const r = await FileOut.deliver(new Uint8Array([80, 75, 3, 4]), '給与明細_2026-08.xlsx');
  eq(r.how, 'share', '共有シートに行った');
  eq(shared.length, 1, '共有シートに1回渡した');
  eq(shared[0].name, '給与明細_2026-08.xlsx', 'ファイル名');
  eq(shared[0].type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '★種類が正しい（これが無いとiPhoneで開けない）');
  eq(downloaded.length, 0, '共有シートに行ったので落としていない');
});

await TA('★③ PCでは今までどおり落ちる（共有シートに行かない）', async () => {
  // ★デスクトップChromeも canShare は true を返す。機能だけで判定すると
  //   PCでファイルが落ちなくなる＝退行。実際に一度そうなったので、ここで固定する。
  const { FileOut, shared, downloaded } = load({ coarse: false, share: true });
  const r = await FileOut.deliver(new Uint8Array([80, 75, 3, 4]), '給与明細_2026-08.xlsx');
  eq(r.how, 'download', '落ちた');
  eq(downloaded.length, 1, '1回落とした');
  eq(downloaded[0].name, '給与明細_2026-08.xlsx', 'ファイル名');
  eq(shared.length, 0, '★共有シートに行っていない');
});

await TA('③ 共有シートが無い端末でも落ちる', async () => {
  const { FileOut, downloaded } = load({ coarse: true, share: false });
  const r = await FileOut.deliver(new Uint8Array([1]), 'furikomi_2026-08.txt');
  eq(r.how, 'download', '落ちた');
  eq(downloaded.length, 1, '1回落とした');
});

await TA('★④ 客が共有シートを閉じただけならエラーにしない（何も起きなくてよい）', async () => {
  const { win, FileOut, downloaded } = load({ coarse: true, share: true });
  win.navigator.share = () => { const e = new Error('cancel'); e.name = 'AbortError'; return Promise.reject(e); };
  const r = await FileOut.deliver(new Uint8Array([1]), 'a.xlsx');
  eq(r.how, 'cancel', '閉じただけ');
  eq(downloaded.length, 0, '勝手に落とし直さない');
});

await TA('共有シートが失敗した時は、黙って諦めず落とす方に切り替える', async () => {
  const { win, FileOut, downloaded } = load({ coarse: true, share: true });
  win.navigator.share = () => Promise.reject(new Error('なぜか失敗'));
  const r = await FileOut.deliver(new Uint8Array([1]), 'a.xlsx');
  eq(r.how, 'download', '落とす方に切り替えた');
  eq(downloaded.length, 1, '落ちた');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
