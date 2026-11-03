/* zip-tasu.test.mjs — ★包みに 部品を 足せるか★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    司さん「★全部 保存しろや、断る 理由が なんか あるんか★」（ア）
 *    `lib/hairanai.js` の 頭に こう 書いて あります（2026-09-06 実測）
 *      「受け取った ファイルを 保存し直す 道は ★元の zip を 持ったまま 値だけ 書き戻す★」
 *      「⇒★でも うちで 足した 物は 書き出す 先に 入らない★」
 *    ⇒★「出来ない から」とは 1文字も 書いて いません★
 *    ⇒★足りないのは `ZipSurgeon` の 「足す」口 1つ★ でした。
 *
 *  ★★この 門が 守る 事★★
 *    ①足した 部品が ★読み戻せる★
 *    ②★元の 部品が 1つも 減らない★
 *    ③★元の 部品の 中身が 1バイトも 変わらない★
 *      ＝★これが 一番 大事★（判子・罫線・グラフは 元の 部品の 中に 在ります）
 *    ④★同じ 名前は 足せない★（★黙って 上書きしない★）
 *    ⑤★同じ 材料から 同じ 包みが 出る★（★時刻を 入れない★）
 *
 *  ★★見て いない 事★★（★書かない 見張りは「全部 守った」と 読まれる★）
 *    ・★実Excel が この 包みを どう 開くかは 測れません★（COM が 要る＝経営者1 の 持ち場）
 *    ・★`[Content_Types].xml` や rels を 直すのは 呼ぶ 側の 仕事★
 *      ＝この 門は ★包みに 入るか★ だけ 見ます
 *    ・★`.xlsb` では 試して いません★
 *
 *  使い方: node tests/zip-tasu.test.mjs
 *          node tests/zip-tasu.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const Z = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

let pass = 0, fail = 0;
const NL = String.fromCharCode(10);
const T = (n, f) => {
  try { f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};
const 待 = async (n, f) => {
  try { await f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};

const 材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');
const 中 = fs.existsSync(材料道) ? fs.readFileSync(材料道) : null;

console.log('');
console.log('[zip-tasu] ★包みに 部品を 足せるか★');

T('★材料が 在る（空振りして いない）★', () => {
  if (!中) throw new Error('★材料が 無い★ ' + 材料道);
  const h = crypto.createHash('sha256').update(中).digest('hex');
  if (h !== 'affd42e9f5c6de9d5565c8ffd623ffd2e37a7ba5ca8820fdd5c971b15ff93053') {
    throw new Error('★材料が 入れ替わって います★ ' + h);
  }
});

T('★`add` と `addText` の 口が 在る★', () => {
  const z = Z.read(new Uint8Array(中));
  for (const k of ['add', 'addText', 'replace', 'remove', 'build']) {
    if (typeof z[k] !== 'function') throw new Error('★' + k + ' が 在りません★');
  }
});

/* ★元の 包みの 中身を 全部 控えて おきます★（★1バイトも 変わらない★を 見る為） */
const 元z = Z.read(new Uint8Array(中));
const 元の名 = 元z.names().slice();
const 元の中身 = {};
for (const n of 元の名) 元の中身[n] = await 元z.bytes(n);

const 足した = await (async () => {
  const z = Z.read(new Uint8Array(中));
  z.addText('xl/worksheets/sheet2.xml', '<x>tameshi</x>');
  const r = await z.build();
  return Z.read(r.bytes);
})();

await 待('★★足した 部品が 読み戻せる★★', async () => {
  const t = await 足した.text('xl/worksheets/sheet2.xml');
  if (t !== '<x>tameshi</x>') throw new Error('★中身が ' + t + '★');
});

T('★★元の 部品が 1つも 減って いない★★', () => {
  const 無 = 元の名.filter((n) => !足した.has(n));
  if (無.length) throw new Error('★' + 無.length + '本 消えました★ ' + 無.join(' '));
  if (足した.names().length !== 元の名.length + 1) {
    throw new Error('★' + 足した.names().length + '本★（' + (元の名.length + 1) + '本 の はず）');
  }
});

await 待('★★★元の 部品の 中身が 1バイトも 変わって いない★★★'
  + '（★判子・罫線・グラフは この 中★）', async () => {
  const 違 = [];
  for (const n of 元の名) {
    const a = 元の中身[n];
    const b = await 足した.bytes(n);
    if (a.length !== b.length) { 違.push(n + ' 大きさ ' + a.length + ' ⇒ ' + b.length); continue; }
    for (let i = 0; i < a.length; i++) {
      if (a[i] !== b[i]) { 違.push(n + ' ' + i + 'バイト目'); break; }
    }
  }
  if (違.length) throw new Error('★' + 違.length + '本 変わりました★' + NL + '       ' + 違.join(NL + '       '));
});

T('★★同じ 名前は 足せない（黙って 上書きしない）★★', () => {
  const z = Z.read(new Uint8Array(中));
  let 投げた = false;
  try { z.addText('xl/styles.xml', '<x/>'); } catch (e) { 投げた = true; }
  if (!投げた) {
    throw new Error('★もう 在る 名前を 足せて しまいました★'
      + '／★同じ 名前が 2つ 在る 包みは Excel が 修復を 言います★');
  }
});

await 待('★1度 消した 名前は 足し直せる★', async () => {
  const z = Z.read(new Uint8Array(中));
  z.remove('xl/calcChain.xml');
  z.addText('xl/calcChain.xml', '<c/>');
  const r = await z.build();
  const t = await Z.read(r.bytes).text('xl/calcChain.xml');
  if (t !== '<c/>') throw new Error('★中身が ' + t + '★');
});

await 待('★★同じ 材料から 同じ 包みが 出る（時刻を 入れて いない）★★', async () => {
  const 作る = async () => {
    const z = Z.read(new Uint8Array(中));
    z.addText('xl/worksheets/sheet2.xml', '<x>tameshi</x>');
    const r = await z.build();
    return crypto.createHash('sha256').update(Buffer.from(r.bytes)).digest('hex');
  };
  const a = await 作る();
  await new Promise((ok) => setTimeout(ok, 1100));   /* ★秒を またがせます★ */
  const b = await 作る();
  if (a !== b) {
    throw new Error('★1.1秒 空けたら 別の 包みに なりました★' + NL
      + '       ' + a + NL + '       ' + b
      + NL + '       ★今の 時刻を 入れて います★＝★同じ 材料で 同じ 物が 出ません★');
  }
});

/* ══ ★★部品の 名前の 区切りは 斜線（/）★★ ══（2026-09-21 経営者1 の 注文）
     ★★なぜ 要るか★★
       経営者1 が 同じ日に 踏みました＝`CreateFromDirectory` は 部品の 名前に
       ★逆斜線★ を 使います（zip の 決まりは 斜線）。
       ⇒前後を 突き合わせると ★全部 「増えた」に 見えました★。
     ⇒`add` に 逆斜線の 名前を 渡すと ★Excel は 開くかも しれませんが★
       ★突き合わせる 道具が 壊れます★。
     ★だから 名前を そのまま 通すのでは なく 数えます★ */
await 待('★★足した 部品の 名前が 斜線で 入る（逆斜線が 混ざらない）★★', async () => {
  const z = Z.read(new Uint8Array(中));
  z.addText('xl/worksheets/sheet9.xml', '<x/>');
  const r = await z.build();
  const 名 = Z.read(r.bytes).names();
  const 逆 = 名.filter((n) => n.indexOf(String.fromCharCode(92)) >= 0);
  if (逆.length) {
    throw new Error('★逆斜線の 名前が ' + 逆.length + '本★ ' + 逆.join(' ')
      + '／★zip の 決まりは 斜線です★');
  }
  if (名.indexOf('xl/worksheets/sheet9.xml') < 0) {
    throw new Error('★足した 名前が 見つかりません★');
  }
});

await 待('★★元の 包みにも 逆斜線が 1本も 無い★★（★材料の 側も 数える★）', async () => {
  const 逆 = 元の名.filter((n) => n.indexOf(String.fromCharCode(92)) >= 0);
  if (逆.length) throw new Error('★元の 包みに 逆斜線が ' + 逆.length + '本★ ' + 逆.join(' '));
});

console.log('');
console.log('  ★見て いない 事★');
console.log('    ・★実Excel が この 包みを どう 開くかは 測れません★（COM が 要る）');
console.log('    ・`[Content_Types].xml` や rels を 直すのは ★呼ぶ 側の 仕事★');
console.log('    ・★`.xlsb` では 試して いません★');
console.log('zip-tasu: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと壊して 赤に なるか★（★lib は 1字も 触りません★）');
  let 悪 = 0;
  const 見る = (な, ok) => { if (ok) console.log('  ok   ' + な); else { console.log('  NG   ' + な); 悪++; } };

  /* ㋐ ★足さずに 組み直したら 部品は 増えない★＝門②が 本当に 数えて いるか */
  const z1 = Z.read(new Uint8Array(中));
  const r1 = await z1.build();
  見る('★足さなければ 部品は 増えない★', Z.read(r1.bytes).names().length === 元の名.length);

  /* ㋑ ★消したら 減る★＝門②が 「減った」を 見つけられるか */
  const z2 = Z.read(new Uint8Array(中));
  z2.remove('xl/calcChain.xml');
  const r2 = await z2.build();
  見る('★消すと 部品が 減る（門が 気づける 形）★',
    Z.read(r2.bytes).names().length === 元の名.length - 1);

  /* ㋒ ★書き換えたら 中身が 変わる★＝門③が 本当に 中を 見て いるか */
  const z3 = Z.read(new Uint8Array(中));
  z3.replaceText('xl/styles.xml', '<x/>');
  const r3 = await z3.build();
  const t3 = await Z.read(r3.bytes).text('xl/styles.xml');
  見る('★書き換えると 中身が 変わる（門が 中を 見て います）★', t3 === '<x/>');

  process.exit(悪 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
