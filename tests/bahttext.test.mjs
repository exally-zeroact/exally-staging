/* bahttext.test.mjs — ★BAHTTEXT が 実Excel と 同じ 字を 出すか★ 2026-09-04
 *
 *  ★何を 正しいと したか★
 *    ★私が タイ語の 決まりを 考えて 作った 表では ありません★。
 *    ★この機械の 実Excel（365）に 116通り＋変な物 9通りを 打たせた 答え★が 真値です。
 *      道具 … tools/bahttext-golden.ps1 ／ tools/bahttext-err.ps1
 *      台帳 … tests/fixtures/bahttext-golden.json
 *      ★司さんの 実物は 1バイトも 触っていません★（新規ブック・保存しない）
 *
 *  ★境界（司さんの 金額で 効く 所）★
 *    0円／1円未満（サタンだけ）／マイナス／丸めの 境目（x.xx5）／
 *    一の位が1（เอ็ด）／十の位が1（สิบ）／十の位が2（ยี่สิบ）／百万・十億・一兆以上
 *    ⇒★どれが 何通り 入っているかを 数えて 出す★（0通りなら 赤）
 *
 *  走らせ方: node tests/bahttext.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const B = require_(path.join(ROOT, 'lib/bahttext.js'));
const G = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/bahttext-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[bahttext] ★実Excel と 同じ 字を 出すか★');

/* ── ① 台帳が 空振りしていない ── */
T('★台帳が 在る（実Excel が 出した 答え）★', Array.isArray(G.本) && G.本.length > 0);
T('★出どころが 書いてある（版・道具）★',
  /Excel 365/.test(G._版 || '') && /bahttext-golden\.ps1/.test((G._道具 || []).join(' ')),
  JSON.stringify(G._版));

/* ── ② ★116通り 全部★ ── */
{
  const ちがい = [];
  for (const x of G.本) {
    const 出 = B.字にする(x.n);
    if (出 !== x.text) ちがい.push(x.n + ' … Excel「' + x.text + '」／うち「' + 出 + '」');
  }
  T('★実Excel の ' + G.本.length + '通りと 1つ残らず 同じ★（ちがい ' + ちがい.length + '）',
    ちがい.length === 0, ちがい.slice(0, 5).join('\n       '));
  console.log('       … 見た 総数 ' + G.本.length + '通り ／ ★合った ' + (G.本.length - ちがい.length) + '通り★');
}

/* ── ③ ★境界が 台帳に 入っているか★（★0通りの 組が 在れば 赤★）── */
{
  const n = G.本.map((x) => x.n);
  const 整 = (x) => Math.abs(Math.trunc(x));
  const 組 = {
    '0円': n.filter((x) => x === 0).length,
    '1円未満（サタンだけ）': n.filter((x) => x > 0 && x < 1).length,
    'マイナス': n.filter((x) => x < 0).length,
    '丸めの境目（x.xx5）': n.filter((x) => /\.\d\d5$/.test(String(x))).length,
    '一の位が1（เอ็ด）': n.filter((x) => 整(x) % 10 === 1).length,
    '十の位が1（สิบ）': n.filter((x) => Math.floor(整(x) / 10) % 10 === 1).length,
    '十の位が2（ยี่สิบ）': n.filter((x) => Math.floor(整(x) / 10) % 10 === 2).length,
    '百万以上': n.filter((x) => Math.abs(x) >= 1e6).length,
    '十億以上': n.filter((x) => Math.abs(x) >= 1e9).length,
    '一兆以上': n.filter((x) => Math.abs(x) >= 1e12).length,
  };
  const 空 = Object.keys(組).filter((k) => !組[k]);
  T('★境界が 台帳に 全部 入っている★', 空.length === 0, '入っていない … ' + 空.join(' / '));
  console.log('       … ' + Object.keys(組).map((k) => k + ' ' + 組[k]).join(' ／ '));
}

/* ── ④ ★変な物を 渡した時★（実Excel に 打たせた 9通り）── */
{
  const 読み = { '文字': 'abc', '数の字': '123', '空の セル': null, '空の 字': '',
    'TRUE': true, 'FALSE': false, '1e20': 1e20, '1e21': 1e21 };
  let 見た = 0;
  for (const x of (G.変な物 || [])) {
    if (!(x.な in 読み)) continue;                 /* エラーは エンジン側で 素通しする */
    見た++;
    const 出 = B.字にする(読み[x.な]);
    const 期待 = /^#/.test(x.text) ? null : x.text;
    T('★' + x.な + ' … ' + x.式 + '★', 出 === 期待,
      'Excel「' + x.text + '」／うち「' + 出 + '」');
  }
  T('★変な物を 1つ以上 見た（空振りしていない）★', 見た >= 6, '見た ' + 見た + '通り');
  /* ★空の セルは 2通りの 形で 来る★（null と undefined）＝★両方 0 に する★
     ★自己確認が 教えてくれた★＝null だけ 見ていたら、その 行を 消しても 赤に ならなかった */
  T('★空の セル（何も 渡されない）も 0★', B.字にする(undefined) === 'ศูนย์บาทถ้วน',
    String(B.字にする(undefined)));
  T('★空の セル（null）も 0★', B.字にする(null) === 'ศูนย์บาทถ้วน');
}

/* ── ⑤ ★丸めは 数の まま やらない★（機械の 中の ずれ）──
   ★私は 最初 2.675 を 例に 書いたが、この 機械では 2.675*100 は 267.50000000000006 で
     ★数の まま でも 268 に なる★＝★例として 間違い だった★（試験が 教えてくれた）。
   ★本当に ずれるのは 1.005★ … 1.005*100 は 100.49999999999999 ⇒ 丸めると 100（＝サタン 0）
     実Excel は ★หนึ่งบาทหนึ่งสตางค์（1.01）★＝★1サタン 少なく なる★ */
T('★1.005 は 1バーツ1サタン（実Excel と 同じ）★',
  B.字にする(1.005) === 'หนึ่งบาทหนึ่งสตางค์', B.字にする(1.005));
T('★数のまま 100倍すると 確かに ずれる（この 罠が 本物だと 示す）★',
  Math.round(1.005 * 100) === 100, String(Math.round(1.005 * 100)));
T('★2.675 も 実Excel と 同じ（この 機械では 数のままでも 合うが、揃えて 見る）★',
  B.字にする(2.675) === 'สองบาทหกสิบแปดสตางค์', B.字にする(2.675));

/* ── ⑥ ★繋いである（画面が 読み込んでいる）★ ── */
{
  const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  T('★画面が lib/bahttext.js を 読み込んでいる★', /lib\/bahttext\.js/.test(book));
  const plug = fs.readFileSync(path.join(ROOT, 'lib/formula-extra-plug.js'), 'utf8');
  T('★エンジンに 繋いである★', /'BAHTTEXT':\s*\{ method: 'bahttext'/.test(plug));
}

/* ── ⑦ ★エンジンに 通して 打つ★（部品だけ 緑では 足りない）──
   ★2026-09-04 ここで 1つ 見つけた★
     ★空の セルは エンジンから ★Symbol★で 渡ってくる★
     ⇒ Number(Symbol) は ★その場で 落ちる★＝★式だけでなく 画面ごと 危ない★
     ⇒★部品（lib）だけ 試していたら 気づけなかった★ */
{
  const M = require_(path.join(ROOT, 'hyperformula.full.min.js'));
  const HF = M.HyperFormula;
  /* ★画面では window.HyperFormula が 全部 持っている★＝node では 同じ 形に 混ぜる */
  const H = Object.assign(Object.create(HF), M,
    { registerFunctionPlugin: HF.registerFunctionPlugin.bind(HF) });
  const FE = require_(path.join(ROOT, 'lib/formula-extra.js'));
  const P = require_(path.join(ROOT, 'lib/formula-extra-plug.js'));
  P.つなぐ(H, FE);
  const 組 = [
    [123.45, 'หนึ่งร้อยยี่สิบสามบาทสี่สิบห้าสตางค์'],
    [0, 'ศูนย์บาทถ้วน'],
    ['', '#VALUE!'],
    [null, 'ศูนย์บาทถ้วน'],                 /* ★空の セル＝Symbol で 来る★ */
    ['abc', '#VALUE!'],
    [-1234.56, 'ลบหนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบหกสตางค์'],
    [1.005, 'หนึ่งบาทหนึ่งสตางค์'],
    [true, 'หนึ่งบาทถ้วน'],
    [1e21, 'หนึ่งพันล้านล้านล้านบาทถ้วน'],
  ];
  let 合 = 0;
  for (const [v, e] of 組) {
    let 出;
    try {
      const hf = HF.buildFromArray([[v, '=BAHTTEXT(A1)']], { licenseKey: 'gpl-v3' });
      const g = hf.getCellValue({ sheet: 0, col: 1, row: 0 });
      出 = (g && g.value !== undefined) ? g.value : String(g);
    } catch (err) { 出 = '★落ちた★ ' + String(err && err.message).slice(0, 60); }
    if (出 === e) 合++;
    else T('★エンジンで ' + JSON.stringify(v) + '★', false, '期待「' + e + '」／出た「' + 出 + '」');
  }
  T('★エンジンに 通して ' + 組.length + '通り 打った（全部 一致）★', 合 === 組.length,
    '合った ' + 合 + ' / ' + 組.length);
  console.log('       … 見た 総数 ' + 組.length + '通り ／ ★合った ' + 合 + '通り★');
}

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★本物の 部品を わざと 壊して 赤に なるか★');
  const { execFileSync } = await import('node:child_process');
  const 道 = path.join(ROOT, 'lib/bahttext.js');
  const 元 = fs.readFileSync(道, 'utf8');
  const 壊す = [
    ['★十の位の 1 を หนึ่งสิบ に する★',
      (t) => t.replace('出 += (d === 1) ? シップ : (d === 2 ? イー + シップ : 数字[d] + シップ);',
        '出 += 数字[d] + シップ;')],
    ['★十の位の 2 を ยี่ に しない★',
      (t) => t.replace('(d === 2 ? イー + シップ : 数字[d] + シップ)', '(数字[d] + シップ)')],
    ['★一の位の 1 を いつも หนึ่ง に する★',
      (t) => t.replace("出 += (d === 1 && (出 !== '' || 上が在る)) ? エット : 数字[d];", '出 += 数字[d];')],
    ['★サタンが 0 の時 ถ้วน を 付けない★',
      (t) => t.replace('return 頭 + 円 + バーツ + チュアン;', 'return 頭 + 円 + バーツ;')],
    ['★整数が 0 でも บาท を 付ける★',
      (t) => t.replace('if (円 === \'\') return 頭 + サ + サタン;', 'if (円 === \'\') return 頭 + バーツ + サ + サタン;')],
    ['★マイナスの ลบ を 落とす★', (t) => t.replace("var 頭 = r.負 ? ロブ : '';", "var 頭 = '';")],
    ['★丸めを 数の まま やる（2.675 が ずれる）★',
      (t) => t.replace('var s = v.toPrecision(15);', 'var s = String(Math.round(v * 100) / 100);')],
    ['★半分を 下へ 落とす★', (t) => t.replace("if (+小[2] >= 5) サ += 1;", "if (+小[2] > 5) サ += 1;")],
    ['★100万ごとの ล้าน を 重ねない★',
      (t) => t.replace('return 大整数を字に(上) + ラーン + 六桁を字に(+下, true);',
        'return 六桁を字に(+上, false) + ラーン + 六桁を字に(+下, true);')],
    ['★空の 字を 0 に する★', (t) => t.replace("if (x === '') return null;", "if (x === '') x = 0;")],
    ['★空の セル（渡されない）を #VALUE! に する★',
      (t) => t.replace('if (x === null || x === undefined) x = 0;', '')],
  ];
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(道, 壊れ);
    let 赤 = false;
    try { execFileSync(process.execPath, [path.join(ROOT, 'tests', 'bahttext.test.mjs')], { stdio: 'pipe' }); }
    catch (e) { 赤 = true; }
    fs.writeFileSync(道, 元);                 /* ★必ず 戻す★ */
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  T('★本物は 壊していない（戻した）★', fs.readFileSync(道, 'utf8') === 元);
  /* ★繋ぐ 側（lib/formula-extra-plug.js）も 壊して 見る★
     ＝★空の セルの Symbol を 見なくすると、画面ごと 落ちる★ */
  const 繋 = path.join(ROOT, 'lib/formula-extra-plug.js');
  const 繋元 = fs.readFileSync(繋, 'utf8');
  /* ★片方ずつ 消しても もう片方が 効くので 赤に ならない★（2026-09-04 実測）
     ⇒★2つ まとめて 外す★＝これが 本当の 守り */
  const 繋壊 = 繋元.replace(/var 空か = \([\s\S]*?\);/,
    'var 空か = (v === null || v === undefined);');
  if (繋壊 === 繋元) { console.log('  ★素通り★  ★空の セル（Symbol）の 守りが 見つからない（印が 古い）★'); fail++; }
  else {
    fs.writeFileSync(繋, 繋壊);
    let 赤 = false;
    try { execFileSync(process.execPath, [path.join(ROOT, 'tests', 'bahttext.test.mjs')], { stdio: 'pipe' }); }
    catch (e) { 赤 = true; }
    fs.writeFileSync(繋, 繋元);
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + '★空の セル（Symbol）の 守りを 2つとも 外す★（lib/formula-extra-plug.js）');
    if (!赤) fail++;
    T('★繋ぐ 側も 壊していない（戻した）★', fs.readFileSync(繋, 'utf8') === 繋元);
  }
}

console.log('');
console.log('bahttext: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
