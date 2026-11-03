/* kinji-ji-kazoeru.mjs — ★決まりで 禁じられた 字を 数える★（2026-09-18）
 *
 *  ★★決まり★★ `team/global-rules.md` §6／§8（★過去の 失敗パターン（再発防止）★）
 *    ・★スマートクォート（" " 「 」）混入禁止 → ASCII 文字のみ★
 *    ・★Unicode 省略記号（|）禁止 → ... （ASCII 3ピリオド）★
 *    ・★バッククォート種別ミス禁止★
 *
 *  ★★なぜ 全部を 直さないか★★
 *    ★2026-09-18 に 数え直しました★（★経営者1 と 私で 別々に★）
 *      `|` は ★覚書き（コメント）の 中だけ★ | 生きた 行 ★0件★
 *    ⇒★★今 壊れる 物は 在りません★★
 *    ⇒★2,000件 近くを 触るのは ★今 動いて いる 物を 触る★★＝★止めました★
 *    ⇒★★代わりに「今日の 数」を 決め打ちに して ★増えたら 赤★★★
 *      ＝★「これから 書く 分は ASCII」が ★心がけ★では なく ★機械★に なります★
 *
 *  ★★私が 1回 踏んだ 罠★★
 *    `grep -o '[||||]'` | ★11,055件★と 出ました
 *    ⇒★★バイトで 当たって いました★★（多バイトの 字を 括弧に 入れた）
 *    ⇒★★だから この 道具は 字で 数えます★★（`String.split` の 長さ）
 *
 *  使い方: node scripts/kinji-ji-kazoeru.mjs
 */
import fs from 'node:fs';
import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/* ★見る 物★（★見る 範囲を 先に 数えて 書く★） */
const 見る = [
  { 名: 'lib', 道: 'lib', 拡: ['.js'] },
  { 名: 'tests', 道: 'tests', 拡: ['.mjs', '.js'] },
  { 名: 'scripts', 道: 'scripts', 拡: ['.mjs', '.js'] },
  /* ══ ★★2026-09-18 に 足しました★★ ══
       ★訳★ ... ★門が 一番 守るべき 場所を 見て いませんでした★
         ＝今日 私が `...` で ★4回★ 転び、経営者1 も ★1回★ 転びました
         ＝★そのうち 3回は `docs/measured` の 中の 道具／紙／出しでした★
         ＝★門の 正規表現が「他の 道具の 出しに 在った `...`」を 掴んで いた★
       ★★足す 前に 数を 出して 見せて あります★★（★4,343個／254本★・経営者1 へ）
       ★★上限は「足した 日の 数 そのまま」★★＝★余裕を 残しません★
       ★見る 型★ ... ★道具（走る 物）と 紙（人が 読む 物）★
         ＝`.tsv` は ★実Excel の 出し★なので 見ません（★中の 字は 答えです★）
       ★★中を 歩きます★★（`kansuu46/` が 在る） */
  { 名: 'docs/measured', 道: 'docs/measured', 拡: ['.mjs', '.js', '.ps1', '.py', '.md'], 再帰: true },
];

/* ══ ★★見ない 物（★名指し＋訳★）★★ ══
     ★2026-09-18 に 1本ずつ 中を 見て 決めました★（★数だけで 決めて いません★）

     ⑴★`*.min.js`★ | ★借り物です★
        `lib/xlsx.full.min.js` | SheetJS 0.20.3（★Apache-2.0★・951,904 B）
        ★配り元と 1バイトも 違いません★（`lib/xlsx-0.20.3-DOKOKARA.md`）
        ⇒★★借り物の 中の 字を 数えても 意味が ありません★★（★触れません★）
        ⇒★ここに スマートクォート 92件★

     ⑵★`lib/symbols.js`★ | ★★字そのものが データです★★
        43〜46行 | ★左二重引用符／右二重引用符／左引用符／右引用符★ の ★名前の 表★
        ⇒★★この 字を 消すと 中身が 壊れます★★
        ⇒★スマートクォート 4件★

     ★★この 2つを 外さないと「守れない 門」に なります★★
       ＝★守れない 門は 消されます★ */
const 見ない名 = [
  { 名: 'xlsx.full.min.js', 訳: '★借り物（SheetJS・Apache-2.0）★' },
  { 名: 'symbols.js', 訳: '★字そのものが データ（43〜46行）★' },
];
/* ★逐語の 写し★＝他の 席の 物を そのまま 写した 物は、★道と 中身（git の blob の id）が 両方 同じ 時だけ★ 数えない
   （2026-10-11 本番前の 対立役＝名前だけで 外すと、同じ 名前の 別の 物が 素通りした）。中身が 変われば 数える＝見直して id を 書き換える */
const 写しの物 = [
  { 道: 'tests/kinshi-ji.test.mjs', blob: 'f4aa3cebe77f8783d9cc991f905f02409310d65e', 訳: '★禁止の 字の 見張り＝全席 共通の 物を 逐語で 写した（2026-10-11・指示役）。字を 替えると 他の 席と 揃わない★' },
];
function blobのid(p) { const b = fs.readFileSync(p); return crypto.createHash('sha1').update('blob ' + b.length + '\0').update(b).digest('hex'); }

/* ★★正規表現を やめました★★（2026-09-18）
     ★`[\/]` と 書いた のに `[\/]` に 落ちて いました★
     ＝★逆斜線の 逃がしが 落ちる★（★記憶に 在る 罠★・今日 2回目）
     ⇒★★名前で 比べます＝逆斜線が 1つも 要りません★★
     ⇒★見分け方★ | ★lib が 111本の はずなのに 110本でした★
              ＝★外れた のは 1本だけ★ */
/* ★★★この 道具自身に 禁字を 書きません★★★
     ★訳★ | ★門が 自分を 数えて 永久に 赤に なります★
     ＝★「守る 紙に 本番の 印を 書くな」と 同じ 型★
     ⇒★★コードで 字を 作ります★★（ファイルに その 字が 1つも 無い） */
const 字 = (n) => String.fromCharCode(n);
const 禁字 = [
  { 字: 字(0x2026), 名: '省略記号 U+2026' },
  { 字: 字(0x201C), 名: 'スマート引用符 左 U+201C' },
  { 字: 字(0x201D), 名: 'スマート引用符 右 U+201D' },
  { 字: 字(0x2018), 名: 'スマート引用符 左一重 U+2018' },
  { 字: 字(0x2019), 名: 'スマート引用符 右一重 U+2019' },
];

/* ══ ★★2026-09-18 の 数（★これを 超えたら 赤★）★★ ══
     ★増やさない為の 門です／減らすのは いつでも よい★ */
const 上限 = {
  'lib':     { 0x2026: 1045, 0x201C: 0, 0x201D: 0, 0x2018: 0, 0x2019: 0 },
  'tests':   { 0x2026: 1863, 0x201C: 0, 0x201D: 0, 0x2018: 0, 0x2019: 0 },
  'scripts': { 0x2026: 343, 0x201C: 0, 0x201D: 0, 0x2018: 0, 0x2019: 0 },
  /* ★★2026-09-18 に 足した 範囲★★（docs/measured）
       ★上限は ★足した 日の 数 そのまま★★（余裕を 残しません）
       ★スマートクォートは 0に しました★
         ＝2件（`toru-jitsubutsu-shoshiki.ps1` の 覚書き 1行）しか 無く
           ★直せる 数だった★ので ★直しました★（`"` に）
       ★U+2026 は 3,659個★＝★今 動いて いる 物を 触らない★ */
  'docs/measured': { 0x2026: 3659, 0x201C: 0, 0x201D: 0, 0x2018: 0, 0x2019: 0 },
};
/* ══ ★★「生きた 行」の 上限★★ ══
     ★2026-09-18 に 数えたら ★0では ありませんでした★★
     ⇒★調べると ★字の 中（'|' の ような 出す 字）★が 大半です★
     ⇒★★これは 壊れる 物では ありません★★（★決まりが 禁じるのは ★混入★★）
     ⇒★★だから 0では なく「今日の 数」を 上限に します★★
     ⇒★減らすのは いつでも よい／★増やさない★のが この 門の 仕事★ */
/* ★`docs/measured` の 「生きた 行」が 多いのは ★紙（.md）だから★です
     ＝この 道具の 「覚書きを 落とす」は ★JS の 書き方しか 知りません★
     ＝★.md は 全部「生きた 行」に なります★（★甘く 見ない 向きの 誤り★）
     ⇒★数え方を 変えず、今日の 数を そのまま 上限に します★ */
const 生きた上限 = { 'lib': 32, 'tests': 1019, 'scripts': 145, 'docs/measured': 3146 };

function 集める(道, 拡, 再帰) {
  const 出 = [];
  const d = path.join(ROOT, 道);
  if (!fs.existsSync(d)) return 出;
  for (const f of fs.readdirSync(d)) {
    const p = path.join(d, f);
    if (fs.statSync(p).isDirectory()) {
      /* ★中を 歩くのは 頼まれた 時だけ★（★前と 同じ 動きを 変えない★） */
      if (再帰 && f !== 'node_modules' && f !== '__pycache__') 出.push(...集める(path.join(道, f), 拡, 再帰));
      continue;
    }
    if (!拡.includes(path.extname(f))) continue;
    if (見ない名.some((x) => x.名 === f)) continue;   /* ★名前で 外す（逆斜線を 使わない）★ */
    const 相対 = path.join(道, f).split(path.sep).join('/');
    if (写しの物.some((x) => x.道 === 相対 && x.blob === blobのid(p))) continue;
    出.push(p);
  }
  return 出;
}

/* ★覚書きを 落とす★（★雑ですが ★生きた 行★ を 甘く 見ない 向き★）
   ＝`//` の 後ろ／`/* | *​/` の 中／`*` で 始まる 行 を 落とす */
function 覚書きを落とす(t) {
  let 出 = '', 中 = false;
  for (const 行 of t.split('\n')) {
    let l = 行;
    if (中) {
      const e = l.indexOf('*/');
      if (e < 0) { 出 += '\n'; continue; }
      l = l.slice(e + 2); 中 = false;
    }
    for (;;) {
      const s = l.indexOf('/*');
      if (s < 0) break;
      const e = l.indexOf('*/', s + 2);
      if (e < 0) { l = l.slice(0, s); 中 = true; break; }
      l = l.slice(0, s) + l.slice(e + 2);
    }
    const c = l.indexOf('//');
    if (c >= 0) l = l.slice(0, c);
    出 += l + '\n';
  }
  return 出;
}

/* ══ ★★2026-09-19 ... ★数える所を 関数に しました★★★ ══
 *   ★訳★ ... `tests/kinji-ji.test.mjs` が この 道具を ★別のプロセスで★ 呼んで いました。
 *          ⇒`scripts/tests-registered.mjs`（★試験は 登録するまで 1本も 走らない★）が
 *            ★「試験の 一覧を 持つ物」だと 読み違えて★ ★CI を 40回 赤に して いました★。
 *          ⇒★別のプロセスを 呼ばなければ 読み違えようが ありません★。
 *   ★返す 物★ ... ★上限を 超えた 件数★（0 なら 緑）
 *   ★命令の 字から 呼んだ 時は 今まで通り★ ... `node scripts/kinji-ji-kazoeru.mjs` */
export function 数を数える() {
  let 赤 = 0;
  console.log('');
  console.log('[kinji-ji] ★決まりで 禁じられた 字を 数える★（★字で 数えます＝バイトでは ない★）');
  console.log('');
  for (const g of 見る) {
    const ファイル = 集める(g.道, g.拡, !!g.再帰);
    const 合 = {}, 生 = {};
    for (const c of 禁字) { 合[c.字] = 0; 生[c.字] = 0; }
    /* ★★どの ファイルに 何個 在るかも 覚えます★★（2026-09-18・経営者1 の 注文）
         ★訳★ ... 上限を 超えた 時 ★どこかを 人が 手で 数えて いました★（★今日 2回★）
         ⇒★超えた 時だけ ★多い 順に 出します★★
         ⇒★同じ 木を 2席で 使って いるので ★誰の 書きかけかも すぐ 分かります★★ */
    const 本ごと = [];
    for (const p of ファイル) {
      const t = fs.readFileSync(p, 'utf-8');
      const k = 覚書きを落とす(t);
      let 本の合 = 0;
      for (const c of 禁字) {
        const a = t.split(c.字).length - 1;
        合[c.字] += a;
        生[c.字] += k.split(c.字).length - 1;
        本の合 += a;
      }
      if (本の合) 本ごと.push({ 道: path.relative(ROOT, p), 数: 本の合 });
    }
    console.log('  ★' + g.名 + '★（' + ファイル.length + '本）');
    for (const c of 禁字) {
      if (!合[c.字] && !生[c.字]) continue;
      const 限 = (上限[g.名] || {})[c.字.charCodeAt(0)];
      const だめ = (限 !== undefined && 合[c.字] > 限);
      const 生限 = (生きた上限[g.名] === undefined) ? 0 : 生きた上限[g.名];
      const 生だめ = 生[c.字] > 生限;
      if (だめ || 生だめ) 赤++;
      console.log('    ' + c.名.padEnd(26)
        + ' 全部 ' + String(合[c.字]).padStart(5) + '（上限 ' + 限 + '）'
        + ' ／ ★生きた 行 ' + 生[c.字] + '★（上限 ' + 生限 + '）'
        + (だめ ? '  ★★増えました★★' : '') + (生だめ ? '  ★★生きた 行に 出ました★★' : ''));
    }
    if (禁字.every((c) => !合[c.字])) console.log('    ★0件★');
    /* ★超えた 時だけ 出どころを 出す★ */
    const 超えた = 禁字.some((c) => {
      const 限 = (上限[g.名] || {})[c.字.charCodeAt(0)];
      return (限 !== undefined && 合[c.字] > 限)
        || 合[c.字] > 0 && 生[c.字] > ((生きた上限[g.名] === undefined) ? 0 : 生きた上限[g.名]);
    });
    if (超えた) {
      console.log('    ★★出どころ（多い 順に 8本）★★');
      for (const x of 本ごと.sort((a, b) => b.数 - a.数).slice(0, 8)) {
        console.log('      ' + String(x.数).padStart(5) + '  ' + x.道);
      }
      console.log('      ★★木（ディスク）の 字を 数えて います★★');
      console.log('        ＝★commit して いない 書きかけも 入ります★');
      console.log('        ＝★同じ 木を 2席で 使う 時は ★相手の 書きかけで 赤に なります★★');
    }
  }
  console.log('');
  if (赤) {
    console.log('★★' + 赤 + '件 上限を 超えて います★★');
    console.log('  ⇒★これから 書く 分は ★ASCII★ に して ください★');
    console.log('  ⇒★`|` は `...` ／ スマートクォートは `"` `\'`★');
    return 赤;
  }
  console.log('★上限の 中です★（★2026-09-18 の 数を 決め打ちに して います★）');
  return 0;
}

/* ★命令の 字から 直に 呼ばれた 時だけ 終わり値を 返す★（★読み込んだだけでは 何も しない★） */
if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  process.exit(数を数える() ? 1 : 0);
}
