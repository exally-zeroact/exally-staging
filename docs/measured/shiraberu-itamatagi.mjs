/* shiraberu-itamatagi.mjs — ★㋖「板またぎ」を 作る前に 形を 数える★（2026-09-15）
 *
 *  ★★なぜ 作る前に 数えるか★★
 *    ㋖は 大きい（★7,372本＋染まり 312本が 一度に 動く★）。
 *    ★作ってから 数えると「作った 物に 合う 数」しか 出ません★
 *    ⇒★先に「どんな 形が 何本 在るか」を 出して から 作ります★。
 *
 *  ★★開き方は `osu-jitsubutsu-dodai.mjs` を 呼ぶだけ★★
 *    ★訳★＝★表の 参照を 直した 後の 式★を 見ないと 数が 合いません
 *      （★生の 字だけだと 3,285本／表の 参照を 直すと 7,372本★）。
 *    ★別の 道を 作れば 印は 毎回 抜ける★＝★開く 道は 1つ★（⑱で 学んだ 形）。
 *
 *  ★★この 道具は 台を 建てません★★
 *    ★式の 字を 読むだけ★＝計算しません。
 *    ⇒ `hakaridai-mon`（`osu-*.mjs` の 建て方を 見る 見張り）の 外です。
 *      ★訳★＝★建てて いないので 建て方を 間違えようが ない★。
 *      ⇒ だから 名前を `osu-` に して いません（★押す 道具と 混ぜない★）。
 *
 *  ★★読むだけ★★ … 開く だけ／★1バイトも 書きません★
 *  ★★出すのは 数だけ★★
 *    ★板の 名前は 1つも 出しません★（★会社名・人の 名前が 板の 名前に なって いる 事が 在る★）
 *    ⇒ 出すのは ★形（かたち）と 本数★だけ。
 *
 *  ★★見て いない 物（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★式の 字だけ★ 見て います＝★答えが 合うかは 見て いません★
 *    ・★`#REF!` は 板名では ない★ので 除きます（★`REF!` が 板名に 見える★）
 *    ・★覚書きの 中の `!` も 式には 出て 来ません★（式に 覚書きは 無い）
 *
 *  使い方: node docs/measured/shiraberu-itamatagi.mjs
 */
import path from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const 土台 = await import(pathToFileURL(path.join(ここ, 'osu-jitsubutsu-dodai.mjs')).href);
const XLSX = 土台.XLSX;

const { wb, 直し, 前 } = await 土台.本を開く();

const 板名 = wb.SheetNames;
const 板の番号 = {};
板名.forEach((n, i) => { 板の番号[n] = i; });
const 板の名 = new Set(板名);

let 式全部 = 0, またぐ = 0, 表直しが効いた = 0;
const 形ごと = {};
const 指し先の板 = new Set();
const 無い板の長さ = [];
const 指し先の形 = { '1マス': 0, '四角': 0, '列ぜんぶ': 0, '行ぜんぶ': 0 };
const 何枚指すか = {};
let 板名に空白 = 0, 板名に記号 = 0, 三D = 0, 外の本 = 0;

/* ★板名を 拾う★（★`#REF!` は 除く★／`'…'` も 素も） */
const 拾 = /(#?)(?:'((?:[^']|'')+)'|([A-Za-z_぀-ヿ一-鿿][A-Za-z0-9_. ぀-ヿ一-鿿]*))!/g;

for (let si = 0; si < 板名.length; si++) {
  const sh = wb.Sheets[板名[si]];
  if (!sh || !sh['!ref']) continue;
  const 範 = XLSX.utils.decode_range(sh['!ref']);
  for (let r = 範.s.r; r <= 範.e.r; r++) {
    for (let c = 範.s.c; c <= 範.e.c; c++) {
      const a = XLSX.utils.encode_cell({ r, c });
      const 中 = sh[a];
      if (!中 || !中.f) continue;
      式全部++;
      const 生 = '=' + String(中.f);
      const k = 板名[si] + '|' + r + ',' + c;      /* ★直しの 鍵は 土台と 同じ 形★ */
      const 直った = (直し[k] !== undefined) ? String(直し[k]) : 生;
      const f = 土台.裸に(直った);                        /* ★文字列の 中の `!` を 消す★ */
      const 生裸 = 土台.裸に(生);
      if (f.replace(/#REF!/g, '').indexOf('!') < 0) continue;
      またぐ++;
      if (生裸.replace(/#REF!/g, '').indexOf('!') < 0) 表直しが効いた++;

      if (/\[[0-9]+\]/.test(f)) 外の本++;
      if (/(?:'[^']+:[^']+'|[A-Za-z0-9_぀-ヿ一-鿿]+:[A-Za-z0-9_぀-ヿ一-鿿]+)!\$?[A-Z]{1,3}\$?[0-9]/.test(f)) 三D++;

      let m; const 本で見た = new Set();
      拾.lastIndex = 0;
      while ((m = 拾.exec(f)) !== null) {
        if (m[1] === '#') continue;                       /* ★`#REF!` は 板では ない★ */
        const 名 = (m[2] !== undefined) ? m[2].split("''").join("'") : m[3];
        if (名.indexOf(':') >= 0) continue;
        if (板の名.has(名)) {
          指し先の板.add(板の番号[名]);
          本で見た.add(板の番号[名]);
          if (/\s/.test(名)) { 板名に空白++; }
          if (/[^A-Za-z0-9_\s぀-ヿ一-鿿]/.test(名)) { 板名に記号++; }
          形ごと[(m[2] !== undefined) ? "㋑ '囲んだ 板名'！" : '㋐ 素の 板名！']
            = (形ごと[(m[2] !== undefined) ? "㋑ '囲んだ 板名'！" : '㋐ 素の 板名！'] || 0) + 1;
        } else {
          無い板の長さ.push(名.length);
        }
      }
      何枚指すか[本で見た.size] = (何枚指すか[本で見た.size] || 0) + 1;

      if (/![$]?[A-Z]{1,3}[$]?[0-9]{1,7}\s*:\s*[$]?[A-Z]{1,3}[$]?[0-9]{1,7}/.test(f)) 指し先の形['四角']++;
      else if (/![$]?[A-Z]{1,3}\s*:\s*[$]?[A-Z]{1,3}(?![0-9])/.test(f)) 指し先の形['列ぜんぶ']++;
      else if (/![$]?[0-9]{1,7}\s*:\s*[$]?[0-9]{1,7}/.test(f)) 指し先の形['行ぜんぶ']++;
      else 指し先の形['1マス']++;
    }
  }
}

const 名前の数 = (wb.Workbook && wb.Workbook.Names) ? wb.Workbook.Names.length : 0;
let 板つきの名前 = 0;
if (wb.Workbook && wb.Workbook.Names) {
  for (const n of wb.Workbook.Names) if (n.Ref && String(n.Ref).indexOf('!') >= 0) 板つきの名前++;
}

const 触った = 土台.触っていないか(前);

console.log('\n# ★㋖ 板またぎ＝★作る前に★ 形を 数えた★（2026-09-15）');
console.log('#   ★読むだけ★ … ★' + (触った === true || 触った === undefined ? '' : '') + '本体は 動いて いません★');
console.log('#   ★板の 名前は 1つも 出して いません★（数と 形だけ）');
console.log('#   ★開き方は `osu-jitsubutsu-dodai.mjs` の 1本だけ★（表の 参照を 直した 後の 式）\n');

console.log('★板の 数★ … ' + 板名.length + '／★指されて いる 板★ … ' + 指し先の板.size);
console.log('★式 全部★ … ' + 式全部);
console.log('★★板を またぐ 式★★ … ' + またぐ + '（'
  + (式全部 ? (またぐ * 100 / 式全部).toFixed(1) : '0') + '%）');
console.log('   ★うち 表の 参照を 直して 初めて 見えた 分★ … ' + 表直しが効いた
  + '（★生の 字だけ 見ると ' + (またぐ - 表直しが効いた) + '本に しか 見えません★）\n');

console.log('★★板名の 書き方（のべ）★★');
for (const [名, n] of Object.entries(形ごと).sort((a, b) => b[1] - a[1])) {
  console.log('   ' + String(n).padStart(6) + '個  ' + 名);
}
console.log('   ' + String(三D).padStart(6) + '本  ㋒ 板：板！（3D）');
console.log('   ' + String(外の本).padStart(6) + '本  ㋓ ［外の 本］！');
console.log('   ' + String(無い板の長さ.length).padStart(6) + '個  ★この 本に 無い 板を 指す★'
  + (無い板の長さ.length ? '（★名前の 長さだけ★ ' + [...new Set(無い板の長さ)].sort((a, b) => a - b).join('/') + '）' : ''));
console.log('   ' + String(板名に空白).padStart(6) + '個  ★板の 名前に 空白★（⇒ `\'…\'` が 要る）');
console.log('   ' + String(板名に記号).padStart(6) + '個  ★板の 名前に 記号★\n');

console.log('★★1本が いくつの 板を 指すか★★');
for (const k of Object.keys(何枚指すか).sort((a, b) => a - b)) {
  console.log('   ' + String(何枚指すか[k]).padStart(6) + '本  ' + k + '枚');
}
console.log('');

console.log('★★指し先の 形★★');
for (const [名, n] of Object.entries(指し先の形)) console.log('   ' + String(n).padStart(6) + '本  ' + 名);
console.log('');

console.log('★定義された 名前★ … ' + 名前の数 + '（うち `!` を 含む … ' + 板つきの名前 + '）');
