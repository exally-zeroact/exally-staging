/* kami-no-zairyou.test.mjs — ★★紙が 古いまま 置かれるのを 止める★★（2026-09-15）
 *
 *  ★★今日 実際に 起きた（これが 作る 訳）★★
 *    `docs/measured/golden-marume-mae-ato-2026-09-08.tsv` は
 *      ★2026-09-10 に commit された★のに
 *      中身は ★2026-09-09 の XIRR の 直し（`e5c3c6c`）より 古い★ ものでした。
 *    ⇒★★走らせずに 置いた 紙★★＝「最後に 走らせた 日で 止まる」より 悪い。
 *    ⇒★★「取った日」は「走らせた 証し」では ありません★★
 *
 *  ★★なぜ これが 一番 静かに 効くか★★
 *    紙は ★突き合わせの「正」★に 使われます。
 *    古い 紙が 残ると ★★自分の 古い 答えと 自分を 比べて 100%★★に なります。
 *    ＝★「偽の 勝ち」の 一番 静かな 形★（★誰も 赤に なりません★）
 *
 *  ★★どう 止めるか★★
 *    紙の 頭に ★その 回に 読んだ 本の 指紋★を 書いて おく（`zairyou-no-yubimon.mjs`）。
 *    ★1つでも 中身が 変わったら 赤★＝★取り直せ★。
 *
 *  ★★git を 使わない 訳（測って 決めました）★★
 *    `.github/workflows/*.yml` に `fetch-depth` の 指定が ★1つも 無い★（実測）
 *    ⇒ `actions/checkout@v4` の 既定は ★深さ 1★
 *    ⇒★CI では `git log -- そのファイル` が 使えません★
 *    ⇒★★手元だけ 緑・CI は 赤★の 形に なります★
 *    ⇒ だから ★中身の 指紋★で やります（★git が 要らない★）。
 *
 *  ★★道具だけ 見ても 足りません★★（指示役1 の 注文・2026-09-15）
 *    ★今日の XIRR は ★道具では なく 台（lib）★が 変わった★のに 紙が 古いまま でした。
 *    ⇒★指紋は ★その 回に 本当に 読んだ 本★を 全部 取ります★
 *      （道具／`exally-formula.js`／`lib/formula-*.js`／借り物 …）
 *    ⇒★広く 取り過ぎない★＝★読んで いない 本は 入れません★（★赤に しすぎない★）
 *
 *  ★★見て いない 事（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★指紋の 行が 無い 紙は 見て いません★
 *      ＝★昔の 紙は 守られて いません★（★取り直す 時に 付きます★）
 *      ⇒★何枚 守れて いて 何枚 守れて いないかを 毎回 出します★
 *    ・★紙の 中身が 正しいかは 見て いません★＝★古いかだけ★
 *    ・★実Excel で 押した 紙（`toru-*.ps1` が 作る 物）は 別★
 *      ＝★あれは うちの 台に 依りません★＝★指紋は 要りません★
 *
 *  使い方: node tests/kami-no-zairyou.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 紙置き = path.join(ROOT, 'docs/measured');
const { 突き合わせ } = await import(pathToFileURL(path.join(紙置き, 'zairyou-no-yubimon.mjs')).href);

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[kami-no-zairyou] ★紙が 古いまま 置かれるのを 止める★');

const 印 = '# ★材料★\t';
function 紙ら() {
  const 出 = [];
  for (const f of fs.readdirSync(紙置き)) {
    if (!/^golden-.*\.tsv$/.test(f)) continue;
    出.push({ 名: 'docs/measured/' + f, 道: path.join(紙置き, f) });
  }
  return 出;
}

T('★★① 指紋の 在る 紙が 1枚でも 古くない★★', () => {
  const 全 = 紙ら();
  if (!全.length) throw new Error('★紙が 1枚も 無い＝空振り★');
  const 守れて = [], 守れず = [], 悪い = [];
  for (const k of 全) {
    const s = fs.readFileSync(k.道, 'utf8');
    if (s.indexOf(印) < 0) { 守れず.push(k.名); continue; }
    守れて.push(k.名);
    const r = 突き合わせ(k.道);
    if (r.古い.length || r.無い.length) {
      悪い.push(k.名 + '\n          ★変わった 材料★ ' + (r.古い.join(' ') || 'なし')
        + (r.無い.length ? '\n          ★消えた 材料★ ' + r.無い.join(' ') : ''));
    }
  }
  console.log('      … 紙 ' + 全.length + '枚 ／ ★指紋 在り ' + 守れて.length + '枚★ ／ '
    + '指紋 無し ' + 守れず.length + '枚（★昔の 紙＝まだ 守れて いません★）');
  if (悪い.length) {
    throw new Error('★' + 悪い.length + '枚が 古い★\n        ' + 悪い.join('\n        ')
      + '\n        ⇒★その 紙を 書き出す 道具を もう1回 走らせて ください★'
      + '\n        ⇒★「取った日」は「走らせた 証し」では ありません★');
  }
});

T('★② 指紋の 行が 読める 形★（★字が 崩れたら 黙って 素通りする★）', () => {
  let 行数 = 0;
  for (const k of 紙ら()) {
    const s = fs.readFileSync(k.道, 'utf8');
    for (const 行 of s.split('\n')) {
      if (行.indexOf(印) !== 0) continue;
      行数++;
      const 割 = 行.split('\t');
      if (割.length !== 3) throw new Error(k.名 + ' … ★列が ' + 割.length + '（3 のはず）★');
      if (!/^[0-9a-f]{16}$/.test(割[2].trim())) {
        throw new Error(k.名 + ' … ★指紋の 形が 違う★ ' + 割[2].trim().slice(0, 20));
      }
    }
  }
  if (!行数) throw new Error('★指紋の 行が 1本も 無い＝空振り★');
  console.log('      … 指紋の 行 ' + 行数 + '本 とも 形は 正しい');
});

/* ══ ★★一番 守りたい 本が 入って いるかを ★名指しで★ 確かめる★★ ══
     ★なぜ 要るか★
       指紋は ★字で 追って★ 集めて います（`docs/measured/○○.mjs` を 拾う）。
       ★字で 追う 手は 必ず 抜けます★
         道を 組み立てる／別の 置き場／引数で 渡る 道 …
       ★抜けたら ★紙は 古いのに 緑★★＝★今日の XIRR と 同じ 形★
     ★実際に 1回 踏んで います★（2026-09-15）
       最初の 作りでは ★`honban-no-michi.mjs` が 1行も 入って いませんでした★
       ＝★`smartRounding:false` と プラグイン 8本を 決める 本★
     ⇒★★集め方が 抜けても ここで 止まります★★ */
T('★★③ 一番 守りたい 本が 指紋に 入って いる★★（★名指し★）', () => {
  /* ★この 紙は これを 読んで 取ったはず★（道具ごとに 決める） */
  const 必ず = {
    'golden-86-karimono-2026-09-15.tsv': [
      'docs/measured/honban-no-michi.mjs',          /* ★建て方（smartRounding／プラグイン 8本）★ */
      'docs/measured/osu-86-karimono-no-ima.mjs',   /* ★書き出した 道具そのもの★ */
      'docs/measured/zairyou-no-yubimon.mjs',       /* ★指紋を 作る 本★ */
      'exally-formula.js',                          /* ★自前層★ */
      'hyperformula.full.min.js',                   /* ★借り物★ */
    ],
  };
  const 悪い = [];
  let 見た = 0;
  for (const [名, 一覧] of Object.entries(必ず)) {
    const p = path.join(紙置き, 名);
    if (!fs.existsSync(p)) { 悪い.push(名 + ' … ★紙が 無い★'); continue; }
    見た++;
    const s = fs.readFileSync(p, 'utf8');
    for (const f of 一覧) {
      if (s.indexOf(印 + f + '\t') < 0) {
        悪い.push(名 + '\n          ★指紋に 無い★ ' + f);
      }
    }
  }
  console.log('      … 名指しで 見た 紙 ' + 見た + '枚 ／ 必ず 入って いるべき 本 '
    + Object.values(必ず).reduce((a, x) => a + x.length, 0) + '本');
  if (悪い.length) {
    throw new Error('★一番 守りたい 本が 指紋に 入って いません★\n        ' + 悪い.join('\n        ')
      + '\n        ⇒★集め方（字で 追う 手）が 抜けて います★'
      + '\n        ⇒★この ままだと 「紙は 古いのに 緑」に なります★');
  }
});

if (process.argv.includes('--self-test')) {
  console.log('\n[kami-no-zairyou --self-test] ★わざと 壊したら 赤に なるか★');

  T('★★③ 材料が 1バイト 変わったら 気づく★★', () => {
    /* ★写しで 試します★＝★本体は 1バイトも 触りません★
       （本体で 試すと ★repo の lib を 書き換える★事に なる） */
    const 仮 = path.join(紙置き, 'zz-kami-no-zairyou-tameshi.tsv');
    const 的 = 'docs/measured/zairyou-no-yubimon.mjs';
    fs.writeFileSync(仮, [
      '# ためし',
      印 + 的 + '\t' + '0'.repeat(16),      /* ★わざと 違う 指紋★ */
      'a\tb',
    ].join('\n'), 'utf8');
    try {
      const r = 突き合わせ(仮);
      if (r.見た.length !== 1) throw new Error('★1本 見るはず★ … ' + r.見た.length);
      if (r.古い.length !== 1) throw new Error('★古いと 言うはず★ … ' + JSON.stringify(r));
      console.log('      … 違う 指紋を ★古い★と 言った');
    } finally { fs.unlinkSync(仮); }
  });

  const crypto = await import('node:crypto');
  T('★④ 合って いる 指紋は 通す★（★要らない 赤を 出さない★）', () => {
    const 仮 = path.join(紙置き, 'zz-kami-no-zairyou-tameshi2.tsv');
    const 的 = 'docs/measured/zairyou-no-yubimon.mjs';
    const h = crypto.createHash('sha256')
      .update(fs.readFileSync(path.join(ROOT, 的))).digest('hex').slice(0, 16);
    fs.writeFileSync(仮, ['# ためし', 印 + 的 + '\t' + h, 'a\tb'].join('\n'), 'utf8');
    try {
      const r = 突き合わせ(仮);
      if (r.古い.length) throw new Error('★合って いるのに 古いと 言った★');
      console.log('      … 合って いる 指紋は 通した');
    } finally { fs.unlinkSync(仮); }
  });

  T('★⑤ 消えた 材料も 気づく★', () => {
    const 仮 = path.join(紙置き, 'zz-kami-no-zairyou-tameshi3.tsv');
    fs.writeFileSync(仮, ['# ためし', 印 + 'lib/mada-nai-hon.js\t' + '0'.repeat(16)].join('\n'), 'utf8');
    try {
      const r = 突き合わせ(仮);
      if (r.無い.length !== 1) throw new Error('★消えたと 言うはず★ … ' + JSON.stringify(r));
      console.log('      … 消えた 材料を ★消えた★と 言った');
    } finally { fs.unlinkSync(仮); }
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
