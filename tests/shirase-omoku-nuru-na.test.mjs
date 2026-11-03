/* shirase-omoku-nuru-na.test.mjs — ★知らせの 箱を 濃い色で 塗らない★（2026-09-10）
 *
 *  ★★司さん（2026-09-10・電話の 絵つき）★★
 *    「全アプリで こんな 濃い色 使うなって 言うてなかったか？」
 *    「色が 濃いすぎるし 背景ボックスの 使い方が 悪くないか？」
 *    ⇒ 直した 後 「★絶対 これが ええ★」
 *    ⇒「Exally や 他の アプリで ★前みたいな 重たい感じに なってる所★ あったら ★先に 直せ★」
 *
 *  ★★調べたら 答えは うちの 中に 既に 在りました★★
 *    ★注意の 知らせ（`#toast.warn`）は 2026-08-25 に 直して 在る★
 *      その時 私は こう 書いて います →
 *      「濃い緑＋白文字＋半透明は、下のグリッドが 透けて 読みにくい。
 *        紙のように 白地・濃い文字にして、行頭を 揃える。左の 橙帯で 警告と 分かる」
 *    ★普通の 知らせだけ その 直しから 取り残されて いました★
 *    ⇒★作る道が 2本 在る時は 両方 直せ★（同じ 型を また 踏んだ）
 *
 *  ★★実測（★絵を 撮って 点を 数えた★）★★
 *    箱の 中 330×148 ＝ 48,840点 の うち ★暗い 点（明るさ 170未満）★
 *      前 … ★42,780点（88%）★ ／ 後 … ★2,900点（6%）★＝★14.8分の1★
 *    絵 … `docs/measured/e-toast-mae-ato-2026-09-10.png`
 *
 *  ★★この 見張りが 守る 事★★
 *    ★お客さんの 画面に かぶせる 知らせは ★白地★★（濃い色は ★左の 帯★だけ）
 *    ⇒★色を 禁じて いません★＝★大きく 塗る のを 禁じて います★
 *      （禁止色は 別の 見張り `tests/no-dark-green.test.mjs` が 持って います。ここに 在ったのは
 *        ★司さんが 2026-08-02 に 選んだ #2E7D54★＝★色では なく 塗る 広さの 話★）
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★ボタン・選ばれている タブは 見て いません★（実Excel も 塗る／小さい）
 *    ・★壊れた時の 赤い 帯★（`#B3261E`）と ★計算していない 印★（`#C0392B`）は ★わざと 強い★
 *    ・★画面 全体を 暗くする 覆い（モーダルの 黒・45か所）は ★まだ 直して いません★★
 *      ＝実Excel は ★ダイアログを 出しても シートを 暗くしません★／★司さんの 決め待ち★
 *    ・★他の repo は この 見張りの 見る 範囲では ありません★（各repo に 同じ物を 置く）
 *
 *  使い方: node tests/shirase-omoku-nuru-na.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★見る 範囲を 先に 数えて 書く★（2026-09-08 の 決まり） */
const 見る = ['book.html', 'hub.html', 'chat.html', 'css/hub.css', 'lib/ribbon.css'];

function 明るさ(r, g, b) { return 0.299 * r + 0.587 * g + 0.114 * b; }

/* ★改行を 逃がし（バックスラッシュ n）で 書きません★
   ＝2026-09-07/10 に ★heredoc で 生の 改行に 化ける★のを 4回 踏んだ */
const 改行と字下げ = String.fromCharCode(10) + '      ';

/* ★色は ★文字で 探さず 値に 直す★★（2026-08-10 の 決まり）
   ①#RRGGBB ②rgb()/rgba() の 2通り。★# 無し と 0〜1の 小数は CSS の 塗りには 出ない★ */
function 色に直す(v) {
  const h = /#([0-9a-fA-F]{6})\b/.exec(v);
  if (h) {
    const x = h[1];
    return { c: [parseInt(x.slice(0, 2), 16), parseInt(x.slice(2, 4), 16), parseInt(x.slice(4, 6), 16)], a: 1 };
  }
  const r = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?/.exec(v);
  if (r) return { c: [+r[1], +r[2], +r[3]], a: r[4] === undefined ? 1 : parseFloat(r[4]) };
  return null;
}

/* ★知らせの 箱の 規則を 集める★（`#toast` / `.toast` で 始まる 規則） */
function 知らせの規則(文) {
  const 出 = [];
  const re = /(^|[\s}])((?:#toast|\.toast)[^{}\n]{0,60})\{([^{}]{0,900})\}/g;
  let m;
  while ((m = re.exec(文)) !== null) {
    /* ★`.warn` など 別の 顔は 別に 見る★ */
    出.push({ 名: m[2].trim(), 中: m[3], 位置: 文.slice(0, m.index).split('\n').length });
  }
  return 出;
}

const 全 = 見る.map((f) => {
  const p = path.join(ROOT, f);
  if (!fs.existsSync(p)) throw new Error('★見る はずの ' + f + ' が 無い★');
  return { 名: f, 字: fs.readFileSync(p, 'utf-8') };
});

console.log('\n[shirase-omoku-nuru-na] ★知らせの 箱を 濃い色で 塗らない★');
console.log('  ★見る 範囲★ … ' + 見る.length + '本 … ' + 見る.join(' / '));

T('★知らせの 箱を 見つけて いる（★空振りして いない★）★', () => {
  const n = 全.reduce((s, f) => s + 知らせの規則(f.字).length, 0);
  if (n < 3) throw new Error('★知らせの 規則が ' + n + '本しか 見つからない★');
  console.log('      … ' + n + '本の 規則');
});

T('★★知らせの 箱を 濃い色で 塗って いない（★これが 本体★）★★', () => {
  const 悪い = [];
  for (const f of 全) {
    for (const r of 知らせの規則(f.字)) {
      const n = /background(?:-color)?\s*:\s*([^;\n}]+)/.exec(r.中);
      if (!n) continue;
      const v = 色に直す(n[1]);
      if (!v) continue;
      if (v.a < 0.5) continue;                       /* ★薄い 覆いは 別の 話★ */
      const 明 = 明るさ(...v.c);
      if (明 < 170) {
        悪い.push(f.名 + ':' + r.位置 + '  ' + r.名 + '  → ' + n[1].trim()
          + '（明るさ ' + 明.toFixed(1) + '）');
      }
    }
  }
  if (悪い.length) {
    throw new Error('★' + 悪い.length + 'か所が 濃い色で 塗って いる★\n      '
      + 悪い.join('\n      ')
      + '\n      ⇒★白地に して 濃い色は ★左の 帯★だけに して ください★'
      + '\n      ⇒ 手本＝`#toast.warn`（2026-08-25）／絵＝docs/measured/e-toast-mae-ato-2026-09-10.png');
  }
  console.log('      … 濃く 塗って いる 箱 0か所');
});

T('★★濃い色は 左の 帯として 残って いる（★色を 消した わけでは ない★）★★', () => {
  /* ★★数を 書き込みません★★（2026-09-10 に 踏んだ）
     ★最初 「帯が 2か所 以上」と 書きました＝Exally の 数を 見張りに 焼き込んで いた★
     ⇒ 給与（知らせ 1本）で ★中身は 正しいのに 赤★に なった
     ⇒★決まりは「★塗りを 持つ 知らせは 帯も 持つ★」＝★数では なく 対★ */
  const 悪い = [], 帯 = [];
  for (const f of 全) {
    for (const r of 知らせの規則(f.字)) {
      if (!/background(?:-color)?\s*:/.test(r.中)) continue;
      const b = /border-left\s*:\s*(\d+)px\s+solid\s+([^;}]+)/.exec(r.中);
      if (!b) { 悪い.push(f.名 + ':' + r.位置 + '  ' + r.名 + '  ★帯が 無い★'); continue; }
      const v = 色に直す(b[2]);
      if (!v) { 悪い.push(f.名 + ':' + r.位置 + '  帯の 色が 読めない'); continue; }
      if (+b[1] > 8) { 悪い.push(f.名 + ':' + r.位置 + '  ★帯が 太すぎる（' + b[1] + 'px）★＝また 塗り面に なる'); continue; }
      帯.push({ 太さ: +b[1], 明: 明るさ(...v.c) });
    }
  }
  if (悪い.length) throw new Error('★' + 悪い.length + 'か所★' + 改行と字下げ + 悪い.join(改行と字下げ));
  if (!帯.length) throw new Error('★帯が 1つも 無い★＝★色が 消えて しまって いる★');
  console.log('      … 塗りを 持つ 知らせ ' + 帯.length + 'か所 とも 帯つき（太さ '
    + [...new Set(帯.map((x) => x.太さ))].join('/') + 'px）');
});

T('★白地に 白い 字を 置いて いない（★読めなく なって いない★）★', () => {
  const 悪い = [];
  for (const f of 全) {
    for (const r of 知らせの規則(f.字)) {
      const n = /background(?:-color)?\s*:\s*([^;\n}]+)/.exec(r.中);
      const c = /(?:^|[;\s])color\s*:\s*([^;\n}]+)/.exec(r.中);
      if (!n || !c) continue;
      const bv = 色に直す(n[1]), cv = 色に直す(c[1]);
      if (!bv || !cv) continue;
      if (明るさ(...bv.c) > 200 && 明るさ(...cv.c) > 200) {
        悪い.push(f.名 + ':' + r.位置 + '  ' + r.名);
      }
    }
  }
  if (悪い.length) throw new Error('★白地に 薄い 字が ' + 悪い.length + 'か所★\n      ' + 悪い.join('\n      '));
  console.log('      … 白地に 薄い 字 0か所');
});

T('★★「モーダルの 黒い 覆いは まだ」の 断りが 残って いる★★', () => {
  const s = fs.readFileSync(path.join(ここ, 'shirase-omoku-nuru-na.test.mjs'), 'utf-8');
  if (!/覆い（モーダルの 黒・\d+か所）は ★まだ 直して いません★/.test(s)) {
    throw new Error('★未完の 断りが 消えた★');
  }
  /* ★実際に まだ 在る事を 数えて 見せる★（★緑に しない★） */
  let n = 0;
  for (const f of 全) {
    const m = f.字.match(/rgba\(0,\s*0,\s*0,\s*0?\.(4[5-9]|[5-9]\d?)\)/g);
    if (m) n += m.length;
  }
  if (!n) throw new Error('★黒い 覆いが 1つも 無い★＝★断りの 方が 古い（直したのに 未完と 書いて 在る）★');
  console.log('      … 黒い 覆いは まだ ' + n + 'か所（★司さんの 決め待ち★）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  const 判じ = (字) => {
    for (const r of 知らせの規則(字)) {
      const n = /background(?:-color)?\s*:\s*([^;\n}]+)/.exec(r.中);
      if (!n) continue;
      const v = 色に直す(n[1]);
      if (v && v.a >= 0.5 && 明るさ(...v.c) < 170) return true;
    }
    return false;
  };

  T('★★前の 塗り（rgba(46,125,84,0.94)）に 戻すと 赤に なる★★', () => {
    const 写し = 全[0].字.replace('background:#fff;\ncolor:#333333;', 'background:rgba(46,125,84,0.94);\ncolor:#fff;');
    if (写し === 全[0].字) throw new Error('★写しを 壊せて いない★＝★この 自己試験は 何も 見て いない★');
    if (!判じ(写し)) throw new Error('★前の 塗りに 戻しても 赤に ならない★');
    console.log('      … 戻すと 見つかる');
  });

  T('★★ハブの 塗り（#2E7D54）に 戻すと 赤に なる★★', () => {
    const hub = 全.find((f) => f.名 === 'css/hub.css');
    const 写し = hub.字.replace('background: #fff; color: #333333;', 'background: #2E7D54; color: #fff;');
    if (写し === hub.字) throw new Error('★写しを 壊せて いない★');
    if (!判じ(写し)) throw new Error('★戻しても 赤に ならない★');
    console.log('      … 戻すと 見つかる');
  });

  T('★★帯を 1本 外したら 赤に なる（★見逃す側★を 塞いだ 証拠）★★', () => {
    /* ★★2026-09-10 監査（経営者）が 見つけた 穴★★
       前は `if (帯.length < 2)` と 書いて いました。
       ⇒★2本 帯が 在れば、帯の 無い 知らせが 何本 混じって いても 緑★
         ＝★守りたい物（塗りだけで 帯が 無い 知らせ）を 通して しまう★
       ⇒★数の 合計では なく ★塗りと 帯の 対★で 見る★に 直した
       ⇒ この 自己試験は ★その 直しが 効いて いる 証拠★ */
    const 数える = (字) => {
      const 悪い = [];
      for (const r of 知らせの規則(字)) {
        if (!/background(?:-color)?\s*:/.test(r.中)) continue;
        if (!/border-left\s*:\s*\d+px\s+solid/.test(r.中)) 悪い.push(r.名);
      }
      return 悪い;
    };
    /* ★塗りも 帯も 持つ 知らせが 2本、その 上に ★帯の 無い 3本目★ を 足す★ */
    const 二本 = '#toast{position:fixed;background:#fff;border-left:5px solid #2E7D54;}'
      + ' .toast{position:fixed;background:#fff;border-left:5px solid #2E7D54;}';
    if (数える(二本).length) throw new Error('★正しい 2本を 悪いと 言って いる★');
    const 三本目 = 二本 + ' .toast-shirase{position:fixed;background:#2E7D54;color:#fff;}';
    const 出 = 数える(三本目);
    if (!出.length) {
      throw new Error('★帯の 無い 知らせを 足しても 赤に ならない★'
        + '＝★前の「2本 在れば 緑」に 戻って いる★');
    }
    console.log('      … 帯の 無い 知らせを 1本 混ぜると 赤（' + 出.join('/') + '）');
  });

  T('★★薄い 覆い（alpha 0.4）は 赤に しない（★狼少年に しない★）★★', () => {
    const 写し = '.toast{position:fixed;background:rgba(0,0,0,0.4);color:#fff;}';
    if (判じ(写し)) throw new Error('★薄い 覆いまで 赤に して いる★');
    console.log('      … 薄い 覆いは 通す');
  });

  T('★★色を 文字で 探して いない（値に 直して いる）★★', () => {
    /* ★同じ 色を 3通りの 書き方で 書いて 全部 捕まえるか★ */
    for (const 書き方 of ['#2E7D54', 'rgb(46,125,84)', 'rgba(46, 125, 84, 0.94)']) {
      const 写し = '.toast{position:fixed;background:' + 書き方 + ';color:#fff;}';
      if (!判じ(写し)) throw new Error('★' + 書き方 + ' を 見落とした★');
    }
    /* ★小文字でも★ */
    if (!判じ('.toast{background:#2e7d54;}')) throw new Error('★小文字を 見落とした★');
    console.log('      … 4通りの 書き方 とも 捕まえる');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
