/* chuki.mjs — ★注記(コメント)を外してから読む★（見張りの共通部品）
 *
 *  ★なぜ在るか（指示役 2026-08-26）★
 *    ★同じ型を3回 踏んだ★ので 決まりにする。
 *      1回目 … 見張りが ★注記に書き写した字★を「配信物に在る」と読んで 嘘の赤（2026-08-22）
 *      2回目 … 検査が ★7,700行の最初の catch★（注記の中）を見ていた（2026-08-25）
 *      3回目 … 当てる道具が ★注記に書いた env:'prod'★ を 本物と読んだ（2026-08-26）
 *    ⇒ ★1本ずつ直すと 4回目が来る★。★外す所を 1か所にして 配る★。
 *
 *  ★注意（ここが 素朴なやり方だと 壊れる所）★
 *    ・`'https://example.com'` の `//` は ★注記ではない★（字の中）
 *    ・`'/* ではない *​/'` … 字の中の /* も 注記ではない
 *    ・正規表現の中の / も 注記ではない（★行の頭の // だけ を消す やり方では 取りこぼす★）
 *    ⇒ ★1文字ずつ読んで「今 字の中か」を持ちながら 外す★
 *
 *  使い方:
 *    import { 注記を外す } from './lib/chuki.mjs';
 *    const 動く所 = 注記を外す(src);              // JS / CSS
 *    const 中身   = 注記を外す(sql, { sql: true }); // SQL（-- も外す）
 *    const 中身   = 注記を外す(html, { html: true }); // HTML（<!-- --> も外す）
 */

/**
 * @param {string} src
 * @param {{sql?:boolean, html?:boolean, 残す?:boolean}} opt
 *   sql   … -- から行末までも 外す
 *   html  … <!-- --> も 外す
 *   残す  … 外した所を 空白ではなく 同じ長さの空白で埋める（行と桁を ずらさない）＝既定で true
 * @returns {string} 注記を外した字
 */
export function 注記を外す(src, opt) {
  opt = opt || {};
  const s = String(src == null ? '' : src);
  const 埋める = (opt.残す === undefined) ? true : !!opt.残す;
  if (!opt.html) return 素で外す(s, { sql: !!opt.sql, 埋める: 埋める });

  /* ★HTML の時★
     ・<script> の中は ★JSとして★ 見る（// も外す）
     ・<style> の中は ★/* … *​/ だけ★
     ・それ以外の本文は ★触らない★
       ＝本文の「don't」のような ひとつだけの ' で 字の中に入ったままになるため。
     ★<script も <style も無い字（.js / .css / JSの切れ端）は JSとして見る★
       （そうしないと 何も外れず、見張りが 注記まで読む＝元の穴に戻る） */
  /* ★<!-- --> は いつでも 先に外す★（<script も <style も無い字でも 注記は注記） */
  let out = s.replace(/<!--[\s\S]*?-->/g, (m) => 空にする(m, 埋める));
  if (out.indexOf('<script') < 0 && out.indexOf('<style') < 0) {
    return 素で外す(out, { sql: !!opt.sql, 埋める: 埋める });
  }
  out = out.replace(/(<style[^>]*>)([\s\S]*?)(<\/style>)/gi,
    (m, a, body, b) => a + 素で外す(body, { だけブロック: true, 埋める: 埋める }) + b);
  out = out.replace(/(<script[^>]*>)([\s\S]*?)(<\/script>)/gi,
    (m, a, body, b) => a + 素で外す(body, { 埋める: 埋める }) + b);
  return out;
}

/** ★1文字ずつ読んで「今 字の中か」を持ちながら 外す★（JS / CSS / SQL） */
function 素で外す(src, o) {
  o = o || {};
  const s = String(src);
  const 埋める = (o.埋める === undefined) ? true : !!o.埋める;
  const 行の注記 = !o.だけブロック;
  let out = '';
  let i = 0;
  const n = s.length;
  while (i < n) {
    const c = s[i];
    const c2 = s[i + 1];
    /* ── 字の中（そのまま通す）── */
    if (c === "'" || c === '"' || c === '`') {
      /* ★' や " は 同じ行の中でしか 閉じられない★（JSの決まり）。
         ⇒ ★その行に 閉じる相手が居なければ 字ではない★＝正規表現の中の ' などで
            ★ずっと字の中に居るつもり★になって 注記を落とし損ねるのを 防ぐ。
            （2026-08-26 実際に踏んだ：book.html の注記の中の1行が 残った） */
      const 端 = 字の終わり(s, i);
      if (端 > i) { out += s.slice(i, 端); i = 端; continue; }
      /* 閉じていない＝字ではない。1文字として そのまま通す */
      out += c; i++;
      continue;
    }
    if (c === '/' && c2 === '*') {
      const 終 = s.indexOf('*/', i + 2);
      const 端 = (終 < 0) ? n : 終 + 2;
      out += 空にする(s.slice(i, 端), 埋める); i = 端; continue;
    }
    if (行の注記 && c === '/' && c2 === '/') {
      const 終 = s.indexOf(String.fromCharCode(10), i);
      const 端 = (終 < 0) ? n : 終;
      out += 空にする(s.slice(i, 端), 埋める); i = 端; continue;
    }
    if (o.sql && c === '-' && c2 === '-') {
      const 終 = s.indexOf(String.fromCharCode(10), i);
      const 端 = (終 < 0) ? n : 終;
      out += 空にする(s.slice(i, 端), 埋める); i = 端; continue;
    }
    out += c;
    i++;
  }
  return out;
}

/** ★字（'…' "…" `…`）の終わりを返す。閉じていなければ -1 と同じ扱い（開始位置を返す）★
 *  ' と " は ★同じ行の中★でしか閉じられない。` は 何行でも またげる。 */
function 字の終わり(s, i) {
  const q = s[i];
  const 行末 = (q === '`') ? s.length : (s.indexOf(String.fromCharCode(10), i) < 0 ? s.length : s.indexOf(String.fromCharCode(10), i));
  let j = i + 1;
  while (j < 行末) {
    if (s[j] === String.fromCharCode(92)) { j += 2; continue; }
    if (s[j] === q) return j + 1;
    j++;
  }
  return i;   /* 閉じていない */
}

/** 外した所を「同じ長さの空白」にする（★行と桁を ずらさない★＝場所を出す検査が狂わない） */
function 空にする(部分, 埋める) {
  if (!埋める) return '';
  let out = '';
  for (const ch of 部分) out += (ch === '\n' || ch === '\r') ? ch : ' ';
  return out;
}

/** ★その字が 注記の外に 在るか★（いちばん よく使う形） */
export function 動く所に在る(src, 探す字, opt) {
  return 注記を外す(src, opt).indexOf(探す字) >= 0;
}
