/* souname-matome.mjs — ★束に 分けて 流した 総なめを 「読む 側」で 判じる★（2026-10-04・経営者と 決めた）
 *
 *  ★★なぜ 在るか★★
 *    `run-tests-batch.mjs` の 束は 途中で 切られる（timeout・メモリで 止められる）と
 *    最後の 行（頼んだ／走らせた／緑）が 出ない まま 終わる。
 *    ⇒ 出しが 空や 途中までに なり、★読む 側が 「赤 0」と 見分けられなかった★（10-04 に 1回 見間違えかけた）。
 *    ⇒★読む 側を 道具に した★。★記録が 無い 束＝走り切っていない＝赤★。
 *
 *  ★★使い方★★
 *    node scripts/souname-matome.mjs --hajime            ... 回の 札を 作る（古い 記録は 全部 消す）
 *    node scripts/run-tests-batch.mjs 1 130 --回          ... 束ごと（走り切った 時だけ 記録を 書く）
 *    node scripts/run-tests-batch.mjs 131 250 --回 ...
 *    node scripts/souname-matome.mjs                      ... ★判じる★
 *
 *  ★★赤に する 物（1つでも 当たれば 赤）★★（経営者の 叩き ⑴〜⑸）
 *    ・回の 札が 無い／今の 頭・名簿の 指紋が 札と 違う／今の 木が 空で ない
 *    ・★名簿の 1〜N を 隙間・重なり 無く 覆って いない★（記録の 無い 束＝隙間）
 *    ・束の 記録が ★この 回の 名札で ない★（前の 回の 完走は 数えない）・頭／名簿の 指紋が 違う
 *    ・束を ★回の 札より 前に★ 始めた／頼んだ ≠ 走らせた／赤が 在る
 *    ・★束の 前か 後の 木が 空で ない★（自己試験が 途中で 落ちると 壊した 値が 残る）
 *
 *  ★★まとめ 自身にも 同じ 穴が 在る★★（経営者 ⑸）
 *    ＝★この 道具の 出しの 最後の 行「名簿 N／覆った M／緑 G」が 無ければ 赤★・★exit が 0 で なければ 赤★
 *    （読む 人が 守る 決まり。引き継ぎにも 書く）
 *
 *  ★置き場★ ＝ repo の 外（一時の フォルダ）＝★「木が 空か」を 自分で 汚さない★
 */
import path from 'node:path';
import fs from 'node:fs';
import os from 'node:os';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

export function 回の置き場() {
  return process.env.EXALLY_SOUNAME_DIR || path.join(os.tmpdir(), 'exally-souname');
}
/** ★名簿の 中身の 指紋★（本数が 同じで 中身が 入れ替わった 時を 捕まえる） */
export function 名簿の指紋(FILES) {
  return crypto.createHash('sha256').update(JSON.stringify(FILES)).digest('hex').slice(0, 16);
}
export function 頭を取る(root) {
  return execFileSync('git', ['-C', root, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
}
/** ★木の 汚れ★（`git status --porcelain` の 字。空なら 綺麗） */
export function 木を取る(root) {
  return execFileSync('git', ['-C', root, 'status', '--porcelain'], { encoding: 'utf8' }).trim();
}

/** ★★判じる★★（試験から 呼べる ように 純な 関数に した）
 *  @param 回 回の 札（無ければ null）
 *  @param 束たち 束の 記録の 並び（読めなかった 物は { 読めない: 名 } ）
 *  @param 今 { 頭, 名簿の指紋, 名簿, 木 }
 *  @returns { 赤: [訳], 名簿, 覆った, 緑 } */
export function 判じる(回, 束たち, 今) {
  const 赤 = [];
  if (!回) return { 赤: ['★回の 札が 無い★（--hajime を 先に）'], 名簿: 今.名簿, 覆った: 0, 緑: 0 };
  if (今.頭 !== 回.頭) 赤.push('★今の 頭が 回の 札と 違う★（' + String(今.頭).slice(0, 7) + ' ／ 札 ' + String(回.頭).slice(0, 7) + '）');
  if (今.名簿の指紋 !== 回.名簿の指紋) 赤.push('★今の 名簿の 指紋が 回の 札と 違う★');
  if (今.木) 赤.push('★今の 木が 空で ない★');
  const 印 = new Array(今.名簿 + 1).fill(0);
  let 緑 = 0;
  for (const b of 束たち) {
    if (b.読めない) { 赤.push('★束の 記録を 読めない★ ' + b.読めない); continue; }
    const 名 = b.始 + '〜' + b.終;
    if (b.名札 !== 回.名札) { 赤.push('★' + 名 + ' は この 回の 記録で ない★（前の 回の 残り＝数えない）'); continue; }
    let 悪い = false;
    if (b.頭 !== 回.頭) { 赤.push('★' + 名 + ' の 頭が 違う★'); 悪い = true; }
    if (b.名簿の指紋 !== 回.名簿の指紋) { 赤.push('★' + 名 + ' の 名簿の 指紋が 違う★'); 悪い = true; }
    if (!(String(b.始めた) >= String(回.始めた))) { 赤.push('★' + 名 + ' は 回の 札より 前に 始めた★'); 悪い = true; }
    if (b.頼んだ !== b.走らせた) { 赤.push('★' + 名 + ' 頼んだ ' + b.頼んだ + ' ≠ 走らせた ' + b.走らせた + '★'); 悪い = true; }
    if (b.赤 && b.赤.length) { 赤.push('★' + 名 + ' に 赤 ' + b.赤.length + '件★ ' + b.赤.join(' ／ ')); 悪い = true; }
    if (b.前の木) { 赤.push('★' + 名 + ' の 前の 木が 空で ない★'); 悪い = true; }
    if (b.後の木) { 赤.push('★' + 名 + ' の 後の 木が 空で ない★（自己試験の 壊した 値が 残った かも）'); 悪い = true; }
    for (let i = b.始; i <= b.終; i++) { if (i >= 1 && i <= 今.名簿) 印[i]++; }
    if (!悪い) 緑 += b.緑;
  }
  let 覆った = 0; const 隙間 = [], 重なり = [];
  for (let i = 1; i <= 今.名簿; i++) {
    if (印[i] === 0) 隙間.push(i); else 覆った++;
    if (印[i] > 1) 重なり.push(i);
  }
  if (隙間.length) 赤.push('★覆って いない 本 ' + 隙間.length + '★（' + 区間(隙間) + '）＝記録の 無い 束＝走り切って いない');
  if (重なり.length) 赤.push('★重なった 本 ' + 重なり.length + '★（' + 区間(重なり) + '）');
  return { 赤, 名簿: 今.名簿, 覆った, 緑 };
}
function 区間(並び) {
  const 出 = []; let s = 並び[0], p = 並び[0];
  for (let k = 1; k <= 並び.length; k++) {
    const x = 並び[k];
    if (x === p + 1) { p = x; continue; }
    出.push(s === p ? String(s) : s + '〜' + p); s = x; p = x;
  }
  return 出.join(',');
}

/* ══ ★走らせ方★ ══（試験から import された 時は 走らない） */
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
  const { FILES } = require_(path.join(ROOT, 'tests/run.js'));
  const 置き場 = 回の置き場();
  const 今 = { 頭: 頭を取る(ROOT), 名簿の指紋: 名簿の指紋(FILES), 名簿: FILES.length, 木: 木を取る(ROOT) };
  if (process.argv.includes('--hajime')) {
    fs.mkdirSync(置き場, { recursive: true });
    for (const f of fs.readdirSync(置き場)) fs.unlinkSync(path.join(置き場, f));   /* ★古い 記録は 全部 消す★ */
    const 回 = { 名札: crypto.randomBytes(8).toString('hex'), 頭: 今.頭, 名簿の指紋: 今.名簿の指紋,
      名簿: 今.名簿, 始めた: new Date().toISOString() };
    const 仮 = path.join(置き場, '回.json.kaki');
    fs.writeFileSync(仮, JSON.stringify(回, null, 1));
    fs.renameSync(仮, path.join(置き場, '回.json'));
    console.log('★回の 札を 作った★ 名札 ' + 回.名札 + ' ／ 頭 ' + 回.頭.slice(0, 7) + ' ／ 名簿 ' + 回.名簿 + '本'
      + (今.木 ? ' ／ ★今の 木が 空で ない★' : ''));
    process.exit(今.木 ? 1 : 0);
  }
  let 回 = null;
  const 束たち = [];
  if (fs.existsSync(path.join(置き場, '回.json'))) {
    回 = JSON.parse(fs.readFileSync(path.join(置き場, '回.json'), 'utf8'));
    for (const f of fs.readdirSync(置き場).filter((x) => /^束-\d+-\d+\.json$/.test(x)).sort()) {
      try { 束たち.push(JSON.parse(fs.readFileSync(path.join(置き場, f), 'utf8'))); }
      catch (e) { 束たち.push({ 読めない: f }); }
    }
  }
  const r = 判じる(回, 束たち, 今);
  console.log('[souname-matome] 束 ' + 束たち.length + '本 ／ 頭 ' + 今.頭.slice(0, 7));
  for (const s of r.赤) console.log('  ' + s);
  /* ★★最後の 行★★＝★これが 無い 出しは 赤★（読む 人の 決まり） */
  console.log((r.赤.length ? '★赤★ ' : '★緑★ ') + '名簿 ' + r.名簿 + '／覆った ' + r.覆った + '／緑 ' + r.緑);
  process.exit(r.赤.length ? 1 : 0);
}
