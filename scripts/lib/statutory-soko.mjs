/* statutory-soko.mjs — ★法定の数値の 倉庫(Supabase public.statutory)を 読む★（検査から 使う）
 *
 *  ★なぜ 検査が 倉庫を 読むのか★（2026-09-05 司さん）
 *    「金関係のこと聞かれたりAIが入力する時だけSupabaseの共有から拾うやないんか」
 *    ⇒ Exally は 法定の数値の ★ファイルを 1本も 持たない★。
 *    ⇒ ★見張りも 本物と 同じ 所を 見る★＝倉庫。
 *      （前は lib を 読んでいた。lib を 消したので、そのままでは 検査が 空振りに なる）
 *
 *  ★fetch を 使わない★
 *    undici が 繋ぎっぱなしの 線を 残し、process.exit と ぶつかって node ごと 落ちる
 *    （Assertion failed: UV_HANDLE_CLOSING / src\win\async.c:76）＝2026-09-05 実測。
 *    ⇒ node:https で 取り、★agent:false で 線を 毎回 閉じる★。
 *
 *  ★読めなかった時は null★＝★呼ぶ側が 緑に しない事★（終了の印 2 で 止める）
 */
import fs from 'node:fs';
import path from 'node:path';
import https from 'node:https';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from './chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');

/** repo が 向いている 倉庫の 口（anon の 鍵＝全アプリ 共通の 棚を 読むだけ） */
export function 倉庫の口() {
  try {
    /* ★注記を外す のは 共通部品を 使う★（自前で 書かない＝tests/chuki.test.mjs の 決まり。
       素朴な 正規表現は 字の中の // や 正規表現の中の /* で 壊れる） */
    const src = 注記を外す(fs.readFileSync(path.join(ROOT, 'js/supa-config.js'), 'utf8'));
    const u = /https:\/\/[a-z0-9]+\.supabase\.co/.exec(src);
    const k = /['"](ey[A-Za-z0-9_.-]{40,}|sb_publishable_[A-Za-z0-9_-]{20,})['"]/.exec(src);
    return (u && k) ? { url: u[0], key: k[1] } : null;
  } catch (e) { return null; }
}

function 取る(u, key, ミリ秒) {
  return new Promise((resolve) => {
    let 済み = false;
    const 終わり = (v) => { if (!済み) { 済み = true; resolve(v); } };
    let req;
    try {
      req = https.get(u, { agent: false, headers: { apikey: key, Authorization: 'Bearer ' + key } }, (res) => {
        if (res.statusCode !== 200) { res.resume(); return 終わり(null); }
        let buf = '';
        res.setEncoding('utf8');
        res.on('data', (c) => { buf += c; });
        res.on('end', () => { try { 終わり(JSON.parse(buf)); } catch (e) { 終わり(null); } });
        res.on('error', () => 終わり(null));
      });
    } catch (e) { return 終わり(null); }
    req.setTimeout(ミリ秒, () => { req.destroy(); 終わり(null); });
    req.on('error', () => 終わり(null));
  });
}

/** 倉庫の 行を 全部 拾う。★読めなければ null★（前の値で ごまかさない） */
export async function 倉庫の行を拾う(ミリ秒) {
  const 口 = 倉庫の口();
  if (!口) return null;
  const j = await 取る(口.url + '/rest/v1/statutory?select=kind,year,data', 口.key, ミリ秒 || 20000);
  return Array.isArray(j) && j.length ? j : null;
}

/** ★読めなかったら その場で 止める★（緑に しない・コードの赤とは 分ける＝終了の印 2） */
export function 読めなければ止まる(行たち, 誰) {
  if (行たち) return 行たち;
  console.log('');
  console.log('★倉庫(Supabase statutory)を 読めませんでした＝' + (誰 || 'この検査') + 'は 1度も 走っていません★');
  console.log('  ★コードが 赤なのでは ありません★／ネットか 倉庫の 側です。もう一度。');
  process.exit(2);
}
