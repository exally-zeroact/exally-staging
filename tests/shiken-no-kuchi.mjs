/* shiken-no-kuchi.mjs — ★実の倉庫に触る道具が使う「試験用の口」の認証と門★（2026-10-10）
 *
 *  ★なぜ★ 受信箱のメールが repo（公開）の コードと 覚書に 書いて あった（司さんの決め 10-10＝repo の 外へ）。
 *    メールと パスワードは repo の 外の cred ファイルだけに 置く。repo には 書かない。
 *  ★置き場★ 環境変数 EXALLY_TEST_CRED、無ければ ~/.exally-e0-test-cred.json（%TEMP% には 置かない＝消える・10-02 の 教え）
 *    中身は { "email": "...", "password": "..." }。作るのは 司さんの 手番（試験用の 口は 作り直さない）。
 *  ★門★ メールに +タグ が 在り、その タグが 白名簿に 在る 時だけ 走る。
 *    （メールの 字を repo に 置かずに「試験用の 口か」を 見る。sha256 を 置く 形は 名前から 元の 箱が 分かるので 使わない）
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export const 許すタグ = ['e0test'];
export const CRED_FILE = process.env.EXALLY_TEST_CRED || path.join(os.homedir(), '.exally-e0-test-cred.json');

/** 試験用の 口か（メールの +タグ が 白名簿に 在るか） */
export function 試験の口か(email) {
  const m = /^[^\s@+]+\+([A-Za-z0-9_-]+)@[^\s@]+\.[^\s@]+$/.exec(String(email || ''));
  return !!m && 許すタグ.includes(m[1]);
}

/** cred を 読む ⇒ { email, password }。無い・形が 違う・試験用の 口で ない 時は 投げる（走らせない） */
export function 試験の口を読む() {
  let cred;
  try { cred = JSON.parse(fs.readFileSync(CRED_FILE, 'utf8')); }
  catch (e) { throw new Error('試験用の 口の cred が 読めない（' + CRED_FILE + '）＝司さんの 手番で 置く。走らせない'); }
  if (!cred || typeof cred.email !== 'string' || typeof cred.password !== 'string') throw new Error('cred の 形が 違う（email と password）');
  if (!試験の口か(cred.email)) throw new Error('中止: 試験用の 口（+タグ が 白名簿）で ない。走らせない');
  return cred;
}
