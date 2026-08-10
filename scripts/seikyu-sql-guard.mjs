/* seikyu-sql-guard.mjs — ★倉庫に当てる前の門番★（請求書の設計図だけを通す）
 *
 * なぜ必要か:
 *   倉庫には Kyually本番 / Exally / 代行請求 / 飲み屋 / アマかせ の実データが同居している。
 *   当てる道具に「何でも流せる口」を作ると、事故の桁が変わる。
 *   ★1本でも門番に落ちたら、1文字も当てない★（部分適用をしない＝中途半端な棚を作らない）
 *
 * 通す物（これ以外は全部 止める）:
 *   ・部屋      kyuyo のみ（＋窓口を置く public、参照だけの auth.users）
 *   ・棚/窓口    pay_invoices / pay_receipts のみ
 *   ・仕掛け    pay_invoices_freeze / pay_invoices_no_delete / pay_receipts_touch のみ
 *   ・窓口(view) は ★security_invoker = true が付いていること★（無ければ止める）
 *
 * 止める物:
 *   ・消す系（drop table / drop schema / truncate / delete from / update…set / drop column …）
 *     ※ drop policy / drop trigger の "if exists" は冪等のために要るので通す
 *   ・許していない棚・部屋の名前が1つでも出たら止める
 *   ・本番倉庫の名前が混ざっていたら止める
 *
 * これは純関数だけの部品（ネットにも倉庫にも触らない）＝ seikyu/tests/sql-guard.test.mjs で
 * ★わざと危ない物を食わせて、止まることを実測する★。
 */

export const ALLOWED_SCHEMA = 'kyuyo';
export const ALLOWED_TABLES = ['pay_invoices', 'pay_receipts'];
export const ALLOWED_FUNCTIONS = ['pay_invoices_freeze', 'pay_invoices_no_delete', 'pay_receipts_touch'];
export const PROD_WAREHOUSE_REF = 'tnfwipbgfgjaymlszeid'; // ★本番。ここへは当てない

/** コメントを外す（門番がコメントの文字で判断しないように） */
export function stripComments(sql) {
  return String(sql)
    .replace(/\/\*[\s\S]*?\*\//g, ' ')
    .replace(/--[^\n]*/g, ' ');
}

const DANGER = [
  [/\bdrop\s+table\b/i, 'drop table（棚を消す）'],
  [/\bdrop\s+schema\b/i, 'drop schema（部屋を消す）'],
  [/\bdrop\s+database\b/i, 'drop database'],
  [/\bdrop\s+view\b/i, 'drop view（窓口を消す）'],
  [/\bdrop\s+function\b/i, 'drop function'],
  [/\bdrop\s+column\b/i, 'drop column（列を消す）'],
  [/\bdrop\s+constraint\b/i, 'drop constraint（決まりを外す）'],
  [/\bdrop\s+index\b/i, 'drop index'],
  [/\btruncate\b/i, 'truncate（中身を全部消す）'],
  [/\bdelete\s+from\b/i, 'delete from（行を消す）'],
  [/\bupdate\s+[a-z_."]+\s+set\b/i, 'update … set（行を書き換える）'],
  [/\binsert\s+into\b/i, 'insert into（行を入れる）'],
  [/\balter\s+role\b/i, 'alter role'],
  [/\bcreate\s+role\b/i, 'create role'],
  [/\bgrant\s+all\b/i, 'grant all（権限を丸ごと渡す）'],
  [/\bsecurity\s+definer\b/i, 'security definer（持ち主の権利で動く）'],
];

/**
 * inspect(sql) → { ok, reasons[], stats }
 *   ok=false なら1文字も当てない。
 */
export function inspect(sql) {
  const reasons = [];
  const raw = String(sql == null ? '' : sql);
  const s = stripComments(raw);

  if (!s.trim()) reasons.push('中身が空です');

  // ① 消す系・危ない書き方
  for (const [re, why] of DANGER) {
    if (re.test(s)) reasons.push('消す系/危ない書き方が混ざっています: ' + why);
  }

  // ② 部屋・棚の名前（schema.table の形を全部拾って、許した物だけか見る）
  const refs = [...s.matchAll(/\b([a-z_][a-z0-9_]*)\.([a-z_][a-z0-9_]*)\b/gi)];
  const seen = new Set();
  for (const m of refs) {
    const sch = m[1].toLowerCase(), name = m[2].toLowerCase();
    // 関数の中の new./old. と、変数っぽい物は棚ではない
    if (sch === 'new' || sch === 'old' || sch === 'pg_catalog' || sch === 'information_schema') continue;
    // Supabase の作り付け。既存の棚(pay_employees 等)と同じ書き方＝ここだけは通す。
    //   auth.users … 参照だけ（誰の行かを繋ぐ）  auth.uid() … 「今ログインしている人」
    if (sch === 'auth' && (name === 'users' || name === 'uid')) continue;
    if (sch === 'jsonb' || sch === 'text') continue;
    seen.add(sch + '.' + name);
    if (sch !== ALLOWED_SCHEMA && sch !== 'public') {
      reasons.push('許していない部屋が出てきました: ' + sch + '.' + name);
      continue;
    }
    if (![...ALLOWED_TABLES, ...ALLOWED_FUNCTIONS].includes(name)) {
      reasons.push('許していない棚/仕掛けが出てきました: ' + sch + '.' + name);
    }
  }

  // ③ 部屋の指定なしで create table / create view していないか（public にこぼれる）
  const bare = [...s.matchAll(/\bcreate\s+(?:or\s+replace\s+)?(?:table|view)\s+(?:if\s+not\s+exists\s+)?([a-z_][a-z0-9_]*)\s*(?:\(|as|with)/gi)];
  for (const m of bare) {
    reasons.push('部屋の指定がありません（public にこぼれます）: create … ' + m[1]);
  }

  // ④ 窓口(view)は security_invoker = true が必須
  const views = [...s.matchAll(/\bcreate\s+(?:or\s+replace\s+)?view\s+([a-z_.]+)([\s\S]*?)\bas\b/gi)];
  for (const m of views) {
    if (!/with\s*\(\s*security_invoker\s*=\s*true\s*\)/i.test(m[2])) {
      reasons.push('★窓口に security_invoker = true が付いていません: ' + m[1] + '（付けないと全アカウントのデータが見えます）');
    }
  }

  // ⑤ 本番倉庫の名前が混ざっていないか
  if (raw.includes(PROD_WAREHOUSE_REF)) reasons.push('本番倉庫の名前が混ざっています: ' + PROD_WAREHOUSE_REF);

  return {
    ok: reasons.length === 0,
    reasons,
    stats: { objects: [...seen].sort(), views: views.length, bytes: raw.length },
  };
}
