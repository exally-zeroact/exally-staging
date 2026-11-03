// ============================================================
// scripts/sql-guard.mjs
// ★本番の倉庫にSQLを当てる前の門番（純ロジック・試験の対象）2026-08-29★
//
//   ★この repo が 向いている 本番の倉庫には
//     Exally本番 / 給与(Kyually)本番 / 代行請求 の 実データが 同居している★。
//   だから ★足すだけ・締めるだけ★ 以外は 1文字も 通さない。
//
//   ★形は 飲み屋(nomiya-app)・ダイコメ(Daikou-app-test)の門番から 借りた★
//   （借りてよいのは 道具・測り方・試験。数字と画面は 各アプリの物）。
//   ★Exally用に 変えた所は 3つ★
//     ①棚の名前ではなく ★部屋(schema)で 締める★＝`exally.` で始まる物だけ
//       （倉庫はアプリごとの部屋に 分けてある。棚の頭文字では 締められない）
//     ②`revoke` を ★締める向きに限って 通す★（借り元は 全部 止めていた）
//       ＝revoke は 権限を 減らす向き。ただし ★exally. の棚に対してだけ★。
//     ③`do $$ … $$` を ★中を読んで 判定する★（借り元は 読めないので 全部 止めていた）
//       ＝2回当てても落ちない書き方（policy が在るか見てから作る）に 必要。
//       中で 走らせてよいのは ★create policy … on exally.… だけ★。
//
//   ▼止める物
//     drop / truncate / delete / update / insert /
//     exally. 以外の 部屋に 触る物 / do $$ の中に 許した物以外が 入っている物
//   ▼通す物
//     create table / create index / create policy / alter table … enable row level security /
//     comment on / grant / revoke（exally. の棚に対してのみ）/ select（確認用）
// ============================================================

// ============================================================
// ★★2026-10-02 作り直し＝★字を 読めるように した ＋ 黒名簿を やめて 白名簿に した★★★
//   （ダイコメの席が 09-29 に 同じ門で 見つけて 塞いだ 穴が Exally にも 開いていた・Exally1 が 実測）
//
//   ★開いていた 穴（直す前に 当てて 確かめた）★
//     ① ドル引用の 中の ' で ★次の 1文が まるごと 門から 消えた★
//          create table exally.dk_note (a text default $q$it's$q$);
//          update daikou.meisai set amount = 0;        ← ★通った★（他の部屋の 書き換え）
//          comment on table exally.dk_note is 'ok';
//        因＝`$$` だけ 先に 剥がし、`$q$` の 中の ' を 外の ' と 対に していた
//     ② 引用した 名前 "public"."meisai" を `""` に 潰し ★棚を 0件★ と 読んでいた
//          alter table "public"."meisai" add column hack int; select 1;   ← ★通った★
//        ＝★部屋の 壁（この 門の 本体）が 丸ごと 外れていた★
//   ★読みで 見つけた 穴（ダイコメの 直しにも 残っていそうな 物・Exally1）★
//     ③ コメントの 中に ドル引用の 印 ⇒ `-- $x$` ～ `-- $x$` の 間の 文が 「文字列」として 消える
//        ⇒★剥がしは 左から 1字ずつ 読む 1回の 読みに した★（出てきた 順に 扱う＝PG と 同じ 読み方）
//     ④ `grant ... on` の 権限名を 落とす 所が ★`;` を 越えて 次の 文を 食う★
//          grant r to u; update daikou.meisai set a = 0 where x; create index i on exally.a (x);
//        ⇒★1文の 中だけ★ に した
//   ★白名簿★＝`;` で 1文ずつ に 分け、★下の ALLOW の 形に 頭から 当たらなければ 赤★
//     その上で 今までどおり 黒名簿（DANGER）と 部屋の 壁も 通す（二重）
// ============================================================

/** ★SQL を 左から 1字ずつ 読む★（コメント・文字列・ドル引用・引用名を 出てきた 順に 扱う）
 *  返り ＝ { 字, do塊[], 読めない }
 *   ・コメント → 空白 ／ 文字列・ドル引用 → '' ／ `do` の 後の ドル引用 → ' $do$ '（中身は do塊へ）
 *   ・引用名 → 安全な 名前なら 引用を 外して 名前に 戻す／危ない 字なら _q_（どの 部屋にも 当たらない）
 *   ・★閉じて いない 物は 「読めない」★＝門は 赤に する（読めない 塊を 通さない） */
function 読む(sql) {
  const s = String(sql == null ? '' : sql);
  let 出 = '', i = 0, 読めない = '';
  const do塊 = [];
  while (i < s.length) {
    const c = s.charAt(i), c2 = s.slice(i, i + 2);
    if (c2 === '--') {
      const j = s.indexOf('\n', i);
      出 += ' '; i = (j < 0 ? s.length : j); continue;
    }
    if (c2 === '/*') {
      /* ★PG の ブロックコメントは 入れ子に なる★ */
      let 深さ = 0, j = i;
      while (j < s.length) {
        if (s.slice(j, j + 2) === '/*') { 深さ++; j += 2; continue; }
        if (s.slice(j, j + 2) === '*/') { 深さ--; j += 2; if (深さ === 0) break; continue; }
        j++;
      }
      if (深さ !== 0) { 読めない = '閉じていない コメント'; break; }
      出 += ' '; i = j; continue;
    }
    if (c === "'") {
      /* ★E'...' は 逆斜線で ' を 逃がせる★ */
      const E = /[eE]/.test(s.charAt(i - 1)) && !/[\w]/.test(s.charAt(i - 2));
      let j = i + 1, 閉 = false;
      while (j < s.length) {
        const d = s.charAt(j);
        if (E && d === '\\') { j += 2; continue; }
        if (d === "'") { if (s.charAt(j + 1) === "'") { j += 2; continue; } 閉 = true; break; }
        j++;
      }
      if (!閉) { 読めない = '閉じていない 文字列'; break; }
      出 += "''"; i = j + 1; continue;
    }
    if (c === '$') {
      /* ★タグは 英字 以外も 受ける★（PG は 名前に \u0080 以上の 字を 許す＝`$あ$` `$é$`）
         前は 英字だけ 見て、`$あ$ ' $あ$` の ' を 文字列の 始まりと 読み ★次の 1文を 食っていた★
         （ダイコメの席が 10-02 に 知らせ・Exally の 門に 当てて 通るのを 確かめてから 直した）
         ★名前の 途中の `$`★（`あ$x$`）は タグでは ない＝前の 字が 名前の 字なら 素通り */
      const t = /^\$([A-Za-z_\u0080-￿][\w\u0080-￿]*)?\$/.exec(s.slice(i));
      if (t && !/[\w$\u0080-￿]/.test(s.charAt(i - 1))) {
        const 印 = t[0];
        const j = s.indexOf(印, i + 印.length);
        if (j < 0) { 読めない = '閉じていない ドル引用 ' + 印; break; }
        const 中 = s.slice(i + 印.length, j);
        if (/\bdo\s*$/i.test(出)) { do塊.push(中); 出 += ' $do$ '; }
        else 出 += " '' ";
        i = j + 印.length; continue;
      }
    }
    if (c === '"') {
      let j = i + 1, 閉 = false, 名 = '';
      while (j < s.length) {
        if (s.charAt(j) === '"') { if (s.charAt(j + 1) === '"') { 名 += '"'; j += 2; continue; } 閉 = true; break; }
        名 += s.charAt(j); j++;
      }
      if (!閉) { 読めない = '閉じていない 引用名'; break; }
      出 += /^[A-Za-z_][A-Za-z0-9_$]*$/.test(名) ? 名 : '_q_';
      i = j + 1; continue;
    }
    出 += c; i++;
  }
  return { 字: 出, do塊, 読めない };
}

/** ★正規表現が 文字列やコメントの中身に 引っかからないように 消す★
 *  do $$ … $$ は ★中を 別に 調べる★ ので、印だけ 残して 場所が 分かるようにする。 */
export function stripNoise(sql, opts) {
  const r = 読む(sql);
  const doBlocks = r.do塊;
  let s = r.字;
  if (opts && opts.collect) opts.collect.読めない = r.読めない;
  /* ★grant/revoke の「権限の名前の並び」を 落とす★
     ＝`revoke update, truncate, references, trigger on exally.recipe from authenticated`
       の update / truncate は ★権限の名前★であって 文ではない。
     ★2026-08-29 実測★＝これを落とさないと 本物の recipe.sql が「truncate が入っている」で
     止まった。★門番を 緩めたのではなく 読み違いを 直した★
     （本物の truncate 文と do $$ の中の truncate は 試験で 止まる事を 見ている）。 */
  /* ★1文の 中だけ★（`;` を 越えると 次の 文を 食う＝上の ④・2026-10-02） */
  s = s.replace(/\b(grant|revoke)\b[^;]*?\bon\b/gi, function (m, kw) { return kw + ' on '; });
  if (opts && opts.collect) opts.collect.doBlocks = doBlocks;
  return s;
}

const DANGER = [
  { name: 'drop', re: /\bdrop\s+(table|policy|column|index|schema|view|function|trigger|type|database|role)\b/i },
  { name: 'truncate', re: /\btruncate\s+(?:table\s+)?[a-z_"]/i },
  { name: 'delete', re: /\bdelete\s+from\b/i },
  /* ★別名つきの update も★（ダイコメ 09-26＝`update 棚 k set ...` が すり抜けた） */
  { name: 'update', re: /\bupdate\s+(?:only\s+)?[a-z_][\w.]*(?:\s+(?:as\s+)?[a-z_]\w*)?\s+set\b/i },
  { name: 'merge', re: /\bmerge\s+into\b/i },
  { name: 'trigger-rule', re: /\bcreate\s+(or\s+replace\s+)?(constraint\s+)?(trigger|rule)\b/i },
  { name: 'procedure', re: /\bcreate\s+(or\s+replace\s+)?procedure\b/i },
  { name: 'security', re: /\b(alter\s+(default\s+)?privileges|set\s+(local\s+)?role|security\s+definer|create\s+role|alter\s+role)\b/i },
  { name: 'copy-from', re: /\bcopy\b[^;]*\bfrom\b/i },
  { name: 'insert', re: /\binsert\s+into\b/i },
  { name: 'alter-drop', re: /\balter\s+table\s+[^;]*\bdrop\b/i },
  { name: 'create-function', re: /\bcreate\s+(or\s+replace\s+)?function\b/i },  // 実行権が既定でPUBLIC
  { name: 'grant-all-public', re: /\bgrant\b[^;]*\bto\s+public\b/i },
];

export function findDangerous(sql) {
  const s = stripNoise(sql);
  const hits = [];
  for (const d of DANGER) {
    const m = s.match(d.re);
    if (m) hits.push({ kind: d.name, at: m[0].replace(/\s+/g, ' ').trim() });
  }
  return hits;
}

/** ★このSQLが 触る棚を 部屋つきで 全部 拾う★（exally.recipe の形） */
export function findTargetTables(sql) {
  const s = stripNoise(sql);
  const out = new Set();
  const N = '([a-z_][\\w]*(?:\\.[a-z_][\\w]*)?)';
  const pats = [
    new RegExp('\\bcreate\\s+table\\s+(?:if\\s+not\\s+exists\\s+)?' + N, 'gi'),
    new RegExp('\\balter\\s+table\\s+(?:if\\s+exists\\s+)?(?:only\\s+)?' + N, 'gi'),
    new RegExp('\\bcreate\\s+(?:unlogged\\s+)?table\\s+(?:if\\s+not\\s+exists\\s+)?' + N, 'gi'),
    new RegExp('\\bcreate\\s+(?:unique\\s+)?index\\s+(?:concurrently\\s+)?(?:if\\s+not\\s+exists\\s+)?[\\w]+\\s+on\\s+' + N, 'gi'),
    new RegExp('\\bcreate\\s+policy\\s+[\\w]+\\s+on\\s+' + N, 'gi'),
    new RegExp('\\bcomment\\s+on\\s+column\\s+' + N + '\\.', 'gi'),
    new RegExp('\\bcomment\\s+on\\s+table\\s+' + N, 'gi'),
    new RegExp('\\bgrant\\s+[^;]*?\\bon\\s+(?:table\\s+)?' + N, 'gi'),
    new RegExp('\\brevoke\\s+[^;]*?\\bon\\s+(?:table\\s+)?' + N, 'gi'),
  ];
  for (const re of pats) {
    let m;
    while ((m = re.exec(s)) !== null) out.add(m[1].toLowerCase());
  }
  return Array.from(out).sort();
}

/** ★do $$ … $$ の 中を 読む★
 *  通すのは 次の形だけ:
 *    begin / end / if … then / end if / 空行
 *    if not exists ( select … )               … 在るか 見るだけ
 *    execute 'create policy … on exally.… '   … 作るだけ
 *  1つでも 外れたら 落とす（★読めない塊を 通さない★）。 */
export function checkDoBlocks(sql, prefix) {
  const box = {};
  stripNoise(sql, { collect: box });
  const 塊 = box.doBlocks || [];
  const reasons = [];
  for (let i = 0; i < 塊.length; i++) {
    const 中 = String(塊[i]);
    // execute の中身（文字列）を 取り出して 別に 見る
    const 走らせる = [];
    let 残り = 中.replace(/\bexecute\s+'((?:[^']|'')*)'/gi, (m, s1) => {
      走らせる.push(String(s1).replace(/''/g, "'"));
      return ' $exec$ ';
    });
    for (const 文 of 走らせる) {
      const re = new RegExp('^\\s*create\\s+policy\\s+[\\w]+\\s+on\\s+' + prefix.replace('.', '\\.') + '[a-z_][\\w]*\\b', 'i');
      if (!re.test(文)) {
        reasons.push('do $$ の中で 許していない物を 走らせようとしている: ' + 文.slice(0, 80));
      }
    }
    // 残り（走らせる物を除いた地の文）に 危ない書き方が 無いか
    残り = 残り.replace(/--[^\n]*/g, ' ').replace(/'(?:[^']|'')*'/g, "''");
    for (const d of DANGER) {
      const m = 残り.match(d.re);
      if (m) reasons.push('do $$ の中に 消す/書き換える書き方: ' + d.kind + ' → ' + m[0].trim());
    }
    // 地の文に 許していない命令が 無いか（select と 制御構文だけ）
    const 許す = /\b(begin|end|if|not|exists|select|from|where|and|or|then|is|null|true|false|declare)\b/gi;
    const 語 = 残り.replace(/\$exec\$/g, ' ').replace(/[^a-z_]+/gi, ' ').trim().split(/\s+/).filter(Boolean);
    for (const w of 語) {
      if (!許す.test(w)) {
        許す.lastIndex = 0;
        // pg_policy / polname / polrelid / regclass など 見るための名前は 通す
        if (/^(pg_[\w]*|pol[\w]*|regclass|exally|recipe|[\w]*_[\w]*)$/i.test(w)) continue;
        reasons.push('do $$ の中に 読めない語: ' + w);
      }
      許す.lastIndex = 0;
    }
  }
  return { 数: 塊.length, reasons };
}

/** ★`;` で 1文ずつ に 分ける★（コメント・文字列・ドル引用は もう 落ちている） */
export function splitStatements(sql) {
  return stripNoise(sql).split(';').map((x) => x.trim().replace(/\s+/g, ' ')).filter((x) => x.length > 0);
}

/* ★★白名簿＝ここに 在る 形だけ 通す★★（2026-10-02）
     `re` は ★文の 頭から★ 当てる。`mi` が 在れば 中身を もう一度 見る。
     ★形は repo の 本物（supabase/2026-08-27_recipe.sql）の 実績と 頭の「通す物」から★ */
const ALLOW = [
  { name: '読むだけ(select)', re: /^(with\s|select\s|table\s|show\s|explain\s)/i,
    mi: (b) => (/\binto\b/i.test(b) ? '読むだけの select に into が 入っている（棚が 作られる）' : null) },
  /* ★`(` を 必ず 要る★＝`create table x as select ...`（他の 部屋の 中身を 複製）を 許さない */
  { name: '棚を 足す', re: /^create\s+(?:unlogged\s+)?table\s+(?:if\s+not\s+exists\s+)?[\w.]+\s*\(/i },
  { name: '索引を 張る', re: /^create\s+(?:unique\s+)?index\s+(?:concurrently\s+)?(?:if\s+not\s+exists\s+)?[\w]+\s+on\s+/i },
  { name: '列/決まりを 足す・RLS を 入れる',
    re: /^alter\s+table\s+(?:if\s+exists\s+)?(?:only\s+)?[\w.]+\s+(?:add\s+column|add\s+constraint|enable\s+row\s+level\s+security|force\s+row\s+level\s+security|alter\s+column\s+[\w]+\s+set\s+default)\b/i,
    mi: (b) => (/\b(disable\s+row\s+level\s+security|no\s+force\s+row\s+level\s+security)\b/i.test(b) ? '行の守り(RLS)を 切ろうとしている' : null) },
  { name: '守り(policy)を 作る', re: /^create\s+policy\s+/i },
  { name: '覚書(comment)', re: /^comment\s+on\s+/i },
  /* ★`on` の 無い grant＝★役(ロール)を 配る★（`grant service_role to anon`）＝誰が 開けるかが 変わる★ */
  { name: '鍵を 配る(grant)', re: /^grant\s+/i,
    mi: (b) => (/\bon\b/i.test(b) ? null : '役(ロール)を 配ろうとしている（誰が 開けるかが 変わる＝先に 司さんに 訊く）') },
  { name: '鍵を 減らす(revoke)', re: /^revoke\s+/i },
  /* ★`do $$ ... $$` は 中身を checkDoBlocks が 読む★（ここは 形だけ） */
  { name: 'do の 塊', re: /^do\s+\$do\$(?:\s+language\s+plpgsql)?$/i },
];

/** ★門★ 通すか 止めるか */
export function guard(sql, opts) {
  const prefix = (opts && opts.prefix) || 'exally.';
  const reasons = [];

  /* ── ⓪ ★字を 読めたか★（閉じていない 物は 読めない＝通さない） ── */
  const box = {};
  stripNoise(sql, { collect: box });
  if (box.読めない) reasons.push('★字を 最後まで 読めない★（' + box.読めない + '）');

  /* ── ① ★白名簿＝1文ずつ★（ここが 本体） ── */
  for (const bun of splitStatements(sql)) {
    const a = ALLOW.find((x) => x.re.test(bun));
    if (!a) { reasons.push('★許した形に 無い文★: ' + (bun.length > 70 ? bun.slice(0, 70) + '...' : bun)); continue; }
    const ng = a.mi ? a.mi(bun) : null;
    if (ng) reasons.push(ng + ' → ' + bun.slice(0, 70));
  }

  /* ── ② 黒名簿（裏打ち） ── */
  for (const d of findDangerous(sql)) {
    reasons.push('消す/書き換える書き方が入っている: ' + d.kind + ' → ' + d.at);
  }

  const tables = findTargetTables(sql);
  const foreign = tables.filter((t) => t.indexOf(prefix) !== 0);
  for (const t of foreign) reasons.push('★他の部屋の棚に触ろうとしている★: ' + t + '（通すのは ' + prefix + ' だけ）');

  const dob = checkDoBlocks(sql, prefix);
  for (const r of dob.reasons) reasons.push(r);

  if (!tables.length && !/\bselect\b/i.test(stripNoise(sql))) {
    reasons.push('何をする物か読み取れない（棚も select も無い）');
  }

  return { ok: reasons.length === 0, reasons, tables, doBlocks: dob.数 };
}
