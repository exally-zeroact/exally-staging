-- ★数え場（AIの回数を数える所）★ 2026-08-26
-- 設計の正本 … docs/SPEC_kazoeba.md（指示役 2026-08-25 OK）
-- ★入れるのは DB-test だけ。本番の倉庫は 声をかけてから★
--
-- ★決まり★
--   ・部屋は exally（他アプリの部屋に作らない）
--   ・★anon には 何も渡さない★／触れるのは service_role だけ
--   ・★人が書いた文そのものは 1文字も入れない★（長さと数だけ）
--   ・★1分10回/1日100回の数字は api/claude.js の1か所★＝ここ(SQL)には 書かない
--   ・★関数は 1つも作らない★（Postgres は 関数の実行権を 既定で PUBLIC に渡すため。
--     作る日が来たら ★revoke execute してから proacl を実測する★）

create table if not exists exally.ai_tsukatta (
  id                   bigint generated always as identity primary key,
  hito                 text        not null,   -- 'u:<Supabaseのsub>' か 'ip:<入口>'
  hito_kind            text        not null,   -- 'user' | 'ip'（★混ぜて数えない★）
  oshita               timestamptz not null default now(),  -- ★サーバ時刻（端末の時計を信じない）★
  kekka                text        not null,   -- ok / tsukaisugi / zandaka / kagi / komiai / jikangire / ai_shippai / ookisugi
  nyuryoku_token       integer     not null default 0,
  shutsuryoku_token    integer     not null default 0,
  oita_token           integer     not null default 0,
  yominaoshita_token   integer     not null default 0,
  watashita_mitsumori  integer     not null default 0,
  kaiwa_kezutta        integer     not null default 0,
  iriguchi             text,
  credit               integer     not null default 0,   -- ★9 クレジット制で使う（今は 0）★
  constraint ai_tsukatta_hito_kind_ok check (hito_kind in ('user','ip'))
);

comment on table exally.ai_tsukatta is
  'AIの窓口を1回 押すごとに1行。回数の上限とクレジットを 同じ表で数える（同じ状態を2か所で判定しない）。人が書いた文は入れない。';

-- ★索引★ 1分/1日を数えるのは これだけ／掃除は oshita で引く
create index if not exists ai_tsukatta_hito_oshita_idx on exally.ai_tsukatta (hito, oshita desc);
create index if not exists ai_tsukatta_oshita_idx      on exally.ai_tsukatta (oshita);

-- ★権限（書いた≠効いた。当てた後に 必ず実測する）★
alter table exally.ai_tsukatta enable row level security;
revoke all on exally.ai_tsukatta from public;
revoke all on exally.ai_tsukatta from anon;
revoke all on exally.ai_tsukatta from authenticated;
grant select, insert, delete on exally.ai_tsukatta to service_role;
-- delete は ★90日の掃除★のため（消してよいのは 90日より古い行だけ）

-- 順番の入れ物（identity）も 同じだけ締める
revoke all on all sequences in schema exally from anon;
revoke all on all sequences in schema exally from authenticated;
