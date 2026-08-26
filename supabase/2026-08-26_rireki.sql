-- ★履歴（見る場所）の置き場★ 2026-08-26
-- 設計の正本 … lib/rireki.js の頭／docs/SPEC_kazoeba.md と同じ流儀
-- ★入れるのは DB-test だけ。本番の倉庫は 声をかけてから★
--
-- ★決まり★
--   ・★倉庫（会社の物）に置く★＝端末を変えても 残る
--     （競合は 客のブラウザの中(IndexedDB)だけ＝端末を変えると消える）
--   ・★客のブックには タブを1つも足さない★
--   ・★本人の物だけ 見える★＝RLS（画面ではなく 倉庫の側で 締める）
--   ・★関数は 1つも作らない★（Postgres は 実行権を 既定で PUBLIC に渡すため）

create table if not exists exally.rireki (
  id          bigint generated always as identity primary key,
  mochinushi  uuid        not null default auth.uid(),   -- 持ち主（ログインした人）
  itsu        timestamptz not null default now(),        -- ★サーバ時刻（端末の時計を信じない）★
  shurui      text        not null,                      -- 関数/自動化/直した所/診断/取り込み・書き出し/その他
  midashi     text        not null,
  nakami      jsonb       not null default '{}'::jsonb,  -- 何を聞いて／何をして／どのセルがどう変わったか
  file_name   text,
  credit      integer     not null default 0,            -- ★AIを呼んだ時だけ 1以上★
  constraint rireki_shurui_ok check (shurui in ('関数','自動化','直した所','診断','取り込み・書き出し','その他'))
);

comment on table exally.rireki is
  '履歴＝見る場所。本人の物だけ見える(RLS)。客のブックにはタブを足さない。人が書いた文そのものは入れない（何をしたか だけ）。';

create index if not exists rireki_mochinushi_itsu_idx on exally.rireki (mochinushi, itsu desc);
create index if not exists rireki_itsu_idx            on exally.rireki (itsu);

-- ★権限（書いた≠効いた。当てた後に 必ず実測する）★
alter table exally.rireki enable row level security;
revoke all on exally.rireki from public;
revoke all on exally.rireki from anon;
grant select, insert on exally.rireki to authenticated;   -- ★行はRLSで 本人の物だけ★
grant select, insert, delete on exally.rireki to service_role;

-- ★本人の物だけ★（anon には ポリシーを1つも作らない＝1行も見えない）
-- ★2回 当てても 落ちない形にする（当てる道具は 何度でも走る）★
--   Postgres に create policy if not exists は 無いので 自分で見てから作る。
do $$
begin
  if not exists (select 1 from pg_policy where polname = 'rireki_jibun_no_mono_yomu'
                   and polrelid = 'exally.rireki'::regclass) then
    execute 'create policy rireki_jibun_no_mono_yomu on exally.rireki for select to authenticated using (mochinushi = auth.uid())';
  end if;
  if not exists (select 1 from pg_policy where polname = 'rireki_jibun_no_mono_kaku'
                   and polrelid = 'exally.rireki'::regclass) then
    execute 'create policy rireki_jibun_no_mono_kaku on exally.rireki for insert to authenticated with check (mochinushi = auth.uid())';
  end if;
end $$;

-- ★2026-08-26 実測して 足した★
--   Supabase は 部屋ごとの既定で authenticated に arwdDxtm（8つ）を渡していた。
--   ★RLS で 行は守られている（Aさんは自分の1行しか見えない）★が、
--   ★要らない権限は 明示的に外す★＝「書いた≠効いた」を 両側から締める。
revoke update, delete, truncate, references, trigger on exally.rireki from authenticated;
