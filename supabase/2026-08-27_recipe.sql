-- ★覚えた手順（レシピ）の置き場★ 2026-08-27
-- 設計の正本 … lib/recipe.js の頭／js/recipe-store.js／docs/SPEC_recipe.md
-- ★入れるのは DB-test だけ。本番の倉庫は 声をかけてから★（まだ 1回も 当てていない）
--
-- ★決まり★
--   ・★入れるのは「手順」だけ★＝客の中身（見本の行・金額・氏名）は 1つも 入れない。
--     lib/recipe.js の レシピを作る が 見本を落としている＝★倉庫が漏れても 客の数字は 出ない★。
--   ・★倉庫（会社の物）に置く★＝端末を変えても 残る（競合は 客のブラウザの中だけ）。
--   ・★本人の物だけ 見える★＝RLS（画面ではなく 倉庫の側で 締める）。
--   ・★関数は 1つも作らない★（Postgres は 実行権を 既定で PUBLIC に渡すため）。

create table if not exists exally.recipe (
  id          bigint generated always as identity primary key,
  mochinushi  uuid        not null default auth.uid(),   -- 持ち主（ログインした人）
  itsu        timestamptz not null default now(),        -- ★サーバ時刻（端末の時計を信じない）★
  na          text        not null,                      -- 名（客が読む名前）
  tanomi      text        not null,                      -- 頼み（「前と同じ事か」を当てる鍵の片方）
  shimon      text        not null,                      -- 指紋＝列の名前と並び（鍵のもう片方）
  yoyaku      jsonb       not null default '{}'::jsonb,  -- 列の名前と並びだけ（★見本は入っていない★）
  tejun       jsonb       not null default '[]'::jsonb   -- 手順（決めた4種類しか入らない）
);

comment on table exally.recipe is
  'レシピ＝覚えた手順。2回目からは AIを呼ばずに 当てる。客の中身(見本の行)は入れない。本人の物だけ見える(RLS)。';

create index if not exists recipe_mochinushi_itsu_idx on exally.recipe (mochinushi, itsu desc);
create index if not exists recipe_shimon_idx          on exally.recipe (mochinushi, shimon);

-- ★権限（書いた≠効いた。当てた後に 必ず proacl と 行数を 実測する）★
alter table exally.recipe enable row level security;
revoke all on exally.recipe from public;
revoke all on exally.recipe from anon;
grant select, insert, delete on exally.recipe to authenticated;  -- ★覚えた物は 客が消せる★
grant select, insert, delete on exally.recipe to service_role;
revoke update, truncate, references, trigger on exally.recipe from authenticated;

-- ★本人の物だけ★（anon には ポリシーを1つも作らない＝1行も見えない）
-- ★2回 当てても 落ちない形にする（当てる道具は 何度でも走る）★
do $$
begin
  if not exists (select 1 from pg_policy where polname = 'recipe_jibun_no_mono_yomu'
                   and polrelid = 'exally.recipe'::regclass) then
    execute 'create policy recipe_jibun_no_mono_yomu on exally.recipe for select to authenticated using (mochinushi = auth.uid())';
  end if;
  if not exists (select 1 from pg_policy where polname = 'recipe_jibun_no_mono_kaku'
                   and polrelid = 'exally.recipe'::regclass) then
    execute 'create policy recipe_jibun_no_mono_kaku on exally.recipe for insert to authenticated with check (mochinushi = auth.uid())';
  end if;
  if not exists (select 1 from pg_policy where polname = 'recipe_jibun_no_mono_kesu'
                   and polrelid = 'exally.recipe'::regclass) then
    execute 'create policy recipe_jibun_no_mono_kesu on exally.recipe for delete to authenticated using (mochinushi = auth.uid())';
  end if;
end $$;
