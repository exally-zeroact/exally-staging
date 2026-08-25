# ★数え場（AIの回数を数える所）の設計 1枚★ — 2026-08-25

> ★2026-08-26：指示役のOKを受けて ★DB-test に当てました★（7章に実測）。
> ★本番の倉庫は まだ 1文字も触っていません★＝声をかけてから。

---

## 0. なぜ Supabase か（指示役の裁定 2026-08-25）

- Vercel KV は ★新しい契約物＝毎月のお金と 管理する物が1つ増える★
- Supabase は もう Pro で動いていて ★部屋(schema)を分ける形も決まっている★
- ★9 クレジット制も 同じ表の上に乗る★＝回数を数える所と クレジットを引く所を
  **別々に作ると 必ずズレる**（同じ状態を2か所で別々に判定するな）

★今の作りの弱点（正直に）★＝数え場は **その機械の中だけ**。Vercel は機械が増えるので
**すり抜けが在る**。1台に集中する連打（実際の事故の形）は 止まっている。

---

## 1. 置き場所

| | |
|---|---|
| 部屋(schema) | ★`exally`★（他アプリの部屋に作らない） |
| 表 | `exally.ai_tsukatta`（＝「AIを使った」1回ぶん の記録） |
| 窓口(view/関数) | ★`security_invoker = true`★ |
| 触れる人 | ★サーバの鍵(service_role)だけ★。★anon には 何も渡さない★ |

---

## 2. 表（1回 押すごとに 1行）

| 列 | 型 | 何を入れるか | なぜ |
|---|---|---|---|
| `id` | `bigint` 自動 | 通し番号 | 並べ替えの軸 |
| `hito` | `text` NOT NULL | ★`u:<Supabaseのsub>` か `ip:<入口>`★ | ★人ごとに数える★（名乗りが無くてもIPで数える） |
| `hito_kind` | `text` NOT NULL | `'user'` / `'ip'` | ★人とIPを 混ぜて数えない★ |
| `oshita` | `timestamptz` NOT NULL | 押した時刻（サーバ時刻） | ★端末の時計を信じない★ |
| `kekka` | `text` NOT NULL | `ok` / `tsukaisugi` / `zandaka` / `kagi` / `komiai` / `jikangire` / `ai_shippai` / `ookisugi` | ★止めた分も残す★（黙って止めない） |
| `nyuryoku_token` | `int` | 入力トークン | 9 の値段の素 |
| `shutsuryoku_token` | `int` | 出力トークン | 同上 |
| `oita_token` / `yominaoshita_token` | `int` | 置いた／読み直した | ★置き賃が違うので 分けて持つ★ |
| `watashita_mitsumori` | `int` | 渡した見積もりトークン | 2万トークンの上限を 後から検算できる |
| `kaiwa_kezutta` | `int` | 捨てた会話の件数 | ★黙って小さくしない★ |
| `iriguchi` | `text` | Origin | どの画面から来たか |
| `credit` | `int` DEFAULT 0 | ★使ったクレジット（9で使う・今は 0）★ | ★同じ表に乗せる★＝ズレない |

★人が書いた文そのものは 1文字も入れない★（客の中身。長さと数だけ）。

### 鍵と索引
- 主キー … `id`
- ★索引 `(hito, oshita DESC)`★ … 「この人の 直近1分／1日」を数えるのは これだけ
- ★索引 `(oshita)`★ … 寿命の掃除用

### 数え方（サーバがやる事）
```sql
-- ★通した分だけ数える★（止めた分まで数えると 窓がいつまでも空かない）
select count(*) from exally.ai_tsukatta
 where hito = $1 and kekka = 'ok' and oshita > now() - interval '1 minute';
```
★1分10回／1日100回★は ★api/claude.js の 事故止め 1か所★のまま（数字を SQL に書かない）。

---

## 3. 消し方・寿命

| | |
|---|---|
| 寿命 | ★90日★（作業用の記録＝原本ではない） |
| 掃除 | 毎日1回 `delete from exally.ai_tsukatta where oshita < now() - interval '90 days'` |
| ★消してよい理由★ | 回数の判定に要るのは ★直近1日★だけ。90日は「後から値段を検算する」ため |
| ★消さない物★ | ★成果物（客のファイル）は ここに1バイトも入らない★＝電帳法の話とは別 |
| 人が消えた時 | `hito` は ★Supabaseのsubの写し★＝退会しても ★この表は本人の中身を持たない★ |

---

## 4. 権限（★書いた≠効いた★）

- `revoke all on exally.ai_tsukatta from public, anon, authenticated;`
- `grant select, insert on exally.ai_tsukatta to service_role;`
- RLS … ★有効にする。ポリシーは service_role だけ★
- ★完成の条件（実測するまで 完成にしない）★
  1. ★`proacl` / `relacl` を実測して anon が 消えている事★
  2. ★`begin … rollback` で anon として押して 0行 である事（行数を数える）★
  3. ★authenticated でも 0行 である事★（画面から直接は触らせない）

---

## 5. ★関数（指示役 2026-08-25 の指摘）★

★Postgres は 関数の実行権を 既定で PUBLIC に渡す★＝表を締めても、後から関数を足した日に
★anon に残る★。

⇒ ★この部屋に 関数を1つも作らない★（2026-08-26 の決定）。
　★実測★＝`select count(*) from pg_proc where pronamespace='exally'::regnamespace` → ★0★
⇒ 作る日が来たら ★`revoke execute on function … from public, anon, authenticated` を書いてから
　proacl を実測する★（この行を消さずに 足す）。

---

## 6. ★掃除（90日）＝誰が いつ 回すか★

| | |
|---|---|
| 誰が | ★サーバ（AIの窓口）★ |
| いつ | ★その日の最初の書き込みの時に 1回だけ★ |
| 何を | `delete from exally.ai_tsukatta where oshita < now() - interval '90 days'` |
| なぜ この形か | ★見張りは 登録するまで1本も回らない★（7本中2本しか回っていなかった前例）。<br>外の見張り（GitHub Actions）は ★鍵(SUPABASE_ACCESS_TOKEN)が この repo に1本も無い★＝<br>作っても ★黙って何もしないまま緑★になる。pg_cron は ★本番にもDB-testにも 入っていない★（実測）。<br>⇒ ★書く物と 消す物を 同じ所（サーバ）で回す★＝書き始めた日から 必ず回る |
| 手で回す口 | `node scripts/apply-sql.mjs --souji`（★消せるのは この1文だけ★・場所と日数は焼き付け） |

★初回を 実際に回した（2026-08-26・DB-test）★
```
 100日前の行を1本 入れてから 回した
 掃除 … 前 4行 → ★消した 1行★ → 後 3行   （★90日より新しい3行は 残った★）
```

---

## 7. ★当てた後の実測（2026-08-26・DB-test khawdrnvssdenumbiwfg）★

★「書いた」ではなく「効いた」を 数で★

| 見た物 | 実測 |
|---|---|
| 表と列 | ★13列 在る★（id/hito/hito_kind/oshita/kekka/…/credit） |
| RLS | ★有効（relrowsecurity = true）★ |
| relacl | `postgres=arwdDxtm/postgres \| service_role=arwdDxtm/postgres` |
| ★anon の権限の数★ | ★0★（値に直して数えた・文字列で探していない） |
| ★authenticated の権限の数★ | ★0★ |
| ★public の権限の数★ | ★0★ |
| service_role | 8（★もともと部屋ごと持っている＝うちが足したのではない★。正直に書く） |
| ★この部屋の関数の数★ | ★0★ |

★中身が3行 在る状態で 押した★（★空だから0行 ではない★）

| 押し方 | 返ってきた物 |
|---|---|
| `begin; set local role anon; select count(*) …; rollback;` | ★42501 permission denied for table ai_tsukatta★（0行 ですらなく ★表に触れない★） |
| `begin; set local role authenticated; select …; rollback;` | ★42501 permission denied★ |
| `begin; set local role anon; insert …; rollback;` | ★42501 permission denied★（★書けない★） |
| ★本物の入口（PostgREST・anon鍵・`Accept-Profile: exally`）★ | ★401 / 42501 permission denied★ |

★最後の1行が いちばん大事★＝客が実際に叩ける口（HTTP）で 弾かれている事を 確かめた。

---

## 8. 入れる順番

1. ★DB-test に当てる★ … ★済（2026-08-26）★
2. 指示役の確認 … ★済★
3. ★本番は 声をかけてから★ … ★まだ★

★サーバが この表を使い始めるには あと1つ要ります★
- ★`SUPABASE_SERVICE_ROLE_KEY`（と `SUPABASE_URL`）を 配信の環境変数に置く事★
  （PostgREST の側は ★`exally` を 公開スキーマとして受け付ける事を 実測済★）
- ★鍵が入るまでは 今の「機械の中だけ」で数える★（1台への連打は 止まっている）。
  ★書き始めた日から 掃除も 一緒に回り始める★（6のとおり）。
