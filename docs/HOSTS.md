# 入口の一覧（どのURLが今の入口で、どれが古いか）

作成 2026-08-04 ／ **★「古い」と書くだけにしない。実際に飛ぶことを機械で確認する★**
機械: `node scripts/check-hosts.mjs`（週1・月曜9時JST＋手動 / `.github/workflows/hosts.yml`）
判定そのものの自己確認: `node scripts/check-hosts.mjs --self-test`

## なぜ要るか

2026-08-04、**司さんが古い入口を「テスト用」だと思って開いた**。中身は古いままだった。
さらに悪いことに、旧本番 `payslip-app-olive` は **中身が古いのに本番のデータ（本物の給料）を見ていた**。
誰かが開けば、古い計算で本物の数字を触れてしまう。

**入口が複数あるのが原因。** どれが今の入口かを1枚にして、機械で見張る。

## 今の入口（★ここを使う★）

| 何 | URL | 倉庫(DB) | repo |
|---|---|---|---|
| ★本番 ハブ | https://exally.vercel.app/hub.html | 本番 | `exally` |
| ★本番 給与(Kyually) | https://exally.vercel.app/kyuyo/ | 本番 | `exally` |
| ★本番 Web明細 | https://exally.vercel.app/kyuyo/meisai.html?t=… | 本番 | `exally` |
| ★本番 グリッド | https://exally.vercel.app/book.html | 本番 | `exally` |
| ★テスト ハブ | https://exally-zeroact.github.io/exally-staging/hub.html | DB-test | `exally-staging` |
| ★テスト 給与(Kyually) | https://exally-zeroact.github.io/exally-staging/kyuyo/ | DB-test | `exally-staging` |

## 古い入口（★使わない。開いたら今の入口へ飛ぶ★）

| 古いURL | 飛び先 | 状態 | やり方 |
|---|---|---|---|
| https://payslip-app-olive.vercel.app/ | 本番 給与 | ★2026-08-04 に塞いだ | `vercel.json` の `redirects` で **308**（`?t=`/`?c=`/`#` を落とさない・`/meisai` は `.html` を付け直す・`sw.js` だけ飛ばさない） |
| https://exally-zeroact.github.io/payslip-app-test/ | テスト 給与 | ★2026-08-04 に塞いだ | 中身を消し、案内1枚＋自動転送（`404.html` で直リンクも拾う） |
| https://exally-test.vercel.app/ （`/index.html` `/home.html` `/exally-home.html`） | 本番 ハブ | ★2026-08-04 に塞いだ | `vercel.json` の `redirects` で **308**。★この4本だけ★ |

### ★exally-test は「古いホームだけ」を塞ぐ（実務が同居している）★

`exally-test.vercel.app` には **代行請求システム（`/daikou-seikyu.html`）が実務で毎日動いています**。
グリッド・各テンプレート・`/api/claude` も生きています。**塞ぐのは古いホーム4本だけ。**
`check-hosts` に「★代行請求が飛ばされていないこと★」を毎週見る行を入れました
（塞ぐ範囲が広がって実務を巻き込んだら、その週のうちに赤で分かります）。
詳しくは `Exally-test` リポジトリの `docs/HOSTS.md`。

### 旧本番を塞ぐ時に気をつけたこと（同じ形で塞ぐ人のために）

1. **うしろを落とさない。** `?t=…`（Web明細のトークン）`?c=…`（初回コード）`#…` が消えると、
   従業員に配ったQR/リンクが死ぬ。Vercel の `redirects` は query をそのまま渡す。**実測で確かめる。**
2. **`/meisai` と `/meisai.html` の両方を書く。**
   旧ホストは `cleanUrls:true` で `/meisai.html` を `/meisai` に化けさせていた。
   新ホストに `/meisai` は無いので、`.html` を付け直さないと **転送は成功したのに 404 に着地する**
   （一番たちの悪い壊れ方＝「飛んだのに真っ白」）。だから `check-hosts` は **飛び先が 200 かまで見る**。
3. **`sw.js` だけは飛ばさない。**
   端末に住み着いた Service Worker は、サーバを塞いだだけでは消えない。
   `sw.js` まで飛ばすと SW の更新が失敗して **古いSWが永久に居座る**。
   ここは「自分を登録解除して全キャッシュを削除する」キルスイッチをそのまま返す。
4. **ログインは引き継がれない（origin が変わるため）。**
   ブラウザの保存領域は URL の家ごとに分かれているので、旧ホストで入れた印は新ホストへ持ち越せない。
   - 会社側: 新しい入口で **もう一度ログイン**（メール／パスワードは同じもので入れる）
   - 従業員(Web明細): **次の1回だけパスワードを入れ直す**（パスワードは端末ではなく倉庫にあるので同じ）
   これは仕様として避けられない。**「切れないようにする」ではなく「1回だけ入れ直す」が正しい説明。**

## 決まり

1. **入口を増やしたら、この表に足す。** 足さない入口は「無い物」として扱う（誰も面倒を見ない）。
2. **古くしたら塞ぐ。** 「古い」と書くだけでは、人は開く（2026-08-04 に実際に起きた）。
3. 塞ぐ時は **うしろ（`?t=…` / `#…`）を落とさない**。
4. 塞ぐ時は **Service Worker と PWA も片付ける**。端末に住み着く物は push だけでは消えない。
5. **機械で確認する。** `scripts/check-hosts.mjs`。

## 落ちた時に誰が見るか

`hosts` ワークフロー（週1・月曜9時JST＋手動）が落ちたら、**その週の実装セッションが見る**。
気づき方を人の記憶に頼らないため、**実装セッションは報告のたびに「入口 OK N／NG M」の1行を載せる**。
通常CIに入れないのは `source-urls` と同じ理由 ―― 外の都合で赤くなると、
**自分のせいでない赤で push が止まり、赤が信用されなくなる**から。

## 他のアプリ（2026-08-04 に Vercel の全プロジェクト15本を叩いて調べた）

| 入口 | 中身 | 判定 |
|---|---|---|
| `nomiya-app` / `nomiya-app-test` | Castally 売上管理 | **正常**。index は同じでも `js/supa-config.js` が別物（md5 が違う）＝倉庫は分かれている |
| `daikou-app` / `daikou-app-test` | ダイコメ | **正常**。1枚物だが `location.hostname` で倉庫を切り替えている |
| `daikome-jimusho` / `daikome-jimusho-test` | ダイコメ事務所 | **正常**。`js/dk-config.js` が別物（md5 が違う） |
| `payslip-app-*-exallysupoort-8848s-projects` | 旧給与 | **公開されていない**（Vercel のログインを求められる）＝客は入れない |
| `payslip-app.vercel.app` | 別物のAPI（`{"status":"API is running"}`） | 給与アプリではない。`/meisai.html?t=` は 404 |
| `shuri-app` / `boat-kelly-dash` / `amazon-ads-automation-test` | それぞれ現役 | 古い入口ではない |
| `project-6suu0` / `zeroact-memory-mcp` | 404 | 中身なし |

### ★見つけたが、このセッションでは触っていないもの（判断が要る）★

| 入口 | 何が出るか | なぜ触らなかったか |
|---|---|---|
| https://amazon-ads-automation.vercel.app/ | 「Adsilio — Amazon PPC Automation」（Next.js） | **★他人の物。うちとは無関係★**（2026-08-04 指示役が確認）。司さんのアマかせは **`amazon-ads-automation-lyart.vercel.app`**（title「アマかせ — Amazon運用アプリ」・Python）。`-lyart` 無しの住所は Vercel プロジェクトが持つ住所の一覧に入っていない＝**他人が先に取った住所**。**触らない。** |

## ★push は反映ではない（2026-08-04 に実際にやらかした）★

`exally-test` の `vercel.json` に説明用の `"//"` キーを足したところ、
**Vercel が設定を読む段階で弾き、ビルドに入る前にデプロイが失敗**した。
このとき **前のデプロイがそのまま配信され続ける**ので、画面は正常に見える。
＝ **「pushした」と「直しが入った」は別**。

気づけたのは `check-hosts` が「まだ 200 のまま（308 になっていない）」と赤で示したから。

**★同じ事故が 2026-08-02 にダイコメでも起きている**（`"_comment"` を足してデプロイ2回失敗）。
別のセッションが同じ日に同じ間違いを踏んだ ＝ 人の記憶では防げない。**だから機械で止める。**

### 決まり（全セッション共通・2026-08-04）

1. **`vercel.json` / `manifest.json` / `package.json` など規格が決まっている設定ファイルに、
   説明文やメモを書かない。** 書きたい事は `docs/` へ。
2. **知らない項目が混ざったら赤にするテストを各リポジトリに置く。**
   このリポジトリでは `tests/config-schema.test.mjs`（CI登録済み・`--self-test` つき）。
3. **塞ぐ・移す・切り替える作業は、「本当にそうなったか」を叩く見張りを先に作る。**
   `check-hosts` を先に作っていたから、今回の失敗に気づけた。
4. **push したら必ず実物のURLを叩いて確かめる。**
