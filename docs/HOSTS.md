# 入口の一覧（どのURLが今の入口で、どれが古いか）

作成 2026-08-04 ／ **★「古い」と書くだけにしない。実際に飛ぶことを機械で確認する★**
機械: `node scripts/check-hosts.mjs`（週1＋手動 / `.github/workflows/hosts.yml`）

## なぜ要るか

2026-08-04、**司さんが `payslip-app-test` を「テスト用」だと思って開いた**。
中身は 2026-08-02 で止まっていて、その日の直しが1つも入っていなかった。
**古い方で見ると「直っていない」ように見える。** 今回は正しい方で見ていたので事故にならなかっただけ。

**入口が複数あるのが原因。** どれが今の入口かを1枚にして、機械で見張る。

## 今の入口（★ここを使う★）

| 何 | URL | 倉庫(DB) | repo |
|---|---|---|---|
| ★本番 ハブ | https://exally.vercel.app/hub.html | 本番 | `exally` |
| ★本番 給与(Kyually) | https://exally.vercel.app/kyuyo/ | 本番 | `exally` |
| ★本番 グリッド | https://exally.vercel.app/book.html | 本番 | `exally` |
| ★テスト ハブ | https://exally-zeroact.github.io/exally-staging/hub.html | DB-test | `exally-staging` |
| ★テスト 給与(Kyually) | https://exally-zeroact.github.io/exally-staging/kyuyo/ | DB-test | `exally-staging` |

## 古い入口（★使わない。開いたら今の入口へ飛ぶ★）

| 古いURL | 飛び先 | 状態 | やり方 |
|---|---|---|---|
| https://exally-zeroact.github.io/payslip-app-test/ | テスト 給与 | ★2026-08-04 に塞いだ | 中身を消し、案内1枚＋自動転送（`?t=`/`#` を落とさない・404.html で直リンクも拾う） |
| https://payslip-app-olive.vercel.app/ | 本番 給与 | **未対応** | ★本番へ今回の直しが入ってから 308 で飛ばす（先に塞ぐと飛び先が古いままになる） |
| https://exally-zeroact.github.io/exally-test/ | — | 中身なし | Jekyll の既定ページ（READMEだけ）。アプリではないが、名前が紛らわしいので一覧に載せる |

## 決まり

1. **入口を増やしたら、この表に足す。** 足さない入口は「無い物」として扱う（誰も面倒を見ない）。
2. **古くしたら塞ぐ。** 「古い」と書くだけでは、人は開く（2026-08-04 に実際に起きた）。
3. 塞ぐ時は **うしろ（`?t=...` / `#...`）を落とさない**。Web明細のリンクが死ぬ。
4. 塞ぐ時は **Service Worker と PWA も片付ける**。端末に住み着く物は push だけでは消えない。
   - `sw.js` は消さず、**自分を登録解除してキャッシュを消すキルスイッチ**として残す
   - 案内ページ自身でも、飛ぶ前に SW 登録解除とキャッシュ削除をする
   - `manifest.json` は案内ページを指す（ホーム画面から開いた人も案内→新しい場所へ）
5. **機械で確認する。** `scripts/check-hosts.mjs` が、今の入口が 200 か・古い入口が今の入口へ飛ぶかを見る。

## 落ちた時に誰が見るか

`hosts` ワークフロー（週1・月曜9時JST＋手動）が落ちたら、その週のうちに実装セッションが気づき、
**次の報告に「入口 OK N／NG M」の1行を必ず載せる**。
通常CIに入れないのは `source-urls` と同じ理由 ―― 外の都合で赤くなると、
**自分のせいでない赤で push が止まり、赤が信用されなくなる**から。

## 他のアプリ（2026-08-04 に調べた）

| repo | GitHub Pages | 判定 |
|---|---|---|
| `payslip-app` | 404 | 公開なし（Vercel の `payslip-app-olive` が入口）→ 上の表のとおり未対応 |
| `nomiya-app-test` | 404 | Pages 公開なし（古い入口は残っていない） |
| `Daikou-app-test` | 404 | 同上 |
| `amazon-ads-automation-test` | 404 | 同上 |
| `exally-test` | 200 | Jekyll の既定ページのみ（アプリなし） |

**★飲み屋・代行請求・アマかせ・ダイコメに「生きたまま残った古い入口」は見つからなかった。**
ただし調べたのは **GitHub Pages と、既知の Vercel 1本だけ**。
各アプリの Vercel 側に旧ドメインが残っていないかは、そのアプリのセッションで確認が要る（未確認）。
