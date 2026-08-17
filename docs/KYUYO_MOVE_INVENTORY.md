# 給与(kyuyo/)を外す準備 — 数えた物の一覧

作成 2026-08-18 ／ **★これは「外す」ための下調べ。1バイトも外していない★**
号令: 指示役。**外す号令は Rakually の本番URLが立って 22人が移ってから。**

数えた道具: `scripts/dep-count.mjs`（★呼ばれる側を数える★。既に在る物を使った）
**★src= だけで判定しない★**（`require` / `import` / `window.○○` も数える）

---

## 0. ★先に言う。今 外すと何が壊れるか★

| # | 壊れる物 | 中身 | 気づけるか |
|---|---|---|---|
| ① | **`api/claude.js`（チャットのサーバ側）** | `require('../kyuyo/lib/shakaihoken-hyo.js')` / `koyo-hoken.js` / `shouhizei-ritsu.js` の**3本** | ★画面では気づけない★。押した時に関数が500 |
| ② | **`hub.html` のタイル** | `<a href="kyuyo/">` | 開けば分かる |
| ③ | **テスト線の `seikyu/`** | `kyuyo/lib/shiharai-chosho.js` / `kyuyo/lib/shouhizei-ritsu.js` の**2本**を読む | 請求書の画面が死ぬ |

★①は「呼ぶ側だけ写して本番を白画面にした」と同じ型★（sha一致・CI緑でも捕まらない）。
**法定の数値は写しを作らない決まりなので、`api/claude.js` は kyuyo の本体を読んでいる。**
⇒ 外す日は **①の3本の行き先を先に決める**（Rakually へ移すのか、Exally 側に残すのか）。

---

## 1. 本番 `exally` で数えた数（2026-08-18 実測）

| 数えた物 | 数 |
|---|---|
| 入口（HTML） | **6枚**（`book.html` `chat.html` `hub.html` `kyuyo/index.html` `kyuyo/admin.html` `kyuyo/meisai.html`） |
| **`seikyu/`** | **★本番repoには存在しない★**（テスト線だけ） |
| `kyuyo/` が **kyuyo/ の外**に必要とする物 | **8本**（下の表） |
| `kyuyo/` を **`src=` / `href=` で読んでいる**配信物 | **1本**（`hub.html`） |
| `kyuyo/` を **それ以外**（require / import / 文字列 / 設定）で書いている物 | **28本** |
| `tests/` のファイル | **75本**／うち **kyuyo を読む 15本** |

### `kyuyo/` が外に必要とする物（★持って出る時に一緒に要る★）

| ファイル | 誰が読む | 3分類 |
|---|---|---|
| `js/exally-login.js` | kyuyo/index | 両方 |
| `js/file-out.js` | kyuyo/index | 両方 |
| `js/supa-config.js` | kyuyo の3枚 全部 | 両方 |
| `lib/access.js` | kyuyo/index | 両方 |
| `lib/periods.js` | kyuyo/index | 両方 |
| `lib/op-registry.js` | kyuyo/index | **kyuyo だけ** |
| `lib/xlsx.full.min.js` | kyuyo/index | **kyuyo だけ** |
| `hub.html` | kyuyo/index（戻るリンク） | 本体 |

### 共通の部品（`lib/ js/ css/ ops/`）を3つに分ける

| 分類 | 数 | ファイル |
|---|---|---|
| **kyuyo だけ** | 2 | `lib/op-registry.js` / `lib/xlsx.full.min.js` |
| **Exally 本体だけ** | 10 | `css/hub.css` `js/auth.js` `js/hub.js` `js/ledger.js` `js/suite-data.js` `lib/aggregate.js` `lib/cross-agg.js` `lib/excel-version.js` `lib/grid-xlsx.js` `lib/ledger-source.js` |
| **両方** | 5 | `js/exally-login.js` `js/file-out.js` `js/supa-config.js` `lib/access.js` `lib/periods.js` |

★`js/exally-login.js` と `css/exally-ui.css` は **Exally 自身の物**。名前はこのままで正しい★
（Rakually は自分の repo の写しを改名する。Exally の repo には触らない）

---

## 2. テスト線 `exally-staging` で数えた数

| 数えた物 | 数 |
|---|---|
| 入口（HTML） | **8枚**（本番の6枚 ＋ `index.html` ＋ `seikyu/index.html`） |
| `seikyu/` が外に必要とする物 | **11本**（うち ★`kyuyo/lib/` が2本★） |
| `kyuyo/` を `src=` / `href=` で読んでいる物 | **4本**（`hub.html` `index.html` `seikyu/index.html` `tests/pages-hosting.test.mjs`） |
| `kyuyo/` を それ以外で書いている物 | **43本** |
| `tests/` のファイル | **74本**／うち **kyuyo を読む 17本** |
| 共通部品の3分類 | kyuyo だけ **1** ／ seikyu だけ **1** ／ 本体だけ **10** ／ 両方 **8** |

★テスト線では **seikyu が kyuyo を読んでいる**★。
**kyuyo を先に外すと請求書が死ぬ。** 外す順番は「seikyu の2本の行き先を決めてから」。

---

## 3. 転送の一覧（★外す日に これをそのまま入れる★）

手本＝`payslip-app-olive` を塞いだ時のやり方（`docs/HOSTS.md`）。
新URLは Rakually の本番URLが決まってから埋める（下の `<新URL>`）。

| 古いURL | 飛び先 | 気をつける事 |
|---|---|---|
| `https://exally.vercel.app/kyuyo/` | `<新URL>/` | 308 |
| `https://exally.vercel.app/kyuyo/index.html` | `<新URL>/` | 308 |
| `https://exally.vercel.app/kyuyo/meisai.html?t=…&c=…` | `<新URL>/meisai.html` | ★うしろ（`?t=` `?c=` `#`）を落とさない★＝従業員に配ったQR/リンクの形 |
| `https://exally.vercel.app/kyuyo/meisai`（`.html` 無し） | `<新URL>/meisai.html` | ★`.html` を付け直す★（付け直さないと 404 に着地する） |
| `https://exally.vercel.app/kyuyo/admin.html` | `<新URL>/admin.html` | 308 |
| `https://exally.vercel.app/sw.js` | **飛ばさない** | ★端末に住み着いた Service Worker はサーバを塞いでも消えない★。`kyuyo/admin.html` が `/sw.js` を登録している |

**★消さずに転送を残す★**（22人が今 使っている）。
`kyuyo/index.html` と `kyuyo/meisai.html` は開くたびに **SWを登録解除**する作りなので、
その2枚は SW が居座らない。**残るのは `admin.html` が登録した `/sw.js` だけ。**

登録先（★先に登録した★・実際に飛ばすのは号令の日）
- `docs/HOSTS.md`「これから塞ぐ入口（号令待ち）」
- `scripts/check-hosts.mjs` の `PLANNED`（**数だけ出す。まだ赤にしない**）
  ★まだ塞いでいない物を赤にすると、自分のせいでない赤が毎週鳴り、赤が信用されなくなる★

---

## 4. 外す日に1コミットで済ませるための順番（設計）

1. **`api/claude.js` の3本の行き先を決める**（★これが決まるまで外せない★）
2. `seikyu/` が読む `kyuyo/lib/` の2本の行き先を決める（テスト線）
3. `hub.html` のタイルの飛び先を新URLへ
4. `vercel.json` に上の転送を入れる（`kyuyo/` の中身は**消さない**）
5. `docs/HOSTS.md` と `check-hosts.mjs` の `PLANNED` を `OLD` へ移す
6. `tests/` の15本（テスト線は17本）の参照を直す
7. **★外す前と後で `dep-count.mjs` を同じ道具で回し、差を出す★**

---

## 5. 数える時に踏んだ罠（次の人へ）

- `dep-count.mjs` は **`require` の相対パスを二重に繋ぐ事がある**。
  `lib/lib/op-registry.js` / `kyuyo/ops/ops/payroll.monthly.js` を「無い物」と報告したが、
  **実際は両方とも在る**（★道具の誤報。ファイルは無事★）。
- `kyuyo/meisai.html` の `data:` から始まるアイコン/manifest も「無い物」に数える。
  **これも誤報**（ファイルではなく埋め込み）。
- ⇒ **★「無い物」が出たら、まず実在を自分で確かめる★**（道具の誤報を欠陥として報告しない）
