# ★この repo から 外した 見張り★（★黙って 消さない 為の 紙★）

> ★★なぜ この 紙が 在るか★★
> 見張りを 黙って 消すと ★「見張りが 在った」事ごと 消えます★。
> 後の 人は 「元から 無かった」と 読み、★同じ 穴を もう 一度 開けます★。
> ⇒★外す 時は ここに 書く★。★何を・なぜ・どこへ 移ったか★。

---

## 2026-09-16 … `.github/workflows/source-urls.yml`

### ★外した 物★
```
  `.github/workflows/source-urls.yml`（出典URLの死活・週1＋手動）
    ①`node kyuyo/scripts/check-source-urls.mjs`
       … 中央 statutory の 出典URLが 全部 生きて いるか（404 を 数える）
    ②`node kyuyo/scripts/check-wage88k-removal.mjs`
       … 社保 適用拡大の 賃金要件（月88,000円）の 撤廃が 法令に 入ったか
```

### ★★なぜ 外したか（★この repo では 1行も 見て いなかった★）★★
```
  ★落ちた 訳★
    Error: Cannot find module '…/kyuyo/scripts/check-source-urls.mjs'
  ★`kyuyo/` は この repo から 出て 行って います★
    `git ls-files 'kyuyo/**'` … ★0本★
  ⇒★★見張りの 顔を して 1行も 見て いない★★
  ⇒★2026-09-14 から ずっと 赤★＝★色が 動かないので 本物の 欠陥が 入っても 分からない★
    （記憶「材料が 無いのに 定時で 走り 緑を 返す物は 止めろ」の ★赤 版★。★赤でも 同じ 害★）
```

### ★★移った 先で 生きて います（★測ってから 外しました★）★★
```
  `rakually`（Rakunally）… https://github.com/exally-zeroact/rakually
    ★道具★ `kyuyo/scripts/check-source-urls.mjs`     … ★在る★
    ★道具★ `kyuyo/scripts/check-wage88k-removal.mjs` … ★在る★
    ★呼ぶ yml★ `.github/workflows/source-urls.yml` ／ `.github/workflows/ci.yml`
    ★走った 跡（gh run list）★
      2026-09-14  success
      2026-09-07  success
      2026-08-31  success
  ⇒★★向こうで 週1で 緑で 走って います＝守りは 効いて います★★
  ⇒★消えるのは この repo の 壊れた 写しだけ★
```

### ★★これは 私（Exally 側）の 持ち物では ありません★★
```
  `check-wage88k-removal.mjs` の 覚書きに ★期限 2026-09-15★ と 在ります
    ＝「賃金要件 88,000円 の 撤廃が 法令に 入ったか、★人が 見て 判断する★」期限
  ★2026-09-16 の 今日、その 期限は 過ぎて います★
  ★但し これは Rakunally（給与）の 持ち物★です。
  ⇒★私は 触りません★。★指示役1／経営者1 へ 上げました★。
  ★向こうの 週1は 09-14 に 緑★＝★法令には まだ 入って いません★（それが 緑の 意味）
```

### ★★同じ 型を 二度と 通さない 為に★★
```
  `tests/workflow-dougu-aru.test.mjs` を 足しました。
  ★見張りが 呼ぶ 道具が repo に 在るか★を ★総なめの 中で★ 数えます。
  ⇒★この 穴は 「週1で しか 走らない」ので 2日 気づきませんでした★
    ⇒★毎回 走る 総なめの 中で 見る★ 形に 移しました。
```
