/* book-open.js — ★受け取ったブックを開く／直して返す★（book.html から使う）
 *
 * ★設計の芯（2026-08-08〜09 の実測で決めた）★
 *   受け取ったファイルを直す → ★元のzipのバイト列を持ったまま、値だけ書き換えて閉じ直す★
 *   ゼロから表を作る        → 今までどおり gridToBook → writeBook（元々書式が無いので問題ない）
 *   ★どちらの道のファイルかを、ファイル単位で覚える★
 *   ★受け取ったファイルが writeBook に入ると判子が消える。分岐はここだけ。★
 *
 * ★読む＝見るだけ★
 *   画面に出すための読み取りは SheetJS で行う（.xlsx/.xlsm/.xlsb/.xls すべて読める）。
 *   ★保存はその結果を使わない★。保存は「元のバイト列」に対して行う。
 */
(function (root) {
  'use strict';

  /* 今 開いているファイル。null なら「ゼロから作った表」＝今までの保存の道 */
  var opened = null;

  function u8(ab) { return ab instanceof Uint8Array ? ab : new Uint8Array(ab); }

  /** ★拡張子を信じない。中身で見分ける★（客が拡張子を変えていても間違えない） */
  function detectKind(bytes, name) {
    var b = bytes;
    if (b.length > 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0) return 'xls';
    if (!(b[0] === 0x50 && b[1] === 0x4B)) return 'unknown';    // zip(PK) でない
    // zip の中身は開いてから見る。ここでは名前で仮置きして、開いた後に上書きする
    return /\.xlsb$/i.test(name || '') ? 'xlsb' : (/\.xlsm$/i.test(name || '') ? 'xlsm' : 'xlsx');
  }

  var MSG_XLS = 'この形式（.xls）は まだ直せません。今は開いて見るだけです。';
  /* ★VBAが入っていた時に 客へ出す言葉（★どうなるかを 先に言う★）★ */
  var MSG_VBA = 'このファイルには マクロ（VBA）が入っています。'
    + 'マクロは そのまま残しますが、ここでは 動きません。'
    + '毎月の繰り返しは、このあと「手順を覚えさせる」で 代わりに出来ます。';

  /** ファイルを開く。★元のバイト列をそのまま持つ★ */
  function openFile(file) {
    return file.arrayBuffer().then(function (ab) {
      var bytes = u8(ab);
      var kind = detectKind(bytes, file.name);
      if (kind === 'unknown') throw new Error('Excelのファイルとして読めませんでした');

      // 中身で確定させる（zip の時だけ）
      /* ★VBA入り(.xlsm)の扱い＝決めた（2026-08-25 司さん方針・指示役の指示）★
         ①★開ける★（読むだけ ではない。直して書き出せる）
         ②★VBAには 触らない。そのまま残す★（実測：書き出しても xl/vbaProject.bin は1バイトも変わらない）
         ③★VBAは 動かさない★（うちはブラウザ。Windows＋マクロ有効化が要る物は 勧めない）
         ④★VBAが要る仕事は うちの側（レシピ）で済ませる★＝画面でも そう言う */
      var hasVba = false;
      var _zip = null;
      if (kind !== 'xls' && root.ZipSurgeon) {
        try {
          var z = root.ZipSurgeon.read(bytes);
          _zip = z;
          if (z.has('xl/workbook.bin')) kind = 'xlsb';
          else if (z.has('xl/vbaProject.bin')) { kind = 'xlsm'; hasVba = true; }
          else if (z.has('xl/workbook.xml')) kind = 'xlsx';
        } catch (e) { /* 読めなければ名前のままにする */ }
      }

      // ★表示用の読み取り（見るだけ）★
      /* ★★`cellStyles: true` が 無いと 列の 幅が 1つも 来ません★★（2026-09-11 実測）
           SheetJS は これが 無いと `ws['!cols']` を ★作りません★
           ⇒ 幅を 読む 所は 前から 書いて 在ったのに ★材料が ずっと 空★でした
           ⇒★実Excel より 列が 広く、`####` に ならず、桁が 多く 出て いた★
             （実測 … Excel「1.001193」／うち「1.001193321」）
         ★重く ならないか 先に 測りました★（司さんの 実物 代行計算表2026.xlsb・15枚）
           今のまま … 309ms / 236ms ／ 山 21MB ／★幅が 来た板 0/15★
           足した後 … 276ms / 263ms ／ 山 26MB ／★幅が 来た板 15/15★
           ⇒★遅く なりません★（山は +5MB） */
      /* ══ ★★字体は ★借り物より 先に★ 読みます★★（2026-09-11）══════════════
         ★後から 読もうとしたら ファイルが 開かなく なりました★（120秒 待っても 開かない）
         ⇒ 切り分けた … ★袋から 1本 取り出す 所★で 返って 来ない
         ⇒★画面の 中で 直に 試すと 速い★（袋 1ms／styles 1ms／字体 3ms・3,432マス）
         ⇒★違いは 順番★＝借り物の `XLSX.read` の 後だと 元の バイト列が 使えなく なる
         ⇒★だから 先に 読みます★（板の 名前も `workbook.bin` から 自分で 取るので
           借り物を 待つ 必要が 在りません）
         ★時間切れ 5秒★＝★ファイルが 開く 事の 方が ずっと 大事★ */
      var 字体待ち = Promise.resolve(null);
      if (kind === 'xlsb' && _zip && root.XlsbEdit && root.XlsbJitai) {
        try {
          var 遅い = new Promise(function (ok) { setTimeout(function () { ok(null); }, 5000); });
          字体待ち = Promise.race([字体を読む(_zip), 遅い]).catch(function () { return null; });
        } catch (e) { 字体待ち = Promise.resolve(null); }
      }
      return 字体待ち.then(function (読んだ物) {
      /* ★★2つ 受け取ります★★（2026-09-25 から 「ゼロを 隠すか」も 一緒に 来ます） */
      var 字体表 = (読んだ物 && 読んだ物.字体) ? 読んだ物.字体 : null;
      var ゼロ表 = (読んだ物 && 読んだ物.ゼロ) ? 読んだ物.ゼロ : null;
      var 全部計算の印 = (読んだ物 && 読んだ物.全部計算 !== undefined) ? 読んだ物.全部計算 : null;
      var wb = root.XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, sheetStubs: false, cellStyles: true });

      /* ★表の名前での参照（Table[列名]）を、実際のA1範囲に直す★（2026-08-18）
         読み込みライブラリは .xlsb で ★表の名前も列名も捨てる★。
           実Excelの真値 =INDEX(R8.1[見本一四], MATCH(B4, R8.1[日付], 0))
           受け取る式     =INDEX(Table1[#Data],  MATCH(B4, Table1[#Data], 0))
         ＝INDEX と MATCH が同じ範囲を指す壊れた式になり、司さんの実物では
         ★式 15,126本のうち 11,669本が 1本残らず #ERROR★ になっていた。
         計算する側をいくら直しても、届く前に消えているので直らない。
         ★合わないセルは直さない（元のまま＝#ERROR のまま）★＝壊すより断る。 */
      var trFixes = {};
      var trStats = null;
      /* ★断った 式のうち 客に 説明できる 物★（2026-09-04）＝★直さない・言うだけ★ */
      var tr断り = [];
      var pre = Promise.resolve();
      if (kind !== 'xls' && root.TableRefs && root.ZipSurgeon) {
        pre = root.TableRefs.resolve(bytes, kind, wb, root.ZipSurgeon).then(function (r) {
          if (r && r.ok) { trFixes = r.fixes || {}; trStats = r.stats || null; tr断り = r.断り || []; }
          else if (r && root.console) root.console.warn('[Exally] 表の参照を直せませんでした: ' + r.why);
        }).catch(function (e) {
          if (root.console) root.console.warn('[Exally] 表の参照を直せませんでした', e);
        });
      }
      /* ★マクロ(VBA)は 開いた時に 読んでおく★（★読むだけ・動かさない★・AIは0回）
         ★読めなくても 画面は そのまま動く★＝読めない時は「未測定」と言う（0本と言わない） */
      /* ★★マスの 飾り（太字・字の色・塗り・罫線）を ★自前で★ 読む★★（2026-09-21）
           ★★なぜ 要るか★★
             実Excel が 作った 飾り付きの ファイルを お客さんの 道で 開いて 数えたら
             ★届いた 7 / 13★ でした（`golden-kazari-gamen-made-2026-09-21.tsv`）。
             ★太字・字の色・罫線2つ・塗り★ が 画面に 出て いませんでした。
           ★★なぜ 借り物では 駄目か★★（2026-09-21 実測）
             `cellStyles: true` を 付けても
               ・`c.s` は ★塗りだけ★
               ・`wb.Styles.Borders` は ★10個 とも `{}`★（★罫線が 空★）
               ・マスが どの `cellXf` を 指すかの ★番号を 捨てます★
           ★★板の 名前から 部品名を 出す 所は 在る 物を 使います★★
             `XlsxEdit.open()` が `xl/workbook.xml` と rels を 読んで
             `{ name, part }` を 作って います＝★新しく 書きません★
           ★★読めなくても 画面は そのまま 動きます★★
             ＝飾りが 付かないだけ（★前と 同じ★）。★開かなく なる 方が ずっと 悪い★
           ★★.xlsb は まだ です★★＝★包みの 中が 別物★（未対応と 書いて おきます） */
      var 飾り表 = null;
      var 図形表 = null;
      var 図形の並び = null;   /* ★名前で 引けなかった 時だけ 使います★ */
      var 図形の当て方 = '';    /* ★どちらで 読んだかを 残します★ */
      if ((kind === 'xlsx' || kind === 'xlsm') && root.XlsxKazari && root.XlsxEdit) {
        pre = pre.then(function () {
          return root.XlsxEdit.open(bytes).then(function (book) {
            /* ★★テーマの 色は `xl/theme/theme1.xml` に 在ります★★（2026-09-21）
                 ＝`.xlsx` の `styles.xml` は ★番号しか 書いて いません★（`<color theme="4"/>`）
                 ＝`.xlsb` は 番号と 実際の 色の 両方を 持ちます（形が 違う）
               ★無くても 動きます★＝テーマの 色が 付かないだけ（前と 同じ） */
            var テーマの字 = '';
            /* ★★外の `ゼロ表` に 入れます★★（2026-09-25 ここで 1回 踏みました）
                 はじめ ここで `var ゼロ表 = {}` と 書いたら
                 ⇒★この 関数の 中だけの 別物★ に なり、
                   `finish(..., ゼロ表)` に 渡るのは ★`.xlsb` の 方の 変数★ でした。
                 ⇒★落ちません／印が いつも false に なるだけ★＝★一番 見つけにくい 形★。
                 ⇒★絵の 点で 数える 試験が 捕まえました★（板の 印 2枚とも false） */
            ゼロ表 = ゼロ表 || {};   /* ★板の 名前 → ゼロを 隠すか★ */
            /* ★★`.xlsx` の 「開いたら 全部 計算しろ」★★（2026-09-25）
                 `xl/workbook.xml` の `<calcPr `fullCalcOnLoad="1"` />`
                 ★札が 無い＝立って いない★（Excel は 立って いる 時だけ 書きます） */
            var 先0 = book.zip.has('xl/workbook.xml')
              ? book.zip.text('xl/workbook.xml').then(function (wx) {
                  var m0 = /<calcPr[\s/>][^>]*>/.exec(wx);
                  if (m0) 全部計算の印 = /fullCalcOnLoad\s*=\s*"(1|true)"/.test(m0[0]);
                  else 全部計算の印 = false;
                }).catch(function () { /* ★読めない 時は 触らない★ */ })
              : Promise.resolve();
            var 先 = 先0.then(function () { return book.zip.has('xl/theme/theme1.xml')
              ? book.zip.text('xl/theme/theme1.xml').then(function (x) { テーマの字 = x; })
                  .catch(function () { テーマの字 = ''; })
              : Promise.resolve(); });
            return 先.then(function () {
            return book.zip.text('xl/styles.xml').then(function (型の字) {
              var 表 = {}, 鎖 = Promise.resolve();   /* ゼロ表 は 上で 作って います */
              book.sheets.forEach(function (板) {
                鎖 = 鎖.then(function () {
                  return book.zip.text(板.part).then(function (xml) {
                    /* ★★ゼロの 印を 先に 読みます★★（2026-09-25 ここで 1回 踏みました）
                         はじめ 飾りの 後ろに 置いたら ★1度も 動きませんでした★。
                         ⇒`飾りを読む` が 断ると 下の `.catch` が 飲み込み、
                           ★その後ろの 行に 来ません★。
                         ⇒★落ちません／印が いつも false に なるだけ★＝★一番 見つけにくい 形★。
                         ⇒★軽くて 断らない 方を 先に 置きます★ */
                    /* ★★ゼロを 隠す 板か★★（2026-09-25）
                         `.xlsx` は 字で 書いて あります ... `<sheetView `showZeros="0"` ほか>`
                         ★Excel は 「出す」時は この 札を 書きません★
                         ⇒★札が 在って 0/false の 時だけ 隠す★
                         ⇒★札が 無い＝出す★（★分からない では ない★） */
                    /* ★★`<sheetViews>` と `<sheetView>` は 別物★★（2026-09-25 ここで 1回 踏みました）
                         はじめ `xml.indexOf('<sheetView')` で 探したら
                         ⇒★先に 在る `<sheetViews>`（複数形の 箱）に 当たりました★
                         ⇒札を 探す 所が ``<sheetViews>`` だけに なり ★いつも 「出す」★。
                         ⇒★落ちません／印が いつも false に なるだけ★＝★一番 見つけにくい 形★。
                       ⇒★名の 後ろが 空白か `/` か `>` の 物だけを 拾います★ */
                    var m0 = /<sheetView[\s/>][^>]*>/.exec(xml);
                    if (m0) {
                      var m = /showZeros\s*=\s*"([^"]*)"/.exec(m0[0]);
                      ゼロ表[板.name] = !!(m && (m[1] === '0' || m[1] === 'false'));
                    }
                    表[板.name] = root.XlsxKazari.飾りを読む(xml, 型の字, テーマの字);
                  }).catch(function () { /* ★1枚 読めなくても 他は 出す★ */ });
                });
              });
              return 鎖.then(function () { 飾り表 = 表; });
            }).then(function () {
              /* ★★図形（判子）も 読みます★★（2026-09-21）
                   ★借り物は 図形を くれません★＝`xl/drawings/` を 読む 所が 1つも 無かった
                   ★板ごとの 図形は rels で 解きます★＝★並び順で 当てません★ */
              if (!root.XlsxZukei) return null;
              var 表2 = {}, 鎖2 = Promise.resolve();
              book.sheets.forEach(function (板) {
                鎖2 = 鎖2.then(function () {
                  var rels = 板.part.split('/');
                  var 名 = rels.pop();
                  var relsの道 = rels.join('/') + '/_rels/' + 名 + '.rels';
                  if (!book.zip.has(relsの道)) return null;
                  return book.zip.text(板.part).then(function (xml) {
                    return book.zip.text(relsの道).then(function (r) {
                      var 部 = root.XlsxZukei.図形の部品名(xml, r, 板.part);
                      if (!部 || !book.zip.has(部)) return null;
                      return book.zip.text(部).then(function (d) {
                        var 図 = root.XlsxZukei.図形を読む(d, root.XlsxKazari
                          ? root.XlsxKazari.テーマを読む(テーマの字) : null);
                        if (図.length) 表2[板.name] = 図;
                      });
                    });
                  }).catch(function () { /* ★1枚 読めなくても 他は 出す★ */ });
                });
              });
              return 鎖2.then(function () { 図形表 = 表2; });
            });
            });
          }).catch(function (e) {
            飾り表 = null;
            if (root.console) root.console.warn('[Exally] 飾りを 読めませんでした', e);
          });
        });
      }
      /* ══ ★★`.xlsb` の 図形★★ ══（2026-09-21）
           ★★なぜ 別の 道が 要るか★★
             `.xlsb` の 包みの 中は ほとんど `.bin` ですが、
             ★`xl/drawings/drawing1.xml` は XML の まま 残ります★（経営者1 の 実測・09-21）
             ⇒★図形は `.xlsb` でも 同じ 道具で 読めます★
           ★★但し 2つ 違います★★（2026-09-21 実測）
             ①板（`sheet1.bin`）が 2進＝`<drawing r:id=>` を ★字として 探せません★
               ⇒`lib/xlsx-zukei.js` が ★rels の `Type` で 解きます★
             ②`XlsxEdit.open()` は `xl/workbook.xml` を 読むので ★使えません★
               ⇒★包みを 直に 見ます★（`ZipSurgeon`）
           ★★板の 名前 → 部品名は 「並び」で 当てて います★★
             ＝★この repo が 前から そうして います★（`saveXlsb`・同じ ファイルの 下の 方）
               「workbook.bin を 読まずに、開いた時の 並びで 対応させる」
             ＝★★板が 2枚 以上 在る `.xlsb` で これが 合うかは 未測定★★
             ⇒★合わないと 判子が 別の 板に 出ます★（★消えるのでは なく ずれる★）
           ★★マスの 飾りは `.xlsb` では まだです★★＝`xl/styles.bin`（2進）
           ★読めなくても 画面は そのまま 動きます★ */
      if (kind === 'xlsb' && root.XlsxZukei && root.ZipSurgeon) {
        pre = pre.then(function () {
          var z2 = root.ZipSurgeon.read(bytes);
          /* ★★板の 名前 ⇒ 部品名は 「引いて」 出します★★（2026-09-21）
               `workbook.bin` の 並び ⇒ `rId` ⇒ `xl/_rels/workbook.bin.rels` ⇒ 部品名
               ＝`lib/xlsb-edit.js` の `板たち()`（★BIFF12 を 歩く 道具は 前から 在りました★）
             ★★なぜ 並びで 当てないか★★
               経営者1 が 板 2枚の `.xlsb` を 2本 作りました（並びを 入れ替えた 物も）。
               ⇒★Excel が 作り直した 物では 番号が 並びに 付いて 来て いました★
               ⇒★★つまり その 2本では 「番号当て」でも 通って しまいます★★
               ⇒★だから あの 材料では 番号当ての 良し悪しは 割れません★
               ⇒それでも やめます＝★「たまたま 合って いる」と 「引いて いる」は 別★
               ⇒★他の 道具（LibreOffice・古い Excel・板を 消した 後）は 未測定★
             ★引けない 時は 並びに 戻します★＝★開かなく なる 方が ずっと 悪い★ */
          var 名で引く = null;
          var しまう = function (出) {
            if (!出) return;
            if (Object.keys(出.名表).length) 図形表 = 出.名表;
            if (Object.keys(出.並表).length) 図形の並び = 出.並表;
            図形の当て方 = 出.当て方;
          };
          try {
            if (root.XlsbEdit && typeof root.XlsbEdit.板たち === 'function'
                && z2.has('xl/workbook.bin') && z2.has('xl/_rels/workbook.bin.rels')) {
              return z2.bytes('xl/workbook.bin').then(function (wbb) {
                return z2.text('xl/_rels/workbook.bin.rels').then(function (rx) {
                  名で引く = root.XlsbEdit.板たち(wbb, rx);
                  return 図形を集める(z2, 名で引く).then(しまう);
                });
              });
            }
          } catch (e) { /* ★引けない 時は 下へ★ */ }
          return 図形を集める(z2, null).then(しまう);
        }).catch(function (e) {
          図形表 = 図形表 || null;
          if (root.console) root.console.warn('[Exally] .xlsb の 図形を 読めませんでした', e);
        });
      }
      var マクロ = null;
      if (hasVba && root.Vba && root.ZipSurgeon) {
        pre = pre.then(function () {
          return root.ZipSurgeon.read(bytes).bytes('xl/vbaProject.bin').then(function (bin) {
            var 読み = root.Vba.読む(bin, root.XLSX && root.XLSX.CFB);
            var 見立て = (読み.ok && root.VbaMikata) ? root.VbaMikata.見立てる(読み.モジュール) : null;
            マクロ = { 読み: 読み, 見立て: 見立て };
          }).catch(function (e) {
            マクロ = { 読み: { ok: false, モジュール: [], なぜ: '読めませんでした' }, 見立て: null };
            if (root.console) root.console.warn('[Exally] マクロを読めませんでした', e);
          });
        });
      }
      return pre.then(function () { return finish(bytes, kind, wb, file, trFixes, trStats, hasVba, マクロ, tr断り, 字体表, 飾り表, 図形表, 図形の並び, 図形の当て方, ゼロ表, 全部計算の印); });
      });
    });
  }

  /** ★★`.xlsb` の 図形を 集める★★（2026-09-21）
   *    `板たち` が 引けた ⇒ ★板の 名前★ で 持ちます（`図形表`）
   *    引けなかった ⇒ ★並びの 番号★ で 持ちます（`図形の並び`・前の やり方）
   *    ★どちらで 読んだかを 画面の 言づてに 出します★＝★黙って 落ちない★ */
  function 図形を集める(z2, 名で引く) {
    var 板たち;
    if (名で引く && 名で引く.length) {
      板たち = 名で引く.map(function (x) { return { 鍵: x.名, 道: x.部品, 名で: true }; });
    } else {
      板たち = z2.names().filter(function (n) {
        return n.indexOf('xl/worksheets/sheet') === 0 && n.indexOf('.bin') === n.length - 4;
      }).sort(function (a, b) {
        var f = function (x) { return parseInt(String(x).replace(/[^0-9]/g, ''), 10) || 0; };
        return f(a) - f(b);
      }).map(function (道, i) { return { 鍵: i, 道: 道, 名で: false }; });
    }
    var 名表 = {}, 並表 = {}, 鎖 = Promise.resolve();
    板たち.forEach(function (板) {
      鎖 = 鎖.then(function () {
        var 名 = 板.道.split('/').pop();
        var relsの道 = 'xl/worksheets/_rels/' + 名 + '.rels';
        if (!z2.has(relsの道)) return null;
        return z2.text(relsの道).then(function (r) {
          /* ★板は 2進なので 字を 渡しません★＝`Type` で 解かせます */
          var 部 = root.XlsxZukei.図形の部品名('', r, 板.道);
          if (!部 || !z2.has(部)) return null;
          return z2.text(部).then(function (d) {
            var 図 = root.XlsxZukei.図形を読む(d);
            if (!図.length) return;
            if (板.名で) 名表[板.鍵] = 図; else 並表[板.鍵] = 図;
          });
        }).catch(function () { /* ★1枚 読めなくても 他は 出す★ */ });
      });
    });
    /* ★★外の 入れ物を 触りません★★＝★返して 呼んだ 側に 入れさせます★
         ＝この 道具は `openFile` の 外に 在るので 中の 変数が 見えません
         ＝★見えないのに 書くと 黙って 窓（global）を 作ります★ */
    return 鎖.then(function () {
      return { 名表: 名表, 並表: 並表, 当て方: 名で引く ? '名前' : '並び' };
    });
  }

  /** 読み終わった物をグリッドの形にして、控え(base)を作る */
  /* ══ ★★開いた 直後に 計算が 要るか★★ ══（2026-09-25）
       ★★なぜ★★
         司さんの 実物は 開くのに ★23秒★。★読み込みは 1〜2%★ で
         ★残りは 「開いた 直後の 計算」★（経営者1 の CPU の 記録）。
         ★実Excel は 保存済みの 答えを 持って いて、開いた だけでは 計算しません★。
         ⇒★だから 「計算する 理由が 1つも 無い 本」は 計算しません★
         ⇒★★これは 「Excel と 同じに する」事です★★（司さんの 決め イ）
       ★★理由が 在る 場合（経営者1 が 挙げて 数えた）★★
         ㋐★開いた 日で 答えが 変わる 関数★（TODAY NOW RAND OFFSET INDIRECT CELL INFO ･･･）
         ㋑★式は 在るのに 答えが 保存されて いない マス★
         ㋓★「開いたら 全部 計算しろ」の 印★
         ⇒★司さんの 実物は ㋐0個 ㋑0個 ㋓立って いない★（★2人が 別々の 道具で 読んだ★）
       ★★1つでも 当たれば 計算します★★＝★迷ったら 計算する 側に 倒す★
         ＝★印が 読めなかった 時（`null`）も 計算します★
         ＝★速さより 答えが 大事★

       ══ ★★数えたが わざと 名簿に 入れて いない 関数が 4つ 在ります★★ ══（2026-09-25）
         ★`WEBSERVICE` ／ `STOCKHISTORY` ／ `TRANSLATE` ／ `DETECTLANGUAGE`★
         （`lib/formula-soto.js:148` が 持って いる ★外へ 出る 関数★・経営者1 が 数えました）
         ★実Excel では この 4つも 「開いた 日で 答えが 変わる」側です★
         ⇒だから ㋐に 入れるのが 「Excel と 同じ」に 見えます。
         ★★それでも 入れて いません＝これは 速さの 話では なく 守りの 話です★★
           `lib/formula-soto.js` の 頭に こう 書いて あります:
             「もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★」
             「★画面から 直に 外へ 出さない／行き先は うちが 決める★」
           ⇒★㋐に 足すと 「知らない 人の 本を 開いた だけで 外へ 出る」に なります★
           ⇒★今の 「計算しない」は それを 止めて います★（★ついでに 止まっただけ★）
         ★★だから 決めるのは 司さんです★★（★「誰が 開けるか」が 変わる 事は 先に 訊く★）
           選べる 形 3つ（★どれも まだ 決まって いません★）
             ⑴㋐に 4つ 足す ......... Excel と 同じ 振る舞い／★但し 外へ 出る★
             ⑵足さない（★今これ★） ... 守りは 固い／★答えは 古いまま★
             ⑶4つが 在る 本だけ 訊く ... お客さんに 決めさせる（経営者1 の 推し）
         ★★但し 他の 理由（㋐㋑㋓）で 計算する 本では この 4つも 走ります★★
           ＝★私が 塞いだ のでは ありません＝前から そうでした★ */
  var 日で変わる関数 = /\b(TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/i;
  /* ══ ★★外へ 出る 関数（4つ）を 数えます★★ ══（2026-09-26）
       ★`lib/formula-soto.js:148` が 持って いる 4つ★（★そこが 正本＝ここで 書き足しません★）
       ★なぜ 数えるか★
         `lib/formula-soto.js` の 頭に こう 書いて あります:
           「もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★」
         ⇒★知らない 人の 本を 開いた だけで 外へ 出る★ 事に なります
         ⇒司さん 09-25「★おすすめで 直せ★」／経営者1 の 推し ⑶
         ⇒★★その 4つが 在る 本だけ お客さんに 訊く★★
       ⇒★ここは 数えるだけ＝出すか どうかは お客さんが 決めます★
       ★見張り★ tests/soto-wo-kiku-webkit.mjs */
  var 外へ出る関数 = /\b(WEBSERVICE|STOCKHISTORY|TRANSLATE|DETECTLANGUAGE)\s*\(/i;
  function 開いた直後に計算が要るか(板たち, 全部計算の印) {
    var 訳 = { 日で変わる: 0, 答えが無い: 0, 外へ出る: 0, 全部計算の印: 全部計算の印, 式: 0 };
    for (var i = 0; i < 板たち.length; i++) {
      var d = (板たち[i] && 板たち[i].data) || {};
      for (var k in d) {
        var c = d[k];
        if (!c || typeof c.f !== 'string' || c.f.charAt(0) !== '=') continue;
        訳.式++;
        if (日で変わる関数.test(c.f)) 訳.日で変わる++;
        if (外へ出る関数.test(c.f)) 訳.外へ出る++;
      }
      /* ★★「答えが 無い」は 板が 数えた 物を 使います★★（2026-09-25 ここで 1回 踏みました）
           はじめ `c.d === ''` で 数えたら ★2,255個★ 出ました。
           ＝`=IFERROR(･･･,"")` の 答えは ★空の 字＝立派な 答え★ です。
           ＝経営者1 が 実Excel で 数えた ㋑は ★0個★。
           ⇒★`sheetToGrid` が 「`c.v` が 無い（undefined/null）」時だけ 数えて います★
           ⇒★これを 直すまで 「計算しない」は 1度も 効いて いませんでした★ */
      訳.答えが無い += (板たち[i] && 板たち[i].答えが無い数) || 0;
    }
    訳.要る = !!(訳.日で変わる || 訳.答えが無い || 全部計算の印 !== false);
    return 訳;
  }

  function finish(bytes, kind, wb, file, trFixes, trStats, hasVba, マクロ, tr断り, 字体表, 飾り表, 図形表, 図形の並び, 図形の当て方, ゼロ表, 全部計算の印) {
      /* ★その本の 既定の 字体を 覚える★＝列の 幅を 点に 直すのに 要る
         （SheetJS は `wb.Styles.Fonts[0]` に 入れる … 実測 2026-09-11
           {"sz":11,"name":"游ゴシック",...}） */
      既定の字体 = (wb.Styles && wb.Styles.Fonts && wb.Styles.Fonts[0]) ? wb.Styles.Fonts[0] : null;
      var out = wb.SheetNames.map(function (nm, i) {
        /* ★`.xlsx` は 名前で／`.xlsb` は 並びで★（上の 注を 見て ください） */
        var 図 = 図形表 ? 図形表[nm] : null;
        if (!図 && 図形の並び) 図 = 図形の並び[i];
        return sheetToGrid(wb.Sheets[nm], nm, trFixes, 字体表 ? 字体表[nm] : null,
          飾り表 ? 飾り表[nm] : null, 図 || null, (ゼロ表 && ゼロ表[nm]) || false);
      });
      /* ★控えは「見せている文字」ではなく「元の生の値」から作る★（2026-08-09）
         画面用に 46043 を "1/21(水)" にして見せているので、その文字を控えにすると
         ★計算し直した瞬間に 46043 と食い違い、全部「変わった」ことになる★
         （実物14シートで 14,424セルが変わった扱いになり、保存が断られた）。 */
      var base = {};
      wb.SheetNames.forEach(function (nm) {
        var ws = wb.Sheets[nm];
        Object.keys(ws).forEach(function (a) {
          if (a.charAt(0) === '!') return;
          var rc = root.XLSX.utils.decode_cell(a);
          var v = ws[a].v;
          base[nm + '|' + rc.r + ',' + rc.c] = (v === undefined || v === null) ? '' : v;
        });
      });

      opened = {
        name: file.name, kind: kind, bytes: bytes,
        sheetNames: wb.SheetNames.slice(),
        /* ★開いた時の値を覚えておく★
           保存の時は「変わったセルだけ」書く。全部書こうとすると
           ・触っていない所まで書き換える危険
           ・答えがエラーの数式セル（記録11）に当たって★保存そのものが断られる★
           が起きる（実物14シート・2万セルで実際に起きた 2026-08-09）。 */
        base: base,
        tableRefs: trStats,        // ★何本 直したか（見張りと報告が読む。画面には出さない）
        /* ★★`.xlsb` の 板を どちらで 当てたか★★（2026-09-21）
             '名前' ＝ `workbook.bin` の 並び ⇒ rId ⇒ rels ⇒ 部品名 を ★引いた★
             '並び' ＝ 引けなかったので ★番号で 当てた★（前の やり方）
             ''     ＝ `.xlsb` では ない／図形を 読んで いない
           ★★なぜ 残すか★★
             経営者1「★あなたの 側で どちらを 読んで いるかを 字で 見せて ください★」
             ＝★板の 並びを 入れ替えた 材料でも Excel が 番号を 付け直す ので★
               ★★答えが 合って いても どちらで 読んだかは 分かりません★★
             ⇒★だから 道具の 側に 書き残します★ */
        図形の当て方: 図形の当て方 || '',
      };
      var 計算の訳 = 開いた直後に計算が要るか(out, 全部計算の印);
      return { kind: kind, sheets: out, opened: opened, hasVba: !!hasVba, マクロ: マクロ || null, 計算の訳: 計算の訳,
        表の断り: tr断り || [] };
  }

  /** ★袋から 1本 取り出す★（縮んで いれば ほどく） */
  function 袋から(z, 名) {
    var e = null, list = z.entries || [];
    for (var i = 0; i < list.length; i++) if (list[i].name === 名) { e = list[i]; break; }
    if (!e) return Promise.resolve(null);
    if (e.method === 0) return Promise.resolve(e.raw);
    return Promise.resolve(root.ZipSurgeon.inflateRaw(e.raw)).then(function (b) {
      return b ? new Uint8Array(b) : null;
    });
  }

  /** ★.xlsb の「板ごと・マスごとの 字体」と「ゼロを 隠すか」を 読む★
   *  返り ･･･ { 字体: { 板の名前: { 'r,c': {pt,名} } }, ゼロ: { 板の名前: true/false } }
   *  ★読めなければ null★＝断って 元のまま（勝手に 直さない） */
  function 字体を読む(z) {
    var XE = root.XlsbEdit, JT = root.XlsbJitai;
    return Promise.all([
      袋から(z, 'xl/styles.bin'),
      袋から(z, 'xl/workbook.bin'),
      袋から(z, 'xl/_rels/workbook.bin.rels'),
    ]).then(function (三つ) {
      if (!三つ[0] || !三つ[1] || !三つ[2]) return null;
      var sp = XE.parse(三つ[0]), wp = XE.parse(三つ[1]);
      if (!sp.ok || !wp.ok) return null;
      var 字 = '';
      for (var i = 0; i < 三つ[2].length; i++) 字 += String.fromCharCode(三つ[2][i]);
      /* ★★「開いたら 全部 計算しろ」の 印★★（2026-09-25）
           ＝`workbook.bin` の 記録157 の 26バイト目 ビット0
           ＝★実Excel で 作った 7冊の 差で 決めました★（`lib/xlsb-jitai.js` の 断り）
           ⇒★立って いなければ 開いた 直後に 計算する 理由は ありません★ */
      var 全部計算 = JT.開いたら全部計算するか(wp.recs);
      var 対応 = JT.板とファイル(wp.recs, 字);
      if (!対応) return null;
      var 名ら = Object.keys(対応), 仕事 = [], 出 = {}, ゼロ = {};
      for (var k = 0; k < 名ら.length; k++) {
        (function (名) {
          仕事.push(袋から(z, 対応[名]).then(function (sb) {
            if (!sb) return;
            var p2 = XE.parse(sb);
            if (!p2.ok) return;
            var 表 = JT.板の字体(sp.recs, p2.recs);
            if (表) 出[名] = 表;
            /* ★★同じ 記録を 歩く ついでに 「ゼロを 隠すか」も 読みます★★（2026-09-25）
                 ★2度 歩きません★＝ここは もう 板の 記録を 開いて います */
            var ゼ = JT.ゼロを隠すか(p2.recs);
            if (ゼ !== null) ゼロ[名] = ゼ;
          }, function () { /* 1枚 読めなくても 他は 使う */ }));
        })(名ら[k]);
      }
      return Promise.all(仕事).then(function () {
        if (!Object.keys(出).length && !Object.keys(ゼロ).length && 全部計算 === null) return null;
        return { 字体: 出, ゼロ: ゼロ, 全部計算: 全部計算 };
      });
    }, function () { return null; });
  }

  /* ★日本語の曜日（aaa / aaaa）を先に本物の文字へ置き換える★
     SheetJS の書式エンジンは この2つを知らないので、そのまま「(aaa)」と出る。
     実物の代行計算表は日付の書式が `m/d(aaa)` なので、★画面に (aaa) が並ぶ★（2026-08-09 実機で確認）。
     中身(シリアル値)は触らない。見せ方だけ直す。 */
  var WD = ['日', '月', '火', '水', '木', '金', '土'];
  function withWeekday(fmt, serial) {
    if (!/a{3,4}/.test(String(fmt))) return fmt;
    var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
    var w = WD[d.getUTCDay()];
    return String(fmt).replace(/a{4}/g, '"' + w + '曜日"').replace(/a{3}/g, '"' + w + '"');
  }

  /** SheetJS の1シート → グリッドの形 { name, data:{'r,c':{v,f,d,numFmt}}, colW, ... }
   *  tableFixes … 'シート名|r,c' → 表の参照を A1 範囲に直した式（TableRefs が作る）。
   *               ★無い物は元のまま★＝直せなかったセルは触らない。 */
  /* ★★うちが 溢れで 作った マスと 同じ 印★★
     ★book.html の `溢れの印` と 同じ 字で なければ なりません★
     ＝違うと 画面が「これは 溢れの 先だ」と 気付かず、古い 数が 残ります。
     ★見張り★ … tests/afureru.test.mjs が 両方が 同じ 字か 突き合わせます */
  var 溢れの印 = '_溢れ元';

  /** ★その マスが「広がった 先」か★＝★元では ない★ */
  function 溢れの先か(c, rc) {
    if (c.f !== undefined && c.f !== null && c.f !== '') return false;   // ★式を 持つ＝元★
    var 元 = 溢れの元の場所(c.F);
    if (!元) return false;
    return 元 !== (rc.r + ',' + rc.c);
  }

  /** ★"G1:G5" → "0,6"★（左上＝溢れの 元）／読めない 時は null */
  function 溢れの元の場所(F) {
    var m = /^([A-Z]+)(\d+):/.exec(String(F || ''));
    if (!m) return null;
    var c = 0, s = m[1];
    for (var i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
    return (Number(m[2]) - 1) + ',' + (c - 1);
  }

  /* ══ ★★国で 中身が 変わる 書式番号を 実Excel に 合わせる★★（2026-09-11）══════
     ★★見つけ方★★ 司さんの 実物（代行計算表2026.xlsb）の ★出る字★を
       実Excel と 1マスずつ 突き合わせた（5,437マス）⇒★323マス 違った★
         実Excel「2,220」／うち「2,220 」（★後ろに 空白★）
     ★正体★＝★書式番号 38番★（この ファイルで ★785マス★ 使って いる）
       ★5〜8番と 37〜40番は 国で 中身が 変わります★（Microsoft の 決め）
       借り物（SheetJS）が 持って いるのは ★世界共通（英語）の 表★
         38番 … `#,##0 ;[Red](#,##0)`  ⇒ 後ろに 空白／マイナスは ★括弧★
       ★日本語の Excel の 38番★（実Excel に 打たせて 測った 2026-09-11）
         `#,##0;[赤]-#,##0`            ⇒ 空白なし／マイナスは ★-★
         1234.5→「1,235」／-1234.5→「★-1,235★」／0→「0」／2220→「2,220」
     ★直し方★＝★借り物の 中は 読まず、外から 表を 入れ替える★（`SSF.load` は 公の 口）
       色の 名は `[Red]` の まま＝借り物は `[赤]` を 知りません。
       ★出る 字は 同じ★（色の 名は 字に 出ない）＝実Excel と 合う 事を 測って 確かめた。
     ★まだ 測って いない★＝5・6・7・8・37・39・40番。
       ★この ファイルでは 1マスも 使って いません★（数えた）⇒ 当て推量で 直しません。 */
  /* ★★`SSF.load` では 届きませんでした★★（2026-09-11 実測）
       表そのものは 書き換わります（前『#,##0 ;[Red](#,##0)』→ 後『#,##0;[Red]-#,##0』）が、
       ★.xlsb を 読む 所は 自前の 写しを 持って いて★ 読んだ マスの `z` は 古い まま。
     ⇒★読んだ 後に 直します★（借り物の 中は 1文字も 触って いません） */
  var 世界共通の字 = {
    '#,##0 ;[Red](#,##0)': '#,##0;[Red]-#,##0'   /* 38番 ★実Excel（日本語）に 打たせて 測った★ */
  };
  /** ★国で 中身が 変わる 書式を 日本の Excel の 物に 直す★
   *  ★断り★＝お客さんが わざと 同じ 字の 書式を 作って いたら、それも 直します。
   *    （日本の Excel で その 形を 作ると ★38番として 入る★ので、実害は 無いと 見て います） */
  function 書式を実Excelに合わせる(z) {
    var s = String(z == null ? '' : z);
    return (世界共通の字[s] !== undefined) ? 世界共通の字[s] : z;
  }

  /* ══ ★★列の 幅を 実Excel と 同じ 点(px)に 直す★★（2026-09-11）══════════════
     ★実Excel に 聞いて 決めました★（docs/measured/toru-hashira-haba.ps1）
       打った 字数  2     3     5     8.44   10    12     15     20     30
       実Excel の点 20.5  28.5  44.5  72     84.5  100.5  124.5  164.5  244.5
       ⇒★一直線★＝`点 = 字数 × 8 + 4.5`（游ゴシック 11pt）
     ★ファイルの 中の 数（width）で 突き合わせた★（8列とも ★差 0.00★）
       width 2.5625 → ×8 = 20.5 ／ width 30.5625 → ×8 = 244.5
       ⇒★点 = ファイルの width × 一字の幅★
     ★SheetJS の `wpx` は 使えません★＝1字を 14点と 思い込んで います
       実測 … 5字の 列 … 実Excel ★44.5点★ ／ SheetJS ★78点★
       ⇒ そのまま 使うと ★幅が 足りて しまい `####` に ならない★
     ★一字の幅は 字体で 変わります★＝★その本の 既定の 字体で「0」を 測って 切り捨て★
       游ゴシック 11pt … 8.153 → ★8★（実Excel と 一致）
       Calibri   11pt … 7.430 → ★7★（実Excel と 一致）
       ⇒★2つの 字体で 実Excel と 合う 事を 確かめました★（当て推量では ありません）
     ★まだ 出来て いない 事★
       ・★幅を 書いて いない 列（標準の 幅）★は ファイルの 既定を 読めません
         （SheetJS が `defaultColWidth` を 出さない）
         ⇒ 実Excel の 標準（8.43字）で 置き換えます＝★その本が 標準を 変えて いたら 合いません★ */

  /** ★実Excel の 標準の 列幅（字）★＝どの 版でも 8.43（実測 2026-09-11 … 8.44 と 出る） */
  var 標準の字数 = 8.43;
  /** ★字体が 読めない 時の 逃げ先★＝游ゴシック 11pt の 実測値 */
  var 逃げの字幅 = 8;

  /** ★その本の 既定の 字体★（SheetJS は `wb.Styles.Fonts[0]` に 入れる） */
  var 既定の字体 = null;

  /** ★「0」1文字の 点(px)★＝★画面と 同じ 測り方★（canvas で 測る）
   *  ★切り捨て★＝実Excel も 整数の 点で 持つ（8.153 → 8 ／ 7.430 → 7 で 一致した） */
  /** ★★実Excel が 持つ「0」1文字の 点★★（2026-10-04 夜・経営者の 実測 ㋑ と 司さんの 本 516列）
   *  ★★なぜ 表か★★ お客さんの 機械で 測ると ★その 機械の 字体★で 決まる。
   *    Linux の WebKit には 游ゴシックが 無く 別の 字に 逃げて「0」が 約10点 ⇒ 標準の 列が 74→93点、
   *    ★#### が 出ず 桁が 多く 出た★（kakareta-ji 幅8.44 の 7組）。Mac・iPhone も 同じ 恐れ（未測定）。
   *    ⇒★表に 在る 字体は 機械で 測らない★（表に 無い 字体だけ 今まで 通り 測る）
   *  ★出どころ★ 游ゴシック 11 ＝8 ･･･ 実Excel 16.0.20430 で 列幅 8.43→72px・1→12.5・10→84.5（㋑）、
   *      司さんの 本 15板 516列が 全部「字×8＋5」と ±1px（道具135）
   *    Calibri 11 ＝7 ･･･ 2026-09-11 の 実測（7.430→7 で 実Excel と 一致）
   *  ★字体の 名前は ファイルに 書かれた まま★で 引く。★測って いない 名前（英語名 Yu Gothic 等）は 置かない★
   *    （表に 無い 字体は 下の 測り方へ 落ちる＝前と 同じ） */
  var 実Excelの一字の幅 = {
    '游ゴシック|11': 8,
    'Calibri|11': 7
  };

  function 一字の幅(字体) {
    var 表の名 = ((字体 && 字体.name) ? 字体.name : '游ゴシック') + '|' + ((字体 && 字体.sz) ? 字体.sz : 11);
    if (Object.prototype.hasOwnProperty.call(実Excelの一字の幅, 表の名)) return 実Excelの一字の幅[表の名];
    try {
      var d = root.document;
      if (!d || !d.createElement) return 逃げの字幅;
      var x = d.createElement('canvas').getContext('2d');
      if (!x) return 逃げの字幅;
      var 名 = (字体 && 字体.name) ? 字体.name : '游ゴシック';
      var 大 = (字体 && 字体.sz) ? 字体.sz : 11;
      x.font = 大 + 'pt "' + 名 + '",sans-serif';
      var w = Math.floor(x.measureText('0').width);
      return (w > 0 && w < 100) ? w : 逃げの字幅;
    } catch (e) { return 逃げの字幅; }
  }

  /** ★1列ぶんの 幅を 点に★／読めない 時は 0（＝画面の 既定に 任せる） */
  function 幅を点に(col, 字幅) {
    if (!col) return 0;
    /* ★ファイルの 中の 数を 使う★＝これが 実Excel と 差 0.00 で 合う */
    if (typeof col.width === 'number' && col.width > 0) return Math.round(col.width * 字幅);
    /* ★width が 無い 時だけ 字数から 作る★（`点 = 字数 × 一字の幅 + 余白5`） */
    if (typeof col.wch === 'number' && col.wch > 0) return Math.round(col.wch * 字幅 + 5);
    return 0;
  }

  /** ★飾りを マスに 写す★（★在る 物だけ★＝空の 鍵を 増やさない） */
  function 飾りを写す(cell, 飾) {
    if (!cell || !飾) return;
    if (飾.bold) cell.bold = true;
    if (飾.italic) cell.italic = true;
    if (飾.underline) cell.underline = true;
    if (飾.color) cell.color = 飾.color;
    if (飾.bgColor) cell.bgColor = 飾.bgColor;
    if (飾.border) cell.border = 飾.border;
  }

  /* ★借り物が 渡す 誤りの 答えを 字に★（`w` が 無い 時は 番号から・Excel の 誤りの 番号） */
  var 誤りの番号 = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?',
    36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };
  function 誤りの字(c) {
    if (typeof c.w === 'string' && c.w.charAt(0) === '#') return c.w;
    return 誤りの番号[c.v] || '#N/A';
  }
  function sheetToGrid(ws, name, tableFixes, 字体, 飾り, 図形, ゼロを隠す) {

    var data = {}, X = root.XLSX, fixes = tableFixes || {};
    var 答えが無い数 = 0;   /* ★式は 在るのに 答えが 保存されて いない マス★ */
    Object.keys(ws).forEach(function (a) {
      if (a.charAt(0) === '!') return;
      var c = ws[a], rc = X.utils.decode_cell(a);
      var cell = { v: '', f: '', d: '' };
      if (c.f !== undefined && c.f !== null && c.f !== '') {
        var fixed = fixes[name + '|' + rc.r + ',' + rc.c];
        cell.f = fixed !== undefined ? fixed : ('=' + c.f);
        cell.d = c.v !== undefined && c.v !== null ? c.v : '';   // ★ファイルの答え（キャッシュ）をそのまま出す★
        /* ══ ★★答えが 誤りで 保存された 式は 誤りの 字で 持ちます★★ ══（2026-10-02）
             ★借り物は ★`t:'e'`・`v`＝誤りの 番号（#REF! なら 23）・`w`＝'#REF!'★ で 渡します★
             ★前は `v` を そのまま 持ち ⇒ ★数の 23★ ⇒ 下で ★書式を 掛けて 数の 字★ に して いました★
             ★実物で 出ました★ ... 司さんの 本の ★記録11（答えが 誤り）69個 とも #REF!（0x17）★
               ＝★実Excel は 開いた だけで #REF! を 出す／うちは 数★（経営者 実測 69個・10-02）
               ＝★計算の 話では ありません★（実Excel は 計算して いない・保存された 答えが #REF!）
             ⇒★字は 計算し直した 時と 同じ 形（'#REF!' の 字）★＝1打ち目の 後と 揃います
             ★見張り★ tests/hozon-no-ayamari-wo-yomu-webkit.mjs */
        if (c.t === 'e') cell.d = 誤りの字(c);
        /* ★★「答えが 無い」と 「答えが 空の 字」は 別物★★（2026-09-25 ここで 1回 踏みました）
             `=IFERROR(･･･,"")` の 答えは ★空の 字★＝★立派な 答え★ です。
             `cell.d === ''` で 数えたら ★2,255個★ 出て、
             ★『答えが 保存されて いない』の 数が 嘘に なりました★
             ⇒経営者1 の 実測は ★0個★（実Excel で 数えた）
             ⇒★`c.v` が 無い（undefined / null）時だけ 数えます★ */
        if (c.v === undefined || c.v === null) 答えが無い数++;
      } else if (c.F && 溢れの先か(c, rc)) {
        /* ══ ★★Excel が 書き込んだ「広がった 先の 答え」★★（2026-09-11）══════
           ★実物で 測った 事★（docs/measured/toru-excel-kara.ps1 で 作った ファイルを
             SheetJS で 読んだ … 2026-09-11）
               G1 … {v:1, f:"_xlws.SORT(E1:E5)", ★F:"G1:G5"★}   ← 溢れの 元
               G2 … {v:1,                        ★F:"G1:G5"★}   ← ★広がった 先★（式は 無い）
               G3 … {v:3,                        ★F:"G1:G5"★}
           ⇒★Excel は 広がった 先の 答えも ファイルに 書きます★
           ⇒ 前は それを ★ただの 数★として 読んで いた
             ⇒ うちの エンジンから 見ると ★溢れ先が 塞がって いる★
             ⇒★#SPILL! ＝ Excel の SORT / UNIQUE / FILTER が 全部 開けない★
                （2026-09-11 実測 … G1・H1・I1 の 3つとも #SPILL!）
           ⇒★うちが 溢れで 作った マスと 同じ 印を 付けます★（book.html の `溢れの印`）
             ＝`v` を 持たせない ⇒ `_pushGrid` は null を 送る ⇒ 塞がない
             ⇒ うちの エンジンが 溢れ直し、`_溢れを写す` が 中身を 入れ直す
           ★断り★＝うちが その 関数を 計算できない 時は、
             ファイルに 在った 答えは ★消えます★（元の 式が エラーに なる）。
             前は「全部 #SPILL!」だったので ★どちらでも 出ません★が、
             ★出なく なる 物が 在る事は 書いて おきます★ */
        cell.d = c.v !== undefined && c.v !== null ? c.v : '';
        if (c.t === 'e') cell.d = 誤りの字(c);   /* ★上の 式の 道と 同じ★（道が 2本） */
        cell[溢れの印] = 溢れの元の場所(c.F);
      } else {
        cell.v = c.v !== undefined && c.v !== null ? c.v : '';
        cell.d = c.w !== undefined ? c.w : cell.v;               // w = Excelが表示していた文字
      }
      /* ══ ★★字の マスは 「答えは 字」の 印を 付ける★★ ══（2026-10-04 夕）
           ★実Excel★ ... 字の マス（'1e-7 ／ ="1e-7"）は ★書式に 関わらず 字の まま「1e-7」★（経営者 真値 組 J 10/10）
           ★前は★ 印が 式の 答え（`d字`＝計算し直した 時 だけ）にしか 付かず、
             画面の forDisplay が 指数の 字を 数に して ★「0.0000001」★（`0.00` なら ★「0.00」★）に 化けて いた
           ★見る 物は 借り物の 型 `t === 's'`★（★typeof v では 見ない★＝打った 数も v は 字の まま 持つ／
             t:'e' 誤りの 値・t:'d' 日付と 取り違えない＝経営者の 叩き⑵）
           ★空の 字（v === ''）には 付けない★（どちらでも 画面は 空）
           ★値の マス・答えが 字の 式・溢れ先 の 3本 とも★（道が 3本＝ここ 1か所で 付ける）
           ★司さんの 本★ ... 式で ない 字の マス 1,571 のうち 数の 形 0（10-04 数えた）＝画面の 字は 動かない 見込み
           ★見張り★ tests/wareki-keika-gamen-webkit.mjs（組 J） */
      if (c.t === 's' && c.v !== undefined && c.v !== null && String(c.v) !== '') cell.d字 = true;
      if (c.z) cell.numFmt = 書式を実Excelに合わせる(c.z);   /* ★38番は 国で 中身が 変わる★ */
      /* ★日付が 46043 という裸の数字で出るのを止める★
         SheetJS が表示用の文字(w)を作らない事がある（.xlsb で実際に起きた 2026-08-09）。
         その時は表示形式(z)を使って自分で作る。★中身(v)はシリアル値のまま持つ★
         ＝保存する時に日付として書き戻せる。 */
      /* ★★2026-09-15＝自前の 台（`lib/shoshiki.js`）に 寄せました★★
         ★★ここは 前から 壊れて いました★★
           `book.html:3466`（★ふだん 通る 道★）は ★書式を 掛ける 前に 15桁に 丸めます★
           （2026-09-11 の 直し＝★実物で お金が 1円 ずれて いた★）。
           ★この 行には その 直しが 在りませんでした★
           ＝★同じ 事を する 道が 2本／片方だけ 直って いた★
           ⇒★15桁の 丸めは 台の 中★に 入れた＝★呼ぶ側が 忘れられません★
         ★測ってから 寄せました★（★壊してから 数えない★）
           司さんの 実物 ★16,536マス（書式つきの 数）★で
           ★台の 字と 今の 字は 1マスも 違いません（16,536/16,536）★
           ★この 行を 通る マスは 0★（ふだんは `c.w`＝ファイルの 字を 出す）
           ⇒★この 1冊の 画面は 1マスも 動きません★
         ★`withWeekday` は 台には 要りません★＝★台が `aaa` を そのまま 読みます★
           （あれは ★SSF が 曜日を 出せない★ から 付けて いた 継ぎ当て） */
      if ((cell.d === '' || typeof cell.d === 'number') && c.z && typeof c.v === 'number') {
        try {
          var t;
          var 書の台 = root.Shoshiki;
          var 台の出 = 書の台 ? 書の台.当てる(c.v, c.z) : null;
          if (台の出 && 台の出.出せる) {
            t = 台の出.字;
          } else {
            /* ★台が まだ 読めない 書式だけ 借り物に 回します★
               ＝`ggge`（和暦）／`[h]`（24時間超）／`?/?`（分数）／`E+`（指数）
               ★司さんの 実物は 1つも 使って いません★（実測）＝★0マス★
               ⇒★同じ 入力を 2か所で 解いて いる のでは なく
                 「★まだ 作って いない 分だけ 借りて いる★」＝★作れば 減ります★ */
            t = root.XLSX.SSF.format(withWeekday(c.z, c.v), c.v);
          }
          /* ══ ★★書式を 掛ける 前の 数を 残します★★ ══（2026-09-25）
               ★★なぜ★★
                 ここで `cell.d` は ★字★ に なり ★元の 数は どこにも 残りません★。
                 画面の `_ゼロを隠すか` は 式の マスで ★その 字を 数に 読み直して★ いました。
                 ⇒書式に 字が 混じると ★`Number('0 円')` ＝ NaN★ ⇒ ★0 と 判じられません★
                 ⇒★実Excel は 空・うちは 「0 円」★＝経営者1 の 実測で 47個★
                    `#,##0" 円"` 22個 ／ `0.00"時間"` 19個 ／ `"¥"#,##0_);…` 6個
                 ⇒★★字を 読み直すのを やめて 元の 数を 渡します★★
               ★数の 時だけ 残します★（字の 答えに 付けると 「字の 0」を 隠す 事に なる）
               ★見張り★ tests/zero-wo-kakusu-webkit.mjs（★書式に 字が 混じった 0★） */
          if (typeof c.v === 'number') cell.書式前の数 = c.v;
          if (t !== undefined && t !== null && t !== '') cell.d = t;
        } catch (e) { /* 作れなければ数のまま出す */ }
      }
      /* ★飾りを 付ける★（`lib/xlsx-kazari.js` が 作った 物）
           ★台が 前から 持って いる 名前に 合わせて います★
           ＝`bold` / `italic` / `underline` / `color` / `bgColor` / `border`
           ＝★画面の 描き手を 1行も 変えずに 出ます★ */
      var 飾 = 飾り ? 飾り[a] : null;
      if (飾) 飾りを写す(cell, 飾);
      data[rc.r + ',' + rc.c] = cell;
    });
    /* ══ ★★中身が 空でも 飾りだけ 在る マス★★ ══（2026-09-21）
         ★実Excel の `BorderAround` は 空の マスにも 罫線を 付けます★
         ⇒`<c r="B3" s="4"/>`（値も 式も 無い）
         ⇒借り物は `sheetStubs: false` で ★この マスを 丸ごと 捨てます★
         ⇒★上の 回では 1つも 拾えません★＝★四角の 下の 辺が 消えます★
         ⇒★だから ここで 入れ直します★ */
    if (飾り) {
      Object.keys(飾り).forEach(function (a) {
        var rc2;
        try { rc2 = X.utils.decode_cell(a); } catch (e) { return; }
        var k2 = rc2.r + ',' + rc2.c;
        if (data[k2]) return;
        var 空 = { v: '', f: '', d: '' };
        飾りを写す(空, 飾り[a]);
        data[k2] = 空;
      });
    }
    /* ══ ★★結合した マスを 読む★★（2026-09-11）══════════════════════════
       ★前は 1組も 読んで いませんでした★（`!merges` を 一度も 見て いない）
       ⇒ 実Excel で ★G1:H1★ と 2マス分 に 広げて ある 所を 1マス分と して 扱い、
         字が 入らず ★`######`★ に なって いました
         実測 … 司さんの 実物「給料表」G1「640,098 円」/ 幅59点（本当は 2マス分）
       ★画面の 持ち方★（book.html と 同じ）
         元の マス … `mergeEnd = {r, c}`（右下の 場所）
         中の マス … `merged = {r, c}`（元の 場所）
       ★SheetJS は `!merges` で くれます★（実測 … 給料表で 24組） */
    var 結合の数 = 0;
    (ws['!merges'] || []).forEach(function (m) {
      if (!m || !m.s || !m.e) return;
      if (m.s.r === m.e.r && m.s.c === m.e.c) return;      /* ★1マスだけは 結合では ない★ */
      var 元 = data[m.s.r + ',' + m.s.c];
      if (!元) 元 = data[m.s.r + ',' + m.s.c] = { v: '', f: '', d: '' };
      元.mergeEnd = { r: m.e.r, c: m.e.c };
      for (var r = m.s.r; r <= m.e.r; r++) {
        for (var c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) continue;
          var 子 = data[r + ',' + c];
          if (!子) 子 = data[r + ',' + c] = { v: '', f: '', d: '' };
          子.merged = { r: m.s.r, c: m.s.c };
        }
      }
      結合の数++;
    });

    var colW = {};
    var 字幅 = 一字の幅(既定の字体);
    (ws['!cols'] || []).forEach(function (col, i) {
      var px = 幅を点に(col, 字幅);
      if (px) colW[i] = px;
    });
    /* ★表の枠（!ref）も 覚えておく★（2026-08-27 指示役の指摘）
       ＝★「値か式が在る所」と「表の枠」は 違う★。
         実物 計算シート … ★値か式 400行×72列／表の枠 404行×152列★
       ★地図が 72列と言うと AIは 73列目から先を 一生 掘らない★ので、
       ★両方を 名前つきで 出す★ため ここで拾う（捨てない）。 */
    var 枠 = null;
    try {
      if (ws['!ref']) {
        var rg = X.utils.decode_range(ws['!ref']);
        枠 = { 行数: rg.e.r + 1, 列数: rg.e.c + 1 };
      }
    } catch (e) { /* 読めない時は null＝「未測定」（0にしない） */ }
    /* ★幅を 書いて いない 列（標準の 幅）も 実Excel に 合わせる★
       ＝画面の 既定は 80点／実Excel の 標準は ★72点★（8.43字 × 8 + 5）
       ⇒ これが 無いと ★Excel 8桁／うち 11桁★の ままです（実測）
       ⇒★この 板だけの 既定★＝新しく 作る ブックの 既定は 動かしません */
    var 標準の点 = Math.round(標準の字数 * 字幅 + 5);
    /* ★★マスごとの 字体を 足す★★（2026-09-11）
       ★借り物は .xlsb の 字体を くれません★（司さんの 実物 35,760マス中 ★0マス★）
       ⇒ `lib/xlsb-jitai.js` で 自分で 読んだ 物を ここで 付けます
       ⇒ 付けないと 実Excel が 9pt で 書いた 所を 11pt で 描き ★`######`★ に なります */
    if (字体) {
      for (var jk in 字体) {
        var jc = data[jk];
        if (!jc) continue;
        if (字体[jk].pt) jc.fontSize = 字体[jk].pt;
        if (字体[jk].名) jc.fontName = 字体[jk].名;
        /* ★★164未満の 書式番号は 国で 中身が 変わります★★＝借り物は 日本の 物を くれない
           ★借り物は ★間違った 字★を 言う 事が 在ります★（実測 2026-09-11）
             55番 … 実Excel「yyyy"年"m"月"」／借り物「m/d/yy」
             ⇒ D1 は 何も 言わず 直せたが、D45 は `m/d/yy` と 言って いて 直せなかった
           ⇒★測った 番号は こちらが 勝ちます★
             （ここに 書くのは ★実Excel に 聞いて 測った 番号だけ★＝当て推量で 増やさない） */
        if (字体[jk].書式) jc.numFmt = 字体[jk].書式;
      }
    }

    /* ★★図形（判子）を 台に 載せる★★（2026-09-21）
         ★場所は `xfrm` では なく ★マスと ずれ★ から 出します★
         ＝`xfrm` は ★置いた 時の 古い 数★（この 材料では 320pt／実Excel は 449.375pt）
         ＝★往復の 穴では ありません★（元の ファイルでも 449.375・経営者1 の 実測）
         ⇒★うちの 列幅で 解くので そのまま 合います★ */
    var objects = [];
    if (図形 && 図形.length && root.XlsxZukei) {
      objects = root.XlsxZukei.台に載せる形(図形, colW, 標準の点, {}, 24);
    }
    return { name: name, data: data, colW: colW, 既定の列幅: 標準の点, objects: objects,
      /* ★★ゼロを 隠す 板か★★（2026-09-25）
           経営者1 の 実測 ... 表示が 実Excel と 違う 9,163個の うち
           ★6,743個（74%）が 「Excel は 空・うちは `0`」★
           ＝★計算の 間違いでは なく 板の 設定を 読んで いなかった★ */
      ゼロを隠す: !!ゼロを隠す,
      答えが無い数: 答えが無い数,
      /* ★その ブックの 既定の 字体★＝画面も 同じ 字で 描く
         ⇒ 同じ 幅に 入る 桁数が 実Excel と 揃う（うちの 字は 細くて 多く 入って いた） */
      既定の字体名: (既定の字体 && 既定の字体.name) ? 既定の字体.name : '',
      既定の字大: (既定の字体 && 既定の字体.sz) ? 既定の字体.sz : 0,
      rowH: {}, hiddenRows: {}, hiddenCols: {}, 枠: 枠, _fromFile: true };
  }

  /* ── 保存：★元のバイト列を書き換える★ ── */

  /** セルの「今の値」（式なら答え、そうでなければ打った値） */
  function valueOf(cell) {
    var isF = typeof cell.f === 'string' && cell.f.charAt(0) === '=';
    var v = isF ? cell.d : (cell.v !== undefined && cell.v !== '' ? cell.v : cell.d);
    return v === undefined || v === null ? '' : v;
  }
  /** 開いた時の値を控えておく（'シート名|r,c' → 値） */
  function baselineOf(sheets) {
    var base = {};
    (sheets || []).forEach(function (sh) {
      Object.keys(sh.data || {}).forEach(function (k) {
        base[sh.name + '|' + k] = valueOf(sh.data[k]);
      });
    });
    return base;
  }
  /** ★開いた時から変わったセルだけ★を集める（触っていない所は1つも書かない） */
  /* ★見た目の違いを「変わった」と数えない★（2026-08-09 指示役が実機で発見）
     開いた時の控えは ファイルの表示文字（例 "1,000"）、
     計算し直した後は 生の数（1000）になる。そのまま比べると
     ★1つも触っていないのに 数式セルが全部「変わった」ことになり、
       文字を返す数式セル（記録8）に当たって保存そのものが断られた★。
     司さんの実物には BrtFmlaString が 2,866個＝必ず当たる。 */
  function normForCompare(v) {
    if (v === undefined || v === null) return '';
    if (typeof v === 'number') return String(v);
    var s = String(v).trim().replace(/[,\s¥￥]/g, '');
    if (s !== '' && !isNaN(Number(s))) return String(Number(s));   // "1,000" と 1000 は同じ
    return String(v).trim();
  }
  function changedCells(sh) {
    var out = {}, base = opened.base || {};
    Object.keys(sh.data || {}).forEach(function (k) {
      var p = k.split(','), r = parseInt(p[0], 10), c = parseInt(p[1], 10);
      if (isNaN(r) || isNaN(c)) return;
      var now = valueOf(sh.data[k]);
      if (normForCompare(now) === normForCompare(base[sh.name + '|' + k])) return;  // 変わっていない
      if (now === '') return;                                        // 空にする操作は まだ扱わない
      /* ★うちの計算が答えを出せなかったセルは、元のまま置いておく★
         #ERROR や #NAME? を書き込むと、★開ける物を自分で壊す★。
         （実物には うちのエンジンが読めない式がある。読めない物は触らないのが正しい） */
      if (typeof now === 'string' && now.charAt(0) === '#') return;
      out[k] = now;
    });
    return out;
  }
  /** ★そのシートの控えを、今の値で取り直す★
   *  うちの計算エンジンは Excel と答えが違う式がある（実物で771セル）。
   *  それを「客が変えた」と数えると、★うちの間違った答えでファイルを上書きしてしまう★。
   *  ⇒ ★シートを計算する側へ流して計算し直した直後に、控えを取り直す★
   *     こうすると「控え＝うちのエンジンから見た今のファイル」になり、
   *     ★そのあとの違い＝客が触った所★だけになる。 */
  /* ══ ★★2つ目の 引数 `除く`（2026-09-25）★★ ══
       ★★何が 起きて いたか（実測・経営者1 が 数えました）★★
         1マスだけ 打って 「書き出す」を 押すと 窓に ★「この3044か所を直して 書き出す」★
         ⇒★実際に 書き込まれたのは 1個だけ★（2通りで 数えて 1個・式は 0個）
         ⇒★★窓の 数が 嘘でした★★＝お客さんは 怖くて 「やめる」を 押します
       ★★因★★
         上の 断りの 通り、控えは ★「計算し直した 直後」に 取り直す★ 決まりです。
         ⇒★開いた 直後に 計算しなく なった ので 控えが 「本の 値」に なりました★
         ⇒1打ちで 旗が 下り、その後 計算し直すと ★答えが 控えと 違う★
         ⇒★うちの 答えが 全部 「客が 変えた」に 数えられました★
       ★★だから 計算し直した 後は また 取り直します★★
         ＝★但し お客さんが 打った マスは 取り直しては いけません★（打った 事が 消える）
         ＝★`除く` に 「板名|行,列」の 表を 渡すと その マスだけ 残します★
       ★見張り★ tests/mado-no-kazu-to-kaita-kazu-webkit.mjs */
  function rebaseSheet(sh, 除く) {
    if (!opened || !sh) return;
    Object.keys(sh.data || {}).forEach(function (k) {
      if (除く && 除く[sh.name + '|' + k]) return;   /* ★お客さんが 打った マス★ */
      opened.base[sh.name + '|' + k] = valueOf(sh.data[k]);
    });
  }

  /** ブック全体で1つでも変わったか（★0件なら元のバイト列をそのまま返す★） */
  /* ══ ★★「変わった」に 板が 増えた事も 入れます★★ ══（2026-09-21）
       ★★何が 起きて いたか★★
         `saveOpened()` は 1つも 変わって いなければ
         ★元の バイト列を そのまま 返します★（作り直さないのが 一番 安全・良い 決め）。
         ⇒でも ここが ★元に 在る 板しか 見て いません★でした
           （`indexOf(...) < 0` なら ★飛ばす★）
         ⇒★★板を 足しただけ では 「変わって いない」★★ に なり、
           ★出た ファイルが 元と 1バイトも 同じ★ でした（2026-09-21 実測・sha256 一致）
         ⇒★板を 足す 所を 直しても ここが 通さなければ 出ません★
       ★★記憶「作る道が 2本 在る時は 両方 直せ」の 形です★★
         ＝`saveXlsxLike()` を 直しただけでは 足りませんでした。
       ★消えた 板は ここでは 見て いません★
         ＝★今の 道は 板を 消せません★（消すのは 別の 話） */
  function anyChanged(sheets) {
    for (var i = 0; i < sheets.length; i++) {
      if (opened.sheetNames.indexOf(sheets[i].name) < 0) return true;   /* ★足した 板★ */
      if (Object.keys(changedCells(sheets[i])).length) return true;
    }
    return false;
  }

  /** .xlsx 用に A1 形式へ直す。★式のセルは <v>(答え)をうちの計算結果で埋める★ */
  function collectValues(sh) {
    var cells = {}, X = root.XLSX, ch = changedCells(sh);
    Object.keys(ch).forEach(function (k) {
      var p = k.split(','), val = ch[k];
      var num = typeof val === 'number' ? val : (val !== '' && !isNaN(Number(val)) ? Number(val) : null);
      cells[X.utils.encode_cell({ r: parseInt(p[0], 10), c: parseInt(p[1], 10) })] =
        num !== null ? { v: num, t: 'n' } : { v: String(val), t: 's' };
    });
    return cells;
  }

  /** .xlsx / .xlsm を直して返す */
  function saveXlsxLike(sheets) {
    return root.XlsxEdit.open(opened.bytes).then(function (book) {
      var chain = Promise.resolve();
      /* ★★うちで 足した 板も 書き出します★★（2026-09-21）
           司さん「★全部 保存しろや、断る 理由が なんか あるんか★」（ア）
           ★前は ここで 「元に無いシートは触らない」と 返して いました★
           ⇒`lib/hairanai.js` が 「入りません」と 数えて 言う 物の 因
           ⇒★「出来ない から」では ありません★＝足りないのは 口 だけ でした
         ★★足す 順に 気を 付けます★★
           ★先に 板を 足してから 値を 入れます★＝足す 前に 値を 入れると 行き先が 無い
         ★★元の 板の 値は 今まで 通り★★＝★触って いない 部品は 1バイトも 変わりません★
           （2026-09-21 実測 … 部品 14本 ⇒ 15本／★減った 0本★／
             変わった 3本＝[Content_Types].xml・rels・workbook.xml だけ／
             ★判子も 飾りも 入った 11本は 中身が 同じ★） */
      sheets.forEach(function (sh) {
        if (opened.sheetNames.indexOf(sh.name) >= 0) return;    // 元に在る板は 下で 直す
        if (typeof root.XlsxEdit.板を足す !== 'function') return;
        chain = chain.then(function () {
          root.XlsxEdit.板を足す(book, sh.name, collectValues(sh));
        });
      });
      sheets.forEach(function (sh) {
        if (opened.sheetNames.indexOf(sh.name) < 0) return;      // 足した 板は 上で 入れた
        chain = chain.then(function () {
          return root.XlsxEdit.setValues(book, sh.name, collectValues(sh));
        });
      });
      return chain.then(function () { return root.XlsxEdit.save(book); });
    });
  }

  /** .xlsb を直して返す（★歩けなければ断る★） */
  function saveXlsb(sheets) {
    var Z = root.ZipSurgeon, E = root.XlsbEdit, X = root.XLSX;
    var zip = Z.read(opened.bytes);
    // シート名 → 部品名（workbook.bin を読まずに、開いた時の並びで対応させる）
    var parts = zip.names().filter(function (n) { return /^xl\/worksheets\/sheet\d+\.bin$/.test(n); })
      .sort(function (a, b) {
        return parseInt(a.replace(/\D+/g, ''), 10) - parseInt(b.replace(/\D+/g, ''), 10);
      });
    var chain = Promise.resolve(), touched = [];
    var 触れなかった数 = 0;   /* ★答えが 字の 式（記録 8）＝触れない マスの 数★ */
    /* ══ ★★うちで 足した 板も 書き出します★★ ══（2026-09-22）
         司さん「★全部 保存しろや★」（ア）＝★司さんの 実物は この 形★
         `.xlsx` の 側は 2026-09-21 に 直しました。こちらは ★別の 道★
           ＝`workbook.bin` も `sheetN.bin` も 2進
         ★★実Excel で 数えて もらいました★★（経営者1・2026-09-22・★赤 0件★）
           `Open` が ★投げない★ ／ 判子 1つ そのまま ／
           ★溢れが 3組 とも 生きて いる★ ／ 元の 板の 値 19マス とも 同じ
         ★★数しか 書けません★★
           ＝字の マスは `sharedStrings.bin` を 指します
           ＝足すと ★元の 板が 指す 番号が ずれる 恐れ★
           ⇒★字の マスは 「入りません」と 言い続けます★（`lib/hairanai.js`）
         ★`xl/metadata.bin` は 1バイトも 触りません★（★触ると 溢れが 壊れます★） */
    if (typeof E.xlsb板を足す === 'function') {
      var 見本 = null;
      sheets.forEach(function (sh) {
        if (opened.sheetNames.indexOf(sh.name) >= 0) return;   /* 元に 在る 板は 下で 直す */
        chain = chain.then(function () {
          /* ★見え方と 行の 頭を 元の 板から 写す為に 1枚目を 見本に します★
             ＝★当て推量で 作らない★ */
          var 先 = 見本 ? Promise.resolve(見本) : zip.bytes(parts[0]).then(function (b0) {
            var r0 = E.parse(b0 instanceof Uint8Array ? b0 : new Uint8Array(b0), E.SHAPE.sheet);
            見本 = r0.ok ? r0.recs : [];
            return 見本;
          });
          return 先.then(function (み) {
            var 数だけ = {}, ch = changedCells(sh);
            Object.keys(ch).forEach(function (k) {
              var p2 = k.split(','), val = ch[k];
              if (typeof val !== 'number' && (val === '' || isNaN(Number(val)))) return;
              数だけ[X.utils.encode_cell({ r: parseInt(p2[0], 10), c: parseInt(p2[1], 10) })] =
                { v: Number(val), t: 'n' };
            });
            return E.xlsb板を足す(zip, み, sh.name, 数だけ).then(function (足) {
              touched.push(足.部品);
            });
          });
        });
      });
    }
    sheets.forEach(function (sh) {
      var i = opened.sheetNames.indexOf(sh.name);
      if (i < 0 || !parts[i]) return;
      chain = chain.then(function () {
        return zip.bytes(parts[i]).then(function (bin) {
          var cells = {}, ch = changedCells(sh);
          Object.keys(ch).forEach(function (k) {
            var val = ch[k];
            if (typeof val !== 'number' && (val === '' || isNaN(Number(val)))) return;  // 数だけ扱う
            cells[k] = Number(val);
          });
          if (!Object.keys(cells).length) return;
          var ed = E.editSheet(bin, cells);
          if (!ed.ok) throw new Error('このファイルは直せません（' + sh.name + '：' + ed.why + '）');
          /* ★★答えが 字の 式（記録 8）は 触れません★★（2026-09-25）
               ＝★前は ここで 本 1冊 まるごと 出ませんでした★
               ＝実測（経営者1・お客さんの 道）･･･ ★入力値を 1つ 直すと 書き出せない★
                 帯 ･･･「この数式セルの答えの形は まだ直せません（記録 8）」
                 ⇒★本番と 同じ 木でも 同じ＝元から です★
               ⇒★今は 触らずに 数えて 先へ 進みます★
               ⇒★★触れなかった 物が 在れば 下で 「開いたら 全部 計算しろ」の 印を 立てます★★ */
          if (ed.触れなかった && ed.触れなかった.length) {
            触れなかった数 += ed.触れなかった.length;
          }
          zip.replace(parts[i], ed.bytes);
          touched.push(parts[i]);
        });
      });
    });
    return chain.then(function () {
      /* ★触れなかった 数も 呼ぶ 側へ 渡します★＝★お客さんに 言う 為★ */
      function 数を添える(出) {
        if (出 && typeof 出 === 'object' && !(出 instanceof Uint8Array)) 出.触れなかった数 = 触れなかった数;
        return 出;
      }
      if (!touched.length) return Promise.resolve(zip.build()).then(数を添える);
      /* ══ ★★触れなかった 式が 在れば 実Excel に 計算し直させます★★ ══（2026-09-25）
           ＝★触れなかった マスの 答えは 古いまま 残ります★
           ⇒★印を 立てて おけば 開いた 時に 直ります★
           ＝印は `workbook.bin` の 記録157 の 26バイト目 ビット0
           ＝★経営者1 が 実Excel で 7冊 作って 決めた 所★（読む 側は `lib/xlsb-jitai.js`）
           ★立てられない 時は 黙って 進みます★＝★本が 出ない より 良い★
             ⇒但し ★何個 触れなかったかは 下で 言います★ */
      /* ★何個 触れなかったかを 声に 出します★＝★後から 数えられる ように★
         （見張りが この 声を 読んで 「印を 立てるべきか」を 決めます） */
      if (root.console) {
        root.console.log('[Exally] 触れなかった 式 ' + 触れなかった数 + '個'
          + '（答えが 字の 式＝記録8 は 書き換えられません）');
      }
      var 印を立てる = Promise.resolve();
      if (触れなかった数 > 0 && typeof E.全部計算の印を立てる === 'function') {
        var wbName = zip.names().filter(function (n) { return /^xl[/]workbook[.]bin$/.test(n); })[0];
        if (wbName) {
          印を立てる = zip.bytes(wbName).then(function (wbb) {
            var 出 = E.全部計算の印を立てる(wbb instanceof Uint8Array ? wbb : new Uint8Array(wbb));
            if (出 && 出.ok && 出.bytes) { zip.replace(wbName, 出.bytes); touched.push(wbName); }
          }).catch(function () { /* ★立てられなくても 本は 出します★ */ });
        }
      }
      return 印を立てる.then(function () {
      /* ★binaryIndex は「何バイト目に何がある」の索引。長さが変わると嘘になるので外す。
         ★部品を消すだけでは足りない。rels と [Content_Types].xml の参照も外す★ */
      return zip.text('[Content_Types].xml').then(function (ct) {
        var relsNames = zip.names().filter(function (n) { return /worksheets\/_rels\/sheet\d+\.bin\.rels$/.test(n); });
        var got = {}, c2 = Promise.resolve();
        relsNames.forEach(function (n) { c2 = c2.then(function () { return zip.text(n).then(function (t) { got[n] = t; }); }); });
        return c2.then(function () {
          var r = E.dropBinaryIndex(zip, ct, got);
          zip.replaceText('[Content_Types].xml', r.contentTypes);
          Object.keys(r.rels).forEach(function (n) { zip.replaceText(n, r.rels[n]); });
          return Promise.resolve(zip.build()).then(数を添える);
        });
      });
      });
    });
  }

  /** 保存の入口。★受け取ったファイルかどうかで道を分ける★ */
  function saveOpened(sheets) {
    if (!opened) return Promise.reject(new Error('開いたファイルがありません'));
    if (opened.kind === 'xls') return Promise.reject(new Error(MSG_XLS));
    /* ★1つも変わっていないなら、元のバイト列をそのまま返す★
       作り直さないのが いちばん安全（1バイトも動かない）。 */
    if (!anyChanged(sheets)) {
      return Promise.resolve({ bytes: opened.bytes, log: { compressed: 0, stored: 0, why: '', noChange: true } });
    }
    return (opened.kind === 'xlsb' ? saveXlsb(sheets) : saveXlsxLike(sheets));
  }

  root.BookOpen = {
    openFile: openFile, saveOpened: saveOpened,
    isOpened: function () { return !!opened; },
    current: function () { return opened; },
    reset: function () { opened = null; },
    detectKind: detectKind, sheetToGrid: sheetToGrid, collectValues: collectValues,
    baselineOf: baselineOf, changedCells: changedCells, valueOf: valueOf, withWeekday: withWeekday,
    anyChanged: anyChanged, normForCompare: normForCompare, rebaseSheet: rebaseSheet,
    MSG_XLS: MSG_XLS,
    MSG_VBA: MSG_VBA,
  };
})(typeof self !== 'undefined' ? self : this);
