/* file-out.js — ★ファイルを客に渡す唯一の口★（種類を正しく付ける／iPhoneは共有シート）
 *
 * なぜ必要か（2026-08-04・司さんの実機で判明）:
 *   iPhone に Excel が入っているのに、落としたファイルを ★Excelで開けなかった★。
 *   原因は端末ではなく実装:
 *     ・Blob の種類が application/octet-stream（＝「種類の分からないデータ」）だった
 *       → iPhone は Excel と紐づけられない。Excelが入っていても開けない。
 *     ・XLSX.writeFile 任せも同じ経路（種類を付けられない）
 *     ・navigator.share を1箇所も使っていなかった
 *   ★iPhoneでファイルを渡す普通のやり方は【共有シートを出す】こと。★
 *   そうすれば「Excelで開く」がその場に並ぶ。他のサイトがやっているのはこれ。
 *
 * 決まり:
 *   ① ★octet-stream を既定にしない。★ 種類の分からない物は落とさせない（拡張子から必ず決める）
 *   ② 共有シートが使える端末（iPhone等）では navigator.share({files})
 *      使えない環境（PC等）は今までどおり <a download>
 *   ③ ★分岐はこの1箇所だけ。★ 他の場所で Blob を作らない・writeFile を呼ばない
 *      （tests/ios-unsupported.test.mjs が破りを赤にする）
 *
 * 【利用】window.FileOut.deliver(bytes, 'name.xlsx') → Promise<{how:'share'|'download'}>
 */
(function (global) {
  'use strict';

  /* 拡張子 → 種類。★ここに無い拡張子は落とさない（黙って octet-stream にしない）。 */
  var MIME = {
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    txt: 'text/plain',
    pdf: 'application/pdf',
    json: 'application/json',
    png: 'image/png',
  };

  function extOf(filename) {
    var m = /\.([A-Za-z0-9]+)$/.exec(String(filename || ''));
    return m ? m[1].toLowerCase() : '';
  }
  function mimeOf(filename) {
    var e = extOf(filename);
    return MIME[e] || null;      // ★分からなければ null（呼ぶ側で止める）
  }

  function toBlob(data, mime) {
    if (typeof Blob === 'undefined') return null;
    if (data instanceof Blob) return data;
    return new Blob([data], { type: mime });
  }

  /* 共有シートが使えるか（ファイル共有に対応しているか）を、実際に聞いて確かめる。
     ★「iOSかどうか」で判定しない。端末や版で変わるので、機能があるかを聞く。 */
  function canShareFile(file) {
    try {
      return !!(global.navigator && global.navigator.canShare && global.navigator.share
        && global.navigator.canShare({ files: [file] }));
    } catch (e) { return false; }
  }

  /* ★共有シートを出すのは【指で触る端末】だけ。
     デスクトップのChromeも canShare は true を返すが、そこで共有シートを出すと
     ★今まで落ちていたファイルが落ちなくなる＝PCの退行★（実測 2026-08-04）。
     指で触る端末(pointer: coarse)＝スマホ/タブレットでは、落としても客が見つけられないので共有シートが正しい。
     UA文字列で「iPhoneか」を見ない（端末や版で変わるうえ、偽装もできる）。 */
  function prefersShare() {
    try {
      if (!global.matchMedia) return false;
      return global.matchMedia('(pointer: coarse)').matches;
    } catch (e) { return false; }
  }

  function anchorDownload(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); if (a.parentNode) a.parentNode.removeChild(a); }, 200);
    return { how: 'download', type: blob.type, filename: filename };
  }

  /* ファイルを渡す。返りは Promise（共有シートは非同期のため）。
     opts.title … 共有シートの見出し（省略時はファイル名） */
  function deliver(data, filename, opts) {
    opts = opts || {};
    var mime = opts.type || mimeOf(filename);
    if (!mime) {
      // ★種類が分からない物は落とさない。落とすと iPhone で「開けないファイル」になる。
      return Promise.reject(new Error('ファイルの種類が分かりません（' + filename + '）。FileOut.MIME に足してください。'));
    }
    var blob = toBlob(data, mime);
    if (!blob) return Promise.reject(new Error('この環境ではファイルを作れません'));

    var file = null;
    try { file = new File([blob], filename, { type: mime }); } catch (e) { file = null; }

    if (file && prefersShare() && canShareFile(file)) {
      // ★iPhone等: 共有シートを出す＝その場に「Excelで開く」が並ぶ
      return global.navigator.share({ files: [file], title: opts.title || filename })
        .then(function () { return { how: 'share', type: mime, filename: filename }; })
        .catch(function (err) {
          // 客が共有シートを閉じただけ＝エラーではない（何も起きなくてよい）
          if (err && (err.name === 'AbortError' || err.name === 'NotAllowedError')) return { how: 'cancel', type: mime, filename: filename };
          // 共有できない端末だった場合は、落とす方に切り替える（無言で失敗させない）
          return anchorDownload(blob, filename);
        });
    }
    return Promise.resolve(anchorDownload(blob, filename));
  }

  global.FileOut = { deliver: deliver, mimeOf: mimeOf, MIME: MIME, canShareFile: canShareFile, prefersShare: prefersShare, extOf: extOf };
})(typeof window !== 'undefined' ? window : globalThis);
