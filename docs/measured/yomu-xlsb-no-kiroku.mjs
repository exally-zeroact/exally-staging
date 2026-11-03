/* yomu-xlsb-no-kiroku.mjs
 *   -- ★`.xlsb` の 中の 1本を ★記録の 並び★に ほどく★（83）（2026-09-21）
 *
 *  ★★なぜ★★
 *    Exally1「★`xl/styles.bin` の 読み解きが 一番 大きい 残りです★」
 *    ⇒私（監査役）が 出すのは ★見立てでは なく 実物の 字★です。
 *
 *  ★★この 道具が する 事／しない 事★★
 *    する  ... 包みから `.bin` を 1本 取り出し、★記録の 頭★（番号・長さ・位置）を 並べる
 *    しない ... ★番号に 名前を 付ける★（`BrtFont` 等）
 *      ⇒★名前は 「そう 呼ばれている」だけで 実物の 証しでは ありません★
 *      ⇒名前は ★飾りを 1つだけ 変えた 物と 突き合わせて★ 決めます（別の 道具）
 *
 *  ★★記録の 形（BIFF12）★★ ... ★これは 仕様の 話なので 実物で 確かめます★
 *    番号 ... 1〜2バイト。1バイト目の 0x80 が 立って いたら もう 1バイト 続く
 *             （番号 ＝ 下位7ビット ＋ 次の バイトの 下位7ビットを 7つ 上へ）
 *    長さ ... 1〜4バイト。同じ 続き方
 *    ⇒★この 読み方が 合って いる 証し★＝★最後の 記録で ぴったり 末尾に 着く★
 *      （ずれて いたら 途中で はみ出すか 余ります ⇒ ★余り 0 を 門に します★）
 *
 *  使い方:
 *    node yomu-xlsb-no-kiroku.mjs <.xlsb の 道> <包みの 中の 名> [--生 N]
 *    例: node yomu-xlsb-no-kiroku.mjs %TEMP%\exally-tameshi-xlsb.xlsb xl/styles.bin --生 12
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const require_ = createRequire(path.join(ここ, '..', '..', 'package.json'));

/** ★zip の 中の 1本を 取り出す★（`yomu-jitsu-excel-no-shirushi.mjs` と 同じ 読み方） */
function zipから取る(buf, 欲しい) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return null;
  const 本数 = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let k = 0; k < 本数; k += 1) {
    if (buf.readUInt32LE(p) !== 0x02014b50) return null;
    const 詰め方 = buf.readUInt16LE(p + 10);
    const 圧んだ = buf.readUInt32LE(p + 20);
    const 名長 = buf.readUInt16LE(p + 28);
    const 追長 = buf.readUInt16LE(p + 30);
    const 注長 = buf.readUInt16LE(p + 32);
    const 頭位置 = buf.readUInt32LE(p + 42);
    const 名 = buf.slice(p + 46, p + 46 + 名長).toString('utf8');
    if (名 === 欲しい) {
      const 名長2 = buf.readUInt16LE(頭位置 + 26);
      const 追長2 = buf.readUInt16LE(頭位置 + 28);
      const 中身 = buf.slice(頭位置 + 30 + 名長2 + 追長2, 頭位置 + 30 + 名長2 + 追長2 + 圧んだ);
      return 詰め方 === 0 ? 中身 : require_('node:zlib').inflateRawSync(中身);
    }
    p += 46 + 名長 + 追長 + 注長;
  }
  return null;
}

/** ★7ビットずつ 継ぎ足す 数★（番号は 最大2バイト／長さは 最大4バイト） */
function 継ぐ数(buf, p, 最大) {
  let v = 0;
  for (let i = 0; i < 最大; i += 1) {
    if (p + i >= buf.length) return null;
    const b = buf[p + i];
    v |= (b & 0x7f) << (7 * i);
    if ((b & 0x80) === 0) return { 値: v >>> 0, 長: i + 1 };
  }
  return null;
}

/** ★記録の 並びに ほどく★（★名前は 付けません★） */
export function 記録に割る(buf) {
  const 出 = [];
  let p = 0;
  while (p < buf.length) {
    const 番 = 継ぐ数(buf, p, 2);
    if (!番) return { 記録: 出, 余り: buf.length - p, 止まった: p };
    const 長 = 継ぐ数(buf, p + 番.長, 4);
    if (!長) return { 記録: 出, 余り: buf.length - p, 止まった: p };
    const 頭 = p + 番.長 + 長.長;
    if (頭 + 長.値 > buf.length) return { 記録: 出, 余り: buf.length - p, 止まった: p };
    出.push({ 番号: 番.値, 長さ: 長.値, 位置: p, 中身: buf.slice(頭, 頭 + 長.値) });
    p = 頭 + 長.値;
  }
  return { 記録: 出, 余り: buf.length - p, 止まった: p };
}

/* ══ ここから下は 自分で 走らせた 時だけ ══ */
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const 道 = process.argv[2];
  const 中の名 = process.argv[3];
  const 生 = process.argv.includes('--生')
    ? Number(process.argv[process.argv.indexOf('--生') + 1]) : 0;
  if (!道 || !中の名) {
    console.log('使い方: node yomu-xlsb-no-kiroku.mjs <.xlsb> <包みの中の名> [--生 N]');
    process.exit(2);
  }
  if (!fs.existsSync(道)) { console.log('★その 包みは 在りません★ ... ' + 道); process.exit(3); }
  const 中 = zipから取る(fs.readFileSync(道), 中の名);
  if (!中) { console.log('★包みの 中に その 名は 在りません★ ... ' + 中の名); process.exit(4); }
  console.log('★包み★ ' + 道);
  console.log('★中の 名★ ' + 中の名 + '（ほどいて ' + 中.length + ' バイト）');

  const { 記録, 余り, 止まった } = 記録に割る(中);
  console.log('★記録★ ' + 記録.length + '本 ／ ★余り★ ' + 余り + ' バイト');
  if (余り !== 0) {
    console.log('★★余りが 0では ありません＝読み方が ずれて います★★（止まった 位置 ' + 止まった + '）');
    process.exit(5);
  }
  /* ★番号ごとの 数★（★名前は 付けません★） */
  const 数 = new Map();
  for (const r of 記録) 数.set(r.番号, (数.get(r.番号) || 0) + 1);
  console.log('');
  console.log('★★番号ごとの 数★★（多い 順 ／ ★名前は 付けて いません★）');
  const 並び = [...数.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  for (const [番, n] of 並び) {
    const 長たち = [...new Set(記録.filter((r) => r.番号 === 番).map((r) => r.長さ))];
    console.log('  番号 ' + String(番).padStart(5) + ' ... ' + String(n).padStart(4) + '本'
      + ' ／ 長さ ' + 長たち.slice(0, 6).join(',') + (長たち.length > 6 ? ',...' : ''));
  }
  /* ★★並びを そのまま 出す★★（★名乗りと 実際の 数を ★両方★ 出します★）
       ＝長さ4の 記録は 中の 本数を 名乗って いる ように 見えます。
       ⇒★見立てを その場で 検算します★＝閉じ（番号+1）までに 実際に 何本 在ったかと 突き合わせ
       ⇒★合わなければ ★合いません★と 出します★（黙って 名乗りを 信じない） */
  if (process.argv.includes('--並び')) {
    console.log('');
    console.log('★★並び★★（★名前は 付けて いません★）');
    /* ★閉じの 位置を 先に 決めます★（番号+1 で 長さ0 の 物を 深さを 見ながら 探す） */
    const 閉じ = new Map();
    for (let i = 0; i < 記録.length; i += 1) {
      if (記録[i].長さ !== 4) continue;
      const 開 = 記録[i].番号;
      let 深 = 0;
      for (let j = i + 1; j < 記録.length; j += 1) {
        if (記録[j].番号 === 開 && 記録[j].長さ === 4) { 深 += 1; continue; }
        if (記録[j].番号 === 開 + 1 && 記録[j].長さ === 0) {
          if (深 === 0) { 閉じ.set(i, j); break; }
          深 -= 1;
        }
      }
    }
    /* ★直下の 本数を 数えます★（入れ子の 中は 数えない） */
    const 直下 = new Map();
    for (const [開i, 閉i] of 閉じ) {
      let n = 0;
      for (let j = 開i + 1; j < 閉i; j += 1) {
        const 飛 = 閉じ.get(j);
        if (飛 !== undefined && 飛 < 閉i) { n += 1; j = 飛; continue; }
        n += 1;
      }
      直下.set(開i, n);
    }
    let ずれ = 0;
    const 深さ = new Array(記録.length).fill(0);
    for (const [開i, 閉i] of 閉じ) for (let j = 開i + 1; j < 閉i; j += 1) 深さ[j] += 1;
    for (let i = 0; i < 記録.length; i += 1) {
      const r = 記録[i];
      let 札 = '';
      if (閉じ.has(i)) {
        const 名乗り = r.中身.readUInt32LE(0);
        /* ★閉じと その 1つ前（例 612 の 前の 38）は 中身では ない ので 名乗りとは 合いません★
             ⇒★合った／合わないを ★そのまま★ 出します（合わせに いかない）★ */
        const 実 = 直下.get(i);
        札 = ' ◀ 名乗り ' + 名乗り + '本 ／ ★実際の 直下 ' + 実 + '本★'
          + (名乗り === 実 ? '（合い）' : '（★合いません★）');
        if (名乗り !== 実) ずれ += 1;
      }
      console.log('  ' + '  '.repeat(深さ[i]) + '番号 ' + String(r.番号).padStart(5)
        + ' 長さ ' + String(r.長さ).padStart(5) + 札);
    }
    console.log('');
    console.log('★名乗りと 実際が ずれた 入れ物★ ... ' + ずれ + '個'
      + (ずれ ? '（★中身で ない 記録が 混ざって いる 印★）' : ''));
  }

  if (生 > 0) {
    console.log('');
    /* ★1本ぶんを 何バイトまで 出すか★（★既定は 全部★／`--字 N` で 抑えられます） */
    const 上限 = process.argv.includes('--字')
      ? Number(process.argv[process.argv.indexOf('--字') + 1]) : 1e9;
    console.log('★★頭から ' + 生 + '本の 生の 字★★'
      + (上限 < 1e9 ? '（1本 ' + 上限 + 'バイトまで）' : '（★1バイトも 切りません★）'));
    for (const r of 記録.slice(0, 生)) {
      console.log('  位置 ' + String(r.位置).padStart(6) + ' 番号 ' + String(r.番号).padStart(5)
        + ' 長さ ' + String(r.長さ).padStart(5) + ' ｜ '
        /* ★★2026-09-25＝ここで 24バイトで 切って いました★★
             ＝記録 157 は ★26バイト★／★最後の 2バイトを 誰も 見て いませんでした★
             ＝`Iteration` と `PrecisionAsDisplayed` を 「変わらない」と 出しかけました
             ⇒★決まった 字数で 切りません★（記憶「見張りを決まった字数で切るな」）
             ⇒★切る 時は ★切ったと 書きます★★ */
        + (r.中身.length > 上限
          ? r.中身.slice(0, 上限).toString('hex').replace(/(..)/g, '$1 ').trim()
            + ' ★...ここで 切りました（全 ' + r.中身.length + 'バイトの うち ' + 上限 + '）★'
          : r.中身.toString('hex').replace(/(..)/g, '$1 ').trim()));
    }
  }
}
