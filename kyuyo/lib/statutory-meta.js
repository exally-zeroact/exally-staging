/* statutory-meta.js — 法定データの【出典・確認日・指紋】。
 *
 * ★なぜ要るか（provenance のため）
 *   オペは「なぜその金額か」を後から言えないといけない。
 *   中央(Supabase statutory)から取れた時は表の source_url / verified_at を返せるが、
 *   ★オフラインで内蔵値を使った時こそ「その数字どこですか」と聞かれる場面★で、
 *   そこだけ答えられないなら provenance を作る意味がない。だから内蔵値にも出典と確認日を持たせる。
 *
 * ★なぜ「日付だけ」ではダメか（2026-08-03 の指摘）
 *   手で書いた確認日は、率を変えた人が更新し忘れれば静かに嘘になる＝「文だけ取り残される」の日付版。
 *   そこで fingerprint（その kind の値そのものから機械で作る指紋）を一緒に持つ。
 *   ★tests/statutory-freshness.test.mjs が指紋を作り直し、ここの記録と違えば【赤】。
 *     率を触った人は、必ず一次情報を見て確認日を打ち直すことになる。
 *
 * ★確認日の書き方（捏造禁止）
 *   その日に一次情報を【実際に開いて突き合わせた】kind だけ verified_at を入れる。
 *   開いていない物は verified_at:null ＋ note に「何が未確認か」を書く。混ざってよい。
 *   ★分からない物を分かったように書く方が悪い。
 *
 * 指紋の作り方: statutory-rows.js が作る「あるべき行」の data を、キー順に安定化したJSONにして
 *   FNV-1a(32bit)。★値の抜き出しは buildStatutoryRows と同じ1本を使う＝
 *   「中央へ投入する値」と「指紋を取る値」がズレようがない。
 *
 * 【利用】ブラウザ window.StatutoryMeta / Node require('./statutory-meta.js')
 */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.StatutoryMeta = api;
  else if (typeof globalThis !== 'undefined') globalThis.StatutoryMeta = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* 安定化JSON → FNV-1a(32bit) を8桁16進で。決定論（OS・改行差で揺れない）。 */
  function sortDeep(v) {
    if (Array.isArray(v)) return v.map(sortDeep);
    if (v && typeof v === 'object') {
      return Object.keys(v).sort().reduce(function (o, k) { o[k] = sortDeep(v[k]); return o; }, {});
    }
    return v;
  }
  function fingerprintOf(data) {
    var s = JSON.stringify(sortDeep(data));
    var h = 0x811c9dc5;
    for (var i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ('0000000' + h.toString(16)).slice(-8);
  }

  /* kind:year → { verified_at, source_url, fingerprint, note }
   *   verified_at … 一次情報を実際に開いて突き合わせた日（開いていなければ null）
   *   source_url  … その時に開いた一次情報のURL（開いていない物は、載せるべき一次情報のURL）
   *   fingerprint … その行の値から機械で作った指紋（値を変えたらここも変わる＝赤になる）
   */
  var META = {
    /* ── 2026-08-03 に一次情報を実際に開いて突き合わせた（3分野） ── */
    'shakaihoken:2026': {
      verified_at: '2026-08-03',
      source_url: 'https://www.kyoukaikenpo.or.jp/about/business/insurance_rate/rate_prefectures/r08/index.html',
      fingerprint: '4554fe20',
      note: '協会けんぽ 令和8年度 都道府県別料率のページを開き、東京9.85%/大阪10.13%/北海道10.28%/佐賀10.55%/沖縄9.44%・'
        + '介護1.62%（全国一律・40〜64歳）・子ども子育て支援金0.23%（令和8年4月分〜）を突き合わせた。'
        + '厚年18.3%（平成29年9月〜固定・労使折半）は日本年金機構のページで確認 '
        + 'https://www.nenkin.go.jp/service/kounen/hokenryo/hoshu/20150515-01.html 。'
        + '★47県のうち突き合わせたのは上の5県。残り42県は同じ一覧ページの表に拠る（1件ずつは確認していない）。',
    },
    'shakaihoken:2025': {
      verified_at: null,
      source_url: 'https://www.kyoukaikenpo.or.jp/g7/cat330/',
      fingerprint: 'c8dda132',
      note: '★未確認: 令和7年度(過去年度)の47県料率は今回開いていない。令和8年度のページには載っていないため。',
    },
    'koyo:2026': {
      verified_at: '2026-08-03',
      source_url: 'https://jsite.mhlw.go.jp/aichi-hellowork/list/okazaki/news/koyouhokennryouR08.html',
      fingerprint: 'd7ad3a44',
      note: '厚労省(愛知労働局ハローワーク)の令和8年度 雇用保険料率のページを開き、'
        + '一般=労働者5/1000・事業主(失業等給付5+二事業3.5=8.5)/1000、'
        + '農林水産清酒=労6/1000・事業主(6+3.5=9.5)/1000、建設=労6/1000・事業主(6+4.5=10.5)/1000 を突き合わせた。'
        + '★厚労省本体のリーフレット https://www.mhlw.go.jp/content/001692566.pdf は取得できたがPDFの中身を機械で読めなかったため、'
        + '突き合わせに使ったのは上のページ。',
    },
    'koyo:2025': {
      verified_at: null,
      source_url: 'https://www.mhlw.go.jp/content/001692566.pdf',
      fingerprint: '36dc8e47',
      note: '★未確認: 令和7年度(過去年度)の料率は今回開いていない。',
    },
    'shouhizei:2019': {
      verified_at: '2026-08-03',
      source_url: 'https://www.nta.go.jp/taxes/shiraberu/taxanswer/shohi/6101.htm',
      fingerprint: '4a4d2a88',
      note: '国税庁 No.6101 を開き、標準税率10%（うち地方消費税2.2%）・軽減税率8%（うち1.76%）を突き合わせた。',
    },

    /* ── ★以下は 2026-08-03 に一次情報を開いていない＝確認日を書かない ── */
    'saitei_chingin:2025': {
      verified_at: null,
      source_url: 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/minimumichiran/',
      fingerprint: '49975691',
      note: '★一部のみ: 一覧ページは開き「令和7年度の額が現行」であることは確認したが、'
        + '47県の金額はPDF(001571192.pdf)側にあり機械で読めなかった＝金額は突き合わせていない。'
        + '令和8年度は目安答申のみで実額未確定＝未収録（推測値を入れない）。',
    },
    'shotokuzei_densan:2025': { verified_at: null, source_url: 'https://www.nta.go.jp/users/gensen/', fingerprint: '2e37e55a', note: '★未確認: 令和7年分の電算機特例パラメータ(扶養控除/給与所得控除/基礎控除)は今回開いていない。' },
    'shotokuzei_densan:2026': { verified_at: null, source_url: 'https://www.nta.go.jp/users/gensen/2026kaisei/index.htm', fingerprint: '17488f72', note: '★未確認: 令和8年分の電算機特例パラメータと税額の算式は今回開いていない。' },
    'shotokuzei_hei:2026': { verified_at: null, source_url: 'https://www.nta.go.jp/publication/pamph/gensen/zeigakuhyo2026/02.htm', fingerprint: 'aa20a8f0', note: '★未確認: 令和8年分 源泉徴収税額表(月額表)の丙欄・別表は今回開いていない。' },
    'shotokuzei_nichi:2026': { verified_at: null, source_url: 'https://www.nta.go.jp/publication/pamph/gensen/zeigakuhyo2026/data/08-14.pdf', fingerprint: '02bdb081', note: '★未確認: 令和8年分 日額表(甲/乙/丙)は今回開いていない。' },
    'shoyo:2026': { verified_at: null, source_url: 'https://www.nta.go.jp/publication/pamph/gensen/zeigakuhyo2026/03.htm', fingerprint: '8c199970', note: '★未確認: 令和8年分 賞与に対する源泉徴収税額の算出率表と社保の上限は今回開いていない。' },
    'nenmatsu:2026': { verified_at: null, source_url: 'https://www.nta.go.jp/users/gensen/2026kaisei/index.htm', fingerprint: '4f3b1829', note: '★未確認: 令和8年分 年末調整のパラメータ(各種控除の額)は今回開いていない。' },
    'warimashi:2023': { verified_at: null, source_url: 'https://www.mhlw.go.jp/hourei/doc/kouji/K060000-A5.pdf', fingerprint: '0423f4c3', note: '★未確認: 労基法37条の割増率(時間外25%/休日35%/深夜25%/月60時間超50%)は今回開いていない。' },
  };

  function keyOf(kind, year) { return kind + ':' + year; }
  function get(kind, year) { return META[keyOf(kind, year)] || null; }
  function keys() { return Object.keys(META).sort(); }

  return { META: META, get: get, keys: keys, keyOf: keyOf, fingerprintOf: fingerprintOf, sortDeep: sortDeep };
});
