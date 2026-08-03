/* ops/payroll.monthly.js — オペレーション「月次給与計算」（headless）
 *
 * ★鉄の掟: ここに業務ロジックを書かない。engine は検証済みの lib を呼ぶだけ。
 *   お金   … lib/payroll-monthly.js（app.js から移設した唯一の真実源）
 *   警告   … lib/payroll-warnings.js
 *   Excel … lib/payslip-xlsx.js の純関数（cells と export は同一ソース＝ズレようがない）
 *
 * 出力: { value, cells, warnings, errors, provenance }
 *   value    … 人ごとの明細と合計
 *   cells    … グリッド用のセル(AOA)。excel.export と同じ物を使う
 *   warnings … [{empId, code, level, scope, text}]
 *   errors   … [{path, code, message}] 検証NG or 個別従業員の計算失敗
 *   provenance … どのエンジンで・どの年度の法定値が実際に選ばれたか
 *
 * 【利用】ブラウザ window.OpPayrollMonthly / Node require('./ops/payroll.monthly.js')
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('../lib/op-contract.js'), require('../lib/payroll-monthly.js'),
      require('../lib/payroll-warnings.js'), require('../lib/payslip-xlsx.js'), require('../lib/shakaihoken-hyo.js'),
      require('../lib/koyo-hoken.js'), require('../lib/saitei-chingin.js'), require('../lib/statutory-meta.js'));
  } else {
    root.OpPayrollMonthly = factory(root.OpContract, root.PayrollMonthly, root.PayrollWarnings, root.PayslipXlsx,
      (typeof SHAKAIHOKEN_HYO !== 'undefined' ? SHAKAIHOKEN_HYO : root.SHAKAIHOKEN_HYO), root.KoyoHoken,
      (typeof SAITEI_CHINGIN !== 'undefined' ? SAITEI_CHINGIN : root.SAITEI_CHINGIN), root.StatutoryMeta);
  }
})(typeof self !== 'undefined' ? self : this, function (OpContract, PM, PW, Xlsx, SHH, KoyoHoken, SAI, SMeta) {
  'use strict';

  var VERSION = '1.0.0';
  var PREFS = ['hokkaido', 'aomori', 'iwate', 'miyagi', 'akita', 'yamagata', 'fukushima', 'ibaraki', 'tochigi', 'gunma',
    'saitama', 'chiba', 'tokyo', 'kanagawa', 'niigata', 'toyama', 'ishikawa', 'fukui', 'yamanashi', 'nagano', 'gifu',
    'shizuoka', 'aichi', 'mie', 'shiga', 'kyoto', 'osaka', 'hyogo', 'nara', 'wakayama', 'tottori', 'shimane', 'okayama',
    'hiroshima', 'yamaguchi', 'tokushima', 'kagawa', 'ehime', 'kochi', 'fukuoka', 'saga', 'nagasaki', 'kumamoto',
    'oita', 'miyazaki', 'kagoshima', 'okinawa'];

  // ── 入力の型（境界はここで弾く） ──
  var EMPLOYEE_SHAPE = {
    id: { type: 'string', required: true, label: '従業員ID' },
    name: { type: 'string', required: true, label: '氏名' },
    payType: { type: 'enum', values: ['月給', '時給', '日給', '歩合', '役員', 'カスタム'], label: '給与形態' },
    taxClass: { type: 'enum', values: ['ko', 'otsu', 'hei'], label: '所得税区分' },
    employmentType: { type: 'enum', values: ['employee', 'contractor'], label: '雇用形態' },
    pref: { type: 'enum', values: PREFS, label: '都道府県' },
    birthYmd: { type: 'ymd', label: '生年月日' },
    joinYmd: { type: 'ymd', label: '入社日' },
    taishokuYmd: { type: 'ymd', label: '退職日' },
    leaveStartYmd: { type: 'ymd', label: '休暇開始日' },
    leaveEndYmd: { type: 'ymd', label: '休暇終了日' },
    fuyou: { type: 'int', min: 0, label: '扶養親族等の数' },
    minWageReduce: { type: 'number', min: 0, max: 100, label: '最賃の減額特例率(%)' },
    weeklyScheduledH: { type: 'number', min: 0, max: 168, label: '週の所定労働時間' },
  };

  var INPUTS = [
    { key: 'month', type: 'ym', required: true, source: 'state.month / 期間選択', label: '対象月' },
    { key: 'company', type: 'object', required: true, source: 'pay_companies', label: '会社設定',
      of: { name: { type: 'string', required: true, label: '会社名' },
            annualHolidays: { type: 'number', min: 0, max: 365, label: '年間休日' },
            dailyWorkH: { type: 'number', min: 0, max: 24, label: '1日の所定(時)' },
            dailyWorkM: { type: 'number', min: 0, max: 59, label: '1日の所定(分)' },
            gyoshu: { type: 'enum', values: ['ippan', 'kensetsu', 'norin'], label: '雇用保険の業種' } } },
    { key: 'employees', type: 'array', required: true, minLength: 1, source: 'pay_employees', label: '従業員', of: EMPLOYEE_SHAPE },
    { key: 'ledger', type: 'array', source: 'pay_ledger(Exally台帳)', label: '台帳行' },
    { key: 'otHistory', type: 'map', source: '過去11ヶ月の確定明細', label: '36協定の履歴' },
    { key: 'options', type: 'object', label: 'オプション' },
  ];

  // ── 法令の根拠（★領域ごとに年度が違う。1枚の札で貼らない） ──
  // ★率は lib から組み立てる。文に数字を書くと、計算が正しいまま【説明文だけ】年度で取り残される。
  //   守り: tests/no-hardcoded-statutory.test.mjs が配信物への法定値の直書きを赤にする。
  var pctOf = function (v, d) { return (v * 100).toFixed(d == null ? 2 : d).replace(/\.?0+$/, '') + '%'; };
  var LAW_KOSEI_PCT = pctOf(SHH.KOSEI_NENKIN_RITSU_TOTAL, 1);           // 厚年 全体率（平成29年9月〜固定）
  var LAW_KAIGO_PCT = pctOf(SHH.getKaigo('2026-06').total, 2);          // 介護 全体率（社保年度=3月起算）
  var LAW_SHIENKIN_PCT = pctOf(SHH.SHIENKIN_TOTAL_FROM_2026_04, 2);     // 子育て支援金 全体率
  var LAW_KOYO_Y = KoyoHoken.LATEST;
  var LAW_KOYO_PER1000 = String(Math.round(KoyoHoken.RATES[LAW_KOYO_Y].ippan * 1000 * 10) / 10); // 告示の書き方(◯/1000)
  var LAW = {
    incomeTax: { basis: '所得税法（電算機計算の特例・別表）', nendo: '令和8年分(2026)', appliedBy: 'payYm の年',
      source: 'https://www.nta.go.jp/taxes/shiraberu/taxanswer/gensen/2502.htm' },
    shahoKenko: { basis: '健康保険法', nendo: '令和8年度（2026年3月分〜）', appliedBy: '社保年度=3月起算', pref: '都道府県別',
      source: 'https://www.kyoukaikenpo.or.jp/about/business/insurance_rate/rate_prefectures/r08/index.html' },
    shahoKosei: { basis: '厚生年金保険法', nendo: '平成29年9月分〜 ' + LAW_KOSEI_PCT + ' 固定', appliedBy: '—',
      source: 'https://www.kyoukaikenpo.or.jp/assets/R8_13tokyo.pdf' },
    kaigo: { basis: '介護保険法', nendo: '令和8年度 ' + LAW_KAIGO_PCT + '（全国一律・40〜64歳）', appliedBy: '社保年度=3月起算',
      source: 'https://www.kyoukaikenpo.or.jp/about/business/insurance_rate/002/index.html' },
    shienkin: { basis: '子ども・子育て支援法（子ども・子育て支援金）', nendo: '2026-04〜 ' + LAW_SHIENKIN_PCT + '（労使折半）', appliedBy: 'ym>=2026-04',
      source: 'https://www.cfa.go.jp/policies/kodomokosodateshienkinseido' },
    koyo: { basis: '雇用保険法', nendo: '令和' + (LAW_KOYO_Y - 2018) + '年度（' + LAW_KOYO_Y + '-04〜' + (LAW_KOYO_Y + 1) + '-03）一般 労働者負担 ' + LAW_KOYO_PER1000 + '/1000', appliedBy: '労働保険年度=4月起算',
      source: 'https://jsite.mhlw.go.jp/yamagata-roudoukyoku/koyouhoken-20260316.html' },
    saiteiChingin: { basis: '最低賃金法', nendo: '令和7年度（2025-10-03 発効）', appliedBy: '最賃年度=10月起算',
      source: 'https://www.mhlw.go.jp/content/11200000/001571192.pdf',
      note: '令和8年度は目安答申（2026-07-28）のみで実額未確定＝未収録。対象月が令和8年度に入ると STATUTORY_STALE で黄警告を出す（推測値を入れない）。',
      noteSource: 'https://www.mhlw.go.jp/stf/newpage_74920.html' },
    roukiho: { basis: '労働基準法 26条(休業手当)/27条(保障給)/32条(法定労働時間)/36条(時間外上限)/37条(割増)/60・61条(年少者)',
      source: 'https://laws.e-gov.go.jp/law/322AC0000000049' },
    tekiyoKakudai: { basis: '健康保険法・厚生年金保険法（短時間労働者の適用拡大）',
      current: '週20時間以上 / 所定内賃金 月8.8万円以上 / 2か月超の雇用見込み / 学生でない / 特定適用事業所(被保険者51人以上)',
      source: 'https://www.nenkin.go.jp/service/kounen/tekiyo/jigyosho/tanjikan.html',
      watch: [
        { item: '賃金要件 月8.8万円以上の撤廃', when: '令和8年10月に撤廃予定（施行日は政令事項）', status: '本エンジン未反映（切替点 WAGE_88K_REMOVED_YM は null のまま）',
          deadline: '2026-09-15 までに日本年金機構・厚労省を再照合して報告する（10月分の給与計算に間に合わせる）',
          source: 'https://www.nenkin.go.jp/service/kounen/tekiyo/jigyosho/tanjikan.html',
          source2: 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000147284_00021.html' },
        { item: '企業規模要件の段階的引下げ', when: '令和9年10月 36人以上 → 令和11年10月 21人以上 → 令和14年10月 11人以上', status: '未収録',
          source: 'https://www.nenkin.go.jp/service/kounen/tekiyo/jigyosho/tanjikan.html' },
      ] },
  };

  function ctxOf(inputs) { return { company: inputs.company, month: inputs.month, otHist: inputs.otHistory || {} }; }

  // 入力の「器」を整える（構造だけ。金額や判定には一切触らない＝業務ロジックではなくアダプタの仕事）。
  //  UI経由では mergeEmp/defEmp が器を用意しているが、オペは外(グリッド/チャット/API)からも呼ばれる。
  //  器が無いだけで例外にするのは不親切なうえ、落ちた人が黙って消えるのが一番危ない。
  function normalizeEmployee(src) {
    var e = JSON.parse(JSON.stringify(src)); // 入力を壊さない（compute は e を書き換える）
    if (!Array.isArray(e.shikyu)) e.shikyu = [];
    if (!Array.isArray(e.kintai)) e.kintai = [];
    if (!Array.isArray(e.extraKojo)) e.extraKojo = [];
    if (!Array.isArray(e.wbInclude)) e.wbInclude = [];
    if (!Array.isArray(e.wbExclude)) e.wbExclude = [];
    if (!e.apply || typeof e.apply !== 'object') e.apply = {};
    if (!e.warimashi || typeof e.warimashi !== 'object') e.warimashi = { mode: 'easy' };
    if (!e.shaho || typeof e.shaho !== 'object') e.shaho = { mode: 'teiji', months: [] };
    if (!e.payType) e.payType = '月給';
    return e;
  }

  // 実行時に「どの年度のどの率が実際に選ばれたか」を記録する（版切替が効いているかを出力で確かめられるように）
  // その kind の値が【どこから来て・いつ確かめた物か】を返す。
  //  中央(Supabase statutory)から流し込まれた時は、面が渡してくれた中央の source_url/verified_at を使う。
  //  流し込まれていない(オフライン等)時は内蔵値＝lib/statutory-meta.js の出典と確認日を使う。
  //  ★どちらの場合も空にしない。「オフラインで内蔵値を使った時こそ出典を聞かれる」ため。
  function originOf(kind, year, src) {
    var key = kind + ':' + year;
    var central = src && src[key];
    if (central) {
      return { origin: 'central', source_url: central.source_url || null, verified_at: central.verified_at || null,
        note: central.source_url ? null : '中央の行に出典URLが入っていません' };
    }
    var m = SMeta && SMeta.get ? SMeta.get(kind, year) : null;
    if (!m) return { origin: 'builtin', source_url: null, verified_at: null, note: '内蔵値・出典未登録(lib/statutory-meta.js に無い)' };
    // ★note は確認日の有無にかかわらず必ず出す。「何をどこまで確かめたか」「何が未確認か」は
    //   日付が入っていても言うべき事だから（例: 最賃は47県を突き合わせたが発効日の扱いが未解決）。
    return { origin: 'builtin', source_url: m.source_url || null, verified_at: m.verified_at || null,
      note: m.note || (m.verified_at ? null : '内蔵値・確認日は未記録') };
  }

  function statutorySnapshot(ctx, employees, statutorySource) {
    var ym = ctx.month;
    var src = statutorySource || {};
    var shahoY = SHH && SHH.shahoYearOf ? SHH.shahoYearOf(ym) : null;
    var koyoY = KoyoHoken && KoyoHoken.employYearOfYm ? KoyoHoken.employYearOfYm(ym) : null;
    var saiY = SAI && SAI.saiteiNendoOf ? SAI.saiteiNendoOf(ym) : null;
    var pref = (employees && employees[0] && employees[0].pref) || 'tokyo';
    var snap = { ym: ym };
    if (SHH && SHH.getKenko) { var k = SHH.getKenko(pref, ym); snap.kenko = Object.assign({ pref: pref, nendo: k.nendo, jugyoin: k.jugyoin, stale: !!k.stale }, originOf('shakaihoken', shahoY, src)); }
    if (SHH && SHH.getKaigo) { var g = SHH.getKaigo(ym); snap.kaigo = Object.assign({ total: g.total, jugyoin: g.jugyoin, stale: !!g.stale }, originOf('shakaihoken', shahoY, src)); }
    if (SHH && SHH.getShienkin) snap.shienkin = Object.assign({ jugyoin: SHH.getShienkin(ym) }, originOf('shakaihoken', shahoY, src));
    if (SHH && SHH.KOSEI_NENKIN_RITSU_JUGYOIN != null) snap.kosei = Object.assign({ jugyoin: SHH.KOSEI_NENKIN_RITSU_JUGYOIN }, originOf('shakaihoken', shahoY, src));
    if (KoyoHoken && KoyoHoken.employRate) snap.koyo = { gyoshu: (ctx.company || {}).gyoshu || 'ippan', rate: KoyoHoken.employRate((ctx.company || {}).gyoshu, KoyoHoken.employYearOfYm(ym)), fy: KoyoHoken.employYearOfYm(ym) }; snap.koyo = Object.assign(snap.koyo, originOf('koyo', koyoY, src));
    if (SAI && SAI.getChingin) snap.saitei = Object.assign({ pref: pref, chingin: SAI.getChingin(pref), nendo: SAI.NENDO, stale: SAI.saiteiStale ? SAI.saiteiStale(ym) : false }, originOf('saitei_chingin', saiY, src));
    return snap;
  }

  // 明細1人ぶん（app.js の buildPeople と同じ組み立て。Excel/グリッド用）
  function personOf(e, r, ctx, warnTexts) {
    var k = (e.kintai || []).filter(function (x) { if (/代休取得|振替休日/.test(x.label || '')) return PM.num(x.value) > 0; return true; });
    var oi = k.findIndex(function (x) { return /出勤/.test(x.label || ''); });
    var wt = { label: '労働時間', value: PM.workedLabel(e) };
    if (oi >= 0) k.splice(oi + 1, 0, wt); else k.unshift(wt);
    return { name: e.name, company: (ctx.company || {}).name, payDate: PM.payDateStr(ctx), kintai: k,
      shikyu: r.shikyu, kojo: r.kojo, net: r.net, shikyuTotal: r.shikyuTotal, kojoTotal: r.kojoTotal, warnings: warnTexts };
  }

  function monthLabelOf(ym) {
    var KAN = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二'];
    var y = parseInt(String(ym).slice(0, 4), 10), m = parseInt(String(ym).slice(5, 7), 10);
    return '令和' + (y - 2018) + '年' + KAN[m] + '月分';
  }

  function engine(inputs) {
    var ctx = ctxOf(inputs);
    var errors = [];
    var people = [], rows = [], warnings = [];

    (inputs.employees || []).forEach(function (src, i) {
      var e;
      try {
        e = normalizeEmployee(src);
        var r = PM.compute(e, ctx);
        var w = PW.collect(e, ctx);                 // ★compute の後に呼ぶ（prorateNote は e._prorate を読む）
        warnings = warnings.concat(w);
        var texts = PW.empWarnings(e, ctx);         // 経理向けサマリ（Excelの「要確認」列に出る文言）
        rows.push({
          empId: src.id, name: src.name,
          shikyu: r.shikyu, shikyuTotal: r.shikyuTotal, nonTaxable: r.nonTaxable,
          hyojun: r.hyojun, hyojunHealth: r.hyojunHealth, hyojunPension: r.hyojunPension,
          hasKaigo: !!r.hasKaigo, kazei: r.kazei, si: r.si,
          incomeTax: r.incomeTax, residentTax: r.residentTax,
          kojo: r.kojo, kojoTotal: r.kojoTotal, net: r.net, netNegative: !!r.netNegative,
        });
        people.push(personOf(e, r, ctx, texts));
      } catch (ex) {
        // ★黙って0円にしない: 落ちた人は errors に出し、value は部分結果と分かるようにする
        errors.push({ path: 'employees[' + i + ']', code: 'ENGINE', message: (src && src.name ? src.name + ': ' : '') + (ex && ex.message ? ex.message : String(ex)) });
      }
    });

    warnings = warnings.concat(PW.collectCompany(ctx));

    var totals = rows.reduce(function (a, r) { return { shikyuTotal: a.shikyuTotal + r.shikyuTotal, kojoTotal: a.kojoTotal + r.kojoTotal, net: a.net + r.net }; }, { shikyuTotal: 0, kojoTotal: 0, net: 0 });

    var value = { month: inputs.month, company: { name: (inputs.company || {}).name }, count: rows.length,
      people: rows, totals: totals, partial: errors.length > 0 };

    var opts = { company: (inputs.company || {}).name, monthLabel: monthLabelOf(inputs.month), filename: '給与明細_' + inputs.month + '.xlsx' };
    var used = {};
    var cells = {
      sheets: [Object.assign({ name: '集計' }, Xlsx.shukeiAOA(people, opts))]
        .concat(people.map(function (p) { return Object.assign({ name: Xlsx.sheetName(p.name, used) }, Xlsx.meishiAOA(p, opts)); })),
      _people: people, _opts: opts,
    };

    var provenance = {
      op: 'payroll.monthly', version: VERSION, validated: true,
      engines: ['lib/payroll-monthly.js', 'lib/payroll-warnings.js', 'lib/calc.js', 'lib/payroll-calc.js',
        'lib/warimashi.js', 'lib/shotokuzei-densan.js', 'lib/shotokuzei-hei.js', 'lib/zaiseki.js', 'lib/juminzei.js',
        'lib/holidays.js', 'lib/shiharai-chosho.js', 'lib/pay-rule.js', 'lib/payslip-xlsx.js'],
      law: LAW,
      statutory: statutorySnapshot(ctx, inputs.employees, (inputs.options || {}).statutorySource),
      watch: LAW.tekiyoKakudai.watch, // ★未反映の法改正を毎回出力して見えるようにする
    };

    return { value: value, cells: cells, warnings: warnings, errors: errors, provenance: provenance };
  }

  var op = OpContract.defineOperation({
    id: 'payroll.monthly',
    version: VERSION,
    title: '月次給与計算',
    desc: '1社1ヶ月の給与を全従業員ぶん計算し、明細・集計・Excel・法令の黄警告を返す',
    inputs: INPUTS,
    engine: engine,
    law: LAW,
    excel: {
      // 純関数。XLSX.writeFile はアダプタ（UI）の責務＝ここでは呼ばない
      export: function (result) {
        if (!result || !result.cells) return null;
        return { sheets: result.cells.sheets, filename: result.cells._opts.filename, opts: result.cells._opts };
      },
    },
    tests: ['tests/ops-payroll-monthly.test.mjs', 'tests/ops-golden-parity.test.mjs', 'tests/law-switchpoints.test.mjs', 'tests/op-contract.test.js',
      'tests/ops-app-parity.test.mjs',        // ★今のapp.jsの道と、このオペの道が同じ物を作るか
      'tests/statutory-freshness.test.mjs'],  // ★provenanceの出典・確認日が値と合っているか(指紋)
  });

  op.monthLabelOf = monthLabelOf;
  return op;
});
