/* seikyu-doc.test.mjs — ★請求書という「物」の決まりを実数で固定する★
 *
 * 代行請求の実物で起きている事故を、ここで根から止める:
 *   ① ★番号が会社の並び順から作られていて、会社を1社足すと過去の番号が変わる★
 *      → 番号は「発行した時に決めて、行に保存する」。並び順から作らない。
 *   ② ★同じ番号を二度使わない★（重複＝入金の消し込みで請求が特定できない＝二重請求・誤督促）
 *      → 形を選べるようにし、★最後の砦は倉庫の一意制約★（このlibは採番と再試行だけ）
 *   ③ ★発行した請求書は直せない・消せない★（明細を1行直すと去年の紙の金額が変わる、を止める）
 *   ④ ★取引先を消しても一覧に名前が出る★（紙と同じ物＝snapshot から出す）
 *   ⑤ ★入金は「1件も無い」と「取れなかった（未確認）」を必ず違う物として返す★
 *   ⑥ ★過入金を0でクランプしない★
 *
 * 使い方: node seikyu/tests/seikyu-doc.test.mjs
 *         node seikyu/tests/seikyu-doc.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const D = require_(path.join(ROOT, 'seikyu/lib/seikyu-doc.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };
const ok = (v, m) => { if (!v) throw new Error(m || 'false'); };
const hasErr = (errs, word) => errs.some(e => String(e.msg || e).indexOf(word) >= 0);

/* ── self-test：わざと壊して赤になるかを先に見せる ───────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n[seikyu-doc --self-test] わざと壊して赤になるか');
  let sp = 0, sf = 0;
  const S = (n, fn) => { try { fn(); sp++; console.log('  ✓ ' + n); } catch (e) { sf++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

  S('★入金が「未確認(null)」と「0件([])」で違う答えを返す（同じなら検査が空振り）', () => {
    const inv = { grand_total: 1000 };
    const a = D.paymentStateOf(inv, null).state;
    const b = D.paymentStateOf(inv, []).state;
    if (a === b) throw new Error('未確認と0件が同じ答え（' + a + '）＝区別できていない');
    if (a !== 'unknown' || b !== 'unpaid') throw new Error('実測値がずれた a=' + a + ' b=' + b);
  });

  S('★過入金が0でクランプされていない（クランプしていたらここが赤）', () => {
    const r = D.paymentStateOf({ grand_total: 1000 }, [{ amount: 1500, ymd: '2026-08-01' }]);
    if (r.remain >= 0) throw new Error('残額が0以上＝クランプされている: ' + r.remain);
    if (r.state !== 'over') throw new Error('過入金と判定していない: ' + r.state);
  });

  S('★取引先が消えた時に空文字を返していない（空なら赤）', () => {
    const inv = { status: 'draft', partner_id: 'pt_x', snapshot: {} };
    const name = D.partnerNameOf(inv, {});
    if (!name) throw new Error('空を返した＝取れなかったのに空欄になる');
  });

  S('★発行済みを編集/削除できると赤', () => {
    if (D.canEdit({ status: 'issued' })) throw new Error('発行済みが編集できてしまう');
    if (D.canDelete({ status: 'issued' })) throw new Error('発行済みが消せてしまう');
    if (!D.canDelete({ status: 'draft' })) throw new Error('下書きが消せない');
  });

  S('★形と連番の実測がズレたら赤（採番が空振りしていない）', () => {
    const n = D.nextNo({ format: 'ym-seq', ymd: '2026-09-30', existing: ['202609-001', '202609-002'] });
    if (n !== '202609-003') throw new Error('採番がずれた: ' + n);
    const first = D.nextNo({ format: 'ym-seq', ymd: '2026-09-30', existing: [] });
    if (first !== '202609-001') throw new Error('最初の番号がずれた: ' + first);
  });

  console.log('\n[self-test] ' + sp + ' passed, ' + sf + ' failed');
  process.exit(sf ? 1 : 0);
}

/* ── 本番の検査 ───────────────────────────────────────────────── */
console.log('\n[seikyu-doc] 請求書という「物」の決まり');

/* ① 番号の形 -------------------------------------------------------- */
T('★形は5つ（連番／年+連番／年月+連番／取引先+年月+連番／自分で決める）', () => {
  eq(D.NUMBER_FORMATS.map(f => f.key).join(','), 'seq,y-seq,ym-seq,p-ym-seq,manual');
});

T('★それぞれの形が指示どおりの見た目になる', () => {
  const o = { ymd: '2026-09-30', partnerCode: 'A001' };
  eq(D.formatNo({ ...o, format: 'seq', seq: 1 }), '00001');
  eq(D.formatNo({ ...o, format: 'y-seq', seq: 1 }), '2026-0001');
  eq(D.formatNo({ ...o, format: 'ym-seq', seq: 1 }), '202609-001');
  eq(D.formatNo({ ...o, format: 'p-ym-seq', seq: 1 }), 'A001-202609-01');
  eq(D.formatNo({ ...o, format: 'manual', seq: 1 }), '', '自由入力は自動で作らない');
});

T('★境界(空)：まだ1通も無ければ 1 番から', () => {
  eq(D.nextNo({ format: 'seq', ymd: '2026-09-30', existing: [] }), '00001');
  eq(D.nextNo({ format: 'y-seq', ymd: '2026-09-30', existing: [] }), '2026-0001');
  eq(D.nextNo({ format: 'p-ym-seq', ymd: '2026-09-30', partnerCode: 'A001', existing: [] }), 'A001-202609-01');
});

T('★境界(端)：桁があふれても切らずに伸ばす（99999 の次は 100000）', () => {
  eq(D.nextNo({ format: 'seq', ymd: '2026-09-30', existing: ['99998'] }), '99999');
  eq(D.nextNo({ format: 'seq', ymd: '2026-09-30', existing: ['99999'] }), '100000');
  eq(D.nextNo({ format: 'ym-seq', ymd: '2026-09-30', existing: ['202609-999'] }), '202609-1000');
});

T('★境界(不明)：末尾が数でない番号が混じっても落ちず、無視して数える', () => {
  eq(D.nextNo({ format: 'ym-seq', ymd: '2026-09-30', existing: ['202609-abc', '202609-002', ''] }), '202609-003');
});

T('★自分で決めた番号も「使用済み」として数える（自動採番がぶつからない）', () => {
  eq(D.nextNo({ format: 'ym-seq', ymd: '2026-09-30', existing: ['202609-001', '202609-050'] }), '202609-051');
});

T('★年をまたぐ：既定は1に戻す／「続ける」を選べば続く', () => {
  eq(D.nextNo({ format: 'y-seq', ymd: '2026-01-05', existing: ['2025-0127'] }), '2026-0001', '既定=戻す');
  eq(D.nextNo({ format: 'y-seq', ymd: '2026-01-05', existing: ['2025-0127'], resetYearly: false }), '2026-0128', '続ける');
  eq(D.nextNo({ format: 'ym-seq', ymd: '2026-01-05', existing: ['202512-007'], resetYearly: false }), '202601-008');
});

T('★「連番だけ」の形で「毎年1に戻す」は選ばせない（去年の番号と必ずぶつかる）', () => {
  const e = D.validateNumbering({ format: 'seq', resetYearly: true });
  ok(e.length > 0, '通ってしまった');
  ok(hasErr(e, 'ぶつか') || hasErr(e, '重複'), '理由が書かれていない: ' + JSON.stringify(e));
  eq(D.validateNumbering({ format: 'seq', resetYearly: false }).length, 0);
  eq(D.validateNumbering({ format: 'y-seq', resetYearly: true }).length, 0);
});

T('★取引先コードが空なら、空欄で作らず赤で止める（p-ym-seq）', () => {
  const e = D.validateNumbering({ format: 'p-ym-seq', partnerCode: '' });
  ok(e.length > 0, '空のコードで通ってしまった');
  eq(D.validateNumbering({ format: 'p-ym-seq', partnerCode: 'A001' }).length, 0);
  // ★空のまま作らせない（'-202609-01' のような番号を生まない）
  eq(D.formatNo({ format: 'p-ym-seq', seq: 1, ymd: '2026-09-30', partnerCode: '' }), '');
});

T('★倉庫に弾かれた時に番号を1つ進めて出し直せる（同時発行の再試行）', () => {
  eq(D.bumpNo('202609-003'), '202609-004');
  eq(D.bumpNo('2026-0001'), '2026-0002');
  eq(D.bumpNo('A001-202609-09'), 'A001-202609-10');
  eq(D.bumpNo('99999'), '100000');
  eq(D.bumpNo('手書き'), '', '数で終わらない番号は自動で進めない（人に決めさせる）');
});

/* ② 支払期限 -------------------------------------------------------- */
T('★決めていなければ期限を作らない（勝手な日付を出さない）', () => {
  eq(D.dueDateFrom('2026-08-31', { kind: 'none' }), '');
  eq(D.dueDateFrom('2026-08-31', null), '');
  eq(D.dueDateFrom('', { kind: 'eom' }), '');
  eq(D.dueDateFrom('2026-13-45', { kind: 'eom' }), '', '読めない日付');
});

T('★境界(月末・うるう年・年またぎ)を実物で測る', () => {
  eq(D.dueDateFrom('2026-02-10', { kind: 'eom' }), '2026-02-28', '平年の2月末');
  eq(D.dueDateFrom('2028-02-10', { kind: 'eom' }), '2028-02-29', 'うるう年の2月末');
  eq(D.dueDateFrom('2026-12-15', { kind: 'nextEom' }), '2027-01-31', '年またぎの翌月末');
  eq(D.dueDateFrom('2026-01-31', { kind: 'days', n: 30 }), '2026-03-02', '30日後');
  eq(D.dueDateFrom('2026-01-15', { kind: 'nextDay', n: 31 }), '2026-02-28', '翌月31日は無い→末日');
  eq(D.dueDateFrom('2026-12-20', { kind: 'nextDay', n: 10 }), '2027-01-10', '年またぎの翌月10日');
});

/* ③ 発行したら固まる ------------------------------------------------ */
T('★下書きだけが直せる・消せる。発行済みは取り消すだけ（行は残す）', () => {
  ok(D.canEdit({ status: 'draft' })); ok(D.canDelete({ status: 'draft' }));
  ok(!D.canEdit({ status: 'issued' })); ok(!D.canDelete({ status: 'issued' }));
  ok(!D.canEdit({ status: 'void' })); ok(!D.canDelete({ status: 'void' }));
  ok(D.canVoid({ status: 'issued' })); ok(!D.canVoid({ status: 'draft' }));
});

T('★固まる列の一覧が在る（倉庫のトリガと突き合わせる元になる）', () => {
  ok(D.FROZEN_FIELDS.length >= 8, '固まる列が少なすぎる');
  for (const f of ['no', 'partner_id', 'issue_ymd', 'due_ymd', 'lines', 'totals', 'snapshot', 'tax_mode', 'rounding']) {
    ok(D.FROZEN_FIELDS.indexOf(f) >= 0, f + ' が固まる列に入っていない');
  }
  ok(D.FROZEN_FIELDS.indexOf('status') < 0, 'status は取り消しのため変えられる必要がある');
  ok(D.FROZEN_FIELDS.indexOf('sent_at') < 0, 'sent_at は送った記録なので後から入る');
});

T('★発行の写し(snapshot)に、紙に出る物が全部入る', () => {
  const s = D.snapshotOf({
    partner: { id: 'pt_1', data: { name: '藤原建設株式会社', honor: '御中', addr: '今治市…', invoiceNo: 'T1', code: 'A001' } },
    org: { data: { yago: '合同会社ZEROact', addr: '今治市…', tel: '090', invoiceNo: 'T3500003003293', bank: '伊予銀行…' } },
    tax: { subtotal: 2000, taxTotal: 180, grandTotal: 2180, byRate: [{ pct: 10, base: 1000, tax: 100 }], exempt: { base: 0 }, hasReduced: true },
    templateId: 'std1',
  });
  eq(s.partner.name, '藤原建設株式会社');
  eq(s.partner.honor, '御中');
  eq(s.org.invoiceNo, 'T3500003003293');
  eq(s.totals.grandTotal, 2180);
  eq(s.templateId, 'std1');
  ok(s.byRate.length === 1, '税率ごとの区分が写しに入っていない');
  ok(s.hasReduced === true, '軽減の印が写しに入っていない');
});

/* ④ 取引先を消しても名前が出る -------------------------------------- */
T('★発行済みは写しの名前を出す（取引先を消しても紙と同じ名前が一覧に出る）', () => {
  const inv = { status: 'issued', partner_id: 'pt_1', snapshot: { partner: { name: '藤原建設株式会社' } } };
  eq(D.partnerNameOf(inv, {}), '藤原建設株式会社');
  // ★マスタを後から直しても、発行済みの一覧は紙のままであること
  eq(D.partnerNameOf(inv, { pt_1: { data: { name: '別の名前' } } }), '藤原建設株式会社');
});

T('★下書きはマスタの名前。マスタから消えていても空欄にしない', () => {
  const draft = { status: 'draft', partner_id: 'pt_1', snapshot: {} };
  eq(D.partnerNameOf(draft, { pt_1: { data: { name: 'Lounge Chouchou' } } }), 'Lounge Chouchou');
  const name = D.partnerNameOf(draft, {});
  ok(name.length > 0, '空欄になった');
  ok(name.indexOf('消え') >= 0 || name.indexOf('不明') >= 0, '取れなかったと分かる文言でない: ' + name);
  eq(D.partnerNameOf({ status: 'draft', partner_id: '', snapshot: {} }, {}).length > 0, true, '取引先未選択でも空にしない');
});

/* ⑤⑥ 入金 ---------------------------------------------------------- */
T('★入金が「未確認(取れなかった)」と「まだ0件」で違う（0件・異常なしにしない）', () => {
  const inv = { grand_total: 1000 };
  eq(D.paymentStateOf(inv, null).state, 'unknown');
  eq(D.paymentStateOf(inv, null).paid, null, '未確認の金額は0ではなくnull');
  eq(D.paymentStateOf(inv, []).state, 'unpaid');
  eq(D.paymentStateOf(inv, []).paid, 0);
});

T('★分けて払われても全部数える（一部入金・複数回）', () => {
  const inv = { grand_total: 10000 };
  const r = D.paymentStateOf(inv, [
    { amount: 3000, ymd: '2026-08-01' },
    { amount: 4000, ymd: '2026-08-20' },
  ]);
  eq(r.paid, 7000); eq(r.remain, 3000); eq(r.state, 'partial'); eq(r.count, 2);
  eq(r.lastYmd, '2026-08-20', '最後に入った日');
});

T('★ちょうど払い終わったら入金済', () => {
  eq(D.paymentStateOf({ grand_total: 10000 }, [{ amount: 10000, ymd: '2026-08-01' }]).state, 'paid');
});

T('★過入金を0でクランプしない（多く入った事実を残す）', () => {
  const r = D.paymentStateOf({ grand_total: 10000 }, [{ amount: 12000, ymd: '2026-08-01' }]);
  eq(r.paid, 12000); eq(r.remain, -2000); eq(r.state, 'over');
});

T('★消した入金・別の請求の入金は数えない', () => {
  const inv = { id: 'iv_1', grand_total: 10000 };
  const r = D.paymentStateOf(inv, [
    { invoice_id: 'iv_1', amount: 3000, ymd: '2026-08-01' },
    { invoice_id: 'iv_1', amount: 5000, ymd: '2026-08-02', deleted_at: '2026-08-03T00:00:00Z' },
    { invoice_id: 'iv_2', amount: 9999, ymd: '2026-08-02' },
  ]);
  eq(r.paid, 3000); eq(r.count, 1);
});

T('★返金(マイナス)も記録として数える', () => {
  const r = D.paymentStateOf({ grand_total: 10000 }, [{ amount: 10000, ymd: '2026-08-01' }, { amount: -2000, ymd: '2026-09-01' }]);
  eq(r.paid, 8000); eq(r.state, 'partial'); eq(r.remain, 2000);
});

T('★0円の請求は「入金済」に化けない（0で割らない・状態が壊れない）', () => {
  eq(D.paymentStateOf({ grand_total: 0 }, []).state, 'unpaid');
  eq(D.paymentStateOf({ grand_total: 0 }, [{ amount: 100, ymd: '2026-08-01' }]).state, 'over');
});

/* ⑦ 発行前の検査 ---------------------------------------------------- */
const goodInv = () => ({
  doc_type: 'invoice', no: '202609-001', partner_id: 'pt_1',
  issue_ymd: '2026-09-30', due_ymd: '2026-10-31',
  tax_mode: 'inclusive', rounding: 'floor',
  lines: [{ name: '運転業務委託料', amount: 1100, rate: 10 }],
});
const goodPartner = { id: 'pt_1', data: { name: '藤原建設株式会社', code: 'A001' } };
const goodOrg = { data: { yago: '合同会社ZEROact', invoiceNo: 'T3500003003293' } };

T('★そろっていれば発行できる', () => {
  const r = D.validateInvoice({ inv: goodInv(), partner: goodPartner, org: goodOrg });
  ok(r.ok, JSON.stringify(r.errors));
});

T('★足りない物は空欄で通さず、1つずつ理由を出す', () => {
  const cases = [
    [{ partner_id: '' }, '取引先'],
    [{ no: '' }, '番号'],
    [{ issue_ymd: '' }, '請求日'],
    [{ lines: [] }, '明細'],
    [{ due_ymd: '2026-09-29' }, '期限'],
  ];
  for (const [patch, word] of cases) {
    const r = D.validateInvoice({ inv: { ...goodInv(), ...patch }, partner: goodPartner, org: goodOrg });
    ok(!r.ok, JSON.stringify(patch) + ' が通ってしまった');
    ok(hasErr(r.errors, word), JSON.stringify(patch) + ' の理由に「' + word + '」が無い: ' + JSON.stringify(r.errors));
  }
});

T('★合計0円の請求書は出せない（マイナスは出せるが注意を出す）', () => {
  const zero = { ...goodInv(), lines: [{ name: 'x', amount: 0, rate: 10 }] };
  ok(!D.validateInvoice({ inv: zero, partner: goodPartner, org: goodOrg }).ok, '0円が通った');
  const minus = { ...goodInv(), lines: [{ name: '値引', amount: -1100, rate: 10 }] };
  const r = D.validateInvoice({ inv: minus, partner: goodPartner, org: goodOrg });
  ok(r.ok, 'マイナスが赤になった（返金の請求書は出せるべき）');
  ok(r.warnings.length > 0, 'マイナスなのに注意が出ていない');
});

T('★登録番号が無い時は「赤」ではなく「注意」（免税事業者も請求書は出す）', () => {
  const r = D.validateInvoice({ inv: goodInv(), partner: goodPartner, org: { data: { yago: 'x' } } });
  ok(r.ok, '登録番号が無いだけで赤になった');
  ok(hasErr(r.warnings, '登録番号'), '注意が出ていない: ' + JSON.stringify(r.warnings));
});

T('★取引先がマスタに無いのに発行しようとしたら赤', () => {
  const r = D.validateInvoice({ inv: goodInv(), partner: null, org: goodOrg });
  ok(!r.ok); ok(hasErr(r.errors, '取引先'));
});

T('★税の計算が通らない明細は、そのまま理由が出る（税率が空など）', () => {
  const bad = { ...goodInv(), lines: [{ name: 'x', amount: 1100 }] };
  const r = D.validateInvoice({ inv: bad, partner: goodPartner, org: goodOrg });
  ok(!r.ok); ok(hasErr(r.errors, '税率'), JSON.stringify(r.errors));
});

T('★1000行までは通り、1001行は赤（黙って切らない）', () => {
  const mk = (n) => Array.from({ length: n }, (_, i) => ({ name: 'r' + i, amount: 100, rate: 10 }));
  ok(D.validateInvoice({ inv: { ...goodInv(), lines: mk(1000) }, partner: goodPartner, org: goodOrg }).ok);
  ok(!D.validateInvoice({ inv: { ...goodInv(), lines: mk(1001) }, partner: goodPartner, org: goodOrg }).ok);
});

T('★見積は後から足せる形になっている（同じ棚・別の番号系列・請求への変換元を持てる）', () => {
  ok(D.DOC_TYPES.indexOf('quote') >= 0, '見積の型が無い');
  const q = { ...goodInv(), doc_type: 'quote' };
  ok(D.validateInvoice({ inv: q, partner: goodPartner, org: goodOrg }).ok, '見積が検査を通らない');
  // 見積→請求：変換元を持つ形であること
  const conv = D.convertQuoteToInvoice({ id: 'iv_q1', doc_type: 'quote', no: 'Q-001', lines: q.lines, partner_id: 'pt_1' });
  eq(conv.doc_type, 'invoice');
  eq(conv.quote_from, 'iv_q1');
  eq(conv.status, 'draft');
  eq(conv.no, '', '番号は請求の系列で採り直す（見積の番号を持ち込まない）');
});

/* ⑧ 実データの形（代行請求で毎日動いている形が入るか） ---------------- */
T('★実測した規模（1通 最大69行）が余裕で入る', () => {
  const lines = Array.from({ length: 69 }, (_, i) => ({ name: '行き先' + i, amount: 1300, rate: 10 }));
  const r = D.validateInvoice({ inv: { ...goodInv(), lines }, partner: goodPartner, org: goodOrg });
  ok(r.ok, JSON.stringify(r.errors));
  eq(r.tax.grandTotal, 89700);
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
