/* org-prefill.test.mjs — 請求書/見積の「自社情報 自動プリフィル」の純関数テスト
 *
 * ねらい（Exallyの旗＝二度手間ゼロ）:
 *   E0で作った共有マスタ pay_org（屋号・住所・電話・インボイス番号）を、
 *   請求書/見積の発行者欄に自動で入れる。毎回手で書かせない。
 *
 * ★守る事★
 *   ・pay_org が未設定なら何もしない（空欄のまま。勝手に埋めない＝捏造しない）
 *   ・すでに人が入力している欄は上書きしない（手入力が勝つ）
 *   ・入れた後も自由に直せる
 *   ・法定データ(消費税率など)の既存の読みには一切関与しない
 */
import assert from 'node:assert';
import OrgPrefill from '../js/org-prefill.js';   // UMD(=CommonJS)なので default 経由で受ける
const { pickOrgFields, planPrefill } = OrgPrefill;

const T = [];
function test(name, fn) { T.push({ name, fn }); }

/* ═══ pay_org から使う項目だけ取り出す ═══ */

test('pay_org から 屋号・住所・電話・インボイス番号 を取り出す', () => {
  const o = pickOrgFields({ yago: '株式会社ゼロアクト', addr: '愛媛県今治市1-2-3', tel: '0898-00-0000', invoiceNo: 'T1234567890123', sealDataUrl: 'x', businesses: ['代行'] });
  assert.deepStrictEqual(o, { issuerName: '株式会社ゼロアクト', issuerAddr: '愛媛県今治市1-2-3', issuerTel: '0898-00-0000', invoiceNo: 'T1234567890123' });
});

test('pay_org が null / 空でも落ちない（空を返す）', () => {
  assert.deepStrictEqual(pickOrgFields(null), {});
  assert.deepStrictEqual(pickOrgFields(undefined), {});
  assert.deepStrictEqual(pickOrgFields({}), {});
});

test('一部だけ設定されている時は、その項目だけ返す', () => {
  assert.deepStrictEqual(pickOrgFields({ yago: 'ゼロアクト' }), { issuerName: 'ゼロアクト' });
  assert.deepStrictEqual(pickOrgFields({ invoiceNo: 'T9' }), { invoiceNo: 'T9' });
});

test('空文字や空白だけの項目は「未設定」として扱う（空白で埋めない）', () => {
  assert.deepStrictEqual(pickOrgFields({ yago: '', addr: '   ', tel: null, invoiceNo: undefined }), {});
});

test('前後の空白は落として入れる', () => {
  assert.deepStrictEqual(pickOrgFields({ yago: '  ゼロアクト  ' }), { issuerName: 'ゼロアクト' });
});

/* ═══ どの欄を実際に埋めるか（手入力を尊重する） ═══ */

test('空の欄には入れる', () => {
  const plan = planPrefill({ issuerName: 'ゼロアクト', issuerAddr: '今治市' }, { issuerName: '', issuerAddr: '' });
  assert.deepStrictEqual(plan, { issuerName: 'ゼロアクト', issuerAddr: '今治市' });
});

test('★すでに人が入れている欄は上書きしない（手入力が勝つ）', () => {
  const plan = planPrefill(
    { issuerName: 'ゼロアクト', issuerAddr: '今治市', issuerTel: '0898' },
    { issuerName: '別の会社名', issuerAddr: '', issuerTel: '090-1111-2222' }
  );
  assert.deepStrictEqual(plan, { issuerAddr: '今治市' }, '手入力済みの欄まで書き換えている');
});

test('空白だけの欄は「空」とみなして入れる', () => {
  const plan = planPrefill({ issuerName: 'ゼロアクト' }, { issuerName: '   ' });
  assert.deepStrictEqual(plan, { issuerName: 'ゼロアクト' });
});

test('★pay_org が未設定なら1つも埋めない（勝手に作らない）', () => {
  assert.deepStrictEqual(planPrefill({}, { issuerName: '', issuerAddr: '' }), {});
  assert.deepStrictEqual(planPrefill(pickOrgFields(null), { issuerName: '' }), {});
});

test('画面に無い欄は計画に入れない（存在しないidを触らない）', () => {
  const plan = planPrefill({ issuerName: 'ゼロアクト', invoiceNo: 'T1' }, { issuerName: '' });  // invoiceNo欄が無い画面
  assert.deepStrictEqual(plan, { issuerName: 'ゼロアクト' });
});

test('全部埋まっていれば何もしない（再訪しても上書きされない）', () => {
  const cur = { issuerName: 'A', issuerAddr: 'B', issuerTel: 'C', invoiceNo: 'D' };
  assert.deepStrictEqual(planPrefill({ issuerName: 'X', issuerAddr: 'Y', issuerTel: 'Z', invoiceNo: 'W' }, cur), {});
});

/* ═══ 実データ相当の通し ═══ */

test('★実データ相当: pay_org 1件 → 請求書の発行者4欄が埋まる', () => {
  const payOrg = { yago: '株式会社ゼロアクト', addr: '愛媛県今治市○○町1-2-3', tel: '0898-00-0000', invoiceNo: 'T1234567890123', businesses: ['代行', '空調'] };
  const 画面 = { issuerName: '', issuerAddr: '', issuerTel: '', invoiceNo: '' };
  const plan = planPrefill(pickOrgFields(payOrg), 画面);
  assert.deepStrictEqual(plan, {
    issuerName: '株式会社ゼロアクト',
    issuerAddr: '愛媛県今治市○○町1-2-3',
    issuerTel: '0898-00-0000',
    invoiceNo: 'T1234567890123'
  });
  // 事業一覧や印影は請求書に関係ないので混ぜない
  assert.strictEqual('businesses' in plan, false);
  assert.strictEqual('sealDataUrl' in plan, false);
});

/* ═══ 実行 ═══ */
let ng = 0;
for (const t of T) {
  try { t.fn(); console.log('  ok   ' + t.name); }
  catch (e) { ng++; console.log('  NG   ' + t.name + '\n       ' + (e && e.message)); }
}
console.log('\norg-prefill: ' + (T.length - ng) + '/' + T.length + ' passed');
process.exit(ng ? 1 : 0);
