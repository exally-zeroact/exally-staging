/* shiki-chizu.test.mjs — ★頼りの 地図（順番）★（2026-09-11）
 *
 *  ★★土台を 自分で 作る ③枚目★★
 *
 *  ★★答えは 実Excel に 聞いて 測りました★★（2026-09-11・新しい 本に 打って 読んだ）
 *    ★逆向きの 鎖★（後ろの マスが 前を 見る）
 *      A1 `=A2+1` → 8 ／ A2 `=A3+1` → 7 ／ A3 `=A4+1` → 6 ／ A4 は 数 5
 *    ★枝分かれ★ B1 `=A4*2` → 10 ／ B2 `=B1+A1` → 18 ／ C1 `=B1+B2` → 28
 *    ★四角で まとめて★ D1 `=A4` → 5 ／ D2 `=D1*2` → 10 ／ D3 `=D2*2` → 20 ／ E1 `=SUM(D1:D3)` → 35
 *    ★列 まるごと★ F1 `=SUM(D:D)` → 35
 *    ★輪（ぐるぐる）★ A1 `=A2+1` と A2 `=A1+1`、B1 `=B1+1`
 *      ⇒実Excel は ★Iteration=False で 正しい 答えを 出さず 名指しで 断る★
 *      ⇒だから この 台も ★輪は 名指しで 返す（勝手に 直さない）★
 *
 *  ★★一番 効く 測り方★★
 *    ★出した 順に 1回 回すだけで 実Excel と 同じ 答えが 出るか★
 *    順番が 1つでも 狂うと ★古い 答え★が 出て すぐ 赤に なります。
 *
 *  ★★実物での 実測（司さんの 実物・写し）★★
 *    式 15,799本 ／ 結んだ 線 470,088本 ／ ★輪 0★ ／ ★読めず 0★ ／ 862ms
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const C = require_(path.join(ROOT, 'lib/shiki-chizu.js'));
const 参 = require_(path.join(ROOT, 'lib/shiki-sansho.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[shiki-chizu] ★頼りの 地図（順番）★');

/* ★実Excel に 打った 物と 同じ 並び★（A4 だけ ただの 数） */
const 式 = {
  'A1': '=A2+1', 'A2': '=A3+1', 'A3': '=A4+1',
  'B1': '=A4*2', 'B2': '=B1+A1', 'C1': '=B1+B2',
  'D1': '=A4', 'D2': '=D1*2', 'D3': '=D2*2', 'E1': '=SUM(D1:D3)',
  'F1': '=SUM(D:D)'
};
/* ★実Excel が 出した 答え★ */
const 実Excel = { A1: 8, A2: 7, A3: 6, B1: 10, B2: 18, C1: 28, D1: 5, D2: 10, D3: 20, E1: 35, F1: 35 };
const 数のマス = { 'A4': 5 };

function 板を作る(表) {
  const 式2 = {};
  for (const a in 表) { const p = 参.読む(a); 式2[p.上 + ',' + p.左] = 表[a]; }
  return [{ 名: 'Sheet1', 式: 式2 }];
}
function 札に(a) { const p = 参.読む(a); return 'Sheet1!' + p.上 + ',' + p.左; }

const 地図 = C.地図を作る(板を作る(式));

T('★★読めない 式が 0本★★', () => {
  if (地図.読めず.length) throw new Error('★' + 地図.読めず.length + '本★ ' + JSON.stringify(地図.読めず[0]));
  if (地図.順番.length !== Object.keys(式).length) throw new Error('順番 ' + 地図.順番.length + ' / ' + Object.keys(式).length);
});

T('★★頼り先が 必ず 前に 居る（逆転 0）★★', () => {
  const 位置 = new Map(); 地図.順番.forEach((k, i) => 位置.set(k, i));
  const 悪 = [];
  for (const k of 地図.順番) for (const d of 地図.頼り[k]) {
    if (!位置.has(d) || 位置.get(d) >= 位置.get(k)) 悪.push(d + ' → ' + k);
  }
  if (悪.length) throw new Error('★' + 悪.length + '本 逆転★ ' + 悪.slice(0, 3).join(' ／ '));
});

/* ══ ★★ここが 芯★★ ══
   ★出した 順に 1回 回すだけで 実Excel と 同じ 答えが 出るか★
   ★下の 計算は「測る 為の ちいさな 物」です＝本番の 計算では ありません★
   （本番の 計算は ④＝まだ 作って いません） */
function ちいさく計算(順番, 頼り) {
  const 値 = {};
  for (const a in 数のマス) 値[札に(a)] = 数のマス[a];
  const 読み値 = (札) => (値[札] === undefined ? 0 : 値[札]);
  for (const 札 of 順番) {
    const [, rc] = 札.split('!');
    const 元 = Object.keys(式).find((a) => 札に(a) === 札);
    let f = 式[元].slice(1);
    /* SUM(四角) */
    f = f.replace(/SUM\(([^)]+)\)/g, (_, 中) => {
      const r = 参.読む(中.trim());
      let s = 0;
      for (let rr = r.上; rr <= Math.min(r.下, 50); rr++) for (let cc = r.左; cc <= Math.min(r.右, 50); cc++) s += 読み値('Sheet1!' + rr + ',' + cc);
      return String(s);
    });
    /* マス を 値に 置き換える */
    f = f.replace(/\$?[A-Z]+\$?\d+/g, (a) => { const r = 参.読む(a); return String(読み値('Sheet1!' + r.上 + ',' + r.左)); });
    /* + と * だけ（測り用） */
    let v = 0;
    for (const 足 of f.split('+')) { let m = 1; for (const 掛 of 足.split('*')) m *= Number(掛); v += m; }
    値[札] = v;
    void rc;
  }
  return 値;
}

T('★★その 順に 1回 回したら 実Excel と 同じ 答えに なる（' + Object.keys(実Excel).length + 'マス）★★', () => {
  const 値 = ちいさく計算(地図.順番, 地図.頼り);
  const 違い = [];
  for (const a in 実Excel) {
    const g = 値[札に(a)];
    if (g !== 実Excel[a]) 違い.push(a + ' … うち ' + g + ' ／ ★実Excel ' + 実Excel[a] + '★');
  }
  if (違い.length) throw new Error('★' + 違い.length + 'マス 違う★  ' + 違い.slice(0, 4).join(' ／ '));
});

T('★★列 まるごと（B:B）でも 104万行を 1本ずつ 数えない★★', () => {
  /* ★中身が 合って いるか★＝B の 中に 在る 式だけ 結ぶ（数の マスは 結ばない） */
  const m0 = C.地図を作る(板を作る({ 'A1': '=SUM(B:B)', 'B1': '=1', 'B9999': '=2' }));
  const 頼 = m0.頼り['Sheet1!0,0'] || [];
  if (頼.length !== 2) throw new Error('★B:B の 中の 式が ' + 頼.length + '本★（2本 のはず）');

  /* ★速さ★＝104万行を 1本ずつ 見て いたら ここで 止まります
     ★まるごとの 参照を 200本★ 並べる（200 × 104万 = 2億回）
     ★行を 二分探索で 挟む 作りなら 一瞬★ */
  const 表 = { 'B1': '=1', 'B9999': '=2' };
  for (let i = 1; i <= 200; i++) 表['A' + i] = '=SUM(B:B)';
  const t0 = Date.now();
  const m = C.地図を作る(板を作る(表));
  const かかった = Date.now() - t0;
  if (m.順番.length !== 202) throw new Error('順番 ' + m.順番.length + ' / 202');
  if (かかった > 500) throw new Error('★' + かかった + 'ms かかった★＝行を 1本ずつ 数えて いる');
});

T('★★輪（ぐるぐる）は 名指しで 返す（勝手に 直さない）★★', () => {
  /* ★実Excel も Iteration=False で 正しい 答えを 出さず 断ります★ */
  const m = C.地図を作る(板を作る({ 'A1': '=A2+1', 'A2': '=A1+1', 'B1': '=B1+1', 'C1': '=5' }));
  const 輪 = m.輪.slice().sort();
  const 欲 = [札に('A1'), 札に('A2'), 札に('B1')].sort();
  if (輪.join('|') !== 欲.join('|')) throw new Error('★輪 = ' + JSON.stringify(m.輪) + '★（A1・A2・B1 のはず）');
  if (m.順番.indexOf(札に('A1')) >= 0) throw new Error('★輪の マスを 順番に 入れて いる★');
  if (m.順番.indexOf(札に('C1')) < 0) throw new Error('★輪と 関係 無い C1 まで 落とした★');
});

T('★★1マス 変えた時、要る 分だけ 計算し直す★★', () => {
  const 出 = C.変えたら(地図, [札に('A3')]);
  /* A3 → A2 → A1 → B2 → C1（D・E・F は 関係 無い） */
  const 欲 = [札に('A2'), 札に('A1'), 札に('B2'), 札に('C1')];
  if (出.length !== 欲.length || 欲.some((x) => 出.indexOf(x) < 0)) {
    throw new Error('★' + 出.length + '本 返った★ ' + JSON.stringify(出));
  }
  /* ★返る 順も 頼りの 順で ないと 古い 答えが 出る★ */
  if (出.indexOf(札に('A2')) > 出.indexOf(札に('A1'))) throw new Error('★順が 逆★');
  if (出.indexOf(札に('B2')) > 出.indexOf(札に('C1'))) throw new Error('★順が 逆（B2 と C1）★');
});

T('★★別の 板を 見る 頼りも 結ぶ★★', () => {
  const m = C.地図を作る([
    { 名: 'Sheet1', 式: { '0,0': '=Sheet2!A1+1' } },
    { 名: 'Sheet2', 式: { '0,0': '=2' } }
  ]);
  if ((m.頼り['Sheet1!0,0'] || []).join() !== 'Sheet2!0,0') {
    throw new Error('★' + JSON.stringify(m.頼り['Sheet1!0,0']) + '★（Sheet2!0,0 のはず）');
  }
  if (m.順番[0] !== 'Sheet2!0,0') throw new Error('★見られる 方を 先に して いない★');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
