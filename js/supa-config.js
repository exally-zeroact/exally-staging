/* supa-config.js — ★テスト用DB★ (exally-staging / GitHub Pages配信専用)
 * 本番倉庫(tnfwipbgfgjaymlszeid)とは別の Supabase「DB-test」を指す。
 * URLとpublishable(公開鍵)はクライアント埋め込みで安全＝RLSで本人ぶんだけ保護。
 * ★このファイルは本番(exally)には絶対にコピーしない(本番は本番倉庫を指す)。
 * ★payslip-app-test と同じ DB-test を共有＝1テストアカウントで両方使える(本番と同じ倉庫共有構成)。
 */
/* ★env = この配信がどの環境か（'test' | 'prod'）★
 *   画面の一番上に「テスト環境」の帯を出すかを、これ1つで決める（js/env-badge.js）。
 *   ★向き先を持っているのはこのファイルだけ★という約束（tests/no-hardcoded-supa.test.mjs）
 *   に合わせて、環境の名札もここに置く。他のファイルに倉庫の名前を書かせない。
 *   ★本番(exally)の supa-config.js は env:'prod'。だから本番に帯は出ない。★ */
window.SUPA = {
  url: 'https://khawdrnvssdenumbiwfg.supabase.co',
  key: 'sb_publishable_UrRIobyVFbaJI_85RBxBOA_GZ4OUxPm',
  env: 'test'
};
