#!/bin/sh
# pre-push-noreply.sh -- ★押す 前に、まだ 遠くに 無い commit の メールが noreply か 見る★（2026-10-11）
#   なぜ：公開 repo の commit の 欄に 受信箱や 司さんの メールが 出て いた（司さん 10-11「アドレスが分からんようにしろや」）。
#         repo の 設定（--local）の user.email は GitHub の noreply に した。設定が 抜けた 席で 押すと また 出るので 止める。
#   見る：作者・commit した 人・本文の Co-authored-by・注釈付きの タグの tagger。
#   通す：<数字>+<口の名>@users.noreply.github.com（bot の [bot] も）・noreply@github.com（GitHub 自身）・
#         noreply@anthropic.com（Claude の 共作の 印＝人の 受信箱で ない）だけ。
#   「遠く」は 押す 先の remote（$1）の 追跡の 枝だけ（ほかの remote に 在る 物は 見る＝10-11 対立役）。
#   出す：SHA と 欄の 名前だけ（メールの 字は 出さない）。
#   置き方：cp scripts/pre-push-noreply.sh .git/hooks/pre-push（git と sh だけで 動く＝古い 枝の 作業場所でも 同じ）
#   読む物：pre-push の 引数 $1＝remote の 名・標準入力「<手元の ref> <手元の sha> <遠くの ref> <遠くの sha>」
remote=${1:-origin}
zero=0000000000000000000000000000000000000000
tab=$(printf '\t')
bad=0
seen=0
ok_mail() {
  m=$(printf '%s' "$1" | tr 'A-Z' 'a-z')
  printf '%s\n' "$m" | grep -Eq '^[0-9]+\+[a-z0-9-]+(\[bot\])?@users\.noreply\.github\.com$' && return 0
  [ "$m" = "noreply@github.com" ] && return 0
  [ "$m" = "noreply@anthropic.com" ] && return 0
  return 1
}
note_bad() { bad=$((bad + 1)); echo "  $(printf '%s' "$1" | cut -c1-12) の $2" >&2; }
while read lref lsha rref rsha; do
  [ -z "$lsha" ] && continue
  [ "$lsha" = "$zero" ] && continue          # 消す 押しは 見る commit 無し
  # 注釈付きの タグ＝tagger の 欄を 見る
  if [ "$(git cat-file -t "$lsha" 2>/dev/null)" = "tag" ]; then
    tm=$(git cat-file tag "$lsha" | sed -n 's/^tagger .*<\([^>]*\)>.*/\1/p')
    ok_mail "$tm" || note_bad "$lsha" 'タグの tagger'
  fi
  list=$(git log --format="%H${tab}%ae${tab}%ce" "$lsha" --not --remotes="$remote" 2>/dev/null) || { echo "★作者の 門：commit を 並べられない＝押さない★" >&2; exit 1; }
  old_ifs=$IFS; IFS='
'
  for line in $list; do
    IFS=$tab
    set -- $line
    IFS='
'
    seen=$((seen + 1))
    ok_mail "$2" || note_bad "$1" '作者'
    ok_mail "$3" || note_bad "$1" 'commit した 人'
    for co in $(git log -1 --format='%(trailers:key=Co-authored-by,valueonly)' "$1" | sed -n 's/.*<\([^>]*\)>.*/\1/p'); do
      ok_mail "$co" || note_bad "$1" '本文の Co-authored-by'
    done
  done
  IFS=$old_ifs
done
if [ "$bad" -gt 0 ]; then
  echo "★押すのを 止めた★ メールが noreply で ない 欄が $bad 件（字は 出さない）" >&2
  echo "  ⇒ git config --local user.email <id>+<口>@users.noreply.github.com にして commit を 作り直す（--amend --reset-author）" >&2
  exit 1
fi
echo "pre-push-noreply: 見た commit $seen 本・全部 noreply"
exit 0
