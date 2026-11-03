# -*- coding: utf-8 -*-
# ★溜まって いる 問いを 1枠に まとめる★（★紙から 読む＝手で 書き写さない★）
#
# ★★2026-09-18 に 1度 落としました★★
#   ★前★ … 行の 先頭が `=` の 物だけ 拾って いた
#   ⇒★★1行に 2つ 在ると 2つ目を 取りこぼす★★
#     （`=PERMUT(0,0)        =PERMUTATIONA(0,0)` で ★6本 落ちた★）
#   ⇒★★見分け方★★ … 紙に `PERMUTATIONA` が 10回 出て くるのに
#                    集めた 中に ★0本★だった
#   ⇒★直し★ … ★行の 中の `=NAME(...)` を 全部 拾う★
#   ⇒★記憶の 型★「★1行に 2つ 在ると 取りこぼす★」（§5 の div と 同じ）
import io, os, re
D = os.path.join(os.path.dirname(os.path.abspath(__file__)))
紙たち = ['junretsu-kiku-koto.md', 'xmatch-kiku-koto.md', 'percentrank-kiku-koto.md',
          'asc-dbcs-kiku-koto.md', 'lenb-kiku-koto.md', 'textafter-kiku-koto.md',
          'aggregate-kiku-koto.md']
# ★=NAME( から 釣り合う ) まで★
PAT = re.compile(u'=[A-Z][A-Z0-9._]*\(')

def 式たちを拾う(l):
    出 = []
    for m in PAT.finditer(l):
        i = m.end() - 1
        深 = 0
        引用 = False
        j = i
        while j < len(l):
            c = l[j]
            if c == u'"':
                引用 = not 引用
            elif not 引用:
                if c == u'(':
                    深 += 1
                elif c == u')':
                    深 -= 1
                    if 深 == 0:
                        出.append(l[m.start():j + 1])
                        break
            j += 1
    return 出

出 = []
for f in 紙たち:
    p = os.path.join(D, f)
    if not os.path.exists(p):
        print('NG missing ' + f)
        continue
    t = io.open(p, encoding='utf-8').read()
    i = t.find(u'次の 枠で 聞く 式')
    # ★★節の 終わりで 止める★★（2026-09-18）
    #   ★前は 紙の 最後まで 読んで いました★
    #   ⇒★④「お客さんに 何が 変わるか」の ★既に 測った 例★を 拾って いました
    #   ⇒★見分け方★ … ★各紙が ちょうど 1本ずつ 多かった★
    if i >= 0:
        j = t.find(chr(10) + u'## ', i)
## ', i)
        本文 = t[i:j] if j > i else t[i:]
    else:
        本文 = t
    式 = []
    for l in 本文.split(u'\n'):
        for x in 式たちを拾う(l):
            if x not in 式:
                式.append(x)
    出.append((f, 式))
    print(u'%-34s %d' % (f, len(式)))

合 = sum(len(x[1]) for x in 出)
print(u'total %d' % 合)
io.open(os.path.join(D, 'toi-atsumeta.txt'), 'w', encoding='utf-8', newline='').write(
    u'\n'.join([f + u'\t' + x for f, 式 in 出 for x in 式]) + u'\n')
