# -*- coding: utf-8 -*-
import io, re
from fractions import Fraction
P = 'C:/Users/zeroa/exally-prod/docs/measured/golden-oddl-6kaime-2026-09-18.tsv'
行 = [l.split(u'\t') for l in io.open(P, encoding='utf-8').read().split(u'\n')
      if l and not l.startswith(u'#')]
KUMI = [u'\u7d441', u'\u7d442', u'\u7d443', u'\u7d444', u'\u7d445', u'\u7d446']
FREQ = dict(zip(KUMI, [4, 2, 2, 4, 1, 2]))
HAMI = dict(zip(KUMI, [3, 61, 1, 48, 28, 1]))
NCC  = dict(zip(KUMI, [3, 1, 1, 2, 1, 3]))
A_, I_ = u'(\u30a2)', u'(\u30a4)'
表 = {}
for c in 行:
    if not c[0].startswith(u'\u7d44'): continue
    m = re.search(u'basis=([0-9])', c[0])
    if not m: continue
    種 = A_ if A_ in c[0] else (I_ if I_ in c[0] else None)
    if 種: 表.setdefault((c[0][:2], int(m.group(1))), {})[種] = c[2]
out = []
out.append(u'  \u7d44   f  \u306f\u307f  NC  b   DSC(\u5229\u73870)      DC-A(\u5229\u56de\u308a0)    \u5dee            \u5206\u6570')
for k in KUMI:
    f = FREQ[k]; 券 = 100.0 * 0.06 / f
    for b in range(5):
        g = 表[(k, b)]
        pa, pi = float(g[A_]), float(g[I_])
        DSC = (100.0/pa - 1.0) * f / 0.05
        DCmA = (pi - 100.0) / 券
        d = DCmA - DSC
        if abs(d) < 1e-12:
            out.append(u'  %s %2d %4d %3d %2d  %13.9f %13.9f   \u2605\u540c\u3058\u2605' % (k, f, HAMI[k], NCC[k], b, DSC, DCmA))
        else:
            F = Fraction(d).limit_denominator(400)
            out.append(u'  %s %2d %4d %3d %2d  %13.9f %13.9f  %.9f  \u2605%s\u2605'
                       % (k, f, HAMI[k], NCC[k], b, DSC, DCmA, d, F))
    out.append(u'')
s = u'\n'.join(out)
io.open('C:/Users/zeroa/AppData/Local/Temp/claude/C--WINDOWS-System32-WindowsPowerShell-v1-0/99bd8909-7111-4a26-b739-c2029ed344d1/scratchpad/hyou.txt',
        'w', encoding='utf-8', newline='').write(s + u'\n')
print(s)
