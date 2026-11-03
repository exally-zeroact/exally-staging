# -*- coding: utf-8 -*-
# 解けた DC / A / DSC が 何で 出来て いるかを 突き止める
import datetime as dt, calendar
D = dt.date

def 月足す(a, m):
    y = a.year + (a.month-1+m)//12
    mm = (a.month-1+m)%12 + 1
    dd = min(a.day, calendar.monthrange(y,mm)[1])   # ★無ければ 月末へ落とす（落ちたら その日で 続く）★
    return D(y,mm,dd)

def 準の並び(last, mat, f):
    step = 12//f
    out = [last]; cur = last
    while cur < mat:
        cur = 月足す(cur, step)      # ★前の 日から 進める＝落ちた 日が 貼り付く★
        out.append(cur)
    return out

組 = {
 'A': dict(settle=D(2009,1,15), mat=D(2009,11,30), last=D(2008,11,30), f=2),
 'B': dict(settle=D(2009,1,15), mat=D(2009,8,31),  last=D(2008,11,30), f=4),
}
解 = {'A': dict(A=0.2541436464, DC=2.0000000000, DSC=1.7458563536),
      'B': dict(A=0.5111111111, DC=3.0326086956, DSC=2.4888888889)}

for nm in ('A','B'):
    g = 組[nm]; f=g['f']; step=12//f
    q = 準の並び(g['last'], g['mat'], f)
    print('[%s] f=%d  準利払日 %s' % (nm, f, ' '.join(str(x) for x in q)))
    print('      期の 長さ %s' % [ (q[i+1]-q[i]).days for i in range(len(q)-1) ])
    # NC ＝ 満期 まで／満期を 越えない まるごとの 期の 数
    NC = sum(1 for i in range(len(q)-1) if q[i+1] <= g['mat'])
    # A ＝ 最初の 期だけ 割る
    i0 = max(i for i in range(len(q)-1) if q[i] <= g['settle'])
    NL0 = (q[i0+1]-q[i0]).days
    A = (g['settle']-q[i0]).days / float(NL0) + i0
    # はみ出し
    はみ = (g['mat'] - q[NC]).days
    NLはみ = (月足す(q[NC], step) - q[NC]).days
    DC = NC + (はみ/float(NLはみ) if はみ else 0.0)
    DSC = NC - A
    for 名, 出, 実 in (('A',A,解[nm]['A']), ('DC',DC,解[nm]['DC']), ('DSC',DSC,解[nm]['DSC'])):
        print('      %-4s 出た %.10f  解いた %.10f  差 %+.1e %s' %
              (名, 出, 実, 出-実, '★合う★' if abs(出-実)<1e-9 else '★★違う★★'))
    # 値段まで 通す
    rate=0.045; red=100.0; c=100.0*rate/f
    for y,p in {'A':[(0.03,101.26138450528185),(0.05,99.557875048624055),(0.08,97.1068588860573)],
                'B':[(0.03,100.9417062656499),(0.05,99.716504638305835),(0.08,97.933167129919369)]}[nm]:
        got = (red+DC*c)/(1.0+DSC*y/f) - A*c
        print('      値段 yld=%.2f  出た %.12f  実 %.12f  差 %+.1e %s' %
              (y,got,p,got-p,'★合う★' if abs(got-p)<1e-9 else '★★違う★★'))
