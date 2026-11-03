# -*- coding: utf-8 -*-
# ★聞く 前に 形を 確かめる★（★狙った 形に なって いるか★）
import datetime as dt, calendar
D=dt.date
def 月足す(a,m):
    y=a.year+(a.month-1+m)//12; mm=(a.month-1+m)%12+1
    return D(y,mm,min(a.day,calendar.monthrange(y,mm)[1]))
組=[('組1 f=4 はみ出す はず', D(2008,11,30), D(2009,8,31), 4),
    ('組2 f=2 はみ出さない はず', D(2007,10,15), D(2008,6,15), 2),
    ('組3 f=2 ★はみ出す はず★', D(2008,11,30), D(2009,5,31), 2),
    ('組4 f=4 ★はみ出さない はず★', D(2008,11,30), D(2009,7,15), 4),
    ('組5 f=1 ★はみ出す はず★', D(2009,1,31), D(2010,2,28), 1),
    ('組6 f=2 NC大 ★割る為★', D(2008,11,30), D(2010,5,31), 2)]
for 名,最,満,f in 組:
    step=12//f; q=[最]; cur=最
    while cur<満:
        cur=月足す(cur,step); q.append(cur)
    最後の準 = max(x for x in q if x<=満)
    はみ = (満-最後の準).days
    NC = sum(1 for i in range(len(q)-1) if q[i+1]<=満)
    print('%-28s 準 %s' % (名, ' '.join(str(x) for x in q)))
    print('%-28s 満 %s ／ 最後の準 %s ／ ★はみ出し %d日★ ／ NC=%d %s'
          % ('', 満, 最後の準, はみ, NC, '★はみ出す★' if はみ else '★はみ出さない★'))
