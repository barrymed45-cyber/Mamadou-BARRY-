import json
from model import run, BASE, N
b = run()
fix = dict(sculpt=False, _sched=b['princ'], _ds1=b['service'][0])
def sc(**kw):
    o = run(**fix, **kw)
    return dict(tri=o['tri_fp'], van=o['van_fp'], dscr=o['dscr_min_amort'], llcr=o['llcr_min'], ebitda2=o['ebitda'][1])
S = [
 ('Cas de base', {}),
 ('Tarif −10 %', dict(tarif_mult=0.9)),
 ('Volume −15 %', dict(vol_mult=0.85)),
 ('Diesel +20 %', dict(diesel_mult=1.2)),
 ('CAPEX +10 %', dict(capex_mult=1.1)),
 ('Coûts fixes +15 %', dict(fixes_mult=1.15)),
 ('Retard de démarrage 6 mois', dict(retard_mois=6)),
 ('Taux d’intérêt +200 pb', dict(taux=0.125)),
 ('Cas bas prêteur (combiné modéré)', dict(tarif_mult=0.95, vol_mult=0.90, diesel_mult=1.15, capex_mult=1.05)),
 ('Stress combiné sévère', dict(tarif_mult=0.9, vol_mult=0.85, diesel_mult=1.2, capex_mult=1.1)),
]
res = [(n, sc(**k)) for n, k in S]
for n, r in res: print(f"{n:38s} TRI {r['tri']*100 if r['tri'] is not None else float('nan'):6.1f}  VAN {r['van']:6.2f}  DSCR {r['dscr']:.2f} LLCR {r['llcr']:.2f} EBITDA2 {r['ebitda2']:.2f}")
def solve(key, target, lo, hi, extra={}):
    f = lambda m: run(**fix, **extra, **{key: m})['dscr_min_amort']
    inc = f(hi) > f(lo)
    for _ in range(60):
        m = (lo+hi)/2
        if (f(m) < target) == inc: lo = m
        else: hi = m
    return m
be = dict(
 tarif130 = solve('tarif_mult', 1.30, 0.5, 1.2)*7.10,
 tarif100 = solve('tarif_mult', 1.00, 0.5, 1.2)*7.10,
 vol130 = solve('vol_mult', 1.30, 0.5, 1.2),
 vol100 = solve('vol_mult', 1.00, 0.4, 1.2),
 diesel100 = solve('diesel_mult', 1.00, 1.0, 3.0),
 diesel130 = solve('diesel_mult', 1.30, 0.8, 3.0),
)
print(be)
# seuil d'exploitation A2
var2 = (b['carb'][1]+b['autres'][1])/b['vol'][1]
print('seuil exploitation A2', b['fixes'][1]/(b['tarif'][1]-var2))
# TRI sans valeur residuelle
print('TRI sans VR', run(valeur_residuelle=0.0)['tri_fp'])
json.dump(dict(base=b, scen=res, be=be, seuil=b['fixes'][1]/(b['tarif'][1]-var2)), open('results.json','w'), default=float, indent=1)
