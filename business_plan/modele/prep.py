import json
from model import run, N
R = json.load(open('results.json')); b = R['base']
fix = dict(sculpt=False, _sched=b['princ'], _ds1=b['service'][0])
prot = run(**fix, tarif_mult=0.95, vol_mult=0.90, diesel_mult=1.0, capex_mult=1.05)
print('protege', prot['tri_fp'], prot['dscr_min_amort'], prot['van_fp'])
# bilan
gross0 = b['hard'] + b['cont'] + b['fees'] + b['idc']
immo, bfr, dsra, cash, fp, dette = [gross0], [0.0], [b['dsra0']], [b['fr0']], [b['equity']], [b['debt']]
g = gross0; cumdep = 0
for i in range(N):
    g += b['cm'][i]; cumdep += b['dep'][i]
    nbv = g - cumdep
    if i == N - 1: nbv -= 5.0  # cession à la valeur résiduelle
    immo.append(nbv); bfr.append(b['bfr'][i] if i < N - 1 else 0.0); dsra.append(b['dsra_bal'][i]); cash.append(b['tres'][i])
    fp.append(fp[-1] + b['rn'][i] - b['flux_act'][i]); dette.append(b['dette_fin'][i])
for i in range(N + 1):
    a = immo[i] + bfr[i] + dsra[i] + cash[i]; l = fp[i] + dette[i]
    print(i, round(a, 3), round(l, 3), round(a - l, 4))
wacc = 0.35 * 0.18 + 0.65 * 0.105 * 0.75
ratios = []
for i in range(N):
    nd = b['dette_fin'][i] - b['tres'][i] - b['dsra_bal'][i]
    ratios.append(dict(nd_ebitda=nd / b['ebitda'][i], icr=(b['ebitda'][i] / b['interets'][i]) if b['interets'][i] > 1e-6 else None,
                       gearing=b['dette_fin'][i] / (b['dette_fin'][i] + fp[i + 1]) if b['dette_fin'][i] > 1e-6 else 0,
                       roe=b['rn'][i] / fp[i + 1], marge_nette=b['rn'][i] / b['ca'][i]))
for r in ratios: print({k: round(v, 3) if v is not None else None for k, v in r.items()})
print('wacc', wacc)
json.dump(dict(prot=dict(tri=prot['tri_fp'], dscr=prot['dscr_min_amort'], van=prot['van_fp'], llcr=prot['llcr_min']),
               bilan=dict(immo=immo, bfr=bfr, dsra=dsra, cash=cash, fp=fp, dette=dette), ratios=ratios, wacc=wacc),
          open('extra.json', 'w'), default=float)
