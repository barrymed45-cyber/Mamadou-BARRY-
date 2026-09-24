"""Modèle financier révisé - Transport minier 100 camions (USD millions).
Année 0 = mobilisation (12 mois), Années 1-7 = exploitation."""
import json, copy

BASE = dict(
    tarif0=7.10, tarif_idx=0.02,
    vol=[2.75, 3.55, 3.85, 3.90, 3.90, 3.90, 3.90],
    diesel_usd_l=1.37, conso=1.05, diesel_esc=0.015,
    autres_var_t=1.46, autres_var_esc=0.03,   # pneus, maintenance, lubrifiants ($/t)
    fixes0=4.80, fixes_esc=0.03,
    capex_hard=24.20, contingence=2.80, fr_initial=3.00,
    capex_maintien=[0.3, 0.5, 0.8, 0.8, 1.0, 1.0, 1.2],
    bfr_pct_ca=0.12,
    valeur_residuelle=5.00,
    is_rate=0.25,
    taux=0.105, frais_montage=0.02, frais_conseils=0.30, idc_util=0.30,
    part_dette=0.65, grace_annees_ops=0, n_remb=6, sculpt=True,
    dsra_mois=6,
    ke=0.18, kd_proj=0.14,
    capex_mult=1.0, tarif_mult=1.0, vol_mult=1.0, diesel_mult=1.0, fixes_mult=1.0,
    retard_mois=0,
)

N = 7

def run(p=None, **over):
    p = copy.deepcopy(p or BASE); p.update(over)
    if not p.get('sculpt'):
        return _run(p)
    p['_sched'] = None
    for _ in range(40):
        o = _run(p)
        # nouvelle forme de remboursement : principal = CFADS/T - intérêts
        g, n = p['grace_annees_ops'], p['n_remb']
        lo, hi = 0.8, 4.0
        for _ in range(60):
            T = (lo + hi) / 2; d = o['debt']; sch = [0.0] * N
            for i in range(g, g + n):
                pr = o['cfads'][i] / T - d * p['taux']; sch[i] = pr; d -= pr
            if d > 0: hi = T
            else: lo = T
        p['_sched'] = [max(x, 0.0) for x in sch]
        p['_ds1'] = o['service'][0]
    o = _run(p); o['dscr_cible'] = T
    return o

def _run(p):
    hard = p['capex_hard'] * p['capex_mult']
    cont = p['contingence'] * p['capex_mult']
    fr0 = p['fr_initial']
    # Sources & emplois (résolution de la circularité frais/IDC/DSRA)
    debt = 20.0
    for _ in range(100):
        fees = debt * p['frais_montage'] + p['frais_conseils']
        idc = debt * p['taux'] * p['idc_util']
        dsra0 = p.get('_ds1', debt * p['taux']) * p['dsra_mois'] / 12   # 6 mois du service A1
        total = hard + cont + fr0 + fees + idc + dsra0
        debt = total * p['part_dette']
    equity = total - debt

    tarif = [p['tarif0'] * p['tarif_mult'] * (1 + p['tarif_idx']) ** i for i in range(N)]
    vol = [v * p['vol_mult'] for v in p['vol']]
    # retard de démarrage : perte de x mois de volume en A1
    if p['retard_mois']:
        vol[0] = vol[0] * (12 - p['retard_mois']) / 12
    ca = [tarif[i] * vol[i] for i in range(N)]
    carb_t = [p['diesel_usd_l'] * p['diesel_mult'] * p['conso'] * (1 + p['diesel_esc']) ** i for i in range(N)]
    autres_t = [p['autres_var_t'] * (1 + p['autres_var_esc']) ** i for i in range(N)]
    carb = [carb_t[i] * vol[i] for i in range(N)]
    autres = [autres_t[i] * vol[i] for i in range(N)]
    fixes = [p['fixes0'] * p['fixes_mult'] * (1 + p['fixes_esc']) ** i for i in range(N)]
    opex = [carb[i] + autres[i] + fixes[i] for i in range(N)]
    ebitda = [ca[i] - opex[i] for i in range(N)]

    # Amortissements (IAS 16 : base amortissable = coût - valeur résiduelle)
    base_amort = hard + cont + fees + idc - p['valeur_residuelle']
    dep = [base_amort / N] * N
    cm = p['capex_maintien']
    for i in range(N):  # CAPEX de maintien amorti sur la durée restante
        for j in range(i, N):
            dep[j] += cm[i] / (N - i)

    bfr = [p['bfr_pct_ca'] * c for c in ca]
    dbfr = [bfr[0]] + [bfr[i] - bfr[i - 1] for i in range(1, N)]
    dbfr[-1] -= bfr[-1]  # récupération du BFR en fin d'horizon

    # Dette : profil sculpté (DSCR cible constant) après la période de mobilisation
    g = p['grace_annees_ops']; nrem = p['n_remb']
    sched = p.get('_sched')
    if sched is None:
        sched = [0.0] * N
        for i in range(g, g + nrem): sched[i] = debt / nrem
    princ = [x * debt / sum(sched) for x in sched]
    dette_deb, dette_fin, interets = [], [], []
    d = debt
    for i in range(N):
        dette_deb.append(d); interets.append(d * p['taux']); d -= princ[i]; dette_fin.append(d)
    service = [interets[i] + princ[i] for i in range(N)]

    ebit = [ebitda[i] - dep[i] for i in range(N)]
    ebt = [ebit[i] - interets[i] for i in range(N)]
    impot, report = [], 0.0
    for i in range(N):
        base = ebt[i] + report
        if base < 0:
            report = base; impot.append(0.0)
        else:
            report = 0.0; impot.append(base * p['is_rate'])
    rn = [ebt[i] - impot[i] for i in range(N)]

    cfads = [ebitda[i] - impot[i] - cm[i] - dbfr[i] for i in range(N)]
    cfads[-1] += p['valeur_residuelle']
    dscr = [cfads[i] / service[i] if service[i] > 1e-9 else None for i in range(N)]
    # LLCR : VA des CFADS sur la durée du prêt / dette début
    llcr = []
    for i in range(N):
        if dette_deb[i] < 1e-9: llcr.append(None); continue
        last = max(k for k in range(N) if princ[k] > 0 or k < g)
        pv = sum(cfads[k] / (1 + p['taux']) ** (k - i + 1) for k in range(i, g + p['n_remb']))
        llcr.append(pv / dette_deb[i])

    # DSRA : 6 mois du service de la dette de l'année suivante
    dsra_bal, dsra_mvt = [], []
    prev = dsra0
    for i in range(N):
        nxt = service[i + 1] * p['dsra_mois'] / 12 if i + 1 < N else 0.0
        dsra_mvt.append(nxt - prev); dsra_bal.append(nxt); prev = nxt
    # trésorerie : fonds de roulement initial = trésorerie d'ouverture
    cash = fr0
    flux_act, tres = [], []
    for i in range(N):
        disp = cfads[i] - service[i] - dsra_mvt[i]
        cash += disp
        # distribution si DSCR >= 1,20x et trésorerie minimale 1,0 M$ (dernière année : tout)
        lock = dscr[i] is not None and dscr[i] < 1.20
        mini = 0.0 if i == N - 1 else 1.0
        dist = max(0.0, cash - mini) if not lock else 0.0
        cash -= dist
        flux_act.append(dist); tres.append(cash)
    eq_flows = [-equity] + flux_act
    proj_flows = [-(hard + cont + fr0)] + [ebitda[i] - cm[i] - dbfr[i] - ebit[i] * p['is_rate'] * (ebit[i] > 0) + (p['valeur_residuelle'] if i == N - 1 else 0) for i in range(N)]
    # le fonds de roulement initial reste en trésorerie : restitué en fin d'horizon
    proj_flows[-1] += fr0
    tri_fp = irr(eq_flows); tri_proj = irr(proj_flows)
    van_fp = npv(p['ke'], eq_flows)
    cum, payback = 0, None
    for i, f in enumerate(eq_flows):
        prev_cum = cum; cum += f
        if cum >= 0 and payback is None and i > 0:
            payback = i - 1 + (-prev_cum) / f
    ds = [x for x in dscr if x is not None]
    out = dict(debt=debt, equity=equity, total=total, fees=fees, idc=idc, dsra0=dsra0, hard=hard, cont=cont, fr0=fr0,
               tarif=tarif, vol=vol, ca=ca, carb=carb, autres=autres, fixes=fixes, opex=opex, ebitda=ebitda,
               dep=dep, ebit=ebit, interets=interets, ebt=ebt, impot=impot, rn=rn, princ=princ,
               dette_deb=dette_deb, dette_fin=dette_fin, service=service, cfads=cfads, dscr=dscr, llcr=llcr,
               bfr=bfr, dbfr=dbfr, cm=cm, dsra_bal=dsra_bal, dsra_mvt=dsra_mvt, flux_act=flux_act, tres=tres,
               eq_flows=eq_flows, proj_flows=proj_flows, tri_fp=tri_fp, tri_proj=tri_proj, van_fp=van_fp,
               payback=payback, dscr_min=min(ds), dscr_moy=sum(ds[1:]) / len(ds[1:]), llcr_min=min(x for x in llcr if x),
               dscr_min_amort=min(ds[1:]), ebitda_marge=[ebitda[i] / ca[i] for i in range(N)])
    return out

def npv(r, fl):
    return sum(f / (1 + r) ** i for i, f in enumerate(fl))

def irr(fl):
    lo, hi = -0.99, 5.0
    if npv(hi, fl) > 0: return None
    for _ in range(200):
        m = (lo + hi) / 2
        if npv(m, fl) > 0: lo = m
        else: hi = m
    return m

def solve(key, target, lo, hi, metric='dscr_min_amort', **over):
    for _ in range(80):
        m = (lo + hi) / 2
        v = run(**{key: m}, **over)[metric]
        if v < target: lo = m
        else: hi = m
    return m

if __name__ == '__main__':
    b = run()
    f = lambda xs: ' '.join(f'{x:7.3f}' if x is not None else '   -   ' for x in xs)
    print('Total %.2f dette %.2f FP %.2f fees %.2f idc %.2f dsra %.2f' % (b['total'], b['debt'], b['equity'], b['fees'], b['idc'], b['dsra0']))
    for k in ['ca', 'carb', 'autres', 'fixes', 'opex', 'ebitda', 'dep', 'ebit', 'interets', 'impot', 'rn', 'dbfr', 'cfads', 'service', 'dscr', 'llcr', 'dsra_mvt', 'flux_act', 'tres', 'dette_fin']:
        print(f'{k:9s}', f(b[k]))
    print('cible %.3f' % b['dscr_cible']); print('TRI FP %.1f%%  TRI projet %.1f%%  VAN18 %.2f  payback %.2f  DSCRmin %.2f moy %.2f LLCRmin %.2f' % (
        b['tri_fp'] * 100, b['tri_proj'] * 100, b['van_fp'], b['payback'], b['dscr_min_amort'], b['dscr_moy'], b['llcr_min']))
    print('eq flows', f(b['eq_flows']))
