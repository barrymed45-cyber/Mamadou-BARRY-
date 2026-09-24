import json, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

R = json.load(open('results.json')); b = R['base']
NAVY, BLUE, GOLD, TEAL, CORAL, GRAY, INK, MUTED = '#0B2545', '#1F4E9E', '#C8962E', '#2A9D8F', '#D9674E', '#8A94A6', '#1F2937', '#6B7280'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.edgecolor': '#CBD2DC', 'axes.linewidth': 0.8,
                     'axes.titleweight': 'bold', 'axes.titlesize': 12, 'axes.titlecolor': NAVY, 'axes.labelcolor': MUTED,
                     'xtick.color': MUTED, 'ytick.color': MUTED, 'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.grid': True, 'grid.color': '#E8ECF2', 'grid.linewidth': 0.8, 'axes.axisbelow': True,
                     'legend.frameon': False, 'savefig.dpi': 200})
fr = lambda x, d=1: f'{x:,.{d}f}'.replace(',', ' ').replace('.', ',')
Y = [f'A{i}' for i in range(1, 8)]

def save(fig, n):
    fig.tight_layout(); fig.savefig(f'fig{n}.png', facecolor='white'); plt.close(fig)

# 1. Production / exportation
fig, ax = plt.subplots(figsize=(8, 3.6))
cats = ['Production', 'Exportation']; v24 = [141.7, 145.9]; v25 = [175.45, 182.83]
x = [0, 1]; w = 0.34
ax.bar([i - w / 2 - 0.01 for i in x], v24, w, color=GRAY, label='2024')
ax.bar([i + w / 2 + 0.01 for i in x], v25, w, color=BLUE, label='2025')
for i in x:
    ax.text(i - w / 2, v24[i] + 3, fr(v24[i]), ha='center', color=INK)
    ax.text(i + w / 2, v25[i] + 3, fr(v25[i], 2), ha='center', color=NAVY, weight='bold')
    ax.text(i + w / 2, v25[i] / 2, f'+{round((v25[i] / v24[i] - 1) * 100)} %', ha='center', color='white', weight='bold')
ax.set_xticks(x, cats); ax.set_ylabel('Millions de tonnes'); ax.set_ylim(0, 210); ax.grid(axis='x', visible=False)
ax.legend(loc='upper left', ncol=2); ax.set_title('Bauxite guinéenne : production et exportations (Mt)', loc='left')
save(fig, 1)

# 2. Exportateurs
fig, ax = plt.subplots(figsize=(8, 3.8))
names = ['SMB', 'CHALCO', 'CBG', 'AGB2A SDM', 'CDM Chine', 'AMC', 'COBAD', 'AGB2A GIC'][::-1]
vals = [71.52, 22.12, 17.37, 17.01, 9.4, 6.8, 6.2, 5.6][::-1]
ax.barh(names, vals, color=[BLUE] * 7 + [NAVY], height=0.66)
for n, v in zip(names, vals): ax.text(v + 0.8, n, fr(v), va='center', color=INK, fontsize=9)
ax.set_xlabel('Millions de tonnes exportées en 2025'); ax.grid(axis='y', visible=False); ax.set_xlim(0, 80)
ax.set_title('Principaux exportateurs de bauxite en 2025 (Mt)', loc='left'); save(fig, 2)

# 3. Saisonnalité
fig, ax = plt.subplots(figsize=(8, 3.4))
m = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc']
prod = [17.42, 16.15, 17.44, 15.54, 16.07, 13.68, 11.30, 11.64, 11.77, 13.33, 15.03, 16.08]
exp = [16.71, 14.52, 17.38, 16.32, 18.61, 15.24, 12.55, 11.61, 13.04, 14.20, 15.27, 17.35]
ax.axvspan(4.5, 9.5, color='#EEF3FB', zorder=0); ax.text(7, 18.6, 'Hivernage', ha='center', color=BLUE, fontsize=9, style='italic')
ax.plot(m, prod, color=BLUE, lw=2, marker='o', ms=5, label='Production')
ax.plot(m, exp, color=GOLD, lw=2, marker='o', ms=5, label='Exportations')
ax.set_ylabel('Millions de tonnes'); ax.set_ylim(10, 19.5); ax.legend(loc='lower left', ncol=2)
ax.set_title('Saisonnalité mensuelle de la bauxite en 2025 (Mt)', loc='left'); save(fig, 3)

# 4. CA / EBITDA
fig, ax = plt.subplots(figsize=(8, 3.6))
x = range(7); w = 0.38
ax.bar([i - w / 2 - 0.01 for i in x], b['ca'], w, color=BLUE, label="Chiffre d'affaires")
ax.bar([i + w / 2 + 0.01 for i in x], b['ebitda'], w, color=GOLD, label='EBITDA')
for i in x:
    ax.text(i - w / 2, b['ca'][i] + 0.5, fr(b['ca'][i]), ha='center', fontsize=8, color=INK)
    ax.text(i + w / 2, b['ebitda'][i] + 0.5, f"{round(b['ebitda_marge'][i] * 100)} %", ha='center', fontsize=8, color=MUTED)
ax.set_xticks(list(x), Y); ax.set_ylabel('M USD'); ax.set_ylim(0, 36); ax.grid(axis='x', visible=False)
ax.legend(loc='upper left', ncol=2); ax.set_title("Chiffre d'affaires et EBITDA (M USD, marge EBITDA en %)", loc='left'); save(fig, 4)

# 5. Structure des coûts A2
v2 = b['vol'][1]; au = b['autres'][1]
parts = [('Carburant', b['carb'][1], GOLD), ('Maintenance et pièces', au * 0.82 / 1.46, BLUE),
         ('Pneumatiques', au * 0.43 / 1.46, TEAL), ('Lubrifiants et divers', au * 0.21 / 1.46, CORAL),
         ('Personnel et coûts fixes', b['fixes'][1], GRAY)]
fig, ax = plt.subplots(figsize=(8, 3.6))
tot = sum(p[1] for p in parts)
wedges, _ = ax.pie([p[1] for p in parts], colors=[p[2] for p in parts], startangle=90, counterclock=False,
                   wedgeprops=dict(width=0.38, edgecolor='white', linewidth=2))
ax.text(0, 0.08, fr(tot), ha='center', fontsize=16, weight='bold', color=NAVY); ax.text(0, -0.18, 'M USD', ha='center', color=MUTED)
ax.legend(wedges, [f"{p[0]} — {fr(p[1], 2)} M$ ({round(p[1] / tot * 100)} %)" for p in parts], loc='center left', bbox_to_anchor=(1.0, 0.5))
ax.set_title("Structure des coûts d'exploitation — Année 2", loc='left'); ax.axis('equal'); save(fig, 5)

# 6. Dette et DSCR (deux panneaux, pas de double axe)
fig, (a1, a2) = plt.subplots(1, 2, figsize=(8.4, 3.4))
a1.bar(['A0'] + Y, [b['debt']] + b['dette_fin'], color=BLUE, width=0.62)
for i, v in enumerate([b['debt']] + b['dette_fin']):
    if v > 0.01: a1.text(i, v + 0.4, fr(v), ha='center', fontsize=8, color=INK)
a1.set_title('Encours de dette fin de période (M USD)', loc='left', fontsize=10.5); a1.grid(axis='x', visible=False)
d = [x for x in b['dscr'] if x is not None]
a2.axhline(1.30, color=CORAL, lw=1.2, ls='--'); a2.text(5.4, 1.33, 'Seuil bancaire 1,30x', ha='right', fontsize=8, color=CORAL)
a2.axhline(1.10, color=MUTED, lw=1, ls=':'); a2.text(5.4, 1.12, 'Cas de défaut 1,10x', ha='right', fontsize=8, color=MUTED)
a2.plot(Y[:6], d, color=NAVY, lw=2, marker='o', ms=6)
for i, v in enumerate(d): a2.text(i, v + 0.05, f'{fr(v, 2)}x', ha='center', fontsize=8, color=INK)
a2.set_ylim(0.9, 2.0); a2.set_title('DSCR annuel (cas de base)', loc='left', fontsize=10.5)
save(fig, 6)

# 7. Tornado TRI
base_tri = R['scen'][0][1]['tri'] * 100
items = [(n, r['tri'] * 100) for n, r in R['scen'][1:8]]
items.sort(key=lambda t: t[1], reverse=True)
fig, ax = plt.subplots(figsize=(8, 3.4))
for i, (n, v) in enumerate(items):
    ax.barh(i, v - base_tri, left=base_tri, color=CORAL if v < base_tri else TEAL, height=0.6)
    ax.text(v - 0.5, i, f'{fr(v)} %', va='center', ha='right', fontsize=8.5, color=INK)
ax.axvline(base_tri, color=NAVY, lw=1.4); ax.text(base_tri + 0.3, len(items) - 0.4, f'Base {fr(base_tri)} %', color=NAVY, fontsize=9, weight='bold')
ax.set_yticks(range(len(items)), [n for n, _ in items]); ax.set_xlim(8, 33); ax.grid(axis='y', visible=False); ax.invert_yaxis()
ax.set_xlabel('TRI fonds propres (%)'); ax.set_title('Sensibilité du TRI fonds propres aux chocs unitaires', loc='left'); save(fig, 7)

# 8. Gantt
tasks = [('Étude de route et chronométrage', 0, 1, 'Préparation'), ('Contrat commercial (term sheet → signature)', 1, 3, 'Préparation'),
         ('Due diligence prêteurs et bouclage financier', 2, 5, 'Financement'), ('Commande et fabrication des camions', 4, 8, 'Mobilisation'),
         ('Construction atelier / dépôt', 4, 8, 'Mobilisation'), ('Recrutement et formation', 6, 10, 'Mobilisation'),
         ('Livraison, dédouanement, homologation', 8, 10, 'Mobilisation'), ('Pilote 20 camions', 10, 11, 'Démarrage'),
         ('Montée en cadence à 100 camions', 11, 12, 'Démarrage')]
col = {'Préparation': GRAY, 'Financement': GOLD, 'Mobilisation': BLUE, 'Démarrage': TEAL}
fig, ax = plt.subplots(figsize=(8.4, 3.8))
for i, (n, s, e, c) in enumerate(tasks):
    ax.barh(i, e - s, left=s, color=col[c], height=0.6)
ax.set_yticks(range(len(tasks)), [t[0] for t in tasks], fontsize=9); ax.invert_yaxis()
ax.set_xticks(range(0, 13)); ax.set_xlim(0, 12.2); ax.set_xlabel('Mois après décision d’investissement (Année 0)')
ax.axvline(5, color=GOLD, lw=1, ls='--'); ax.text(5.1, 8.6, 'Bouclage financier', color='#8A6516', fontsize=8)
ax.axvline(12, color=TEAL, lw=1, ls='--'); ax.text(11.9, 8.6, 'Mise en service', color=TEAL, fontsize=8, ha='right')
from matplotlib.patches import Patch
ax.legend([Patch(color=v) for v in col.values()], col.keys(), loc='upper right', ncol=4, fontsize=8, bbox_to_anchor=(1, 1.13))
ax.grid(axis='y', visible=False); ax.set_title('Calendrier de mobilisation — Année 0', loc='left', pad=22); save(fig, 8)
print('ok')
