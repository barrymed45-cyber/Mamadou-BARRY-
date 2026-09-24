const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Header, Footer, AlignmentType,
  HeadingLevel, WidthType, ShadingType, BorderStyle, PageNumber, PageBreak, LevelFormat, VerticalAlign,
  TableLayoutType, HeightRule,
} = require('docx');

const R = JSON.parse(fs.readFileSync('results.json')); const b = R.base;
const X = JSON.parse(fs.readFileSync('extra.json'));
const NAVY = '0B2545', BLUE = '1F4E9E', GOLD = 'C8962E', INK = '1F2937', MUTED = '6B7280', LIGHT = 'EEF3FB', ZEBRA = 'F6F8FB', LINE = 'D5DCE6', GOLDL = 'FBF4E6', RED = 'B42318', GREEN = '1B7F5A', REDL = 'FDECEA', GREENL = 'E7F5EE';
const FONT = 'Calibri';
const W = 9638; // largeur utile A4, marges 2 cm

// ---------- formatage ----------
const fr = (x, d = 1) => {
  if (x === null || x === undefined) return '–';
  const neg = x < 0; let s = Math.abs(x).toFixed(d);
  if (Math.abs(x) < 0.5 * Math.pow(10, -d)) s = (0).toFixed(d);
  let [i, f] = s.split('.'); i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return (neg && parseFloat(s) !== 0 ? '−' : '') + i + (f ? ',' + f : '');
};
const pc = (x, d = 1) => fr(x * 100, d) + ' %';
const xx = (x) => (x === null || x === undefined) ? '–' : fr(x, 2) + 'x';

// ---------- primitives ----------
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color || INK, ...o.extra });
function rich(parts, o = {}) {
  // parts : chaîne avec **gras**
  const out = []; const re = /\*\*(.+?)\*\*/g; let last = 0, m;
  while ((m = re.exec(parts))) {
    if (m.index > last) out.push(run(parts.slice(last, m.index), o));
    out.push(run(m[1], { ...o, bold: true, color: o.boldColor || o.color || NAVY })); last = re.lastIndex;
  }
  if (last < parts.length) out.push(run(parts.slice(last), o));
  return out;
}
const P = (t, o = {}) => new Paragraph({ children: rich(t, o), alignment: o.align || AlignmentType.JUSTIFIED, spacing: { after: o.after ?? 120, before: o.before || 0, line: 276 }, keepNext: o.keepNext });
const SRC = (t) => new Paragraph({ children: [run(t, { size: 16, italics: true, color: MUTED })], spacing: { after: 160 } });
const BUL = (t, lvl = 0) => new Paragraph({ children: rich(t), numbering: { reference: 'bul', level: lvl }, alignment: AlignmentType.LEFT, spacing: { after: 60, line: 264 } });
const NUM = (t) => new Paragraph({ children: rich(t), numbering: { reference: 'num', level: 0 }, spacing: { after: 60, line: 264 } });
const SP = (a = 120) => new Paragraph({ children: [], spacing: { after: a } });
const BR = () => new Paragraph({ children: [new PageBreak()] });

let secNo = 0;
function H1(t) {
  secNo++;
  return [
  new Paragraph({
    heading: HeadingLevel.HEADING_1, keepNext: true, pageBreakBefore: true,
    children: [new TextRun({ text: String(secNo).padStart(2, '0') + '  ', font: FONT, size: 40, bold: true, color: GOLD }), new TextRun({ text: t.toUpperCase(), font: FONT, size: 32, bold: true, color: NAVY })],
    spacing: { before: 0, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GOLD, space: 6 } },
  }), SP(160)];
}
function H1u(t) { // titre non numéroté
  return [new Paragraph({
    heading: HeadingLevel.HEADING_1, keepNext: true, pageBreakBefore: true, children: [new TextRun({ text: t.toUpperCase(), font: FONT, size: 32, bold: true, color: NAVY })],
    spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GOLD, space: 6 } },
  }), SP(160)];
}
let h2n = 0; let lastSec = 0;
function H2(t) {
  if (lastSec !== secNo) { h2n = 0; lastSec = secNo; } h2n++;
  return new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 240, after: 100 }, children: [new TextRun({ text: `${secNo}.${h2n}  `, font: FONT, size: 25, bold: true, color: GOLD }), new TextRun({ text: t, font: FONT, size: 25, bold: true, color: BLUE })] });
}
const H2u = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 240, after: 100 }, children: [new TextRun({ text: t, font: FONT, size: 25, bold: true, color: BLUE })] });

const nob = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOB = { top: nob, bottom: nob, left: nob, right: nob, insideHorizontal: nob, insideVertical: nob };
const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };

function cell(content, w, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map(c => typeof c === 'string'
    ? new Paragraph({ alignment: o.align || AlignmentType.LEFT, spacing: { before: 0, after: 0, line: 252 }, children: rich(c, { size: o.size || 18, bold: o.bold, color: o.color || INK, boldColor: o.color }) })
    : c);
  return new TableCell({
    children: paras, width: { size: w, type: WidthType.DXA }, verticalAlign: o.valign || VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    margins: { top: o.pad ?? 60, bottom: o.pad ?? 60, left: 100, right: 100 }, columnSpan: o.span,
    borders: o.borders || { top: thin, bottom: thin, left: nob, right: nob },
  });
}
// Tableau standard : en-tête marine, lignes zébrées, colonnes numériques alignées à droite
function T(head, rows, widths, o = {}) {
  const tw = widths.reduce((a, c) => a + c, 0);
  const numCol = (j) => o.num ? (j > 0) : (o.numCols || []).includes(j);
  const hdr = new TableRow({ tableHeader: true, children: head.map((h, j) => cell(h, widths[j], { fill: NAVY, color: 'FFFFFF', bold: true, size: o.size || 18, align: numCol(j) ? AlignmentType.RIGHT : AlignmentType.LEFT })) });
  const body = rows.map((r, i) => {
    const isTot = o.totals && o.totals.includes(i); const hl = o.highlight && o.highlight.includes(i);
    return new TableRow({ cantSplit: true, children: r.map((c, j) => cell(c, widths[j], { fill: isTot ? LIGHT : hl ? GOLDL : (i % 2 ? ZEBRA : undefined), bold: isTot || (j === 0 && o.boldFirst), color: isTot ? NAVY : INK, size: o.size || 18, align: numCol(j) ? AlignmentType.RIGHT : AlignmentType.LEFT })) });
  });
  return [new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: widths, rows: [hdr, ...body], borders: NOB, layout: TableLayoutType.FIXED }), SP(o.after ?? 180)];
}
// Encadré (callout)
function BOX(title, text, o = {}) {
  const col = o.color || GOLD, fill = o.fill || GOLDL;
  const lines = Array.isArray(text) ? text : [text];
  return [new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W], borders: NOB,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill },
      margins: { top: 140, bottom: 140, left: 240, right: 200 },
      borders: { left: { style: BorderStyle.SINGLE, size: 36, color: col }, top: nob, bottom: nob, right: nob },
      children: [new Paragraph({ spacing: { after: 60 }, children: [run(title.toUpperCase(), { bold: true, size: 18, color: o.titleColor || NAVY, extra: { characterSpacing: 20 } })] }),
      ...lines.map(l => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, line: 264 }, children: rich(l, { size: 20 }) }))],
    })] })],
  }), SP(180)];
}
// Tuiles KPI
function KPI(items, o = {}) {
  const n = items.length; const w = Math.floor(W / n); const ws = items.map((_, i) => i === n - 1 ? W - w * (n - 1) : w);
  const white = { style: BorderStyle.SINGLE, size: 24, color: 'FFFFFF' };
  return [new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: ws, borders: NOB,
    rows: [new TableRow({ children: items.map((it, i) => new TableCell({
      width: { size: ws[i], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: o.fill || LIGHT },
      margins: { top: 140, bottom: 140, left: 80, right: 80 }, borders: { top: { style: BorderStyle.SINGLE, size: 18, color: o.accent || GOLD }, bottom: white, left: white, right: white },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [run(it[0], { bold: true, size: o.big || 30, color: o.valColor || NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [run(it[1], { size: 16, color: o.labColor || MUTED })] })],
    })) })],
  }), SP(o.after ?? 160)];
}
let figNo = 0;
function FIG(file, caption, source, wpx = 620) {
  figNo++; const img = fs.readFileSync(file);
  // ratio à partir de l'en-tête PNG
  const wi = img.readUInt32BE(16), hi = img.readUInt32BE(20);
  return [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 80, after: 40 }, children: [new ImageRun({ type: 'png', data: img, transformation: { width: wpx, height: Math.round(wpx * hi / wi) }, altText: { title: caption, description: caption, name: 'fig' + figNo } })] }),
  new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: source ? 0 : 180 }, children: [run(`Figure ${figNo} — `, { bold: true, size: 17, color: BLUE }), run(caption, { size: 17, italics: true, color: MUTED })] }),
  ...(source ? [SRC('Source : ' + source)] : [])];
}

// ---------- données dérivées ----------
const Y7 = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7'];
const w8 = [2638, 1000, 1000, 1000, 1000, 1000, 1000, 1000];
const S = Object.fromEntries(R.scen.map(([n, r]) => [n, r]));
const base = S['Cas de base'];
const vol130 = R.be.vol130, vol100 = R.be.vol100;
const TRI = pc(b.tri_fp), VAN = fr(b.van_fp, 2), DSCR = xx(b.dscr_min_amort), PAY = fr(b.payback, 1), TRIP = pc(b.tri_proj);
const TOT = fr(b.total, 1), DEBT = fr(b.debt, 1), EQ = fr(b.equity, 1);

// ---------- page de garde ----------
const logo = fs.readFileSync('v2/word/media/logo_alphab.jpg');
function cover() {
  const bandRow = (children, fill, pad = 200) => new TableRow({ children: [new TableCell({ width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill }, margins: { top: pad, bottom: pad, left: 400, right: 400 }, borders: { top: nob, bottom: nob, left: nob, right: nob }, children })] });
  const kp = [['100', 'camions-bennes lourds'], ['3,9 Mt/an', 'capacité stabilisée'], [TOT + ' M USD', 'coût total du projet'], [TRI, 'TRI fonds propres'], [DSCR, 'DSCR minimum']];
  return [
    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 360 }, children: [new ImageRun({ type: 'jpg', data: logo, transformation: { width: 190, height: 107 }, altText: { title: 'Logo', description: 'Logo ALPHA B HANOK SARL', name: 'logo' } })] }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], borders: NOB, rows: [
      bandRow([
        new Paragraph({ spacing: { after: 120 }, children: [run('BUSINESS PLAN  ·  DOSSIER DE FINANCEMENT', { bold: true, size: 20, color: GOLD, extra: { characterSpacing: 40 } })] }),
        new Paragraph({ spacing: { after: 80 }, children: [run('Société de transport minier', { bold: true, size: 56, color: 'FFFFFF' })] }),
        new Paragraph({ spacing: { after: 200 }, children: [run('Flotte de 100 camions-bennes lourds — République de Guinée', { size: 30, color: 'DCE6F5' })] }),
        new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 10 } }, spacing: { after: 0 }, children: [run('Corridor bauxitique Boké–Boffa  ·  Contrat de transport au tonnage  ·  Financement structuré sans recours limité', { size: 19, color: 'DCE6F5' })] }),
      ], NAVY, 560),
    ] }),
    SP(240),
    ...KPI(kp, { big: 26 }),
    SP(400),
    ...T(['Élément', 'Information'], [
      ['Promoteur (sponsor)', 'ALPHA B HANOK SARL — Construction (BTP), consultation, négoce, commerce général'],
      ['Emprunteur', 'Société de projet dédiée (SPV) de droit guinéen, régie par l’Acte uniforme OHADA'],
      ['Financement recherché', `Dette senior ${DEBT} M USD (65 %) — fonds propres ${EQ} M USD (35 %)`],
      ['Version', 'Version 3.0 consolidée — harmonisation des versions 1 et 2 et révision du modèle financier'],
      ['Date', 'Septembre 2026'],
      ['Classification', 'Strictement confidentiel — diffusion restreinte aux investisseurs et prêteurs sollicités'],
    ], [2800, W - 2800], { boldFirst: true }),
  ];
}

// ---------- contenu ----------
const C = [];
const add = (...xs) => xs.forEach(x => Array.isArray(x) ? C.push(...x) : C.push(x));

// Fiche synthétique (term sheet)
add(H1u('Fiche synthétique de la demande de financement'));
add(P('Cette fiche résume, au format d’une **term sheet indicative**, les paramètres de financement soumis aux prêteurs et investisseurs. Elle ne constitue pas une offre ; les conditions définitives résulteront de la due diligence et de la négociation avec les établissements financiers.'));
add(T(['Paramètre', 'Proposition indicative'], [
  ['Emprunteur', 'SPV de droit guinéen dédiée exclusivement au projet (société anonyme OHADA), détenue majoritairement par des personnes guinéennes (loi sur le contenu local)'],
  ['Sponsor', 'ALPHA B HANOK SARL, avec ouverture possible à un partenaire technique ou financier minoritaire'],
  ['Objet', 'Acquisition et mise en exploitation de 100 camions-bennes lourds, d’une flotte de soutien et d’un atelier central pour le transport de bauxite sous contrat pluriannuel'],
  ['Coût total du projet', `${TOT} M USD, dont investissements ${fr(b.hard + b.cont, 1)} M USD, fonds de roulement ${fr(b.fr0, 1)} M USD, frais et intérêts intercalaires ${fr(b.fees + b.idc, 2)} M USD, dotation initiale DSRA ${fr(b.dsra0, 2)} M USD`],
  ['Instrument', `Prêt senior à terme amortissable ou crédit-bail équivalent, ${DEBT} M USD, en USD`],
  ['Durée', '7 ans à compter du premier tirage, dont 12 mois de période de mobilisation (différé en capital)'],
  ['Taux indicatif', '10,5 % fixe (ou équivalent SOFR + marge), commission de montage 2,0 %, commission d’engagement sur non-tiré'],
  ['Remboursement', `Semestriel, profil sculpté sur un DSCR cible constant de ${xx(b.dscr_cible)} ; extinction à la fin de l’année 6 d’exploitation`],
  ['Fonds propres', `${EQ} M USD (35 %), apportés avant le premier tirage (capital et avances d’actionnaires subordonnées)`],
  ['Sûretés', 'Nantissement de la flotte et des équipements, cession des créances du contrat minier, nantissement des comptes et des titres de la SPV, délégation des polices d’assurance'],
  ['Comptes de réserve', 'DSRA de 6 mois de service de la dette ; compte de réserve de maintenance (MRA) pour les révisions majeures'],
  ['Engagements financiers', 'DSCR historique ≥ 1,30x (blocage des distributions < 1,20x ; défaut < 1,10x) ; LLCR ≥ 1,40x ; ratio dette / (dette + fonds propres) ≤ 70 %'],
  ['Conditions préalables clés', 'Contrat de transport signé (≥ 5 ans, take-or-pay, indexation carburant), audit de corridor, offres fermes constructeurs, avis fiscal, ESIA/ESMS, apport intégral des fonds propres'],
], [2600, W - 2600], { boldFirst: true }));
add(KPI([[TRI, 'TRI fonds propres'], [TRIP, 'TRI projet après impôt'], [VAN + ' M$', 'VAN fonds propres à 18 %'], [DSCR, 'DSCR min. / LLCR ' + xx(b.llcr_min)], [PAY + ' ans', 'retour des fonds propres']]));

// Avertissement
add(H1u('Avertissement et périmètre du document'));
add(BOX('Nature du document', 'Ce business plan est un dossier de décision et de recherche de financement. Il ne constitue ni une offre ferme d’un constructeur, ni un engagement bancaire, ni un avis juridique ou fiscal. Les statistiques sectorielles et les textes cités proviennent de sources publiques identifiées (section 17). Les prix des camions, le tarif de transport, les coûts unitaires, la productivité, les conditions de financement et les projections sont des **hypothèses de modélisation** à confirmer par des devis, une étude de route, un contrat minier et une due diligence indépendante.'));
add(T(['Catégorie', 'Traitement dans ce document'], [
  ['Données publiques vérifiées', 'Production et exportations minières, principaux opérateurs, poids du secteur, taux de change, cadre fiscal de référence et contenu local'],
  ['Hypothèses techniques', 'Charge utile, distance, rotations, disponibilité mécanique, consommation, jours d’exploitation'],
  ['Hypothèses commerciales', 'Tarif par tonne, volume minimum garanti, indexation, durée du contrat et délai de paiement'],
  ['Hypothèses financières', 'CAPEX, OPEX, conditions de la dette, amortissements, valeur résiduelle, fonds de roulement et comptes de réserve'],
], [3000, W - 3000], { boldFirst: true }));
add(P('**Normes de présentation.** Les projections sont établies en dollars américains (USD), en millions sauf mention contraire, en termes nominaux. Les états financiers prévisionnels suivent les principes des normes IFRS (IAS 16 pour les immobilisations, IAS 23 pour la capitalisation des intérêts intercalaires) ; les comptes statutaires de la SPV seront tenus selon le SYSCOHADA révisé. Les ratios de crédit (DSCR, LLCR, ratio d’endettement) sont calculés selon les conventions usuelles du financement de projet.'));
add(BOX('Condition de décision', 'La commande des 100 camions ne doit être lancée qu’après signature d’un contrat ferme, vérification du corridor, réception d’offres fournisseurs comparables, confirmation écrite du régime fiscal et douanier, et bouclage du financement.', { color: RED, fill: REDL }));

// Sommaire
add(H1u('Sommaire'));
const toc = ['Résumé exécutif', 'Présentation du projet et du promoteur', 'Analyse du marché minier guinéen', 'Stratégie commerciale et contractuelle', 'Plan technique et opérationnel', 'Organisation, ressources humaines et contenu local', 'HSE, normes environnementales et sociales', 'Cadre juridique, fiscal et institutionnel', 'Coût du projet et plan de financement', 'Hypothèses et prévisions financières', 'Analyse de sensibilité et tests de résistance', 'Analyse du risque de crédit', 'Gouvernance, conformité et reporting', 'Plan de mise en œuvre', 'Risques stratégiques et mesures de maîtrise', 'Conclusion et recommandation d’investissement', 'Sources documentaires'];
const tocRows = []; for (let i = 0; i < toc.length; i += 2) tocRows.push([String(i + 1).padStart(2, '0'), toc[i], toc[i + 1] ? String(i + 2).padStart(2, '0') : '', toc[i + 1] || '']);
add(T(['N°', 'Section', 'N°', 'Section'], tocRows, [600, 4219, 600, 4219]));
add(T(['Annexe', 'Contenu'], [
  ['A', 'Matrice complète des hypothèses'], ['B', 'Plan de due diligence avant financement'], ['C', 'États financiers prévisionnels détaillés'],
  ['D', 'Cahier des charges minimal des camions'], ['E', 'Note de révision : harmonisation des versions 1 et 2'], ['F', 'Glossaire financier'],
], [1200, W - 1200]));

// 1. Résumé exécutif
add(H1('Résumé exécutif'));
add(P('Le projet consiste à créer un **opérateur guinéen de transport minier intégré**, doté de 100 camions-bennes lourds, d’un atelier central, d’une flotte d’assistance, d’un système numérique de gestion de flotte et d’une organisation fonctionnant en continu. Le marché prioritaire est le transport de bauxite sur les corridors de Boké et de Boffa, entre sites miniers, stocks, installations de traitement et terminaux portuaires.'));
add(P('Le marché est profond : en 2025, la Guinée a produit **175,45 millions de tonnes** de bauxite et en a exporté **182,83 millions**, en hausse respective de 24 % et 25 % sur 2024. À capacité stabilisée (3,90 Mt/an), la flotte proposée représente environ **2,1 % des exportations nationales**, un objectif significatif mais compatible avec la taille du marché.'));
add(SRC('Source : Ministère des Mines et de la Géologie / ITIE-Guinée, Bulletin des Statistiques Minières et Carrières n°30, avril 2026.'));
add(KPI([['3,90 Mt/an', 'volume stabilisé'], ['7,10 USD/t', 'tarif initial modélisé'], [fr(b.ca[1], 1) + ' M$', "chiffre d'affaires A2"], [fr(b.ebitda[1], 1) + ' M$', 'EBITDA A2 (' + pc(b.ebitda_marge[1], 0) + ')']], { after: 60 }));
add(KPI([[TRI, 'TRI fonds propres'], [VAN + ' M$', 'VAN à 18 %'], [DSCR, 'DSCR minimum'], [PAY + ' ans', 'retour des fonds propres']], { accent: BLUE }));
add(H2u('Demande de financement'));
add(P(`Le coût total du projet s’élève à **${TOT} M USD**, financé à 65 % par une dette senior de **${DEBT} M USD** sur 7 ans et à 35 % par des fonds propres de **${EQ} M USD**. Le service de la dette est couvert en moyenne ${xx(b.dscr_moy)} sur la durée du prêt dans le cas de base, au-dessus du seuil de 1,30x usuellement exigé par les prêteurs pour un actif de ce profil.`));
add(BOX('Thèse d’investissement', `Le projet est **bancable sous quatre protections contractuelles** : volume minimum garanti (take-or-pay), indexation du carburant, responsabilité du client sur l’état de la route et mécanisme de paiement sécurisé (cession de créances et compte séquestre). La valeur ne réside pas dans la flotte elle-même, mais dans le **droit contractuel de transporter un volume rentable pendant une durée au moins égale à celle de la dette**.`));
add(T(['Critère de décision', 'Seuil recommandé', 'Cas de base'], [
  ['Durée ferme du contrat', '≥ 7 ans (durée de la dette) ; 5 ans minimum avec clause de reprise de dette', '7 ans modélisés'],
  ['Volume annuel garanti (take-or-pay)', `≥ ${fr(3.9 * vol130, 2)} Mt/an en régime (DSCR 1,30x)`, '3,55 Mt en A2 puis 3,90 Mt'],
  ['Tarif contractuel plancher', `≥ ${fr(R.be.tarif130, 2)} USD/t en A1, indexé (DSCR 1,30x)`, '7,10 USD/t en A1, indexé 2 %/an'],
  ['Indexation carburant', 'Mensuelle, ou carburant fourni par le client', 'Clause obligatoire'],
  ['Délai de paiement', '30 jours maximum, compte séquestre', '30 jours modélisés'],
  ['Apport des promoteurs', '≥ 35 % du coût total, avant tout tirage', `${EQ} M USD`],
], [3000, 3638, 3000], { boldFirst: true }));

// 2. Présentation
add(H1('Présentation du projet et du promoteur'));
add(H2('Objet et proposition de valeur'));
add(P('La société fournira un **service de transport minier facturé à la tonne**, incluant les véhicules, les chauffeurs, le carburant, la maintenance, les pneumatiques, la supervision, le suivi GPS, la sécurité, l’assurance et le reporting. Le client minier conserve la responsabilité du chargement, du déchargement, de la disponibilité de la route et des autorisations d’accès au site, sauf stipulation contraire.'));
add(H2('Périmètre des services'));
add(T(['Service', 'Description', 'Facturation'], [
  ['Transport principal de bauxite', 'Mine vers stock, concasseur, port ou point de transfert', 'USD par tonne validée'],
  ['Transport interne de site', 'Mouvements entre fosse, ROM pad, concasseur et stock', 'USD par heure ou par tonne'],
  ['Transport de matériaux', 'Latérite, stériles, agrégats, matériaux de route', 'USD par tonne-kilomètre'],
  ['Mise à disposition de flotte', 'Camions avec chauffeurs et maintenance', 'USD par camion-jour'],
  ['Prestations complémentaires', 'Dépannage, gestion de flotte, reporting, formation à la conduite', 'Forfait ou coût remboursable'],
], [2700, 4238, 2700], { boldFirst: true }));
add(H2('Positionnement'));
add(P('Le positionnement retenu est celui d’un **opérateur local structuré selon les standards miniers internationaux** : forte disponibilité mécanique, contrôle des consommations, prévention de la fatigue, traçabilité des tonnages, discipline de maintenance et production continue. Ce positionnement est plus défendable qu’une simple location de camions, car il permet de vendre une performance mesurable et contractualisée.'));
add(H2('Implantation'));
add(P('Le corridor Boké–Boffa est la zone prioritaire, en raison de la concentration des opérateurs bauxitiques et des infrastructures portuaires. La localisation définitive du dépôt sera arrêtée après analyse multicritère : distance au client, accès au carburant, disponibilité foncière, sécurité, drainage, proximité des pièces et acceptabilité communautaire.'));
add(BOX('Choix de marché', 'Le transport du minerai de fer de Simandou repose sur l’infrastructure ferroviaire transguinéenne. Le projet vise donc en priorité la bauxite, les mouvements internes de mine, les matériaux de construction et les services de soutien, plutôt qu’un transport routier longue distance du minerai de fer.'));
add(H2('Promoteur et capacité de mise en œuvre'));
add(P('Le projet est porté par **ALPHA B HANOK SARL**, société guinéenne active dans la construction (BTP), la consultation, le négoce et le commerce général. Ce positionnement apporte une connaissance du tissu économique local, des chantiers et des circuits d’approvisionnement.'));
add(P('Les prêteurs internationaux évaluent en priorité **l’expérience opérationnelle du sponsor dans le transport minier**. Pour sécuriser ce point, le plan retient les mesures suivantes :'));
add(BUL('recrutement d’un directeur général et d’un directeur technique justifiant chacun d’au moins 10 ans d’expérience en gestion de flottes minières ;'));
add(BUL('contrat de maintenance et d’assistance technique avec le constructeur (OEM) pendant les 24 premiers mois, avec techniciens résidents ;'));
add(BUL('ouverture possible du capital à un partenaire technique minoritaire disposant de références en Afrique de l’Ouest ;'));
add(BUL('constitution d’un dossier sponsor complet : statuts, états financiers certifiés des 3 derniers exercices, références de chantiers, identification des bénéficiaires effectifs (KYC).'));

// 3. Marché
add(H1('Analyse du marché minier guinéen'));
add(H2('Poids économique et dynamique sectorielle'));
add(P('Le secteur extractif est central dans l’économie guinéenne. Selon l’ITIE, il représentait 20 % du PIB, 92 % des exportations et plus de 17 % des recettes publiques en 2022 ; pour 2023, l’ITIE indique 93 % des exportations et 22 % des recettes publiques. Cette concentration crée un marché de services important, mais expose les sous-traitants aux cycles miniers, aux décisions réglementaires et à la dépendance envers quelques grands donneurs d’ordre.'));
add(SRC('Source : Initiative pour la Transparence dans les Industries Extractives, fiche pays Guinée et décision de validation 2026.'));
add(FIG('fig1.png', 'Production et exportations de bauxite, 2024–2025', 'MMG / ITIE-Guinée, Bulletin n°30.'));
add(H2('Structure concurrentielle et clients cibles'));
add(P('Les exportations de 2025 sont concentrées autour de quelques opérateurs majeurs : **SMB (71,52 Mt), CHALCO (22,12 Mt), CBG (17,37 Mt) et AGB2A SDM (17,01 Mt)**. Les cibles commerciales seront priorisées selon la maturité de leur infrastructure, la distance de roulage, leur politique de sous-traitance, leur qualité de paiement et leur disposition à signer des contrats pluriannuels.'));
add(FIG('fig2.png', 'Principaux exportateurs de bauxite en 2025', 'MMG / ITIE-Guinée, Bulletin n°30, tableau 2.'));
add(H2('Saisonnalité et contraintes physiques'));
add(P('La baisse des volumes pendant l’hivernage (juin à octobre) confirme l’importance du drainage, de la maintenance préventive, du stock de pneumatiques et d’une réserve de flotte. Le modèle ne suppose pas une production uniforme : la première année est volontairement limitée et **dix camions sont maintenus en réserve technique**.'));
add(FIG('fig3.png', 'Profil mensuel de la production et des exportations de bauxite en 2025', 'MMG / ITIE-Guinée, Bulletin n°30.'));
add(H2('Taille de marché adressable'));
add(T(['Indicateur', 'Valeur', 'Lecture pour le projet'], [
  ['Exportations nationales de bauxite 2025', '182,83 Mt', 'Base de marché publique'],
  ['Capacité du projet en régime stabilisé', '3,90 Mt/an', 'Environ 2,1 % du volume exporté en 2025'],
  ['Volume du contrat en année 2', '3,55 Mt', 'Environ 1,9 % du volume exporté en 2025'],
  ['Portefeuille clients recommandé', '1 contrat principal + 1 secondaire', 'Réduction du risque de concentration'],
], [3600, 2000, 4038], { boldFirst: true }));
add(H2('Opportunités et points de vigilance'));
add(T(['Opportunités', 'Points de vigilance'], [
  ['Croissance récente de la production et des exportations', 'Marché concentré autour de quelques opérateurs'],
  ['Besoins de sous-traitance locale et exigences de contenu local', 'Forte pression concurrentielle sur les tarifs'],
  ['Développement de raffineries et d’infrastructures', 'Risque de substitution par convoyeur ou rail'],
  ['Besoin de professionnalisation HSE et numérique', 'Exposition au diesel, au change, aux pneus et aux pièces importées'],
], [W / 2, W / 2]));

// 4. Stratégie commerciale
add(H1('Stratégie commerciale et contractuelle'));
add(H2('Approche de commercialisation'));
add(P('La stratégie consiste à **sécuriser d’abord un contrat principal adossé à un volume, puis à mobiliser la flotte**, et non l’inverse. Les étapes comprennent la préqualification, l’étude du corridor, la remise d’une offre technique, la négociation d’un modèle tarifaire transparent et la validation des garanties de paiement.'));
add(H2('Architecture contractuelle minimale'));
add(T(['Clause', 'Exigence recommandée', 'Justification'], [
  ['Durée', '7 ans ferme (5 ans minimum)', 'Adossement à la durée de la dette'],
  ['Volume garanti', `≥ ${fr(3.9 * vol130, 2)} Mt/an en régime`, 'Maintien d’un DSCR ≥ 1,30x'],
  ['Take-or-pay', '85 % du volume mensuel programmé', 'Protection contre l’indisponibilité imputable au client'],
  ['Tarif initial', `≥ ${fr(R.be.tarif130, 2)} USD/t pour 35 km aller simple`, 'Seuil bancaire du cas de base'],
  ['Carburant', 'Indexation mensuelle ou fourniture par le client', 'Neutralisation du principal coût variable'],
  ['Paiement', '30 jours, facturation bimensuelle, compte séquestre', 'Protection du fonds de roulement'],
  ['Route', 'Entretien et disponibilité à la charge du client', 'Préservation de la productivité et des pneus'],
  ['Temps d’attente', 'Franchise définie puis facturation', 'Protection contre les files de chargement'],
  ['Résiliation anticipée', 'Indemnité couvrant l’encours de dette et la démobilisation', 'Protection des prêteurs et des actionnaires'],
  ['Cession de créances', 'Acceptation par le client de la cession au profit des prêteurs', 'Condition standard de bancabilité'],
  ['Pesage', 'Pont-bascule certifié et rapprochement quotidien', 'Sécurisation du chiffre d’affaires'],
], [2300, 3669, 3669], { boldFirst: true }));
add(H2('Formule tarifaire proposée'));
add(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: LIGHT }, children: [run('Tₙ = T₀ × [ 40 % + 40 % × (Dieselₙ / Diesel₀) + 20 % × (FXₙ / FX₀) ]', { bold: true, size: 24, color: NAVY })] }));
add(P('La formule répartit le tarif entre une composante fixe, une composante carburant et une composante de change. Elle sera recalibrée sur la structure réelle des coûts (le carburant représente environ un tiers des coûts d’exploitation en année 2) et sur les conditions de fourniture du diesel.'));
add(H2('Pipeline commercial'));
add(T(['Segment cible', 'Proposition', 'Priorité'], [
  ['Grand producteur bauxitique', 'Contrat principal de 70 à 80 camions', 'Très élevée'],
  ['Deuxième opérateur ou projet de raffinerie', 'Contrat complémentaire de 20 à 30 camions', 'Élevée'],
  ['Entreprises de travaux miniers', 'Mouvements de stériles et de matériaux', 'Moyenne'],
  ['Projets d’infrastructure', 'Latérite, agrégats, terrassement', 'Opportuniste'],
], [3600, 4038, 2000], { boldFirst: true }));

// 5. Technique
add(H1('Plan technique et opérationnel'));
add(H2('Hypothèses de dimensionnement'));
add(T(['Paramètre', 'Hypothèse centrale', 'Statut'], [
  ['Flotte totale', '100 camions-bennes lourds', 'Dimensionnement du projet'], ['Flotte programmée', '90 camions/jour', 'Hypothèse'],
  ['Réserve technique', '10 camions', 'Hypothèse'], ['Charge utile contractuelle', '45 tonnes', 'À valider (audit de route et constructeur)'],
  ['Distance aller simple', '35 km', 'Corridor de référence, à remplacer par le trajet réel'], ['Rotations moyennes', '3 par jour', 'À confirmer par chronométrage'],
  ['Jours productifs', '321 jours/an', 'Inclut arrêts et saisonnalité'], ['Consommation', '1,05 litre par tonne transportée', 'Hypothèse de modélisation'],
  ['Disponibilité mécanique', '≥ 85 %', 'KPI contractuel'],
], [3000, 3319, 3319], { boldFirst: true }));
add(H2('Calcul de capacité'));
add(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 120 }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: LIGHT }, children: [run('90 camions × 45 t × 3 rotations × 321 jours = 3 900 000 tonnes/an', { bold: true, size: 24, color: NAVY })] }));
add(P('Cette capacité est un maximum contractuel stabilisé et non une garantie indépendante des conditions de chargement, de la route et du port. La productivité sera mesurée par cycle complet, et non uniquement par la disponibilité des camions.'));
add(H2('Composition de la flotte'));
add(T(['Équipement', 'Quantité', 'Fonction'], [
  ['Camions-bennes lourds', '100', 'Transport principal'], ['Camions-ateliers mobiles', '4', 'Dépannage et maintenance de proximité'],
  ['Citernes carburant', '2', 'Distribution contrôlée du diesel'], ['Dépanneuse lourde', '1', 'Récupération des véhicules immobilisés'],
  ['Véhicules HSE et supervision', '8', 'Contrôle opérationnel et interventions'], ['Minibus du personnel', '2', 'Relève et transport des équipes'],
  ['Ambulance', '1', 'Réponse médicale d’urgence'],
], [3600, 1400, 4638], { boldFirst: true }));
add(H2('Processus opérationnel'));
add(T(['Étape', 'Contrôle essentiel'], [
  ['1', 'Planification de la production et affectation des chauffeurs'], ['2', 'Inspection avant départ et autorisation de mise en service'],
  ['3', 'Chargement et contrôle de la charge utile'], ['4', 'Transport sous surveillance GPS/IVMS'],
  ['5', 'Pesage, déchargement et validation du tonnage'], ['6', 'Retour, gestion de file et nouvelle rotation'],
  ['7', 'Rapprochement journalier tonnage – carburant – kilométrage'], ['8', 'Maintenance préventive selon kilométrage et heures moteur'],
], [1000, W - 1000]));
add(H2('Maintenance et pièces'));
add(P('L’atelier central couvrira la maintenance préventive, le diagnostic électronique, les pneumatiques, le graissage, le lavage, la soudure, les réparations de benne et la gestion des composants majeurs. Le stock initial sera établi par criticité et délai d’approvisionnement. Les moteurs, boîtes, ponts, suspensions, freins et pneus feront l’objet d’un programme de remplacement fondé sur l’état réel, financé par un **compte de réserve de maintenance**.'));
add(T(['KPI technique', 'Cible'], [
  ['Disponibilité mécanique', '≥ 85 %'], ['Respect de la maintenance planifiée', '≥ 95 %'], ['Taux de panne en ligne', '≤ 3 % des départs'],
  ['Immobilisation moyenne', 'Suivie par type de panne'], ['Exactitude du stock de pièces', '≥ 98 %'], ['Écart carburant physique / système', '≤ 1 %'],
], [W / 2, W / 2], { boldFirst: true }));

// 6. Organisation
add(H1('Organisation, ressources humaines et contenu local'));
add(H2('Structure de gouvernance'));
add(T(['Niveau', 'Fonctions clés'], [
  ['Conseil d’administration', 'Stratégie, financement, audit, nomination de la direction ; comité d’audit et des risques'],
  ['Direction générale', 'Contrat client, performance globale, conformité, relation prêteurs'],
  ['Direction des opérations', 'Production, dispatch, chauffeurs, relation avec le site'],
  ['Direction technique', 'Maintenance, pièces, pneus, carburant'],
  ['Direction HSE-Q', 'Sécurité, qualité, environnement, social, enquêtes incidents'],
  ['Direction administrative et financière', 'Comptabilité IFRS/SYSCOHADA, trésorerie, fiscalité, contrôle de gestion, reporting prêteurs'],
  ['Ressources humaines', 'Recrutement, formation, paie, discipline, relations sociales'],
], [3200, W - 3200], { boldFirst: true }));
add(H2('Effectif cible'));
add(T(['Catégorie', 'Effectif indicatif'], [
  ['Chauffeurs titulaires, relèves et réserve', '220'], ['Mécaniciens, électriciens et soudeurs', '30'], ['Techniciens pneus, graissage et lavage', '16'],
  ['Dispatchers et contrôleurs de flotte', '12'], ['HSE, qualité et sécurité routière', '15'], ['Magasin, achats et carburant', '10'],
  ['Administration, finance et RH', '18'], ['Direction et responsables', '8'], ['Sécurité, nettoyage et soutien', '15'], ['Total indicatif', '344'],
], [6638, 3000], { numCols: [1], totals: [9] }));
add(P('L’effectif sera recalculé après validation des cycles de travail, de la politique de relève et des services externalisés. Le budget salarial du modèle est un plafond agrégé, et non une grille de rémunération individuelle.'));
add(H2('Contenu local'));
add(P('La loi L/2022/0010/CNT définit l’entreprise locale comme une entreprise établie et immatriculée en Guinée, dont le capital est détenu à au moins **51 %** par des personnes guinéennes, avec au moins **50 %** de personnel dirigeant et **75 %** de personnel d’exécution guinéens. Elle prévoit également l’approvisionnement local en biens et services et la transmission de plans d’approvisionnement.'));
add(SRC('Source : Conseil National de la Transition, loi L/2022/0010/CNT portant contenu local, articles 1 à 3 et article 8.'));
add(BOX('Orientation de structuration', 'La SPV sera constituée en droit guinéen, avec un plan de nationalisation des postes, un programme de formation des chauffeurs et techniciens et une politique documentée d’achats locaux. Toute entrée d’un partenaire étranger sera plafonnée de façon à préserver la qualification d’entreprise locale.'));
add(H2('Formation'));
add(T(['Population', 'Formation minimale'], [
  ['Chauffeurs', 'Conduite défensive, route minière, fatigue, freinage, inspection, premiers secours'],
  ['Mécaniciens', 'Procédures constructeur, consignation, levage, diagnostic et sécurité atelier'],
  ['Dispatchers', 'Planification, système de flotte, gestion des incidents et reporting'],
  ['Managers', 'Leadership HSE, gestion contractuelle, coûts, enquête et conformité'],
], [2400, W - 2400], { boldFirst: true }));

// 7. HSE / E&S
add(H1('HSE, normes environnementales et sociales'));
add(H2('Référentiel international applicable'));
add(P('Les institutions financières internationales conditionnent leur financement au respect de standards environnementaux et sociaux reconnus. Le projet est, selon toute vraisemblance, de **catégorie B** au sens des **Principes de l’Équateur** (impacts limités, largement réversibles et maîtrisables). Le système de management sera aligné sur :'));
add(T(['Référentiel', 'Application au projet'], [
  ['Normes de performance IFC (PS1 à PS4)', 'Système de gestion E&S (ESMS), conditions de travail, prévention de la pollution, santé et sécurité des communautés'],
  ['Principes de l’Équateur IV', 'Catégorisation, évaluation E&S, plan d’action E&S (ESAP), mécanisme de plainte, suivi indépendant'],
  ['Directives EHS du Groupe Banque mondiale', 'Directives générales et directives transport routier / hydrocarbures'],
  ['ISO 45001, ISO 14001, ISO 9001', 'Certification visée dans les 24 mois suivant la mise en service'],
  ['ISO 39001', 'Management de la sécurité routière, recommandé pour un opérateur de flotte lourde'],
  ['Conventions fondamentales de l’OIT', 'Interdiction du travail forcé et du travail des enfants, liberté syndicale, non-discrimination'],
], [3400, W - 3400], { boldFirst: true }));
add(H2('Principaux dangers et mesures critiques'));
add(T(['Danger', 'Mesure critique'], [
  ['Collision ou renversement', 'Limiteurs, IVMS, contrôle de vitesse, conduite défensive, voies séparées'],
  ['Fatigue', 'Rosters, contrôle des heures, pauses, détection de somnolence, aptitude médicale'],
  ['Surcharge', 'Pont-bascule, refus de départ, contrôle de la benne'],
  ['Défaillance de freinage', 'Inspection quotidienne, maintenance, tests et zones de secours'],
  ['Incendie / carburant', 'Extincteurs, mise à la terre, permis, confinement et plan d’urgence'],
  ['Poussière et visibilité', 'Arrosage, distance de sécurité, éclairage et limitation de vitesse'],
  ['Interaction communautaire', 'Plan de circulation, signalisation, sensibilisation et mécanisme de plainte'],
  ['Travaux en atelier', 'Consignation, levage certifié, EPI et permis de travail'],
], [3000, W - 3000], { boldFirst: true }));
add(H2('Objectifs de performance HSE'));
add(KPI([['0', 'accident mortel'], ['100 %', 'chauffeurs habilités'], ['≥ 95 %', 'inspections réalisées'], ['≤ 1 %', 'excès de vitesse critiques']], { big: 26, after: 60 }));
add(KPI([['100 %', 'incidents enquêtés'], ['≥ 98 %', 'port des EPI'], ['≤ 3 %', 'pannes en ligne'], ['24/7', 'capacité d’urgence']], { big: 26, accent: BLUE }));
add(H2('Reporting quotidien'));
add(T(['Domaine', 'Indicateurs'], [
  ['Production', 'Tonnes, rotations, temps de cycle, attente, distance'], ['Flotte', 'Disponibilité, utilisation, pannes, immobilisation'],
  ['Carburant', 'Litres, L/t, écarts, ravitaillements'], ['HSE', 'Incidents, presque-accidents, vitesse, fatigue'],
  ['Qualité', 'Écarts de pesage, réclamations, conformité documentaire'], ['Finance', 'CA journalier, coût variable, marge par camion et par trajet'],
], [2400, W - 2400], { boldFirst: true }));

// 8. Juridique & fiscal
add(H1('Cadre juridique, fiscal et institutionnel'));
add(H2('Textes de référence'));
add(T(['Domaine', 'Référence principale', 'Incidence pour le projet'], [
  ['Droit des sociétés et sûretés', 'Actes uniformes OHADA (sociétés commerciales, sûretés)', 'Forme de la SPV, nantissements, cession de créances'],
  ['Secteur minier', 'Code minier 2011 amendé en 2013', 'Accès aux sites, contrats avec les titulaires, exigences du donneur d’ordre'],
  ['Contenu local', 'Loi L/2022/0010/CNT', 'Actionnariat, emploi, achats locaux, reporting'],
  ['Fiscalité', 'Code général des impôts 2021 et lois de finances', 'Impôt sur les sociétés, TVA, retenues, obligations déclaratives'],
  ['Travail et sécurité sociale', 'Code du travail et textes CNSS', 'Contrats, temps de travail, paie, santé et sécurité'],
  ['Transport et circulation', 'Textes nationaux et règles du site minier', 'Immatriculation, assurance, permis, charge et circulation'],
  ['Environnement', 'Code de l’environnement et EIES du client', 'Déchets, hydrocarbures, poussière, bruit, urgences'],
], [2400, 3419, 3819], { boldFirst: true }));
add(H2('Fiscalité retenue dans le modèle'));
add(P('L’article 229 du Code général des impôts 2021 fixe l’impôt sur les sociétés à **25 %** pour les personnes morales autres que celles des secteurs expressément soumis à 30 % ou 35 %. Une société de transport indépendante, non titulaire d’un titre minier, est donc modélisée à 25 %, sous réserve de confirmation écrite du traitement fiscal, des éventuels agréments (Code des investissements) et des lois de finances ultérieures. Les droits et taxes à l’importation des camions sont réputés inclus dans le prix « rendu Guinée » et doivent être confirmés par un avis douanier.'));
add(SRC('Source : Direction Générale des Impôts, Code général des impôts 2021, article 229.'));
add(H2('Taux de change'));
add(P('Le modèle est établi en USD afin d’aligner les investissements importés, la dette et les contrats miniers. À titre de conversion, le fixing BCRG du 31 juillet 2026 est de **8 762,7692 GNF pour 1 USD**. La société tiendra une comptabilité conforme aux obligations guinéennes et gérera le risque de conversion des flux payés en GNF (salaires, achats locaux).'));
add(SRC('Source : Banque Centrale de la République de Guinée, fixing du 31 juillet 2026.'));
add(H2('Autorisations et conformité préalables'));
add(T(['Action', 'Responsable', 'Avant mobilisation'], [
  ['Constitution, RCCM, NIF, comptes bancaires', 'Promoteurs / conseil juridique', 'Oui'], ['Agrément ou enregistrement contenu local', 'Direction / conseil', 'À confirmer'],
  ['Contrats de travail et affiliation CNSS', 'RH', 'Oui'], ['Immatriculation et assurance de la flotte', 'Logistique / assureur', 'Oui'],
  ['Permis de conduire et habilitations site', 'RH / HSE', 'Oui'], ['Permis environnementaux du dépôt et de l’atelier', 'HSE / autorités', 'Oui'],
  ['Conformité fiscale et douanière des importations', 'DAF / transitaire', 'Oui'], ['Accord d’accès au site et plan de circulation', 'Client minier', 'Oui'],
], [4638, 3200, 1800], { boldFirst: true }));

// 9. Financement
add(H1('Coût du projet et plan de financement'));
add(H2('Budget d’investissement'));
const capex = [['100 camions-bennes rendus en Guinée', 16.5], ['Véhicules et équipements de soutien', 2.0], ['Atelier, dépôt, magasin et aire de lavage', 2.2], ['Stock initial de pièces stratégiques, pneus et lubrifiants', 2.0], ['GPS, IVMS, logiciels et informatique', 0.5], ['Mobilisation, formation, assurance et homologation', 1.0]];
add(T(['Poste', 'M USD', 'Part'], [
  ...capex.map(([n, v]) => [n, fr(v, 2), pc(v / 30, 1)]),
  ['Sous-total investissements physiques', fr(24.2, 2), pc(24.2 / 30, 1)],
  ['Provision pour aléas et imprévus (11,6 % des investissements)', fr(2.8, 2), pc(2.8 / 30, 1)],
  ['Fonds de roulement et trésorerie de démarrage', fr(3.0, 2), pc(0.1, 1)],
  ['Enveloppe d’investissement', fr(30, 2), '100,0 %'],
], [6638, 1500, 1500], { numCols: [1, 2], totals: [6, 9] }));
add(P('Le prix moyen implicite d’un camion rendu en Guinée est de **165 000 USD**. Ce montant sera remplacé par des offres fermes incluant benne, options minières, transport, assurance, droits et mise en service, issues d’une mise en concurrence d’au moins trois constructeurs.'));
add(H2('Tableau des emplois et ressources'));
add(P('Conformément aux pratiques des prêteurs, le coût total à financer intègre, au-delà de l’enveloppe d’investissement, les **frais de financement**, les **intérêts intercalaires** de la période de mobilisation (capitalisés selon IAS 23) et la **dotation initiale du compte de réserve du service de la dette (DSRA)**.'));
add(T(['Emplois', 'M USD', 'Ressources', 'M USD'], [
  ['Investissements physiques', fr(b.hard, 2), 'Fonds propres (capital et avances d’actionnaires)', fr(b.equity, 2)],
  ['Provision pour aléas', fr(b.cont, 2), 'Dette senior à terme', fr(b.debt, 2)],
  ['Fonds de roulement initial', fr(b.fr0, 2), '', ''],
  ['Frais de montage et conseils des prêteurs', fr(b.fees, 2), '', ''],
  ['Intérêts intercalaires (année 0)', fr(b.idc, 2), '', ''],
  ['Dotation initiale DSRA', fr(b.dsra0, 2), '', ''],
  ['Total des emplois', fr(b.total, 2), 'Total des ressources', fr(b.total, 2)],
], [3500, 1319, 3500, 1319], { numCols: [1, 3], totals: [6] }));
add(H2('Structure de financement'));
add(T(['Source', 'M USD', 'Part', 'Conditions modélisées'], [
  ['Fonds propres', fr(b.equity, 2), '35 %', 'Apport intégral avant le premier tirage de la dette'],
  ['Dette senior / crédit-bail', fr(b.debt, 2), '65 %', '10,5 %, 7 ans dont 12 mois de mobilisation, remboursement sculpté'],
  ['Total', fr(b.total, 2), '100 %', ''],
], [2600, 1300, 1000, 4738], { numCols: [1, 2], totals: [2] }));
add(P('**Sources de financement ciblées.** Banques commerciales de la place (financement local en USD), institutions de financement du développement actives dans le secteur minier ouest-africain (IFC, Afreximbank, Banque africaine de développement, Proparco, etc.), financements adossés à une agence de crédit à l’exportation du pays du constructeur, crédit-bail fournisseur, et couverture du risque politique (ATI, MIGA) pour améliorer la notation du risque.'));
add(H2('Sûretés et garanties'));
add(T(['Instrument', 'Objet'], [
  ['Nantissement de la flotte et des équipements', 'Garantie sur les actifs financés'],
  ['Cession des créances du contrat minier', 'Paiement direct sur un compte séquestre au profit des prêteurs'],
  ['Nantissement des comptes et des titres de la SPV', 'Contrôle des flux et possibilité de substitution (step-in)'],
  ['Compte de réserve du service de la dette (DSRA)', '6 mois de service de la dette'],
  ['Compte de réserve de maintenance (MRA)', 'Financement des révisions majeures'],
  ['Assurances tous risques et perte d’exploitation', 'Délégation au profit des prêteurs'],
  ['Engagement de volume du client (take-or-pay)', 'Sécurisation du chiffre d’affaires'],
  ['Garantie des actionnaires pendant la mobilisation', 'Couverture du risque de dépassement de coûts et de retard'],
], [4400, W - 4400], { boldFirst: true }));
add(H2('Cascade des paiements (waterfall)'));
add(P('Toutes les recettes du contrat sont encaissées sur un compte de recettes nanti. Elles sont affectées selon l’ordre de priorité suivant, standard en financement de projet :'));
add(T(['Rang', 'Affectation'], [
  ['1', 'Charges d’exploitation, impôts et taxes'], ['2', 'Frais, commissions et intérêts de la dette senior'], ['3', 'Remboursement du principal de la dette senior'],
  ['4', 'Reconstitution du DSRA à son niveau cible'], ['5', 'Dotation au compte de réserve de maintenance'],
  ['6', 'Distributions aux actionnaires, si DSCR ≥ 1,20x, réserves pleines et absence de défaut'],
], [1000, W - 1000]));
add(H2('Engagements financiers (covenants)'));
add(T(['Ratio', 'Seuil de distribution', 'Cas de défaut', 'Cas de base'], [
  ['DSCR historique (12 mois glissants)', '≥ 1,20x', '< 1,10x', `${xx(b.dscr_min_amort)} min.`],
  ['LLCR', '≥ 1,40x', '< 1,20x', `${xx(b.llcr_min)} min.`],
  ['Dette / (dette + fonds propres)', '—', '> 70 %', `${pc(X.ratios[0].gearing, 0)} max. (A1)`],
  ['DSRA', 'Niveau cible atteint', 'Non reconstitué sous 60 jours', '6 mois'],
], [3200, 2146, 2146, 2146], { boldFirst: true }));

// 10. Prévisions financières
add(H1('Hypothèses et prévisions financières'));
add(H2('Hypothèses économiques centrales'));
add(T(['Hypothèse', 'Valeur'], [
  ['Calendrier', 'Année 0 : mobilisation de 12 mois ; années 1 à 7 : exploitation'],
  ['Tarif initial', '7,10 USD/t, indexé de 2 % par an'],
  ['Volumes', '2,75 Mt (A1, montée en cadence), 3,55 Mt (A2), 3,85 Mt (A3), 3,90 Mt/an ensuite'],
  ['Prix du diesel de référence', '1,37 USD/litre (12 000 GNF au fixing retenu), + 1,5 % par an'],
  ['Consommation', '1,05 litre par tonne'],
  ['Autres coûts variables', '1,46 USD/t (maintenance 0,82 ; pneus 0,43 ; lubrifiants et divers 0,21), + 3 % par an'],
  ['Coûts fixes (personnel, assurances, structure)', '4,80 M USD en A1, + 3 % par an'],
  ['CAPEX de maintien', '0,3 à 1,2 M USD par an (révisions majeures), soit 5,6 M USD sur 7 ans'],
  ['Besoin en fonds de roulement', '12 % du chiffre d’affaires (≈ 44 jours), récupéré en fin d’horizon'],
  ['Amortissements', 'Linéaire sur 7 ans jusqu’à une valeur résiduelle de 5,0 M USD (IAS 16)'],
  ['Impôt sur les sociétés', '25 %, avec report des déficits'],
  ['Coût des fonds propres (actualisation)', '18 % ; coût moyen pondéré du capital ≈ ' + pc(X.wacc, 1)],
], [3600, W - 3600], { boldFirst: true }));
add(H2('Compte de résultat prévisionnel'));
const row = (lab, arr, d = 2) => [lab, ...arr.map(v => fr(v, d))];
add(T(['M USD', ...Y7], [
  row('Volume (Mt)', b.vol, 2), row('Tarif (USD/t)', b.tarif, 2), row("Chiffre d'affaires", b.ca),
  row('Carburant', b.carb.map(v => -v)), row('Autres coûts variables', b.autres.map(v => -v)), row('Coûts fixes', b.fixes.map(v => -v)),
  row('EBITDA', b.ebitda), ['Marge EBITDA', ...b.ebitda_marge.map(v => pc(v, 0))],
  row('Amortissements', b.dep.map(v => -v)), row("Résultat d'exploitation (EBIT)", b.ebit), row('Charges financières', b.interets.map(v => -v)),
  row('Impôt sur les sociétés', b.impot.map(v => -v)), row('Résultat net', b.rn),
], w8, { num: true, size: 17, totals: [6, 12], highlight: [2] }));
add(FIG('fig4.png', "Chiffre d'affaires et EBITDA prévisionnels", null));
add(H2('Structure des coûts'));
add(FIG('fig5.png', "Répartition des coûts d'exploitation de l'année 2", null, 600));
add(H2('Tableau des flux de trésorerie et service de la dette'));
add(T(['M USD', ...Y7], [
  row('EBITDA', b.ebitda), row('Impôt payé', b.impot.map(v => -v)), row('CAPEX de maintien', b.cm.map(v => -v)),
  row('Variation du BFR', b.dbfr.map(v => -v)), row('Cession des actifs (valeur résiduelle)', [0, 0, 0, 0, 0, 0, 5]),
  row('Flux disponible pour le service de la dette (CFADS)', b.cfads), row('Intérêts', b.interets.map(v => -v)), row('Principal', b.princ.map(v => -v)),
  ['DSCR', ...b.dscr.map(xx)], row('Mouvement du DSRA', b.dsra_mvt.map(v => -v)), row('Flux disponibles pour les actionnaires', b.flux_act),
], w8, { num: true, size: 17, totals: [5, 10], highlight: [8] }));
add(P('Le fonds de roulement initial (3,0 M USD) finance la constitution du BFR en année 1. Une trésorerie minimale de 1,0 M USD est conservée en permanence ; le solde est distribué sous réserve des tests de distribution. La dernière année intègre la cession de la flotte à sa valeur résiduelle et la récupération du BFR.'));
add(H2('Profil de la dette et couverture'));
add(T(['M USD', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6'], [
  row('Encours début', b.dette_deb.slice(0, 6)), row('Intérêts', b.interets.slice(0, 6)), row('Principal', b.princ.slice(0, 6)),
  row('Encours fin', b.dette_fin.slice(0, 6)), ['DSCR', ...b.dscr.slice(0, 6).map(xx)], ['LLCR', ...b.llcr.slice(0, 6).map(xx)],
], [2638, 1167, 1167, 1167, 1167, 1166, 1166], { num: true, size: 17, highlight: [4] }));
add(FIG('fig6.png', 'Encours de dette et ratio de couverture du service de la dette', null));
add(H2('Bilan prévisionnel simplifié'));
const Bl = X.bilan; const col7 = [0, 1, 2, 3, 4, 5, 6];
add(T(['M USD', 'A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6'], [
  row('Immobilisations nettes', col7.map(i => Bl.immo[i])), row('Besoin en fonds de roulement', col7.map(i => Bl.bfr[i])),
  row('Compte DSRA', col7.map(i => Bl.dsra[i])), row('Trésorerie', col7.map(i => Bl.cash[i])),
  row('Total actif', col7.map(i => Bl.immo[i] + Bl.bfr[i] + Bl.dsra[i] + Bl.cash[i])),
  row('Fonds propres (y c. avances d’actionnaires)', col7.map(i => Bl.fp[i])), row('Dette senior', col7.map(i => Bl.dette[i])),
  row('Total passif', col7.map(i => Bl.fp[i] + Bl.dette[i])),
], w8, { num: true, size: 17, totals: [4, 7] }));
add(H2('Rentabilité et ratios clés'));
add(KPI([[TRI, 'TRI fonds propres'], [TRIP, 'TRI projet après impôt'], [VAN + ' M$', 'VAN fonds propres à 18 %'], [PAY + ' ans', 'retour après mise en service']]));
add(T(['Ratio', ...Y7.slice(0, 6)], [
  ['Dette nette / EBITDA', ...X.ratios.slice(0, 6).map(r => xx(r.nd_ebitda))],
  ['Couverture des intérêts (EBITDA / intérêts)', ...X.ratios.slice(0, 6).map(r => xx(r.icr))],
  ['Dette / (dette + fonds propres)', ...X.ratios.slice(0, 6).map(r => pc(r.gearing, 0))],
  ['Rentabilité des fonds propres (ROE)', ...X.ratios.slice(0, 6).map(r => pc(r.roe, 0))],
  ['Marge nette', ...X.ratios.slice(0, 6).map(r => pc(r.marge_nette, 0))],
], [3238, 1067, 1067, 1067, 1067, 1066, 1066], { num: true, size: 17 }));
add(BOX('Lecture financière', `Le rendement du cas de base (TRI fonds propres de ${TRI}, TRI projet de ${TRIP} contre un coût moyen pondéré du capital d’environ ${pc(X.wacc, 0)}) est attractif, mais il dépend fortement du contrat. Le scénario ne doit pas être présenté comme une promesse de performance : il sert à fixer les seuils de négociation, les protections bancaires et les variables à auditer. Sans valeur résiduelle en fin d’horizon, le TRI fonds propres reste de 28,8 %.`));

// 11. Sensibilité
add(H1('Analyse de sensibilité et tests de résistance'));
add(H2('Résultats des scénarios'));
add(P('Chaque scénario est calculé **à profil de dette constant** (celui du cas de base), conformément à la pratique des prêteurs : seul l’impact du choc sur la capacité de remboursement est mesuré.'));
add(T(['Scénario', 'TRI FP', 'VAN M$', 'DSCR min.', 'EBITDA A2'], R.scen.map(([n, r]) => [n, r.tri === null ? 'n.s.' : pc(r.tri), fr(r.van, 2), xx(r.dscr), fr(r.ebitda2, 2)]).concat([['Cas bas avec protections contractuelles*', pc(X.prot.tri), fr(X.prot.van, 2), xx(X.prot.dscr), '–']]),
  [3838, 1450, 1450, 1450, 1450], { numCols: [1, 2, 3, 4], totals: [0], highlight: [8, 9] }));
add(SRC('Cas bas prêteur : tarif −5 %, volume −10 %, diesel +15 %, CAPEX +5 %. Stress sévère : tarif −10 %, volume −15 %, diesel +20 %, CAPEX +10 %. * Cas bas avec indexation carburant effective (diesel neutralisé).'));
add(FIG('fig7.png', 'Sensibilité du TRI fonds propres aux chocs unitaires', null));
add(H2('Points morts et seuils critiques'));
add(T(['Seuil', 'Valeur', 'Interprétation'], [
  ['Point mort d’exploitation (A2)', fr(R.seuil, 2) + ' Mt/an', 'Volume couvrant les coûts d’exploitation, avant dette et impôt'],
  ['Volume pour DSCR 1,30x', `${pc(vol130, 0)} du plan (${fr(3.9 * vol130, 2)} Mt en régime)`, 'Minimum de sécurité bancaire au tarif central'],
  ['Volume pour DSCR 1,00x', `${pc(vol100, 0)} du plan (${fr(3.9 * vol100, 2)} Mt en régime)`, 'Seuil de défaut de paiement hors réserves'],
  ['Tarif pour DSCR 1,30x', fr(R.be.tarif130, 2) + ' USD/t (A1)', 'Tarif plancher de négociation au volume central'],
  ['Tarif pour DSCR 1,00x', fr(R.be.tarif100, 2) + ' USD/t (A1)', 'Seuil de défaut de paiement hors réserves'],
  ['Diesel pour DSCR 1,30x', fr(1.37 * R.be.diesel130, 2) + ' USD/l', `Hausse tolérable de ${pc(R.be.diesel130 - 1, 0)} sans indexation`],
], [3000, 3000, 3638], { boldFirst: true }));
add(BOX('Enseignement pour la négociation', `Le projet est **plus sensible au tarif et au volume qu’au coût d’investissement**. Une baisse de 10 % du tarif ou de 15 % du volume ramène le DSCR sous 1,30x ; le stress combiné sévère ne permet pas de servir la dette (DSCR ${xx(S['Stress combiné sévère'].dscr)}). Le take-or-pay, l’indexation carburant et le compte séquestre ne sont donc pas des options : ce sont des **conditions de bancabilité**.`, { color: RED, fill: REDL }));
add(H2('Priorité des leviers'));
add(T(['Rang', 'Levier', 'Action'], [
  ['1', 'Tarif et indexation', 'Refuser un tarif non indexé ou inférieur au plancher bancaire'], ['2', 'Volume garanti', 'Introduire un take-or-pay et un calendrier mensuel'],
  ['3', 'Productivité réelle', 'Chronométrer le cycle avant tout engagement'], ['4', 'Prix d’acquisition', 'Mettre en concurrence au moins trois constructeurs'],
  ['5', 'Disponibilité mécanique', 'Garantie, pièces, techniciens constructeur et stock critique'], ['6', 'Délai de paiement', 'Compte séquestre ou garantie bancaire de paiement'],
], [900, 2800, 5938]));

// 12. Risque de crédit
add(H1('Analyse du risque de crédit'));
add(P('La grille ci-dessous reprend l’analyse des « 5 C » utilisée par les comités de crédit. Elle identifie les points forts du dossier et les conditions à remplir pour obtenir une décision favorable.'));
add(T(['Critère', 'Appréciation', 'Mesures d’atténuation'], [
  ['Capacité de remboursement', `DSCR de base ${xx(b.dscr_min_amort)}, LLCR ${xx(b.llcr_min)} ; sensibilité forte au tarif et au volume`, 'Take-or-pay, indexation, DSRA 6 mois, remboursement sculpté'],
  ['Caractère du sponsor', 'Promoteur local diversifié ; expérience en transport minier à démontrer', 'Management expérimenté, assistance technique constructeur, partenaire technique, KYC complet'],
  ['Capital', '35 % de fonds propres, apportés avant la dette', 'Garantie de dépassement de coûts pendant la mobilisation'],
  ['Collatéral', 'Flotte et équipements ≈ 18,5 M USD à neuf ; valeur de réalisation forcée estimée à 60–70 %, soit une couverture inférieure à l’encours initial', 'La sûreté principale est la cession de créances du contrat minier ; assurances déléguées'],
  ['Conditions de marché', 'Marché bauxitique en croissance mais concentré ; risque pays Guinée', 'Contrat en USD, clients de premier rang, assurance risque politique (ATI, MIGA)'],
], [2200, 3719, 3719], { boldFirst: true }));
add(BOX('Point de vigilance des prêteurs', 'Au démarrage, la valeur de revente de la flotte ne couvre pas à elle seule l’encours de la dette. Le crédit repose donc sur la **qualité du contrat et du client minier** (financement adossé aux flux). La solidité financière du donneur d’ordre, sa notation et son historique de paiement seront au cœur de la due diligence.', { color: BLUE, fill: LIGHT }));

// 13. Gouvernance & conformité
add(H1('Gouvernance, conformité et reporting'));
add(H2('Gouvernance d’entreprise'));
add(BUL('Conseil d’administration comprenant au moins un administrateur indépendant ; comité d’audit et des risques.'));
add(BUL('Séparation des fonctions d’engagement, de paiement et de contrôle ; double signature au-delà de seuils définis.'));
add(BUL('Politique d’achats avec mise en concurrence documentée et registre des conflits d’intérêts.'));
add(H2('Conformité et intégrité'));
add(T(['Domaine', 'Exigence'], [
  ['Connaissance du client (KYC)', 'Identification des actionnaires et bénéficiaires effectifs, vérification des personnes politiquement exposées'],
  ['Lutte contre le blanchiment (LBC/FT)', 'Procédures alignées sur les recommandations du GAFI et la réglementation BCRG'],
  ['Anti-corruption', 'Code de conduite, politique cadeaux et paiements de facilitation, clauses anti-corruption dans tous les contrats ; conformité aux conventions OCDE et ONU'],
  ['Sanctions internationales', 'Filtrage des contreparties et fournisseurs (listes ONU, UE, OFAC)'],
  ['Transparence', 'Déclaration des paiements au secteur extractif conforme aux exigences ITIE'],
], [3200, W - 3200], { boldFirst: true }));
add(H2('Reporting aux prêteurs et investisseurs'));
add(T(['Document', 'Fréquence', 'Délai'], [
  ['Rapport opérationnel (tonnages, disponibilité, HSE)', 'Mensuel', '15 jours'],
  ['Comptes de gestion et suivi budgétaire', 'Trimestriel', '45 jours'],
  ['Certificat de conformité aux ratios', 'Semestriel', '60 jours'],
  ['États financiers audités (IFRS et SYSCOHADA)', 'Annuel', '120 jours'],
  ['Rapport de suivi environnemental et social', 'Annuel', '90 jours'],
  ['Modèle financier mis à jour et budget', 'Annuel', 'Avant le début de l’exercice'],
], [4838, 2400, 2400], { boldFirst: true }));
add(P('Les comptes seront audités par un cabinet d’audit de réputation internationale. Un ingénieur indépendant des prêteurs pourra être mandaté pour valider les hypothèses techniques et suivre la mobilisation.'));

// 14. Mise en œuvre
add(H1('Plan de mise en œuvre'));
add(H2('Phasage'));
add(FIG('fig8.png', 'Calendrier indicatif de mobilisation sur 12 mois (année 0)', null));
add(H2('Jalons de décision'));
add(T(['Jalon', 'Livrable', 'Décision'], [
  ['J1 — Préfaisabilité', 'Étude de corridor, besoin client, modèle économique', 'Poursuivre ou arrêter'],
  ['J2 — Offre commerciale', 'Tarif, SLA, volume, projet de contrat', 'Entrer en négociation finale'],
  ['J3 — Due diligence', 'Juridique, fiscale, technique, E&S, contrepartie', 'Approuver l’investissement'],
  ['J4 — Bouclage financier', 'Documentation de crédit signée, conditions préalables levées', 'Commande ferme'],
  ['J5 — Réception usine', 'Inspection des véhicules', 'Autoriser l’expédition'],
  ['J6 — Pilote 20 camions', 'Productivité et sécurité validées', 'Monter à 100 camions'],
  ['J7 — Réception définitive', 'Performance sur 90 jours', 'Clôturer la mobilisation'],
], [2800, 4038, 2800], { boldFirst: true }));
add(H2('Plan des 100 premiers jours d’exploitation'));
add(T(['Période', 'Priorités'], [
  ['Jours 1–30', 'Stabiliser inspections, carburant, dispatch, pesage, maintenance et reporting'],
  ['Jours 31–60', 'Réduire les temps d’attente, analyser les pannes récurrentes, ajuster stocks et roster'],
  ['Jours 61–100', 'Valider les coûts unitaires, la performance contractuelle, le bonus/malus et le plan d’amélioration'],
], [2000, W - 2000], { boldFirst: true }));

// 15. Risques
add(H1('Risques stratégiques et mesures de maîtrise'));
const lvl = (t) => t;
add(T(['Risque', 'Impact', 'Probabilité', 'Mesure principale'], [
  ['Absence de contrat ferme', 'Très élevé', 'Élevée', 'Ne pas commander avant contrat et garanties'],
  ['Volume insuffisant', 'Élevé', 'Moyenne', 'Take-or-pay et diversification client'],
  ['Hausse du diesel', 'Très élevé', 'Élevée', 'Indexation ou fourniture par le client'],
  ['Dégradation de la route', 'Très élevé', 'Élevée', 'SLA route, audit et pénalités'],
  ['Retard de paiement', 'Très élevé', 'Moyenne', 'Séquestre, garantie bancaire, facturation fréquente'],
  ['Pannes et pièces', 'Élevé', 'Moyenne', 'Garantie constructeur, stock critique, atelier, MRA'],
  ['Accident grave', 'Critique', 'Moyenne', 'Système HSE, fatigue, vitesse, formation'],
  ['Risque de change', 'Élevé', 'Moyenne', 'Tarif en USD ou indexation FX'],
  ['Retard de mobilisation', 'Élevé', 'Moyenne', 'Garantie de dépassement des sponsors, pénalités constructeur'],
  ['Douane / importation', 'Moyen', 'Moyenne', 'Transitaire expérimenté et provision'],
  ['Conflit communautaire', 'Élevé', 'Moyenne', 'Engagement, recrutement local, gestion des plaintes'],
  ['Substitution par rail ou convoyeur', 'Élevé', 'Faible à moyenne', 'Durée ferme et diversification des services'],
  ['Résiliation anticipée', 'Critique', 'Faible à moyenne', 'Indemnité, reprise de dette et valeur résiduelle'],
  ['Risque pays et réglementaire', 'Élevé', 'Moyenne', 'Assurance risque politique, clauses de stabilisation'],
], [2800, 1300, 1500, 4038], { boldFirst: true }));
add(BOX('Principe de maîtrise', 'Tout risque non transférable au client doit être tarifé, assuré, couvert par une réserve ou absorbé par des fonds propres suffisants. Un risque ignoré ne disparaît pas : il se transforme en besoin de trésorerie.'));

// 16. Conclusion
add(H1('Conclusion et recommandation d’investissement'));
add(P('Le marché guinéen de la bauxite est suffisamment profond pour soutenir un opérateur de 100 camions. La croissance récente des volumes, la concentration des activités dans les corridors côtiers et les exigences de contenu local créent une opportunité réelle pour un prestataire local professionnel.'));
add(P(`Le projet est **recommandé sous conditions**. Le coût total de ${TOT} M USD ne doit être engagé qu’après validation simultanée du contrat, du corridor, des devis fournisseurs, de la fiscalité, du financement et du système HSE. Le cas de base produit un TRI fonds propres de ${TRI} et un DSCR minimum de ${DSCR}, mais le stress combiné sévère détruit de la valeur et ne couvre pas la dette.`));
add(T(['Décision', 'Condition'], [
  ['GO conditionnel', `Contrat ≥ 5 ans (idéalement 7), volume garanti ≥ ${fr(3.9 * vol130, 2)} Mt/an, tarif ≥ ${fr(R.be.tarif130, 2)} USD/t, indexation diesel, paiement sécurisé, fonds propres apportés`],
  ['NO-GO', 'Commande sans contrat, route non auditée, absence de garantie de paiement, dette sans différé'],
  ['Option prudente', 'Démarrage avec 20 à 30 camions et option ferme d’extension après validation du pilote'],
], [2400, W - 2400], { boldFirst: true }));
add(BOX('Recommandation finale', 'Négocier et financer le projet comme un **contrat d’infrastructure logistique**, et non comme un simple achat de véhicules. La valeur réside dans le droit contractuel de transporter un volume rentable pendant une durée suffisante, sécurisé par des flux cédés aux prêteurs.', { color: GREEN, fill: GREENL }));

// 17. Sources
add(H1('Sources documentaires'));
add(T(['Réf.', 'Institution', 'Document et lien'], [
  ['S1', 'Ministère des Mines et de la Géologie / ITIE-Guinée', 'Bulletin des Statistiques Minières et Carrières n°30, janvier–décembre 2025, publié le 11 avril 2026. itie-guinee.org'],
  ['S2', 'Initiative pour la Transparence dans les Industries Extractives', 'Fiche pays Guinée : contribution du secteur extractif à l’économie. eiti.org/countries/guinea'],
  ['S3', 'ITIE internationale', 'Guinea achieves a good score in EITI implementation, 15 mai 2026. eiti.org/news'],
  ['S4', 'Conseil National de la Transition', 'Loi L/2022/0010/CNT portant contenu local en République de Guinée. cnt.gov.gn'],
  ['S5', 'Ministère des Mines et de la Géologie / CPDM', 'Code minier 2011 amendé en 2013. cpdm.mines.gov.gn'],
  ['S6', 'Direction Générale des Impôts', 'Code général des impôts 2021, notamment article 229. dgi.gov.gn'],
  ['S7', 'Banque Centrale de la République de Guinée', 'Fixing du 31 juillet 2026 : USD/GNF 8 762,7692. bcrg-guinee.org'],
  ['S8', 'Equator Principles Association / IFC', 'Principes de l’Équateur IV (2020) ; Normes de performance environnementale et sociale de l’IFC (2012)'],
], [700, 3400, 5538], { boldFirst: true }));
add(BOX('Traçabilité', 'Les chiffres officiels repris dans le document renvoient à ces sources. Les hypothèses de projet sont identifiées comme telles et seront remplacées par les données issues du contrat et de la due diligence.', { color: BLUE, fill: LIGHT }));

// Annexes
add(H1u('Annexe A — Matrice complète des hypothèses'));
add(T(['Famille', 'Hypothèse', 'Statut'], [
  ['Marché', 'Bauxite, corridor Boké–Boffa', 'Choix stratégique'], ['Flotte', '100 camions (90 programmés)', 'Projet'], ['Charge utile', '45 t', 'À confirmer'],
  ['Distance aller', '35 km', 'À remplacer par la route réelle'], ['Rotations/jour', '3', 'À chronométrer'], ['Jours productifs', '321', 'Hypothèse'],
  ['Tarif initial', '7,10 USD/t', 'À négocier'], ['Indexation', '2 %/an + formule diesel/FX', 'À contractualiser'],
  ['Diesel', '1,37 USD/l, + 1,5 %/an', 'Référence budgétaire'], ['Consommation', '1,05 l/t', 'À mesurer'],
  ['Autres coûts variables', '1,46 USD/t, + 3 %/an', 'À valider par devis'], ['Coûts fixes', '4,80 M USD, + 3 %/an', 'Budget'],
  ['Enveloppe d’investissement', '30,0 M USD', 'Budget'], ['Coût total du projet', TOT + ' M USD', 'Calculé'],
  ['Dette', DEBT + ' M USD, 10,5 %, 7 ans', 'Hypothèse'], ['Profil de remboursement', 'Sculpté, DSCR cible ' + xx(b.dscr_cible), 'Hypothèse'],
  ['DSRA', '6 mois de service', 'Standard prêteurs'], ['BFR', '12 % du CA', 'Hypothèse'],
  ['IS', '25 %', 'Référence fiscale à confirmer'], ['Valeur résiduelle', '5,0 M USD en A7', 'Hypothèse prudente'],
], [2800, 3919, 2919], { boldFirst: true }));
add(H1u('Annexe B — Plan de due diligence avant financement'));
add(T(['Volet', 'Vérifications minimales'], [
  ['Commercial', 'Contrat, volumes, historique de paiement, solidité financière et notation du client'],
  ['Technique', 'Profil de route, pentes, revêtement, cycle, ponts, aires de retournement (ingénieur indépendant)'],
  ['Flotte', 'Devis, garantie, pièces, homologation, consommation, charge utile'],
  ['Financier', 'Audit du modèle financier par un tiers, fiscalité, change, BFR, assurances, dette'],
  ['Juridique', 'Société, permis, sûretés, résiliation, force majeure, litiges'],
  ['Environnemental et social', 'ESIA ou audit E&S, ESMS, ESAP, mécanisme de plainte, dépôt carburant'],
  ['Ressources humaines', 'Disponibilité des chauffeurs, salaires, roster, formation'],
  ['Intégrité', 'KYC/UBO, filtrage sanctions, vérification de réputation des sponsors et contreparties'],
], [2800, W - 2800], { boldFirst: true }));
add(H1u('Annexe C — États financiers prévisionnels détaillés'));
add(T(['M USD', ...Y7], [
  row("Chiffre d'affaires", b.ca, 3), row('Coûts variables', b.carb.map((v, i) => v + b.autres[i]), 3), row('Coûts fixes', b.fixes, 3),
  row('EBITDA', b.ebitda, 3), row('Amortissements', b.dep, 3), row('Intérêts', b.interets, 3), row('Impôt', b.impot, 3), row('Résultat net', b.rn, 3),
  row('CFADS', b.cfads, 3), row('Principal', b.princ, 3), row('Flux actionnaires', b.flux_act, 3), row('Encours de dette fin', b.dette_fin, 3),
], w8, { num: true, size: 16, totals: [3, 7, 10] }));
add(P(`Flux des actionnaires : apport de ${fr(b.equity, 3)} M USD en année 0, puis flux ci-dessus. TRI fonds propres ${TRI} ; TRI projet après impôt ${TRIP} ; VAN des fonds propres à 18 % : ${VAN} M USD.`));
add(H1u('Annexe D — Cahier des charges minimal des camions'));
add(T(['Élément', 'Exigence'], [
  ['Configuration', 'Camion-benne adapté au corridor et à la charge utile contractuelle'], ['Motorisation', 'Puissance et refroidissement adaptés à la pente, à la poussière et à la chaleur'],
  ['Freinage', 'Frein moteur / ralentisseur, ABS selon configuration, avertisseurs'], ['Benne', 'Volume compatible avec la densité de la bauxite, protection anti-usure, bâchage si requis'],
  ['Sécurité', 'ROPS/FOPS si exigé, caméras, alarme de recul, extincteurs, coupe-batterie'], ['Télématique', 'GPS, IVMS, lecture CAN, carburant, vitesse, géorepérage'],
  ['Maintenance', 'Garantie, formation, outillage, documentation et stock initial'], ['Confort', 'Climatisation, siège suspendu, réduction de la fatigue et visibilité'],
], [2400, W - 2400], { boldFirst: true }));
add(H1u('Annexe E — Note de révision : harmonisation des versions 1 et 2'));
add(P('La présente version consolide les deux versions antérieures du business plan et révise le modèle financier pour le rendre conforme aux attentes des prêteurs internationaux. Les principales modifications sont les suivantes :'));
add(T(['Point', 'Versions 1 et 2', 'Version 3.0 consolidée'], [
  ['Présentation des nombres', 'Mélange de points et virgules décimales (v1)', 'Format français homogène (virgule décimale)'],
  ['Part de marché', '« environ 2,2 % » et « 2,13 % » selon les pages', '2,1 % partout (3,90 / 182,83 Mt)'],
  ['Calendrier', 'Exploitation dès l’année 1, alors que la mobilisation dure 12 mois', 'Année 0 de mobilisation ajoutée, cohérente avec le planning'],
  ['Coût total', '30,0 M USD (investissements seuls)', `${TOT} M USD avec frais, intérêts intercalaires et DSRA`],
  ['Financement', 'Dette 19,5 M USD / fonds propres 10,5 M USD', `Dette ${DEBT} M USD / fonds propres ${EQ} M USD (65/35 maintenu)`],
  ['Remboursement', 'Linéaire sur 5 ans après 12 mois de grâce', 'Sculpté sur DSCR constant, 7 ans au total'],
  ['BFR et amortissements', 'BFR fixe ; valeur résiduelle et BFR agrégés (8,0 M USD)', 'BFR à 12 % du CA ; amortissement jusqu’à VR 5,0 M USD (IAS 16)'],
  ['Hypothèses implicites', 'Escalade des coûts variables non documentée', 'Diesel + 1,5 %/an et autres variables + 3 %/an explicités'],
  ['Normes et gouvernance', 'Non traitées', 'Term sheet, covenants, waterfall, IFC PS, Principes de l’Équateur, KYC/LBC, reporting'],
  ['Résultats clés', 'TRI 37,3 % ; VAN 8,45 M$ ; DSCR min. 1,44x', `TRI ${TRI} ; VAN ${VAN} M$ ; DSCR min. ${DSCR}`],
  ['Seuils de négociation', 'Tarif ≥ 6,93 USD/t ; volume ≥ 3,30 Mt', `Tarif ≥ ${fr(R.be.tarif130, 2)} USD/t ; volume ≥ ${fr(3.9 * vol130, 2)} Mt en régime`],
], [2200, 3419, 4019], { boldFirst: true, size: 17 }));
add(P('La baisse du TRI par rapport aux versions précédentes résulte principalement de l’intégration de la période de mobilisation (les fonds propres sont investis un an avant les premiers revenus) et des coûts de financement. Elle rend le dossier **plus crédible** auprès d’un comité de crédit, qui aurait de toute façon retraité ces éléments.'));
add(H1u('Annexe F — Glossaire financier'));
add(T(['Terme', 'Définition'], [
  ['CFADS', 'Cash Flow Available for Debt Service : flux de trésorerie disponible pour le service de la dette (EBITDA − impôts − CAPEX de maintien − variation du BFR)'],
  ['DSCR', 'Debt Service Coverage Ratio : CFADS de la période / service de la dette (intérêts + principal) de la période'],
  ['LLCR', 'Loan Life Coverage Ratio : valeur actualisée des CFADS sur la durée restante du prêt / encours de la dette'],
  ['DSRA', 'Debt Service Reserve Account : compte de réserve nanti couvrant 6 mois de service de la dette'],
  ['MRA', 'Maintenance Reserve Account : compte de réserve pour les révisions majeures de la flotte'],
  ['IDC', 'Interest During Construction : intérêts intercalaires de la période de mobilisation, capitalisés'],
  ['Take-or-pay', 'Engagement du client de payer un volume minimum, qu’il soit transporté ou non'],
  ['TRI / VAN', 'Taux de rendement interne / valeur actuelle nette des flux de trésorerie'],
  ['Remboursement sculpté', 'Échéancier dont le principal est ajusté aux flux prévus pour maintenir un DSCR constant'],
  ['Waterfall', 'Cascade contractuelle d’affectation des recettes du projet'],
  ['Step-in', 'Droit des prêteurs de se substituer à l’emprunteur défaillant dans l’exécution du contrat'],
], [2200, W - 2200], { boldFirst: true }));

// ---------- document ----------
const smallLogo = new ImageRun({ type: 'jpg', data: logo, transformation: { width: 64, height: 36 }, altText: { title: 'Logo', description: 'Logo', name: 'logo2' } });
const header = new Header({ children: [new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [2000, W - 2000], borders: NOB,
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 2000, type: WidthType.DXA }, borders: { top: nob, left: nob, right: nob, bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD } }, children: [new Paragraph({ children: [smallLogo] })] }),
    new TableCell({ width: { size: W - 2000, type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: { top: nob, left: nob, right: nob, bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD } }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 60 }, children: [run('Business plan — Transport minier, 100 camions — Guinée', { size: 16, color: MUTED })] })] }),
  ] })],
})] });
const footer = new Footer({ children: [new Paragraph({
  alignment: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 6 } },
  children: [run('Strictement confidentiel  ·  ALPHA B HANOK SARL  ·  Version 3.0 — septembre 2026  ·  Page ', { size: 16, color: MUTED }),
  new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: NAVY, bold: true }), run(' / ', { size: 16, color: MUTED }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: MUTED })],
})] });

const page = { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134, header: 567, footer: 567 } };
const doc = new Document({
  creator: 'ALPHA B HANOK SARL', title: 'Business plan — Transport minier 100 camions — Guinée', description: 'Dossier de financement, version 3.0',
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true, color: NAVY }, paragraph: { outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 25, bold: true, color: BLUE }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT, style: { run: { color: GOLD }, paragraph: { indent: { left: 500, hanging: 280 } } } }] },
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 320 } } } }] },
  ] },
  sections: [
    { properties: { page }, children: cover() },
    { properties: { page }, headers: { default: header }, footers: { default: footer }, children: C },
  ],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync('out.docx', buf); console.log('ok', C.length); });
