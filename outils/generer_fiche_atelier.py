#!/usr/bin/env python3
"""KTE – Fiches atelier (Fraisage, Tournage, Fraisage Fil) et fiche d'achats des éléments standards.

Usage :
    python3 generer_fiche_atelier.py FA_<affaire>.json logo_kte.png <dossier de sortie>

Le JSON décrit l'affaire et une ou plusieurs fiches :
{
  "affaire": "D26xxxx", "client": "Client (site)", "edition": "15/10/2026",
  "delai_atelier": "10/11/2026", "expedition": "13/11/2026", "referent": "Prénom Nom",
  "dossier_nom": "D26xxxx <libellé>", "dossier_url": "https://<sharepoint>/…",
  "vignettes": "chemin/du/dossier/Vignettes",          # PNG carrés, nom libre (champ "vignette" des lignes)
  "fiches": [
    {"metier": "Fraisage",                              # Fraisage | Tournage | Fraisage Fil | Contrôle
     "lignes": [{"vignette": "D26xxxx_Rep01_Plaque_base.png", "rep": "1", "designation": "Plaque base",
                 "ref": "Réf. <n° de plan>", "qte": 1, "matiere": "C45E", "traitement": "—",
                 "dimensions": "395 × 370 × 20"}],
     "gravure": {"oui": true, "texte": "N° de rep. + D26xxxx", "procede": "Laser"},   # ou {"oui": false}
     "observations": ["…"]},
    {"metier": "Achats",                                # fiche d'achats des éléments standards
     "reception": "03/11/2026", "referent_achat": "Prénom Nom",
     "lignes": [{"rep": "9", "designation": "Colonne", "ref": "601-14-80", "fournisseur": "<fournisseur>",
                 "dimensions": "Ø14 × 80", "qte": 2, "pour": "Montage"}],
     "observations": ["…"]}
  ]
}
Sortie : FA_<affaire>_<Métier>.pdf (FA_<affaire>_Achats_standards.pdf pour les achats).

Règles de présentation (validées le 15/10/2026) :
- n° d'affaire seul (sans le mot « Affaire »), gras 22 pt, sous le titre ;
- une couleur par métier, uniquement sur le titre et l'en-tête du tableau (pas d'aplat ni de bande : économie d'encre) ;
- vignette 24 mm par repère (colonne Visuel) pour les fiches de fabrication ;
- ligne Gravure (21/10/2026) sur toutes les fiches de fabrication : Non / Oui, et si Oui le texte à graver et le
  procédé (Machine / Laser) ; champ "gravure" absent = cases vides à cocher à la main ;
- document interne sans prix ;
- fiche sur plusieurs pages (22/10/2026) : flèche « Suite au verso » en bas de chaque page sauf la dernière.
Mise en page du 28/10/2026 :
- plus de ligne d'en-tête « KTE – Fiche de suivi… / Page n » ; « Page x/N » en bas au centre de chaque page ;
- titre = le métier seul (Fraisage, Tournage, Fraisage Fil, Contrôle, Éléments standards), en couleur ;
- client en gras 15 pt, sans la ville / le site entre parenthèses ; plus de ligne « Fiche FA-… – éditée le … » ;
- Expédition client en 15 pt gras noir ; plus de ligne « Dossier / plans » ;
- cadre « Observations » toujours présent, vide pour une remarque manuelle ; une observation
  du JSON (rare) y est écrite en 13 pt ;
- fiches de fabrication (sauf Contrôle) : en bas de la page 1, tableaux Appro matière (2 lignes) et Appro outillage
  (1 ligne), colonnes Fournisseur / ARC (croix à réception de l'ARC) / Date de livraison, remplis à la main par l'opérateur.
"""
import glob, json, os, re, sys
from typo_kte import rl, brut  # règles de lisibilité communes (insécables, ==surlignage==)
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping
from reportlab.platypus import KeepTogether, BaseDocTemplate, PageTemplate, Frame, Table, TableStyle, Paragraph, Spacer, Image, Flowable
from reportlab.lib.styles import ParagraphStyle

COUL = {'Fraisage': '#1F5FA8',       # bleu
        'Tournage': '#2E7D32',       # vert
        'Fraisage Fil': '#D96C00',   # orange
        'Contrôle': '#00838F',       # bleu canard
        'Achats': '#6A3D9A'}         # violet
GRIS = colors.HexColor('#6B6B6B'); TRAIT = colors.HexColor('#D0D0D0'); FOND = colors.HexColor('#F0F3F8')


def _polices():
    def f(nom):
        l = [p for p in glob.glob('/usr/share/fonts/**/*.ttf', recursive=True) + glob.glob('/Library/Fonts/*.ttf')
             + glob.glob(os.path.expanduser('~/Library/Fonts/*.ttf')) if p.endswith(nom)]
        return l[0] if l else None
    r, b = f('Carlito-Regular.ttf'), f('Carlito-Bold.ttf')
    if r and b:
        pdfmetrics.registerFont(TTFont('Carlito', r)); pdfmetrics.registerFont(TTFont('Carlito-Bold', b))
        for i, it in ((0, 0), (1, 0), (0, 1), (1, 1)):
            addMapping('Carlito', i, it, 'Carlito-Bold' if i else 'Carlito')
        return 'Carlito', 'Carlito-Bold'
    return 'Helvetica', 'Helvetica-Bold'


FR, FB = _polices()


def st(**k):
    b = dict(fontName=FR, fontSize=10, leading=12.5); b.update(k); return ParagraphStyle('s', **b)


N = st(); B = st(fontName=FB); P = st(fontSize=8.5, leading=10.5)


def _pied(aff, total=None, appro=False, K=None):
    """Pied de page (28/10/2026) : plus de ligne d'en-tête ; « Page x/N » en bas au centre, flèche « Suite au verso »
    en bas à droite sauf sur la dernière page, et tableaux Appro matière / Appro outillage en bas de la page 1."""
    def f(c, doc):
        c.saveState(); c.setFont(FR, 9); c.setFillColor(GRIS)
        c.drawCentredString(105*mm, 7*mm, f'Page {doc.page}' + (f'/{total}' if total else ''))
        c.restoreState()
        if total and doc.page < total:
            _verso(c)
        if appro and doc.page == 1:
            t1, t2 = _appro(K)
            w2, h2 = t2.wrapOn(c, 182*mm, 60*mm); w1, h1 = t1.wrapOn(c, 182*mm, 60*mm)
            t2.drawOn(c, 14*mm, 14*mm); t1.drawOn(c, 14*mm, 14*mm + h2 + 3*mm)
    return f


BAS_APPRO = 50*mm   # place réservée en bas de la page 1 pour les tableaux d'appro


def _appro(K):
    """Appro matière (2 lignes) et Appro outillage (1 ligne) : fournisseur, ARC (case étroite cochée à réception de l'ARC
    de la commande fournisseur, 05/11/2026) et date de livraison, remplis à la main."""
    L = st(fontName=FB, fontSize=12, leading=14, textColor=K); H = st(fontName=FB, fontSize=9.5, leading=11, alignment=1)
    def tab(titre, n):
        rows = [[Paragraph(titre, L), Paragraph('Fournisseur', H), Paragraph('ARC', H),
                 Paragraph('Date de livraison', H)]] + [['', '', '', '']] * n
        t = Table(rows, colWidths=[36*mm, 84*mm, 12*mm, 50*mm], rowHeights=[7*mm] + [10*mm] * n)
        t.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), .6, colors.HexColor('#9A9A9A')),
                               ('SPAN', (0, 0), (0, n)), ('BACKGROUND', (1, 0), (-1, 0), FOND),
                               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
        return t
    return tab('Appro matière', 2), tab('Appro outillage', 1)


def _verso(c):
    """Flèche en bas de page : la fiche continue au verso (règle du 22/10/2026)."""
    c.saveState()
    x, y = 194*mm, 6*mm
    c.setFillColor(colors.black); c.setStrokeColor(colors.black)
    c.setFont(FB, 11); c.drawRightString(x - 14*mm, y + 1.5*mm, 'Suite au verso')
    c.setLineWidth(2.2); c.line(x - 12*mm, y + 3*mm, x - 3*mm, y + 3*mm)
    pth = c.beginPath(); pth.moveTo(x, y + 3*mm); pth.lineTo(x - 4.5*mm, y + 6*mm); pth.lineTo(x - 4.5*mm, y); pth.close()
    c.drawPath(pth, fill=1, stroke=0)
    c.restoreState()


def _doc(o, titre, appro):
    W, H = A4; haut, bas = 12*mm, 14*mm
    d = BaseDocTemplate(o, pagesize=A4, leftMargin=14*mm, rightMargin=14*mm, topMargin=haut, bottomMargin=bas,
                        title=titre, author='KTE')
    r1 = BAS_APPRO if appro else 0
    d.addPageTemplates([
        PageTemplate('p1', [Frame(14*mm, bas + r1, W - 28*mm, H - haut - bas - r1, id='f1')], autoNextPageTemplate='p2'),
        PageTemplate('p2', [Frame(14*mm, bas, W - 28*mm, H - haut - bas, id='f2')])])
    return d


def _construire(out, titre, story, aff, appro=False, K=None):
    """Deux passes : nombre de pages, puis « Page x/N » et flèche « Suite au verso » sur toutes les pages sauf la dernière."""
    import io
    tmp = _doc(io.BytesIO(), titre, appro); e = _pied(aff, None, appro, K)
    for pt in tmp.pageTemplates: pt.onPage = e
    tmp.build(story()); n = tmp.page
    doc = _doc(out, titre, appro); e = _pied(aff, n, appro, K)
    for pt in doc.pageTemplates: pt.onPage = e
    doc.build(story())
    return n


def _client(c):
    """Nom du client sans la ville / le site entre parenthèses (28/10/2026)."""
    c = re.split(r'\s+[–—-]\s+', c or '')[0]          # « Client – commande 12345 … » → « Client »
    return re.sub(r'\s*\([^)]*\)', '', c).strip()


def _haut(a, logo, titre, K):
    """En-tête (28/10/2026) : logo | métier seul en couleur, n° d'affaire, client plus grand (sans ligne « Fiche FA-… »)."""
    T = st(fontName=FB, fontSize=24, leading=28, textColor=K, alignment=2)
    A = st(fontName=FB, fontSize=22, leading=26, alignment=2)
    C = st(fontName=FB, fontSize=15, leading=18, alignment=2)
    t = Table([[Image(logo, 24*mm, 24*mm),
                [Paragraph(titre, T), Paragraph(a['affaire'], A), Paragraph(rl(_client(a['client'])), C)]]],
              colWidths=[40*mm, 142*mm])
    t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    return t


def _info(lignes):
    t = Table([[Paragraph(k, B), v] for k, v in lignes], colWidths=[40*mm, 142*mm])
    t.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), .5, TRAIT), ('BACKGROUND', (0, 0), (0, -1), FOND),
                           ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('TOPPADDING', (0, 0), (-1, -1), 4),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 4)]))
    return t


class CadreObs(Flowable):
    """Cadre Observations vide : 24 mm de haut, réduit (14 mm mini) s'il reste moins de place en bas de page,
    pour ne pas faire passer la fiche sur deux pages à cause du seul cadre."""
    TITRE = "Observations"
    def __init__(self, hmax=24*mm, hmin=14*mm): super().__init__(); self.hmax, self.hmin = hmax, hmin; self.hAlign = "CENTER"
    def wrap(self, aW, aH):
        self.width = 182*mm; dispo = aH - 1
        self.height = self.hmax if dispo >= self.hmax else (dispo if dispo >= self.hmin else self.hmax)
        return self.width, self.height
    def draw(self):
        c = self.canv; w, h, t = self.width, self.height, 6.5*mm
        c.saveState(); c.setFillColor(FOND); c.rect(0, h - t, w, t, stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor('#9A9A9A')); c.setLineWidth(.6); c.rect(0, 0, w, h)
        c.setStrokeColor(TRAIT); c.setLineWidth(.5); c.line(0, h - t, w, h - t)
        c.setFillColor(colors.black); c.setFont(FB, 10); c.drawString(6, h - t + 2*mm, self.TITRE); c.restoreState()


def _observations(obs):
    """Cadre Observations toujours présent (28/10/2026) : réservé à une remarque manuelle du référent de l'affaire.
    Aucune observation automatique ; si le JSON en contient une, elle est écrite plus grand (13 pt)."""
    if not obs:
        return CadreObs()
    O = st(fontSize=13, leading=16.5)
    t = Table([[Paragraph(CadreObs.TITRE, B)], [[Paragraph(rl(o), O) for o in obs]]], colWidths=[182*mm])
    t.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), .6, colors.HexColor('#9A9A9A')),
                           ('LINEBELOW', (0, 0), (-1, 0), .5, TRAIT), ('BACKGROUND', (0, 0), (-1, 0), FOND),
                           ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('TOPPADDING', (0, 1), (-1, 1), 5),
                           ('BOTTOMPADDING', (0, 1), (-1, 1), 6)]))
    return KeepTogether([t])


class Case(Flowable):
    """Case à cocher 3,6 mm (croix si cochée)."""
    def __init__(self, coche): super().__init__(); self.coche = coche; self.width = self.height = 3.6*mm
    def draw(self):
        c = self.canv; c.setLineWidth(.8); c.rect(0, 0, self.width, self.height)
        if self.coche:
            c.setLineWidth(1.3); c.line(.6*mm, .6*mm, 3*mm, 3*mm); c.line(.6*mm, 3*mm, 3*mm, .6*mm)


def _gravure(g, K):
    """Ligne Gravure : Non / Oui (+ texte à graver) puis procédé Machine / Laser.
    g = {"oui": bool, "texte": "...", "procede": "Machine" | "Laser"} ; None = cases vides (à remplir à la main)."""
    oui = bool(g.get('oui')) if isinstance(g, dict) else (bool(g) if g is not None else None)
    txt = (g.get('texte') or g.get('precision') or '').strip() if isinstance(g, dict) else ''
    pro = (g.get('procede') or '').strip().lower() if isinstance(g, dict) else ''
    rens = g is not None
    row = [Paragraph('Gravure :', B), Case(rens and oui is False), Paragraph('Non', N), Case(bool(oui)), Paragraph('Oui', N),
           Paragraph(f'<b>{rl(txt)}</b>' if oui and txt else ('' if rens else '<font color="#6B6B6B">À graver : ……………………………</font>'), N),
           Paragraph('Procédé :', B), Case(bool(oui) and pro.startswith('mach')), Paragraph('Machine', N),
           Case(bool(oui) and pro.startswith('las')), Paragraph('Laser', N)]
    t = Table([row], colWidths=[17*mm, 5*mm, 9*mm, 5*mm, 9*mm, 72*mm, 17*mm, 5*mm, 16*mm, 5*mm, 22*mm])
    t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('LEFTPADDING', (0, 0), (-1, -1), 2),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 2), ('TOPPADDING', (0, 0), (-1, -1), 5),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                           ('BOX', (0, 0), (-1, -1), 1.2 if oui else .8, K if oui else TRAIT)]))
    return t


def _rouge(d): return Paragraph(f'<font color="#B0002A" size="15"><b>{d}</b></font>', N)
def _noir(d): return Paragraph(f'<font size="15"><b>{d}</b></font>', N)


def fabrication(a, f, logo, out):
    m = f['metier']; K = colors.HexColor(COUL[m]); appro = m != 'Contrôle'
    info = _info([('Délai atelier', _rouge(a['delai_atelier'])), ('Expédition client', _noir(a['expedition'])),
                  ('Référent affaire', Paragraph(a['referent'], N))])
    H = st(fontName=FB, textColor=colors.white)
    rows = [[Paragraph(x, H) for x in ['Visuel', 'Rep.', 'Désignation / réf.', 'Qté', 'Matière', 'Traitement', 'Dimensions finies']]]
    tot = 0
    for l in f['lignes']:
        tot += l['qte']
        v = os.path.join(a.get('vignettes', ''), l.get('vignette', ''))
        rows.append([Image(v, 24*mm, 24*mm) if l.get('vignette') and os.path.exists(v) else '',
                     Paragraph(str(l['rep']), N), [Paragraph(rl(l['designation']), B), Paragraph(rl(l.get('ref', '')), P)],
                     Paragraph(f"<b>{l['qte']}</b>", N), Paragraph(rl(l.get('matiere', '')), P),
                     Paragraph(rl(l.get('traitement', '—')), P), Paragraph(rl(l.get('dimensions', '')), P)])
    rows.append(['', '', Paragraph('Total pièces', B), Paragraph(f'<b>{tot}</b>', N), '', '', ''])
    t = Table(rows, colWidths=[27*mm, 12*mm, 52*mm, 12*mm, 27*mm, 24*mm, 28*mm], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), K), ('GRID', (0, 0), (-1, -1), .5, TRAIT),
                           ('BACKGROUND', (0, -1), (-1, -1), FOND), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                           ('ALIGN', (0, 1), (0, -2), 'CENTER'), ('TOPPADDING', (0, 1), (-1, -2), 3),
                           ('BOTTOMPADDING', (0, 1), (-1, -2), 3)]))
    obs = _observations(f.get('observations'))
    def story():
        import copy as _c
        return [_haut(a, logo, m, K), Spacer(1, 6*mm), _c.deepcopy(info), Spacer(1, 4*mm), _gravure(f.get('gravure'), K),
                Spacer(1, 5*mm), _c.deepcopy(t), Spacer(1, 4*mm), _c.deepcopy(obs)]
    _construire(out, f"Fiche atelier {m} {a['affaire']}", story, a['affaire'], appro, K)


def achats(a, f, logo, out):
    K = colors.HexColor(COUL['Achats'])
    info = _info([('Réception au plus tard', _rouge(f['reception'])), ('Délai atelier', Paragraph(a['delai_atelier'], B)),
                  ('Expédition client', _noir(a['expedition'])),
                  ('Référent achat', Paragraph(f.get('referent_achat', a['referent']), N))])
    H = st(fontName=FB, textColor=colors.white, fontSize=9.5, leading=11)
    rows = [[Paragraph(x, H) for x in ['Rep.', 'Désignation', 'Réf. fournisseur', 'Fournisseur', 'Dimensions', 'Qté', 'Pour',
                                       'Commandé le', 'Reçu le']]]
    tot = 0
    for l in f['lignes']:
        tot += l['qte']
        rows.append([Paragraph(str(l['rep']), N), Paragraph(rl(l['designation']), B), Paragraph(rl(l.get('ref', '—')), N),
                     Paragraph(l.get('fournisseur', ''), P), Paragraph(l.get('dimensions', ''), P),
                     Paragraph(f"<b>{l['qte']}</b>", N), Paragraph(l.get('pour', ''), P), '', ''])
    rows.append(['', Paragraph('Total articles', B), '', '', '', Paragraph(f'<b>{tot}</b>', N), '', '', ''])
    t = Table(rows, colWidths=[11*mm, 25*mm, 21*mm, 21*mm, 24*mm, 10*mm, 26*mm, 22*mm, 22*mm], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), K), ('GRID', (0, 0), (-1, -1), .5, TRAIT),
                           ('BACKGROUND', (0, -1), (-1, -1), FOND), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                           ('TOPPADDING', (0, 1), (-1, -2), 7), ('BOTTOMPADDING', (0, 1), (-1, -2), 7)]))
    obs = _observations(f.get('observations'))
    def story():
        import copy as _c
        return [_haut(a, logo, 'Éléments standards', K), Spacer(1, 6*mm), _c.deepcopy(info),
                Spacer(1, 7*mm), _c.deepcopy(t), Spacer(1, 6*mm), _c.deepcopy(obs)]
    _construire(out, f"Fiche d'achats standards {a['affaire']}", story, a['affaire'])


def main():
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    a = json.load(open(sys.argv[1], encoding='utf-8')); logo = sys.argv[2]; dos = sys.argv[3]
    os.makedirs(dos, exist_ok=True)
    for f in a['fiches']:
        # JSON écrits par Akteev : délais / référent parfois dans la fiche plutôt qu'au niveau de l'affaire
        a = {**a, **{k: f[k] for k in ('delai_atelier', 'expedition', 'referent', 'client') if f.get(k) and not a.get(k)}}
        if f['metier'] == 'Achats':
            out = os.path.join(dos, f"FA_{a['affaire']}_Achats_standards.pdf"); achats(a, f, logo, out)
        else:
            out = os.path.join(dos, f"FA_{a['affaire']}_{f['metier'].replace(' ', '_').replace('ô', 'o')}.pdf"); fabrication(a, f, logo, out)
        print(out)


if __name__ == '__main__':
    main()
