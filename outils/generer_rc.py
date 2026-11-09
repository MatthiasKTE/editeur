#!/usr/bin/env python3
"""KTE – Rapports de contrôle (RC) Excel d'une affaire, avec la vignette de chaque pièce.

Un classeur RC_<affaire>.xlsx, une feuille par repère, construite à partir du modèle « RC Vierge.xlsx »
(mise en page, logo, colonnes de mesure, légende des instruments inchangés). En-tête pré-rempli :
OF (n° d'affaire), réf. pièce (rep. + désignation + réf.), client, quantité, matière, traitement.
La colonne « Instrument » (moyens de contrôle) est laissée VIDE : c'est au métrologue de la remplir.
La vignette de la pièce est posée en haut à droite (encadrée) pour éviter de contrôler la mauvaise pièce.

Usage :
  python3 generer_rc.py <FA_<affaire>.json | RC_<affaire>.json> "<RC Vierge.xlsx>" <sortie.xlsx> [--vignettes <dossier>]
                        [--fournisseurs <fournisseurs.json>]

Liste des fournisseurs (09/11/2026) : elle n'est plus dans le script. Elle est lue dans un JSON (liste de noms,
ou {"fournisseurs": [...]}) donné par --fournisseurs ; à défaut, « fournisseurs.json » à côté du script.
Le fichier de référence est « _Outils chiffrage/fournisseurs.json » sur SharePoint (jamais publié sur GitHub).
Sans liste, le script s'arrête : un RC client ne doit pas sortir sans le filtre des fournisseurs.

Entrée :
  - soit le JSON des fiches atelier (FA_<affaire>.json : affaire, client, vignettes, fiches[].lignes[]) ;
    les fiches « Achats » sont ignorées (sauf lignes avec "rc": true) ;
  - soit un JSON RC : {"affaire", "client", "commande_client"?, "vignettes"?, "lignes": [{"rep", "designation",
    "ref"?, "qte", "matiere"?, "traitement"?, "vignette"?, "intervenant"?}]}.
  Une ligne avec "rc": false est ignorée. Le dossier des vignettes est relatif au JSON.

RÈGLE (21/10/2026) : le RC est destiné au CLIENT → aucune trace des fournisseurs / sous-traitants
(pas de ligne FABRICATION, pas de n° de CF, pas de marque catalogue) et aucun caractère chinois/CJK
(y compris noms de polices et de styles hérités du modèle). Le nettoyage est appliqué à chaque génération ;
pour un RC existant : python3 generer_rc.py --nettoyer <RC.xlsx> [...] [--fournisseurs <fournisseurs.json>]
"""
import argparse, copy, io, json, os, re, shutil, sys, tempfile, zipfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage, ImageOps

# Fournisseurs / sous-traitants / marques catalogue à ne jamais faire apparaître sur un RC client :
# liste lue dans fournisseurs.json (cf. charger_fournisseurs), jamais écrite dans le script.
FOURNISSEURS = []
RE_FOURN = None


def charger_fournisseurs(chemin=None):
    """Charge la liste des fournisseurs (JSON : liste de noms ou {"fournisseurs": [...]}) et prépare RE_FOURN."""
    global FOURNISSEURS, RE_FOURN
    chemin = chemin or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fournisseurs.json')
    if not os.path.isfile(chemin):
        sys.exit(f"Liste des fournisseurs introuvable : {chemin}\n"
                 "Donner --fournisseurs \"<_Outils chiffrage>/fournisseurs.json\" (fichier SharePoint).")
    d = json.load(open(chemin, encoding='utf-8'))
    L = d.get('fournisseurs', []) if isinstance(d, dict) else d
    L = [str(x).strip() for x in L if str(x).strip()]
    if not L:
        sys.exit(f"Liste des fournisseurs vide : {chemin}")
    FOURNISSEURS = L
    RE_FOURN = re.compile(r'\b(' + '|'.join(re.escape(f) for f in FOURNISSEURS) + r')\b', re.I)
    return L


RE_CJK = re.compile(r'[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]')
RE_CF = re.compile(r'\(?\bCF-[\w-]+\)?')


def propre(t):
    """Texte affichable sur un RC client : sans fournisseur, sans n° de CF, sans CJK."""
    if t is None:
        return t
    t = str(t)
    if RE_FOURN.search(t):
        # segments « – » contenant un fournisseur (ex. « Réf. <fournisseur> 1005-14-23 ») supprimés
        seg = [x for x in re.split(r'\s+–\s+', t) if not RE_FOURN.search(x)]
        t = ' – '.join(seg) if seg else RE_FOURN.sub('', t)
    t = RE_CF.sub('', RE_CJK.sub('', t))
    return re.sub(r'\s{2,}', ' ', t).strip(' –-')


def matiere_propre(m):
    if m and RE_FOURN.search(str(m)):
        return 'Élément standard repris' if 'achat' in str(m).lower() else propre(m)
    return propre(m)


def nettoyer_textes(wb):
    """Nettoie les cellules texte de toutes les feuilles (en-têtes A4–A6 surtout)."""
    for ws in wb.worksheets:
        ws.title = RE_CJK.sub('', ws.title) or ws.title
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or v.startswith('='):
                    continue
                n = v
                if 'FABRICATION :' in n:
                    n = re.sub(r'\s*–\s*FABRICATION :.*$', '', n)
                if RE_FOURN.search(n) or RE_CJK.search(n) or RE_CF.search(n):
                    if n.startswith('QUANTITE'):
                        parts = re.split(r'\s+–\s+', n)
                        parts = [('MATIÈRE : ' + matiere_propre(x[10:])) if x.startswith('MATIÈRE : ') else propre(x) for x in parts]
                        n = '     –     '.join(x for x in parts if x)
                    else:
                        n = propre(n)
                if n != v:
                    c.value = n


def nettoyer_xlsx(chemin):
    """Supprime toute trace CJK des parties XML du classeur (polices 宋体 du modèle, style « 常规 »,
    polices est-asiatiques du thème) et les éventuels noms de fournisseurs restés dans les textes."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    with zipfile.ZipFile(chemin) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename.endswith('.xml') or it.filename.endswith('.rels'):
                x = data.decode('utf-8')
                if it.filename == 'xl/styles.xml':
                    x = re.sub(r'<name val="[^"]*[\u2e80-\u9fff\uac00-\ud7af\uff00-\uffef][^"]*"\s*/>', '<name val="Calibri"/>', x)
                    x = re.sub(r'<cellStyle name="[^"]*[\u2e80-\u9fff\uff00-\uffef][^"]*"', '<cellStyle name="Normal"', x)
                    x = re.sub(r'\s*<charset val="134"\s*/>', '', x)
                if 'theme' in it.filename:
                    x = re.sub(r'<a:font script="\w+" typeface="[^"]*"\s*/>', '', x)
                    x = re.sub(r'(<a:ea typeface=")[^"]*(")', r'\1\2', x)
                x = RE_CJK.sub('', x)
                data = x.encode('utf-8')
            zo.writestr(it, data)
    shutil.move(tmp, chemin)


TAILLE_VIGNETTE = 165  # px (≈ 4,4 cm) ; la ligne 1 de l'en-tête est agrandie pour la recevoir


def lignes_rc(d):
    if 'fiches' in d:
        out = []
        for f in d['fiches']:
            achats = str(f.get('metier', '')).lower().startswith('achat')
            for l in f.get('lignes', []):
                if l.get('rc') is False or (achats and not l.get('rc')):
                    continue
                out.append(dict(l))
        return out
    return [l for l in d.get('lignes', []) if l.get('rc') is not False]


def cle_rep(l):
    m = re.match(r'\s*(\d+)', str(l.get('rep', '')))
    return (int(m.group(1)) if m else 9999, str(l.get('rep', '')))


def nom_feuille(l, pris):
    m = re.match(r'\s*(\d+)(.*)', str(l.get('rep', '')))
    rep = f"Rep{int(m.group(1)):02d}{m.group(2).strip()}" if m else f"Rep{l.get('rep', '')}"
    base = re.sub(r'[\\/*?:\[\]]', '-', f"{rep} {l.get('designation', '')}")[:31].strip()
    n, i = base, 2
    while n.lower() in pris:
        suf = f" ({i})"; n = base[:31 - len(suf)] + suf; i += 1
    pris.add(n.lower()); return n


def vide(v):
    return v in (None, '', '—', '-')


def vignette_xl(chemin):
    im = PILImage.open(chemin).convert('RGB')
    im = ImageOps.contain(im, (TAILLE_VIGNETTE * 3, TAILLE_VIGNETTE * 3))
    cadre = PILImage.new('RGB', (im.width + 8, im.height + 8), (120, 120, 120))
    blanc = PILImage.new('RGB', (im.width + 4, im.height + 4), (255, 255, 255))
    blanc.paste(im, (2, 2)); cadre.paste(blanc, (2, 2))
    b = io.BytesIO(); cadre.save(b, 'PNG'); b.seek(0)
    x = XLImage(b)
    r = TAILLE_VIGNETTE / max(cadre.width, cadre.height)
    x.width, x.height = round(cadre.width * r), round(cadre.height * r)
    return x


def ancrer(ws, img, col, row, dx=0, dy=0):
    m = AnchorMarker(col=col, colOff=pixels_to_EMU(dx), row=row, rowOff=pixels_to_EMU(dy))
    img.anchor = OneCellAnchor(_from=m, ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
    ws.add_image(img)


def main():
    if len(sys.argv) > 2 and sys.argv[1] == '--nettoyer':
        ap = argparse.ArgumentParser()
        ap.add_argument('--nettoyer', nargs='+', required=True); ap.add_argument('--fournisseurs')
        a = ap.parse_args()
        charger_fournisseurs(a.fournisseurs)
        for f in a.nettoyer:
            wb = openpyxl.load_workbook(f); nettoyer_textes(wb); wb.save(f)
            nettoyer_xlsx(f); print('nettoyé :', f)
        return
    ap = argparse.ArgumentParser()
    ap.add_argument('json'); ap.add_argument('modele'); ap.add_argument('sortie')
    ap.add_argument('--vignettes')
    ap.add_argument('--fournisseurs', help='JSON de la liste des fournisseurs (défaut : fournisseurs.json à côté du script)')
    a = ap.parse_args()
    charger_fournisseurs(a.fournisseurs)
    d = json.load(open(a.json, encoding='utf-8'))
    base = os.path.dirname(os.path.abspath(a.json))
    dv = a.vignettes or os.path.join(base, d.get('vignettes', 'Vignettes'))
    L = sorted(lignes_rc(d), key=cle_rep)
    if not L:
        sys.exit('Aucune pièce à contrôler dans ' + a.json)

    wb = openpyxl.load_workbook(a.modele)
    mod = wb.worksheets[0]
    logos = [(i._data(), copy.deepcopy(i.anchor), i.width, i.height) for i in mod._images]
    pris, manquantes = set(), []
    for l in L:
        ws = wb.copy_worksheet(mod)
        ws.title = nom_feuille({**l, 'designation': propre(l.get('designation', ''))}, pris)
        # réglages d'impression du modèle (copy_worksheet ne les reprend pas tous)
        ws.page_setup.orientation = mod.page_setup.orientation
        ws.page_setup.paperSize = mod.page_setup.paperSize
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
        ws.sheet_properties.pageSetUpPr = copy.copy(mod.sheet_properties.pageSetUpPr)
        ws.page_margins = copy.copy(mod.page_margins)
        ws.print_options.horizontalCentered = mod.print_options.horizontalCentered
        for data, anc, w, h in logos:
            x = XLImage(io.BytesIO(data)); x.width, x.height = w, h; x.anchor = copy.deepcopy(anc); ws.add_image(x)
        rep = str(l.get('rep', '')).strip()
        ref = '' if vide(propre(l.get('ref'))) else f" – {propre(l['ref'])}"
        ws['A4'] = f"OF : {d.get('affaire', '')}" + ('' if vide(d.get('commande_client')) else f"   –   Cde client : {d['commande_client']}")
        ws['A5'] = f"REF. PIECE : Rep. {rep} – {propre(l.get('designation', ''))}{ref}"
        ws['P5'] = f"  {d.get('client', '')}"
        q = [f"QUANTITE : {l.get('qte', '')}"]
        if not vide(l.get('matiere')): q.append(f"MATIÈRE : {matiere_propre(l['matiere'])}")
        if not vide(l.get('traitement')): q.append(f"TRAITEMENT : {propre(l['traitement'])}")
        # pas de ligne FABRICATION / intervenant : le RC part chez le client (aucun fournisseur visible)
        ws['A6'] = '     –     '.join(q)
        # colonne H « Instrument » (moyens de contrôle) laissée vide : c'est au métrologue de la renseigner
        for r in range(9, ws.max_row + 1):
            if str(ws.cell(r, 1).value or '').strip().lower().startswith('contr'):
                break
            ws.cell(r, 8).value = None
        for c in ('A5', 'A6'):
            al = copy.copy(ws[c].alignment); al.shrink_to_fit, al.wrap_text = True, False; ws[c].alignment = al
        v = l.get('vignette')
        chemin = os.path.join(dv, v) if v else None
        if chemin and os.path.isfile(chemin):
            ws.row_dimensions[1].height = 100  # 56 pt dans le modèle
            ancrer(ws, vignette_xl(chemin), col=15, row=0, dx=30, dy=4)  # colonnes P à S, lignes 1 à 3
        else:
            manquantes.append(f"Rep. {rep} ({v or 'pas de vignette indiquée'})")
            ws['Q2'] = 'VIGNETTE MANQUANTE'
            ws['Q2'].font = openpyxl.styles.Font(b=True, color='C00000', sz=9)
    wb.remove(mod)
    wb.active = 0
    nettoyer_textes(wb)
    wb.save(a.sortie)
    nettoyer_xlsx(a.sortie)
    print(f"{a.sortie} : {len(L)} feuille(s) – " + ', '.join(ws.title for ws in wb.worksheets))
    if manquantes:
        print('ATTENTION vignettes manquantes : ' + '; '.join(manquantes))


if __name__ == '__main__':
    main()
