#!/usr/bin/env python3
"""Règles de lisibilité communes à tous les documents KTE (devis, ARC, BL, commandes fournisseurs, fiches
atelier, synthèses). Même logique que la fonction `lisible()` de l'éditeur en ligne (index.html).

- Jamais de retour à la ligne pour 1 ou 2 caractères : le dernier mot d'une ligne (≤ 6 caractères) reste
  collé au précédent ; « réf. », « rep. », « n° », « ép. », « Ø » restent collés à ce qui suit ; un nombre
  reste collé à son unité ou au mot qui suit (« 20 mm », « 10 rails ») ; pas de « – », « : » ni « / » en début
  de ligne.
- Surlignage jaune : texte entre ==…== dans une description ou une précision (ex. « ==hors trempe== »).

  from typo_kte import lisible, rl
  lisible(texte)   -> texte brut avec espaces insécables (Excel, mails…)
  rl(texte)        -> balisage ReportLab échappé + surlignage (Paragraph)
"""
import re

NB = " "
SURLIGNE = "#FFF200"
_ABR = r"(?:réf\.|Réf\.|rep\.|Rep\.|repère|n°|N°|ép\.|Ø|M|DN|Ra|art\.)"


def lisible(t):
    if t is None:
        return ""
    out = []
    for ligne in str(t).split("\n"):
        s = re.sub(r"[ \t]+", " ", ligne).strip()
        s = re.sub(r"(?<!\S)(%s) (?=\S)" % _ABR, r"\1" + NB, s)          # réf. X, rep. 4, ép. 20, Ø 12
        s = re.sub(r"(\d) (?=[^\s\d–—-][^\s]{0,8}(?:\s|$))", r"\1" + NB, s)  # 20 mm, 10 rails, 3 semaines
        s = re.sub(r" (?=[–—:;!?/×x]\s)", NB, s)                       # pas de « – » ou « : » en début de ligne
        m = re.search(r" (\S{1,6})$", s)                                  # dernier mot court collé
        if m:
            s = s[:m.start()] + NB + m.group(1)
        out.append(s)
    return "\n".join(out)


def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def rl(t, br="<br/>"):
    """Texte prêt pour reportlab.Paragraph : échappé, insécables, ==surlignage==, sauts de ligne."""
    s = _esc(lisible(t))
    s = re.sub(r"==(.+?)==", lambda m: f'<font backColor="{SURLIGNE}">{m.group(1)}</font>', s)
    return s.replace("\n", br)


def brut(t):
    """Texte sans marqueurs de surlignage (Excel, Axonaut, mails)."""
    return re.sub(r"==(.+?)==", r"\1", lisible(t))


if __name__ == "__main__":
    for x in ["Plaque alu R1, réf. Préhenseur R1 rep. 1", "Rail L440, réf. Préhenseurs R1/R2 rep. 4",
              "Matière : XC38 / XC48 – ==hors trempe==", "Matière : aluminium – ép. 20 mm"]:
        print(repr(rl(x)))
