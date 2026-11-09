#!/usr/bin/env python3
"""KTE – Plans de contrôle bullés + cotes pour les RC (validé 22/10/2026).

Usage :
  python3 plans_controle.py <plans.json> <dossier sortie>
plans.json = {"affaire": "D26xxxx", "plans": {"<réf>": {"pdf": "<plan client .pdf>", "drop": [...], "set": {"<texte>": {...}}}}}
Sortie : PC_<affaire>_<réf>.pdf (plan client avec bulles rouges numérotées) et cotes_<réf>.json
(n, txt, sym, dim, sup, inf, note) pour remplir_rc.py.
Lecture de la couche texte des PDF SolidWorks : cotes ±, tolérances empilées (+0,2 / 0), ajustements H7/J7,
taraudages (M, G), perçages/lamages (drilled, counterbored), rayons, chanfreins, tolérances géométriques,
tableaux de trous (REPERE / X / Y / Tol. / Holes). Cotes sans tolérance : ISO 2768-m.
Rugosités Ra non reprises. TOUJOURS contrôler visuellement chaque plan bullé (cotes oubliées / parasites) et
corriger par "drop" / "set" dans plans.json.
Puis : python3 remplir_rc.py RC_<affaire>.xlsx <dossier des cotes_*.json>  (après generer_rc.py).
"""
import sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bulles import traiter
cfg = json.load(open(sys.argv[1], encoding='utf-8')); out = sys.argv[2]; os.makedirs(out, exist_ok=True)
base = os.path.dirname(os.path.abspath(sys.argv[1]))
for ref, c in cfg['plans'].items():
    src = c['pdf'] if os.path.isabs(c['pdf']) else os.path.join(base, c['pdf'])
    k = traiter(src, os.path.join(out, f"PC_{cfg['affaire']}_{ref}.pdf"), os.path.join(out, f'cotes_{ref}.json'),
                {'drop': c.get('drop', []), 'set': c.get('set', {})})
    print(ref, len(k), 'cotes')
