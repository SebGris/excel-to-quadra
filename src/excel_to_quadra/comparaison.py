# -*- coding: utf-8 -*-
"""Comparaison entre la génération courante et une version de référence.

On compare les écritures de type M, identifiées par la clé
(dossier, compte, sens, libellé) — le dossier vient du nom de fichier
``NNN_ecriture_Quadra*.txt`` — et on classe chaque différence de montant en
NOUVELLE / SUPPRIMEE / MONTANT_MODIFIE. Les lignes I ne sont pas comparées dans
cette version. Le rapport est écrit en CSV (séparateur « ; », UTF-8 BOM).
"""

import csv
import glob
import os
from collections import defaultdict, namedtuple
from decimal import Decimal

from .moteur import MOTIF_SORTIE

#: Une différence d'écriture. Montants en centimes ; avant/apres = None si absent.
Difference = namedtuple("Difference",
                        "type dossier compte sens libelle avant apres ecart")

#: Synthèse globale de la comparaison (compteurs + totaux en centimes).
Synthese = namedtuple("Synthese",
                      "dossiers nouvelles supprimees modifiees "
                      "total_avant total_apres ecart")


def lire_ecritures_m(dossier: str) -> dict:
    """Lit les `*_ecriture_Quadra*.txt` d'un dossier et renvoie un dict
    (dossier, compte, sens, libellé) -> montant en centimes (lignes M cumulées).
    """
    ecritures: dict = defaultdict(int)
    if not os.path.isdir(dossier):
        return {}
    for chemin in glob.glob(os.path.join(dossier, MOTIF_SORTIE)):
        code = os.path.basename(chemin).split("_ecriture_Quadra")[0]
        contenu = open(chemin, "rb").read().decode("cp1252")
        for ligne in contenu.split("\r\n"):
            if not ligne.startswith("M"):
                continue
            cle = (code, ligne[1:9], ligne[41], ligne[21:41])
            ecritures[cle] += int(ligne[42:55])
    return dict(ecritures)


def comparer(reference: dict, courant: dict) -> list:
    """Renvoie la liste des Difference entre référence et courant.

    Les écritures identiques (même clé, même montant) ne sont pas listées.
    """
    diffs = []
    for cle in set(reference) | set(courant):
        avant = reference.get(cle)
        apres = courant.get(cle)
        if avant is None:
            diffs.append(Difference("NOUVELLE", *cle, None, apres, apres))
        elif apres is None:
            diffs.append(Difference("SUPPRIMEE", *cle, avant, None, -avant))
        elif avant != apres:
            diffs.append(Difference("MONTANT_MODIFIE", *cle, avant, apres, apres - avant))
    return diffs


def synthetiser(diffs: list, reference: dict, courant: dict) -> Synthese:
    """Construit la synthèse globale (compteurs + totaux avant/après/écart)."""
    total_avant = sum(reference.values())
    total_apres = sum(courant.values())
    return Synthese(
        dossiers=len({d.dossier for d in diffs}),
        nouvelles=sum(1 for d in diffs if d.type == "NOUVELLE"),
        supprimees=sum(1 for d in diffs if d.type == "SUPPRIMEE"),
        modifiees=sum(1 for d in diffs if d.type == "MONTANT_MODIFIE"),
        total_avant=total_avant, total_apres=total_apres,
        ecart=total_apres - total_avant,
    )


def _euros(centimes) -> str:
    """Centimes -> euros « 1234,56 » (virgule décimale, Excel FR) ; None -> «»."""
    if centimes is None:
        return ""
    return f"{Decimal(centimes) / 100:.2f}".replace(".", ",")


def ecrire_rapport_csv(diffs: list, chemin: str, synthese: Synthese) -> None:
    """Écrit le rapport CSV : différences triées par dossier puis type, puis une
    section récapitulative. Séparateur « ; », encodage UTF-8 BOM (Excel FR)."""
    diffs = sorted(diffs, key=lambda d: (d.dossier, d.type, d.compte, d.sens))
    with open(chemin, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["Type", "Dossier", "Compte", "Sens", "Libelle",
                    "Montant_avant", "Montant_apres", "Ecart"])
        for d in diffs:
            w.writerow([d.type, d.dossier, d.compte, d.sens, d.libelle.strip(),
                        _euros(d.avant), _euros(d.apres), _euros(d.ecart)])
        w.writerow([])
        w.writerow(["RECAPITULATIF"])
        w.writerow(["Dossiers touches", synthese.dossiers])
        w.writerow(["Ecritures nouvelles", synthese.nouvelles])
        w.writerow(["Ecritures supprimees", synthese.supprimees])
        w.writerow(["Ecritures modifiees", synthese.modifiees])
        w.writerow(["Total avant", _euros(synthese.total_avant)])
        w.writerow(["Total apres", _euros(synthese.total_apres)])
        w.writerow(["Ecart global", _euros(synthese.ecart)])


def comparer_dossiers(dossier_reference: str, dossier_courant: str,
                      chemin_csv: str):
    """Compare référence vs courant et écrit le CSV. Renvoie la Synthese, ou
    None si la référence est absente/vide (pas de diff, non bloquant)."""
    reference = lire_ecritures_m(dossier_reference)
    if not reference:
        return None
    courant = lire_ecritures_m(dossier_courant)
    diffs = comparer(reference, courant)
    synthese = synthetiser(diffs, reference, courant)
    ecrire_rapport_csv(diffs, chemin_csv, synthese)
    return synthese
