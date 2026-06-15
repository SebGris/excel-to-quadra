# -*- coding: utf-8 -*-
"""Mise en forme des enregistrements du fichier ASCII QuadraCOMPTA.

Référence : spécification Quadratus « Fichier d'entrée ASCII dans QuadraCOMPTA »
(QC_ASC, mise à jour 02/2015) :
  - enregistrement M (écriture comptable)  : 146 caractères ;
  - enregistrement I (ligne analytique)    : 39 caractères, placé immédiatement
    après l'enregistrement M qu'il ventile.
"""

from decimal import ROUND_HALF_UP, Decimal
from typing import Optional

LONGUEUR_LIGNE_M = 146
LONGUEUR_LIGNE_I = 39

#: Position (1-indexée) et longueur du n° de pièce dans l'enregistrement M.
POS_NUMERO_PIECE = 100
LONGUEUR_NUMERO_PIECE = 8


def euros_vers_centimes(montant_euros: float) -> int:
    """Convertit un montant en euros vers des centimes entiers (valeur absolue).

    Utilise Decimal avec arrondi commercial (ROUND_HALF_UP) pour éviter les
    erreurs de représentation binaire des flottants (ex. 0.005 € doit donner
    1 centime, alors que round(0.005 * 100) en float donnerait 0).
    """
    d = Decimal(str(abs(montant_euros))) * 100
    return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def formater_ligne_m(compte: str, journal: str, date_jjmmaa: str,
                     libelle: str, sens: str, montant_euros: float,
                     numero_piece: Optional[str] = None) -> str:
    """Construit une ligne d'écriture Quadra de type M (146 caractères).

    Champs obligatoires de la spécification : type, compte (8), journal (2),
    folio « 000 », date JJMMAA, sens D/C, montant en centimes sur 13 caractères.

    Le n° de pièce optionnel est inscrit dans la zone blanche en position 100
    (8 c., justifié à gauche), sans décaler les champs 1-55 ni changer la
    longueur. Absent (None) : la zone reste à blanc (comportement inchangé).
    """
    if sens not in ("D", "C"):
        raise ValueError(f"Sens invalide : {sens!r} (attendu 'D' ou 'C')")
    ligne = (
        "M"                                          # pos 1     : type
        + str(compte).ljust(8)[:8]                   # pos 2-9   : n° de compte
        + str(journal).ljust(2)[:2]                  # pos 10-11 : code journal
        + "000"                                      # pos 12-14 : folio
        + str(date_jjmmaa).ljust(6)[:6]              # pos 15-20 : date JJMMAA
        + " "                                        # pos 21    : code libellé
        + str(libelle).ljust(20)[:20]                # pos 22-41 : libellé (20 c.)
        + sens                                       # pos 42    : sens D/C
        + str(euros_vers_centimes(montant_euros)).rjust(13, "0")  # pos 43-55
    )
    ligne = ligne.ljust(LONGUEUR_LIGNE_M)
    if numero_piece:                                 # zone blanche en position 100
        i = POS_NUMERO_PIECE - 1                     # passage en index 0
        piece = str(numero_piece).ljust(LONGUEUR_NUMERO_PIECE)[:LONGUEUR_NUMERO_PIECE]
        ligne = ligne[:i] + piece + ligne[i + LONGUEUR_NUMERO_PIECE:]
    return ligne


def formater_ligne_i(centre: str, montant_euros: float, pourcent: float = 100.0) -> str:
    """Construit une ligne analytique Quadra de type I (39 caractères)."""
    pct = str(int(round(pourcent * 100))).rjust(5, "0")   # 100,00 % -> « 10000 »
    ligne = (
        "I"                                          # pos 1     : type
        + pct                                        # pos 2-6   : % de répartition
        + str(euros_vers_centimes(montant_euros)).rjust(13, "0")  # pos 7-19
        + str(centre).ljust(10)[:10]                 # pos 20-29 : code centre
    )
    return ligne.ljust(LONGUEUR_LIGNE_I)             # pos 30-39 : nature (à blanc)
