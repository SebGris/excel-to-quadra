# -*- coding: utf-8 -*-
"""Normalisation des codes établissement et des montants lus dans Excel."""

from typing import Optional


def normaliser_code(valeur, extraire: bool = False, strip_zeros: bool = False) -> Optional[str]:
    """Renvoie le code dossier normalisé, ou None si la cellule n'est pas un dossier.

    Formats pris en charge :
      - code brut                      : 704, 17641 ;
      - code préfixé entité (extraire) : « 7702 - NOM »  -> 702 ;
      - code zéro-paddé (strip_zeros)  : « 000702 »      -> 702.

    Les lignes de totaux (« Total… », « Somme… ») et les valeurs non numériques
    sont écartées.
    """
    if valeur is None:
        return None
    texte = str(valeur).strip()
    bas = texte.lower()
    if bas == "" or bas.startswith("total") or bas.startswith("somme"):
        return None
    if extraire:                                   # format « 7702 -  NOM »
        texte = texte.split("-")[0].strip()
    if strip_zeros and texte.isdigit():            # 000702 -> 702
        texte = texte.lstrip("0") or "0"
    if not texte.isdigit():                        # un code dossier est numérique
        return None
    if len(texte) == 4 and texte.startswith("7"):  # 7702 -> 702, 7730 -> 730
        texte = texte[1:]
    return texte


def lire_montant(valeur) -> Optional[float]:
    """Renvoie un montant float exploitable (≠ 0), ou None si vide / nul / texte."""
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, (int, float)):
        montant = round(float(valeur), 2)
        return montant if abs(montant) > 0.0001 else None
    return None
