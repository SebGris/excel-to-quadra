# -*- coding: utf-8 -*-
"""Moteur de génération des écritures Quadra.

Chaque écriture est produite en paire équilibrée (un débit, un crédit du même
montant). Une ligne analytique (type I) est ajoutée immédiatement après la part
de charge ou de produit (comptes de classe 6 ou 7), ventilée à 100 % vers le
centre du dossier.

Règles de gestion :
  - montant négatif (régularisation)  -> sens débit/crédit inversés ;
  - extourne (contre-passation)        -> sens inversés à la date d'extourne ;
  - extourne d'un montant négatif      -> double inversion = sens d'origine ;
  - centre analytique inconnu          -> ligne M produite, absence signalée.
"""

import glob
import os
from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal
from typing import Dict, List, Optional, Tuple

#: Motif des fichiers générés (un par dossier, + variante _contrepass).
MOTIF_SORTIE = "*_ecriture_Quadra*.txt"

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .config import Composante, Configuration, Source, SourcePaie
from .format_quadra import formater_ligne_i, formater_ligne_m
from .normalisation import lire_montant, normaliser_code

ParDossier = Dict[str, List[str]]


def _est_charge_ou_produit(compte: str) -> bool:
    return str(compte)[:1] in ("6", "7")


def ajouter_ecriture_pair(par_dossier: ParDossier, dossier: str, journal: str,
                          date: str, libelle: str, compte_credit: str,
                          compte_debit: str, montant: float,
                          centre: Optional[str],
                          sans_centre: Optional[list],
                          extourne: bool = False, facteur: float = 1.0,
                          ventilation: Optional[List[dict]] = None) -> None:
    """Ajoute une écriture équilibrée (M crédit + M débit) et sa/ses ligne(s) I.

    Sans ventilation : une ligne I à 100 % vers `centre`. Avec ventilation (liste
    de {centre, pourcent}) : une ligne I par centre, la dernière recevant le solde
    pour que la somme des lignes I égale exactement la ligne M.
    """
    montant = round(montant * facteur, 2)      # prorata avant calcul du sens
    inverser = extourne ^ (montant < 0)        # extourne et négatif se cumulent
    montant = abs(montant)
    if inverser:                               # l'ex-débit passe au crédit
        legs = [(compte_debit, "C"), (compte_credit, "D")]
    else:
        legs = [(compte_credit, "C"), (compte_debit, "D")]
    # part de charge/produit (classe 6/7) écrite en premier, suivie de sa ligne I
    legs.sort(key=lambda x: 0 if _est_charge_ou_produit(x[0]) else 1)
    for compte, sens in legs:
        par_dossier[dossier].append(formater_ligne_m(compte, journal, date, libelle, sens, montant))
        if not _est_charge_ou_produit(compte):
            continue
        if ventilation:                        # ventilation prioritaire sur le centre
            # Arrondi commercial (Decimal/ROUND_HALF_UP) comme pour les centimes :
            # jamais d'arithmétique flottante sur les montants comptables.
            reste = Decimal(str(montant))
            dernier = len(ventilation) - 1
            for i, v in enumerate(ventilation):
                if i < dernier:
                    part = (Decimal(str(montant)) * Decimal(str(v["pourcent"])) / 100
                            ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    reste -= part
                else:
                    part = reste               # le solde garantit somme(I) == M
                par_dossier[dossier].append(formater_ligne_i(v["centre"], part, v["pourcent"]))
        elif centre:
            par_dossier[dossier].append(formater_ligne_i(centre, montant))
        elif sans_centre is not None:
            sans_centre.append((dossier, libelle))


def generer_ecritures(sources: List[Source], cfg: Configuration,
                      extourne: bool = False) -> Tuple[ParDossier, list]:
    """Traite les sources « une ligne = un établissement ».

    Renvoie (par_dossier, sans_centre) où sans_centre liste les couples
    (dossier, libellé) dont la ligne analytique n'a pas pu être générée.
    """
    par_dossier: ParDossier = defaultdict(list)
    sans_centre: list = []
    cache_wb = {}

    for src in sources:
        if extourne and not src.contre_passation:
            continue
        chemin = os.path.join(cfg.dossier_entree, src.fichier)
        if chemin not in cache_wb:
            cache_wb[chemin] = load_workbook(chemin, data_only=True)
        ws = cache_wb[chemin][src.feuille]

        c_dossier = column_index_from_string(src.col_dossier)
        c_montant = column_index_from_string(src.col_montant)
        date = src.contre_passation if extourne else src.date_ecriture

        # En mode agrégé, on cumule (dossier, centre) -> montant et on émet
        # une seule écriture par dossier après la boucle, au lieu de ligne à ligne.
        cumul: Dict[Tuple[str, Optional[str]], float] = defaultdict(float)

        for r in range(src.ligne_debut, ws.max_row + 1):
            code_brut = normaliser_code(ws.cell(r, c_dossier).value,
                                        src.extraire_code, src.strip_zeros)
            montant = lire_montant(ws.cell(r, c_montant).value)
            if code_brut is None or montant is None:
                continue
            # Alias : le dossier lu est remplacé par sa cible (centre résolu sur
            # la cible) — contrairement au remap, le code lu n'est pas conservé.
            code_brut = cfg.alias_dossiers.get(code_brut, code_brut)
            dossier = src.remap.get(code_brut, code_brut)
            # Si le code lu EST un code analytique remappé, il sert de centre ;
            # sinon le centre vient de la table dossier -> centre.
            centre = code_brut if code_brut in src.remap else cfg.analytique.get(dossier)
            if src.agreger:
                cumul[(dossier, centre)] += montant
            else:
                ajouter_ecriture_pair(par_dossier, dossier, src.journal, date, src.libelle,
                                      src.compte_credit, src.compte_debit, montant,
                                      centre, sans_centre, extourne, facteur=src.facteur,
                                      ventilation=src.ventilation.get(dossier))

        for dossier, centre in sorted(cumul, key=lambda k: (k[0], k[1] or "")):
            ajouter_ecriture_pair(par_dossier, dossier, src.journal, date, src.libelle,
                                  src.compte_credit, src.compte_debit,
                                  round(cumul[(dossier, centre)], 2),
                                  centre, sans_centre, extourne, facteur=src.facteur,
                                  ventilation=src.ventilation.get(dossier))
    return par_dossier, sans_centre


def generer_ecritures_paie(sources: List[SourcePaie], cfg: Configuration,
                           extourne: bool = False) -> Tuple[ParDossier, list, list]:
    """Traite les classeurs de paie détaillés par salarié.

    Les montants sont agrégés par centre de coût, le dossier est retrouvé via
    la table inverse centre -> dossier, et le centre de coût sert directement
    de centre analytique. Renvoie (par_dossier, centres_inconnus, en_attente).
    """
    par_dossier: ParDossier = defaultdict(list)
    centres_inconnus: list = []
    en_attente: list = []
    cache_wb = {}

    for src in sources:
        if extourne and not src.contre_passation:
            continue
        chemin = os.path.join(cfg.dossier_entree, src.fichier)
        if chemin not in cache_wb:
            cache_wb[chemin] = load_workbook(chemin, data_only=True)
        ws = cache_wb[chemin][src.feuille]
        c_centre = column_index_from_string(src.col_centre)
        date = src.contre_passation if extourne else src.date_ecriture

        # Agrégation des montants par centre de coût, composante par composante
        cumul: Dict[str, Dict[str, float]] = defaultdict(lambda: defaultdict(float))
        for r in range(src.ligne_debut, ws.max_row + 1):
            centre = ws.cell(r, c_centre).value
            if centre is None:
                continue
            centre = str(centre).strip()
            if not centre.isdigit():
                continue
            for comp in src.composantes:
                v = lire_montant(ws.cell(r, column_index_from_string(comp.col)).value)
                if v:
                    cumul[centre][comp.col] += v

        for centre in sorted(cumul):
            dossier = cfg.centre_vers_dossier.get(centre)
            if dossier is None:
                centres_inconnus.append(centre)
                continue
            dossier = cfg.alias_dossiers.get(dossier, dossier)   # alias de dossier
            for comp in src.composantes:
                montant = round(cumul[centre].get(comp.col, 0.0), 2)
                if abs(montant) < 0.005:
                    continue
                if not comp.complete:
                    en_attente.append(comp.libelle)
                    continue
                ajouter_ecriture_pair(par_dossier, dossier, src.journal, date, comp.libelle,
                                      comp.compte_credit, comp.compte_debit, montant,
                                      centre, None, extourne)
    return par_dossier, sorted(set(centres_inconnus)), sorted(set(en_attente))


def controler_equilibre(lignes: List[str]) -> Tuple[int, int]:
    """Renvoie (total débits, total crédits) en centimes — lignes M uniquement."""
    debit = credit = 0
    for l in lignes:
        if not l.startswith("M"):
            continue
        centimes = int(l[42:55])
        if l[41] == "D":
            debit += centimes
        else:
            credit += centimes
    return debit, credit


def nettoyer_sortie(dossier_sortie: str) -> list:
    """Supprime du dossier de sortie les seuls fichiers `*_ecriture_Quadra*.txt`.

    Évite qu'un fichier orphelin d'un run antérieur (dossier disparu, aliasé ou
    renommé) ne soit importé par erreur. Tout autre fichier est préservé. À
    appeler une fois au démarrage, avant la première écriture (les deux passes
    arrêté + contre-passation écrivent toutes deux des fichiers de ce motif).
    """
    if not os.path.isdir(dossier_sortie):
        return []
    supprimes = []
    for chemin in glob.glob(os.path.join(dossier_sortie, MOTIF_SORTIE)):
        os.remove(chemin)
        supprimes.append(os.path.basename(chemin))
    return supprimes


def ecrire_fichiers(par_dossier: ParDossier, dossier_sortie: str,
                    suffixe: str = "") -> Tuple[int, int, list]:
    """Écrit un fichier texte par dossier (cp1252, CRLF).

    Renvoie (total débits, total crédits, déséquilibres) en centimes.
    """
    os.makedirs(dossier_sortie, exist_ok=True)
    total_d = total_c = 0
    desequilibres = []
    for code in sorted(par_dossier):
        lignes = par_dossier[code]
        d, c = controler_equilibre(lignes)
        total_d += d
        total_c += c
        if d != c:
            desequilibres.append((code, d, c))
        nom = f"{code}_ecriture_Quadra{suffixe}.txt"
        with open(os.path.join(dossier_sortie, nom), "w",
                  encoding="cp1252", newline="") as fh:
            fh.write("\r\n".join(lignes) + "\r\n")
    return total_d, total_c, desequilibres
