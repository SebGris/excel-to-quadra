# -*- coding: utf-8 -*-
"""Chargement et validation de la configuration YAML.

La configuration externalise tout le référentiel métier (chemins, tables
analytiques, sources et comptes) afin que le code reste générique et
publiable : aucune donnée d'organisation n'est codée en dur dans le package.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

import yaml

#: Marqueur d'un compte non renseigné : la source est ignorée (et signalée).
A_RENSEIGNER = "XXXXXXXX"


@dataclass
class Source:
    """Une colonne de montant d'un classeur « une ligne = un établissement »."""
    fichier: str
    feuille: str
    ligne_debut: int
    col_dossier: str
    col_montant: str
    compte_credit: str
    compte_debit: str
    libelle: str
    journal: str
    date_ecriture: str
    extraire_code: bool = False
    strip_zeros: bool = False
    contre_passation: Optional[str] = None
    remap: Dict[str, str] = field(default_factory=dict)
    agreger: bool = False                 # cumuler les lignes d'un même dossier
    facteur: float = 1.0                  # prorata appliqué au montant (ex. 7/12)
    # ventilation analytique multi-centres : dossier -> [{centre, pourcent}, ...]
    ventilation: Dict[str, List[dict]] = field(default_factory=dict)
    # filtre de dates optionnel (bornes AAAAMMJJ incluses) sur une colonne date
    col_date: Optional[str] = None
    date_min: Optional[str] = None
    date_max: Optional[str] = None
    numero_piece: Optional[str] = None    # surcharge le n° de pièce global

    @property
    def complete(self) -> bool:
        return A_RENSEIGNER not in (self.compte_credit, self.compte_debit)


@dataclass
class Composante:
    """Une composante d'une provision de paie (ex. prime / charges sociales)."""
    col: str
    compte_debit: str
    compte_credit: str
    libelle: str

    @property
    def complete(self) -> bool:
        return A_RENSEIGNER not in (self.compte_credit, self.compte_debit)


@dataclass
class SourcePaie:
    """Un classeur de paie détaillé par salarié, agrégé par centre de coût."""
    fichier: str
    feuille: str
    ligne_debut: int
    col_centre: str
    journal: str
    date_ecriture: str
    composantes: List[Composante]
    contre_passation: Optional[str] = None
    numero_piece: Optional[str] = None    # surcharge le n° de pièce global
    col_matricule: str = "G"              # colonne du matricule (détection de doublons)


@dataclass
class Configuration:
    dossier_entree: str
    dossier_sortie: str
    analytique: Dict[str, str]            # dossier -> centre analytique
    centre_vers_dossier: Dict[str, str]   # centre  -> dossier (table inverse)
    sources: List[Source]
    sources_paie: List[SourcePaie]
    alias_dossiers: Dict[str, str] = field(default_factory=dict)  # dossier lu -> cible
    numero_piece: Optional[str] = None    # n° de pièce global (journal partagé)
    numero_piece_incremental: bool = False  # accole un compteur de run au n° de pièce

    def centres_connus(self) -> set:
        """Ensemble des centres analytiques valides.

        Union de (a) la table analytique et (b) les centres supplémentaires
        — tous deux présents comme clés de `centre_vers_dossier` — et (c) de
        tous les centres cités dans les `ventilation` des sources.
        """
        connus = set(self.centre_vers_dossier)            # (a) + (b)
        for src in self.sources:
            for liste in src.ventilation.values():        # (c)
                for entree in liste:
                    connus.add(entree["centre"])
        return connus


def _normaliser_ventilation(brut) -> Dict[str, List[dict]]:
    """Normalise une table de ventilation : dossier -> [{centre, pourcent}, ...]."""
    return {
        str(dossier): [{"centre": str(e["centre"]), "pourcent": float(e["pourcent"])}
                       for e in liste]
        for dossier, liste in (brut or {}).items()
    }


def charger_configuration(chemin: str) -> Configuration:
    """Charge et valide un fichier de configuration YAML."""
    with open(chemin, encoding="utf-8") as fh:
        brut = yaml.safe_load(fh) or {}

    for cle in ("dossier_entree", "dossier_sortie"):
        if not brut.get(cle):
            raise ValueError(f"Configuration : clé obligatoire manquante « {cle} »")

    analytique = {str(k): str(v) for k, v in (brut.get("analytique") or {}).items()}

    # Table inverse construite depuis « analytique », complétée par les centres
    # supplémentaires (clusters multi-activités : plusieurs centres -> un dossier).
    centre_vers_dossier = {centre: dossier for dossier, centre in analytique.items()}
    for centre, dossier in (brut.get("centres_supplementaires") or {}).items():
        centre_vers_dossier[str(centre)] = str(dossier)

    alias_dossiers = {str(k): str(v) for k, v in (brut.get("alias_dossiers") or {}).items()}

    sources = [Source(**{**s,
                         "remap": {str(k): str(v) for k, v in (s.get("remap") or {}).items()},
                         "ventilation": _normaliser_ventilation(s.get("ventilation"))})
               for s in (brut.get("sources") or [])]
    sources_paie = [
        SourcePaie(**{**s, "composantes": [Composante(**c) for c in s["composantes"]]})
        for s in (brut.get("sources_paie") or [])
    ]
    return Configuration(
        dossier_entree=brut["dossier_entree"],
        dossier_sortie=brut["dossier_sortie"],
        analytique=analytique,
        centre_vers_dossier=centre_vers_dossier,
        sources=sources,
        sources_paie=sources_paie,
        alias_dossiers=alias_dossiers,
        numero_piece=brut.get("numero_piece"),
        numero_piece_incremental=bool(brut.get("numero_piece_incremental", False)),
    )
