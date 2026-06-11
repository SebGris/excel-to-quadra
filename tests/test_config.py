# -*- coding: utf-8 -*-
"""Tests du chargement de la configuration YAML."""

import pytest

from excel_to_quadra.config import A_RENSEIGNER, charger_configuration

YAML_MINIMAL = """
dossier_entree: "/entree"
dossier_sortie: "/sortie"
analytique:
  "704": "770401"
centres_supplementaires:
  "770202": "702"
sources:
  - fichier: "test.xlsx"
    feuille: "Récap"
    ligne_debut: 4
    col_dossier: "B"
    col_montant: "D"
    compte_credit: "74140000"
    compte_debit: "44170400"
    libelle: "Forfait"
    journal: "OS"
    date_ecriture: "310526"
sources_paie:
  - fichier: "paie.xlsx"
    feuille: "Feuil1"
    ligne_debut: 6
    col_centre: "D"
    journal: "OS"
    date_ecriture: "310526"
    composantes:
      - col: "N"
        compte_debit: "64133820"
        compte_credit: "42822000"
        libelle: "Prime"
      - col: "O"
        compte_debit: "XXXXXXXX"
        compte_credit: "43822000"
        libelle: "En attente"
"""


def test_chargement_nominal(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text(YAML_MINIMAL, encoding="utf-8")
    cfg = charger_configuration(str(chemin))
    assert cfg.analytique["704"] == "770401"
    assert cfg.centre_vers_dossier["770401"] == "704"     # table inverse
    assert cfg.centre_vers_dossier["770202"] == "702"     # centres supplémentaires
    assert cfg.sources[0].extraire_code is False          # valeur par défaut
    assert cfg.sources[0].complete
    assert cfg.sources_paie[0].composantes[0].complete
    assert not cfg.sources_paie[0].composantes[1].complete  # XXXXXXXX détecté


def test_cle_obligatoire_manquante(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text("dossier_entree: '/entree'", encoding="utf-8")
    with pytest.raises(ValueError, match="dossier_sortie"):
        charger_configuration(str(chemin))
