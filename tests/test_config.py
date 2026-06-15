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


YAML_OPTIONS = """
dossier_entree: "/entree"
dossier_sortie: "/sortie"
alias_dossiers:
  "736": "723"
analytique:
  "723": "772301"
sources:
  - fichier: "test.xlsx"
    feuille: "Feuil1"
    ligne_debut: 2
    col_dossier: "B"
    col_montant: "H"
    compte_credit: "40810000"
    compte_debit: "62280000"
    libelle: "Ventile"
    journal: "OS"
    date_ecriture: "310526"
    ventilation:
      "723":
        - {centre: "772301", pourcent: 59.04}
        - {centre: "772302", pourcent: 13.34}
        - {centre: "772303", pourcent: 27.62}
sources_paie: []
"""


def test_chargement_alias_et_ventilation(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text(YAML_OPTIONS, encoding="utf-8")
    cfg = charger_configuration(str(chemin))
    assert cfg.alias_dossiers == {"736": "723"}        # option globale
    ventil = cfg.sources[0].ventilation["723"]
    assert len(ventil) == 3
    assert ventil[0]["centre"] == "772301"
    assert ventil[0]["pourcent"] == 59.04              # converti en float


def test_alias_dossiers_absent_par_defaut(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text(YAML_MINIMAL, encoding="utf-8")
    cfg = charger_configuration(str(chemin))
    assert cfg.alias_dossiers == {}                    # défaut : dict vide
    assert cfg.sources[0].ventilation == {}


YAML_CENTRES = """
dossier_entree: "/entree"
dossier_sortie: "/sortie"
analytique:
  "790": "179101"
centres_supplementaires:
  "179102": "790"
sources:
  - fichier: "x.xlsx"
    feuille: "F"
    ligne_debut: 2
    col_dossier: "A"
    col_montant: "C"
    compte_credit: "40810000"
    compte_debit: "62280000"
    libelle: "X"
    journal: "OS"
    date_ecriture: "310526"
    ventilation:
      "790":
        - {centre: "179103", pourcent: 100.0}
sources_paie: []
"""


def test_centres_connus_union_des_trois_sources(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text(YAML_CENTRES, encoding="utf-8")
    cfg = charger_configuration(str(chemin))
    # (a) analytique, (b) centres_supplementaires, (c) centres des ventilations
    assert cfg.centres_connus() == {"179101", "179102", "179103"}


def test_numero_piece_global_et_defaut_source(tmp_path):
    chemin = tmp_path / "cfg.yaml"
    chemin.write_text(YAML_MINIMAL + '\nnumero_piece: "IMPORT"\n', encoding="utf-8")
    cfg = charger_configuration(str(chemin))
    assert cfg.numero_piece == "IMPORT"                # option globale chargée
    assert cfg.sources[0].numero_piece is None         # pas de surcharge -> None par défaut
