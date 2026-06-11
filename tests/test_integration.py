# -*- coding: utf-8 -*-
"""Tests d'intégration : classeurs Excel générés à la volée, chaîne complète."""

import pytest
from openpyxl import Workbook

from quadra_ecritures.config import (Composante, Configuration, Source, SourcePaie)
from quadra_ecritures.moteur import (controler_equilibre, ecrire_fichiers,
                                     generer_ecritures, generer_ecritures_paie)

CRLF = b"\r\n"


@pytest.fixture
def environnement(tmp_path):
    """Crée deux classeurs sources (simple + paie) et la configuration associée."""
    entree = tmp_path / "entree"
    sortie = tmp_path / "sortie"
    entree.mkdir()

    # --- Classeur « une ligne = un établissement » (type OETH) ---
    wb = Workbook()
    ws = wb.active
    ws.title = "CRE"
    ws["A1"] = "Provision"                       # en-têtes hors zone de données
    donnees = [("7704 - CRECHE A", 100.0), ("7705 - CRECHE B", 0.0),
               ("7799 - SANS CENTRE", 50.0), ("Total", 150.0)]
    for i, (code, montant) in enumerate(donnees, start=3):
        ws.cell(i, 1, code)
        ws.cell(i, 3, montant)
    wb.save(entree / "provision.xlsx")

    # --- Classeur de paie détaillé par salarié (type PRECA) ---
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "Feuil1"
    paie = [  # (centre de coût, prime) — 2 salariés même centre + 1 négatif + 1 inconnu
        ("770401", 60.0), ("770401", 40.0), ("770501", -30.0), ("999999", 10.0)]
    for i, (centre, prime) in enumerate(paie, start=2):
        ws2.cell(i, 4, centre)
        ws2.cell(i, 14, prime)                    # colonne N
    wb2.save(entree / "preca.xlsx")

    cfg = Configuration(
        dossier_entree=str(entree),
        dossier_sortie=str(sortie),
        analytique={"704": "770401", "705": "770501"},
        centre_vers_dossier={"770401": "704", "770501": "705"},
        sources=[Source(
            fichier="provision.xlsx", feuille="CRE", ligne_debut=3,
            col_dossier="A", col_montant="C", extraire_code=True,
            compte_credit="43785000", compte_debit="63388100",
            libelle="OETH TEST", journal="OS", date_ecriture="310526",
            contre_passation="010626")],
        sources_paie=[SourcePaie(
            fichier="preca.xlsx", feuille="Feuil1", ligne_debut=2, col_centre="D",
            journal="OS", date_ecriture="310526", contre_passation=None,
            composantes=[Composante(col="N", compte_debit="64133820",
                                    compte_credit="42822000", libelle="PRIME TEST")])],
    )
    return cfg


class TestChaineComplete:
    def test_lignes_filtrees_et_dossiers_produits(self, environnement):
        par_dossier, sans_centre = generer_ecritures(environnement.sources, environnement)
        assert "704" in par_dossier                 # 100 € comptabilisés
        assert "705" not in par_dossier             # montant nul écarté
        assert "799" in par_dossier                 # écriture produite même sans centre
        assert ("799", "OETH TEST") in sans_centre  # ... mais signalée

    def test_agregation_paie_et_routage(self, environnement):
        par_paie, inconnus, attente = generer_ecritures_paie(
            environnement.sources_paie, environnement)
        # 60 + 40 agrégés sur le centre 770401 -> dossier 704
        debit = next(l for l in par_paie["704"] if l.startswith("M") and l[41] == "D")
        assert int(debit[42:55]) == 10000
        # montant négatif agrégé seul -> sens inversés sur le dossier 705
        charge_705 = next(l for l in par_paie["705"] if l[1:9] == "64133820")
        assert charge_705[41] == "C"
        assert inconnus == ["999999"]               # centre inconnu signalé
        assert attente == []

    def test_fichiers_disque_format_et_encodage(self, environnement):
        par_dossier, _ = generer_ecritures(environnement.sources, environnement)
        td, tc, deseq = ecrire_fichiers(par_dossier, environnement.dossier_sortie)
        assert td == tc and deseq == []
        brut = (open(f"{environnement.dossier_sortie}/704_ecriture_Quadra.txt", "rb").read())
        assert CRLF in brut and brut.endswith(CRLF)
        assert brut.count(b"\n") == brut.count(CRLF)            # CRLF strict
        lignes = brut.decode("cp1252").rstrip("\r\n").split("\r\n")
        assert all(len(l) in (39, 146) for l in lignes)
        assert all(l[0] in "MI" for l in lignes)

    def test_contre_passation_inverse_les_sens(self, environnement):
        normal, _ = generer_ecritures(environnement.sources, environnement)
        extourne, _ = generer_ecritures(environnement.sources, environnement, extourne=True)
        sens = lambda lots, cpt: next(
            l[41] for l in lots["704"] if l.startswith("M") and l[1:9] == cpt)
        assert sens(normal, "63388100") == "D" and sens(extourne, "63388100") == "C"
        assert sens(normal, "43785000") == "C" and sens(extourne, "43785000") == "D"
        # date d'extourne appliquée
        assert all(l[14:20] == "010626" for l in extourne["704"] if l.startswith("M"))

    def test_equilibre_global(self, environnement):
        par_dossier, _ = generer_ecritures(environnement.sources, environnement)
        par_paie, _, _ = generer_ecritures_paie(environnement.sources_paie, environnement)
        for d, lignes in par_paie.items():
            par_dossier[d].extend(lignes)
        for code, lignes in par_dossier.items():
            d, c = controler_equilibre(lignes)
            assert d == c, f"dossier {code} déséquilibré"
