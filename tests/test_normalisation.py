# -*- coding: utf-8 -*-
"""Tests de la normalisation des codes établissement et des montants."""

import pytest

from quadra_ecritures.normalisation import lire_montant, normaliser_code


class TestNormaliserCode:
    @pytest.mark.parametrize("valeur,attendu", [
        ("704", "704"), (704, "704"), ("17641", "17641"),   # codes bruts
        ("7702", "702"), ("7730", "730"),                   # préfixe entité retiré
        ("  705 ", "705"),                                  # espaces
    ])
    def test_codes_bruts(self, valeur, attendu):
        assert normaliser_code(valeur) == attendu

    @pytest.mark.parametrize("valeur,attendu", [
        ("7702 - LE CHAT PERCHE", "702"),
        ("7730-AUTRE", "730"),
        ("17751 - DSP", "17751"),
    ])
    def test_format_prefixe_avec_nom(self, valeur, attendu):
        assert normaliser_code(valeur, extraire=True) == attendu

    @pytest.mark.parametrize("valeur,attendu", [
        ("000702", "702"), ("097881", "97881"), ("0", "0"),
    ])
    def test_zero_padding(self, valeur, attendu):
        assert normaliser_code(valeur, strip_zeros=True) == attendu

    @pytest.mark.parametrize("valeur", [
        None, "", "  ", "Total", "TOTAL GENERAL", "Somme", "texte", "70A4",
    ])
    def test_valeurs_ecartees(self, valeur):
        assert normaliser_code(valeur) is None


class TestLireMontant:
    def test_nombres(self):
        assert lire_montant(12.5) == 12.5
        assert lire_montant(-493.71) == -493.71
        assert lire_montant(100) == 100.0

    def test_arrondi_2_decimales(self):
        assert lire_montant(10.456) == 10.46

    @pytest.mark.parametrize("valeur", [None, 0, 0.0, 0.00001, "12,5", "texte", True, False])
    def test_valeurs_ecartees(self, valeur):
        assert lire_montant(valeur) is None
