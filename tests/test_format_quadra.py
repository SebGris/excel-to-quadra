# -*- coding: utf-8 -*-
"""Tests des enregistrements M et I (positions de la spécification QC_ASC)."""

import pytest

from excel_to_quadra.format_quadra import (
    LONGUEUR_LIGNE_I, LONGUEUR_LIGNE_M, euros_vers_centimes,
    formater_ligne_i, formater_ligne_m)


class TestLigneM:
    def test_longueur(self):
        assert len(formater_ligne_m("62280000", "OS", "310526", "Lib", "D", 1.0)) == LONGUEUR_LIGNE_M

    def test_positions_de_la_specification(self):
        l = formater_ligne_m("62280000", "OS", "310526", "Cout paie 0526", "D", 1234.56)
        assert l[0] == "M"                          # pos 1  : type
        assert l[1:9] == "62280000"                 # pos 2  : compte (8)
        assert l[9:11] == "OS"                      # pos 10 : journal (2)
        assert l[11:14] == "000"                    # pos 12 : folio
        assert l[14:20] == "310526"                 # pos 15 : date JJMMAA
        assert l[20] == " "                         # pos 21 : code libellé
        assert l[21:41] == "Cout paie 0526      "   # pos 22 : libellé (20)
        assert l[41] == "D"                         # pos 42 : sens
        assert l[42:55] == "0000000123456"          # pos 43 : centimes (13)
        assert l[55:] == " " * 91                   # reste à blanc

    def test_libelle_tronque_a_20(self):
        l = formater_ligne_m("60000000", "OS", "310526", "X" * 30, "C", 1.0)
        assert l[21:41] == "X" * 20
        assert len(l) == LONGUEUR_LIGNE_M

    def test_montant_en_valeur_absolue(self):
        l = formater_ligne_m("60000000", "OS", "310526", "Lib", "D", -10.0)
        assert l[42:55] == "0000000001000"

    def test_sens_invalide(self):
        with pytest.raises(ValueError):
            formater_ligne_m("60000000", "OS", "310526", "Lib", "X", 1.0)

    def test_numero_piece_aux_positions_100_107(self):
        l = formater_ligne_m("62280000", "OS", "310526", "Lib", "D", 1.0, numero_piece="IMPORT")
        assert l[99:107] == "IMPORT  "          # pos 100-107 : « IMPORT » + 2 espaces
        assert len(l) == LONGUEUR_LIGNE_M

    def test_numero_piece_ne_decale_pas_les_positions_1_55(self):
        sans = formater_ligne_m("62280000", "OS", "310526", "Cout paie 0526", "D", 1234.56)
        avec = formater_ligne_m("62280000", "OS", "310526", "Cout paie 0526", "D", 1234.56,
                                numero_piece="IMPORT")
        assert avec[:55] == sans[:55]            # compte/journal/date/libellé/sens/montant
        assert avec[99:107] == "IMPORT  "
        assert len(avec) == LONGUEUR_LIGNE_M

    def test_sans_numero_piece_zone_piece_a_blanc(self):
        l = formater_ligne_m("62280000", "OS", "310526", "Lib", "D", 1.0)
        assert l[99:107] == " " * 8
        assert l[55:] == " " * 91                # comportement strictement inchangé

    def test_numero_piece_tronque_a_8(self):
        l = formater_ligne_m("62280000", "OS", "310526", "Lib", "D", 1.0, numero_piece="ABCDEFGHIJ")
        assert l[99:107] == "ABCDEFGH"
        assert len(l) == LONGUEUR_LIGNE_M

    def test_numero_piece_ne_pollue_pas_le_reste_de_la_zone(self):
        l = formater_ligne_m("62280000", "OS", "310526", "Lib", "D", 1.0, numero_piece="IMPORT")
        assert l[99:107] == "IMPORT  "              # seule la zone pièce porte le n°
        assert l[55:99] == " " * 44                 # pos 56-99 à blanc
        assert l[107:] == " " * 39                  # pos 108-146 à blanc
        assert len(l) == LONGUEUR_LIGNE_M


class TestLigneI:
    def test_longueur(self):
        assert len(formater_ligne_i("770401", 1.0)) == LONGUEUR_LIGNE_I

    def test_positions_de_la_specification(self):
        l = formater_ligne_i("770401", 1234.56)
        assert l[0] == "I"                          # pos 1  : type
        assert l[1:6] == "10000"                    # pos 2  : % (100,00)
        assert l[6:19] == "0000000123456"           # pos 7  : montant (13)
        assert l[19:29] == "770401    "             # pos 20 : centre (10)
        assert l[29:] == " " * 10                   # pos 30 : nature à blanc

    def test_pourcentage_partiel(self):
        assert formater_ligne_i("770202", 100.0, pourcent=40.0)[1:6] == "04000"


class TestConversionCentimes:
    @pytest.mark.parametrize("euros,centimes", [
        (12.0, 1200), (0.01, 1), (1234.56, 123456),
        (-493.71, 49371),            # valeur absolue
        (0.005, 1), (2.675, 268),    # arrondis bancaires float -> round propre
    ])
    def test_conversion(self, euros, centimes):
        assert euros_vers_centimes(euros) == centimes
