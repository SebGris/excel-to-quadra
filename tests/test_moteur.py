# -*- coding: utf-8 -*-
"""Tests du moteur : paires d'écritures, analytique, négatifs, extourne, équilibre."""

from collections import defaultdict

from excel_to_quadra.moteur import ajouter_ecriture_pair, controler_equilibre


def _emettre(montant, extourne=False, centre="770401",
             compte_credit="40810000", compte_debit="62280000", facteur=1.0,
             ventilation=None):
    par_dossier = defaultdict(list)
    sans_centre = []
    ajouter_ecriture_pair(par_dossier, "704", "OS", "310526", "Lib",
                          compte_credit, compte_debit, montant,
                          centre, sans_centre, extourne, facteur=facteur,
                          ventilation=ventilation)
    return par_dossier["704"], sans_centre


def _sens(lignes, compte):
    return next(l[41] for l in lignes if l.startswith("M") and l[1:9] == compte)


class TestEcritureNominale:
    def test_paire_equilibree(self):
        lignes, _ = _emettre(100.0)
        d, c = controler_equilibre(lignes)
        assert d == c == 10000

    def test_sens_nominal(self):
        lignes, _ = _emettre(100.0)
        assert _sens(lignes, "62280000") == "D"     # charge au débit
        assert _sens(lignes, "40810000") == "C"     # provision au crédit

    def test_ligne_i_apres_la_charge_uniquement(self):
        lignes, _ = _emettre(100.0)
        idx_charge = next(i for i, l in enumerate(lignes) if l[1:9] == "62280000")
        assert lignes[idx_charge + 1].startswith("I")
        assert sum(1 for l in lignes if l.startswith("I")) == 1

    def test_ligne_i_sur_produit_classe_7(self):
        lignes, _ = _emettre(100.0, compte_credit="74140000", compte_debit="44170400")
        idx_produit = next(i for i, l in enumerate(lignes) if l[1:9] == "74140000")
        assert lignes[idx_produit + 1].startswith("I")

    def test_centre_inconnu_signale_sans_bloquer(self):
        lignes, sans_centre = _emettre(100.0, centre=None)
        assert sum(1 for l in lignes if l.startswith("M")) == 2   # écritures produites
        assert not any(l.startswith("I") for l in lignes)         # pas de ligne I
        assert sans_centre == [("704", "Lib")]                    # absence signalée


class TestMontantsNegatifs:
    def test_sens_inverses(self):
        lignes, _ = _emettre(-100.0)
        assert _sens(lignes, "62280000") == "C"     # charge passe au crédit
        assert _sens(lignes, "40810000") == "D"

    def test_montant_en_valeur_absolue_et_equilibre(self):
        lignes, _ = _emettre(-100.0)
        d, c = controler_equilibre(lignes)
        assert d == c == 10000


class TestExtourne:
    def test_sens_inverses(self):
        lignes, _ = _emettre(100.0, extourne=True)
        assert _sens(lignes, "62280000") == "C"
        assert _sens(lignes, "40810000") == "D"

    def test_extourne_d_un_negatif_revient_au_sens_d_origine(self):
        lignes, _ = _emettre(-100.0, extourne=True)
        assert _sens(lignes, "62280000") == "D"
        assert _sens(lignes, "40810000") == "C"


class TestFacteur:
    def test_facteur_proratise_le_montant(self):
        # 7/12 sur 13839,10 € -> 8072,81 € (807281 centimes)
        lignes, _ = _emettre(13839.10, facteur=7 / 12)
        d, c = controler_equilibre(lignes)
        assert d == c == 807281

    def test_facteur_par_defaut_neutre(self):
        lignes, _ = _emettre(13839.10)              # facteur 1.0 implicite
        d, c = controler_equilibre(lignes)
        assert d == c == 1383910

    def test_facteur_applique_avant_inversion_du_signe(self):
        # facteur négatif rend le montant négatif -> sens inversés
        lignes, _ = _emettre(100.0, facteur=-0.5)
        assert _sens(lignes, "62280000") == "C"     # charge passe au crédit
        d, c = controler_equilibre(lignes)
        assert d == c == 5000                        # |100 * -0,5| = 50 €


class TestVentilation:
    def _i_lignes(self, lignes):
        return [l for l in lignes if l.startswith("I")]

    def _m_charge(self, lignes):
        return next(int(l[42:55]) for l in lignes
                    if l.startswith("M") and l[1:9] == "62280000")

    def test_50_50_avec_solde_sur_la_derniere(self):
        ventil = [{"centre": "770401", "pourcent": 50.0},
                  {"centre": "770402", "pourcent": 50.0}]
        lignes, _ = _emettre(145.35, ventilation=ventil)
        i = self._i_lignes(lignes)
        assert len(i) == 2
        assert [int(l[6:19]) for l in i] == [7268, 7267]   # 72,68 puis solde 72,67
        assert [l[1:6] for l in i] == ["05000", "05000"]   # pourcentages réels
        assert sum(int(l[6:19]) for l in i) == self._m_charge(lignes)

    def test_trois_centres_somme_exacte_au_montant_m(self):
        ventil = [{"centre": "C1", "pourcent": 59.04},
                  {"centre": "C2", "pourcent": 13.34},
                  {"centre": "C3", "pourcent": 27.62}]
        lignes, _ = _emettre(1000.0, ventilation=ventil)
        i = self._i_lignes(lignes)
        assert len(i) == 3
        assert [l[1:6] for l in i] == ["05904", "01334", "02762"]
        assert sum(int(l[6:19]) for l in i) == self._m_charge(lignes)

    def test_dossier_sans_ventilation_inchange(self):
        lignes, _ = _emettre(100.0)                        # une seule ligne I à 100 %
        i = self._i_lignes(lignes)
        assert len(i) == 1
        assert i[0][1:6] == "10000"
        assert i[0][19:29].strip() == "770401"

    def test_ventilation_prime_meme_sans_centre_par_defaut(self):
        ventil = [{"centre": "880001", "pourcent": 100.0}]
        lignes, sans_centre = _emettre(100.0, centre=None, ventilation=ventil)
        i = self._i_lignes(lignes)
        assert len(i) == 1 and i[0][19:29].strip() == "880001"
        assert sans_centre == []                           # pas de signalement


class TestControlerEquilibre:
    def test_ignore_les_lignes_i(self):
        lignes, _ = _emettre(100.0)
        d, c = controler_equilibre(lignes)
        assert (d, c) == (10000, 10000)             # la ligne I ne compte pas

    def test_detecte_un_desequilibre(self):
        lignes, _ = _emettre(100.0)
        lignes_m = [l for l in lignes if l.startswith("M") and l[41] == "D"]
        d, c = controler_equilibre(lignes_m)
        assert d != c
