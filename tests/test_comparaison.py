# -*- coding: utf-8 -*-
"""Tests de la comparaison entre génération courante et version de référence."""

from excel_to_quadra.comparaison import (Difference, Synthese, comparer,
                                         comparer_dossiers, ecrire_rapport_csv,
                                         lire_ecritures_m, synthetiser)
from excel_to_quadra.format_quadra import formater_ligne_m


def _ecrire_fichier(dossier, code, lignes_m):
    chemin = dossier / f"{code}_ecriture_Quadra.txt"
    chemin.write_bytes(("\r\n".join(lignes_m) + "\r\n").encode("cp1252"))


class TestComparer:
    def test_ecriture_ajoutee_est_nouvelle(self):
        ref = {("704", "62280000", "D", "Lib"): 10000}
        cur = {("704", "62280000", "D", "Lib"): 10000,
               ("704", "40810000", "C", "Lib"): 10000}
        diffs = comparer(ref, cur)
        assert len(diffs) == 1
        d = diffs[0]
        assert d.type == "NOUVELLE" and d.compte == "40810000"
        assert d.avant is None and d.apres == 10000 and d.ecart == 10000

    def test_ecriture_retiree_est_supprimee(self):
        diffs = comparer({("704", "62280000", "D", "Lib"): 10000}, {})
        assert len(diffs) == 1 and diffs[0].type == "SUPPRIMEE"
        assert diffs[0].avant == 10000 and diffs[0].apres is None
        assert diffs[0].ecart == -10000

    def test_montant_modifie_avec_ecart(self):
        ref = {("704", "62280000", "D", "Lib"): 10000}
        cur = {("704", "62280000", "D", "Lib"): 15000}
        diffs = comparer(ref, cur)
        assert len(diffs) == 1 and diffs[0].type == "MONTANT_MODIFIE"
        assert diffs[0].avant == 10000 and diffs[0].apres == 15000
        assert diffs[0].ecart == 5000

    def test_ecriture_identique_absente_du_rapport(self):
        ref = {("704", "62280000", "D", "Lib"): 10000}
        assert comparer(ref, dict(ref)) == []


class TestLectureEtDossiers:
    def test_nouveau_dossier_entier_en_nouvelle(self, tmp_path):
        ref = tmp_path / "ref"
        cur = tmp_path / "cur"
        ref.mkdir()
        cur.mkdir()
        m1 = formater_ligne_m("62280000", "OS", "310526", "Charge", "D", 100.0)
        m2 = formater_ligne_m("40810000", "OS", "310526", "Charge", "C", 100.0)
        _ecrire_fichier(cur, "999", [m1, m2])         # dossier absent de la référence
        diffs = comparer(lire_ecritures_m(str(ref)), lire_ecritures_m(str(cur)))
        assert diffs and all(d.type == "NOUVELLE" and d.dossier == "999" for d in diffs)

    def test_dossier_extrait_du_nom_de_fichier(self, tmp_path):
        cur = tmp_path / "cur"
        cur.mkdir()
        _ecrire_fichier(cur, "17641", [formater_ligne_m("74140000", "OS", "310526",
                                                         "Forfait", "C", 50.0)])
        ecritures = lire_ecritures_m(str(cur))
        assert list(ecritures)[0][0] == "17641"       # dossier = préfixe du nom

    def test_numero_piece_hors_de_la_cle(self, tmp_path):
        # Deux générations identiques SAUF le n° de pièce (position 100, incrémental)
        # -> aucune différence : le n° de pièce ne fait pas partie de la clé.
        ref = tmp_path / "ref"
        cur = tmp_path / "cur"
        ref.mkdir()
        cur.mkdir()
        m_ref = formater_ligne_m("62280000", "OS", "310526", "Charge", "D", 100.0,
                                 numero_piece="IMPORT01")
        m_cur = formater_ligne_m("62280000", "OS", "310526", "Charge", "D", 100.0,
                                 numero_piece="IMPORT02")
        assert m_ref != m_cur                         # les lignes diffèrent (pos 100)
        _ecrire_fichier(ref, "704", [m_ref])
        _ecrire_fichier(cur, "704", [m_cur])
        assert comparer(lire_ecritures_m(str(ref)), lire_ecritures_m(str(cur))) == []


class TestComparerDossiers:
    def test_reference_vide_pas_de_diff_ni_csv(self, tmp_path):
        ref = tmp_path / "ref"
        cur = tmp_path / "cur"
        ref.mkdir()
        cur.mkdir()
        _ecrire_fichier(cur, "704", [formater_ligne_m("62280000", "OS", "310526",
                                                      "Lib", "D", 100.0)])
        csv_path = tmp_path / "diff.csv"
        res = comparer_dossiers(str(ref), str(cur), str(csv_path))
        assert res is None                            # référence vide -> rien
        assert not csv_path.exists()


class TestRapportCsv:
    def test_entete_et_separateur(self, tmp_path):
        diffs = [Difference("NOUVELLE", "704", "40810000", "C", "Lib".ljust(20),
                            None, 10000, 10000)]
        synth = Synthese(1, 1, 0, 0, 0, 10000, 10000)
        csv_path = tmp_path / "diff.csv"
        ecrire_rapport_csv(diffs, str(csv_path), synth)
        lignes = csv_path.read_text(encoding="utf-8-sig").splitlines()
        assert lignes[0] == ("Type;Dossier;Compte;Sens;Libelle;"
                             "Montant_avant;Montant_apres;Ecart")
        assert "NOUVELLE;704;40810000;C;Lib;;100,00;100,00" in lignes

    def test_synthese_calculee(self):
        ref = {("704", "62280000", "D", "Lib"): 10000,
               ("705", "62280000", "D", "Lib"): 20000}
        cur = {("704", "62280000", "D", "Lib"): 15000,           # modifié
               ("706", "40810000", "C", "Lib"): 5000}            # nouveau
        diffs = comparer(ref, cur)
        s = synthetiser(diffs, ref, cur)
        assert s.nouvelles == 1 and s.supprimees == 1 and s.modifiees == 1
        assert s.dossiers == 3                                   # 704, 705, 706
        assert s.total_avant == 30000 and s.total_apres == 20000
        assert s.ecart == -10000
