# -*- coding: utf-8 -*-
"""Point d'entrée en ligne de commande.

Usage :
    excel-to-quadra --config config/situation.local.yaml
    python -m excel_to_quadra.cli --config config/situation.local.yaml
"""

import argparse
import os
import sys
from collections import Counter
from datetime import datetime

from .comparaison import comparer_dossiers
from .config import charger_configuration
from .moteur import (EnteteInvalide, archiver_entree, ecrire_fichiers,
                     formater_numero_piece, generer_ecritures,
                     generer_ecritures_paie, nettoyer_sortie, prochain_compteur)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Génère un fichier d'import ASCII QuadraCOMPTA par établissement.")
    parser.add_argument("--config", required=True,
                        help="Chemin du fichier de configuration YAML")
    parser.add_argument("--reference", nargs="?", const="reference", default=None,
                        help="Dossier de référence à comparer (défaut « reference » "
                             "si l'option est passée sans valeur)")
    args = parser.parse_args(argv)

    cfg = charger_configuration(args.config)
    pretes = [s for s in cfg.sources if s.complete]
    attente_simple = [s.libelle for s in cfg.sources if not s.complete]

    # Archivage optionnel du dossier entrée (confort, jamais bloquant) : un ZIP
    # horodaté à la seconde, avant toute lecture des sources.
    if cfg.dossier_archives or cfg.archiver_entree:
        dossier_archives = cfg.dossier_archives or os.path.join(
            os.path.dirname(os.path.normpath(cfg.dossier_entree)), "archives")
        try:
            chemin = archiver_entree(cfg.dossier_entree, dossier_archives,
                                     datetime.now().strftime("%Y%m%d_%H%M%S"))
            print(f"  Archive de « entree » : {chemin}" if chemin
                  else "  Archivage : dossier entrée vide ou absent — ignoré.")
        except Exception as e:                       # disque plein, droits…
            print(f"  !! Archivage impossible (ignoré) : {e}")

    # Purge des fichiers générés d'un run précédent (jamais les autres fichiers),
    # avant toute écriture, pour ne pas réimporter un dossier devenu orphelin.
    nettoyer_sortie(cfg.dossier_sortie)

    # N° de pièce incrémental : compteur de run accolé au n° de pièce de base.
    # Calculé une seule fois pour que toutes les écritures du run (passe normale
    # ET contre-passation) portent le même numéro. Le compteur est rangé à côté
    # du fichier de configuration (dossier stable, non purgé comme la sortie).
    if cfg.numero_piece_incremental and cfg.numero_piece:
        dossier_config = os.path.dirname(os.path.abspath(args.config))
        chemin_compteur = os.path.join(dossier_config, "compteur_import.txt")
        cfg.numero_piece = formater_numero_piece(cfg.numero_piece,
                                                 prochain_compteur(chemin_compteur))
        print(f"  N° de pièce de ce run : {cfg.numero_piece}")

    print("Écritures d'arrêté :")
    centres_invalides: list = []
    doublons_paie: list = []
    try:
        par_dossier, sans_centre = generer_ecritures(pretes, cfg,
                                                     centres_inconnus=centres_invalides)
        par_paie, centres_inconnus, attente_paie = generer_ecritures_paie(
            cfg.sources_paie, cfg, doublons=doublons_paie)
    except EnteteInvalide as e:
        print(f"\n  !! ERREUR DE STRUCTURE — génération interrompue :\n     {e}")
        return 2
    for dossier, lignes in par_paie.items():
        par_dossier[dossier].extend(lignes)

    td, tc, deseq = ecrire_fichiers(par_dossier, cfg.dossier_sortie)
    for code, d, c in deseq:
        print(f"  !! DÉSÉQUILIBRE dossier {code} : débit {d/100:.2f} / crédit {c/100:.2f}")
    print(f"  {len(par_dossier)} fichier(s) — total débit {td/100:.2f} € / crédit {tc/100:.2f} €")

    if sans_centre:
        print("\n  Centre analytique manquant (lignes I non générées, à compléter) :")
        for dossier, libelle, fichier in sorted(set(sans_centre)):
            print(f"   - dossier {dossier} : {libelle} (fichier : {fichier})")
    if centres_invalides:
        print("\n  Centres analytiques inconnus (à vérifier — écritures produites) :")
        for centre, dossier, libelle, fichier in sorted(set(centres_invalides)):
            print(f"   - centre {centre} (dossier {dossier}) : {libelle} — {fichier}")
    if centres_inconnus:
        print("\n  Centres de coût non rattachés à un dossier (écritures non générées) :")
        for centre, fichier in centres_inconnus:
            print(f"   - {centre} (fichier : {fichier})")
    if attente_simple or attente_paie:
        print("\n  Sources en attente de comptes (non générées) :")
        for lib in attente_simple + list(attente_paie):
            print("   -", lib)
    if doublons_paie:
        print("\n  Doublons potentiels détectés (à vérifier) :")
        for matricule, centre, fichiers in sorted(doublons_paie):
            parties = [f"{f} (x{n})" if n > 1 else f
                       for f, n in sorted(Counter(fichiers).items())]
            print(f"   - matricule {matricule}, centre {centre} : {', '.join(parties)}")

    if any(s.contre_passation for s in cfg.sources + cfg.sources_paie):
        print("\nContre-passations :")
        ex, _ = generer_ecritures([s for s in pretes if s.contre_passation], cfg, extourne=True)
        ex_p, _, _ = generer_ecritures_paie(
            [s for s in cfg.sources_paie if s.contre_passation], cfg, extourne=True)
        for dossier, lignes in ex_p.items():
            ex[dossier].extend(lignes)
        td2, tc2, deseq2 = ecrire_fichiers(ex, cfg.dossier_sortie, suffixe="_contrepass")
        for code, d, c in deseq2:
            print(f"  !! DÉSÉQUILIBRE dossier {code} : débit {d/100:.2f} / crédit {c/100:.2f}")
        print(f"  {len(ex)} fichier(s) — total débit {td2/100:.2f} € / crédit {tc2/100:.2f} €")

    # Comparaison optionnelle avec une version de référence (CLI > config).
    reference = args.reference or cfg.dossier_reference
    if reference:
        chemin_csv = os.path.join(
            cfg.dossier_sortie, f"diff_situation_{datetime.now():%Y%m%d}.csv")
        synth = comparer_dossiers(reference, cfg.dossier_sortie, chemin_csv)
        if synth is None:
            print(f"\nComparaison : référence « {reference} » absente ou vide — ignorée.")
        else:
            print(f"\nComparaison avec « {reference} » -> {chemin_csv}")
            print(f"  {synth.dossiers} dossier(s) touché(s) ; "
                  f"{synth.nouvelles} nouvelle(s), {synth.supprimees} supprimée(s), "
                  f"{synth.modifiees} modifiée(s)")
            print(f"  total avant {synth.total_avant/100:.2f} € / "
                  f"après {synth.total_apres/100:.2f} € / "
                  f"écart {synth.ecart/100:.2f} €")

    print(f"\nTerminé. Fichiers dans : {cfg.dossier_sortie}")
    return 1 if deseq else 0


if __name__ == "__main__":
    sys.exit(main())
