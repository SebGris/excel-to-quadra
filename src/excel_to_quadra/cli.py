# -*- coding: utf-8 -*-
"""Point d'entrée en ligne de commande.

Usage :
    excel-to-quadra --config config/situation.local.yaml
    python -m excel_to_quadra.cli --config config/situation.local.yaml
"""

import argparse
import sys

from .config import charger_configuration
from .moteur import (ecrire_fichiers, generer_ecritures, generer_ecritures_paie,
                     nettoyer_sortie)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Génère un fichier d'import ASCII QuadraCOMPTA par établissement.")
    parser.add_argument("--config", required=True,
                        help="Chemin du fichier de configuration YAML")
    args = parser.parse_args(argv)

    cfg = charger_configuration(args.config)
    pretes = [s for s in cfg.sources if s.complete]
    attente_simple = [s.libelle for s in cfg.sources if not s.complete]

    # Purge des fichiers générés d'un run précédent (jamais les autres fichiers),
    # avant toute écriture, pour ne pas réimporter un dossier devenu orphelin.
    nettoyer_sortie(cfg.dossier_sortie)

    print("Écritures d'arrêté :")
    centres_invalides: list = []
    par_dossier, sans_centre = generer_ecritures(pretes, cfg,
                                                 centres_inconnus=centres_invalides)
    par_paie, centres_inconnus, attente_paie = generer_ecritures_paie(cfg.sources_paie, cfg)
    for dossier, lignes in par_paie.items():
        par_dossier[dossier].extend(lignes)

    td, tc, deseq = ecrire_fichiers(par_dossier, cfg.dossier_sortie)
    for code, d, c in deseq:
        print(f"  !! DÉSÉQUILIBRE dossier {code} : débit {d/100:.2f} / crédit {c/100:.2f}")
    print(f"  {len(par_dossier)} fichier(s) — total débit {td/100:.2f} € / crédit {tc/100:.2f} €")

    if sans_centre:
        print("\n  Centre analytique manquant (lignes I non générées, à compléter) :")
        for dossier, libelle in sorted(set(sans_centre)):
            print(f"   - dossier {dossier} : {libelle}")
    if centres_invalides:
        print("\n  Centres analytiques inconnus (à vérifier — écritures produites) :")
        for centre, dossier, libelle, fichier in sorted(set(centres_invalides)):
            print(f"   - centre {centre} (dossier {dossier}) : {libelle} — {fichier}")
    if centres_inconnus:
        print("\n  Centres de coût non rattachés à un dossier (écritures non générées) :")
        for centre in centres_inconnus:
            print("   -", centre)
    if attente_simple or attente_paie:
        print("\n  Sources en attente de comptes (non générées) :")
        for lib in attente_simple + list(attente_paie):
            print("   -", lib)

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

    print(f"\nTerminé. Fichiers dans : {cfg.dossier_sortie}")
    return 1 if deseq else 0


if __name__ == "__main__":
    sys.exit(main())
