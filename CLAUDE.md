# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Vue d'ensemble

`excel-to-quadra` génère des fichiers d'import **ASCII QuadraCOMPTA (Cegid)** à
partir de classeurs Excel de provisions de situation comptable : un fichier
texte par établissement (dossier Quadra), avec lignes analytiques et contrôle
d'équilibre débit/crédit. Tout le référentiel métier (chemins, comptes, tables
analytiques, sources) est **externalisé en YAML** — le code Python reste
générique et ne contient aucune donnée d'organisation en dur.

Le code, les commentaires et les libellés sont en **français** : conserver cette
langue dans le code et les messages.

## Commandes

```bash
# Installation (Windows)
python -m venv .venv && .venv\Scripts\activate
pip install -e ".[dev]"

# Exécution
excel-to-quadra --config config/situation.local.yaml
python -m excel_to_quadra.cli --config config/situation.local.yaml

# Tests
pytest                       # suite complète (autonome, ~62 tests)
pytest -v                    # détail
pytest tests/test_moteur.py  # un seul fichier
pytest tests/test_moteur.py::test_nom_du_test   # un seul test
pytest -k "extourne"         # filtrer par mot-clé
```

Il n'y a pas de linter ni de formateur configuré dans le projet.

## Architecture

Pipeline en une passe, orchestré par [cli.py](src/excel_to_quadra/cli.py) :
**config YAML → lecture Excel → génération des paires d'écritures → contrôle
d'équilibre → écriture d'un fichier texte par dossier**.

- **[config.py](src/excel_to_quadra/config.py)** — dataclasses (`Source`,
  `SourcePaie`, `Composante`, `Configuration`) et chargement/validation YAML.
  Construit `centre_vers_dossier` (table inverse centre→dossier) à partir de
  `analytique` (dossier→centre) plus les `centres_supplementaires`. Un compte
  valant le marqueur `A_RENSEIGNER` (`"XXXXXXXX"`) rend une source/composante
  `incomplete` : elle est ignorée et signalée, sans bloquer le reste.

- **[normalisation.py](src/excel_to_quadra/normalisation.py)** — `normaliser_code`
  (extrait/nettoie le code dossier ; écarte totaux et non-numériques ; règle
  spéciale `7xxx → xxx` sur 4 chiffres commençant par 7) et `lire_montant`
  (ignore vides, nuls, texte, booléens).

- **[moteur.py](src/excel_to_quadra/moteur.py)** — cœur de la logique.
  - `ajouter_ecriture_pair` produit chaque écriture en **paire équilibrée**
    (un M débit + un M crédit). La part de charge/produit (compte de classe 6
    ou 7) est écrite en premier et **immédiatement suivie de sa ligne I**
    analytique (ventilée à 100 % vers le centre du dossier).
  - Inversion du sens débit/crédit : `inverser = extourne ^ (montant < 0)` —
    extourne (contre-passation) et montant négatif (régularisation) se cumulent
    par XOR, donc une extourne d'un négatif retrouve le sens d'origine.
  - `generer_ecritures` traite les sources « une ligne = un établissement » ;
    `generer_ecritures_paie` traite les classeurs de paie détaillés par salarié,
    **agrégés par centre de coût** puis routés via la table inverse (le centre
    de coût sert directement de centre analytique).
  - Éléments non générables (centre analytique manquant, centre de coût inconnu,
    source en attente de comptes) sont **collectés et renvoyés** pour le rapport
    final, jamais une exception.

- **[format_quadra.py](src/excel_to_quadra/format_quadra.py)** — mise en forme
  des enregistrements à position fixe selon la spec Quadratus : ligne **M**
  (146 c.) et ligne **I** (39 c.). `euros_vers_centimes` convertit via `Decimal`
  + `ROUND_HALF_UP` (jamais de float : 0,005 € → 1 centime).

## Invariants à préserver

- **Format de sortie** : encodage **Windows-1252** (`cp1252`), fins de ligne
  **CRLF**, longueurs M=146 / I=39 au caractère près. Les tests de
  [test_format_quadra.py](tests/test_format_quadra.py) vérifient les positions
  exactes des champs — toute modification du format doit y être répercutée.
- **Équilibre** : débit = crédit par dossier et global ; `cli.main` renvoie un
  code de retour non nul en cas de déséquilibre.
- **Montants** : toujours en `Decimal`/centimes pour les calculs comptables,
  jamais d'arithmétique flottante sur les totaux.
- **Tolérance** : ne jamais interrompre la génération sur une donnée incomplète
  — ignorer l'élément et l'ajouter à la liste de signalement correspondante.

## Contrôles qualité — erreurs silencieuses

Un traitement comptable doit attraper les erreurs que l'équilibre débit/crédit
ne révèle pas. Il y a **deux cas bloquants** — le **déséquilibre débit/crédit**
(code retour 1) et l'**en-tête de fichier invalide** (`EnteteInvalide`, code
retour 2, quand `entete_attendu` ne correspond pas) ; les autres contrôles
restent **non bloquants** (avertissement dans le rapport, code retour 0). Chaque
signalement précise le **fichier source** concerné.

- **Déséquilibre débit/crédit** (bloquant) — par dossier et global.
- **Structure d'en-tête** (`entete_attendu`, bloquant) — vérifie *avant*
  traitement que les colonnes clés d'un fichier portent les libellés attendus ;
  un fichier mal structuré produirait sinon des écritures fausses en silence.
- **Centre analytique manquant** — une écriture de classe 6/7 sans centre.
- **Centre de coût non rattaché** — un centre d'un fichier source ne pointe vers
  aucun dossier (écriture non générée).
- **Centre analytique inconnu** — un centre émis n'appartient pas à l'ensemble
  des centres connus (union de la table analytique, des centres supplémentaires
  et des centres cités dans les ventilations).
- **Doublons en entrée** — une même clé métier (p. ex. matricule + centre de
  coût) présente dans plusieurs fichiers sources, ou plusieurs fois dans le même.

> **Principe directeur (à respecter pour toute évolution et tout nouvel outil) :**
> un doublon en entrée est **équilibré** en débit/crédit — il augmente d'autant
> le débit et le crédit — donc il est **invisible au contrôle d'équilibre**.
> Tout traitement qui agrège des données financières doit donc embarquer un
> contrôle de doublons dédié, distinct du contrôle d'équilibre. Ne jamais
> supposer que l'équilibre garantit l'absence de double comptage.

Enfin, **hors « contrôles » au sens strict** mais dans le même esprit qualité :
la **comparaison à une référence** (rapport diff CSV) rend compte des écarts
d'écritures entre deux générations (NOUVELLE / SUPPRIMEE / MONTANT_MODIFIE).

## Données et configuration

- Les classeurs réels vivent dans `entree/` ; les fichiers générés dans
  `sortie/`. Lecture via `openpyxl` avec `data_only=True` : les classeurs
  doivent avoir été enregistrés par Excel pour que les valeurs de formules
  soient disponibles.
- Copier [config/exemple_situation.yaml](config/exemple_situation.yaml) vers
  `config/situation.local.yaml` (les `*.local.yaml` sont git-ignorés et
  contiennent les vrais comptes/chemins). Les noms de fichiers sources
  contiennent des dates : la config est à mettre à jour à chaque situation.
- Les tests d'intégration créent leurs propres classeurs Excel dans un
  répertoire temporaire — aucune donnée réelle n'est nécessaire pour `pytest`.

## Données spécifiques à l'organisation

Les chiffres de référence (totaux attendus, codes dossiers, nom de
l'organisation) sont consignés dans `CLAUDE.local.md` — fichier **non versionné**
(git-ignoré) pour ne jamais publier de données internes. S'y reporter pour les
valeurs de non-régression sur les données réelles.
