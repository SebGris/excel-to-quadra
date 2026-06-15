# excel-to-quadra

Génération de fichiers d'import **ASCII QuadraCOMPTA (Cegid)** à partir de
classeurs Excel de provisions de situation comptable — un fichier texte par
établissement (dossier Quadra), avec lignes analytiques et contrôle
d'équilibre débit/crédit.

Conçu pour un contexte multi-établissements (réseau de ~60 structures,
plusieurs entités juridiques) où chaque établissement correspond à un dossier
comptable distinct, mais **entièrement piloté par configuration** : aucun
référentiel d'organisation n'est codé en dur.

## Fonctionnalités

- **Sources déclaratives** : chaque colonne de montant d'un classeur Excel est
  décrite en YAML (feuille, colonnes, comptes, libellé, journal, date). Le
  moteur n'a jamais besoin d'être modifié pour ajouter une provision.
- **Deux familles de sources** :
  - classeurs « une ligne = un établissement » (provisions, produits) ;
  - classeurs de **paie détaillés par salarié**, agrégés automatiquement par
    centre de coût et routés vers le bon dossier via une table inverse.
- **Lignes analytiques** (type I) générées sur la part de charge ou de produit
  (classes 6/7), ventilées à 100 % vers le centre du dossier.
- **Règles comptables** : montants négatifs (régularisations) inversant le sens
  débit/crédit, contre-passation optionnelle par source, exclusion des lignes
  de totaux et des montants nuls.
- **Tolérance aux données incomplètes** : une source dont un compte n'est pas
  renseigné (`XXXXXXXX`) ou un centre de coût inconnu n'interrompt pas la
  génération — l'élément est ignoré et **signalé dans le rapport final**.
- **Contrôles** : équilibre débit = crédit par dossier et global, conversion
  euros → centimes en `Decimal` (arrondi commercial, pas d'erreur de flottant).

## Format de sortie

Conforme à la spécification Quadratus « Fichier d'entrée ASCII dans
QuadraCOMPTA » :

| Enregistrement | Longueur | Contenu |
|---|---|---|
| `M` (écriture) | 146 c. | compte (8), journal, folio `000`, date `JJMMAA`, libellé (20), sens `D`/`C`, montant en centimes (13) |
| `I` (analytique) | 39 c. | % de répartition, montant en centimes, code centre (10) — suit immédiatement sa ligne `M` |

Encodage **Windows-1252**, fins de ligne **CRLF**.

## Installation

```bash
git clone https://github.com/SebGris/excel-to-quadra.git
cd excel-to-quadra
python -m venv .venv && .venv\Scripts\activate    # Windows
pip install -e ".[dev]"
```

## Utilisation

1. Copier `config/exemple_situation.yaml` vers `config/situation.local.yaml`
   (les fichiers `*.local.yaml` sont exclus du dépôt par `.gitignore`).
2. Renseigner chemins, table analytique, sources et comptes.
3. Lancer :

```bash
excel-to-quadra --config config/situation.local.yaml
# ou
python -m excel_to_quadra.cli --config config/situation.local.yaml
```

Le rapport de fin d'exécution affiche le nombre de fichiers générés, les
totaux débit/crédit, puis les éventuels éléments à compléter (centres
analytiques manquants, centres de coût inconnus, sources en attente de
comptes). Le code de retour est non nul si un déséquilibre est détecté.

Pour faciliter la correction, les signalements « centre de coût non rattaché à
un dossier » et « centre analytique manquant » précisent le **fichier source**
concerné.

Une section **« Centres analytiques inconnus (à vérifier) »** liste les centres
produits sur une ligne `I` mais absents de l'ensemble des centres connus (union
de la table `analytique`, des `centres_supplementaires` et des `ventilation`).
C'est un **avertissement** : l'écriture est tout de même produite et le code de
retour reste `0` (à la différence du centre *manquant*, qui concerne l'absence
de centre, pas un centre invalide).

Au démarrage, le dossier de sortie est **purgé des fichiers générés d'un run
précédent** (motif `*_ecriture_Quadra*.txt` uniquement ; tout autre fichier est
préservé), afin qu'un dossier devenu orphelin — supprimé, renommé ou aliasé — ne
soit pas réimporté par erreur.

### Utilisation par double-clic (Windows, sans droits admin)

Pour un poste sans environnement Python préparé, deux scripts évitent la ligne
de commande :

1. **`Installer.bat`** — à lancer **une seule fois** : vérifie la présence de
   Python et installe le programme via `pip install --user` (aucun droit
   administrateur requis).
2. **`Lancer.bat`** — à double-cliquer **à chaque situation** : lance la
   génération sur `config/situation.local.yaml` et garde la fenêtre ouverte
   pour afficher le rapport.

### Options de configuration des sources

Outre les champs de base (feuille, colonnes, comptes, libellé, journal, date),
une source « une ligne = un établissement » accepte plusieurs options :

- **`agreger`** : cumule toutes les lignes d'un même dossier avant émission,
  pour ne produire qu'une seule écriture par dossier (au lieu d'une par ligne).
- **`facteur`** : multiplicateur appliqué au montant avant le calcul du sens
  débit/crédit (ex. lissage de charges sur 7 mois : `facteur: 0.5833` = 7/12).
- **`ventilation`** : répartit la ligne analytique d'un dossier sur plusieurs
  centres (clé = dossier, valeur = liste de `{centre, pourcent}`). La ligne `M`
  de classe 6/7 est alors suivie d'une ligne `I` par centre ; la dernière reçoit
  le solde pour que la somme des lignes `I` égale exactement la ligne `M`.
  Prioritaire sur le centre par défaut, et compatible avec `agreger` / `facteur`.
- **`col_date` / `date_min` / `date_max`** : filtre de dates optionnel. Quand
  `col_date` (lettre de colonne) et au moins une borne (`date_min` / `date_max`,
  format `AAAAMMJJ`, incluses) sont renseignés, seules les lignes dont la date
  tombe dans la période sont retenues — utile pour un exercice non encore
  clôturé dont l'export mélange plusieurs années. Le filtre s'applique **avant**
  `agreger`, donc les cumuls ne portent que sur la période.

```yaml
sources:
  - fichier: "Charges_a_lisser.xlsx"
    feuille: "Feuil1"
    ligne_debut: 2
    col_dossier: "B"
    col_montant: "H"
    compte_debit: "64140100"    # charge (classe 6) -> ligne analytique
    compte_credit: "42868000"
    libelle: "Lissage charges"
    journal: "OS"
    date_ecriture: "311226"
    agreger: true               # une seule écriture par dossier sur le cumul
    facteur: 0.5833333333       # multiplicateur du montant (7/12)
    col_date: "G"               # colonne contenant la date de l'écriture
    date_min: "20260101"        # bornes AAAAMMJJ incluses : ne garder que 2026
    date_max: "20261231"
    ventilation:                # répartition analytique multi-centres
      "702":
        - {centre: "770201", pourcent: 59.04}
        - {centre: "770202", pourcent: 13.34}
        - {centre: "770203", pourcent: 27.62}
```

### Alias de dossiers (option globale)

`alias_dossiers` (clé de premier niveau) redirige un code dossier lu vers un
autre : les écritures sont produites dans le dossier **cible**, avec le centre
analytique de la cible. Contrairement au `remap`, le code lu n'est pas conservé
comme centre.

```yaml
alias_dossiers:
  "736": "723"   # tout 736 lu est comptabilisé dans le dossier 723
```

## Structure du projet

```
excel-to-quadra/
├── pyproject.toml              # métadonnées, dépendances, point d'entrée CLI
├── Installer.bat               # installation par double-clic (Windows, pip --user)
├── Lancer.bat                  # génération par double-clic (Windows)
├── config/
│   └── exemple_situation.yaml  # configuration d'exemple commentée
├── src/excel_to_quadra/
│   ├── format_quadra.py        # enregistrements M et I (spécification Quadra)
│   ├── normalisation.py        # codes établissement et montants
│   ├── config.py               # dataclasses + chargement/validation YAML
│   ├── moteur.py               # génération, agrégation paie, équilibre, sortie
│   └── cli.py                  # point d'entrée en ligne de commande
└── tests/
    ├── test_format_quadra.py   # positions des enregistrements au caractère près
    ├── test_normalisation.py   # formats de codes, montants
    ├── test_moteur.py          # paires, analytique, négatifs, extourne
    ├── test_config.py          # chargement YAML, table inverse, validation
    └── test_integration.py     # chaîne complète sur classeurs générés à la volée
```

## Tests

```bash
pytest          # 91 tests
pytest -v       # détail
```

Les tests d'intégration créent leurs propres classeurs Excel dans un
répertoire temporaire : la suite est autonome, aucune donnée réelle n'est
nécessaire ni incluse.

## Limitations connues

- La lecture s'appuie sur `openpyxl` avec `data_only=True` : les classeurs
  doivent avoir été enregistrés par Excel pour que les valeurs des formules
  soient disponibles.
- Les noms de fichiers sources contiennent souvent des dates : la
  configuration est à mettre à jour à chaque situation.
- Enregistrements Quadra de type `C` (comptes), `R` (règlements) et `Y`/`X`/`Z`
  non gérés.

## Licence

MIT — voir [LICENSE](LICENSE).
