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

Une section **« Doublons potentiels détectés (à vérifier) »** signale, pour les
sources de paie, les couples (matricule, centre de coût) présents plus d'une
fois — dans un même fichier ou dans plusieurs — avec les fichiers concernés. Un
tel doublon double la provision d'un salarié **sans rompre l'équilibre**, donc
sans être détecté par le contrôle débit/crédit. Le matricule est lu dans la
colonne `col_matricule` de la source de paie (défaut `G`, **surchargeable par
source** — ex. `H` pour les fichiers STC dont la colonne `G` est une catégorie ;
`col_matricule: null` **désactive** la détection pour un fichier sans matricule
individuel). Un même matricule sur des centres *différents* (salarié réparti)
n'est pas un doublon. C'est un **avertissement** non bloquant (code inchangé).

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
- **`entete_attendu`** (+ `ligne_entete`) : contrôle de structure. Dictionnaire
  `colonne → libellé attendu` vérifié **avant** de traiter la source : chaque
  cellule de la ligne d'en-tête (`ligne_entete`, défaut `ligne_debut - 1`) doit
  contenir le libellé attendu (comparaison insensible à la casse et aux espaces
  de début/fin). Tout écart **interrompt** la génération avec une erreur claire
  (fichier, colonne, attendu, trouvé) : c'est un garde-fou contre un classeur
  qui ne correspond pas à ce que la config croit lire (colonnes décalées, mauvais
  onglet…). Absent, aucun contrôle (comportement inchangé).

```yaml
sources_paie:
  - fichier: "CRE PRECA 310526.xlsx"
    feuille: "Feuil1"
    ligne_debut: 6
    col_centre: "D"
    col_matricule: "G"        # ou "H" (STC), ou null pour désactiver la détection
    entete_attendu:           # vérifié en ligne 5 (ligne_debut - 1)
      D: "centre de coût"
      G: "Mat-Id-Es"
      N: "Montant Prime Précarité"
    # ... composantes ...
```

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

### Numéro de pièce (option globale, surchargeable par source)

`numero_piece` (clé de premier niveau) inscrit un n° de pièce sur chaque ligne
`M` générée (position 100, 8 caractères ; les lignes `I` analytiques n'en
portent pas). Une source peut le **surcharger** via son propre champ
`numero_piece` (la valeur de la source prime sur la globale).

Usage : dans un journal partagé (ex. `OS`), cela distingue les écritures
**importées** par le programme (avec n° de pièce) de celles **saisies
manuellement** (quasi toujours sans n° de pièce). Absent, le champ reste à
blanc — les fichiers existants ne changent pas d'un octet.

```yaml
numero_piece: "IMPORT"      # appliqué à toutes les écritures générées
sources:
  - fichier: "Cas_particulier.xlsx"
    # ... champs de base ...
    numero_piece: "IMPORT2" # surcharge pour cette source uniquement
```

#### Compteur de génération incrémental

Avec `numero_piece_incremental: true`, un compteur sur 2 chiffres est accolé au
n° de pièce de base : `IMPORT01`, `IMPORT02`… Le compteur **s'incrémente à
chaque exécution** du programme, et **toutes les écritures d'un même run**
(passe normale *et* contre-passation) portent le même numéro — pratique pour
isoler un import donné dans le journal.

Le dernier compteur utilisé est persisté dans `compteur_import.txt`, **à côté du
fichier de configuration** (dossier stable, contrairement à la sortie qui est
purgée à chaque run) ; absent, vide ou corrompu, on repart de `1`.

Format : base + compteur, tronqué à **8 caractères** (donc base de 6 c. max pour
un compteur 2 chiffres : `IMPORT` → `IMPORT01`). Au-delà de 99 le compteur passe
à 3 chiffres et c'est la **base** qui est rognée (`IMPORT` → `IMPOR100`), jamais
le compteur. Sans l'option, le n° de pièce reste fixe (comportement par défaut).

```yaml
numero_piece: "IMPORT"
numero_piece_incremental: true   # IMPORT01, IMPORT02, … (un n° par exécution)
```

### Comparaison avec une version de référence

Pour repérer ce qui a changé entre deux générations (correction d'un classeur,
nouvelle clé de ventilation…), le programme peut comparer la sortie courante à
une **version de référence** et produire un **rapport CSV des différences**.

Workflow :

1. Déposer les `*_ecriture_Quadra*.txt` de la génération précédente dans un
   dossier de référence (ex. `reference/`).
2. Relancer la génération en activant la comparaison, soit en ligne de commande
   `excel-to-quadra --config … --reference reference`, soit via la clé de config
   `dossier_reference: "reference"` (l'option CLI a priorité).
3. Lire `diff_situation_AAAAMMJJ.csv` dans le dossier de sortie.

La comparaison se fait au niveau de chaque **écriture M**, identifiée par
(dossier, compte, sens, libellé) ; la valeur comparée est le montant. Le **n° de
pièce** (position 100, incrémental) **ne fait pas partie de la clé** — sinon
toutes les écritures sembleraient « modifiées » d'une génération à l'autre. Chaque
différence est classée **NOUVELLE**, **SUPPRIMEE** ou **MONTANT_MODIFIE**
(avec `montant_avant`, `montant_apres`, `ecart`) ; les écritures identiques ne
figurent pas dans le rapport. Le CSV (`;`, UTF-8 BOM, montants en euros) se
termine par un récapitulatif (dossiers touchés, compteurs, totaux avant/après et
écart global), également affiché à l'écran.

Sans référence (ni option ni clé), aucune comparaison — comportement inchangé.
Référence absente ou vide : message informatif, run normal (non bloquant).

### Archivage automatique du dossier d'entrée

Au démarrage de chaque génération, le programme peut archiver les classeurs du
dossier `entree` dans un **ZIP horodaté** `entree_AAAAMMJJ_HHMMSS.zip` (à la
seconde, pour qu'un nouvel essai le même jour n'écrase pas le précédent) — utile
pour conserver l'état exact des sources ayant produit un import.

Activation : renseigner `dossier_archives` (dossier cible) **ou** `archiver_entree:
true` (archive alors dans `archives/`, à côté de `entree`). Sans l'une de ces
options, aucun archivage (comportement inchangé). Seuls les fichiers à la racine
de `entree` sont archivés ; `entree` vide ou absent → pas d'archive. L'archivage
est un **confort, jamais bloquant** : en cas d'échec (disque plein, droits), un
avertissement est émis et la génération continue.

```yaml
dossier_archives: "archives"   # ou : archiver_entree: true
```

> ⚠️ Les archives contiennent des **données de paie** : elles ne doivent **pas**
> être versionnées. Le `.gitignore` couvre déjà `*.zip`.

### Périmètre restreint à une source (`--source`)

`--source <nom_ou_motif>` restreint la génération aux sources dont le `fichier:`
correspond (sous-chaîne **ou** motif glob, insensible à la casse), exactement
comme si seul ce fichier était présent dans `entree` : seuls les
dossiers/établissements alimentés par cette source ont un `.txt` en sortie.
Tous les contrôles (équilibre compris) s'appliquent au périmètre restreint.
Sans l'option, **toutes** les sources sont traitées (comportement inchangé).
La clé config `filtre_source` joue le même rôle (l'option CLI est prioritaire).

Usage : permettre aux collègues ayant déjà importé la version précédente de
n'importer que les écritures du **nouveau** fichier, sans doublon. Combiner avec
`--output <dossier>` pour écrire dans un dossier distinct et **ne pas écraser**
la génération complète :

```bash
excel-to-quadra --config config/situation.local.yaml \
    --source "*PRECA*" --output sortie_preca
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
│   ├── comparaison.py          # diff CSV courant vs référence
│   ├── normalisation.py        # codes établissement et montants
│   ├── config.py               # dataclasses + chargement/validation YAML
│   ├── moteur.py               # génération, agrégation paie, équilibre, sortie
│   └── cli.py                  # point d'entrée en ligne de commande
└── tests/
    ├── test_format_quadra.py   # positions des enregistrements au caractère près
    ├── test_normalisation.py   # formats de codes, montants
    ├── test_moteur.py          # paires, analytique, négatifs, extourne
    ├── test_comparaison.py     # diff courant vs référence, rapport CSV
    ├── test_config.py          # chargement YAML, table inverse, validation
    └── test_integration.py     # chaîne complète sur classeurs générés à la volée
```

## Tests

```bash
pytest          # 142 tests
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
