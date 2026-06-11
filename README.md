# quadra-ecritures

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
git clone <url-du-depot>
cd quadra-ecritures
python -m venv .venv && .venv\Scripts\activate    # Windows
pip install -e ".[dev]"
```

## Utilisation

1. Copier `config/exemple_situation.yaml` vers `config/situation.local.yaml`
   (les fichiers `*.local.yaml` sont exclus du dépôt par `.gitignore`).
2. Renseigner chemins, table analytique, sources et comptes.
3. Lancer :

```bash
quadra-ecritures --config config/situation.local.yaml
# ou
python -m quadra_ecritures.cli --config config/situation.local.yaml
```

Le rapport de fin d'exécution affiche le nombre de fichiers générés, les
totaux débit/crédit, puis les éventuels éléments à compléter (centres
analytiques manquants, centres de coût inconnus, sources en attente de
comptes). Le code de retour est non nul si un déséquilibre est détecté.

## Structure du projet

```
quadra-ecritures/
├── pyproject.toml              # métadonnées, dépendances, point d'entrée CLI
├── config/
│   └── exemple_situation.yaml  # configuration d'exemple commentée
├── src/quadra_ecritures/
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
pytest          # 62 tests
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
