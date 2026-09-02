# Phone Allowance Monitoring

Suivi et rapprochement des consommations téléphoniques Orange / MTN par
collaborateur, avec alertes configurables et export d'un rapport Excel
mensuel. Aucune base de données : toute la persistance se fait via des
classeurs Excel (`.xlsx`) dans `data/`.

## Installation

**macOS : ne pas utiliser le Python fourni par les Command Line Tools
d'Apple** (`/usr/bin/python3` ou celui des CLT) — son Tk 8.5 embarqué est
cassé sur macOS récent et affiche une fenêtre blanche. Utiliser un Python
Homebrew (Tk moderne) :

```bash
brew install python@3.12 python-tk@3.12
/opt/homebrew/bin/python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows / Linux : l'installateur officiel python.org embarque déjà un Tk
correct, `python3 -m venv .venv` suffit.

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

## Lancer l'application

```bash
python main.py
```

## Lancer les tests

```bash
pytest
```

## Structure

```
app/    moteur métier (normalisation, Excel, rapprochement, alertes, export...)
ui/     interface graphique Tkinter (aucune logique métier)
data/   fichiers Excel persistants (créés automatiquement au premier lancement)
tests/  tests unitaires et d'intégration (pytest)
main.py point d'entrée
```

Voir les docstrings de chaque module dans `app/` pour le détail des règles
métier (normalisation des numéros, robustesse de lecture des fichiers
opérateurs, calcul du mois précédent, seuils d'alerte...).

## Application Windows prête à l'emploi (aucun terminal, aucun Python requis)

Un `.exe` autonome (`PhoneAllowanceMonitoring.exe`) est compilé automatiquement
par GitHub Actions (`.github/workflows/build-windows.yml`) à chaque mise à
jour du code — l'utilisateur final n'a qu'à le télécharger et double-cliquer
dessus, comme n'importe quel logiciel Windows.

**Pour récupérer le fichier :**
- Version stable (recommandé) : onglet **Releases** du dépôt GitHub, dernier
  lien `PhoneAllowanceMonitoring.exe`. Une nouvelle Release est publiée à
  chaque tag `vX.Y.Z` (ex : `git tag v1.0.0 && git push origin v1.0.0`), et
  le lien de téléchargement reste valable indéfiniment.
- Dernière version du code (à chaque push sur `main`) : onglet **Actions**
  du dépôt -> dernier run -> section **Artifacts** en bas de page. Ces
  artefacts expirent après 90 jours.

Le dossier `data/` (base de référence, historique, rapports) est créé à côté
de l'exécutable au premier lancement — copier/déplacer `PhoneAllowanceMonitoring.exe`
dans un dossier normal (ex : `Documents\Phone Allowance Monitoring\`), pas
dans un dossier temporaire de téléchargement.

**Premier lancement** : Windows affichera probablement un avertissement
SmartScreen ("Windows a protégé votre ordinateur") car l'exécutable n'est
pas signé numériquement — c'est normal pour un logiciel interne non publié
sur le Store. Cliquer sur **Informations complémentaires** puis **Exécuter
quand même**. Cette étape n'apparaît qu'une seule fois par machine.

### Compiler soi-même (optionnel, si on ne veut pas passer par GitHub Actions)

Depuis un PC Windows avec Python installé :

```bash
pip install -r requirements.txt pyinstaller
pyinstaller PhoneAllowanceMonitoring.spec
```

Produit `dist/PhoneAllowanceMonitoring.exe`. PyInstaller ne fait pas de
cross-compilation : ce build doit être lancé sous Windows pour produire un
`.exe` Windows (sous macOS/Linux, le même `.spec` produit un exécutable pour
l'OS local — utile pour vérifier que l'empaquetage fonctionne avant de
pousser le code).
