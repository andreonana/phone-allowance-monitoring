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

## Empaqueter en exécutable Windows

```bash
pyinstaller --onefile --noconsole main.py
```

Le dossier `data/` est créé à côté de l'exécutable au premier lancement.
