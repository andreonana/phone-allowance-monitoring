"""
Résolution des chemins de données de l'application.

L'application ne dépend d'aucune base de données : toutes les données
persistantes (base de référence, historique, configuration) sont des
fichiers Excel stockés dans un dossier `data/` situé à côté de l'exécutable
(ou du script principal en développement).

Ce module centralise cette résolution pour rester portable entre
Windows (poste cible de l'utilisateur), macOS et Linux (développement),
et pour fonctionner aussi bien en exécution "python main.py" qu'une fois
empaqueté en .exe avec PyInstaller (auquel cas sys.frozen est vrai et les
données doivent être stockées à côté de l'exécutable, pas dans le dossier
temporaire d'extraction de PyInstaller).
"""

from __future__ import annotations

import sys
from pathlib import Path


def get_app_root() -> Path:
    """Retourne le dossier racine de l'application (où vit main.py ou le .exe)."""
    if getattr(sys, "frozen", False):
        # Exécutable PyInstaller : les données doivent vivre à côté du .exe,
        # pas dans le dossier temporaire d'extraction (sys._MEIPASS).
        return Path(sys.executable).resolve().parent
    # Exécution depuis les sources : racine du projet = parent de app/
    return Path(__file__).resolve().parent.parent


def get_data_dir() -> Path:
    """Dossier contenant les fichiers Excel persistants de l'application."""
    data_dir = get_app_root() / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def get_exports_dir() -> Path:
    """Dossier où sont écrits les rapports Excel exportés."""
    exports_dir = get_data_dir() / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)
    return exports_dir


def get_reference_base_path() -> Path:
    return get_data_dir() / "reference_base.xlsx"


def get_history_path() -> Path:
    return get_data_dir() / "history.xlsx"


def get_logs_dir() -> Path:
    logs_dir = get_data_dir() / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir
