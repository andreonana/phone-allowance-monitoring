"""
Génère (ou régénère) `data/reference_base.xlsx` à partir d'un classeur Excel
"DRAFT" qui contient déjà une colonne STAFF ID / matricule, avec un onglet
ORANGE et un onglet MTN (ex: le classeur fourni par le client).

Usage :
    python scripts/generate_reference_base.py "/chemin/vers/DRAFT.xlsx"

Sans argument, utilise le chemin par défaut ci-dessous.

Ce script est le point d'entrée "simple" pour remplir la base de référence :
pas de saisie manuelle ligne par ligne dans l'IHM, on repart d'un classeur
Excel existant. Réutilise `ReferenceRepository.bootstrap_from_operator_file`
(app/reference_repository.py) pour les feuilles ORANGE puis MTN.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.models import Operateur
from app.paths import get_reference_base_path
from app.reference_repository import ReferenceRepository

DEFAULT_SOURCE = "/Users/fayacomputer/Downloads/DRAFT PHONE ALLOWANCE MONITORING V1.xlsx"


def main() -> None:
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(DEFAULT_SOURCE)
    if not source.exists():
        print(f"Fichier source introuvable : {source}")
        sys.exit(1)

    destination = get_reference_base_path()
    if destination.exists():
        print(f"ATTENTION : {destination} existe déjà et va être écrasé/complété.")

    repo = ReferenceRepository(destination)

    total = {"nb_numeros_ajoutes": 0, "nb_ignores": 0, "nb_collaborateurs_crees": 0}
    for operateur in (Operateur.ORANGE, Operateur.MTN):
        resume = repo.bootstrap_from_operator_file(str(source), operateur)
        print(f"[{operateur.value}] {resume}")
        for cle in total:
            total[cle] += resume[cle]

    repo.save()
    print(f"\nTotal : {total}")
    print(f"{len(repo.collaborateurs)} collaborateur(s), {len(repo.numeros)} numéro(s) au total.")
    print(f"Écrit dans : {destination}")


if __name__ == "__main__":
    main()
