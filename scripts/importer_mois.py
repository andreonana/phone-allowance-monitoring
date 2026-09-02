"""
Importe un ou deux fichiers opérateur (Orange / MTN) pour un mois donné dans
l'historique (`data/history.xlsx`), en s'appuyant sur la base de référence
(`data/reference_base.xlsx`) pour le rapprochement numéro -> matricule.

C'est exactement ce que fait le bouton "Analyser" de l'onglet Import de
l'IHM (ui/app_window.py) — ce script permet de le faire en ligne de commande,
utile pour rejouer un import (ex: après avoir complété la base de référence)
sans repasser par l'interface graphique.

Usage :
    python scripts/importer_mois.py 2026-06 \
        --orange "/chemin/vers/fichier_orange.xlsx" \
        --mtn "/chemin/vers/fichier_mtn.xlsx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.excel_io import FichierOperateurInvalide, lire_fichier_operateur
from app.history_repository import HistoryRepository
from app.models import Operateur
from app.paths import get_history_path, get_reference_base_path
from app.reconciliation_engine import rapprocher_fichier
from app.reference_repository import ReferenceRepository


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mois", help="Mois analysé, format AAAA-MM (ex: 2026-06)")
    parser.add_argument("--orange", help="Chemin du fichier Orange du mois")
    parser.add_argument("--mtn", help="Chemin du fichier MTN du mois")
    args = parser.parse_args()

    if not args.orange and not args.mtn:
        parser.error("au moins --orange ou --mtn doit être fourni")

    reference = ReferenceRepository(get_reference_base_path())
    history = HistoryRepository(get_history_path())

    for operateur, chemin in ((Operateur.ORANGE, args.orange), (Operateur.MTN, args.mtn)):
        if not chemin:
            continue
        try:
            df = lire_fichier_operateur(chemin, operateur.value)
            lignes = rapprocher_fichier(df, args.mois, operateur, reference)
            lot = history.enregistrer_lot(args.mois, operateur, Path(chemin).name, lignes)
            print(f"[{operateur.value}] {lot.nb_lignes} ligne(s) — "
                  f"{lot.nb_non_identifies} non identifiée(s), {lot.nb_invalides} invalide(s).")
        except FichierOperateurInvalide as exc:
            print(f"[{operateur.value}] ÉCHEC : {exc}")

    history.save()
    print(f"Historique sauvegardé dans : {history.chemin_fichier}")


if __name__ == "__main__":
    main()
