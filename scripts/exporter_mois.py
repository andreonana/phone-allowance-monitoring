"""
Consolide un mois (Orange + MTN, comparaison au mois précédent, alertes) et
exporte le rapport Excel final dans `data/exports/`. Équivalent en ligne de
commande du bouton "Exporter" de l'IHM.

Usage :
    python scripts/exporter_mois.py 2026-07
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.alerts_engine import generer_alertes
from app.config import ConfigRepository
from app.consolidation import consolider
from app.history_repository import HistoryRepository
from app.paths import get_data_dir, get_history_path, get_reference_base_path
from app.reference_repository import ReferenceRepository
from app.report_export import exporter_rapport


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage : python scripts/exporter_mois.py AAAA-MM")
        sys.exit(1)
    mois = sys.argv[1]

    reference = ReferenceRepository(get_reference_base_path())
    history = HistoryRepository(get_history_path())
    config = ConfigRepository(get_data_dir() / "config.xlsx")

    syntheses = consolider(mois, reference, history)
    alertes = generer_alertes(mois, syntheses, history, reference, config)
    chemin = exporter_rapport(mois, syntheses, alertes, history)

    print(f"{len(syntheses)} collaborateur(s) dans la synthèse, {len(alertes)} alerte(s).")
    print(f"Rapport écrit dans : {chemin}")


if __name__ == "__main__":
    main()
