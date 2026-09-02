"""
Point d'entrée de Phone Allowance Monitoring.

Exécution en développement :  python main.py
Empaquetage Windows          :  pyinstaller --onefile --noconsole main.py
"""

from __future__ import annotations

from ui.app_window import lancer

if __name__ == "__main__":
    lancer()
