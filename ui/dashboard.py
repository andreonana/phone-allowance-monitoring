"""
Onglet "Tableau de bord" : vue d'ensemble en un coup d'œil de l'évolution de
la consommation (cartes chiffrées + courbes/diagrammes matplotlib).

Aucune logique métier propre à ce module : il ne fait qu'appeler les
fonctions déjà existantes (`app.consolidation.consolider`,
`app.alerts_engine.generer_alertes`, `HistoryRepository.totaux_globaux_par_mois`)
et met en forme leur résultat.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # noqa: E402
from matplotlib.figure import Figure  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

from app.alerts_engine import generer_alertes  # noqa: E402
from app.consolidation import consolider  # noqa: E402
from ui import theme  # noqa: E402

_FORMAT_MONTANT = FuncFormatter(lambda x, _pos: f"{x:,.0f}".replace(",", " "))


class DashboardFrame(ttk.Frame):
    def __init__(self, parent, reference, history, config_repo) -> None:
        super().__init__(parent, style="TFrame")
        self.reference = reference
        self.history = history
        self.config_repo = config_repo

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        # -- Barre du haut : sélection du mois observé ------------------ #
        entete = ttk.Frame(self)
        entete.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 6))
        ttk.Label(entete, text="Tableau de bord", style="Titre.TLabel").pack(side="left")

        droite = ttk.Frame(entete)
        droite.pack(side="right")
        ttk.Label(droite, text="Mois :").pack(side="left", padx=(0, 4))
        self.var_mois = tk.StringVar()
        self.combo_mois = ttk.Combobox(droite, textvariable=self.var_mois, state="readonly", width=10)
        self.combo_mois.pack(side="left")
        self.combo_mois.bind("<<ComboboxSelected>>", lambda _e: self.rafraichir())
        ttk.Button(droite, text="Actualiser", command=self.rafraichir).pack(side="left", padx=(8, 0))

        # -- Cartes KPI --------------------------------------------------- #
        self.cadre_cartes = ttk.Frame(self)
        self.cadre_cartes.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
        for i in range(4):
            self.cadre_cartes.columnconfigure(i, weight=1)
        self._cartes: list[dict] = []
        for i, (titre, sous_titre) in enumerate([
            ("Total mois en cours", "toutes lignes facturées"),
            ("Total mois précédent", ""),
            ("Variation", "vs mois précédent"),
            ("Alertes actives", "à examiner"),
        ]):
            self._cartes.append(self._construire_carte(self.cadre_cartes, titre, sous_titre, i))

        # -- Graphiques ---------------------------------------------------- #
        cadre_graph = ttk.Frame(self, style="Carte.TFrame")
        cadre_graph.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.figure = Figure(figsize=(9.5, 5.2), dpi=100, facecolor=theme.FOND_CARTE)
        self.ax_evolution = self.figure.add_subplot(2, 2, (1, 2))
        self.ax_repartition = self.figure.add_subplot(2, 2, 3)
        self.ax_top_variations = self.figure.add_subplot(2, 2, 4)
        self.figure.subplots_adjust(left=0.08, right=0.97, top=0.92, bottom=0.12, hspace=0.55, wspace=0.35)

        self.canvas = FigureCanvasTkAgg(self.figure, master=cadre_graph)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

        self.rafraichir()

    # ------------------------------------------------------------------ #
    def _construire_carte(self, parent, titre: str, sous_titre: str, colonne: int) -> dict:
        cadre = tk.Frame(parent, background=theme.FOND_CARTE, highlightbackground=theme.BORDURE,
                          highlightthickness=1, bd=0)
        cadre.grid(row=0, column=colonne, sticky="nsew", padx=(0 if colonne == 0 else 8, 0))
        label_titre = tk.Label(cadre, text=titre, background=theme.FOND_CARTE,
                                foreground=theme.TEXTE_ATTENUE, font=("TkDefaultFont", 9), anchor="w")
        label_titre.pack(fill="x", padx=12, pady=(10, 0))
        label_valeur = tk.Label(cadre, text="—", background=theme.FOND_CARTE,
                                 foreground=theme.TEXTE, font=("TkDefaultFont", 17, "bold"), anchor="w")
        label_valeur.pack(fill="x", padx=12, pady=(2, 0))
        label_sous = tk.Label(cadre, text=sous_titre, background=theme.FOND_CARTE,
                               foreground=theme.TEXTE_ATTENUE, font=("TkDefaultFont", 8), anchor="w")
        label_sous.pack(fill="x", padx=12, pady=(0, 10))
        return {"cadre": cadre, "valeur": label_valeur, "sous_titre": label_sous}

    def _maj_carte(self, index: int, valeur: str, sous_titre: str | None = None,
                    couleur: str | None = None) -> None:
        carte = self._cartes[index]
        carte["valeur"].configure(text=valeur, foreground=couleur or theme.TEXTE)
        if sous_titre is not None:
            carte["sous_titre"].configure(text=sous_titre)

    # ------------------------------------------------------------------ #
    def rafraichir(self, mois: str | None = None) -> None:
        """Recalcule les cartes et redessine les graphiques pour `mois`
        (dernier mois disponible par défaut). À appeler après tout import ou
        toute modification de la base de référence."""
        totaux = self.history.totaux_globaux_par_mois()
        mois_tries = sorted(totaux)

        self.combo_mois["values"] = mois_tries
        mois = mois or self.var_mois.get().strip() or (mois_tries[-1] if mois_tries else "")
        if mois and mois not in mois_tries:
            mois = mois_tries[-1] if mois_tries else ""
        self.var_mois.set(mois)

        self._maj_cartes(mois, totaux)
        self._dessiner_evolution(mois_tries, totaux)
        self._dessiner_repartition(mois, totaux)
        self._dessiner_top_variations(mois)
        self.canvas.draw_idle()

    def _maj_cartes(self, mois: str, totaux: dict[str, dict]) -> None:
        if not mois:
            for i in range(4):
                self._maj_carte(i, "—", "aucune donnée importée")
            return

        mois_prec = self.history.mois_precedent(mois)
        total_actuel = totaux.get(mois, {}).get("total", 0.0)
        total_prec = totaux.get(mois_prec, {}).get("total") if mois_prec else None

        self._maj_carte(0, f"{total_actuel:,.0f} FCFA".replace(",", " "), mois)
        self._maj_carte(1, f"{total_prec:,.0f} FCFA".replace(",", " ") if total_prec else "—",
                         mois_prec or "aucun mois antérieur")

        if total_prec:
            variation = total_actuel - total_prec
            pct = (variation / total_prec) * 100
            couleur = theme.HAUSSE if variation > 0 else (theme.BAISSE if variation < 0 else theme.TEXTE_ATTENUE)
            signe = "+" if variation >= 0 else ""
            self._maj_carte(2, f"{signe}{pct:.1f} %", f"{signe}{variation:,.0f} FCFA".replace(",", " "), couleur)
        else:
            self._maj_carte(2, "—", "pas de comparaison possible")

        try:
            syntheses = consolider(mois, self.reference, self.history)
            alertes = generer_alertes(mois, syntheses, self.history, self.reference, self.config_repo)
            nb_critiques = sum(1 for a in alertes if a.type.value in
                                ("FORTE_AUGMENTATION", "FORTE_CONSOMMATION", "DOUBLON"))
            couleur = theme.HAUSSE if nb_critiques else theme.BAISSE
            self._maj_carte(3, str(len(alertes)), f"dont {nb_critiques} critique(s)", couleur)
        except Exception:
            self._maj_carte(3, "—", "")

    # ------------------------------------------------------------------ #
    def _style_axe(self, ax) -> None:
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(theme.BORDURE)
        ax.spines["bottom"].set_color(theme.BORDURE)
        ax.tick_params(colors=theme.TEXTE_ATTENUE, labelsize=8)
        ax.set_facecolor(theme.FOND_CARTE)

    def _dessiner_evolution(self, mois_tries: list[str], totaux: dict[str, dict]) -> None:
        ax = self.ax_evolution
        ax.clear()
        self._style_axe(ax)
        ax.set_title("Évolution de la consommation par mois", loc="left",
                      color=theme.TEXTE, fontsize=10, fontweight="bold")

        if len(mois_tries) < 1:
            ax.text(0.5, 0.5, "Aucun mois importé pour l'instant", ha="center", va="center",
                     color=theme.TEXTE_ATTENUE, transform=ax.transAxes)
            return

        orange = [totaux[m]["ORANGE"] for m in mois_tries]
        mtn = [totaux[m]["MTN"] for m in mois_tries]
        total = [totaux[m]["total"] for m in mois_tries]

        ax.plot(mois_tries, total, marker="o", color=theme.COURBE_TOTAL, linewidth=2.4, label="Total")
        ax.plot(mois_tries, orange, marker="o", color=theme.COURBE_ORANGE, linewidth=1.6, label="Orange")
        ax.plot(mois_tries, mtn, marker="o", color="#B58A00", linewidth=1.6, label="MTN")
        ax.yaxis.set_major_formatter(_FORMAT_MONTANT)
        ax.set_ylim(bottom=0)
        ax.legend(loc="upper left", frameon=False, fontsize=8)
        if len(mois_tries) > 8:
            for label in ax.get_xticklabels():
                label.set_rotation(45)
                label.set_ha("right")

    def _dessiner_repartition(self, mois: str, totaux: dict[str, dict]) -> None:
        ax = self.ax_repartition
        ax.clear()
        ax.set_title("Orange vs MTN", loc="left", color=theme.TEXTE, fontsize=10, fontweight="bold")

        valeurs_mois = totaux.get(mois, {"ORANGE": 0.0, "MTN": 0.0})
        orange, mtn = valeurs_mois["ORANGE"], valeurs_mois["MTN"]
        if orange + mtn <= 0:
            ax.axis("off")
            ax.text(0.5, 0.5, "Aucune donnée", ha="center", va="center", color=theme.TEXTE_ATTENUE)
            return

        ax.pie([orange, mtn], labels=["Orange", "MTN"], colors=[theme.COURBE_ORANGE, "#E6B800"],
               autopct="%1.0f%%", startangle=90, wedgeprops={"width": 0.45, "edgecolor": theme.FOND_CARTE},
               textprops={"fontsize": 8, "color": theme.TEXTE})

    def _dessiner_top_variations(self, mois: str) -> None:
        ax = self.ax_top_variations
        ax.clear()
        self._style_axe(ax)
        ax.set_title("Plus fortes variations (matricule)", loc="left",
                      color=theme.TEXTE, fontsize=10, fontweight="bold")

        if not mois:
            ax.text(0.5, 0.5, "Aucune donnée", ha="center", va="center",
                     color=theme.TEXTE_ATTENUE, transform=ax.transAxes)
            return

        syntheses = [s for s in consolider(mois, self.reference, self.history) if s.variation_montant]
        top = sorted(syntheses, key=lambda s: abs(s.variation_montant), reverse=True)[:8]
        if not top:
            ax.text(0.5, 0.5, "Pas de mois précédent à comparer", ha="center", va="center",
                     color=theme.TEXTE_ATTENUE, transform=ax.transAxes)
            return

        top = list(reversed(top))  # plus forte variation en haut du barh
        libelles = [s.nom or s.matricule for s in top]
        valeurs = [s.variation_montant for s in top]
        couleurs = [theme.HAUSSE if v > 0 else theme.BAISSE for v in valeurs]

        ax.barh(libelles, valeurs, color=couleurs)
        ax.axvline(0, color=theme.BORDURE, linewidth=1)
        ax.xaxis.set_major_formatter(_FORMAT_MONTANT)
        ax.tick_params(axis="y", labelsize=8)
