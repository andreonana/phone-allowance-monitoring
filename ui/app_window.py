"""
Fenêtre principale de l'application (Tkinter, bibliothèque standard pour les
widgets + matplotlib pour les graphiques, portable Windows/Mac/Linux et
empaquetable telle quelle avec PyInstaller).

Cinq onglets :
    Tableau de bord     : cartes chiffrées + courbes d'évolution, en un coup d'œil
    Import & Analyse    : sélection des fichiers Orange/MTN du mois, rapprochement
    Base de référence   : gestion des collaborateurs et de leurs numéros
    Rapport & Alertes   : synthèse par collaborateur, alertes, export du rapport Excel
    Paramètres          : seuils d'alerte configurables
"""

from __future__ import annotations

import os
import platform
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from app.alerts_engine import generer_alertes
from app.config import ConfigRepository
from app.consolidation import consolider
from app.excel_io import FichierOperateurInvalide, lire_fichier_operateur
from app.history_repository import HistoryRepository
from app.models import Operateur, TypeCollaborateur
from app.paths import get_data_dir, get_history_path, get_reference_base_path
from app.reconciliation_engine import rapprocher_fichier
from app.reference_repository import ReferenceRepository
from app.report_export import exporter_rapport
from ui import theme
from ui.dashboard import DashboardFrame


def _ouvrir_dossier(chemin: Path) -> None:
    """Ouvre un dossier dans l'explorateur de fichiers du système."""
    systeme = platform.system()
    try:
        if systeme == "Windows":
            os.startfile(chemin)  # type: ignore[attr-defined]
        elif systeme == "Darwin":
            subprocess.run(["open", str(chemin)], check=False)
        else:
            subprocess.run(["xdg-open", str(chemin)], check=False)
    except Exception:
        pass  # ouverture du dossier = confort, ne doit jamais faire planter l'appli


class AppWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Phone Allowance Monitoring")
        self.geometry("1200x780")
        self.minsize(980, 640)
        theme.appliquer_theme(self)

        # Chargement de la couche métier (fichiers créés au premier `save()`).
        self.reference = ReferenceRepository(get_reference_base_path())
        self.history = HistoryRepository(get_history_path())
        self.config_repo = ConfigRepository(get_data_dir() / "config.xlsx")

        self._dernieres_syntheses = []
        self._dernieres_alertes = []

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.onglet_dashboard = ttk.Frame(notebook)
        self.onglet_import = ttk.Frame(notebook)
        self.onglet_reference = ttk.Frame(notebook)
        self.onglet_rapport = ttk.Frame(notebook)
        self.onglet_parametres = ttk.Frame(notebook)

        notebook.add(self.onglet_dashboard, text="  Tableau de bord  ")
        notebook.add(self.onglet_import, text="  Import & Analyse  ")
        notebook.add(self.onglet_reference, text="  Base de référence  ")
        notebook.add(self.onglet_rapport, text="  Rapport & Alertes  ")
        notebook.add(self.onglet_parametres, text="  Paramètres  ")

        self._construire_onglet_dashboard()
        self._construire_onglet_import()
        self._construire_onglet_reference()
        self._construire_onglet_rapport()
        self._construire_onglet_parametres()

    # ------------------------------------------------------------------ #
    # Onglet Tableau de bord
    # ------------------------------------------------------------------ #

    def _construire_onglet_dashboard(self) -> None:
        self.onglet_dashboard.columnconfigure(0, weight=1)
        self.onglet_dashboard.rowconfigure(0, weight=1)
        self.dashboard = DashboardFrame(self.onglet_dashboard, self.reference, self.history, self.config_repo)
        self.dashboard.grid(row=0, column=0, sticky="nsew")

    # ------------------------------------------------------------------ #
    # Onglet Import & Analyse
    # ------------------------------------------------------------------ #

    def _construire_onglet_import(self) -> None:
        cadre = self.onglet_import
        for col in range(3):
            cadre.columnconfigure(col, weight=1 if col == 1 else 0)

        ttk.Label(cadre, text="Mois analysé (AAAA-MM) :").grid(row=0, column=0, sticky="w", padx=10, pady=(15, 5))
        self.var_mois = tk.StringVar()
        ttk.Entry(cadre, textvariable=self.var_mois, width=12).grid(row=0, column=1, sticky="w", pady=(15, 5))

        self.var_fichier_orange = tk.StringVar()
        self.var_fichier_mtn = tk.StringVar()

        ttk.Label(cadre, text="Fichier Orange :").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(cadre, textvariable=self.var_fichier_orange, state="readonly").grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(cadre, text="Parcourir…",
                   command=lambda: self._choisir_fichier(self.var_fichier_orange)).grid(row=1, column=2, padx=10)

        ttk.Label(cadre, text="Fichier MTN :").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        ttk.Entry(cadre, textvariable=self.var_fichier_mtn, state="readonly").grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(cadre, text="Parcourir…",
                   command=lambda: self._choisir_fichier(self.var_fichier_mtn)).grid(row=2, column=2, padx=10)

        ttk.Button(cadre, text="Analyser", command=self._analyser).grid(row=3, column=0, sticky="w", padx=10, pady=15)

        self.texte_log_import = tk.Text(cadre, height=20, state="disabled", wrap="word")
        self.texte_log_import.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))
        cadre.rowconfigure(4, weight=1)

    def _choisir_fichier(self, variable: tk.StringVar) -> None:
        chemin = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")],
        )
        if chemin:
            variable.set(chemin)

    def _log_import(self, message: str) -> None:
        self.texte_log_import.configure(state="normal")
        self.texte_log_import.insert("end", message + "\n")
        self.texte_log_import.configure(state="disabled")
        self.texte_log_import.see("end")

    def _analyser(self) -> None:
        mois = self.var_mois.get().strip()
        if not mois:
            messagebox.showerror("Mois manquant", "Merci de saisir le mois analysé, au format AAAA-MM.")
            return

        fichiers = [
            (Operateur.ORANGE, self.var_fichier_orange.get().strip()),
            (Operateur.MTN, self.var_fichier_mtn.get().strip()),
        ]
        fichiers = [(op, chemin) for op, chemin in fichiers if chemin]
        if not fichiers:
            messagebox.showerror("Aucun fichier", "Merci de sélectionner au moins un fichier Orange ou MTN.")
            return

        self.texte_log_import.configure(state="normal")
        self.texte_log_import.delete("1.0", "end")
        self.texte_log_import.configure(state="disabled")

        for operateur, chemin in fichiers:
            try:
                df = lire_fichier_operateur(chemin, operateur.value)
                lignes = rapprocher_fichier(df, mois, operateur, self.reference)
                lot = self.history.enregistrer_lot(mois, operateur, Path(chemin).name, lignes)
                self._log_import(
                    f"[{operateur.value}] {lot.nb_lignes} ligne(s) importée(s) — "
                    f"{lot.nb_non_identifies} non identifiée(s), {lot.nb_invalides} invalide(s)."
                )
            except FichierOperateurInvalide as exc:
                self._log_import(f"[{operateur.value}] ÉCHEC : {exc}")
                messagebox.showerror(f"Fichier {operateur.value} invalide", str(exc))

        self.history.save()
        self._log_import("Historique sauvegardé.")
        self._rafraichir_mois_disponibles()
        self.dashboard.rafraichir(mois)

    # ------------------------------------------------------------------ #
    # Onglet Base de référence
    # ------------------------------------------------------------------ #

    def _construire_onglet_reference(self) -> None:
        cadre = self.onglet_reference
        cadre.columnconfigure(0, weight=1)
        cadre.rowconfigure(0, weight=1)

        colonnes = ("matricule", "nom", "direction", "fonction", "type", "actif")
        self.arbre_collaborateurs = ttk.Treeview(cadre, columns=colonnes, show="headings", height=15)
        for col, largeur in zip(colonnes, (90, 160, 120, 140, 90, 60)):
            self.arbre_collaborateurs.heading(col, text=col.upper())
            self.arbre_collaborateurs.column(col, width=largeur)
        self.arbre_collaborateurs.grid(row=0, column=0, columnspan=4, sticky="nsew", padx=10, pady=10)
        self.arbre_collaborateurs.bind("<<TreeviewSelect>>", self._selection_collaborateur)

        formulaire = ttk.LabelFrame(cadre, text="Collaborateur")
        formulaire.grid(row=1, column=0, columnspan=4, sticky="ew", padx=10, pady=5)
        for col in range(4):
            formulaire.columnconfigure(col, weight=1)

        self.var_matricule = tk.StringVar()
        self.var_nom = tk.StringVar()
        self.var_direction = tk.StringVar()
        self.var_fonction = tk.StringVar()
        self.var_type = tk.StringVar(value=TypeCollaborateur.INDIVIDUEL.value)

        champs = [
            ("Matricule", self.var_matricule), ("Nom", self.var_nom),
            ("Direction", self.var_direction), ("Fonction", self.var_fonction),
        ]
        for i, (libelle, variable) in enumerate(champs):
            ttk.Label(formulaire, text=libelle).grid(row=0, column=i, sticky="w", padx=5)
            ttk.Entry(formulaire, textvariable=variable).grid(row=1, column=i, sticky="ew", padx=5, pady=(0, 5))

        ttk.Label(formulaire, text="Type").grid(row=0, column=4, sticky="w", padx=5)
        ttk.Combobox(formulaire, textvariable=self.var_type, state="readonly",
                     values=[t.value for t in TypeCollaborateur]).grid(row=1, column=4, sticky="ew", padx=5)

        boutons = ttk.Frame(cadre)
        boutons.grid(row=2, column=0, columnspan=4, sticky="w", padx=10, pady=(0, 10))
        ttk.Button(boutons, text="Enregistrer", command=self._enregistrer_collaborateur).pack(side="left", padx=5)
        ttk.Button(boutons, text="Supprimer", command=self._supprimer_collaborateur).pack(side="left", padx=5)
        ttk.Button(boutons, text="Nouveau", command=self._vider_formulaire_collaborateur).pack(side="left", padx=5)

        numeros = ttk.LabelFrame(cadre, text="Ajouter un numéro au collaborateur sélectionné")
        numeros.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=(0, 10))

        self.var_numero_brut = tk.StringVar()
        self.var_operateur_numero = tk.StringVar(value=Operateur.ORANGE.value)

        ttk.Label(numeros, text="Numéro").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(numeros, textvariable=self.var_numero_brut, width=18).grid(row=0, column=1, padx=5)
        ttk.Label(numeros, text="Opérateur").grid(row=0, column=2, padx=5)
        ttk.Combobox(numeros, textvariable=self.var_operateur_numero, state="readonly",
                     values=[o.value for o in Operateur], width=10).grid(row=0, column=3, padx=5)
        ttk.Button(numeros, text="Ajouter", command=self._ajouter_numero).grid(row=0, column=4, padx=10)

        self._rafraichir_arbre_collaborateurs()

    def _rafraichir_arbre_collaborateurs(self) -> None:
        self.arbre_collaborateurs.delete(*self.arbre_collaborateurs.get_children())
        for c in sorted(self.reference.collaborateurs.values(), key=lambda c: c.matricule):
            self.arbre_collaborateurs.insert("", "end", iid=c.matricule, values=(
                c.matricule, c.nom, c.direction, c.fonction, c.type.value, "Oui" if c.actif else "Non",
            ))

    def _selection_collaborateur(self, _event=None) -> None:
        selection = self.arbre_collaborateurs.selection()
        if not selection:
            return
        collab = self.reference.get_collaborateur(selection[0])
        if collab is None:
            return
        self.var_matricule.set(collab.matricule)
        self.var_nom.set(collab.nom)
        self.var_direction.set(collab.direction)
        self.var_fonction.set(collab.fonction)
        self.var_type.set(collab.type.value)

    def _vider_formulaire_collaborateur(self) -> None:
        self.arbre_collaborateurs.selection_remove(self.arbre_collaborateurs.selection())
        for variable in (self.var_matricule, self.var_nom, self.var_direction, self.var_fonction):
            variable.set("")
        self.var_type.set(TypeCollaborateur.INDIVIDUEL.value)

    def _enregistrer_collaborateur(self) -> None:
        matricule = self.var_matricule.get().strip()
        if not matricule:
            messagebox.showerror("Matricule manquant", "Le matricule est obligatoire.")
            return
        self.reference.upsert_collaborateur(
            matricule, nom=self.var_nom.get().strip(), direction=self.var_direction.get().strip(),
            fonction=self.var_fonction.get().strip(), type=TypeCollaborateur(self.var_type.get()),
        )
        self.reference.save()
        self._rafraichir_arbre_collaborateurs()
        self.dashboard.rafraichir()

    def _supprimer_collaborateur(self) -> None:
        matricule = self.var_matricule.get().strip()
        if not matricule or matricule not in self.reference.collaborateurs:
            return
        if not messagebox.askyesno("Confirmer", f"Supprimer le collaborateur {matricule} et ses numéros ?"):
            return
        self.reference.supprimer_collaborateur(matricule)
        self.reference.save()
        self._vider_formulaire_collaborateur()
        self._rafraichir_arbre_collaborateurs()
        self.dashboard.rafraichir()

    def _ajouter_numero(self) -> None:
        matricule = self.var_matricule.get().strip()
        if not matricule:
            messagebox.showerror("Collaborateur manquant", "Sélectionnez ou saisissez d'abord un matricule.")
            return
        operateur = Operateur(self.var_operateur_numero.get())
        ref = self.reference.upsert_numero(self.var_numero_brut.get(), operateur, matricule)
        if ref is None:
            messagebox.showerror("Numéro invalide", f"'{self.var_numero_brut.get()}' n'est pas un numéro valide.")
            return
        self.reference.save()
        self.var_numero_brut.set("")
        messagebox.showinfo("Numéro ajouté", f"{ref.numero_normalise} rattaché à {matricule}.")

    # ------------------------------------------------------------------ #
    # Onglet Rapport & Alertes
    # ------------------------------------------------------------------ #

    def _construire_onglet_rapport(self) -> None:
        cadre = self.onglet_rapport
        cadre.columnconfigure(0, weight=1)
        cadre.rowconfigure(1, weight=3)
        cadre.rowconfigure(2, weight=2)

        entete = ttk.Frame(cadre)
        entete.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        ttk.Label(entete, text="Mois :").pack(side="left")
        self.var_mois_rapport = tk.StringVar()
        self.combo_mois_rapport = ttk.Combobox(entete, textvariable=self.var_mois_rapport, state="readonly", width=12)
        self.combo_mois_rapport.pack(side="left", padx=5)
        ttk.Button(entete, text="Générer le rapport", command=self._generer_rapport).pack(side="left", padx=10)

        # -- Synthèse par collaborateur (Orange + MTN, comparaison au mois précédent) --
        cadre_synthese = ttk.LabelFrame(cadre, text="Synthèse par collaborateur")
        cadre_synthese.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 8))
        cadre_synthese.columnconfigure(0, weight=1)
        cadre_synthese.rowconfigure(0, weight=1)

        colonnes_synthese = ("matricule", "nom", "orange", "mtn", "total_actuel",
                              "total_precedent", "variation", "variation_pct", "statut")
        entetes_synthese = ("Matricule", "Nom", "Orange", "MTN", "Total mois actuel",
                             "Total mois précédent", "Variation (FCFA)", "Variation (%)", "Statut")
        self.arbre_synthese = ttk.Treeview(cadre_synthese, columns=colonnes_synthese, show="headings")
        for col, entete_col, largeur in zip(colonnes_synthese, entetes_synthese,
                                             (90, 150, 100, 100, 130, 140, 130, 100, 90)):
            self.arbre_synthese.heading(col, text=entete_col)
            self.arbre_synthese.column(col, width=largeur, anchor="e" if col not in ("matricule", "nom", "statut") else "w")
        self.arbre_synthese.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        defilement = ttk.Scrollbar(cadre_synthese, orient="vertical", command=self.arbre_synthese.yview)
        defilement.grid(row=0, column=1, sticky="ns", pady=8)
        self.arbre_synthese.configure(yscrollcommand=defilement.set)
        theme.configurer_tags_couleur(self.arbre_synthese, theme.COULEURS_STATUT)

        # -- Alertes du mois --
        cadre_alertes = ttk.LabelFrame(cadre, text="Alertes")
        cadre_alertes.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 8))
        cadre_alertes.columnconfigure(0, weight=1)
        cadre_alertes.rowconfigure(0, weight=1)

        colonnes_alertes = ("type", "matricule", "nom", "numero", "message")
        self.arbre_alertes = ttk.Treeview(cadre_alertes, columns=colonnes_alertes, show="headings")
        for col, largeur in zip(colonnes_alertes, (170, 90, 150, 110, 420)):
            self.arbre_alertes.heading(col, text=col.upper())
            self.arbre_alertes.column(col, width=largeur)
        self.arbre_alertes.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        defilement_alertes = ttk.Scrollbar(cadre_alertes, orient="vertical", command=self.arbre_alertes.yview)
        defilement_alertes.grid(row=0, column=1, sticky="ns", pady=8)
        self.arbre_alertes.configure(yscrollcommand=defilement_alertes.set)
        theme.configurer_tags_couleur(self.arbre_alertes, theme.COULEURS_ALERTE)

        self.label_resultat_export = ttk.Label(cadre, text="")
        self.label_resultat_export.grid(row=3, column=0, sticky="w", padx=10, pady=(0, 10))

        self._rafraichir_mois_disponibles()

    def _rafraichir_mois_disponibles(self) -> None:
        mois = self.history.mois_disponibles()
        self.combo_mois_rapport["values"] = mois
        if mois and not self.var_mois_rapport.get():
            self.var_mois_rapport.set(mois[-1])

    def _generer_rapport(self) -> None:
        mois = self.var_mois_rapport.get().strip()
        if not mois:
            messagebox.showerror("Mois manquant", "Sélectionnez un mois à analyser.")
            return

        syntheses = consolider(mois, self.reference, self.history)
        alertes = generer_alertes(mois, syntheses, self.history, self.reference, self.config_repo)
        self._dernieres_syntheses, self._dernieres_alertes = syntheses, alertes

        self.arbre_synthese.delete(*self.arbre_synthese.get_children())
        for s in sorted(syntheses, key=lambda s: s.total_actuel, reverse=True):
            self.arbre_synthese.insert("", "end", tags=(s.statut,), values=(
                s.matricule, s.nom,
                f"{s.montant_orange:,.0f}".replace(",", " "), f"{s.montant_mtn:,.0f}".replace(",", " "),
                f"{s.total_actuel:,.0f}".replace(",", " "),
                f"{s.total_precedent:,.0f}".replace(",", " ") if s.total_precedent is not None else "—",
                f"{s.variation_montant:,.0f}".replace(",", " ") if s.variation_montant is not None else "—",
                f"{s.variation_pct:.1f}" if s.variation_pct is not None else "—",
                s.statut,
            ))

        self.arbre_alertes.delete(*self.arbre_alertes.get_children())
        for a in alertes:
            self.arbre_alertes.insert("", "end", tags=(a.type.value,),
                                       values=(a.type.value, a.matricule or "", a.nom, a.numero, a.message))

        try:
            chemin = exporter_rapport(mois, syntheses, alertes, self.history)
        except Exception as exc:
            messagebox.showerror("Échec de l'export", str(exc))
            return

        self.label_resultat_export.configure(text=f"Rapport généré : {chemin}")
        self.dashboard.rafraichir(mois)
        if messagebox.askyesno("Rapport généré", f"Rapport généré :\n{chemin}\n\nOuvrir le dossier ?"):
            _ouvrir_dossier(chemin.parent)

    # ------------------------------------------------------------------ #
    # Onglet Paramètres
    # ------------------------------------------------------------------ #

    def _construire_onglet_parametres(self) -> None:
        cadre = self.onglet_parametres

        ttk.Label(cadre, text="Seuil de forte augmentation (%)").grid(row=0, column=0, sticky="w", padx=10, pady=(20, 5))
        self.var_seuil_hausse = tk.StringVar(value=str(self.config_repo.get("SEUIL_FORTE_HAUSSE_PCT")))
        ttk.Entry(cadre, textvariable=self.var_seuil_hausse, width=12).grid(row=0, column=1, sticky="w", padx=10)

        ttk.Label(cadre, text="Seuil de forte consommation (FCFA)").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.var_seuil_conso = tk.StringVar(value=str(self.config_repo.get("SEUIL_FORTE_CONSOMMATION_FCFA")))
        ttk.Entry(cadre, textvariable=self.var_seuil_conso, width=12).grid(row=1, column=1, sticky="w", padx=10)

        ttk.Button(cadre, text="Enregistrer", command=self._enregistrer_parametres).grid(
            row=2, column=0, sticky="w", padx=10, pady=15)

    def _enregistrer_parametres(self) -> None:
        try:
            seuil_hausse = float(self.var_seuil_hausse.get().replace(",", "."))
            seuil_conso = float(self.var_seuil_conso.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Valeur invalide", "Les seuils doivent être des nombres.")
            return
        self.config_repo.set("SEUIL_FORTE_HAUSSE_PCT", seuil_hausse)
        self.config_repo.set("SEUIL_FORTE_CONSOMMATION_FCFA", seuil_conso)
        self.config_repo.save()
        messagebox.showinfo("Paramètres", "Paramètres enregistrés.")


def lancer() -> None:
    app = AppWindow()
    app.mainloop()
