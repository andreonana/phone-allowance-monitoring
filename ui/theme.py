"""
Thème visuel et code couleur de l'IHM (Tkinter). Aucune logique métier ici :
uniquement des constantes de style et leur application à un `ttk.Style`.

Le code couleur est partagé entre les listes de l'IHM (Treeview) et le
tableau de bord (matplotlib), pour rester cohérent d'un écran à l'autre :

    rouge  -> ça coûte plus cher / à surveiller (hausse, alerte forte)
    vert   -> ça baisse (bonne nouvelle)
    gris   -> stable, rien à signaler
    bleu   -> nouveauté (nouveau numéro/collaborateur)
    ambre  -> donnée à corriger dans la base de référence (non identifié...)
    violet -> numéro disparu du mois
"""

from __future__ import annotations

from tkinter import ttk

# ---------------------------------------------------------------------- #
# Palette générale
# ---------------------------------------------------------------------- #

FOND = "#F3F4F6"           # fond général des fenêtres/onglets
FOND_CARTE = "#FFFFFF"     # fond des cadres "carte"
BORDURE = "#D1D5DB"
TEXTE = "#111827"
TEXTE_ATTENUE = "#6B7280"

PRIMAIRE = "#1F4E78"       # bleu nuit (même teinte que l'en-tête du rapport Excel)
PRIMAIRE_SURVOL = "#163A5A"
ACCENT = "#2563EB"

ROUGE_FOND, ROUGE_TEXTE = "#FEE2E2", "#991B1B"
ROUGE_VIF_FOND, ROUGE_VIF_TEXTE = "#FFEDD5", "#9A3412"
VERT_FOND, VERT_TEXTE = "#DCFCE7", "#166534"
GRIS_FOND, GRIS_TEXTE = "#F3F4F6", "#374151"
BLEU_FOND, BLEU_TEXTE = "#DBEAFE", "#1E40AF"
AMBRE_FOND, AMBRE_TEXTE = "#FEF3C7", "#92400E"
VIOLET_FOND, VIOLET_TEXTE = "#EDE9FE", "#5B21B6"
ROUGE_FONCE_FOND, ROUGE_FONCE_TEXTE = "#FEE2E2", "#7F1D1D"

# Couleurs "pleines" utilisées pour les courbes/barres du tableau de bord.
COURBE_ORANGE = "#FF7900"   # couleur opérateur Orange
COURBE_MTN = "#FFCC00"      # couleur opérateur MTN
COURBE_TOTAL = "#1F4E78"
HAUSSE = "#DC2626"
BAISSE = "#16A34A"

# ---------------------------------------------------------------------- #
# Code couleur par statut de synthèse / type d'alerte, réutilisé par les
# Treeview (tag_configure) ET par le tableau de bord.
# ---------------------------------------------------------------------- #

COULEURS_STATUT: dict[str, tuple[str, str]] = {
    "Hausse": (ROUGE_FOND, ROUGE_TEXTE),
    "Baisse": (VERT_FOND, VERT_TEXTE),
    "Stable": (GRIS_FOND, GRIS_TEXTE),
    "Nouveau": (BLEU_FOND, BLEU_TEXTE),
    "Disparu": (AMBRE_FOND, AMBRE_TEXTE),
    "": (FOND_CARTE, TEXTE_ATTENUE),
}

COULEURS_ALERTE: dict[str, tuple[str, str]] = {
    "FORTE_AUGMENTATION": (ROUGE_FOND, ROUGE_TEXTE),
    "FORTE_CONSOMMATION": (ROUGE_VIF_FOND, ROUGE_VIF_TEXTE),
    "NUMERO_NON_IDENTIFIE": (AMBRE_FOND, AMBRE_TEXTE),
    "NUMERO_INVALIDE": (GRIS_FOND, GRIS_TEXTE),
    "NOUVEAU_NUMERO": (BLEU_FOND, BLEU_TEXTE),
    "NUMERO_DISPARU": (VIOLET_FOND, VIOLET_TEXTE),
    "DOUBLON": (ROUGE_FONCE_FOND, ROUGE_FONCE_TEXTE),
}


def configurer_tags_couleur(tree: ttk.Treeview, couleurs: dict[str, tuple[str, str]]) -> None:
    """Déclare un tag Treeview par clé de `couleurs` (statut ou type
    d'alerte) : la ligne insérée avec ce tag prend le fond/texte associé."""
    for cle, (fond, texte) in couleurs.items():
        tree.tag_configure(cle, background=fond, foreground=texte)


def appliquer_theme(racine) -> ttk.Style:
    """Configure un thème clair, sobre et cohérent pour toute l'application.
    `racine` est la fenêtre Tk principale (pour ajuster la taille de police
    par défaut, qui se propage à tous les widgets standards)."""
    import tkinter.font as tkfont

    racine.configure(background=FOND)

    police_defaut = tkfont.nametofont("TkDefaultFont")
    police_defaut.configure(size=10)
    tkfont.nametofont("TkTextFont").configure(size=10)
    police_titre = tkfont.Font(family=police_defaut.actual("family"), size=15, weight="bold")
    police_soustitre = tkfont.Font(family=police_defaut.actual("family"), size=11, weight="bold")
    racine.option_add("*Font", police_defaut)

    style = ttk.Style(racine)
    try:
        style.theme_use("clam")
    except Exception:
        pass  # thème non disponible sur cette plateforme -> on garde le défaut

    style.configure(".", background=FOND, foreground=TEXTE)
    style.configure("TFrame", background=FOND)
    style.configure("Carte.TFrame", background=FOND_CARTE, relief="flat")
    style.configure("TLabelframe", background=FOND, bordercolor=BORDURE)
    style.configure("TLabelframe.Label", background=FOND, foreground=PRIMAIRE, font=police_soustitre)
    style.configure("TLabel", background=FOND, foreground=TEXTE)
    style.configure("Carte.TLabel", background=FOND_CARTE, foreground=TEXTE)
    style.configure("Titre.TLabel", background=FOND, foreground=PRIMAIRE, font=police_titre)
    style.configure("SousTitre.TLabel", background=FOND, foreground=TEXTE_ATTENUE)

    style.configure("TButton", padding=(12, 6), background=PRIMAIRE, foreground="#FFFFFF",
                     borderwidth=0, focusthickness=0)
    style.map("TButton", background=[("active", PRIMAIRE_SURVOL), ("disabled", BORDURE)],
              foreground=[("disabled", TEXTE_ATTENUE)])

    style.configure("TEntry", fieldbackground="#FFFFFF", padding=4)
    style.configure("TCombobox", fieldbackground="#FFFFFF", padding=4)

    style.configure("TNotebook", background=FOND, borderwidth=0, tabmargins=(6, 6, 6, 0))
    style.configure("TNotebook.Tab", padding=(16, 8), background="#E5E7EB", foreground=TEXTE_ATTENUE)
    style.map("TNotebook.Tab",
              background=[("selected", FOND_CARTE)],
              foreground=[("selected", PRIMAIRE)])

    style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                     foreground=TEXTE, rowheight=26, borderwidth=0)
    style.configure("Treeview.Heading", background=PRIMAIRE, foreground="#FFFFFF",
                     font=police_soustitre, relief="flat", padding=(6, 6))
    style.map("Treeview.Heading", background=[("active", PRIMAIRE_SURVOL)])
    style.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "#FFFFFF")])

    return style
