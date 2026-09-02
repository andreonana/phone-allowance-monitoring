"""
Structures de données métier (dataclasses), indépendantes du support de
persistance (Excel) et de l'interface graphique.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Optional


class Operateur(str, Enum):
    ORANGE = "ORANGE"
    MTN = "MTN"


class TypeCollaborateur(str, Enum):
    INDIVIDUEL = "INDIVIDUEL"
    POOL = "POOL"           # compte de service / ligne partagée (agence, projet...)


class StatutRapprochement(str, Enum):
    OK = "OK"
    NON_IDENTIFIE = "NON_IDENTIFIE"     # numéro absent de la base de référence
    NUMERO_INVALIDE = "NUMERO_INVALIDE"  # numéro brut non normalisable


class TypeAlerte(str, Enum):
    NUMERO_NON_IDENTIFIE = "NUMERO_NON_IDENTIFIE"
    NUMERO_INVALIDE = "NUMERO_INVALIDE"
    FORTE_AUGMENTATION = "FORTE_AUGMENTATION"
    FORTE_CONSOMMATION = "FORTE_CONSOMMATION"
    NOUVEAU_NUMERO = "NOUVEAU_NUMERO"
    NUMERO_DISPARU = "NUMERO_DISPARU"
    DOUBLON = "DOUBLON"


@dataclass
class Collaborateur:
    """Une entrée de la base de référence interne (collaborateur ou compte pool)."""

    matricule: str
    nom: str = ""
    direction: str = ""
    fonction: str = ""
    type: TypeCollaborateur = TypeCollaborateur.INDIVIDUEL
    actif: bool = True
    date_creation: datetime = field(default_factory=datetime.now)
    date_modification: datetime = field(default_factory=datetime.now)


@dataclass
class NumeroReference:
    """Association numéro normalisé <-> matricule dans la base de référence."""

    numero_normalise: str
    operateur: Operateur
    matricule: str
    actif: bool = True
    date_creation: datetime = field(default_factory=datetime.now)


@dataclass
class LigneConsommation:
    """
    Une ligne de consommation d'un fichier opérateur importé, après
    normalisation du numéro et rapprochement avec la base de référence.
    Rien n'est jamais supprimé : même les numéros non identifiés ou invalides
    sont conservés ici (cf. cahier des charges §7).
    """

    mois: str                      # format "YYYY-MM"
    operateur: Operateur
    numero_brut: str
    numero_normalise: Optional[str]
    montant: float
    matricule: Optional[str]                    # None si non rapproché
    statut: StatutRapprochement
    staff_id_source_brut: str = ""              # STAFF ID tel que lu dans le fichier opérateur, si présent
    commentaire: str = ""


@dataclass
class SyntheseCollaborateur:
    """Ligne de synthèse consolidée (Orange + MTN) pour un collaborateur, un mois donné."""

    matricule: str
    nom: str
    direction: str
    fonction: str
    type: TypeCollaborateur
    numero_orange: str = ""
    montant_orange: float = 0.0
    numero_mtn: str = ""
    montant_mtn: float = 0.0
    total_actuel: float = 0.0
    total_precedent: Optional[float] = None
    variation_montant: Optional[float] = None
    variation_pct: Optional[float] = None
    statut: str = ""          # libellé lisible : Hausse / Baisse / Stable / Nouveau...
    alertes: list = field(default_factory=list)


@dataclass
class Alerte:
    mois: str
    type: TypeAlerte
    matricule: Optional[str]
    nom: str
    numero: str
    message: str
    valeur: Optional[float] = None
    seuil_applique: Optional[float] = None


@dataclass
class LotImport:
    """Trace d'un import mensuel (un par opérateur et par mois)."""

    mois: str
    operateur: Operateur
    nom_fichier: str
    date_import: datetime
    nb_lignes: int
    nb_non_identifies: int
    nb_invalides: int
