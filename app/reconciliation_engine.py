"""
Moteur de rapprochement : transforme les lignes brutes d'un fichier opérateur
(déjà extraites par `app.excel_io`) en lignes de consommation exploitables
(`app.models.LigneConsommation`), en normalisant le numéro (§5) et en le
rapprochant de la base de référence (`app.reference_repository`).

Aucune ligne n'est jamais écartée (§7) : un numéro invalide ou absent de la
base de référence donne une ligne avec un statut explicite
(`NUMERO_INVALIDE` / `NON_IDENTIFIE`) plutôt que d'être supprimée.
"""

from __future__ import annotations

import re

import pandas as pd

from app.models import LigneConsommation, Operateur, StatutRapprochement
from app.normalization import normalize_phone_number
from app.reference_repository import ReferenceRepository

# Caractères qu'on tolère dans un montant brut (séparateurs de milliers,
# symbole monétaire, espaces insécables...) en plus des chiffres et du
# séparateur décimal.
_MONTANT_BRUIT = re.compile(r"[^\d,.\-]")


def normaliser_montant(raw) -> float:
    """
    Convertit un montant brut (int, float, str potentiellement bruitée par
    Excel : espaces, symbole FCFA, virgule décimale, parenthèses pour les
    négatifs...) en float. Ne lève jamais d'exception : une valeur non
    interprétable est traitée comme 0.0 (montant absent, pas une consommation
    supprimée : la ligne elle-même reste conservée par l'appelant).
    """
    if raw is None:
        return 0.0
    if isinstance(raw, (int, float)):
        try:
            if isinstance(raw, float) and raw != raw:  # NaN
                return 0.0
        except Exception:
            pass
        return float(raw)

    text = str(raw).strip()
    if not text:
        return 0.0

    negatif = text.startswith("(") and text.endswith(")")
    if negatif:
        text = text[1:-1]

    text = text.replace("FCFA", "").replace("XAF", "").strip()
    text = _MONTANT_BRUIT.sub("", text)
    if not text:
        return 0.0

    # Une seule virgule et pas de point -> virgule décimale (format FR).
    # Sinon la virgule est un séparateur de milliers -> on la supprime.
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")

    try:
        valeur = float(text)
    except ValueError:
        return 0.0

    return -valeur if negatif else valeur


def rapprocher_fichier(df: pd.DataFrame, mois: str, operateur: Operateur,
                        reference: ReferenceRepository) -> list[LigneConsommation]:
    """
    Rapproche chaque ligne du DataFrame extrait par `excel_io.lire_fichier_operateur`
    avec la base de référence. Retourne une ligne de consommation par ligne
    source, quel que soit le résultat du rapprochement.
    """
    lignes: list[LigneConsommation] = []

    for _, row in df.iterrows():
        numero_brut = row.get("numero_brut")
        montant = normaliser_montant(row.get("montant_brut"))
        staff_id_brut = row.get("staff_id_brut")
        commentaire_brut = row.get("commentaire_brut")

        resultat = normalize_phone_number(numero_brut)

        if not resultat.is_valid:
            lignes.append(LigneConsommation(
                mois=mois, operateur=operateur,
                numero_brut="" if numero_brut is None else str(numero_brut),
                numero_normalise=None, montant=montant, matricule=None,
                statut=StatutRapprochement.NUMERO_INVALIDE,
                staff_id_source_brut="" if staff_id_brut is None else str(staff_id_brut),
                commentaire="" if commentaire_brut is None else str(commentaire_brut),
            ))
            continue

        matricule = reference.get_matricule_pour_numero(resultat.normalized)
        statut = StatutRapprochement.OK if matricule else StatutRapprochement.NON_IDENTIFIE

        lignes.append(LigneConsommation(
            mois=mois, operateur=operateur,
            numero_brut="" if numero_brut is None else str(numero_brut),
            numero_normalise=resultat.normalized, montant=montant, matricule=matricule,
            statut=statut,
            staff_id_source_brut="" if staff_id_brut is None else str(staff_id_brut),
            commentaire="" if commentaire_brut is None else str(commentaire_brut),
        ))

    return lignes
