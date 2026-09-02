"""
Consolidation mensuelle : agrège les consommations Orange + MTN par
collaborateur et compare au mois précédent (celui retourné par
`HistoryRepository.mois_precedent`, pas forcément M-1 calendaire — §16).

Ne considère que les collaborateurs ayant eu une consommation rapprochée
(statut OK) le mois analysé et/ou le mois précédent : un collaborateur de la
base de référence qui n'a jamais eu de ligne de consommation n'apparaît pas
dans la synthèse (rien à comparer).
"""

from __future__ import annotations

from app.history_repository import HistoryRepository
from app.models import SyntheseCollaborateur, TypeCollaborateur
from app.reference_repository import ReferenceRepository

STATUT_NOUVEAU = "Nouveau"
STATUT_DISPARU = "Disparu"
STATUT_HAUSSE = "Hausse"
STATUT_BAISSE = "Baisse"
STATUT_STABLE = "Stable"
STATUT_INCONNU = ""  # pas de mois précédent en base -> aucune comparaison possible


def _determiner_statut(total_actuel: float, total_precedent: float | None) -> str:
    if total_precedent is None:
        return STATUT_NOUVEAU if total_actuel > 0 else STATUT_INCONNU
    if total_actuel == 0 and total_precedent > 0:
        return STATUT_DISPARU
    if total_actuel > total_precedent:
        return STATUT_HAUSSE
    if total_actuel < total_precedent:
        return STATUT_BAISSE
    return STATUT_STABLE


def consolider(mois: str, reference: ReferenceRepository,
                history: HistoryRepository) -> list[SyntheseCollaborateur]:
    """Construit la synthèse consolidée du mois pour chaque collaborateur
    ayant eu une consommation ce mois-ci et/ou le mois précédent."""
    totaux_actuels = history.totaux_par_matricule(mois)
    mois_prec = history.mois_precedent(mois)
    totaux_precedents = history.totaux_par_matricule(mois_prec) if mois_prec else {}

    matricules = set(totaux_actuels) | set(totaux_precedents)
    resultats: list[SyntheseCollaborateur] = []

    for matricule in sorted(matricules):
        collab = reference.get_collaborateur(matricule)
        actuel = totaux_actuels.get(matricule)
        precedent = totaux_precedents.get(matricule)

        total_actuel = actuel["total"] if actuel else 0.0
        total_precedent = precedent["total"] if precedent else None

        if total_precedent:
            variation_montant = total_actuel - total_precedent
            variation_pct = (variation_montant / total_precedent) * 100
        else:
            variation_montant = None
            variation_pct = None

        resultats.append(SyntheseCollaborateur(
            matricule=matricule,
            nom=collab.nom if collab else "",
            direction=collab.direction if collab else "",
            fonction=collab.fonction if collab else "",
            type=collab.type if collab else TypeCollaborateur.INDIVIDUEL,
            numero_orange=actuel["numero_orange"] if actuel else "",
            montant_orange=actuel["ORANGE"] if actuel else 0.0,
            numero_mtn=actuel["numero_mtn"] if actuel else "",
            montant_mtn=actuel["MTN"] if actuel else 0.0,
            total_actuel=total_actuel,
            total_precedent=total_precedent,
            variation_montant=variation_montant,
            variation_pct=variation_pct,
            statut=_determiner_statut(total_actuel, total_precedent),
        ))

    return resultats
