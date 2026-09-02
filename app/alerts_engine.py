"""
Génération des alertes configurables du mois (cahier des charges §10 pour
les seuils). Deux familles de sources :

    - les lignes brutes de l'historique (`HistoryRepository.lignes_du_mois`) :
      numéros invalides, non identifiés, doublons ;
    - la synthèse consolidée (`app.consolidation.consolider`) : forte
      augmentation, forte consommation, nouveau numéro, numéro disparu.

Les alertes générées sont aussi rattachées à `SyntheseCollaborateur.alertes`
pour affichage direct dans le rapport / l'IHM, sans devoir les recroiser.
"""

from __future__ import annotations

from collections import Counter

from app.config import ConfigRepository
from app.history_repository import HistoryRepository
from app.models import Alerte, StatutRapprochement, SyntheseCollaborateur, TypeAlerte
from app.reference_repository import ReferenceRepository
from app import consolidation


def _alertes_lignes_brutes(mois: str, history: HistoryRepository,
                            reference: ReferenceRepository) -> list[Alerte]:
    alertes: list[Alerte] = []
    lignes_mois = history.lignes_du_mois(mois)

    for c in lignes_mois:
        if c["statut"] == StatutRapprochement.NUMERO_INVALIDE.value:
            alertes.append(Alerte(
                mois=mois, type=TypeAlerte.NUMERO_INVALIDE, matricule=None, nom="",
                numero=c["numero_brut"] or "",
                message=(f"Numéro invalide dans le fichier {c['operateur']} "
                         f"(valeur brute : '{c['numero_brut']}')"),
            ))
        elif c["statut"] == StatutRapprochement.NON_IDENTIFIE.value:
            numero = c["numero_normalise"] or c["numero_brut"] or ""
            alertes.append(Alerte(
                mois=mois, type=TypeAlerte.NUMERO_NON_IDENTIFIE, matricule=None, nom="",
                numero=numero,
                message=(f"Numéro {numero} ({c['operateur']}) absent de la base de "
                         f"référence, consommation non rattachée"),
                valeur=c["montant"],
            ))

    # Doublons : un même numéro rapproché apparaît plusieurs fois pour le
    # même opérateur dans le mois (fichier source mal dédupliqué en amont).
    compteur = Counter(
        (c["operateur"], c["numero_normalise"]) for c in lignes_mois
        if c["numero_normalise"] and c["statut"] == StatutRapprochement.OK.value
    )
    for (operateur, numero), nb in compteur.items():
        if nb <= 1:
            continue
        matricule = reference.get_matricule_pour_numero(numero)
        collab = reference.get_collaborateur(matricule) if matricule else None
        alertes.append(Alerte(
            mois=mois, type=TypeAlerte.DOUBLON, matricule=matricule,
            nom=collab.nom if collab else "", numero=numero,
            message=f"Numéro {numero} ({operateur}) présent {nb} fois dans le fichier importé",
            valeur=float(nb),
        ))

    return alertes


def _alertes_synthese(mois: str, syntheses: list[SyntheseCollaborateur],
                       config: ConfigRepository) -> list[Alerte]:
    alertes: list[Alerte] = []
    seuil_hausse = config.get("SEUIL_FORTE_HAUSSE_PCT")
    seuil_conso = config.get("SEUIL_FORTE_CONSOMMATION_FCFA")

    for s in syntheses:
        numero = s.numero_orange or s.numero_mtn

        if s.statut == consolidation.STATUT_NOUVEAU:
            a = Alerte(mois=mois, type=TypeAlerte.NOUVEAU_NUMERO, matricule=s.matricule,
                       nom=s.nom, numero=numero,
                       message=f"Nouveau numéro actif pour {s.nom or s.matricule}",
                       valeur=s.total_actuel)
            alertes.append(a)
            s.alertes.append(a)

        if s.statut == consolidation.STATUT_DISPARU:
            a = Alerte(mois=mois, type=TypeAlerte.NUMERO_DISPARU, matricule=s.matricule,
                       nom=s.nom, numero=numero,
                       message=(f"Plus aucune consommation ce mois-ci pour {s.nom or s.matricule} "
                                f"(consommait {s.total_precedent:,.0f} FCFA le mois précédent)".replace(",", " ")),
                       valeur=s.total_precedent)
            alertes.append(a)
            s.alertes.append(a)

        if s.variation_pct is not None and s.variation_pct >= seuil_hausse:
            a = Alerte(mois=mois, type=TypeAlerte.FORTE_AUGMENTATION, matricule=s.matricule,
                       nom=s.nom, numero=numero,
                       message=(f"Forte augmentation pour {s.nom or s.matricule} : "
                                f"+{s.variation_pct:.0f}% (seuil {seuil_hausse:.0f}%)"),
                       valeur=s.variation_pct, seuil_applique=seuil_hausse)
            alertes.append(a)
            s.alertes.append(a)

        if s.total_actuel >= seuil_conso:
            a = Alerte(mois=mois, type=TypeAlerte.FORTE_CONSOMMATION, matricule=s.matricule,
                       nom=s.nom, numero=numero,
                       message=(f"Forte consommation pour {s.nom or s.matricule} : "
                                f"{s.total_actuel:,.0f} FCFA (seuil {seuil_conso:,.0f} FCFA)".replace(",", " ")),
                       valeur=s.total_actuel, seuil_applique=seuil_conso)
            alertes.append(a)
            s.alertes.append(a)

    return alertes


def generer_alertes(mois: str, syntheses: list[SyntheseCollaborateur],
                     history: HistoryRepository, reference: ReferenceRepository,
                     config: ConfigRepository) -> list[Alerte]:
    """
    Génère toutes les alertes du mois. Rattache également chaque alerte
    concernant un collaborateur identifié à `SyntheseCollaborateur.alertes`
    (mutation en place des objets de `syntheses`).
    """
    alertes = _alertes_lignes_brutes(mois, history, reference)
    alertes += _alertes_synthese(mois, syntheses, config)
    return alertes
