import pytest

from app.consolidation import consolider
from app.history_repository import HistoryRepository
from app.models import LigneConsommation, Operateur, StatutRapprochement
from app.reference_repository import ReferenceRepository


@pytest.fixture
def reference(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    ref.upsert_collaborateur("CM0814", nom="Jean Dupont", direction="IT")
    ref.upsert_collaborateur("CM0900", nom="Awa Ngo", direction="RH")
    return ref


@pytest.fixture
def history(tmp_path):
    return HistoryRepository(tmp_path / "history.xlsx")


def _ligne(mois, operateur, matricule, montant, numero="699123456", statut=StatutRapprochement.OK):
    return LigneConsommation(mois=mois, operateur=operateur, numero_brut=numero,
                              numero_normalise=numero, montant=montant, matricule=matricule,
                              statut=statut)


def test_hausse_detectee(reference, history):
    history.enregistrer_lot("2026-07", Operateur.ORANGE, "juillet.xlsx",
                             [_ligne("2026-07", Operateur.ORANGE, "CM0814", 10000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 20000)])

    syntheses = consolider("2026-08", reference, history)
    s = next(s for s in syntheses if s.matricule == "CM0814")
    assert s.total_actuel == 20000
    assert s.total_precedent == 10000
    assert s.variation_montant == 10000
    assert s.variation_pct == 100.0
    assert s.statut == "Hausse"
    assert s.nom == "Jean Dupont"


def test_nouveau_collaborateur_sans_mois_precedent(reference, history):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0900", 5000)])
    syntheses = consolider("2026-08", reference, history)
    s = next(s for s in syntheses if s.matricule == "CM0900")
    assert s.total_precedent is None
    assert s.statut == "Nouveau"


def test_disparu_quand_plus_de_consommation(reference, history):
    history.enregistrer_lot("2026-07", Operateur.ORANGE, "juillet.xlsx",
                             [_ligne("2026-07", Operateur.ORANGE, "CM0814", 10000)])
    # Rien pour CM0814 en août -> il doit quand même apparaître (comparaison au mois précédent)
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0900", 3000)])

    syntheses = consolider("2026-08", reference, history)
    s = next(s for s in syntheses if s.matricule == "CM0814")
    assert s.total_actuel == 0
    assert s.total_precedent == 10000
    assert s.statut == "Disparu"


def test_consolidation_orange_et_mtn(reference, history):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout_orange.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 10000, numero="699111111")])
    history.enregistrer_lot("2026-08", Operateur.MTN, "aout_mtn.xlsx",
                             [_ligne("2026-08", Operateur.MTN, "CM0814", 7000, numero="677222222")])

    syntheses = consolider("2026-08", reference, history)
    s = next(s for s in syntheses if s.matricule == "CM0814")
    assert s.montant_orange == 10000
    assert s.montant_mtn == 7000
    assert s.total_actuel == 17000
    assert s.numero_orange == "699111111"
    assert s.numero_mtn == "677222222"


def test_lignes_non_identifiees_exclues_de_la_synthese(reference, history):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, None, 5000,
                                     statut=StatutRapprochement.NON_IDENTIFIE)])
    syntheses = consolider("2026-08", reference, history)
    assert syntheses == []
