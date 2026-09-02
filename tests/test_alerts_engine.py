import pytest

from app.alerts_engine import generer_alertes
from app.config import ConfigRepository
from app.consolidation import consolider
from app.history_repository import HistoryRepository
from app.models import LigneConsommation, Operateur, StatutRapprochement, TypeAlerte
from app.reference_repository import ReferenceRepository


@pytest.fixture
def reference(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    ref.upsert_collaborateur("CM0814", nom="Jean Dupont")
    return ref


@pytest.fixture
def history(tmp_path):
    return HistoryRepository(tmp_path / "history.xlsx")


@pytest.fixture
def config(tmp_path):
    return ConfigRepository(tmp_path / "config.xlsx")


def _ligne(mois, operateur, matricule, montant, numero="699123456", statut=StatutRapprochement.OK):
    return LigneConsommation(mois=mois, operateur=operateur, numero_brut=numero,
                              numero_normalise=numero, montant=montant, matricule=matricule,
                              statut=statut)


def test_alerte_forte_augmentation(reference, history, config):
    config.set("SEUIL_FORTE_HAUSSE_PCT", 50.0)
    history.enregistrer_lot("2026-07", Operateur.ORANGE, "j.xlsx",
                             [_ligne("2026-07", Operateur.ORANGE, "CM0814", 10000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 20000)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)

    types = [a.type for a in alertes]
    assert TypeAlerte.FORTE_AUGMENTATION in types
    a = next(a for a in alertes if a.type == TypeAlerte.FORTE_AUGMENTATION)
    assert a.matricule == "CM0814"
    assert a.valeur == 100.0

    s = next(s for s in syntheses if s.matricule == "CM0814")
    assert any(al.type == TypeAlerte.FORTE_AUGMENTATION for al in s.alertes)


def test_pas_alerte_hausse_sous_le_seuil(reference, history, config):
    config.set("SEUIL_FORTE_HAUSSE_PCT", 50.0)
    history.enregistrer_lot("2026-07", Operateur.ORANGE, "j.xlsx",
                             [_ligne("2026-07", Operateur.ORANGE, "CM0814", 10000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 11000)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert not any(a.type == TypeAlerte.FORTE_AUGMENTATION for a in alertes)


def test_alerte_forte_consommation(reference, history, config):
    config.set("SEUIL_FORTE_CONSOMMATION_FCFA", 100000.0)
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 150000)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.FORTE_CONSOMMATION and a.matricule == "CM0814" for a in alertes)


def test_alerte_numero_non_identifie(reference, history, config):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, None, 5000,
                                     statut=StatutRapprochement.NON_IDENTIFIE)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.NUMERO_NON_IDENTIFIE for a in alertes)


def test_alerte_numero_invalide(reference, history, config):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, None, 5000, numero=None,
                                     statut=StatutRapprochement.NUMERO_INVALIDE)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.NUMERO_INVALIDE for a in alertes)


def test_alerte_doublon(reference, history, config):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx", [
        _ligne("2026-08", Operateur.ORANGE, "CM0814", 5000),
        _ligne("2026-08", Operateur.ORANGE, "CM0814", 3000),
    ])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.DOUBLON and a.numero == "699123456" for a in alertes)


def test_alerte_numero_disparu(reference, history, config):
    history.enregistrer_lot("2026-07", Operateur.ORANGE, "j.xlsx",
                             [_ligne("2026-07", Operateur.ORANGE, "CM0814", 10000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx", [])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.NUMERO_DISPARU and a.matricule == "CM0814" for a in alertes)


def test_alerte_nouveau_numero(reference, history, config):
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "a.xlsx",
                             [_ligne("2026-08", Operateur.ORANGE, "CM0814", 8000)])
    syntheses = consolider("2026-08", reference, history)
    alertes = generer_alertes("2026-08", syntheses, history, reference, config)
    assert any(a.type == TypeAlerte.NOUVEAU_NUMERO and a.matricule == "CM0814" for a in alertes)
