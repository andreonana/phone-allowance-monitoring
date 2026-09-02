import pandas as pd
import pytest

from app.models import Operateur, StatutRapprochement
from app.reconciliation_engine import normaliser_montant, rapprocher_fichier
from app.reference_repository import ReferenceRepository


@pytest.fixture
def reference(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    ref.upsert_collaborateur("CM0814", nom="Jean Dupont")
    ref.upsert_numero("699123456", Operateur.ORANGE, "CM0814")
    return ref


def _df(records):
    return pd.DataFrame.from_records(
        records, columns=["numero_brut", "montant_brut", "staff_id_brut", "commentaire_brut"]
    )


def test_ligne_ok_rattachee_au_matricule(reference):
    df = _df([{"numero_brut": "699123456", "montant_brut": 15000, "staff_id_brut": None, "commentaire_brut": None}])
    lignes = rapprocher_fichier(df, "2026-08", Operateur.ORANGE, reference)
    assert len(lignes) == 1
    l = lignes[0]
    assert l.statut == StatutRapprochement.OK
    assert l.matricule == "CM0814"
    assert l.numero_normalise == "699123456"
    assert l.montant == 15000.0


def test_ligne_non_identifiee(reference):
    df = _df([{"numero_brut": "677000000", "montant_brut": 5000, "staff_id_brut": None, "commentaire_brut": None}])
    lignes = rapprocher_fichier(df, "2026-08", Operateur.ORANGE, reference)
    assert lignes[0].statut == StatutRapprochement.NON_IDENTIFIE
    assert lignes[0].matricule is None
    assert lignes[0].numero_normalise == "677000000"


def test_ligne_numero_invalide_conservee(reference):
    df = _df([{"numero_brut": "abc", "montant_brut": 5000, "staff_id_brut": None, "commentaire_brut": None}])
    lignes = rapprocher_fichier(df, "2026-08", Operateur.ORANGE, reference)
    assert len(lignes) == 1  # jamais supprimée (§7)
    assert lignes[0].statut == StatutRapprochement.NUMERO_INVALIDE
    assert lignes[0].numero_normalise is None


@pytest.mark.parametrize("brut,attendu", [
    (15000, 15000.0),
    (15000.0, 15000.0),
    ("15 000", 15000.0),
    ("15000 FCFA", 15000.0),
    ("15,5", 15.5),
    ("(500)", -500.0),
    (None, 0.0),
    ("", 0.0),
])
def test_normaliser_montant(brut, attendu):
    assert normaliser_montant(brut) == attendu
