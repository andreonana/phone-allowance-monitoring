from app.models import Operateur, TypeCollaborateur
from app.reference_repository import ReferenceRepository


def test_upsert_et_persistance(tmp_path):
    chemin = tmp_path / "reference_base.xlsx"
    ref = ReferenceRepository(chemin)
    ref.upsert_collaborateur("CM0814", nom="Jean Dupont", direction="IT")
    ref.upsert_numero("699 123 456", Operateur.ORANGE, "CM0814")
    ref.save()

    ref2 = ReferenceRepository(chemin)
    assert ref2.get_matricule_pour_numero("699123456") == "CM0814"
    collab = ref2.get_collaborateur("CM0814")
    assert collab.nom == "Jean Dupont"
    assert collab.direction == "IT"
    assert collab.type == TypeCollaborateur.INDIVIDUEL


def test_type_pool_devine_pour_matricule_non_standard(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    collab = ref.upsert_collaborateur("Agence Douala")
    assert collab.type == TypeCollaborateur.POOL


def test_numero_invalide_non_ajoute(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    ref.upsert_collaborateur("CM0814")
    resultat = ref.upsert_numero("abc", Operateur.ORANGE, "CM0814")
    assert resultat is None
    assert ref.get_matricule_pour_numero("abc") is None


def test_suppression_collaborateur_supprime_ses_numeros(tmp_path):
    ref = ReferenceRepository(tmp_path / "reference_base.xlsx")
    ref.upsert_collaborateur("CM0814")
    ref.upsert_numero("699123456", Operateur.ORANGE, "CM0814")
    ref.supprimer_collaborateur("CM0814")
    assert ref.get_collaborateur("CM0814") is None
    assert ref.get_matricule_pour_numero("699123456") is None


def test_fichier_absent_donne_base_vide(tmp_path):
    ref = ReferenceRepository(tmp_path / "inexistant.xlsx")
    assert ref.collaborateurs == {}
    assert ref.numeros == {}
