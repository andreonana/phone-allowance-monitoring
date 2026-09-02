from app.history_repository import HistoryRepository
from app.models import LigneConsommation, Operateur, StatutRapprochement


def _ligne(mois, matricule, montant, numero="699123456"):
    return LigneConsommation(mois=mois, operateur=Operateur.ORANGE, numero_brut=numero,
                              numero_normalise=numero, montant=montant, matricule=matricule,
                              statut=StatutRapprochement.OK)


def test_enregistrer_lot_et_persistance(tmp_path):
    chemin = tmp_path / "history.xlsx"
    history = HistoryRepository(chemin)
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx",
                             [_ligne("2026-08", "CM0814", 15000)])
    history.save()

    history2 = HistoryRepository(chemin)
    assert history2.mois_disponibles() == ["2026-08"]
    totaux = history2.totaux_par_matricule("2026-08")
    assert totaux["CM0814"]["total"] == 15000


def test_reimport_remplace_sans_dupliquer(tmp_path):
    history = HistoryRepository(tmp_path / "history.xlsx")
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "v1.xlsx", [_ligne("2026-08", "CM0814", 15000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "v2.xlsx", [_ligne("2026-08", "CM0814", 20000)])

    lignes = history.lignes_du_mois("2026-08")
    assert len(lignes) == 1
    assert lignes[0]["montant"] == 20000
    assert len(history.imports) == 1
    assert history.imports[0].nom_fichier == "v2.xlsx"


def test_mois_precedent_non_calendaire(tmp_path):
    history = HistoryRepository(tmp_path / "history.xlsx")
    history.enregistrer_lot("2026-05", Operateur.ORANGE, "mai.xlsx", [_ligne("2026-05", "CM0814", 1000)])
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx", [_ligne("2026-08", "CM0814", 2000)])
    # Juin et juillet n'existent pas en base -> le mois précédent d'août est mai.
    assert history.mois_precedent("2026-08") == "2026-05"
    assert history.mois_precedent("2026-05") is None


def test_numero_non_identifie_exclu_des_totaux(tmp_path):
    history = HistoryRepository(tmp_path / "history.xlsx")
    ligne = LigneConsommation(mois="2026-08", operateur=Operateur.ORANGE, numero_brut="677000000",
                               numero_normalise="677000000", montant=5000, matricule=None,
                               statut=StatutRapprochement.NON_IDENTIFIE)
    history.enregistrer_lot("2026-08", Operateur.ORANGE, "aout.xlsx", [ligne])
    assert history.totaux_par_matricule("2026-08") == {}
