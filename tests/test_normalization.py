from app.normalization import normalize_phone_number


def test_numero_local_simple():
    r = normalize_phone_number("699123456")
    assert r.is_valid
    assert r.normalized == "699123456"


def test_numero_avec_espaces_et_tirets():
    assert normalize_phone_number("699 123 456").normalized == "699123456"
    assert normalize_phone_number("699-123-456").normalized == "699123456"


def test_numero_avec_indicatif_pays():
    assert normalize_phone_number("+237699123456").normalized == "699123456"
    assert normalize_phone_number("237699123456").normalized == "699123456"
    assert normalize_phone_number("+237 699 123 456").normalized == "699123456"
    assert normalize_phone_number("00237699123456").normalized == "699123456"


def test_numero_artefact_excel_float():
    assert normalize_phone_number(699123456.0).normalized == "699123456"


def test_numero_notation_scientifique_texte():
    r = normalize_phone_number("6.99123456E+8")
    assert r.is_valid
    assert r.normalized == "699123456"


def test_numero_zero_de_tete_errone():
    assert normalize_phone_number("0699123456").normalized == "699123456"


def test_numero_vide_invalide():
    r = normalize_phone_number(None)
    assert not r.is_valid
    assert r.normalized is None
    assert "vide" in r.reason.lower()

    r2 = normalize_phone_number("")
    assert not r2.is_valid


def test_numero_ne_commencant_pas_par_6_invalide():
    r = normalize_phone_number("512345678")
    assert not r.is_valid


def test_numero_lettres_invalide():
    r = normalize_phone_number("ABC123456")
    assert not r.is_valid
    assert "ABC123456" in r.reason
