"""
Normalisation et validation des numéros de téléphone.

Point CRITIQUE du système (cf. cahier des charges §5) : tout rapprochement
entre les fichiers opérateurs et la base de référence se fait exclusivement
sur le numéro NORMALISÉ, jamais sur le numéro brut.

Format interne unique retenu : 9 chiffres, sans indicatif pays, sans espace,
commençant par 6 (numérotation mobile camerounaise). Exemple : "699123456".

Formats bruts pris en charge (observés dans les fichiers réels + formats
annoncés par le cahier des charges) :
    699123456
    699 123 456
    699-123-456
    +237699123456
    237699123456
    +237 699 123 456
    699123456.0        (artefact Excel : numéro lu comme float)
    6.99123456E+8       (artefact Excel : notation scientifique)
    0699123456          (zéro de tête erroné, tolérance de saisie)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

# Un numéro mobile camerounais valide, une fois normalisé, a 9 chiffres et
# commence par 6 (toutes les plages Orange/MTN/Camtel observées commencent
# par 65x-69x). On ne verrouille pas plus finement le 2e chiffre : le
# cahier des charges des opérateurs évolue et un verrou trop strict casserait
# la robustesse demandée en §16.
_VALID_LOCAL_PATTERN = re.compile(r"^6\d{8}$")

# Préfixe pays Cameroun
_COUNTRY_CODE = "237"


@dataclass(frozen=True)
class NormalizationResult:
    """Résultat de la normalisation d'un numéro brut."""

    raw_value: str          # valeur brute telle que lue dans le fichier (str)
    normalized: Optional[str]  # numéro normalisé sur 9 chiffres, ou None si invalide
    is_valid: bool
    reason: str = ""         # explication si invalide (pour affichage utilisateur / logs)

    def __str__(self) -> str:  # pragma: no cover - confort de debug
        return self.normalized if self.is_valid else f"INVALIDE({self.raw_value})"


def _to_digit_string(raw) -> str:
    """
    Convertit une valeur brute (int, float, str, None...) telle que renvoyée
    par openpyxl/pandas en une chaîne de chiffres, en gérant les artefacts
    Excel classiques (float avec .0, notation scientifique, espaces, tirets,
    parenthèses, points, apostrophes de séparation de milliers).
    """
    if raw is None:
        return ""

    if isinstance(raw, float):
        # Excel stocke parfois les numéros de téléphone comme des nombres.
        # 640117015.0 -> "640117015"
        if raw.is_integer():
            return str(int(raw))
        # Notation scientifique résiduelle éventuelle -> on tente quand même
        return re.sub(r"\D", "", f"{raw:.0f}")

    if isinstance(raw, int):
        return str(raw)

    text = str(raw).strip()
    if not text:
        return ""

    # Notation scientifique textuelle type "6.40117E+08"
    sci_match = re.match(r"^[\d.]+E\+?\d+$", text, flags=re.IGNORECASE)
    if sci_match:
        try:
            return str(int(float(text)))
        except ValueError:
            pass

    # Cas "699123456.0" fourni comme texte
    if re.match(r"^\d+\.0+$", text):
        return text.split(".")[0]

    # On garde le '+' de tête (indicatif) le temps du traitement, on
    # supprime tout le reste qui n'est pas un chiffre (espaces, tirets,
    # points, parenthèses, apostrophes).
    has_plus = text.startswith("+")
    digits = re.sub(r"\D", "", text)
    return ("+" + digits) if has_plus else digits


def normalize_phone_number(raw) -> NormalizationResult:
    """
    Normalise un numéro de téléphone brut vers le format interne unique
    (9 chiffres, sans indicatif pays).

    Ne lève jamais d'exception : toute entrée invalide renvoie un
    NormalizationResult avec is_valid=False et une raison explicite, afin
    que l'appelant puisse conserver la ligne (cf. §7 : "ne pas supprimer
    la ligne") et l'afficher clairement à l'utilisateur.
    """
    raw_str = "" if raw is None else str(raw).strip()
    digits_with_plus = _to_digit_string(raw)

    if not digits_with_plus:
        reason = "Numéro vide" if not raw_str else f"Aucun chiffre trouvé dans '{raw_str}'"
        return NormalizationResult(raw_value=raw_str, normalized=None, is_valid=False,
                                    reason=reason)

    digits = digits_with_plus.lstrip("+")

    candidate: Optional[str] = None

    if digits.startswith(_COUNTRY_CODE) and len(digits) == len(_COUNTRY_CODE) + 9:
        # +237699123456 / 237699123456
        candidate = digits[len(_COUNTRY_CODE):]
    elif len(digits) == 9:
        # 699123456
        candidate = digits
    elif len(digits) == 10 and digits.startswith("0"):
        # Tolérance : saisie avec un 0 de tête erroné (ex: 0699123456)
        candidate = digits[1:]
    elif len(digits) == 2 + len(_COUNTRY_CODE) + 9 and digits.startswith("00" + _COUNTRY_CODE):
        # 00237699123456 (format international avec préfixe de sortie 00)
        candidate = digits[2 + len(_COUNTRY_CODE):]
    else:
        candidate = digits  # on tente la validation telle quelle, message d'erreur clair sinon

    if candidate and _VALID_LOCAL_PATTERN.match(candidate):
        return NormalizationResult(raw_value=raw_str, normalized=candidate, is_valid=True)

    return NormalizationResult(
        raw_value=raw_str,
        normalized=None,
        is_valid=False,
        reason=(
            f"Format de numéro non reconnu ou invalide (attendu : 9 chiffres "
            f"commençant par 6, avec ou sans indicatif +237) — valeur lue : '{raw_str}'"
        ),
    )
