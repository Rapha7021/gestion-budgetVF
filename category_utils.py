from functools import lru_cache
from typing import Dict, Iterable, List, Tuple

from database import get_connection

DEFAULT_CATEGORIES: List[Tuple[str, str]] = [
    ("STP", "Stagiaire Projet"),
    ("AOP", "Assistante / opérateur"),
    ("TEP", "Technicien"),
    ("IJP", "Junior"),
    ("ISP", "Senior"),
    ("EDP", "Expert"),
    ("MOY", "Collaborateur moyen"),
]


def _normalize_pair(code: str, label: str) -> Tuple[str, str]:
    code = (code or "").strip()
    label = (label or "").strip()
    if not code:
        return "", ""
    if not label:
        label = code
    return code, label


def _fetch_db_categories() -> Iterable[Tuple[str, str]]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT DISTINCT categorie, libelle
            FROM categorie_cout
            WHERE categorie IS NOT NULL AND TRIM(categorie) != ''
            """
        )
        for code, label in cursor.fetchall():
            normalized = _normalize_pair(code, label)
            if normalized[0]:
                yield normalized


@lru_cache(maxsize=1)
def get_category_mappings() -> Tuple[Dict[str, str], Dict[str, str]]:
    """
    Retourne deux dictionnaires :
      - code -> libellé (affichage)
      - libellé/code (lowercase) -> code canonique
    """
    code_to_label: Dict[str, str] = {}
    label_to_code: Dict[str, str] = {}

    def register(pairs: Iterable[Tuple[str, str]]) -> None:
        for code, label in pairs:
            if not code:
                continue
            code_to_label[code] = label
            label_to_code[label.lower()] = code
            label_to_code[code.lower()] = code

    register(DEFAULT_CATEGORIES)
    register(_fetch_db_categories())
    return code_to_label, label_to_code


def resolve_category_code(value: str) -> str:
    """
    Retourne le code de catégorie correspondant au libellé ou au code fourni.
    Si la valeur n'est pas connue, elle est renvoyée telle quelle (après strip).
    """
    if value is None:
        return ""
    stripped = value.strip()
    if not stripped:
        return ""
    _, label_to_code = get_category_mappings()
    return label_to_code.get(stripped.lower(), stripped)


def get_category_label(code: str) -> str:
    """Retourne le libellé lisible associé au code."""
    if code is None:
        return ""
    code = code.strip()
    code_to_label, _ = get_category_mappings()
    return code_to_label.get(code, code)


def list_category_labels() -> List[str]:
    """Renvoie tous les libellés connus (défaut + personnalisés)."""
    code_to_label, _ = get_category_mappings()
    unique_labels = {label for label in code_to_label.values() if label}
    return sorted(unique_labels)


def invalidate_category_cache() -> None:
    """Vide le cache pour refléter les modifications utilisateur."""
    get_category_mappings.cache_clear()
