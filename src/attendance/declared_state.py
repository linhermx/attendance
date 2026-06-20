from __future__ import annotations

import re
import unicodedata


DECLARED_STATE_TO_EVENT = {
    "entrada": "entry",
    "salida a descanso": "lunch_out",
    "regreso descanso": "lunch_return",
    "salida": "exit",
}


def normalize_declared_state(value: object) -> str:
    text = "" if value is None else str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text)


def map_declared_state(value: object) -> str | None:
    normalized = normalize_declared_state(value)
    if not normalized:
        return None
    return DECLARED_STATE_TO_EVENT.get(normalized)


def is_valid_declared_state(value: object) -> bool:
    return map_declared_state(value) is not None

