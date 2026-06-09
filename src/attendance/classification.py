from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from itertools import product
import json
from pathlib import Path
import re
from typing import Mapping, Sequence
import unicodedata


EVENT_KEYS = ("entry", "lunch_out", "lunch_return", "exit")
RIGID_EVENT_KEYS = ("entry", "exit")
FLEXIBLE_EVENT_KEYS = ("lunch_out", "lunch_return")
PROTECTED_LUNCH_BEFORE_MINUTES = 15
PROTECTED_LUNCH_AFTER_MINUTES = 15
DUPLICATE_PUNCH_SECONDS = 5 * 60
EVENT_LABELS = {
    "entry": "entrada",
    "lunch_out": "inicio de comida",
    "lunch_return": "fin de comida",
    "exit": "salida final",
}
PUBLIC_EVENT_KEYS = {
    "entry": "entrada",
    "lunch_out": "inicio_comida",
    "lunch_return": "fin_comida",
    "exit": "salida",
}
EVENT_ALIASES = {
    "entry": "entry",
    "entrada": "entry",
    "lunch_out": "lunch_out",
    "inicio_comida": "lunch_out",
    "inicio comida": "lunch_out",
    "salida_comida": "lunch_out",
    "salida comida": "lunch_out",
    "lunch_return": "lunch_return",
    "fin_comida": "lunch_return",
    "fin comida": "lunch_return",
    "regreso_comida": "lunch_return",
    "regreso comida": "lunch_return",
    "exit": "exit",
    "salida": "exit",
    "salida_final": "exit",
    "salida final": "exit",
}


@dataclass(frozen=True)
class EventWindow:
    before_minutes: int
    after_minutes: int
    max_before_minutes: int
    max_after_minutes: int

    def __post_init__(self) -> None:
        if min(
            self.before_minutes,
            self.after_minutes,
            self.max_before_minutes,
            self.max_after_minutes,
        ) < 0:
            raise ValueError("Las ventanas no pueden contener minutos negativos.")
        if self.max_before_minutes < self.before_minutes:
            raise ValueError("max_before_minutes debe ser mayor o igual que before_minutes.")
        if self.max_after_minutes < self.after_minutes:
            raise ValueError("max_after_minutes debe ser mayor o igual que after_minutes.")


# Entrada y salida usan estas ventanas como límites de clasificación.
# Las ventanas de comida son referencias débiles y nunca reglas operativas.
DEFAULT_WINDOWS: dict[str, EventWindow] = {
    "entry": EventWindow(120, 120, 180, 210),
    "lunch_out": EventWindow(60, 60, 120, 180),
    "lunch_return": EventWindow(60, 90, 120, 180),
    "exit": EventWindow(120, 240, 180, 360),
}


@dataclass(frozen=True)
class ClassificationPolicy:
    windows: Mapping[str, EventWindow] = field(default_factory=lambda: dict(DEFAULT_WINDOWS))
    minimum_score: float = 30.0
    ambiguity_margin: float = 8.0
    max_candidates_per_event: int = 12
    isolated_lunch_decisive_minutes: int = 15
    maximum_lunch_pair_minutes: int = 240
    duplicate_punch_seconds: int = DUPLICATE_PUNCH_SECONDS

    def __post_init__(self) -> None:
        missing = [event_key for event_key in EVENT_KEYS if event_key not in self.windows]
        if missing:
            raise ValueError("Faltan ventanas para: " + ", ".join(missing))
        if self.minimum_score < 0 or self.ambiguity_margin < 0:
            raise ValueError("Los parámetros de score no pueden ser negativos.")
        if self.max_candidates_per_event < 1:
            raise ValueError("max_candidates_per_event debe ser al menos 1.")
        if self.isolated_lunch_decisive_minutes < 0:
            raise ValueError("isolated_lunch_decisive_minutes no puede ser negativo.")
        if self.maximum_lunch_pair_minutes < 45:
            raise ValueError("maximum_lunch_pair_minutes debe ser al menos 45.")
        if self.duplicate_punch_seconds < 0:
            raise ValueError("duplicate_punch_seconds no puede ser negativo.")


@dataclass(frozen=True)
class Punch:
    punch_id: int
    checked_at: datetime
    state: str = ""


@dataclass(frozen=True)
class DuplicatePunch:
    duplicate: Punch
    original: Punch
    seconds_apart: int
    block: str


@dataclass(frozen=True)
class NormalizedPunches:
    usable_punches: list[Punch]
    duplicate_punches: list[DuplicatePunch]


@dataclass(frozen=True)
class ExpectedEvent:
    key: str
    expected_at: datetime
    window: EventWindow


@dataclass(frozen=True)
class CandidateScore:
    punch_id: int
    event_key: str
    score: float
    signed_distance_minutes: float
    inside_window: bool
    state_match: bool
    state_conflict: bool
    basis: str


@dataclass(frozen=True)
class Hypothesis:
    score: float
    assignments: Mapping[str, int | None]
    sequence_invalid: bool


@dataclass
class ClassificationResult:
    assignments: dict[str, Punch | None]
    candidates: dict[tuple[int, str], CandidateScore]
    confidence: dict[int, CandidateScore]
    unused_punches: list[Punch]
    ambiguous_punches: list[Punch]
    duplicate_punches: list[DuplicatePunch]
    technical_flags: list[str]
    hypothesis_score: float
    alternative_score: float | None


@dataclass(frozen=True)
class ClassificationConfiguration:
    default_policy: ClassificationPolicy | Mapping[str, object] | None = None
    schedule_policies: Mapping[str, ClassificationPolicy | Mapping[str, object]] = field(default_factory=dict)
    employee_policies: Mapping[str, ClassificationPolicy | Mapping[str, object]] = field(default_factory=dict)


def _normalize_text(value: object) -> str:
    text = "" if value is None else str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text)


def _event_key(value: str) -> str:
    normalized = _normalize_text(value).replace("-", "_")
    key = EVENT_ALIASES.get(normalized)
    if key is None:
        raise ValueError(f"Evento de turno no reconocido: {value!r}.")
    return key


def _parse_time(value: object) -> time:
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    if not isinstance(value, str):
        raise TypeError(f"Hora no soportada: {value!r}.")
    for pattern in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(value.strip(), pattern).time()
        except ValueError:
            continue
    raise ValueError(f"Hora inválida: {value!r}.")


def _parse_window(value: object, default: EventWindow) -> EventWindow:
    if isinstance(value, EventWindow):
        return value
    if isinstance(value, int):
        return EventWindow(value, value, value, value)
    if not isinstance(value, Mapping):
        raise TypeError("Cada tolerancia debe ser un entero, EventWindow o diccionario.")
    before = int(value.get("antes", value.get("before_minutes", default.before_minutes)))
    after = int(value.get("despues", value.get("after_minutes", default.after_minutes)))
    max_before = int(value.get("max_antes", value.get("max_before_minutes", default.max_before_minutes)))
    max_after = int(value.get("max_despues", value.get("max_after_minutes", default.max_after_minutes)))
    return EventWindow(before, after, max_before, max_after)


def resolve_policy(
    tolerances: ClassificationPolicy | Mapping[str, object] | None = None,
    base_policy: ClassificationPolicy | None = None,
) -> ClassificationPolicy:
    if tolerances is None:
        return base_policy or ClassificationPolicy()
    if isinstance(tolerances, ClassificationPolicy):
        return tolerances
    if not isinstance(tolerances, Mapping):
        raise TypeError("tolerancias debe ser ClassificationPolicy, diccionario o None.")

    base_policy = base_policy or ClassificationPolicy()
    windows = dict(base_policy.windows)
    for raw_key, value in tolerances.items():
        try:
            key = _event_key(str(raw_key))
        except ValueError:
            continue
        windows[key] = _parse_window(value, windows[key])
    return ClassificationPolicy(
        windows=windows,
        minimum_score=float(
            tolerances.get("score_minimo", tolerances.get("minimum_score", base_policy.minimum_score))
        ),
        ambiguity_margin=float(
            tolerances.get(
                "margen_ambiguedad",
                tolerances.get("ambiguity_margin", base_policy.ambiguity_margin),
            )
        ),
        max_candidates_per_event=int(
            tolerances.get(
                "max_candidatos_por_evento",
                tolerances.get("max_candidates_per_event", base_policy.max_candidates_per_event),
            )
        ),
        isolated_lunch_decisive_minutes=int(
            tolerances.get(
                "comida_aislada_minutos_decisivos",
                tolerances.get(
                    "isolated_lunch_decisive_minutes",
                    base_policy.isolated_lunch_decisive_minutes,
                ),
            )
        ),
        maximum_lunch_pair_minutes=int(
            tolerances.get(
                "maximo_par_comida_minutos",
                tolerances.get(
                    "maximum_lunch_pair_minutes",
                    base_policy.maximum_lunch_pair_minutes,
                ),
            )
        ),
        duplicate_punch_seconds=int(
            tolerances.get(
                "duplicado_segundos",
                tolerances.get(
                    "duplicate_punch_seconds",
                    base_policy.duplicate_punch_seconds,
                ),
            )
        ),
    )


def load_classification_configuration(path: str | Path) -> ClassificationConfiguration:
    config_path = Path(path)
    data = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(data, Mapping):
        raise ValueError("La configuración de clasificación debe ser un objeto JSON.")

    default_policy = data.get("predeterminada", data.get("default"))
    schedule_policies = data.get("turnos", data.get("schedules", {}))
    employee_policies = data.get("empleados", data.get("employees", {}))
    if default_policy is not None and not isinstance(default_policy, Mapping):
        raise ValueError("'predeterminada' debe ser un objeto JSON.")
    if not isinstance(schedule_policies, Mapping) or not isinstance(employee_policies, Mapping):
        raise ValueError("'turnos' y 'empleados' deben ser objetos JSON.")

    base_policy = resolve_policy(default_policy)
    for label, policy in {**schedule_policies, **employee_policies}.items():
        if not isinstance(policy, Mapping):
            raise ValueError(f"La política {label!r} debe ser un objeto JSON.")
        resolve_policy(policy, base_policy)
    return ClassificationConfiguration(
        default_policy=default_policy,
        schedule_policies={str(key): value for key, value in schedule_policies.items()},
        employee_policies={str(key): value for key, value in employee_policies.items()},
    )


def build_expected_events(
    turno: Mapping[str, object],
    work_date: date,
    policy: ClassificationPolicy,
) -> list[ExpectedEvent]:
    times: dict[str, time] = {}
    for raw_key, value in turno.items():
        try:
            key = _event_key(str(raw_key))
        except ValueError:
            continue
        times[key] = _parse_time(value)
    missing = [EVENT_LABELS[key] for key in EVENT_KEYS if key not in times]
    if missing:
        raise ValueError("El turno no contiene todos los eventos esperados: " + ", ".join(missing))
    return [
        ExpectedEvent(key, datetime.combine(work_date, times[key]), policy.windows[key])
        for key in EVENT_KEYS
    ]


def _state_hint(state: str) -> tuple[str | None, bool]:
    normalized = _normalize_text(state)
    strong_hints = (
        ("salida a descanso", "lunch_out"),
        ("salida de comida", "lunch_out"),
        ("inicio comida", "lunch_out"),
        ("regreso descanso", "lunch_return"),
        ("regreso de comida", "lunch_return"),
        ("fin comida", "lunch_return"),
        ("salida final", "exit"),
        ("entrada laboral", "entry"),
    )
    for phrase, event_key in strong_hints:
        if phrase in normalized:
            return event_key, True
    if normalized == "entrada":
        return "entry", False
    if normalized == "salida":
        return "exit", False
    return None, False


def _rigid_candidate_score(punch: Punch, event: ExpectedEvent) -> CandidateScore | None:
    signed_distance = (punch.checked_at - event.expected_at).total_seconds() / 60
    distance = abs(signed_distance)
    strict_limit = event.window.before_minutes if signed_distance < 0 else event.window.after_minutes
    maximum_limit = event.window.max_before_minutes if signed_distance < 0 else event.window.max_after_minutes
    inside_window = distance <= strict_limit
    if inside_window:
        score = 100.0 - 45.0 * (distance / max(strict_limit, 1))
    elif distance <= maximum_limit:
        score = 50.0 - 25.0 * ((distance - strict_limit) / max(maximum_limit - strict_limit, 1))
    else:
        return None
    hint, strong_hint = _state_hint(punch.state)
    state_match = hint == event.key
    state_conflict = bool(strong_hint and hint is not None and hint != event.key)
    if state_match:
        score += 30.0 if strong_hint else 6.0
    elif state_conflict:
        score -= 45.0
    if score < 0:
        return None
    return CandidateScore(
        punch.punch_id,
        event.key,
        round(score, 3),
        signed_distance,
        inside_window,
        state_match,
        state_conflict,
        "ventana rígida",
    )


def _contextual_late_entry_candidate(punch: Punch, event: ExpectedEvent) -> CandidateScore:
    signed_distance = (punch.checked_at - event.expected_at).total_seconds() / 60
    hint, strong_hint = _state_hint(punch.state)
    state_match = hint == "entry"
    state_conflict = bool(strong_hint and hint is not None and hint != "entry")
    score = max(65.0, 90.0 - max(0.0, signed_distance) * 0.08)
    if state_match:
        score += 30.0 if strong_hint else 6.0
    elif state_conflict:
        score -= 45.0
    return CandidateScore(
        punch.punch_id,
        "entry",
        round(score, 3),
        signed_distance,
        False,
        state_match,
        state_conflict,
        "entrada tardía contextual",
    )


def _flexible_candidate_score(punch: Punch, event: ExpectedEvent) -> CandidateScore:
    signed_distance = (punch.checked_at - event.expected_at).total_seconds() / 60
    distance = abs(signed_distance)
    strict_limit = event.window.before_minutes if signed_distance < 0 else event.window.after_minutes
    # La referencia horaria pesa poco y nunca excluye una comida coherente.
    score = max(18.0, 42.0 - min(distance, 480.0) * 0.05)
    hint, strong_hint = _state_hint(punch.state)
    state_match = hint == event.key
    state_conflict = bool(strong_hint and hint is not None and hint != event.key)
    if state_match:
        score += 30.0 if strong_hint else 6.0
    elif state_conflict:
        score -= 35.0
    return CandidateScore(
        punch.punch_id,
        event.key,
        round(score, 3),
        signed_distance,
        distance <= strict_limit,
        state_match,
        state_conflict,
        "referencia flexible",
    )


def _inside_protected_lunch_zone(
    punch: Punch,
    events_by_key: Mapping[str, ExpectedEvent],
) -> bool:
    protected_start = events_by_key["lunch_out"].expected_at - timedelta(
        minutes=PROTECTED_LUNCH_BEFORE_MINUTES
    )
    protected_end = events_by_key["lunch_return"].expected_at + timedelta(
        minutes=PROTECTED_LUNCH_AFTER_MINUTES
    )
    return protected_start <= punch.checked_at <= protected_end


def _has_strong_exit_hint(punch: Punch) -> bool:
    hint, strong_hint = _state_hint(punch.state)
    return strong_hint and hint == "exit"


def _duplicate_block_label(block: str) -> str:
    if block in FLEXIBLE_EVENT_KEYS or block == "lunch":
        return "comida"
    return EVENT_LABELS.get(block, block)


def _same_strong_hint_block(left: Punch, right: Punch) -> str | None:
    left_hint, left_strong = _state_hint(left.state)
    right_hint, right_strong = _state_hint(right.state)
    if left_strong and right_strong and left_hint != right_hint:
        return None
    strong_hints = [hint for hint, strong in ((left_hint, left_strong), (right_hint, right_strong)) if strong]
    if not strong_hints:
        return None
    hint = strong_hints[0]
    if hint in FLEXIBLE_EVENT_KEYS:
        return "lunch"
    return hint


def _rigid_duplicate_block(
    left: Punch,
    right: Punch,
    event: ExpectedEvent,
    policy: ClassificationPolicy,
) -> str | None:
    left_candidate = _rigid_candidate_score(left, event)
    right_candidate = _rigid_candidate_score(right, event)
    if (
        left_candidate is not None
        and right_candidate is not None
        and left_candidate.score >= policy.minimum_score
        and right_candidate.score >= policy.minimum_score
    ):
        return event.key
    return None


def _flexible_duplicate_block(
    left: Punch,
    right: Punch,
    event: ExpectedEvent,
) -> str | None:
    left_candidate = _flexible_candidate_score(left, event)
    right_candidate = _flexible_candidate_score(right, event)
    if left_candidate.inside_window and right_candidate.inside_window:
        return "lunch"
    return None


def _duplicate_block_for_pair(
    left: Punch,
    right: Punch,
    expected_events: Sequence[ExpectedEvent],
    policy: ClassificationPolicy,
) -> str | None:
    hinted_block = _same_strong_hint_block(left, right)
    if hinted_block is not None:
        return hinted_block

    events_by_key = {event.key: event for event in expected_events}
    if _inside_protected_lunch_zone(left, events_by_key) and _inside_protected_lunch_zone(right, events_by_key):
        return "lunch"

    for event_key in ("entry", "exit"):
        block = _rigid_duplicate_block(left, right, events_by_key[event_key], policy)
        if block is not None:
            return block

    for event_key in FLEXIBLE_EVENT_KEYS:
        block = _flexible_duplicate_block(left, right, events_by_key[event_key])
        if block is not None:
            return block

    return None


def normalize_punches(
    punches: Sequence[Punch],
    expected_events: Sequence[ExpectedEvent],
    policy: ClassificationPolicy,
) -> NormalizedPunches:
    if policy.duplicate_punch_seconds <= 0:
        return NormalizedPunches(list(punches), [])

    usable_punches: list[Punch] = []
    duplicate_punches: list[DuplicatePunch] = []
    for punch in punches:
        if not usable_punches:
            usable_punches.append(punch)
            continue
        previous = usable_punches[-1]
        seconds_apart = int((punch.checked_at - previous.checked_at).total_seconds())
        if 0 <= seconds_apart <= policy.duplicate_punch_seconds:
            block = _duplicate_block_for_pair(previous, punch, expected_events, policy)
            if block is not None:
                duplicate_punches.append(
                    DuplicatePunch(
                        duplicate=punch,
                        original=previous,
                        seconds_apart=seconds_apart,
                        block=block,
                    )
                )
                continue
        usable_punches.append(punch)
    return NormalizedPunches(usable_punches, duplicate_punches)


def _contextual_late_entry_punch_id(
    punches: Sequence[Punch],
    expected_events: Sequence[ExpectedEvent],
    policy: ClassificationPolicy,
) -> int | None:
    if len(punches) < 2:
        return None
    events = {event.key: event for event in expected_events}
    first = punches[0]
    if first.checked_at >= events["lunch_out"].expected_at:
        return None
    hint, strong_hint = _state_hint(first.state)
    if strong_hint and hint in FLEXIBLE_EVENT_KEYS:
        return None
    later_exit_exists = any(
        (candidate := _rigid_candidate_score(punch, events["exit"])) is not None
        and candidate.score >= policy.minimum_score
        for punch in punches[1:]
    )
    return first.punch_id if later_exit_exists else None


def _score_hypothesis(
    assignments: Mapping[str, int | None],
    punches_by_id: Mapping[int, Punch],
    candidates: Mapping[tuple[int, str], CandidateScore],
    policy: ClassificationPolicy,
    reference_lunch_seconds: int,
) -> tuple[float, bool]:
    assigned = [
        (key, punches_by_id[punch_id].checked_at)
        for key in EVENT_KEYS
        if (punch_id := assignments[key]) is not None
    ]
    sequence_invalid = any(
        left_time >= right_time
        for index, (_, left_time) in enumerate(assigned)
        for _, right_time in assigned[index + 1 :]
    )
    if sequence_invalid:
        return -10000.0, True

    score = sum(
        candidates[(punch_id, event_key)].score
        for event_key, punch_id in assignments.items()
        if punch_id is not None
    )
    lunch_out_id = assignments["lunch_out"]
    lunch_return_id = assignments["lunch_return"]
    if lunch_out_id is not None and lunch_return_id is not None:
        duration_seconds = (
            punches_by_id[lunch_return_id].checked_at - punches_by_id[lunch_out_id].checked_at
        ).total_seconds()
        if duration_seconds <= 0 or duration_seconds > policy.maximum_lunch_pair_minutes * 60:
            return -10000.0, True
        score += 180.0
        if duration_seconds <= reference_lunch_seconds:
            score += 80.0
        else:
            excess_minutes = (duration_seconds - reference_lunch_seconds) / 60
            score += max(-120.0, 80.0 - excess_minutes * 2.0)
        if duration_seconds < 5 * 60:
            score -= 90.0
    elif lunch_return_id is not None:
        score -= 35.0
    elif lunch_out_id is not None:
        score += 20.0
    if assignments["entry"] is not None and assignments["exit"] is not None:
        score += 20.0
    return round(score, 3), False


def _build_hypotheses(
    punches: Sequence[Punch],
    expected_events: Sequence[ExpectedEvent],
    policy: ClassificationPolicy,
) -> tuple[list[Hypothesis], dict[tuple[int, str], CandidateScore]]:
    punches_by_id = {punch.punch_id: punch for punch in punches}
    events_by_key = {event.key: event for event in expected_events}
    reference_lunch_seconds = int(
        (
            events_by_key["lunch_return"].expected_at
            - events_by_key["lunch_out"].expected_at
        ).total_seconds()
    )
    if reference_lunch_seconds <= 0:
        raise ValueError("La referencia de regreso de comida debe ser posterior al inicio.")
    contextual_entry_id = _contextual_late_entry_punch_id(punches, expected_events, policy)
    candidates: dict[tuple[int, str], CandidateScore] = {}
    choices: list[list[int | None]] = []
    for event in expected_events:
        event_candidates: list[CandidateScore] = []
        for punch in punches:
            if event.key == "entry" and punch.punch_id == contextual_entry_id:
                rigid_candidate = _rigid_candidate_score(punch, event)
                contextual_candidate = _contextual_late_entry_candidate(punch, event)
                candidate = max(
                    (item for item in (rigid_candidate, contextual_candidate) if item is not None),
                    key=lambda item: item.score,
                )
            else:
                candidate = (
                    _rigid_candidate_score(punch, event)
                    if event.key in RIGID_EVENT_KEYS
                    else _flexible_candidate_score(punch, event)
                )
            if candidate is None:
                continue
            if (
                event.key == "exit"
                and _inside_protected_lunch_zone(punch, events_by_key)
                and not _has_strong_exit_hint(punch)
            ):
                continue
            if event.key in RIGID_EVENT_KEYS and candidate.score < policy.minimum_score:
                continue
            candidates[(punch.punch_id, event.key)] = candidate
            event_candidates.append(candidate)
        event_candidates.sort(key=lambda item: item.score, reverse=True)
        choices.append([None] + [item.punch_id for item in event_candidates[: policy.max_candidates_per_event]])

    hypotheses: list[Hypothesis] = []
    for selected in product(*choices):
        selected_ids = [punch_id for punch_id in selected if punch_id is not None]
        if len(selected_ids) != len(set(selected_ids)):
            continue
        assignments = dict(zip(EVENT_KEYS, selected))
        if contextual_entry_id is not None and assignments["entry"] != contextual_entry_id:
            continue
        score, invalid = _score_hypothesis(
            assignments,
            punches_by_id,
            candidates,
            policy,
            reference_lunch_seconds,
        )
        if invalid:
            continue
        hypotheses.append(Hypothesis(score, assignments, False))
    hypotheses.sort(key=lambda item: item.score, reverse=True)
    return hypotheses, candidates


def classify_punches(
    punches: Sequence[Punch],
    expected_events: Sequence[ExpectedEvent],
    policy: ClassificationPolicy | None = None,
) -> ClassificationResult:
    policy = policy or ClassificationPolicy()
    events_by_key = {event.key: event for event in expected_events}
    if set(events_by_key) != set(EVENT_KEYS):
        raise ValueError("Deben proporcionarse entrada, inicio/fin de comida y salida.")
    sorted_punches = sorted(punches, key=lambda item: (item.checked_at, item.punch_id))
    punches_by_id = {punch.punch_id: punch for punch in sorted_punches}
    if len(punches_by_id) != len(sorted_punches):
        raise ValueError("Cada checada debe tener un punch_id único.")

    normalized = normalize_punches(
        sorted_punches,
        [events_by_key[key] for key in EVENT_KEYS],
        policy,
    )
    usable_punches = normalized.usable_punches
    usable_punches_by_id = {punch.punch_id: punch for punch in usable_punches}

    hypotheses, candidates = _build_hypotheses(
        usable_punches,
        [events_by_key[key] for key in EVENT_KEYS],
        policy,
    )
    best = hypotheses[0]
    second = hypotheses[1] if len(hypotheses) > 1 else None
    selected = dict(best.assignments)
    ambiguous_ids: set[int] = set()
    partial_meal_ambiguous = False

    if second is not None and best.score - second.score < policy.ambiguity_margin:
        best_by_punch = {punch_id: key for key, punch_id in best.assignments.items() if punch_id is not None}
        second_by_punch = {punch_id: key for key, punch_id in second.assignments.items() if punch_id is not None}
        ambiguous_ids = {
            punch_id
            for punch_id in set(best_by_punch) | set(second_by_punch)
            if best_by_punch.get(punch_id) != second_by_punch.get(punch_id)
        }

    contextual_entry = selected["entry"]
    contextual_entry_selected = (
        contextual_entry is not None
        and candidates[(contextual_entry, "entry")].basis == "entrada tardía contextual"
    )
    if contextual_entry_selected and selected["exit"] is not None:
        intermediate_punches = [
            punch
            for punch in usable_punches
            if punch.punch_id not in {selected["entry"], selected["exit"]}
        ]
        if len(intermediate_punches) == 1:
            intermediate = intermediate_punches[0]
            hint, strong_hint = _state_hint(intermediate.state)
            if not (strong_hint and hint in FLEXIBLE_EVENT_KEYS):
                for event_key in FLEXIBLE_EVENT_KEYS:
                    if selected[event_key] == intermediate.punch_id:
                        selected[event_key] = None
                ambiguous_ids.add(intermediate.punch_id)
                partial_meal_ambiguous = True

    lunch_out_id = selected["lunch_out"]
    if len(usable_punches) == 1 and lunch_out_id is not None:
        lunch_candidate = candidates[(lunch_out_id, "lunch_out")]
        if abs(lunch_candidate.signed_distance_minutes) > policy.isolated_lunch_decisive_minutes:
            ambiguous_ids.add(lunch_out_id)

    for event_key, punch_id in list(selected.items()):
        if punch_id in ambiguous_ids:
            selected[event_key] = None

    assignments = {
        event_key: usable_punches_by_id[punch_id] if punch_id is not None else None
        for event_key, punch_id in selected.items()
    }
    confidence = {
        punch_id: candidates[(punch_id, event_key)]
        for event_key, punch_id in selected.items()
        if punch_id is not None
    }
    selected_ids = set(confidence)
    ambiguous_punches = [punch for punch in sorted_punches if punch.punch_id in ambiguous_ids]
    unused_punches = [punch for punch in sorted_punches if punch.punch_id not in selected_ids]

    technical_flags: list[str] = []
    if normalized.duplicate_punches:
        technical_flags.append("Checada duplicada omitida por normalizacion")
    if ambiguous_punches:
        technical_flags.append("Asignación ambigua")
    if contextual_entry_selected:
        technical_flags.append("Entrada tardía contextual")
    if partial_meal_ambiguous:
        technical_flags.append("Comida parcial ambigua")
    if any(
        candidate.event_key in FLEXIBLE_EVENT_KEYS and not candidate.inside_window
        for candidate in confidence.values()
    ):
        technical_flags.append("Comida fuera de referencia horaria")
    if any(punch.punch_id not in ambiguous_ids for punch in unused_punches):
        technical_flags.append("Checada no utilizada por el clasificador")
    selected_lunch_out = assignments["lunch_out"]
    if assignments["entry"] is None and selected_lunch_out is not None:
        lunch_candidate = confidence[selected_lunch_out.punch_id]
        if abs(lunch_candidate.signed_distance_minutes) > policy.isolated_lunch_decisive_minutes:
            technical_flags.append("Posible entrada tardía")

    return ClassificationResult(
        assignments=assignments,
        candidates=candidates,
        confidence=confidence,
        unused_punches=unused_punches,
        ambiguous_punches=ambiguous_punches,
        duplicate_punches=normalized.duplicate_punches,
        technical_flags=technical_flags,
        hypothesis_score=best.score,
        alternative_score=second.score if second is not None else None,
    )


def _format_time(value: datetime | None) -> str | None:
    return value.strftime("%H:%M:%S") if value is not None else None


def _assignment_reason(
    punch: Punch,
    candidate: CandidateScore,
    expected_events: Sequence[ExpectedEvent],
) -> str:
    distance = abs(candidate.signed_distance_minutes)
    if candidate.event_key in FLEXIBLE_EVENT_KEYS:
        reason = (
            f"Asignada como {EVENT_LABELS[candidate.event_key]} por contexto cronológico flexible; "
            f"la referencia teórica está a {distance:.1f} min y solo actuó como desempate débil."
        )
    elif candidate.basis == "entrada tardía contextual":
        reason = (
            "Asignada como entrada tardía contextual porque es la primera checada anterior "
            "a la referencia de comida y existe una salida final posterior plausible."
        )
    else:
        position = "dentro" if candidate.inside_window else "en la extensión"
        reason = (
            f"Asignada como {EVENT_LABELS[candidate.event_key]} por cercanía a la ventana rígida; "
            f"está {position} y a {distance:.1f} min de la hora objetivo."
        )
    if candidate.event_key == "lunch_out":
        entry = next(event for event in expected_events if event.key == "entry")
        entry_distance = abs((punch.checked_at - entry.expected_at).total_seconds() / 60)
        reason += f" Está a {entry_distance:.1f} min de la entrada esperada."
    if candidate.state_match:
        reason += " El tipo registrado por el checador coincide."
    return reason


def format_classification_audit(
    punches: Sequence[Punch],
    expected_events: Sequence[ExpectedEvent],
    result: ClassificationResult,
) -> str:
    parts = [
        f"Hipótesis ganadora={result.hypothesis_score:.1f}"
        + (f"; alternativa={result.alternative_score:.1f}" if result.alternative_score is not None else "")
    ]
    if result.technical_flags:
        parts.append("Alertas técnicas=" + ", ".join(result.technical_flags))
    ambiguous_ids = {punch.punch_id for punch in result.ambiguous_punches}
    duplicates_by_id = {duplicate.duplicate.punch_id: duplicate for duplicate in result.duplicate_punches}
    for punch in sorted(punches, key=lambda item: (item.checked_at, item.punch_id)):
        checked_at = punch.checked_at.strftime("%H:%M:%S")
        selected = result.confidence.get(punch.punch_id)
        duplicate = duplicates_by_id.get(punch.punch_id)
        punch_candidates = sorted(
            (candidate for (punch_id, _), candidate in result.candidates.items() if punch_id == punch.punch_id),
            key=lambda item: item.score,
            reverse=True,
        )
        alternatives = ", ".join(
            f"{EVENT_LABELS[candidate.event_key]}={candidate.score:.1f}"
            for candidate in punch_candidates
        ) or "ninguna"
        if selected is not None:
            parts.append(
                f"{checked_at} -> {EVENT_LABELS[selected.event_key]} (score={selected.score:.1f}). "
                f"{_assignment_reason(punch, selected, expected_events)} Alternativas: {alternatives}."
            )
        elif duplicate is not None:
            original_at = duplicate.original.checked_at.strftime("%H:%M:%S")
            parts.append(
                f"{checked_at} -> duplicada/no utilizada. "
                f"Se conservo {original_at}; diferencia={duplicate.seconds_apart} s; "
                f"bloque probable={_duplicate_block_label(duplicate.block)}. Alternativas: {alternatives}."
            )
        elif punch.punch_id in ambiguous_ids:
            parts.append(f"{checked_at} -> sin asignar por ambigüedad. Alternativas: {alternatives}.")
        else:
            parts.append(f"{checked_at} -> no utilizada. Alternativas: {alternatives}.")
    return " | ".join(parts)


def _unique_time_key(punch: Punch, used_keys: set[str]) -> str:
    base = punch.checked_at.strftime("%H:%M:%S")
    if base not in used_keys:
        used_keys.add(base)
        return base
    suffix = 2
    while f"{base}#{suffix}" in used_keys:
        suffix += 1
    key = f"{base}#{suffix}"
    used_keys.add(key)
    return key


def clasificar_checadas(
    checadas: Sequence[str | datetime],
    turno: Mapping[str, object],
    tolerancias: ClassificationPolicy | Mapping[str, object] | None = None,
) -> dict[str, object]:
    from .business import BusinessPolicy, evaluate_business

    policy = resolve_policy(tolerancias)
    explicit_dates = {value.date() for value in checadas if isinstance(value, datetime)}
    if len(explicit_dates) > 1:
        raise ValueError("Todas las checadas deben pertenecer al mismo día.")
    base_date = next(iter(explicit_dates), date(2000, 1, 3))
    parsed = [
        value if isinstance(value, datetime) else datetime.combine(base_date, _parse_time(value))
        for value in checadas
    ]
    punches = [Punch(index, checked_at) for index, checked_at in enumerate(parsed)]
    expected_events = build_expected_events(turno, base_date, policy)
    result = classify_punches(punches, expected_events, policy)
    expected_by_key = {event.key: event for event in expected_events}
    maximum_lunch_seconds = int(
        (
            expected_by_key["lunch_return"].expected_at
            - expected_by_key["lunch_out"].expected_at
        ).total_seconds()
    )
    evaluation = evaluate_business(
        result,
        expected_events,
        policy=BusinessPolicy(maximum_lunch_seconds=maximum_lunch_seconds),
    )

    confidence: dict[str, dict[str, object]] = {}
    used_keys: set[str] = set()
    duplicates_by_id = {duplicate.duplicate.punch_id: duplicate for duplicate in result.duplicate_punches}
    for punch in sorted(punches, key=lambda item: (item.checked_at, item.punch_id)):
        key = _unique_time_key(punch, used_keys)
        candidate = result.confidence.get(punch.punch_id)
        duplicate = duplicates_by_id.get(punch.punch_id)
        if candidate is not None:
            confidence[key] = {
                "evento_asignado": PUBLIC_EVENT_KEYS[candidate.event_key],
                "score": f"{candidate.score:.1f}",
                "razon": _assignment_reason(punch, candidate, expected_events),
            }
        elif duplicate is not None:
            original_at = duplicate.original.checked_at.strftime("%H:%M:%S")
            confidence[key] = {
                "evento_asignado": None,
                "score": "0.0",
                "razon": (
                    f"Checada duplicada cercana a {original_at}; "
                    f"diferencia={duplicate.seconds_apart} s; "
                    f"bloque probable={_duplicate_block_label(duplicate.block)}."
                ),
            }
        else:
            confidence[key] = {
                "evento_asignado": None,
                "score": "0.0",
                "razon": "La checada no se asignó porque la evidencia fue ambigua o insuficiente.",
            }

    return {
        "entrada": _format_time(evaluation.assignments["entry"]),
        "inicio_comida": _format_time(evaluation.assignments["lunch_out"]),
        "fin_comida": _format_time(evaluation.assignments["lunch_return"]),
        "salida": _format_time(evaluation.assignments["exit"]),
        "status": evaluation.status,
        "detalle": evaluation.detail,
        "incidencias": evaluation.operational_incidents,
        "checadas_no_utilizadas": [
            punch.checked_at.strftime("%H:%M:%S") for punch in result.unused_punches
        ],
        "checadas_duplicadas": [
            {
                "duplicada": duplicate.duplicate.checked_at.strftime("%H:%M:%S"),
                "original": duplicate.original.checked_at.strftime("%H:%M:%S"),
                "diferencia_segundos": duplicate.seconds_apart,
                "bloque_probable": _duplicate_block_label(duplicate.block),
            }
            for duplicate in result.duplicate_punches
        ],
        "confianza": confidence,
        "auditoria": format_classification_audit(punches, expected_events, result),
    }
