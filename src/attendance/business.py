from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import math
from typing import Mapping, Sequence

from .classification import ClassificationResult, EVENT_KEYS, ExpectedEvent
from .time_calculation import calculate_worked_time


@dataclass(frozen=True)
class BusinessPolicy:
    maximum_lunch_seconds: int = 45 * 60
    severe_tardy_minutes: int = 60
    entry_grace_seconds: int = 59


@dataclass
class BusinessEvaluation:
    assignments: dict[str, datetime | None]
    operational_incidents: list[str]
    status: str
    detail: str
    tardy_minutes: int
    lunch_duration_seconds: int | None
    lunch_minutes: int | None
    early_leave_minutes: int | None
    worked_minutes: int | None


def calculate_worked_minutes(
    assignments: Mapping[str, datetime | None],
    *,
    scheduled_entry: datetime | None = None,
    entry_grace_seconds: int = 59,
) -> int | None:
    return calculate_worked_time(
        assignments,
        scheduled_entry=scheduled_entry,
        entry_grace_seconds=entry_grace_seconds,
    ).worked_minutes


def evaluate_business(
    classification: ClassificationResult,
    expected_events: Sequence[ExpectedEvent],
    *,
    cutoff_time: datetime | None = None,
    policy: BusinessPolicy | None = None,
) -> BusinessEvaluation:
    policy = policy or BusinessPolicy()
    expected = {event.key: event.expected_at for event in expected_events}
    assignments = {
        key: classification.assignments[key].checked_at
        if classification.assignments[key] is not None
        else None
        for key in EVENT_KEYS
    }
    entry = assignments["entry"]
    lunch_out = assignments["lunch_out"]
    lunch_return = assignments["lunch_return"]
    exit_time = assignments["exit"]
    finalized = cutoff_time is None or cutoff_time >= expected["exit"] or exit_time is not None
    partial_meal_ambiguous = "Comida parcial ambigua" in classification.technical_flags
    declared_state_mode = "Modo declarativo por estado" in classification.technical_flags

    incidents: list[str] = []
    if entry is None and (finalized or any((lunch_out, lunch_return, exit_time))):
        incidents.append("Sin entrada")
    if (
        not partial_meal_ambiguous
        and lunch_out is None
        and (finalized or lunch_return is not None or exit_time is not None)
    ):
        incidents.append("Sin inicio de comida")
    if not partial_meal_ambiguous and lunch_return is None and (finalized or exit_time is not None):
        incidents.append("Sin regreso de comida")
    if exit_time is None and finalized:
        incidents.append("Sin salida final")

    assigned_times = [assignments[key] for key in EVENT_KEYS if assignments[key] is not None]
    sequence_invalid = any(left >= right for left, right in zip(assigned_times, assigned_times[1:]))
    if lunch_return is not None and lunch_out is None and (entry is not None or exit_time is not None):
        sequence_invalid = True
    if any("Secuencia inv" in flag for flag in classification.technical_flags):
        sequence_invalid = True
    if sequence_invalid:
        incidents.append("Secuencia inválida")

    lunch_duration_seconds: int | None = None
    lunch_minutes: int | None = None
    if lunch_out is not None and lunch_return is not None and lunch_return > lunch_out:
        lunch_duration_seconds = int((lunch_return - lunch_out).total_seconds())
        lunch_minutes = lunch_duration_seconds // 60
        if lunch_duration_seconds > policy.maximum_lunch_seconds:
            excess_minutes = math.ceil((lunch_duration_seconds - policy.maximum_lunch_seconds) / 60)
            incidents.append(f"Exceso de comida (+{excess_minutes} min)")

    early_leave_minutes: int | None = None
    if exit_time is not None and exit_time < expected["exit"]:
        early_leave_minutes = int((expected["exit"] - exit_time).total_seconds() // 60)
        if early_leave_minutes > 0:
            incidents.append(f"Salida anticipada ({early_leave_minutes} min)")

    if partial_meal_ambiguous:
        incidents.append("Registro incompleto")

    ambiguous = (
        bool(classification.ambiguous_punches)
        or "Posible entrada tardía" in classification.technical_flags
    ) and not partial_meal_ambiguous
    if ambiguous:
        incidents.append("Registro ambiguo")
    non_ambiguous_unused = [
        punch
        for punch in classification.unused_punches
        if punch not in classification.ambiguous_punches
        and all(punch != duplicate.duplicate for duplicate in classification.duplicate_punches)
    ]
    if non_ambiguous_unused and not declared_state_mode:
        incidents.append("Checada no reconocida")

    unresolved_visible_punches = [
        punch
        for punch in classification.unused_punches
        if all(punch != duplicate.duplicate for duplicate in classification.duplicate_punches)
    ]
    detail_annotations: list[str] = []
    if unresolved_visible_punches and (ambiguous or bool(non_ambiguous_unused)):
        registered_times = ", ".join(
            punch.checked_at.strftime("%H:%M:%S")
            for punch in sorted(unresolved_visible_punches, key=lambda item: item.checked_at)
        )
        detail_annotations.append(f"Checada registrada sin clasificar ({registered_times})")

    tardy_minutes = 0
    tardy_detail = ""
    if entry is not None and entry > expected["entry"]:
        tardy_minutes = int((entry - expected["entry"]).total_seconds() // 60)
        if tardy_minutes > 0:
            label = "Retardo grave" if tardy_minutes >= policy.severe_tardy_minutes else "Retardo"
            tardy_detail = f"{label} ({tardy_minutes} min)"

    if tardy_minutes > 0 and incidents:
        status = "Retardo + incidencia"
    elif tardy_minutes > 0:
        status = "Retardo"
    elif ambiguous:
        status = "Ambiguo"
    elif incidents:
        status = "Incidencia"
    else:
        status = "Puntual"

    worked_time = calculate_worked_time(
        assignments,
        scheduled_entry=expected["entry"],
        entry_grace_seconds=policy.entry_grace_seconds,
    )

    detail_items = ([tardy_detail] if tardy_detail else []) + incidents + detail_annotations
    return BusinessEvaluation(
        assignments=assignments,
        operational_incidents=incidents,
        status=status,
        detail=" | ".join(detail_items),
        tardy_minutes=tardy_minutes,
        lunch_duration_seconds=lunch_duration_seconds,
        lunch_minutes=lunch_minutes,
        early_leave_minutes=early_leave_minutes,
        worked_minutes=worked_time.worked_minutes,
    )
