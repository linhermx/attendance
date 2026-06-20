from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Mapping


@dataclass(frozen=True)
class WorkedTimeResult:
    entry_real: datetime | None
    entry_computable: datetime | None
    exit_real: datetime | None
    lunch_seconds: int | None
    worked_minutes: int | None
    calculable: bool
    non_calculation_reason: str | None


def compute_entry_for_worked_time(
    entry_real: datetime,
    scheduled_entry: datetime,
    grace_seconds: int = 59,
) -> datetime:
    if entry_real <= scheduled_entry + timedelta(seconds=grace_seconds):
        return scheduled_entry
    return entry_real


def calculate_worked_time(
    assignments: Mapping[str, datetime | None],
    *,
    scheduled_entry: datetime | None,
    entry_grace_seconds: int = 59,
) -> WorkedTimeResult:
    entry_real = assignments.get("entry")
    lunch_out = assignments.get("lunch_out")
    lunch_return = assignments.get("lunch_return")
    exit_real = assignments.get("exit")

    if entry_real is None or exit_real is None:
        return WorkedTimeResult(
            entry_real=entry_real,
            entry_computable=None,
            exit_real=exit_real,
            lunch_seconds=None,
            worked_minutes=None,
            calculable=False,
            non_calculation_reason="missing_entry_or_exit",
        )

    if scheduled_entry is None:
        entry_computable = entry_real
    else:
        entry_computable = compute_entry_for_worked_time(
            entry_real,
            scheduled_entry,
            grace_seconds=entry_grace_seconds,
        )

    if not entry_computable < exit_real:
        return WorkedTimeResult(
            entry_real=entry_real,
            entry_computable=entry_computable,
            exit_real=exit_real,
            lunch_seconds=None,
            worked_minutes=None,
            calculable=False,
            non_calculation_reason="invalid_entry_exit_order",
        )

    if (lunch_out is None) != (lunch_return is None):
        return WorkedTimeResult(
            entry_real=entry_real,
            entry_computable=entry_computable,
            exit_real=exit_real,
            lunch_seconds=None,
            worked_minutes=None,
            calculable=False,
            non_calculation_reason="incomplete_lunch",
        )

    if lunch_out is None and lunch_return is None:
        worked_seconds = (exit_real - entry_computable).total_seconds()
        return WorkedTimeResult(
            entry_real=entry_real,
            entry_computable=entry_computable,
            exit_real=exit_real,
            lunch_seconds=None,
            worked_minutes=max(0, int(worked_seconds // 60)),
            calculable=True,
            non_calculation_reason=None,
        )

    assert lunch_out is not None and lunch_return is not None
    if not entry_computable < lunch_out < lunch_return < exit_real:
        return WorkedTimeResult(
            entry_real=entry_real,
            entry_computable=entry_computable,
            exit_real=exit_real,
            lunch_seconds=int((lunch_return - lunch_out).total_seconds()),
            worked_minutes=None,
            calculable=False,
            non_calculation_reason="invalid_lunch_sequence",
        )

    lunch_seconds = int((lunch_return - lunch_out).total_seconds())
    worked_seconds = (exit_real - entry_computable).total_seconds() - lunch_seconds
    return WorkedTimeResult(
        entry_real=entry_real,
        entry_computable=entry_computable,
        exit_real=exit_real,
        lunch_seconds=lunch_seconds,
        worked_minutes=max(0, int(worked_seconds // 60)),
        calculable=True,
        non_calculation_reason=None,
    )

