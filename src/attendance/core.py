from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
import math
import re
from typing import Callable, Mapping
import unicodedata

import pandas as pd

from .classification import (
    ClassificationPolicy,
    Punch,
    build_expected_events,
    classify_punches,
    format_classification_audit,
    resolve_policy,
)
from .business import BusinessPolicy, calculate_worked_minutes as calculate_business_worked_minutes, evaluate_business


SHEET_NAME = "data"
REPORT_NAME = "reporte_asistencia.xlsx"
LOG_NAME = "run_log.txt"
RANGE_REPORT_NAME = "reporte_asistencia_rango.xlsx"
RANGE_LOG_NAME = "run_log_rango.txt"
QUICK_VIEW_BLOCK_SIZE = 4
ClassificationPolicyInput = ClassificationPolicy | Mapping[str, object]
INCIDENT_STATUSES = {"Incidencia", "Retardo + incidencia", "Ambiguo"}
ATTENDANCE_STATUSES = {"Puntual", "Retardo", "Retardo + incidencia", "Incidencia", "Ambiguo"}
NON_WORKDAY_STATUS = "Día no laborable"
NON_WORKDAY_REVIEW_STATUS = "Revisión"
NON_OPERATIONAL_SOURCE_FOLDERS = {
    "case",
    "cases",
    "caso",
    "casos",
    "demo",
    "demos",
    "evidence",
    "evidencia",
    "evidencias",
    "example",
    "examples",
    "fixture",
    "fixtures",
    "mock",
    "mocks",
    "test",
    "test data",
    "testdata",
    "tests",
}
PERSONAL_COLUMN_ALIASES = {
    "id de usuario": "id_usuario",
    "nombre": "nombre",
    "apellido": "apellido",
    "numero de tarjeta": "numero_tarjeta",
    "no. de departamento": "departamento_id",
    "departamento": "departamento",
}

EVENT_COLUMN_ALIASES = {
    "tiempo": "tiempo",
    "id de usuario": "id_usuario",
    "nombre": "nombre",
    "apellido": "apellido",
    "numero de tarjeta": "numero_tarjeta",
    "dispositivo": "dispositivo",
    "punto del evento": "punto_evento",
    "verificacion": "verificacion",
    "estado": "estado",
    "evento": "evento",
    "notas": "notas",
}

REQUIRED_PERSONAL_COLUMNS = ["id_usuario", "nombre", "apellido"]
REQUIRED_EVENT_COLUMNS = ["tiempo", "id_usuario", "nombre", "apellido", "estado"]


@dataclass
class RunIssue:
    level: str
    message: str


@dataclass(frozen=True)
class WorkSchedule:
    label: str
    is_workday: bool
    entry_time: time | None
    lunch_out_time: time | None
    lunch_return_time: time | None
    exit_time: time | None
    lunch_max_minutes: int | None


@dataclass
class RunResult:
    work_date: date | None
    work_date_label: str
    schedule_label: str
    total_employees: int
    attendance_count: int
    tardy_count: int
    absence_count: int
    incident_employee_count: int
    summary_frame: pd.DataFrame
    quick_view_frame: pd.DataFrame
    absence_frame: pd.DataFrame
    tardy_frame: pd.DataFrame
    incident_frame: pd.DataFrame
    daily_frame: pd.DataFrame
    report_file: Path
    log_file: Path | None
    issues: list[RunIssue]


@dataclass
class RangeRunResult:
    start_date: date | None
    end_date: date | None
    range_label: str
    total_employees: int
    workday_count: int
    operational_day_count: int
    non_operational_day_count: int
    partial_cutoff: bool
    attendance_count: int
    tardy_count: int
    absence_count: int
    incident_employee_count: int
    summary_frame: pd.DataFrame
    historical_preview_frame: pd.DataFrame
    absence_frame: pd.DataFrame
    tardy_frame: pd.DataFrame
    incident_frame: pd.DataFrame
    detail_frame: pd.DataFrame
    report_file: Path
    log_file: Path | None
    issues: list[RunIssue]


def normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text)


def normalize_user_id(value: object) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return str(value).strip()


def display_time(value: datetime | None) -> str:
    if value is None or pd.isna(value):
        return ""
    return value.strftime("%H:%M:%S")


def display_duration_minutes(total_minutes: int | None) -> str:
    if total_minutes is None:
        return ""
    hours, minutes = divmod(max(0, int(total_minutes)), 60)
    return f"{hours:02d}:{minutes:02d}"


def minutes_floor(delta_seconds: float) -> int:
    return max(0, int(delta_seconds // 60))


def load_table(path: str | Path, sheet_name: str = SHEET_NAME) -> tuple[pd.DataFrame, str]:
    path = Path(path)
    with pd.ExcelFile(path) as excel:
        selected_sheet = sheet_name if sheet_name in excel.sheet_names else excel.sheet_names[0]
        frame = pd.read_excel(excel, sheet_name=selected_sheet)
    return frame, selected_sheet


def standardize_columns(frame: pd.DataFrame, aliases: dict[str, str]) -> pd.DataFrame:
    rename_map: dict[object, str] = {}
    for column in frame.columns:
        normalized = normalize_text(column)
        rename_map[column] = aliases.get(normalized, normalized)
    return frame.rename(columns=rename_map)


def validate_required_columns(frame: pd.DataFrame, required_columns: list[str], label: str) -> list[RunIssue]:
    issues: list[RunIssue] = []
    missing = [column for column in required_columns if column not in frame.columns]
    if missing:
        issues.append(
            RunIssue(
                "error",
                f"{label} no contiene las columnas requeridas: {', '.join(missing)}.",
            )
        )
    return issues


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def build_output_path(output_dir: Path, file_name: str, overwrite: bool) -> Path:
    path = output_dir / file_name
    if overwrite:
        return path
    return unique_path(path)


def build_output_paths(output_dir: Path, overwrite: bool) -> tuple[Path, Path]:
    report_path = build_output_path(output_dir, REPORT_NAME, overwrite)
    log_path = build_output_path(output_dir, LOG_NAME, overwrite)
    return report_path, log_path


def build_range_output_paths(output_dir: Path, overwrite: bool) -> tuple[Path, Path]:
    report_path = build_output_path(output_dir, RANGE_REPORT_NAME, overwrite)
    log_path = build_output_path(output_dir, RANGE_LOG_NAME, overwrite)
    return report_path, log_path


def validate_operational_report_paths(*paths: str | Path) -> None:
    blocked: list[Path] = []
    for value in paths:
        path = Path(value).resolve(strict=False)
        normalized_parts = {
            normalize_text(part).replace("_", " ").replace("-", " ")
            for part in path.parts
        }
        if normalized_parts & NON_OPERATIONAL_SOURCE_FOLDERS:
            blocked.append(path)
    if blocked:
        raise ValueError(
            "Un reporte operativo no puede usar archivos ni carpetas ubicados en "
            "testing, fixtures, mocks, demo, examples o evidence: "
            + ", ".join(str(path) for path in blocked)
        )


def detect_role_mismatch_issues(
    personal_frame: pd.DataFrame,
    events_frame: pd.DataFrame,
    events_label: str,
) -> list[RunIssue]:
    issues: list[RunIssue] = []
    missing_personal = [column for column in REQUIRED_PERSONAL_COLUMNS if column not in personal_frame.columns]
    missing_events = [column for column in REQUIRED_EVENT_COLUMNS if column not in events_frame.columns]
    personal_looks_like_events = all(column in personal_frame.columns for column in REQUIRED_EVENT_COLUMNS)
    events_looks_like_personal = all(column in events_frame.columns for column in REQUIRED_PERSONAL_COLUMNS)

    if missing_personal and personal_looks_like_events:
        issues.append(
            RunIssue(
                "error",
                "El archivo cargado en Personal parece ser un archivo de eventos del checador. Revisa que no estén invertidos.",
            )
        )
    if missing_events and events_looks_like_personal:
        issues.append(
            RunIssue(
                "error",
                f"El archivo cargado en {events_label} parece ser una BBDD de personal. Revisa que no estén invertidos.",
            )
        )
    return issues


def refresh_summary(result: RunResult) -> None:
    result.summary_frame = build_summary_frame(result)


def write_with_fallback(
    path: Path,
    writer: Callable[[Path], None],
    issues: list[RunIssue],
    label: str,
    before_retry: Callable[[], None] | None = None,
) -> Path:
    try:
        writer(path)
        return path
    except PermissionError:
        fallback_path = unique_path(path)
        issues.append(
            RunIssue(
                "warning",
                f"{label} estaba abierto o en uso. Se guardó como {fallback_path.name}.",
            )
        )
        if before_retry is not None:
            before_retry()
        writer(fallback_path)
        return fallback_path


def schedule_for_date(work_date: date, issues: list[RunIssue]) -> WorkSchedule:
    weekday = work_date.weekday()
    if weekday == 5:
        return WorkSchedule(
            label="Sábado",
            is_workday=True,
            entry_time=time(8, 0),
            lunch_out_time=time(12, 0),
            lunch_return_time=time(12, 30),
            exit_time=time(14, 0),
            lunch_max_minutes=30,
        )
    if weekday == 6:
        return WorkSchedule(
            label="Domingo - día no laborable",
            is_workday=False,
            entry_time=None,
            lunch_out_time=None,
            lunch_return_time=None,
            exit_time=None,
            lunch_max_minutes=None,
        )
    return WorkSchedule(
        label="Lunes a viernes",
        is_workday=True,
        entry_time=time(8, 0),
        lunch_out_time=time(12, 0),
        lunch_return_time=time(12, 45),
        exit_time=time(17, 0),
        lunch_max_minutes=45,
    )


def resolve_context_classification_policy(
    schedule: WorkSchedule,
    employee_id: str,
    classification_policy: ClassificationPolicyInput | None = None,
    schedule_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
    employee_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
) -> tuple[ClassificationPolicy, str]:
    policy = resolve_policy(classification_policy)
    source = "predeterminada"
    if classification_policy is not None:
        source = "general"
    if schedule_classification_policies and schedule.label in schedule_classification_policies:
        policy = resolve_policy(schedule_classification_policies[schedule.label], policy)
        source = f"turno:{schedule.label}"
    if employee_classification_policies and employee_id in employee_classification_policies:
        policy = resolve_policy(employee_classification_policies[employee_id], policy)
        source = f"empleado:{employee_id}"
    return policy, source


def format_policy_audit(policy: ClassificationPolicy, source: str) -> str:
    event_labels = {
        "entry": "entrada",
        "lunch_out": "inicio comida",
        "lunch_return": "fin comida",
        "exit": "salida",
    }
    window_summary = ", ".join(
        (
            f"{event_labels.get(event_key, event_key)} "
            f"({'ventana rígida' if event_key in {'entry', 'exit'} else 'referencia flexible'})"
            f"=-{window.before_minutes}/+{window.after_minutes}"
            f" (máx -{window.max_before_minutes}/+{window.max_after_minutes})"
        )
        for event_key, window in policy.windows.items()
    )
    return (
        f"Política={source}; parámetros[{window_summary}]; "
        f"score mínimo={policy.minimum_score:.1f}; margen ambigüedad={policy.ambiguity_margin:.1f}"
    )


def combine_day_time(work_date: date, value: time) -> datetime:
    return datetime.combine(work_date, value)


def calculate_worked_minutes(
    entry_real: datetime | None,
    lunch_out_real: datetime | None,
    lunch_return_real: datetime | None,
    exit_real: datetime | None,
) -> int | None:
    return calculate_business_worked_minutes(
        {
            "entry": entry_real,
            "lunch_out": lunch_out_real,
            "lunch_return": lunch_return_real,
            "exit": exit_real,
        }
    )


def prepare_personal_frame(frame: pd.DataFrame, issues: list[RunIssue]) -> tuple[pd.DataFrame, set[str]]:
    personal = frame.copy()
    personal["id_usuario"] = personal["id_usuario"].apply(normalize_user_id)
    personal["nombre"] = personal["nombre"].fillna("").astype(str).str.strip()
    personal["apellido"] = personal["apellido"].fillna("").astype(str).str.strip()
    if "departamento" not in personal.columns:
        personal["departamento"] = ""
    personal["departamento"] = personal["departamento"].fillna("").astype(str).str.strip()
    personal["nombre_completo"] = (
        personal["nombre"].str.cat(personal["apellido"], sep=" ").str.replace(r"\s+", " ", regex=True).str.strip()
    )

    blank_id_rows = int((personal["id_usuario"] == "").sum())
    if blank_id_rows:
        issues.append(
            RunIssue(
                "warning",
                f"Se descartaron {blank_id_rows} fila(s) vacía(s) o residual(es) en la BBDD sin ID de usuario.",
            )
        )
    personal = personal[personal["id_usuario"] != ""].copy()

    invalid_name_mask = (
        personal["nombre_completo"].eq("")
        | personal["nombre_completo"].str.contains(r"\d", regex=True, na=False)
    )
    excluded_invalid_ids = set(personal.loc[invalid_name_mask, "id_usuario"].tolist())
    invalid_name_count = len(excluded_invalid_ids)
    if invalid_name_count:
        issues.append(
            RunIssue(
                "warning",
                f"Se excluyeron {invalid_name_count} empleado(s) con nombre vacío o numérico en la BBDD: "
                + ", ".join(sorted(excluded_invalid_ids, key=lambda value: int(value) if value.isdigit() else value)),
            )
        )
        personal = personal[~invalid_name_mask].copy()

    if personal["id_usuario"].duplicated().any():
        duplicated_ids = personal.loc[personal["id_usuario"].duplicated(), "id_usuario"].tolist()
        issues.append(
            RunIssue(
                "warning",
                "Se encontraron IDs duplicados en personal. Se conserva el primer registro para: "
                + ", ".join(duplicated_ids),
            )
        )
        personal = personal.drop_duplicates(subset=["id_usuario"], keep="first")

    return personal.sort_values(["nombre_completo", "id_usuario"]).reset_index(drop=True), excluded_invalid_ids


def prepare_events_frame(frame: pd.DataFrame, issues: list[RunIssue]) -> pd.DataFrame:
    events = frame.copy()
    events["id_usuario"] = events["id_usuario"].apply(normalize_user_id)
    events["nombre"] = events["nombre"].fillna("").astype(str).str.strip()
    events["apellido"] = events["apellido"].fillna("").astype(str).str.strip()
    events["estado"] = events["estado"].fillna("").astype(str).str.strip()
    if "dispositivo" not in events.columns:
        events["dispositivo"] = ""
    if "evento" not in events.columns:
        events["evento"] = ""
    if "notas" not in events.columns:
        events["notas"] = ""
    events["dispositivo"] = events["dispositivo"].fillna("").astype(str).str.strip()
    events["evento"] = events["evento"].fillna("").astype(str).str.strip()
    events["notas"] = events["notas"].fillna("").astype(str).str.strip()
    events["estado_normalizado"] = events["estado"].apply(normalize_text)
    events["tiempo"] = pd.to_datetime(events["tiempo"], errors="coerce")
    invalid_times = events["tiempo"].isna().sum()
    if invalid_times:
        issues.append(
            RunIssue(
                "warning",
                f"Se descartaron {invalid_times} eventos con tiempo inválido.",
            )
        )
    blank_id_rows = int((events["id_usuario"] == "").sum())
    if blank_id_rows:
        issues.append(
            RunIssue(
                "warning",
                f"Se descartaron {blank_id_rows} evento(s) sin ID de usuario.",
            )
        )
    events = events[(events["id_usuario"] != "") & events["tiempo"].notna()].copy()
    events = events.sort_values(["id_usuario", "tiempo"]).reset_index(drop=True)
    return events


def validate_report_personnel_source(
    report_frame: pd.DataFrame,
    personal_frame: pd.DataFrame,
) -> None:
    if report_frame.empty:
        return
    required_report_columns = {"ID", "Nombre"}
    required_personal_columns = {"id_usuario", "nombre_completo"}
    if not required_report_columns.issubset(report_frame.columns):
        raise ValueError("El reporte no contiene ID y Nombre para validar su procedencia.")
    if not required_personal_columns.issubset(personal_frame.columns):
        raise ValueError("La fuente de personal no contiene ID y nombre completo para validación.")

    authorized = {
        (normalize_user_id(row["id_usuario"]), normalize_text(row["nombre_completo"]))
        for _, row in personal_frame.iterrows()
    }
    reported = {
        (normalize_user_id(row["ID"]), normalize_text(row["Nombre"]))
        for _, row in report_frame[["ID", "Nombre"]].drop_duplicates().iterrows()
    }
    unauthorized = sorted(reported - authorized)
    if unauthorized:
        values = ", ".join(f"{employee_id}: {name}" for employee_id, name in unauthorized)
        raise ValueError(
            "El reporte contiene personal que no proviene de la BBDD de personal cargada: "
            + values
        )


def select_work_date(events: pd.DataFrame, issues: list[RunIssue]) -> date | None:
    if events.empty:
        return None
    unique_dates = sorted(events["tiempo"].dt.date.unique().tolist())
    if len(unique_dates) > 1:
        issues.append(
            RunIssue(
                "warning",
                "El archivo de eventos contiene más de una fecha. Se usará la fecha más reciente: "
                + unique_dates[-1].strftime("%Y-%m-%d"),
            )
        )
    return unique_dates[-1]


def build_non_operational_day_rows(
    personal: pd.DataFrame,
    work_date: date,
    schedule: WorkSchedule,
    detail_message: str,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for _, employee in personal.iterrows():
        rows.append(
            {
                "Fecha": work_date.strftime("%Y-%m-%d"),
                "Día": work_date.strftime("%A"),
                "Horario": schedule.label,
                "ID": employee["id_usuario"],
                "Nombre": employee["nombre_completo"],
                "Entrada": "",
                "Inicio comida": "",
                "Fin comida": "",
                "Salida": "",
                "Horas trabajadas": "",
                "Retardo min": 0,
                "Comida min": "",
                "Salida anticipada min": "",
                "Estatus": "Sin operación",
                "Detalle": detail_message,
                "Auditoría clasificación": "",
            }
        )
    return pd.DataFrame(rows)


def build_non_workday_review_rows(
    personal: pd.DataFrame,
    events: pd.DataFrame,
    work_date: date,
    schedule: WorkSchedule,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for _, employee in personal.iterrows():
        employee_id = employee["id_usuario"]
        employee_events = events[events["id_usuario"] == employee_id].copy().sort_values("tiempo")
        checked_times = [
            checked_at.strftime("%H:%M:%S")
            for checked_at in employee_events["tiempo"].tolist()
        ]
        has_punches = bool(checked_times)
        rows.append(
            {
                "Fecha": work_date.strftime("%Y-%m-%d"),
                "Día": work_date.strftime("%A"),
                "Horario": schedule.label,
                "ID": employee_id,
                "Nombre": employee["nombre_completo"],
                "Entrada": "",
                "Inicio comida": "",
                "Fin comida": "",
                "Salida": "",
                "Horas trabajadas": "",
                "Retardo min": 0,
                "Comida min": "",
                "Salida anticipada min": "",
                "Estatus": NON_WORKDAY_REVIEW_STATUS if has_punches else NON_WORKDAY_STATUS,
                "Detalle": "Checadas en día no laborable" if has_punches else "",
                "Auditoría clasificación": (
                    "Día no laborable; no se evaluaron horarios ni incidencias. "
                    + "Checadas conservadas para revisión: "
                    + ", ".join(checked_times)
                    if has_punches
                    else ""
                ),
            }
        )
    return pd.DataFrame(rows)


def analyze_operational_day(
    personal: pd.DataFrame,
    deduped_events: pd.DataFrame,
    work_date: date,
    schedule: WorkSchedule,
    cutoff_time: datetime | None = None,
    classification_policy: ClassificationPolicyInput | None = None,
    schedule_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
    employee_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
) -> pd.DataFrame:
    if not schedule.is_workday:
        return build_non_workday_review_rows(personal, deduped_events, work_date, schedule)

    assert schedule.entry_time is not None
    assert schedule.lunch_out_time is not None
    assert schedule.lunch_return_time is not None
    assert schedule.exit_time is not None
    assert schedule.lunch_max_minutes is not None
    employee_rows: list[dict[str, object]] = []

    for _, employee in personal.iterrows():
        employee_id = employee["id_usuario"]
        employee_name = employee["nombre_completo"]
        employee_events = deduped_events[deduped_events["id_usuario"] == employee_id].copy().sort_values("tiempo")
        if employee_events.empty:
            if cutoff_time is not None:
                status = "Pendiente"
                detail = "Corte parcial"
            else:
                status = "Falta"
                detail = summarize_detail("Sin checadas en el día.")
            employee_rows.append(
                {
                    "Fecha": work_date.strftime("%Y-%m-%d"),
                    "Día": work_date.strftime("%A"),
                    "Horario": schedule.label,
                    "ID": employee_id,
                    "Nombre": employee_name,
                    "Entrada": "",
                    "Inicio comida": "",
                    "Fin comida": "",
                    "Salida": "",
                    "Horas trabajadas": "",
                    "Retardo min": 0,
                    "Comida min": "",
                    "Salida anticipada min": "",
                    "Estatus": status,
                    "Detalle": detail,
                    "Auditoría clasificación": "",
                }
            )
            continue

        employee_policy, policy_source = resolve_context_classification_policy(
            schedule,
            employee_id,
            classification_policy,
            schedule_classification_policies,
            employee_classification_policies,
        )
        expected_events = build_expected_events(
            {
                "entry": schedule.entry_time,
                "lunch_out": schedule.lunch_out_time,
                "lunch_return": schedule.lunch_return_time,
                "exit": schedule.exit_time,
            },
            work_date,
            employee_policy,
        )
        punches = [
            Punch(
                punch_id=position,
                checked_at=row["tiempo"],
                state=str(row.get("estado", row.get("estado_normalizado", ""))),
                device=str(row.get("dispositivo", "")),
            )
            for position, (_, row) in enumerate(employee_events.iterrows())
        ]
        classification = classify_punches(punches, expected_events, employee_policy)
        evaluation = evaluate_business(
            classification,
            expected_events,
            cutoff_time=cutoff_time,
            policy=BusinessPolicy(maximum_lunch_seconds=schedule.lunch_max_minutes * 60),
        )
        entry_real = evaluation.assignments["entry"]
        lunch_out_real = evaluation.assignments["lunch_out"]
        lunch_return_real = evaluation.assignments["lunch_return"]
        exit_real = evaluation.assignments["exit"]
        business_audit_parts: list[str] = []
        if lunch_out_real is not None:
            allowed_return = lunch_out_real + timedelta(minutes=schedule.lunch_max_minutes)
            business_audit_parts.append(
                "Regreso permitido="
                + allowed_return.strftime("%H:%M:%S")
                + f" ({schedule.lunch_max_minutes} min desde inicio real)"
            )
        if evaluation.lunch_duration_seconds is not None:
            business_audit_parts.append(
                f"Duración real de comida={evaluation.lunch_duration_seconds} s"
            )
        classification_audit = (
            format_policy_audit(employee_policy, policy_source)
            + " | "
            + format_classification_audit(punches, expected_events, classification)
            + (" | " + " | ".join(business_audit_parts) if business_audit_parts else "")
        )

        employee_rows.append(
            {
                "Fecha": work_date.strftime("%Y-%m-%d"),
                "Día": work_date.strftime("%A"),
                "Horario": schedule.label,
                "ID": employee_id,
                "Nombre": employee_name,
                "Entrada": display_time(entry_real),
                "Inicio comida": display_time(lunch_out_real),
                "Fin comida": display_time(lunch_return_real),
                "Salida": display_time(exit_real),
                "Horas trabajadas": display_duration_minutes(evaluation.worked_minutes),
                "Retardo min": evaluation.tardy_minutes,
                "Comida min": evaluation.lunch_minutes if evaluation.lunch_minutes is not None else "",
                "Salida anticipada min": (
                    evaluation.early_leave_minutes
                    if evaluation.early_leave_minutes is not None
                    else ""
                ),
                "Estatus": evaluation.status,
                "Detalle": evaluation.detail,
                "Auditoría clasificación": classification_audit,
            }
        )

    return pd.DataFrame(employee_rows)


def build_summary_frame(result: RunResult) -> pd.DataFrame:
    punctual_count = 0
    if not result.daily_frame.empty and "Estatus" in result.daily_frame.columns:
        punctual_count = int((result.daily_frame["Estatus"] == "Puntual").sum())
    return pd.DataFrame(
        [
            {"Indicador": "Fecha analizada", "Valor": result.work_date_label},
            {"Indicador": "Horario aplicado", "Valor": result.schedule_label},
            {"Indicador": "Total de empleados", "Valor": result.total_employees},
            {"Indicador": "Asistencias", "Valor": result.attendance_count},
            {"Indicador": "Asistencias limpias", "Valor": punctual_count},
            {"Indicador": "Retardos", "Valor": result.tardy_count},
            {"Indicador": "Faltas", "Valor": result.absence_count},
            {"Indicador": "Incidencias", "Valor": result.incident_employee_count},
            {"Indicador": "Observaciones globales", "Valor": len(result.issues)},
        ]
    )


def auto_fit_columns(worksheet, frame: pd.DataFrame) -> None:
    for column_index, column_name in enumerate(frame.columns):
        values = [str(column_name)] + frame[column_name].fillna("").astype(str).tolist()
        width = min(max(len(value) for value in values) + 2, 60)
        worksheet.set_column(column_index, column_index, width)


def apply_status_highlights(workbook, worksheet, frame: pd.DataFrame) -> None:
    if frame.empty or "Estatus" not in frame.columns:
        return

    status_column = frame.columns.get_loc("Estatus")
    first_row = 1
    last_row = len(frame)
    status_range = xl_range(first_row, status_column, last_row, status_column)

    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": "Falta",
            "format": workbook.add_format({"bg_color": "#7A2048", "font_color": "#FFFFFF"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": "Retardo",
            "format": workbook.add_format({"bg_color": "#D94841", "font_color": "#FFFFFF"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": "Incidencia",
            "format": workbook.add_format({"bg_color": "#F3B63A", "font_color": "#000000"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": NON_WORKDAY_REVIEW_STATUS,
            "format": workbook.add_format({"bg_color": "#DBEAFE", "font_color": "#1D4ED8"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": NON_WORKDAY_STATUS,
            "format": workbook.add_format({"bg_color": "#E5E7EB", "font_color": "#374151"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": "Ambiguo",
            "format": workbook.add_format({"bg_color": "#F3B63A", "font_color": "#000000"}),
        },
    )
    worksheet.conditional_format(
        status_range,
        {
            "type": "text",
            "criteria": "containing",
            "value": "Puntual",
            "format": workbook.add_format({"bg_color": "#DCEBD3", "font_color": "#1D4D2F"}),
        },
    )


def xl_range(first_row: int, first_col: int, last_row: int, last_col: int) -> str:
    to_name = lambda n: chr(ord("A") + n) if n < 26 else chr(ord("A") + (n // 26) - 1) + chr(ord("A") + (n % 26))
    return f"{to_name(first_col)}{first_row + 1}:{to_name(last_col)}{last_row + 1}"


def write_standard_sheet(writer: pd.ExcelWriter, sheet_name: str, frame: pd.DataFrame) -> None:
    frame.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    auto_fit_columns(worksheet, frame)
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))
    apply_status_highlights(writer.book, worksheet, frame)


def write_classification_audit_sheet(
    writer: pd.ExcelWriter,
    frame: pd.DataFrame,
) -> None:
    sheet_name = "Auditoría clasificación"
    write_standard_sheet(writer, sheet_name, frame)
    if frame.empty or "Auditoría clasificación" not in frame.columns:
        return
    worksheet = writer.sheets[sheet_name]
    column_widths = {"Fecha": 14, "ID": 12, "Nombre": 32, "Horario": 22}
    for column_name, width in column_widths.items():
        if column_name in frame.columns:
            column_index = frame.columns.get_loc(column_name)
            worksheet.set_column(column_index, column_index, width)
    audit_column = frame.columns.get_loc("Auditoría clasificación")
    audit_format = writer.book.add_format({"text_wrap": True, "valign": "top"})
    worksheet.set_column(audit_column, audit_column, 110, audit_format)
    for row_number, audit_text in enumerate(frame["Auditoría clasificación"].astype(str), start=1):
        estimated_lines = max(3, math.ceil(len(audit_text) / 105))
        worksheet.set_row(row_number, min(estimated_lines * 16, 240))


def write_frame_block(
    worksheet,
    start_row: int,
    title: str,
    frame: pd.DataFrame,
    section_format,
    header_format,
    cell_format,
    wrap_format,
) -> int:
    worksheet.write(start_row, 0, title, section_format)
    if frame.empty:
        worksheet.write(start_row + 1, 0, "Sin registros.", cell_format)
        return start_row + 3

    for column_index, column_name in enumerate(frame.columns):
        worksheet.write(start_row + 1, column_index, column_name, header_format)

    for row_offset, row in enumerate(frame.fillna("").itertuples(index=False), start=2):
        for column_index, value in enumerate(row):
            text = value if not isinstance(value, float) else f"{value:g}"
            fmt = wrap_format if isinstance(text, str) and len(str(text)) > 45 else cell_format
            worksheet.write(start_row + row_offset, column_index, text, fmt)

    end_row = start_row + len(frame) + 4
    return end_row


def summarize_detail(detail: str) -> str:
    return detail or ""


def sort_by_id(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "ID" not in frame.columns:
        return frame.reset_index(drop=True)
    sorted_frame = frame.copy()
    sorted_frame["_id_sort"] = pd.to_numeric(sorted_frame["ID"], errors="coerce")
    sorted_frame["_id_sort"] = sorted_frame["_id_sort"].fillna(10**9)
    sorted_frame = sorted_frame.sort_values(["_id_sort", "ID"]).drop(columns="_id_sort")
    return sorted_frame.reset_index(drop=True)


def sort_by_date_and_id(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.reset_index(drop=True)
    sorted_frame = frame.copy()
    sorted_frame["_fecha_sort"] = pd.to_datetime(sorted_frame.get("Fecha"), errors="coerce")
    if "ID" in sorted_frame.columns:
        sorted_frame["_id_sort"] = pd.to_numeric(sorted_frame["ID"], errors="coerce").fillna(10**9)
        sorted_frame = sorted_frame.sort_values(["_fecha_sort", "_id_sort", "ID"])
        sorted_frame = sorted_frame.drop(columns=["_fecha_sort", "_id_sort"])
    else:
        sorted_frame = sorted_frame.sort_values(["_fecha_sort"]).drop(columns="_fecha_sort")
    return sorted_frame.reset_index(drop=True)


def build_daily_frame(full_frame: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "ID",
        "Nombre",
        "Entrada",
        "Inicio comida",
        "Fin comida",
        "Salida",
        "Horas trabajadas",
        "Retardo min",
        "Comida min",
        "Salida anticipada min",
        "Estatus",
        "Detalle",
        "Auditoría clasificación",
    ]
    return sort_by_id(full_frame[columns].copy())


def build_absence_frame(daily_frame: pd.DataFrame) -> pd.DataFrame:
    frame = daily_frame[daily_frame["Estatus"] == "Falta"][["ID", "Nombre", "Detalle"]].copy()
    return sort_by_id(frame)


def build_tardy_frame(daily_frame: pd.DataFrame) -> pd.DataFrame:
    frame = daily_frame[daily_frame["Retardo min"].fillna(0) > 0][
        ["ID", "Nombre", "Entrada", "Retardo min", "Detalle"]
    ].copy()
    return sort_by_id(frame)


def build_incident_frame(daily_frame: pd.DataFrame) -> pd.DataFrame:
    frame = daily_frame[daily_frame["Estatus"].isin(INCIDENT_STATUSES)][
        ["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]
    ].copy()
    return sort_by_id(frame)


def build_classification_audit_frame(daily_frame: pd.DataFrame) -> pd.DataFrame:
    columns = ["ID", "Nombre", "Horario", "Auditoría clasificación"]
    if "Auditoría clasificación" not in daily_frame.columns:
        return pd.DataFrame(columns=columns)
    available = [column for column in columns if column in daily_frame.columns]
    frame = daily_frame[daily_frame["Auditoría clasificación"].astype(str) != ""][available].copy()
    return sort_by_id(frame)


def build_operational_detail_frame(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.drop(columns=["Auditoría clasificación"], errors="ignore").copy()


def build_quick_view_source(
    daily_frame: pd.DataFrame,
    authorized_personnel: pd.DataFrame,
) -> pd.DataFrame:
    validate_report_personnel_source(daily_frame, authorized_personnel)
    source = daily_frame.copy()

    base_columns = ["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Horas trabajadas", "Estatus", "Detalle"]
    for column in base_columns:
        if column not in source.columns:
            source[column] = ""
    return sort_by_id(source[base_columns].copy())


def build_quick_view_frame(source_frame: pd.DataFrame) -> pd.DataFrame:
    columns = ["Campo"] + [f"Empleado {index}" for index in range(1, QUICK_VIEW_BLOCK_SIZE + 1)]
    if source_frame.empty:
        return pd.DataFrame([{"Campo": "Detalle", "Empleado 1": "Sin novedades para mostrar."}], columns=columns)

    rows: list[dict[str, object]] = []
    for start in range(0, len(source_frame), QUICK_VIEW_BLOCK_SIZE):
        chunk = source_frame.iloc[start : start + QUICK_VIEW_BLOCK_SIZE].reset_index(drop=True)
        row_templates = [
            ("ID", "ID"),
            ("Nombre", "Nombre"),
            ("Entrada", "Entrada"),
            ("Inicio comida", "Inicio comida"),
            ("Fin comida", "Fin comida"),
            ("Salida", "Salida"),
            ("Horas trabajadas", "Horas trabajadas"),
            ("Estatus", "Estatus"),
            ("Detalle", "Detalle"),
        ]
        for label, field in row_templates:
            row = {"Campo": label}
            for offset in range(QUICK_VIEW_BLOCK_SIZE):
                column_name = f"Empleado {offset + 1}"
                if offset < len(chunk):
                    value = chunk.loc[offset, field]
                    row[column_name] = summarize_detail(value) if field == "Detalle" else value
                else:
                    row[column_name] = ""
            rows.append(row)
        rows.append({column: "" for column in columns})
    return pd.DataFrame(rows, columns=columns)


def write_summary_sheet(writer: pd.ExcelWriter, result: RunResult) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Resumen")
    writer.sheets["Resumen"] = worksheet

    title_format = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#0F3B78"})
    subtitle_format = workbook.add_format({"font_color": "#4B6482"})
    section_format = workbook.add_format({"bold": True, "bg_color": "#DCE9FF", "border": 1})
    header_format = workbook.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1})
    cell_format = workbook.add_format({"border": 1, "valign": "top"})
    wrap_format = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    metric_label = workbook.add_format({"bold": True, "bg_color": "#EDF3FB", "border": 1})
    metric_value = workbook.add_format({"border": 1})

    worksheet.write(0, 0, "Reporte diario de asistencia", title_format)
    worksheet.write(1, 0, f"Fecha: {result.work_date_label}", subtitle_format)
    worksheet.write(1, 2, f"Horario: {result.schedule_label}", subtitle_format)
    daily_rule_text = (
        "Domingo no laborable: las checadas se conservan únicamente para revisión."
        if result.schedule_label == "Domingo - día no laborable"
        else "La comida se evalúa por duración real; entrada y salida conservan horario rígido."
    )
    worksheet.write(2, 0, daily_rule_text, subtitle_format)

    for row_index, row in enumerate(result.summary_frame.itertuples(index=False), start=4):
        worksheet.write(row_index, 0, row[0], metric_label)
        worksheet.write(row_index, 1, row[1], metric_value)

    next_row = 4 + len(result.summary_frame) + 2
    next_row = write_frame_block(
        worksheet,
        next_row,
        "Retardos del día",
        result.tardy_frame,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )
    next_row = write_frame_block(
        worksheet,
        next_row,
        "Faltas del día",
        result.absence_frame,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )
    next_row = write_frame_block(
        worksheet,
        next_row,
        "Incidencias del día",
        result.incident_frame[["ID", "Nombre", "Estatus", "Detalle"]] if not result.incident_frame.empty else result.incident_frame,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )

    global_issues = pd.DataFrame(
        [{"Observación global": issue.message} for issue in result.issues]
    )
    write_frame_block(
        worksheet,
        next_row,
        "Observaciones globales",
        global_issues,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )
    worksheet.set_column(0, 0, 24)
    worksheet.set_column(1, 4, 28)
    worksheet.freeze_panes(4, 0)


def write_quick_view_sheet(writer: pd.ExcelWriter, result: RunResult) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Vista rápida")
    writer.sheets["Vista rápida"] = worksheet

    title_format = workbook.add_format({"bold": True, "font_size": 15, "font_color": "#FFFFFF", "bg_color": "#162033"})
    block_label = workbook.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1})
    name_label = workbook.add_format({"bold": True, "bg_color": "#0F3B78", "font_color": "#FFFFFF", "border": 1})
    value_format = workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    blank_format = workbook.add_format({"bg_color": "#F7F9FC"})
    status_falta = workbook.add_format({"border": 1, "bg_color": "#7A2048", "font_color": "#FFFFFF", "bold": True})
    status_retardo = workbook.add_format({"border": 1, "bg_color": "#D94841", "font_color": "#FFFFFF", "bold": True})
    status_incidencia = workbook.add_format({"border": 1, "bg_color": "#F3B63A", "font_color": "#000000", "bold": True})
    status_puntual = workbook.add_format({"border": 1, "bg_color": "#DCEBD3", "font_color": "#1D4D2F", "bold": True})
    status_neutral = workbook.add_format({"border": 1, "bg_color": "#E5E7EB", "font_color": "#374151", "bold": True})
    status_review = workbook.add_format({"border": 1, "bg_color": "#DBEAFE", "font_color": "#1D4ED8", "bold": True})

    source_frame = result.quick_view_frame
    worksheet.merge_range(0, 0, 0, QUICK_VIEW_BLOCK_SIZE, "Vista rápida para compartir", title_format)
    worksheet.write(1, 0, f"Fecha: {result.work_date_label}")

    placeholder_fields = {"Entrada", "Inicio comida", "Fin comida", "Salida", "Horas trabajadas"}
    block_start_row = 3
    block_rows = 10
    for start in range(0, len(source_frame), block_rows):
        chunk = source_frame.iloc[start : start + block_rows]
        if chunk.empty:
            continue
        if "Campo" not in chunk.columns:
            continue
        header_row = block_start_row
        values_by_row = {row["Campo"]: row for _, row in chunk.iterrows() if str(row["Campo"]).strip() != ""}
        header_values = values_by_row.get("ID")
        worksheet.write(header_row, 0, "CAMPO", block_label)
        for offset in range(QUICK_VIEW_BLOCK_SIZE):
            column = offset + 1
            employee_key = f"Empleado {offset + 1}"
            value = header_values.get(employee_key, "") if header_values is not None else ""
            worksheet.write(header_row, column, value, block_label)

        detail_rows = [
            ("NOMBRE", "Nombre"),
            ("ENTRADA", "Entrada"),
            ("INICIO COMIDA", "Inicio comida"),
            ("FIN COMIDA", "Fin comida"),
            ("SALIDA", "Salida"),
            ("HORAS TRABAJADAS", "Horas trabajadas"),
            ("ESTATUS", "Estatus"),
            ("DETALLE", "Detalle"),
        ]
        for row_offset, (label, field) in enumerate(detail_rows, start=1):
            current_row = header_row + row_offset
            label_format = name_label if label == "NOMBRE" else block_label
            worksheet.write(current_row, 0, label, label_format)
            row_values = values_by_row.get(field)
            for offset in range(QUICK_VIEW_BLOCK_SIZE):
                column = offset + 1
                employee_key = f"Empleado {offset + 1}"
                if row_values is None:
                    worksheet.write(current_row, column, "", blank_format)
                    continue
                value = row_values.get(employee_key, "")
                if pd.isna(value):
                    value = ""
                if field in placeholder_fields and str(value).strip() == "":
                    value = "--"
                if field == "Estatus":
                    text = str(value)
                    if text == NON_WORKDAY_STATUS:
                        fmt = status_neutral
                    elif text == NON_WORKDAY_REVIEW_STATUS:
                        fmt = status_review
                    elif "Falta" in text:
                        fmt = status_falta
                    elif "Retardo" in text:
                        fmt = status_retardo
                    elif "Incidencia" in text or "Ambiguo" in text:
                        fmt = status_incidencia
                    else:
                        fmt = status_puntual
                    worksheet.write(current_row, column, value, fmt)
                else:
                    worksheet.write(current_row, column, value, value_format)
        block_start_row += len(detail_rows) + 3

    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, QUICK_VIEW_BLOCK_SIZE, 28)


def write_main_report(path: Path, result: RunResult) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        write_summary_sheet(writer, result)
        write_quick_view_sheet(writer, result)
        write_standard_sheet(writer, "Faltas", result.absence_frame)
        write_standard_sheet(writer, "Retardos", result.tardy_frame)
        write_standard_sheet(writer, "Incidencias", result.incident_frame)
        write_standard_sheet(writer, "Detalle diario", build_operational_detail_frame(result.daily_frame))
        write_classification_audit_sheet(writer, build_classification_audit_frame(result.daily_frame))


def write_log(path: Path, result: RunResult) -> None:
    lines = [
        f"Fecha analizada: {result.work_date_label}",
        f"Horario aplicado: {result.schedule_label}",
        f"Total empleados: {result.total_employees}",
        f"Asistencias: {result.attendance_count}",
        f"Retardos: {result.tardy_count}",
        f"Faltas: {result.absence_count}",
        f"Personal con incidencias: {result.incident_employee_count}",
        "",
        "Observaciones globales:",
    ]
    if result.issues:
        for issue in result.issues:
            lines.append(f"- [{issue.level.upper()}] {issue.message}")
    else:
        lines.append("- Ninguna")
    path.write_text("\n".join(lines), encoding="utf-8")


def persist_result_outputs(result: RunResult) -> RunResult:
    refresh_summary(result)
    result.report_file = write_with_fallback(
        result.report_file,
        lambda target: write_main_report(target, result),
        result.issues,
        "El reporte principal",
        before_retry=lambda: refresh_summary(result),
    )
    return result


def empty_result(
    report_file: Path,
    issues: list[RunIssue],
) -> RunResult:
    result = RunResult(
        work_date=None,
        work_date_label="No determinada",
        schedule_label="No determinado",
        total_employees=0,
        attendance_count=0,
        tardy_count=0,
        absence_count=0,
        incident_employee_count=0,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        quick_view_frame=pd.DataFrame(columns=["Campo", "Empleado 1"]),
        absence_frame=pd.DataFrame(columns=["ID", "Nombre", "Detalle"]),
        tardy_frame=pd.DataFrame(columns=["ID", "Nombre", "Entrada", "Retardo min", "Detalle"]),
        incident_frame=pd.DataFrame(columns=["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]),
        daily_frame=pd.DataFrame(),
        report_file=report_file,
        log_file=None,
        issues=issues,
    )
    result.quick_view_frame = build_quick_view_frame(result.daily_frame)
    return persist_result_outputs(result)


def calculate_attendance(
    personal_path: str | Path,
    events_path: str | Path,
    output_dir: str | Path,
    overwrite: bool = True,
    classification_policy: ClassificationPolicyInput | None = None,
    schedule_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
    employee_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
) -> RunResult:
    validate_operational_report_paths(personal_path, events_path, output_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_file, _log_file = build_output_paths(output_dir, overwrite)

    issues: list[RunIssue] = []

    personal_raw, personal_sheet = load_table(personal_path, SHEET_NAME)
    events_raw, events_sheet = load_table(events_path, SHEET_NAME)
    if personal_sheet != SHEET_NAME:
        issues.append(
            RunIssue(
                "warning",
                f"El archivo de personal no tenía la hoja '{SHEET_NAME}'. Se usó '{personal_sheet}'.",
            )
        )
    if events_sheet != SHEET_NAME:
        issues.append(
            RunIssue(
                "warning",
                f"El archivo de eventos no tenía la hoja '{SHEET_NAME}'. Se usó '{events_sheet}'.",
            )
        )

    personal_std = standardize_columns(personal_raw, PERSONAL_COLUMN_ALIASES)
    events_std = standardize_columns(events_raw, EVENT_COLUMN_ALIASES)

    issues.extend(detect_role_mismatch_issues(personal_std, events_std, "Eventos"))
    issues.extend(validate_required_columns(personal_std, REQUIRED_PERSONAL_COLUMNS, "Personal"))
    issues.extend(validate_required_columns(events_std, REQUIRED_EVENT_COLUMNS, "Eventos"))
    if any(issue.level == "error" for issue in issues):
        return empty_result(report_file, issues)

    personal, excluded_invalid_ids = prepare_personal_frame(personal_std, issues)
    events = prepare_events_frame(events_std, issues)
    work_date = select_work_date(events, issues)
    if work_date is None:
        issues.append(RunIssue("error", "No se encontraron eventos válidos para analizar."))
        return empty_result(report_file, issues)

    events = events[events["tiempo"].dt.date == work_date].copy().reset_index(drop=True)
    schedule = schedule_for_date(work_date, issues)
    analysis_events = events.copy()
    if not schedule.is_workday and not analysis_events.empty:
        issues.append(
            RunIssue(
                "warning",
                (
                    f"Se detectaron {len(analysis_events)} checada(s) en domingo no laborable. "
                    "Se conservaron únicamente para revisión."
                ),
            )
        )

    known_ids = set(personal["id_usuario"].tolist())
    unknown_ids = sorted(set(analysis_events["id_usuario"].tolist()) - known_ids - excluded_invalid_ids, key=lambda value: int(value) if value.isdigit() else value)
    if unknown_ids:
        issues.append(
            RunIssue(
                "warning",
                "Se detectaron eventos de usuarios que no existen en la BBDD de personal: "
                + ", ".join(unknown_ids),
            )
        )

    full_employee_frame = analyze_operational_day(
        personal,
        analysis_events,
        work_date,
        schedule,
        classification_policy=classification_policy,
        schedule_classification_policies=schedule_classification_policies,
        employee_classification_policies=employee_classification_policies,
    )
    daily_frame = build_daily_frame(full_employee_frame)
    validate_report_personnel_source(daily_frame, personal)
    absence_frame = build_absence_frame(daily_frame)
    tardy_frame = build_tardy_frame(daily_frame)
    incident_frame = build_incident_frame(daily_frame)
    quick_view_source = build_quick_view_source(daily_frame, personal)
    quick_view_frame = build_quick_view_frame(quick_view_source)

    total_employees = len(daily_frame)
    absence_count = int((daily_frame["Estatus"] == "Falta").sum())
    attendance_count = int(daily_frame["Estatus"].isin(ATTENDANCE_STATUSES).sum())
    tardy_count = int((daily_frame["Retardo min"] > 0).sum())
    incident_employee_count = int(daily_frame["Estatus"].isin(INCIDENT_STATUSES).sum())

    result = RunResult(
        work_date=work_date,
        work_date_label=work_date.strftime("%Y-%m-%d"),
        schedule_label=schedule.label,
        total_employees=total_employees,
        attendance_count=attendance_count,
        tardy_count=tardy_count,
        absence_count=absence_count,
        incident_employee_count=incident_employee_count,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        quick_view_frame=quick_view_frame,
        absence_frame=absence_frame,
        tardy_frame=tardy_frame,
        incident_frame=incident_frame,
        daily_frame=daily_frame,
        report_file=report_file,
        log_file=None,
        issues=issues,
    )
    return persist_result_outputs(result)


def build_range_summary_frame(result: RangeRunResult) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Indicador": "Período analizado", "Valor": result.range_label},
            {"Indicador": "Total de empleados", "Valor": result.total_employees},
            {"Indicador": "Días laborales en rango", "Valor": result.workday_count},
            {"Indicador": "Días con operación", "Valor": result.operational_day_count},
            {"Indicador": "Días sin registros globales", "Valor": result.non_operational_day_count},
            {"Indicador": "Corte parcial detectado", "Valor": "Sí" if result.partial_cutoff else "No"},
            {"Indicador": "Asistencias", "Valor": result.attendance_count},
            {"Indicador": "Retardos", "Valor": result.tardy_count},
            {"Indicador": "Faltas", "Valor": result.absence_count},
            {"Indicador": "Incidencias", "Valor": result.incident_employee_count},
            {"Indicador": "Observaciones globales", "Valor": len(result.issues)},
        ]
    )


def build_range_alerts_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    if detail_frame.empty:
        return pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Estatus", "Detalle"])
    neutral_statuses = {"Sin operación", "Pendiente", NON_WORKDAY_STATUS}
    alert_mask = (
        (detail_frame["Retardo min"].fillna(0) > 0)
        | (detail_frame["Estatus"] == "Falta")
        | ((detail_frame["Detalle"] != "") & ~detail_frame["Estatus"].isin(neutral_statuses))
    )
    frame = detail_frame.loc[alert_mask, ["Fecha", "ID", "Nombre", "Estatus", "Detalle"]].copy()
    return sort_by_date_and_id(frame)


def build_range_absence_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    frame = detail_frame[detail_frame["Estatus"] == "Falta"][["Fecha", "ID", "Nombre", "Detalle"]].copy()
    return sort_by_date_and_id(frame)


def build_range_tardy_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    frame = detail_frame[detail_frame["Retardo min"].fillna(0) > 0][
        ["Fecha", "ID", "Nombre", "Entrada", "Retardo min", "Detalle"]
    ].copy()
    return sort_by_date_and_id(frame)


def build_range_incident_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    frame = detail_frame[detail_frame["Estatus"].isin(INCIDENT_STATUSES)][
        ["Fecha", "ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]
    ].copy()
    return sort_by_date_and_id(frame)


def build_range_classification_audit_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    columns = ["Fecha", "ID", "Nombre", "Horario", "Auditoría clasificación"]
    if "Auditoría clasificación" not in detail_frame.columns:
        return pd.DataFrame(columns=columns)
    frame = detail_frame[detail_frame["Auditoría clasificación"].astype(str) != ""][columns].copy()
    return sort_by_date_and_id(frame)


def build_range_detail_frame(full_frame: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Fecha",
        "Día",
        "Horario",
        "ID",
        "Nombre",
        "Entrada",
        "Inicio comida",
        "Fin comida",
        "Salida",
        "Horas trabajadas",
        "Retardo min",
        "Comida min",
        "Salida anticipada min",
        "Estatus",
        "Detalle",
        "Auditoría clasificación",
    ]
    return sort_by_date_and_id(full_frame[columns].copy())


def build_range_preview_frame(detail_frame: pd.DataFrame) -> pd.DataFrame:
    if detail_frame.empty:
        return pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Entrada", "Salida", "Horas trabajadas", "Estatus", "Detalle"])
    columns = ["Fecha", "ID", "Nombre", "Entrada", "Salida", "Horas trabajadas", "Estatus", "Detalle"]
    return sort_by_date_and_id(detail_frame[columns].copy())


def write_range_summary_sheet(writer: pd.ExcelWriter, result: RangeRunResult) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Resumen")
    writer.sheets["Resumen"] = worksheet

    title_format = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#0F3B78"})
    subtitle_format = workbook.add_format({"font_color": "#4B6482"})
    metric_label = workbook.add_format({"bold": True, "bg_color": "#EDF3FB", "border": 1})
    metric_value = workbook.add_format({"border": 1})
    section_format = workbook.add_format({"bold": True, "bg_color": "#DCE9FF", "border": 1})
    header_format = workbook.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1})
    cell_format = workbook.add_format({"border": 1, "valign": "top"})
    wrap_format = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})

    worksheet.write(0, 0, "Reporte por rango de asistencia", title_format)
    worksheet.write(1, 0, f"Período: {result.range_label}", subtitle_format)
    worksheet.write(
        2,
        0,
        (
            "Los domingos no cuentan como días laborales; sus checadas se conservan para revisión. "
            "Los días laborales sin registros globales no se cuentan como falta."
        ),
        subtitle_format,
    )

    for row_index, row in enumerate(result.summary_frame.itertuples(index=False), start=4):
        worksheet.write(row_index, 0, row[0], metric_label)
        worksheet.write(row_index, 1, row[1], metric_value)

    next_row = 4 + len(result.summary_frame) + 2
    next_row = write_frame_block(
        worksheet,
        next_row,
        "Incidencias del periodo",
        result.incident_frame[["Fecha", "ID", "Nombre", "Estatus", "Detalle"]] if not result.incident_frame.empty else result.incident_frame,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )
    global_issues = pd.DataFrame([{"Observación global": issue.message} for issue in result.issues])
    write_frame_block(
        worksheet,
        next_row,
        "Observaciones globales",
        global_issues,
        section_format,
        header_format,
        cell_format,
        wrap_format,
    )
    worksheet.set_column(0, 0, 28)
    worksheet.set_column(1, 4, 28)


def write_range_historical_sheet(writer: pd.ExcelWriter, result: RangeRunResult) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Vista histórica")
    writer.sheets["Vista histórica"] = worksheet

    title_format = workbook.add_format({"bold": True, "font_size": 15, "font_color": "#FFFFFF", "bg_color": "#162033"})
    block_label = workbook.add_format({"bold": True, "bg_color": "#1F2A44", "font_color": "#FFFFFF", "border": 1})
    name_label = workbook.add_format({"bold": True, "bg_color": "#0F3B78", "font_color": "#FFFFFF", "border": 1})
    date_format = workbook.add_format({"bold": True, "bg_color": "#EEF3FA", "border": 1, "align": "center", "valign": "vcenter"})
    value_format = workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    status_falta = workbook.add_format({"border": 1, "bg_color": "#7A2048", "font_color": "#FFFFFF", "bold": True})
    status_retardo = workbook.add_format({"border": 1, "bg_color": "#D94841", "font_color": "#FFFFFF", "bold": True})
    status_incidencia = workbook.add_format({"border": 1, "bg_color": "#F3B63A", "font_color": "#000000", "bold": True})
    status_puntual = workbook.add_format({"border": 1, "bg_color": "#DCEBD3", "font_color": "#1D4D2F", "bold": True})
    status_neutral = workbook.add_format({"border": 1, "bg_color": "#E5E7EB", "font_color": "#374151", "bold": True})
    status_pending = workbook.add_format({"border": 1, "bg_color": "#DBEAFE", "font_color": "#1D4ED8", "bold": True})

    detail_frame = result.detail_frame.copy()
    unique_dates = sorted(detail_frame["Fecha"].unique().tolist()) if not detail_frame.empty else []
    employee_meta = (
        sort_by_id(detail_frame[["ID", "Nombre"]].drop_duplicates())
        if not detail_frame.empty
        else pd.DataFrame(columns=["ID", "Nombre"])
    )
    last_data_column = max(2, len(employee_meta) + 1)

    worksheet.merge_range(0, 0, 0, last_data_column, "Vista histórica por rango", title_format)
    worksheet.write(1, 0, f"Período: {result.range_label}")

    lookup = {
        (row["Fecha"], str(row["ID"])): row
        for _, row in detail_frame.iterrows()
    }

    detail_rows = [
        ("ENTRADA", "Entrada"),
        ("INICIO COMIDA", "Inicio comida"),
        ("FIN COMIDA", "Fin comida"),
        ("SALIDA", "Salida"),
        ("HORAS TRABAJADAS", "Horas trabajadas"),
        ("ESTATUS", "Estatus"),
        ("DETALLE", "Detalle"),
    ]
    placeholder_fields = {"Entrada", "Inicio comida", "Fin comida", "Salida", "Horas trabajadas"}

    header_row = 3
    worksheet.write(header_row, 0, "FECHA", block_label)
    worksheet.write(header_row, 1, "CAMPO", block_label)
    for column, employee in enumerate(employee_meta.itertuples(index=False), start=2):
        worksheet.write(header_row, column, str(employee.ID), block_label)

    name_row = header_row + 1
    worksheet.write(name_row, 1, "NOMBRE", name_label)
    for column, employee in enumerate(employee_meta.itertuples(index=False), start=2):
        worksheet.write(name_row, column, employee.Nombre, value_format)

    current_row = name_row + 1
    for day_label in unique_dates:
        date_start_row = current_row
        date_end_row = current_row + len(detail_rows) - 1
        worksheet.merge_range(date_start_row, 0, date_end_row, 0, day_label, date_format)
        for row_offset, (label, field) in enumerate(detail_rows):
            row_idx = current_row + row_offset
            worksheet.write(row_idx, 1, label, block_label)
            for column, employee in enumerate(employee_meta.itertuples(index=False), start=2):
                employee_id = str(employee.ID)
                row = lookup.get((day_label, employee_id))
                value = "" if row is None else row[field]
                if pd.isna(value):
                    value = ""
                if field in placeholder_fields and str(value).strip() == "":
                    value = "--"
                if field == "Estatus":
                    text = str(value)
                    if text in {"Sin operación", NON_WORKDAY_STATUS}:
                        fmt = status_neutral
                    elif text in {"Pendiente", NON_WORKDAY_REVIEW_STATUS}:
                        fmt = status_pending
                    elif "Falta" in text:
                        fmt = status_falta
                    elif "Retardo" in text:
                        fmt = status_retardo
                    elif "Incidencia" in text or "Ambiguo" in text:
                        fmt = status_incidencia
                    else:
                        fmt = status_puntual
                    worksheet.write(row_idx, column, value, fmt)
                else:
                    worksheet.write(row_idx, column, value, value_format)
        current_row = date_end_row + 1

    worksheet.set_column(0, 0, 14)
    worksheet.set_column(1, 1, 22)
    worksheet.set_column(2, last_data_column, 28)
    worksheet.freeze_panes(5, 2)

def write_range_report(path: Path, result: RangeRunResult) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        write_range_summary_sheet(writer, result)
        write_range_historical_sheet(writer, result)
        write_standard_sheet(writer, "Faltas", result.absence_frame)
        write_standard_sheet(writer, "Retardos", result.tardy_frame)
        write_standard_sheet(writer, "Incidencias", result.incident_frame)
        write_standard_sheet(
            writer,
            "Detalle consolidado",
            build_operational_detail_frame(result.detail_frame),
        )
        write_classification_audit_sheet(writer, build_range_classification_audit_frame(result.detail_frame))


def write_range_log(path: Path, result: RangeRunResult) -> None:
    lines = [
        f"Período analizado: {result.range_label}",
        f"Total empleados: {result.total_employees}",
        f"Días laborales: {result.workday_count}",
        f"Días con operación: {result.operational_day_count}",
        f"Días sin registros globales: {result.non_operational_day_count}",
        f"Corte parcial: {'Sí' if result.partial_cutoff else 'No'}",
        f"Asistencias: {result.attendance_count}",
        f"Retardos: {result.tardy_count}",
        f"Faltas: {result.absence_count}",
        f"Incidencias: {result.incident_employee_count}",
        "",
        "Observaciones globales:",
    ]
    if result.issues:
        for issue in result.issues:
            lines.append(f"- [{issue.level.upper()}] {issue.message}")
    else:
        lines.append("- Ninguna")
    path.write_text("\n".join(lines), encoding="utf-8")


def persist_range_result_outputs(result: RangeRunResult) -> RangeRunResult:
    result.summary_frame = build_range_summary_frame(result)
    result.report_file = write_with_fallback(
        result.report_file,
        lambda target: write_range_report(target, result),
        result.issues,
        "El reporte principal por rango",
        before_retry=lambda: setattr(result, "summary_frame", build_range_summary_frame(result)),
    )
    return result


def empty_range_result(report_file: Path, issues: list[RunIssue]) -> RangeRunResult:
    result = RangeRunResult(
        start_date=None,
        end_date=None,
        range_label="No determinado",
        total_employees=0,
        workday_count=0,
        operational_day_count=0,
        non_operational_day_count=0,
        partial_cutoff=False,
        attendance_count=0,
        tardy_count=0,
        absence_count=0,
        incident_employee_count=0,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        historical_preview_frame=pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Entrada", "Salida", "Horas trabajadas", "Estatus", "Detalle"]),
        absence_frame=pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Detalle"]),
        tardy_frame=pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Entrada", "Retardo min", "Detalle"]),
        incident_frame=pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]),
        detail_frame=pd.DataFrame(columns=["Fecha", "ID", "Nombre", "Entrada", "Salida", "Estatus", "Detalle"]),
        report_file=report_file,
        log_file=None,
        issues=issues,
    )
    result.summary_frame = build_range_summary_frame(result)
    return persist_range_result_outputs(result)


def calculate_attendance_range(
    personal_path: str | Path,
    range_events_path: str | Path,
    output_dir: str | Path,
    overwrite: bool = True,
    classification_policy: ClassificationPolicyInput | None = None,
    schedule_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
    employee_classification_policies: Mapping[str, ClassificationPolicyInput] | None = None,
) -> RangeRunResult:
    validate_operational_report_paths(personal_path, range_events_path, output_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_file, _log_file = build_range_output_paths(output_dir, overwrite)
    issues: list[RunIssue] = []

    personal_raw, personal_sheet = load_table(personal_path, SHEET_NAME)
    events_raw, events_sheet = load_table(range_events_path, SHEET_NAME)
    if personal_sheet != SHEET_NAME:
        issues.append(RunIssue("warning", f"El archivo de personal no tenía la hoja '{SHEET_NAME}'. Se usó '{personal_sheet}'."))
    if events_sheet != SHEET_NAME:
        issues.append(RunIssue("warning", f"El archivo de rango no tenía la hoja '{SHEET_NAME}'. Se usó '{events_sheet}'."))

    personal_std = standardize_columns(personal_raw, PERSONAL_COLUMN_ALIASES)
    events_std = standardize_columns(events_raw, EVENT_COLUMN_ALIASES)
    issues.extend(detect_role_mismatch_issues(personal_std, events_std, "Rango"))
    issues.extend(validate_required_columns(personal_std, REQUIRED_PERSONAL_COLUMNS, "Personal"))
    issues.extend(validate_required_columns(events_std, REQUIRED_EVENT_COLUMNS, "Rango"))
    if any(issue.level == "error" for issue in issues):
        return empty_range_result(report_file, issues)

    personal, excluded_invalid_ids = prepare_personal_frame(personal_std, issues)
    events = prepare_events_frame(events_std, issues)
    if events.empty:
        issues.append(RunIssue("error", "No se encontraron eventos válidos para analizar en el rango."))
        return empty_range_result(report_file, issues)

    min_date = events["tiempo"].dt.date.min()
    max_date = events["tiempo"].dt.date.max()
    all_dates = list(pd.date_range(min_date, max_date, freq="D").date)
    work_dates = [single_date for single_date in all_dates if single_date.weekday() != 6]
    sunday_review_dates = [
        single_date
        for single_date in all_dates
        if single_date.weekday() == 6
        and bool((events["tiempo"].dt.date == single_date).any())
    ]
    analysis_dates = sorted(work_dates + sunday_review_dates)
    if sunday_review_dates:
        sunday_punch_count = int(events["tiempo"].dt.date.isin(sunday_review_dates).sum())
        issues.append(
            RunIssue(
                "warning",
                (
                    f"Se detectaron {sunday_punch_count} checada(s) en domingo no laborable. "
                    "Se conservaron únicamente para revisión."
                ),
            )
        )

    known_ids = set(personal["id_usuario"].tolist())
    unknown_ids = sorted(
        set(events["id_usuario"].tolist()) - known_ids - excluded_invalid_ids,
        key=lambda value: int(value) if str(value).isdigit() else str(value),
    )
    if unknown_ids:
        issues.append(
            RunIssue(
                "warning",
                "Se detectaron eventos de usuarios que no existen en la BBDD de personal: " + ", ".join(unknown_ids),
            )
        )

    all_day_frames: list[pd.DataFrame] = []
    operational_day_count = 0
    non_operational_day_count = 0
    partial_cutoff = False
    last_work_date = work_dates[-1] if work_dates else None

    for single_date in analysis_dates:
        schedule = schedule_for_date(single_date, issues)
        day_events = events[events["tiempo"].dt.date == single_date].copy().reset_index(drop=True)
        if not schedule.is_workday:
            all_day_frames.append(
                analyze_operational_day(
                    personal,
                    day_events,
                    single_date,
                    schedule,
                )
            )
            continue
        if day_events.empty:
            non_operational_day_count += 1
            all_day_frames.append(
                build_non_operational_day_rows(
                    personal,
                    single_date,
                    schedule,
                    "Sin registros globales / posible día no laborable",
                )
            )
            continue

        operational_day_count += 1
        cutoff_time: datetime | None = None
        day_last_event = day_events["tiempo"].max()
        assert schedule.exit_time is not None
        if single_date == last_work_date and day_last_event < combine_day_time(single_date, schedule.exit_time):
            cutoff_time = day_last_event
            partial_cutoff = True
        all_day_frames.append(
            analyze_operational_day(
                personal,
                day_events,
                single_date,
                schedule,
                cutoff_time=cutoff_time,
                classification_policy=classification_policy,
                schedule_classification_policies=schedule_classification_policies,
                employee_classification_policies=employee_classification_policies,
            )
        )

    if partial_cutoff:
        issues.append(
            RunIssue(
                "warning",
                f"Se detectó corte parcial en la última fecha del rango ({last_work_date.strftime('%Y-%m-%d')}).",
            )
        )
    if non_operational_day_count:
        issues.append(
            RunIssue(
                "warning",
                f"Se detectaron {non_operational_day_count} día(s) laborales sin registros globales. Se marcaron como sin operación.",
            )
        )

    full_range_frame = pd.concat(all_day_frames, ignore_index=True) if all_day_frames else pd.DataFrame()
    detail_frame = build_range_detail_frame(full_range_frame)
    validate_report_personnel_source(detail_frame, personal)
    preview_frame = build_range_preview_frame(detail_frame)
    absence_frame = build_range_absence_frame(detail_frame)
    tardy_frame = build_range_tardy_frame(detail_frame)
    incident_frame = build_range_incident_frame(detail_frame)

    attendance_count = int(detail_frame["Estatus"].isin(ATTENDANCE_STATUSES).sum())
    tardy_count = int((detail_frame["Retardo min"].fillna(0) > 0).sum())
    absence_count = int((detail_frame["Estatus"] == "Falta").sum())
    incident_employee_count = int(detail_frame["Estatus"].isin(INCIDENT_STATUSES).sum())

    result = RangeRunResult(
        start_date=min_date,
        end_date=max_date,
        range_label=f"{min_date.strftime('%Y-%m-%d')} a {max_date.strftime('%Y-%m-%d')}",
        total_employees=len(personal),
        workday_count=len(work_dates),
        operational_day_count=operational_day_count,
        non_operational_day_count=non_operational_day_count,
        partial_cutoff=partial_cutoff,
        attendance_count=attendance_count,
        tardy_count=tardy_count,
        absence_count=absence_count,
        incident_employee_count=incident_employee_count,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        historical_preview_frame=preview_frame,
        absence_frame=absence_frame,
        tardy_frame=tardy_frame,
        incident_frame=incident_frame,
        detail_frame=detail_frame,
        report_file=report_file,
        log_file=None,
        issues=issues,
    )
    result.summary_frame = build_range_summary_frame(result)
    return persist_range_result_outputs(result)
