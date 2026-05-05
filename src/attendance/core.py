from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
import math
import re
from typing import Callable
import unicodedata

import pandas as pd


SHEET_NAME = "data"
REPORT_NAME = "reporte_asistencia.xlsx"
OVERTIME_REPORT_NAME = "reporte_horas_extra.xlsx"
LOG_NAME = "run_log.txt"
DUPLICATE_WINDOW_SECONDS = 3 * 60
QUICK_VIEW_BLOCK_SIZE = 4
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
    entry_time: time
    lunch_out_time: time
    lunch_return_time: time
    exit_time: time
    workday_minutes: int
    lunch_min_minutes: int
    lunch_max_minutes: int


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
    overtime_employee_count: int
    total_overtime_hours: int
    summary_frame: pd.DataFrame
    quick_view_frame: pd.DataFrame
    absence_frame: pd.DataFrame
    tardy_frame: pd.DataFrame
    incident_frame: pd.DataFrame
    daily_frame: pd.DataFrame
    overtime_frame: pd.DataFrame
    report_file: Path
    overtime_report_file: Path | None
    log_file: Path
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


def minutes_floor(delta_seconds: float) -> int:
    return max(0, int(delta_seconds // 60))


def load_table(path: str | Path, sheet_name: str = SHEET_NAME) -> tuple[pd.DataFrame, str]:
    path = Path(path)
    excel = pd.ExcelFile(path)
    selected_sheet = sheet_name if sheet_name in excel.sheet_names else excel.sheet_names[0]
    frame = pd.read_excel(path, sheet_name=selected_sheet)
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
            entry_time=time(8, 0),
            lunch_out_time=time(12, 0),
            lunch_return_time=time(12, 30),
            exit_time=time(14, 0),
            workday_minutes=6 * 60,
            lunch_min_minutes=30,
            lunch_max_minutes=30,
        )
    if weekday == 6:
        issues.append(
            RunIssue(
                "warning",
                "La fecha analizada cae en domingo. Se aplicó el horario de sábado para no bloquear el reporte.",
            )
        )
        return WorkSchedule(
            label="Domingo (horario sábado)",
            entry_time=time(8, 0),
            lunch_out_time=time(12, 0),
            lunch_return_time=time(12, 30),
            exit_time=time(14, 0),
            workday_minutes=6 * 60,
            lunch_min_minutes=30,
            lunch_max_minutes=30,
        )
    return WorkSchedule(
            label="Lunes a viernes",
        entry_time=time(8, 0),
        lunch_out_time=time(12, 0),
        lunch_return_time=time(12, 45),
        exit_time=time(17, 0),
        workday_minutes=9 * 60,
        lunch_min_minutes=30,
        lunch_max_minutes=45,
    )


def combine_day_time(work_date: date, value: time) -> datetime:
    return datetime.combine(work_date, value)


def calculate_payable_overtime_hours(
    entry_real: datetime,
    exit_real: datetime,
    work_date: date,
    schedule: WorkSchedule,
) -> int:
    scheduled_entry_dt = combine_day_time(work_date, schedule.entry_time)
    payable_start_dt = max(scheduled_entry_dt, entry_real) + timedelta(minutes=schedule.workday_minutes)
    if exit_real <= payable_start_dt:
        return 0
    return int((exit_real - payable_start_dt).total_seconds() // 3600)


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


def dedupe_events(events: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    kept_rows: list[dict[str, object]] = []
    removed_counts: dict[str, int] = {}
    for user_id, group in events.groupby("id_usuario", sort=False):
        group = group.sort_values("tiempo")
        last_kept_time: datetime | None = None
        for _, row in group.iterrows():
            current_time = row["tiempo"]
            if last_kept_time is None:
                kept_rows.append(row.to_dict())
                last_kept_time = current_time
                continue
            if (current_time - last_kept_time).total_seconds() <= DUPLICATE_WINDOW_SECONDS:
                removed_counts[user_id] = removed_counts.get(user_id, 0) + 1
                continue
            kept_rows.append(row.to_dict())
            last_kept_time = current_time
    deduped = pd.DataFrame(kept_rows)
    if deduped.empty:
        deduped = events.head(0).copy()
    return deduped.reset_index(drop=True), removed_counts


def find_lunch_out_index(events: pd.DataFrame, work_date: date, schedule: WorkSchedule) -> int | None:
    if len(events) < 2:
        return None
    start_window = combine_day_time(work_date, schedule.lunch_out_time) - timedelta(minutes=60)
    end_window = combine_day_time(work_date, schedule.lunch_return_time) + timedelta(minutes=45)
    candidates = events.iloc[1:].copy()
    explicit = candidates[
        candidates["estado_normalizado"].str.contains("salida a descanso", regex=False)
        & candidates["tiempo"].between(start_window, end_window)
    ]
    if not explicit.empty:
        return int(explicit.index[0])
    generic = candidates[candidates["tiempo"].between(start_window, end_window)]
    if not generic.empty:
        return int(generic.index[0])
    return None


def find_lunch_return_index(
    events: pd.DataFrame,
    lunch_out_index: int | None,
    work_date: date,
    schedule: WorkSchedule,
) -> int | None:
    if lunch_out_index is None:
        return None
    start_row = events.loc[lunch_out_index, "tiempo"] + timedelta(minutes=1)
    latest_return = min(
        combine_day_time(work_date, schedule.lunch_return_time) + timedelta(minutes=90),
        combine_day_time(work_date, schedule.exit_time) - timedelta(minutes=30),
    )
    tail = events.loc[lunch_out_index + 1 :]
    if tail.empty:
        return None
    explicit = tail[
        tail["estado_normalizado"].str.contains("regreso descanso", regex=False)
        & tail["tiempo"].between(start_row, latest_return)
    ]
    if not explicit.empty:
        return int(explicit.index[0])
    generic = tail[tail["tiempo"].between(start_row, latest_return)]
    if not generic.empty:
        return int(generic.index[0])
    return None


def find_exit_index(
    events: pd.DataFrame,
    start_index: int | None,
    work_date: date,
    schedule: WorkSchedule,
) -> int | None:
    if start_index is None:
        return None
    tail = events.loc[start_index + 1 :]
    if tail.empty:
        return None
    earliest_exit = combine_day_time(work_date, schedule.exit_time) - timedelta(
        minutes=45 if schedule.label.startswith("Sabado") or schedule.label.startswith("Domingo") else 90
    )
    strong_candidates = tail[tail["tiempo"] >= earliest_exit]
    if not strong_candidates.empty:
        return int(strong_candidates.index[-1])
    if start_index == 0 and len(tail) == 1:
        return int(tail.index[-1])
    return None


def compose_status(absent: bool, tardy_minutes: int, observations: list[str]) -> str:
    if absent:
        return "Falta"
    if tardy_minutes > 0 and observations:
        return "Retardo + incidencia"
    if tardy_minutes > 0:
        return "Retardo"
    if observations:
        return "Incidencia"
    return "Puntual"


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
            {"Indicador": "Personal con horas extra", "Valor": result.overtime_employee_count},
            {"Indicador": "Horas extra totales", "Valor": result.total_overtime_hours},
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
    if not detail:
        return ""
    replacements = {
        "Sin checadas en el día.": "Sin checadas",
        "Sin checada de salida a comida.": "Sin salida a comida",
        "Sin checada de regreso de comida.": "Sin regreso de comida",
        "Sin checada de salida final.": "Sin salida final",
        "Tiempo de comida mayor al máximo permitido": "Comida mayor al máximo",
        "Salida antes del horario programado": "Salida anticipada",
    }
    text = detail
    for original, replacement in replacements.items():
        text = text.replace(original, replacement)
    return text


def sort_by_id(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "ID" not in frame.columns:
        return frame.reset_index(drop=True)
    sorted_frame = frame.copy()
    sorted_frame["_id_sort"] = pd.to_numeric(sorted_frame["ID"], errors="coerce")
    sorted_frame["_id_sort"] = sorted_frame["_id_sort"].fillna(10**9)
    sorted_frame = sorted_frame.sort_values(["_id_sort", "ID"]).drop(columns="_id_sort")
    return sorted_frame.reset_index(drop=True)


def build_daily_frame(full_frame: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "ID",
        "Nombre",
        "Entrada",
        "Inicio comida",
        "Fin comida",
        "Salida",
        "Retardo min",
        "Comida min",
        "Salida anticipada min",
        "Horas extra",
        "Estatus",
        "Detalle",
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
    frame = daily_frame[
        (daily_frame["Detalle"] != "") & (daily_frame["Estatus"] != "Falta")
    ][["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]].copy()
    return sort_by_id(frame)


def build_quick_view_source(
    absence_frame: pd.DataFrame,
    tardy_frame: pd.DataFrame,
    incident_frame: pd.DataFrame,
    daily_frame: pd.DataFrame,
) -> pd.DataFrame:
    source = daily_frame.copy()

    base_columns = ["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Horas extra", "Estatus", "Detalle"]
    for column in base_columns:
        if column not in source.columns:
            source[column] = ""
    if "Horas extra" in source.columns:
        source["Horas extra"] = source["Horas extra"].apply(
            lambda value: "" if pd.isna(value) or int(value) <= 0 else f"{int(value)} h"
        )
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
            ("Horas extra", "Horas extra"),
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


def build_overtime_report_frame(full_frame: pd.DataFrame) -> pd.DataFrame:
    frame = full_frame[full_frame["Horas extra"] > 0].copy()
    if frame.empty:
        return pd.DataFrame(columns=["ID", "Nombre", "Salida", "Horas extra"])
    return frame[["ID", "Nombre", "Salida", "Horas extra"]].reset_index(drop=True)


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
    worksheet.write(2, 0, "Horas extra se generan en un reporte separado y solo cuentan tras cumplir la jornada.", subtitle_format)

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

    source_frame = result.quick_view_frame
    worksheet.merge_range(0, 0, 0, QUICK_VIEW_BLOCK_SIZE, "Vista rápida para compartir", title_format)
    worksheet.write(1, 0, f"Fecha: {result.work_date_label}")

    placeholder_fields = {"Entrada", "Inicio comida", "Fin comida", "Salida", "Horas extra"}
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
            ("HORAS EXTRA", "Horas extra"),
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
                    if "Falta" in text:
                        fmt = status_falta
                    elif "Retardo" in text:
                        fmt = status_retardo
                    elif "Incidencia" in text:
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
        write_standard_sheet(writer, "Detalle diario", result.daily_frame)


def write_overtime_report(path: Path, result: RunResult) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        summary = pd.DataFrame(
            [
                {"Indicador": "Fecha analizada", "Valor": result.work_date_label},
                {"Indicador": "Horario aplicado", "Valor": result.schedule_label},
                {"Indicador": "Personal con horas extra", "Valor": result.overtime_employee_count},
                {"Indicador": "Horas extra totales", "Valor": result.total_overtime_hours},
            ]
        )
        write_standard_sheet(writer, "Resumen", summary)
        write_standard_sheet(writer, "Horas extra", result.overtime_frame)


def write_log(path: Path, result: RunResult) -> None:
    lines = [
        f"Fecha analizada: {result.work_date_label}",
        f"Horario aplicado: {result.schedule_label}",
        f"Total empleados: {result.total_employees}",
        f"Asistencias: {result.attendance_count}",
        f"Retardos: {result.tardy_count}",
        f"Faltas: {result.absence_count}",
        f"Personal con incidencias: {result.incident_employee_count}",
        f"Personal con horas extra: {result.overtime_employee_count}",
        f"Horas extra totales: {result.total_overtime_hours}",
        (
            f"Reporte horas extra: {result.overtime_report_file.name}"
            if result.overtime_report_file
            else "Reporte horas extra: No generado"
        ),
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
    if result.overtime_report_file is not None:
        result.overtime_report_file = write_with_fallback(
            result.overtime_report_file,
            lambda target: write_overtime_report(target, result),
            result.issues,
            "El reporte de horas extra",
        )

    refresh_summary(result)
    result.report_file = write_with_fallback(
        result.report_file,
        lambda target: write_main_report(target, result),
        result.issues,
        "El reporte principal",
        before_retry=lambda: refresh_summary(result),
    )
    result.log_file = write_with_fallback(
        result.log_file,
        lambda target: write_log(target, result),
        result.issues,
        "El archivo de log",
    )
    return result


def empty_result(
    report_file: Path,
    log_file: Path,
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
        overtime_employee_count=0,
        total_overtime_hours=0,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        quick_view_frame=pd.DataFrame(columns=["Campo", "Empleado 1"]),
        absence_frame=pd.DataFrame(columns=["ID", "Nombre", "Detalle"]),
        tardy_frame=pd.DataFrame(columns=["ID", "Nombre", "Entrada", "Retardo min", "Detalle"]),
        incident_frame=pd.DataFrame(columns=["ID", "Nombre", "Entrada", "Inicio comida", "Fin comida", "Salida", "Estatus", "Detalle"]),
        daily_frame=pd.DataFrame(),
        overtime_frame=pd.DataFrame(),
        report_file=report_file,
        overtime_report_file=None,
        log_file=log_file,
        issues=issues,
    )
    result.quick_view_frame = build_quick_view_frame(result.daily_frame)
    return persist_result_outputs(result)


def calculate_attendance(
    personal_path: str | Path,
    events_path: str | Path,
    output_dir: str | Path,
    overwrite: bool = True,
) -> RunResult:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_file, log_file = build_output_paths(output_dir, overwrite)
    overtime_report_file = build_output_path(output_dir, OVERTIME_REPORT_NAME, overwrite)

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

    issues.extend(validate_required_columns(personal_std, REQUIRED_PERSONAL_COLUMNS, "Personal"))
    issues.extend(validate_required_columns(events_std, REQUIRED_EVENT_COLUMNS, "Eventos"))
    if any(issue.level == "error" for issue in issues):
        return empty_result(report_file, log_file, issues)

    personal, excluded_invalid_ids = prepare_personal_frame(personal_std, issues)
    events = prepare_events_frame(events_std, issues)
    work_date = select_work_date(events, issues)
    if work_date is None:
        issues.append(RunIssue("error", "No se encontraron eventos válidos para analizar."))
        return empty_result(report_file, log_file, issues)

    events = events[events["tiempo"].dt.date == work_date].copy().reset_index(drop=True)
    schedule = schedule_for_date(work_date, issues)
    deduped_events, _removed_counts = dedupe_events(events)

    known_ids = set(personal["id_usuario"].tolist())
    unknown_ids = sorted(set(deduped_events["id_usuario"].tolist()) - known_ids - excluded_invalid_ids, key=lambda value: int(value) if value.isdigit() else value)
    if unknown_ids:
        issues.append(
            RunIssue(
                "warning",
                "Se detectaron eventos de usuarios que no existen en la BBDD de personal: "
                + ", ".join(unknown_ids),
            )
        )

    employee_rows: list[dict[str, object]] = []

    entry_dt = combine_day_time(work_date, schedule.entry_time)
    exit_dt = combine_day_time(work_date, schedule.exit_time)

    for _, employee in personal.iterrows():
        employee_id = employee["id_usuario"]
        employee_name = employee["nombre_completo"]
        employee_events = deduped_events[deduped_events["id_usuario"] == employee_id].copy().sort_values("tiempo")
        observations: list[str] = []

        if employee_events.empty:
            observations.append("Sin checadas en el día.")
            status = "Falta"
            row = {
                "ID": employee_id,
                "Nombre": employee_name,
                "Entrada": "",
                "Inicio comida": "",
                "Fin comida": "",
                "Salida": "",
                "Retardo min": 0,
                "Comida min": "",
                "Salida anticipada min": "",
                "Horas extra": 0,
                "Estatus": status,
                "Detalle": summarize_detail(" | ".join(observations)),
            }
            employee_rows.append(row)
            continue

        entry_index = int(employee_events.index[0])
        lunch_out_index = find_lunch_out_index(employee_events, work_date, schedule)
        lunch_return_index = find_lunch_return_index(employee_events, lunch_out_index, work_date, schedule)
        start_for_exit = lunch_return_index
        if start_for_exit is None:
            start_for_exit = lunch_out_index
        if start_for_exit is None:
            start_for_exit = entry_index
        exit_index = find_exit_index(employee_events, start_for_exit, work_date, schedule)

        entry_real = employee_events.loc[entry_index, "tiempo"]
        lunch_out_real = employee_events.loc[lunch_out_index, "tiempo"] if lunch_out_index is not None else None
        lunch_return_real = (
            employee_events.loc[lunch_return_index, "tiempo"] if lunch_return_index is not None else None
        )
        exit_real = employee_events.loc[exit_index, "tiempo"] if exit_index is not None else None

        tardy_minutes = 0
        if entry_real > entry_dt:
            tardy_minutes = minutes_floor((entry_real - entry_dt).total_seconds())

        lunch_minutes: int | str = ""
        if lunch_out_real is None:
            observations.append("Sin checada de salida a comida.")
        if lunch_out_real is not None and lunch_return_real is None:
            observations.append("Sin checada de regreso de comida.")

        if lunch_out_real is not None and lunch_return_real is not None:
            lunch_minutes = minutes_floor((lunch_return_real - lunch_out_real).total_seconds())
            if lunch_minutes < schedule.lunch_min_minutes:
                pass
            if lunch_minutes > schedule.lunch_max_minutes:
                observations.append(
                    f"Tiempo de comida mayor al máximo permitido ({lunch_minutes} min)."
                )

        early_leave_minutes: int | str = ""
        overtime_hours = 0
        if exit_real is None:
            observations.append("Sin checada de salida final.")
        else:
            if exit_real < exit_dt:
                early_leave_minutes = minutes_floor((exit_dt - exit_real).total_seconds())
                if early_leave_minutes > 0:
                    observations.append(f"Salida antes del horario programado ({early_leave_minutes} min).")
            if exit_real > exit_dt:
                overtime_hours = calculate_payable_overtime_hours(entry_real, exit_real, work_date, schedule)

        status = compose_status(
            absent=False,
            tardy_minutes=tardy_minutes,
            observations=observations,
        )

        employee_rows.append(
            {
                "ID": employee_id,
                "Nombre": employee_name,
                "Entrada": display_time(entry_real),
                "Inicio comida": display_time(lunch_out_real),
                "Fin comida": display_time(lunch_return_real),
                "Salida": display_time(exit_real),
                "Retardo min": tardy_minutes,
                "Comida min": lunch_minutes,
                "Salida anticipada min": early_leave_minutes,
                "Horas extra": overtime_hours,
                "Estatus": status,
                "Detalle": summarize_detail(" | ".join(observations)),
            }
        )

    full_employee_frame = pd.DataFrame(employee_rows)
    daily_frame = build_daily_frame(full_employee_frame)
    absence_frame = build_absence_frame(daily_frame)
    tardy_frame = build_tardy_frame(daily_frame)
    incident_frame = build_incident_frame(daily_frame)
    quick_view_source = build_quick_view_source(absence_frame, tardy_frame, incident_frame, daily_frame)
    quick_view_frame = build_quick_view_frame(quick_view_source)
    overtime_frame = build_overtime_report_frame(full_employee_frame)

    total_employees = len(daily_frame)
    absence_count = int((daily_frame["Estatus"] == "Falta").sum())
    attendance_count = total_employees - absence_count
    tardy_count = int((daily_frame["Retardo min"] > 0).sum())
    incident_employee_count = int(((daily_frame["Detalle"] != "") & (daily_frame["Estatus"] != "Falta")).sum())
    overtime_employee_count = len(overtime_frame)
    total_overtime_hours = int(full_employee_frame["Horas extra"].fillna(0).sum())

    final_overtime_report_file: Path | None = None
    if not overtime_frame.empty:
        final_overtime_report_file = overtime_report_file

    result = RunResult(
        work_date=work_date,
        work_date_label=work_date.strftime("%Y-%m-%d"),
        schedule_label=schedule.label,
        total_employees=total_employees,
        attendance_count=attendance_count,
        tardy_count=tardy_count,
        absence_count=absence_count,
        incident_employee_count=incident_employee_count,
        overtime_employee_count=overtime_employee_count,
        total_overtime_hours=total_overtime_hours,
        summary_frame=pd.DataFrame(columns=["Indicador", "Valor"]),
        quick_view_frame=quick_view_frame,
        absence_frame=absence_frame,
        tardy_frame=tardy_frame,
        incident_frame=incident_frame,
        daily_frame=daily_frame,
        overtime_frame=overtime_frame,
        report_file=report_file,
        overtime_report_file=final_overtime_report_file,
        log_file=log_file,
        issues=issues,
    )
    return persist_result_outputs(result)
