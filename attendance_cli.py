from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if SRC.exists() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from attendance.core import calculate_attendance, calculate_attendance_range
from attendance.classification import load_classification_configuration


def main():
    parser = argparse.ArgumentParser(
        description="Analiza asistencias usando los archivos exportados por el checador."
    )
    parser.add_argument(
        "--mode",
        choices=("daily", "range"),
        default="daily",
        help="Modo de análisis: daily o range",
    )
    parser.add_argument("--personal", required=True, help="Ruta al archivo de personal")
    parser.add_argument("--events", help="Ruta al archivo de eventos del día")
    parser.add_argument("--range-events", help="Ruta al archivo de eventos por rango")
    parser.add_argument("--outdir", default="salida", help="Carpeta de salida")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescribir si el reporte ya existe")
    parser.add_argument(
        "--classification-config",
        help="JSON opcional con políticas predeterminadas, por turno y por empleado",
    )
    args = parser.parse_args()

    classification_kwargs = {}
    if args.classification_config:
        classification_config = load_classification_configuration(args.classification_config)
        classification_kwargs = {
            "classification_policy": classification_config.default_policy,
            "schedule_classification_policies": classification_config.schedule_policies,
            "employee_classification_policies": classification_config.employee_policies,
        }

    if args.mode == "daily":
        if not args.events:
            parser.error("--events es obligatorio cuando --mode daily")
        result = calculate_attendance(
            personal_path=Path(args.personal),
            events_path=Path(args.events),
            output_dir=Path(args.outdir),
            overwrite=args.overwrite,
            **classification_kwargs,
        )
        print(f"Fecha analizada: {result.work_date_label}")
        print(f"Horario aplicado: {result.schedule_label}")
        print(f"Total empleados: {result.total_employees}")
        print(f"Asistencias: {result.attendance_count}")
        print(f"Retardos: {result.tardy_count}")
        print(f"Faltas: {result.absence_count}")
        print(f"Incidencias: {result.incident_employee_count}")
        print(f"Reporte: {result.report_file}")
    else:
        if not args.range_events:
            parser.error("--range-events es obligatorio cuando --mode range")
        result = calculate_attendance_range(
            personal_path=Path(args.personal),
            range_events_path=Path(args.range_events),
            output_dir=Path(args.outdir),
            overwrite=args.overwrite,
            **classification_kwargs,
        )
        print(f"Periodo analizado: {result.range_label}")
        print(f"Total empleados: {result.total_employees}")
        print(f"Días laborales: {result.workday_count}")
        print(f"Días con operación: {result.operational_day_count}")
        print(f"Días sin registros globales: {result.non_operational_day_count}")
        print(f"Asistencias: {result.attendance_count}")
        print(f"Retardos: {result.tardy_count}")
        print(f"Faltas: {result.absence_count}")
        print(f"Incidencias: {result.incident_employee_count}")
        print(f"Reporte: {result.report_file}")

    if result.issues:
        print("\nObservaciones globales:")
        for issue in result.issues:
            print(f"- [{issue.level.upper()}] {issue.message}")


if __name__ == "__main__":
    main()
