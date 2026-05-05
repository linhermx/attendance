from __future__ import annotations

import argparse
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if SRC.exists() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from attendance.core import calculate_attendance


def main():
    parser = argparse.ArgumentParser(
        description="Analiza asistencias diarias usando los archivos exportados por el checador."
    )
    parser.add_argument("--personal", required=True, help="Ruta al archivo de personal")
    parser.add_argument("--events", required=True, help="Ruta al archivo de eventos del dia")
    parser.add_argument("--outdir", default="salida", help="Carpeta de salida")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescribir si el reporte ya existe")
    args = parser.parse_args()

    result = calculate_attendance(
        personal_path=Path(args.personal),
        events_path=Path(args.events),
        output_dir=Path(args.outdir),
        overwrite=args.overwrite,
    )

    print(f"Fecha analizada: {result.work_date_label}")
    print(f"Horario aplicado: {result.schedule_label}")
    print(f"Total empleados: {result.total_employees}")
    print(f"Asistencias: {result.attendance_count}")
    print(f"Retardos: {result.tardy_count}")
    print(f"Faltas: {result.absence_count}")
    print(f"Incidencias: {result.incident_employee_count}")
    print(f"Horas extra totales: {result.total_overtime_hours}")
    print(f"Reporte: {result.report_file}")
    if result.issues:
        print("\nObservaciones globales:")
        for issue in result.issues:
            print(f"- [{issue.level.upper()}] {issue.message}")


if __name__ == "__main__":
    main()
