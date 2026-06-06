from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from attendance.core import build_quick_view_source, calculate_attendance


class ReportPersonnelValidationTests(unittest.TestCase):
    def test_quick_view_rejects_personnel_not_in_loaded_source(self) -> None:
        authorized = pd.DataFrame(
            [{"id_usuario": "1", "nombre_completo": "EMPLEADO FUENTE"}]
        )
        daily = pd.DataFrame(
            [
                {"ID": "1", "Nombre": "EMPLEADO FUENTE"},
                {"ID": "99", "Nombre": "REGISTRO NO AUTORIZADO"},
            ]
        )
        with self.assertRaisesRegex(ValueError, "no proviene de la BBDD"):
            build_quick_view_source(daily, authorized)

    def test_operational_report_rejects_fixture_source_paths(self) -> None:
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            fixture_dir = root / "tests" / "fixtures"
            fixture_dir.mkdir(parents=True)
            personal = fixture_dir / "personal.xlsx"
            events = fixture_dir / "events.xlsx"
            personal.touch()
            events.touch()
            with self.assertRaisesRegex(ValueError, "reporte operativo"):
                calculate_attendance(personal, events, root / "out")


if __name__ == "__main__":
    unittest.main()
