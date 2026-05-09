from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import tkinter as tk
from ctypes import wintypes
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from attendance.core import RangeRunResult, RunIssue, RunResult, calculate_attendance, calculate_attendance_range
from attendance.version import __version__


ResultType = RunResult | RangeRunResult


class LinherAttendanceApp(tk.Tk):
    def __init__(self):
        self._set_windows_app_id()
        super().__init__()
        self.withdraw()
        self.title(f"LINHER Attendance | Control de asistencia (v{__version__})")
        self._apply_window_icon()

        self.personal_path = tk.StringVar(value="")
        self.events_path = tk.StringVar(value="")
        self.range_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value="")
        self.overwrite = tk.BooleanVar(value=False)

        self.last_dir = Path.home()
        self.result: ResultType | None = None
        self.report_file: Path | None = None
        self.overtime_report_file: Path | None = None
        self.log_file: Path | None = None

        self._apply_style()
        self._build_layout()
        self._configure_window()
        self.deiconify()

    def _set_windows_app_id(self):
        if os.name != "nt":
            return
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("LINHER.Attendance")
        except Exception:
            pass

    def _resource_path(self, relative_name: str) -> Path:
        if getattr(sys, "frozen", False):
            bundle_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
            return bundle_dir / "resources" / relative_name
        return Path(__file__).resolve().parent / "resources" / relative_name

    def _apply_window_icon(self):
        icon_path = self._resource_path("attendance.ico")
        if not icon_path.exists():
            return
        try:
            self.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def _apply_style(self):
        style = ttk.Style(self)
        for theme in ("vista", "xpnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break

        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#436057")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        style.configure("MetricValue.TLabel", font=("Segoe UI", 24, "bold"), foreground="#0D5E4A")
        style.configure("MetricCaption.TLabel", font=("Segoe UI", 10), foreground="#4A605A")
        style.configure("HeroTitle.TLabel", font=("Segoe UI", 16, "bold"), foreground="#0D5E4A")
        style.configure("HeroText.TLabel", font=("Segoe UI", 10), foreground="#36574D")

    def _configure_window(self):
        width = 1240
        height = 880
        work_area = self._get_work_area()
        if work_area:
            left, top, right, bottom = work_area
            available_width = max(980, right - left)
            available_height = max(740, bottom - top)
            width = min(width, available_width - 60)
            height = min(height, available_height - 60)
            x = left + max(0, (available_width - width) // 2)
            y = top + max(0, (available_height - height) // 5)
            self.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.geometry(f"{width}x{height}")
        self.minsize(1100, 780)

    def _get_work_area(self):
        if os.name != "nt":
            return None
        try:
            rect = wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
            return (rect.left, rect.top, rect.right, rect.bottom)
        except Exception:
            return None

    def _build_layout(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(20, 18, 20, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Centro de Control de Asistencia", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=(
                "Analiza exportaciones del checador en modo diario o por rango para generar cortes claros "
                "de asistencia, retardos, omisiones de checada, salidas anticipadas y horas extra pagables."
            ),
            style="Subtitle.TLabel",
            wraplength=1080,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            header,
            text="Flujo sugerido: 1. Elegir modo  2. Cargar archivos  3. Ejecutar análisis  4. Compartir el reporte",
            style="Subtitle.TLabel",
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        body = ttk.Frame(self, padding=(20, 8, 20, 12))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(body, padding=(0, 0, 18, 0))
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.columnconfigure(0, weight=1)

        self._build_input_panel(sidebar)
        self._build_action_panel(sidebar)

        guide_panel = ttk.LabelFrame(sidebar, text="Qué vigila el sistema", style="Section.TLabelframe", padding=12)
        guide_panel.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(
            guide_panel,
            text=(
                "Sí importa:\n"
                "- faltas\n"
                "- retardos\n"
                "- checadas faltantes\n"
                "- salida anticipada\n"
                "- comida excedida\n\n"
                "También ayuda a detectar:\n"
                "- archivos invertidos\n"
                "- corte parcial\n"
                "- días sin registros globales"
            ),
            justify="left",
            style="Subtitle.TLabel",
        ).grid(row=0, column=0, sticky="w")

        content = ttk.Frame(body)
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        result_block = ttk.LabelFrame(content, text="Resultado", style="Section.TLabelframe", padding=12)
        result_block.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        result_block.columnconfigure(0, weight=1)

        hero = tk.Frame(result_block, bg="#E9F5EF", bd=1, relief="solid")
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        hero.columnconfigure(0, weight=1)
        self.hero_title = ttk.Label(
            hero,
            text="Carga los archivos para iniciar el análisis.",
            style="HeroTitle.TLabel",
        )
        self.hero_text = ttk.Label(
            hero,
            text=(
                "Aquí verás el corte del modo activo, los totales importantes y las alertas que "
                "sí requieren seguimiento."
            ),
            style="HeroText.TLabel",
            wraplength=1080,
            justify="left",
        )
        self.hero_title.grid(row=0, column=0, sticky="w", padx=16, pady=(14, 4))
        self.hero_text.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 14))

        metrics = ttk.Frame(result_block)
        metrics.grid(row=1, column=0, sticky="ew")
        for column_index in range(5):
            metrics.columnconfigure(column_index, weight=1)

        self.metric_cards = {
            "empleados": self._create_metric_card(metrics, 0, "0", "Plantilla válida"),
            "asistencias": self._create_metric_card(metrics, 1, "0", "Presentes"),
            "retardos": self._create_metric_card(metrics, 2, "0", "Retardos"),
            "faltas": self._create_metric_card(metrics, 3, "0", "Faltas"),
            "incidencias": self._create_metric_card(metrics, 4, "0", "Alertas"),
        }

        notebook_frame = ttk.Frame(content)
        notebook_frame.grid(row=2, column=0, sticky="nsew")
        notebook_frame.columnconfigure(0, weight=1)
        notebook_frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.trees: dict[str, ttk.Treeview] = {}
        self._add_tree_tab("resumen", "Resumen")
        self._add_tree_tab("vista", "Vista rápida")
        self._add_tree_tab("faltas", "Faltas")
        self._add_tree_tab("retardos", "Retardos")
        self._add_tree_tab("incidencias", "Incidencias")
        self._add_tree_tab("detalle", "Detalle diario")
        self._set_daily_result_tabs()

        status_bar = ttk.Frame(content, padding=(0, 10, 0, 0))
        status_bar.grid(row=3, column=0, sticky="ew")
        status_bar.columnconfigure(0, weight=1)
        status_bar.columnconfigure(1, weight=1)
        self.status_left = ttk.Label(status_bar, text="Listo.")
        self.status_right = ttk.Label(status_bar, text="Aún no se ha generado ningún reporte.", anchor="e")
        self.status_left.grid(row=0, column=0, sticky="w")
        self.status_right.grid(row=0, column=1, sticky="e")

    def _build_input_panel(self, parent: ttk.Frame):
        panel = ttk.LabelFrame(parent, text="Carga de archivos", style="Section.TLabelframe", padding=12)
        panel.grid(row=0, column=0, sticky="ew")
        panel.columnconfigure(0, weight=1)

        self.input_notebook = ttk.Notebook(panel)
        self.input_notebook.grid(row=0, column=0, sticky="ew")
        self.input_notebook.bind("<<NotebookTabChanged>>", self._on_mode_tab_changed)

        daily_tab = ttk.Frame(self.input_notebook, padding=(4, 6, 4, 6))
        daily_tab.columnconfigure(1, weight=1)
        self.input_notebook.add(daily_tab, text="Diario")
        self._build_file_row(daily_tab, 0, "Personal", self.personal_path, self.pick_personal)
        self._build_file_row(daily_tab, 1, "Eventos del día", self.events_path, self.pick_events)
        self._build_file_row(daily_tab, 2, "Carpeta de salida", self.output_dir, self.pick_output_dir)
        ttk.Label(
            daily_tab,
            text=(
                "Usa la BBDD de personal y el archivo del día. Si cargas archivos invertidos, "
                "la app te avisará antes de dejarte con un corte vacío."
            ),
            foreground="#4B6482",
            wraplength=300,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        range_tab = ttk.Frame(self.input_notebook, padding=(4, 6, 4, 6))
        range_tab.columnconfigure(1, weight=1)
        self.input_notebook.add(range_tab, text="Rango")
        self._build_file_row(range_tab, 0, "Personal", self.personal_path, self.pick_personal)
        self._build_file_row(range_tab, 1, "Archivo de rango", self.range_path, self.pick_range)
        self._build_file_row(range_tab, 2, "Carpeta de salida", self.output_dir, self.pick_output_dir)
        ttk.Label(
            range_tab,
            text=(
                "El modo Rango toma la fecha mínima y máxima del archivo, excluye domingos e incluye "
                "corte parcial o días sin registros globales cuando aplique."
            ),
            foreground="#4B6482",
            wraplength=300,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _build_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(2, 10))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(12, 8))
        ttk.Button(parent, text="Seleccionar...", command=command).grid(row=row, column=2, sticky="ew")

    def _build_action_panel(self, parent: ttk.Frame):
        panel = ttk.LabelFrame(parent, text="Ejecución", style="Section.TLabelframe", padding=12)
        panel.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        panel.columnconfigure(0, weight=1)

        self.execution_label = ttk.Label(panel, text="Configuración del reporte diario")
        self.execution_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
        ttk.Checkbutton(panel, text="Sobrescribir si existe", variable=self.overwrite).grid(
            row=1, column=0, sticky="w", pady=(0, 14)
        )
        self.run_button = ttk.Button(panel, text="Analizar asistencia diaria", command=self.run_analysis)
        self.run_button.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        ttk.Button(panel, text="Restablecer", command=self.reset_form).grid(
            row=3, column=0, sticky="ew", pady=(0, 8)
        )
        self.open_output_button = ttk.Button(panel, text="Abrir carpeta...", command=self.open_output_dir)
        self.open_output_button.grid(row=4, column=0, sticky="ew")
        self.open_output_button.state(["disabled"])
        ttk.Label(
            panel,
            text=(
                "El reporte resaltará faltas, retardos, omisiones de checada, comida excedida y salida anticipada. "
                "Las horas extra se generan por separado y solo cuentan después de cumplir la jornada."
            ),
            foreground="#4B6482",
            wraplength=300,
            justify="left",
        ).grid(row=5, column=0, sticky="w", pady=(12, 0))

    def _create_metric_card(self, parent: ttk.Frame, column: int, value: str, caption: str):
        card = ttk.Frame(parent, padding=14, relief="solid")
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
        card.columnconfigure(0, weight=1)
        value_label = ttk.Label(card, text=value, style="MetricValue.TLabel")
        caption_label = ttk.Label(card, text=caption, style="MetricCaption.TLabel", wraplength=180, justify="left")
        value_label.grid(row=0, column=0, sticky="w", pady=(0, 10))
        caption_label.grid(row=1, column=0, sticky="w")
        return {"value": value_label, "caption": caption_label}

    def _add_tree_tab(self, key: str, label: str):
        frame = ttk.Frame(self.notebook, padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, show="headings", height=10)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.notebook.add(frame, text=label)
        self.trees[key] = tree

    def _set_daily_result_tabs(self):
        self.notebook.tab(0, text="Resumen")
        self.notebook.tab(1, text="Vista rápida")
        self.notebook.tab(2, text="Faltas")
        self.notebook.tab(3, text="Retardos")
        self.notebook.tab(4, text="Incidencias")
        self.notebook.tab(5, text="Detalle diario")

    def _set_range_result_tabs(self):
        self.notebook.tab(0, text="Resumen")
        self.notebook.tab(1, text="Vista histórica")
        self.notebook.tab(2, text="Faltas")
        self.notebook.tab(3, text="Retardos")
        self.notebook.tab(4, text="Incidencias")
        self.notebook.tab(5, text="Detalle consolidado")

    def _get_current_mode(self) -> str:
        current_index = self.input_notebook.index(self.input_notebook.select())
        return "daily" if current_index == 0 else "range"

    def _on_mode_tab_changed(self, _event=None):
        mode = self._get_current_mode()
        if mode == "daily":
            self.execution_label.configure(text="Configuración del reporte diario")
            self.run_button.configure(text="Analizar asistencia diaria")
            self._set_daily_result_tabs()
        else:
            self.execution_label.configure(text="Configuración del reporte por rango")
            self.run_button.configure(text="Analizar rango")
            self._set_range_result_tabs()

    def pick_personal(self):
        path = self._pick_file("Selecciona el archivo de personal")
        if path:
            self.personal_path.set(path)

    def pick_events(self):
        path = self._pick_file("Selecciona el archivo de eventos del día")
        if path:
            self.events_path.set(path)

    def pick_range(self):
        path = self._pick_file("Selecciona el archivo de rango")
        if path:
            self.range_path.set(path)

    def pick_output_dir(self):
        try:
            path = filedialog.askdirectory(title="Selecciona la carpeta de salida", initialdir=str(self.last_dir))
        except KeyboardInterrupt:
            return
        if path:
            self.last_dir = Path(path)
            self.output_dir.set(path)

    def _pick_file(self, title: str) -> str | None:
        try:
            path = filedialog.askopenfilename(
                title=title,
                initialdir=str(self.last_dir),
                filetypes=[("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")],
            )
        except KeyboardInterrupt:
            return None
        if not path:
            return None
        self.last_dir = Path(path).resolve().parent
        return path

    def reset_form(self):
        self.personal_path.set("")
        self.events_path.set("")
        self.range_path.set("")
        self.output_dir.set("")
        self.overwrite.set(False)
        self.result = None
        self.report_file = None
        self.overtime_report_file = None
        self.log_file = None
        self.hero_title.configure(text="Carga los archivos para iniciar el análisis.")
        self.hero_text.configure(
            text="Aquí verás el corte del modo activo, los totales importantes y las alertas que sí requieren seguimiento."
        )
        self._set_metric_values("0", "0", "0", "0", "0")
        self._set_metric_captions("Plantilla válida", "Presentes", "Retardos", "Faltas", "Alertas")
        self._set_daily_result_tabs()
        for tree in self.trees.values():
            tree.delete(*tree.get_children())
            tree.configure(columns=())
        self.status_left.configure(text="Listo.")
        self.status_right.configure(text="Aún no se ha generado ningún reporte.")
        self.open_output_button.state(["disabled"])
        self.notebook.select(0)

    def open_output_dir(self):
        target = self.output_dir.get().strip()
        if not target:
            return
        try:
            subprocess.Popen(["explorer", str(Path(target))])
        except Exception as exc:
            messagebox.showerror("No se pudo abrir la carpeta", str(exc))

    def run_analysis(self):
        mode = self._get_current_mode()
        personal_path = self.personal_path.get().strip()
        output_dir = self.output_dir.get().strip()
        data_path = self.events_path.get().strip() if mode == "daily" else self.range_path.get().strip()

        if not personal_path or not data_path or not output_dir:
            required_label = "el archivo de eventos del día" if mode == "daily" else "el archivo de rango"
            messagebox.showwarning(
                "Faltan datos",
                f"Debes seleccionar el archivo de personal, {required_label} y la carpeta de salida.",
            )
            return

        status_text = "Analizando asistencia diaria..." if mode == "daily" else "Analizando rango de asistencia..."
        self.status_left.configure(text=status_text)
        self.update_idletasks()

        try:
            if mode == "daily":
                result: ResultType = calculate_attendance(
                    personal_path=Path(personal_path),
                    events_path=Path(data_path),
                    output_dir=Path(output_dir),
                    overwrite=self.overwrite.get(),
                )
            else:
                result = calculate_attendance_range(
                    personal_path=Path(personal_path),
                    range_events_path=Path(data_path),
                    output_dir=Path(output_dir),
                    overwrite=self.overwrite.get(),
                )
        except Exception as exc:
            self.status_left.configure(text="Ocurrió un error durante el análisis.")
            messagebox.showerror("Análisis fallido", str(exc))
            return

        self.result = result
        self.report_file = result.report_file
        self.overtime_report_file = result.overtime_report_file
        self.log_file = result.log_file
        self._update_result_view(result, mode)
        self.open_output_button.state(["!disabled"])
        self.status_right.configure(text=self._build_output_status(result))

        errors, cautions, notes = self._classify_issues(result.issues)
        if errors:
            self.status_left.configure(text="Análisis con errores.")
            self._show_issue_message("error", result, errors, [])
            return

        if cautions or notes:
            self.status_left.configure(text="Análisis completado con observaciones.")
            self._show_issue_message("notice", result, cautions, notes)
            return

        self.status_left.configure(text="Análisis completado.")
        self._show_issue_message("info", result, [], [])

    def _update_result_view(self, result: ResultType, mode: str):
        if mode == "daily":
            self._update_daily_result_view(result)
        else:
            self._update_range_result_view(result)

    def _update_daily_result_view(self, result: ResultType):
        assert isinstance(result, RunResult)
        self._set_daily_result_tabs()
        self.hero_title.configure(
            text=(
                f"Fecha {result.work_date_label} | Asistencias: {result.attendance_count} | "
                f"Retardos: {result.tardy_count} | Faltas: {result.absence_count}"
            )
        )
        hero_lines = [
            f"Horario aplicado: {result.schedule_label}",
            f"Personal con incidencias: {result.incident_employee_count}",
            (
                f"Horas extra: {result.total_overtime_hours} hora(s) en reporte separado"
                if result.overtime_report_file
                else "Horas extra: sin reporte separado para este día"
            ),
        ]
        if result.issues:
            hero_lines.append(f"Observaciones globales detectadas: {len(result.issues)}")
        self.hero_text.configure(text=" | ".join(hero_lines))

        self._set_metric_captions("Plantilla válida", "Presentes", "Retardos", "Faltas", "Alertas")
        self._set_metric_values(
            str(result.total_employees),
            str(result.attendance_count),
            str(result.tardy_count),
            str(result.absence_count),
            str(result.incident_employee_count),
        )

        self._fill_tree(self.trees["resumen"], result.summary_frame)
        self._fill_tree(self.trees["vista"], result.quick_view_frame)
        self._fill_tree(self.trees["faltas"], result.absence_frame)
        self._fill_tree(self.trees["retardos"], result.tardy_frame)
        self._fill_tree(self.trees["incidencias"], result.incident_frame)
        self._fill_tree(self.trees["detalle"], result.daily_frame)
        self.notebook.select(1)

    def _update_range_result_view(self, result: ResultType):
        assert isinstance(result, RangeRunResult)
        self._set_range_result_tabs()
        self.hero_title.configure(
            text=(
                f"Período {result.range_label} | Días laborales: {result.workday_count} | "
                f"Retardos: {result.tardy_count} | Faltas: {result.absence_count}"
            )
        )
        hero_lines = [
                f"Días con operación: {result.operational_day_count}",
                f"Días sin registros globales: {result.non_operational_day_count}",
            "Corte parcial detectado" if result.partial_cutoff else "Corte parcial no detectado",
            (
                f"Horas extra: {result.total_overtime_hours} hora(s) en reporte separado"
                if result.overtime_report_file
                else "Horas extra: sin reporte separado para este rango"
            ),
        ]
        if result.issues:
            hero_lines.append(f"Observaciones globales detectadas: {len(result.issues)}")
        self.hero_text.configure(text=" | ".join(hero_lines))

        self._set_metric_captions("Plantilla válida", "Días laborales", "Retardos", "Faltas", "Alertas")
        self._set_metric_values(
            str(result.total_employees),
            str(result.workday_count),
            str(result.tardy_count),
            str(result.absence_count),
            str(result.incident_employee_count),
        )

        self._fill_tree(self.trees["resumen"], result.summary_frame)
        self._fill_tree(self.trees["vista"], result.historical_preview_frame)
        self._fill_tree(self.trees["faltas"], result.absence_frame)
        self._fill_tree(self.trees["retardos"], result.tardy_frame)
        self._fill_tree(self.trees["incidencias"], result.incident_frame)
        self._fill_tree(self.trees["detalle"], result.detail_frame)
        self.notebook.select(0)

    def _set_metric_values(self, empleados: str, asistencias: str, retardos: str, faltas: str, incidencias: str):
        self.metric_cards["empleados"]["value"].configure(text=empleados)
        self.metric_cards["asistencias"]["value"].configure(text=asistencias)
        self.metric_cards["retardos"]["value"].configure(text=retardos)
        self.metric_cards["faltas"]["value"].configure(text=faltas)
        self.metric_cards["incidencias"]["value"].configure(text=incidencias)

    def _set_metric_captions(self, empleados: str, asistencias: str, retardos: str, faltas: str, incidencias: str):
        self.metric_cards["empleados"]["caption"].configure(text=empleados)
        self.metric_cards["asistencias"]["caption"].configure(text=asistencias)
        self.metric_cards["retardos"]["caption"].configure(text=retardos)
        self.metric_cards["faltas"]["caption"].configure(text=faltas)
        self.metric_cards["incidencias"]["caption"].configure(text=incidencias)

    def _build_output_status(self, result: ResultType) -> str:
        overtime_text = f" | Horas extra: {result.overtime_report_file.name}" if result.overtime_report_file else ""
        return f"Reporte: {result.report_file.name}{overtime_text}"

    def _classify_issues(self, issues: list[RunIssue]) -> tuple[list[str], list[str], list[str]]:
        errors: list[str] = []
        cautions: list[str] = []
        notes: list[str] = []
        caution_markers = (
            "corte parcial",
            "usuarios que no existen en la bbdd",
            "ids duplicados en personal",
            "eventos con tiempo invalido",
        )
        for issue in issues:
            text = issue.message
            normalized = text.lower()
            if issue.level.lower() == "error":
                errors.append(text)
            elif any(marker in normalized for marker in caution_markers):
                cautions.append(text)
            else:
                notes.append(text)
        return errors, cautions, notes

    def _show_issue_message(
        self,
        level: str,
        result: ResultType,
        issue_messages: list[str],
        note_messages: list[str],
    ):
        lines = [
            "Carpeta de salida:",
            str(result.report_file.parent),
            "",
            f"Reporte principal: {result.report_file.name}",
        ]
        if result.overtime_report_file:
            lines.append(f"Reporte de horas extra: {result.overtime_report_file.name}")

        if issue_messages:
            section_title = "Detalle del error:" if level == "error" else "Observaciones importantes:"
            lines.extend(["", section_title])
            lines.extend(f"- {message}" for message in issue_messages)
        if note_messages:
            lines.extend(["", "Notas informativas:"])
            lines.extend(f"- {message}" for message in note_messages)

        message = "\n".join(lines)
        if level == "error":
            messagebox.showerror("Análisis con errores", message)
        elif level == "notice":
            messagebox.showinfo("Reporte generado con observaciones", message)
        else:
            messagebox.showinfo("Reporte generado", message)

    def _fill_tree(self, tree: ttk.Treeview, frame):
        tree.delete(*tree.get_children())
        columns = list(frame.columns)
        tree.configure(columns=columns)
        for column in columns:
            tree.heading(column, text=column)
            width = max(110, min(300, len(column) * 11))
            if column in {"Campo", "Fecha", "Dia", "Horario"}:
                width = 130
            if column.startswith("Empleado"):
                width = 220
            if column == "Nombre":
                width = 240
            if column == "Detalle":
                width = 340
            if column in {"Entrada", "Inicio comida", "Fin comida", "Salida", "Horas trabajadas"}:
                width = 125
            tree.column(column, width=width, anchor="w")

        if frame.empty:
            return

        safe_frame = frame.fillna("")
        for row in safe_frame.itertuples(index=False):
            values = [value if not isinstance(value, float) else f"{value:g}" for value in row]
            tree.insert("", "end", values=values)


def main():
    app = LinherAttendanceApp()
    app.mainloop()


if __name__ == "__main__":
    main()
