from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import tkinter as tk
from ctypes import wintypes
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from attendance.core import RunResult, calculate_attendance
from attendance.version import __version__


class LinherAttendanceApp(tk.Tk):
    def __init__(self):
        self._set_windows_app_id()
        super().__init__()
        self.withdraw()
        self.title(f"LINHER Attendance | Control diario (v{__version__})")
        self._apply_window_icon()

        self.personal_path = tk.StringVar(value="")
        self.events_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value="")
        self.overwrite = tk.BooleanVar(value=False)

        self.last_dir = Path.home()
        self.result: RunResult | None = None
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
        width = 1220
        height = 860
        work_area = self._get_work_area()
        if work_area:
            left, top, right, bottom = work_area
            available_width = max(960, right - left)
            available_height = max(720, bottom - top)
            width = min(width, available_width - 60)
            height = min(height, available_height - 60)
            x = left + max(0, (available_width - width) // 2)
            y = top + max(0, (available_height - height) // 5)
            self.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.geometry(f"{width}x{height}")
        self.minsize(1080, 760)

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

        ttk.Label(header, text="Control Diario de Asistencia", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=(
                "Carga los archivos del checador para generar un corte claro del día: faltas, "
                "retardos, omisiones de checada, salidas anticipadas y horas extra pagables."
            ),
            style="Subtitle.TLabel",
            wraplength=1080,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(
            header,
            text="Flujo sugerido: 1. Cargar archivos  2. Ejecutar análisis  3. Compartir la Vista rápida",
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

        guide_panel = ttk.LabelFrame(sidebar, text="Qué se vigila", style="Section.TLabelframe", padding=12)
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
                "No se eleva como alerta:\n"
                "- comida menor al mínimo"
            ),
            justify="left",
            style="Subtitle.TLabel",
        ).grid(row=0, column=0, sticky="w")

        content = ttk.Frame(body)
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        result_block = ttk.LabelFrame(content, text="Corte del día", style="Section.TLabelframe", padding=12)
        result_block.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        result_block.columnconfigure(0, weight=1)

        hero = tk.Frame(result_block, bg="#E9F5EF", bd=1, relief="solid")
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        hero.columnconfigure(0, weight=1)
        self.hero_title = ttk.Label(
            hero,
            text="Carga personal y eventos para iniciar el análisis.",
            style="HeroTitle.TLabel",
        )
        self.hero_text = ttk.Label(
            hero,
            text=(
                "Aquí verás con rapidez la fecha analizada, el tamaño de la plantilla, las faltas, "
                "los retardos y las alertas que sí requieren seguimiento."
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
        self._add_tree_tab("vista", "Vista rápida")
        self._add_tree_tab("faltas", "Faltas")
        self._add_tree_tab("retardos", "Retardos")
        self._add_tree_tab("incidencias", "Incidencias")
        self._add_tree_tab("detalle", "Detalle diario")
        self._add_tree_tab("resumen", "Resumen")

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
        panel.columnconfigure(1, weight=1)

        ttk.Label(panel, text="Personal").grid(row=0, column=0, sticky="w", pady=(2, 10))
        ttk.Entry(panel, textvariable=self.personal_path).grid(row=0, column=1, sticky="ew", padx=(12, 8))
        ttk.Button(panel, text="Seleccionar...", command=self.pick_personal).grid(row=0, column=2, sticky="ew")

        ttk.Label(panel, text="Eventos del día").grid(row=1, column=0, sticky="w", pady=(2, 10))
        ttk.Entry(panel, textvariable=self.events_path).grid(row=1, column=1, sticky="ew", padx=(12, 8))
        ttk.Button(panel, text="Seleccionar...", command=self.pick_events).grid(row=1, column=2, sticky="ew")

        ttk.Label(panel, text="Carpeta de salida").grid(row=2, column=0, sticky="w", pady=(2, 8))
        ttk.Entry(panel, textvariable=self.output_dir).grid(row=2, column=1, sticky="ew", padx=(12, 8))
        ttk.Button(panel, text="Seleccionar...", command=self.pick_output_dir).grid(row=2, column=2, sticky="ew")

        ttk.Label(
            panel,
            text=(
                "Se leen directo los .xls del checador. Si pegas rutas manualmente, la app las respeta."
            ),
            foreground="#4B6482",
            wraplength=300,
            justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _build_action_panel(self, parent: ttk.Frame):
        panel = ttk.LabelFrame(parent, text="Ejecución", style="Section.TLabelframe", padding=12)
        panel.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        panel.columnconfigure(0, weight=1)

        ttk.Label(panel, text="Configuración del reporte diario").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )
        ttk.Checkbutton(panel, text="Sobrescribir si existe", variable=self.overwrite).grid(row=1, column=0, sticky="w", pady=(0, 14))
        ttk.Button(panel, text="Analizar asistencia", command=self.run_analysis).grid(
            row=2, column=0, sticky="ew", pady=(0, 12)
        )
        ttk.Button(panel, text="Restablecer", command=self.reset_form).grid(
            row=3, column=0, sticky="ew", pady=(0, 8)
        )
        self.open_output_button = ttk.Button(panel, text="Abrir carpeta...", command=self.open_output_dir)
        self.open_output_button.grid(row=4, column=0, sticky="ew")
        self.open_output_button.state(["disabled"])
        ttk.Label(
            panel,
            text=(
                "El reporte destacará faltas, retardos, omisiones de checada, comida excedida y salida anticipada. "
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
        return value_label

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

    def pick_personal(self):
        path = self._pick_file("Selecciona el archivo de personal")
        if path:
            self.personal_path.set(path)

    def pick_events(self):
        path = self._pick_file("Selecciona el archivo de eventos")
        if path:
            self.events_path.set(path)

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
        self.output_dir.set("")
        self.overwrite.set(False)
        self.result = None
        self.report_file = None
        self.overtime_report_file = None
        self.log_file = None
        self.hero_title.configure(text="Carga personal y eventos para iniciar el análisis.")
        self.hero_text.configure(
            text=(
                "Aquí verás con rapidez la fecha analizada, el tamaño de la plantilla, las faltas, "
                "los retardos y las alertas que sí requieren seguimiento."
            )
        )
        for label in self.metric_cards.values():
            label.configure(text="0")
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
        personal_path = self.personal_path.get().strip()
        events_path = self.events_path.get().strip()
        output_dir = self.output_dir.get().strip()
        if not personal_path or not events_path or not output_dir:
            messagebox.showwarning(
                "Faltan datos",
                "Debes seleccionar el archivo de personal, el archivo de eventos y la carpeta de salida.",
            )
            return

        self.status_left.configure(text="Analizando asistencia...")
        self.update_idletasks()
        try:
            result = calculate_attendance(
                personal_path=Path(personal_path),
                events_path=Path(events_path),
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
        self._update_result_view(result)
        self.open_output_button.state(["!disabled"])
        self.status_left.configure(text="Análisis completado.")
        overtime_text = (
            f" | Horas extra: {result.overtime_report_file.name}"
            if result.overtime_report_file
            else ""
        )
        self.status_right.configure(text=f"Reporte: {result.report_file.name}{overtime_text} | Log: {result.log_file.name}")
        self._show_report_saved_message(result)

    def _update_result_view(self, result: RunResult):
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

        self.metric_cards["empleados"].configure(text=str(result.total_employees))
        self.metric_cards["asistencias"].configure(text=str(result.attendance_count))
        self.metric_cards["retardos"].configure(text=str(result.tardy_count))
        self.metric_cards["faltas"].configure(text=str(result.absence_count))
        self.metric_cards["incidencias"].configure(text=str(result.incident_employee_count))

        self._fill_tree(self.trees["resumen"], result.summary_frame)
        self._fill_tree(self.trees["vista"], result.quick_view_frame)
        self._fill_tree(self.trees["faltas"], result.absence_frame)
        self._fill_tree(self.trees["retardos"], result.tardy_frame)
        self._fill_tree(self.trees["incidencias"], result.incident_frame)
        self._fill_tree(self.trees["detalle"], result.daily_frame)
        self.notebook.select(self.notebook.tabs()[0])

    def _show_report_saved_message(self, result: RunResult):
        lines = [
            "El reporte se generó correctamente.",
            "",
            f"Carpeta de salida: {result.report_file.parent}",
            f"Reporte principal: {result.report_file.name}",
        ]
        if result.overtime_report_file:
            lines.append(f"Reporte de horas extra: {result.overtime_report_file.name}")
        lines.append(f"Log: {result.log_file.name}")
        renamed_outputs = [
            issue.message for issue in result.issues if "Se guardó como" in issue.message or "se guardó como" in issue.message
        ]
        if renamed_outputs:
            lines.extend(["", "Notas de guardado:"])
            lines.extend(f"- {message}" for message in renamed_outputs)
        messagebox.showinfo("Reporte generado", "\n".join(lines))

    def _fill_tree(self, tree: ttk.Treeview, frame: pd.DataFrame):
        tree.delete(*tree.get_children())
        columns = list(frame.columns)
        tree.configure(columns=columns)
        for column in columns:
            tree.heading(column, text=column)
            width = max(110, min(280, len(column) * 11))
            if column == "Campo":
                width = 140
            if column.startswith("Empleado"):
                width = 220
            if column == "Nombre":
                width = 220
            if column == "Detalle":
                width = 320
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
