from __future__ import annotations

import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import requests


REPO = "linhermx/attendance"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"

ASSET_NAME = "attendance_windows.zip"
APP_DIR_PREFIX = "attendance_v"
APP_EXE_NAME = "attendance_windows.exe"
BUNDLED_ASSET_DIR = "bundled_assets"
BUNDLED_RELEASE_META = "attendance_release.json"
STATE_FILE_NAME = "launcher_state.json"
CACHE_MAX_AGE_SECONDS = 30 * 60
LATEST_TIMEOUT_SECONDS = 8
BACKGROUND_LATEST_TIMEOUT_SECONDS = 4
DOWNLOAD_TIMEOUT_SECONDS = 60
LAUNCHER_MUTEX_NAME = "Local\\LINHER.Attendance.Launcher"


@dataclass
class InstalledApp:
    version: tuple[int, int, int]
    exe_path: Path
    home_path: Path


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def runtime_root() -> Path:
    override = os.getenv("LINHER_ATTENDANCE_RUNTIME_ROOT", "").strip()
    if override:
        return Path(override)

    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / "LINHER" / "Attendance"
    return Path.home() / "AppData" / "Local" / "LINHER" / "Attendance"


def work_area() -> tuple[int, int, int, int] | None:
    if os.name != "nt":
        return None

    rect = RECT()
    SPI_GETWORKAREA = 48
    if ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0):
        return rect.left, rect.top, rect.right, rect.bottom
    return None


def ensure_dirs(root: Path) -> dict[str, Path]:
    app_dir = root / "app"
    downloads_dir = root / "downloads"
    logs_dir = root / "logs"
    state_dir = root / "state"
    app_dir.mkdir(parents=True, exist_ok=True)
    downloads_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    state_dir.mkdir(parents=True, exist_ok=True)
    return {
        "root": root,
        "app": app_dir,
        "downloads": downloads_dir,
        "logs": logs_dir,
        "state": state_dir,
    }


def state_file(root: Path) -> Path:
    return root / "state" / STATE_FILE_NAME


def parse_version(tag: str) -> tuple[int, int, int] | None:
    match = re.match(r"^v?(\d+)\.(\d+)\.(\d+)$", tag.strip())
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def parse_version_from_dir_name(name: str) -> tuple[int, int, int] | None:
    match = re.match(rf"^{re.escape(APP_DIR_PREFIX)}(\d+)\.(\d+)\.(\d+)$", name)
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def format_version(version: tuple[int, int, int] | None) -> str:
    if version is None:
        return "ninguna"
    return "v" + ".".join(map(str, version))


def find_installed_app(app_dir: Path) -> InstalledApp | None:
    candidates: list[InstalledApp] = []

    for item in app_dir.iterdir():
        if not item.is_dir():
            continue
        version = parse_version_from_dir_name(item.name)
        if not version:
            continue
        exe_path = item / APP_EXE_NAME
        if exe_path.exists():
            candidates.append(InstalledApp(version=version, exe_path=exe_path, home_path=item))

    # Backward compatibility with old onefile installs.
    for item in app_dir.glob(f"{APP_DIR_PREFIX}*.exe"):
        version = parse_version_from_dir_name(item.stem)
        if not version:
            continue
        candidates.append(InstalledApp(version=version, exe_path=item, home_path=item.parent))

    if not candidates:
        return None

    candidates.sort(key=lambda value: value.version, reverse=True)
    return candidates[0]


def load_cached_release(path: Path) -> dict[str, object] | None:
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

    tag = str(data.get("tag", "")).strip()
    version = parse_version(tag)
    url = str(data.get("url", "")).strip()
    checked_at = data.get("checked_at")
    asset_name = str(data.get("asset_name", ASSET_NAME)).strip()
    if not version or not url or not isinstance(checked_at, (int, float)):
        return None

    return {
        "version": version,
        "tag": tag,
        "url": url,
        "asset_name": asset_name,
        "checked_at": float(checked_at),
    }


def save_cached_release(
    path: Path,
    *,
    version: tuple[int, int, int],
    tag: str,
    url: str,
    asset_name: str,
) -> None:
    payload = {
        "version": list(version),
        "tag": tag,
        "url": url,
        "asset_name": asset_name,
        "checked_at": time.time(),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def cache_is_fresh(cache: dict[str, object] | None) -> bool:
    if not cache:
        return False
    checked_at = float(cache.get("checked_at", 0.0) or 0.0)
    return (time.time() - checked_at) < CACHE_MAX_AGE_SECONDS


def get_latest_release(
    timeout_seconds: int = LATEST_TIMEOUT_SECONDS,
) -> tuple[tuple[int, int, int], str, str, str]:
    response = requests.get(LATEST_API, timeout=timeout_seconds)
    response.raise_for_status()
    data = response.json()

    tag = data.get("tag_name", "")
    version = parse_version(tag)
    if not version:
        raise RuntimeError(f"Tag inválido en latest release: {tag!r}")

    assets = data.get("assets", []) or []
    asset = next((item for item in assets if item.get("name") == ASSET_NAME), None)
    if not asset:
        raise RuntimeError(
            f"No se encontró el asset '{ASSET_NAME}' en el latest release.\n"
            "Sube ese paquete como asset en GitHub Releases."
        )

    url = asset.get("browser_download_url", "")
    if not url:
        raise RuntimeError("El asset publicado no tiene browser_download_url.")
    return version, tag, url, ASSET_NAME


def human_size(size_in_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB"]
    value = float(size_in_bytes)
    unit_index = 0
    while value >= 1024 and unit_index < len(units) - 1:
        value /= 1024.0
        unit_index += 1
    return f"{value:.1f} {units[unit_index]}"


class LoadingWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.brand_image: tk.PhotoImage | None = None
        self.window_width = 520
        self.window_height = 220
        self.root.title("LINHER Attendance")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        self._apply_icon()
        self._build_ui()
        self._center()
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(300, lambda: self.root.attributes("-topmost", False))
        self.root.update_idletasks()
        self.root.update()

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        style = ttk.Style(self.root)
        for theme in ("vista", "xpnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break

        row_index = 0
        brand = self._load_brand_image()
        if brand is not None:
            self.brand_image = brand
            ttk.Label(frame, image=self.brand_image).grid(row=row_index, column=0, sticky="w", pady=(0, 10))
            row_index += 1
            self.window_height = 250

        ttk.Label(
            frame,
            text="Preparando Attendance",
            font=("Segoe UI", 14, "bold"),
        ).grid(row=row_index, column=0, sticky="w")
        row_index += 1

        self.status_var = tk.StringVar(value="Inicializando...")
        self.detail_var = tk.StringVar(value="Preparando el entorno del launcher.")

        ttk.Label(
            frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "bold"),
        ).grid(row=row_index, column=0, sticky="w", pady=(14, 4))
        row_index += 1
        ttk.Label(
            frame,
            textvariable=self.detail_var,
            wraplength=470,
            justify="left",
        ).grid(row=row_index, column=0, sticky="w")
        row_index += 1

        self.progress = ttk.Progressbar(frame, orient="horizontal", length=470, mode="indeterminate")
        self.progress.grid(row=row_index, column=0, sticky="ew", pady=(16, 0))
        self.progress.start(10)

    def _apply_icon(self):
        icon_path = base_dir() / "resources" / "attendance.ico"
        if not icon_path.exists():
            icon_path = base_dir() / "src" / "attendance" / "resources" / "attendance.ico"
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(default=str(icon_path))
        except Exception:
            pass

    def _load_brand_image(self) -> tk.PhotoImage | None:
        candidate_paths = [
            base_dir() / "resources" / "attendance_brand.png",
            base_dir() / "resources" / "attendance_brand.gif",
            base_dir() / "src" / "attendance" / "resources" / "attendance_brand.png",
            base_dir() / "src" / "attendance" / "resources" / "attendance_brand.gif",
        ]
        brand_path = next((path for path in candidate_paths if path.exists()), None)
        if not brand_path:
            return None

        try:
            image = tk.PhotoImage(file=str(brand_path))
        except Exception:
            return None

        max_width = 470
        image_width = max(1, image.width())
        if image_width > max_width:
            scale = max(1, (image_width + max_width - 1) // max_width)
            image = image.subsample(scale, scale)
        return image

    def _center(self):
        self.root.update_idletasks()
        width = max(self.window_width, self.root.winfo_reqwidth())
        height = max(self.window_height, self.root.winfo_reqheight())
        usable_area = work_area()
        if usable_area:
            left, top, right, bottom = usable_area
            x = max(left, left + ((right - left - width) // 2))
            y = max(top, top + ((bottom - top - height) // 2))
        else:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = max(0, (screen_width - width) // 2)
            y = max(0, (screen_height - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def set_status(self, title: str, detail: str, *, determinate: bool = False, progress: float | None = None):
        self.status_var.set(title)
        self.detail_var.set(detail)
        if determinate:
            if str(self.progress.cget("mode")) != "determinate":
                self.progress.stop()
                self.progress.configure(mode="determinate", maximum=100)
            self.progress["value"] = 0 if progress is None else max(0, min(progress, 100))
        else:
            if str(self.progress.cget("mode")) != "indeterminate":
                self.progress.configure(mode="indeterminate")
                self.progress.start(10)
        self.pump()

    def ask_yes_no(self, title: str, message: str) -> bool:
        self.progress.stop()
        try:
            return messagebox.askyesno(title, message, parent=self.root)
        finally:
            if str(self.progress.cget("mode")) == "indeterminate":
                self.progress.start(10)

    def show_info(self, title: str, message: str):
        messagebox.showinfo(title, message, parent=self.root)

    def show_warning(self, title: str, message: str):
        messagebox.showwarning(title, message, parent=self.root)

    def show_error(self, title: str, message: str):
        messagebox.showerror(title, message, parent=self.root)

    def pump(self):
        self.root.update_idletasks()
        self.root.update()

    def close(self):
        try:
            self.progress.stop()
        except Exception:
            pass
        self.root.destroy()


class SingleInstanceGuard:
    def __init__(self, name: str):
        self.name = name
        self.handle = None

    def acquire(self) -> bool:
        if os.name != "nt":
            return True
        kernel32 = ctypes.windll.kernel32
        self.handle = kernel32.CreateMutexW(None, False, self.name)
        return kernel32.GetLastError() != 183

    def release(self):
        if os.name != "nt" or not self.handle:
            return
        ctypes.windll.kernel32.ReleaseMutex(self.handle)
        ctypes.windll.kernel32.CloseHandle(self.handle)
        self.handle = None


def bundled_assets_dir() -> Path:
    return base_dir() / BUNDLED_ASSET_DIR


def bundled_archive_path() -> Path:
    return bundled_assets_dir() / ASSET_NAME


def bundled_release_metadata() -> dict[str, object] | None:
    metadata_path = bundled_assets_dir() / BUNDLED_RELEASE_META
    if not metadata_path.exists():
        return None
    try:
        data = json.loads(metadata_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return None

    tag = str(data.get("tag", "")).strip()
    version = parse_version(tag)
    asset_name = str(data.get("asset_name", ASSET_NAME)).strip()
    if not version:
        return None
    return {"version": version, "tag": tag, "asset_name": asset_name}


def download_file(url: str, dst: Path, ui: LoadingWindow | None = None) -> None:
    downloaded = 0
    with requests.get(url, stream=True, timeout=DOWNLOAD_TIMEOUT_SECONDS) as response:
        response.raise_for_status()
        total_bytes = int(response.headers.get("Content-Length", "0") or 0)
        with dst.open("wb") as file:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if not chunk:
                    continue
                file.write(chunk)
                downloaded += len(chunk)
                if ui:
                    if total_bytes > 0:
                        progress = (downloaded / total_bytes) * 100
                        ui.set_status(
                            "Descargando actualización...",
                            f"{human_size(downloaded)} de {human_size(total_bytes)} descargados.",
                            determinate=True,
                            progress=progress,
                        )
                    else:
                        ui.set_status(
                            "Descargando actualización...",
                            f"{human_size(downloaded)} descargados.",
                            determinate=False,
                        )


def launcher_command(*args: str) -> list[str]:
    if getattr(sys, "frozen", False):
        return [str(Path(sys.executable).resolve()), *args]
    return [sys.executable, str(base_dir() / "attendance_launcher.py"), *args]


def spawn_background_refresh() -> None:
    creation_flags = 0
    if os.name == "nt":
        creation_flags = (
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "DETACHED_PROCESS", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    try:
        subprocess.Popen(
            launcher_command("--refresh-cache"),
            cwd=str(base_dir()),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
            close_fds=True,
            creationflags=creation_flags,
        )
    except Exception:
        pass


def refresh_release_cache(cache_path: Path) -> None:
    try:
        latest_version, latest_tag, asset_url, asset_name = get_latest_release(
            timeout_seconds=BACKGROUND_LATEST_TIMEOUT_SECONDS
        )
    except Exception:
        return
    save_cached_release(
        cache_path,
        version=latest_version,
        tag=latest_tag,
        url=asset_url,
        asset_name=asset_name,
    )


def locate_extracted_app_dir(extract_root: Path) -> Path:
    direct_exe = extract_root / APP_EXE_NAME
    if direct_exe.exists():
        return extract_root

    for exe_path in extract_root.rglob(APP_EXE_NAME):
        return exe_path.parent

    raise RuntimeError(
        f"El paquete descargado no contiene '{APP_EXE_NAME}'. "
        "Verifica el asset publicado en GitHub Releases."
    )


def install_archive(dirs: dict[str, Path], version: tuple[int, int, int], archive_path: Path) -> Path:
    target_dir = dirs["app"] / f"{APP_DIR_PREFIX}{version[0]}.{version[1]}.{version[2]}"
    target_exe = target_dir / APP_EXE_NAME
    if target_exe.exists():
        return target_exe

    staging_root = Path(tempfile.mkdtemp(prefix="attendance_extract_", dir=str(dirs["downloads"])))
    try:
        with zipfile.ZipFile(archive_path, "r") as package:
            package.extractall(staging_root)

        extracted_dir = locate_extracted_app_dir(staging_root)
        target_dir.parent.mkdir(parents=True, exist_ok=True)
        if target_dir.exists():
            shutil.rmtree(target_dir, ignore_errors=True)
        shutil.move(str(extracted_dir), str(target_dir))
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)

    if not target_exe.exists():
        raise RuntimeError(
            "La versión instalada no quedó completa. "
            f"No se encontró '{APP_EXE_NAME}' después de extraer el paquete."
        )
    return target_exe


def install_release(
    *,
    dirs: dict[str, Path],
    version: tuple[int, int, int],
    url: str,
    ui: LoadingWindow | None = None,
) -> Path:
    tmp_archive = dirs["downloads"] / ASSET_NAME
    download_file(url, tmp_archive, ui=ui)
    try:
        return install_archive(dirs, version, tmp_archive)
    finally:
        tmp_archive.unlink(missing_ok=True)


def install_bundled_release(
    *,
    dirs: dict[str, Path],
    version: tuple[int, int, int],
) -> Path:
    archive_path = bundled_archive_path()
    if not archive_path.exists():
        raise RuntimeError("No se encontró el paquete incluido por el instalador.")
    return install_archive(dirs, version, archive_path)


def run_app(exe_path: Path, ui: LoadingWindow | None = None) -> None:
    if ui:
        ui.set_status("Abriendo aplicación...", f"Se abrirá {exe_path.name}.")
        ui.close()
    subprocess.Popen([str(exe_path)], cwd=str(exe_path.parent))
    raise SystemExit(0)


def run_local_source(ui: LoadingWindow | None = None) -> None:
    source_entry = base_dir() / "attendance_gui.py"
    if ui:
        ui.set_status("Abriendo entorno local...", "Se ejecutará la GUI local para soporte.")
        ui.close()
    subprocess.Popen([sys.executable, str(source_entry)], cwd=str(source_entry.parent))
    raise SystemExit(0)


def show_single_instance_message():
    root = tk.Tk()
    root.withdraw()
    try:
        messagebox.showinfo(
            "Attendance ya se está iniciando",
            "Ya hay un proceso del launcher en ejecución. Espera a que termine de abrir la aplicación.",
            parent=root,
        )
    finally:
        root.destroy()


def main() -> None:
    guard = SingleInstanceGuard(LAUNCHER_MUTEX_NAME)
    if not guard.acquire():
        return

    try:
        dirs = ensure_dirs(runtime_root())
        cache_path = state_file(dirs["root"])

        if "--refresh-cache" in sys.argv[1:]:
            refresh_release_cache(cache_path)
            return

        ui = LoadingWindow()
        ui.set_status("Inicializando Attendance...", "Preparando el entorno del launcher.")

        installed = find_installed_app(dirs["app"])
        installed_version = installed.version if installed else None
        installed_exe = installed.exe_path if installed else None
        cached_release = load_cached_release(cache_path)
        bundled_release = bundled_release_metadata()

        if installed_exe:
            ui.set_status("Revisando instalación local...", f"Versión detectada: {format_version(installed_version)}.")

            if cached_release and cached_release["version"] > installed_version:
                latest_version = cached_release["version"]
                latest_tag = str(cached_release["tag"])
                asset_url = str(cached_release["url"])
                message = (
                    f"Hay una versión disponible: {latest_tag}\n"
                    f"Instalada: {format_version(installed_version)}\n\n"
                    "¿Deseas actualizar ahora?"
                )
                if ui.ask_yes_no("Actualización disponible", message):
                    try:
                        installed_exe = install_release(
                            dirs=dirs,
                            version=latest_version,
                            url=asset_url,
                            ui=ui,
                        )
                    except Exception as exc:
                        ui.show_warning(
                            "Actualización fallida",
                            f"No se pudo actualizar:\n{exc}\n\nSe abrirá la versión instalada.",
                        )

                if not cache_is_fresh(cached_release):
                    spawn_background_refresh()
                run_app(installed_exe, ui)

            if not cache_is_fresh(cached_release):
                ui.set_status("Verificando actualizaciones...", "Se refrescará la información en segundo plano.")
                spawn_background_refresh()
            run_app(installed_exe, ui)

        if bundled_release:
            try:
                ui.set_status(
                    "Instalando la versión incluida...",
                    f"Se preparará {bundled_release['tag']} incluida por el instalador.",
                )
                installed_exe = install_bundled_release(
                    dirs=dirs,
                    version=bundled_release["version"],
                )
                spawn_background_refresh()
                run_app(installed_exe, ui)
            except Exception as exc:
                ui.show_warning(
                    "No se pudo usar la versión incluida",
                    f"La instalación inicial no se pudo preparar desde el paquete local:\n{exc}\n\nSe intentará descargar la versión publicada.",
                )

        try:
            ui.set_status("Verificando actualizaciones...", "Consultando la última versión publicada.")
            if cached_release and cache_is_fresh(cached_release):
                latest_version = cached_release["version"]
                latest_tag = str(cached_release["tag"])
                asset_url = str(cached_release["url"])
                asset_name = str(cached_release["asset_name"])
            else:
                latest_version, latest_tag, asset_url, asset_name = get_latest_release()
                save_cached_release(
                    cache_path,
                    version=latest_version,
                    tag=latest_tag,
                    url=asset_url,
                    asset_name=asset_name,
                )
        except Exception as exc:
            if installed_exe:
                ui.show_warning(
                    "Actualización no disponible",
                    f"No se pudo verificar la actualización:\n{exc}\n\nSe abrirá la versión instalada.",
                )
                run_app(installed_exe, ui)

            if not getattr(sys, "frozen", False):
                ui.show_warning(
                    "Sin conexión a releases",
                    f"No se pudo verificar la actualización:\n{exc}\n\nSe abrirá la GUI local.",
                )
                run_local_source(ui)

            ui.show_error(
                "No hay app instalada",
                "No se pudo verificar la actualización y no existe una app instalada.\n\n"
                f"Detalle:\n{exc}"
            )
            ui.close()
            return

        message = (
            f"Hay una versión disponible: {latest_tag}\n"
            f"Instalada: {format_version(installed_version)}\n\n"
            "¿Deseas instalarla ahora?"
        )
        if installed_exe is None or ui.ask_yes_no("Actualización disponible", message):
            try:
                installed_exe = install_release(
                    dirs=dirs,
                    version=latest_version,
                    url=asset_url,
                    ui=ui,
                )
            except Exception as exc:
                if not getattr(sys, "frozen", False):
                    ui.show_warning(
                        "Actualización fallida",
                        f"No se pudo descargar la versión publicada:\n{exc}\n\nSe abrirá la GUI local.",
                    )
                    run_local_source(ui)
                ui.show_error("Actualización fallida", f"No se pudo descargar o instalar:\n{exc}")
                ui.close()
                return

        if installed_exe:
            run_app(installed_exe, ui)

        if not getattr(sys, "frozen", False):
            run_local_source(ui)

        ui.show_error("No se encontró la app", "No hay ejecutable para abrir.")
        ui.close()
    finally:
        guard.release()


if __name__ == "__main__":
    main()
