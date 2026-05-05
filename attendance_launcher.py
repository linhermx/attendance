from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

import requests


REPO = "linhermx/attendance"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"

ASSET_NAME = "attendance_windows.exe"
APP_EXE_PREFIX = "attendance_v"
APP_EXE_SUFFIX = ".exe"
STATE_FILE_NAME = "launcher_state.json"
CACHE_MAX_AGE_SECONDS = 30 * 60
LATEST_TIMEOUT_SECONDS = 8
BACKGROUND_LATEST_TIMEOUT_SECONDS = 4
DOWNLOAD_TIMEOUT_SECONDS = 60


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def runtime_root() -> Path:
    local_appdata = os.getenv("LOCALAPPDATA")
    if local_appdata:
        return Path(local_appdata) / "LINHER" / "Attendance"
    return Path.home() / "AppData" / "Local" / "LINHER" / "Attendance"


def ensure_dirs(root: Path) -> dict[str, Path]:
    app_dir = root / "app"
    downloads_dir = root / "downloads"
    logs_dir = root / "logs"
    app_dir.mkdir(parents=True, exist_ok=True)
    downloads_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    return {"root": root, "app": app_dir, "downloads": downloads_dir, "logs": logs_dir}


def state_file(root: Path) -> Path:
    return root / STATE_FILE_NAME


def parse_version(tag: str) -> tuple[int, int, int] | None:
    match = re.match(r"^v?(\d+)\.(\d+)\.(\d+)$", tag.strip())
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def parse_version_from_name(name: str) -> tuple[int, int, int] | None:
    match = re.match(
        rf"^{re.escape(APP_EXE_PREFIX)}(\d+)\.(\d+)\.(\d+){re.escape(APP_EXE_SUFFIX)}$",
        name,
    )
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def format_version(version: tuple[int, int, int] | None) -> str:
    if version is None:
        return "ninguna"
    return "v" + ".".join(map(str, version))


def find_installed_app(app_dir: Path) -> tuple[tuple[int, int, int], Path] | None:
    candidates = []
    for item in app_dir.glob(f"{APP_EXE_PREFIX}*{APP_EXE_SUFFIX}"):
        version = parse_version_from_name(item.name)
        if version:
            candidates.append((version, item))
    if not candidates:
        return None
    candidates.sort(key=lambda value: value[0], reverse=True)
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
    if not version or not url or not isinstance(checked_at, (int, float)):
        return None

    return {
        "version": version,
        "tag": tag,
        "url": url,
        "checked_at": float(checked_at),
    }


def save_cached_release(path: Path, *, version: tuple[int, int, int], tag: str, url: str) -> None:
    payload = {
        "version": list(version),
        "tag": tag,
        "url": url,
        "checked_at": time.time(),
    }
    path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")


def cache_is_fresh(cache: dict[str, object] | None) -> bool:
    if not cache:
        return False
    checked_at = float(cache.get("checked_at", 0.0) or 0.0)
    return (time.time() - checked_at) < CACHE_MAX_AGE_SECONDS


def get_latest_release(timeout_seconds: int = LATEST_TIMEOUT_SECONDS) -> tuple[tuple[int, int, int], str, str]:
    response = requests.get(LATEST_API, timeout=timeout_seconds)
    response.raise_for_status()
    data = response.json()

    tag = data.get("tag_name", "")
    version = parse_version(tag)
    if not version:
        raise RuntimeError(f"Tag invalido en latest release: {tag!r}")

    assets = data.get("assets", []) or []
    asset = next((item for item in assets if item.get("name") == ASSET_NAME), None)
    if not asset:
        raise RuntimeError(
            f"No se encontro el asset '{ASSET_NAME}' en el latest release.\n"
            "Sube ese ejecutable como asset en GitHub Releases."
        )

    url = asset.get("browser_download_url", "")
    if not url:
        raise RuntimeError("Asset sin browser_download_url")
    return version, tag, url


def download_file(url: str, dst: Path) -> None:
    with requests.get(url, stream=True, timeout=DOWNLOAD_TIMEOUT_SECONDS) as response:
        response.raise_for_status()
        with dst.open("wb") as file:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if chunk:
                    file.write(chunk)


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
        latest_version, latest_tag, asset_url = get_latest_release(
            timeout_seconds=BACKGROUND_LATEST_TIMEOUT_SECONDS
        )
    except Exception:
        return
    save_cached_release(cache_path, version=latest_version, tag=latest_tag, url=asset_url)


def install_release(
    *,
    dirs: dict[str, Path],
    version: tuple[int, int, int],
    url: str,
) -> Path:
    tmp = dirs["downloads"] / ASSET_NAME
    target = dirs["app"] / (
        f"{APP_EXE_PREFIX}{version[0]}.{version[1]}.{version[2]}{APP_EXE_SUFFIX}"
    )
    download_file(url, tmp)
    if target.exists():
        tmp.unlink(missing_ok=True)
        return target
    shutil.move(str(tmp), str(target))
    return target


def run_app(exe_path: Path) -> None:
    subprocess.Popen([str(exe_path)], cwd=str(exe_path.parent))
    raise SystemExit(0)


def run_local_source() -> None:
    source_entry = base_dir() / "attendance_gui.py"
    subprocess.Popen([sys.executable, str(source_entry)], cwd=str(source_entry.parent))
    raise SystemExit(0)


def main() -> None:
    dirs = ensure_dirs(runtime_root())
    cache_path = state_file(dirs["root"])

    if "--refresh-cache" in sys.argv[1:]:
        refresh_release_cache(cache_path)
        return

    tk_root = tk.Tk()
    tk_root.withdraw()

    installed = find_installed_app(dirs["app"])
    installed_version = installed[0] if installed else None
    installed_exe = installed[1] if installed else None
    cached_release = load_cached_release(cache_path)

    if installed_exe:
        if cached_release and cached_release["version"] > installed_version:
            latest_version = cached_release["version"]
            latest_tag = str(cached_release["tag"])
            asset_url = str(cached_release["url"])
            message = (
                f"Hay una version disponible: {latest_tag}\n"
                f"Instalada: {format_version(installed_version)}\n\n"
                "Deseas actualizar ahora?"
            )
            if messagebox.askyesno("Actualizacion disponible", message):
                try:
                    installed_exe = install_release(dirs=dirs, version=latest_version, url=asset_url)
                except Exception as exc:
                    messagebox.showwarning(
                        "Update fallo",
                        f"No se pudo actualizar:\n{exc}\n\nSe abrira la version instalada.",
                    )

            if not cache_is_fresh(cached_release):
                spawn_background_refresh()
            run_app(installed_exe)

        if not cache_is_fresh(cached_release):
            spawn_background_refresh()
        run_app(installed_exe)

    try:
        if cached_release and cache_is_fresh(cached_release):
            latest_version = cached_release["version"]
            latest_tag = str(cached_release["tag"])
            asset_url = str(cached_release["url"])
        else:
            latest_version, latest_tag, asset_url = get_latest_release()
            save_cached_release(cache_path, version=latest_version, tag=latest_tag, url=asset_url)
    except Exception as exc:
        if installed_exe:
            messagebox.showwarning(
                "Actualizacion no disponible",
                f"No se pudo verificar update:\n{exc}\n\nSe abrira la version instalada.",
            )
            run_app(installed_exe)

        if not getattr(sys, "frozen", False):
            messagebox.showwarning(
                "Sin conexion a releases",
                f"No se pudo verificar update:\n{exc}\n\nSe abrira la GUI local.",
            )
            run_local_source()

        messagebox.showerror(
            "No hay app instalada",
            f"No se pudo verificar update y no existe app instalada.\n\nDetalle:\n{exc}",
        )
        return

    message = (
        f"Hay una version disponible: {latest_tag}\n"
        f"Instalada: {format_version(installed_version)}\n\n"
        "Deseas instalarla ahora?"
    )
    if installed_exe is None or messagebox.askyesno("Actualizacion disponible", message):
        try:
            installed_exe = install_release(dirs=dirs, version=latest_version, url=asset_url)
        except Exception as exc:
            if not getattr(sys, "frozen", False):
                messagebox.showwarning(
                    "Update fallo",
                    f"No se pudo descargar la version publicada:\n{exc}\n\nSe abrira la GUI local.",
                )
                run_local_source()
            messagebox.showerror("Update fallo", f"No se pudo descargar/instalar:\n{exc}")
            return

    if installed_exe:
        run_app(installed_exe)

    if not getattr(sys, "frozen", False):
        run_local_source()

    messagebox.showerror("No se encontro la app", "No hay ejecutable para abrir.")


if __name__ == "__main__":
    main()
