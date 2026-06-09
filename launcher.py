"""Windows launcher for the bundled Streamlit attendance app."""

from __future__ import annotations

import ctypes
import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


APP_NAME = "OutsourceAttendance"
HOST = "127.0.0.1"
PREFERRED_PORT = 8501
REQUIRE_MONGODB_ENV_VAR = "DATALENS_REQUIRE_MONGODB"
NO_BROWSER_ENV_VAR = "OUTSOURCE_ATTENDANCE_NO_BROWSER"


def _bundled_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


def _app_data_dir() -> Path:
    root = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("APPDATA")
        or str(Path.home() / "AppData" / "Local")
    )
    return Path(root) / APP_NAME


def _prepare_runtime_paths() -> Path:
    app_data = _app_data_dir()
    log_dir = app_data / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    os.environ.setdefault(REQUIRE_MONGODB_ENV_VAR, "1")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

    if getattr(sys, "frozen", False):
        log_file = log_dir / "launcher.log"
        log_handle = open(log_file, "a", encoding="utf-8", buffering=1)
        sys.stdout = log_handle
        sys.stderr = log_handle

    return app_data


def _find_available_port(preferred: int = PREFERRED_PORT) -> int:
    for port in range(preferred, preferred + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind((HOST, port))
            except OSError:
                continue
            return port
    raise RuntimeError("No local port was available for the attendance app.")


def _open_browser_when_ready(port: int) -> None:
    time.sleep(2)
    webbrowser.open(f"http://{HOST}:{port}")


def _show_startup_error(message: str) -> None:
    if os.name == "nt":
        ctypes.windll.user32.MessageBoxW(0, message, APP_NAME, 0x10)
    else:
        print(message, file=sys.stderr)


def main() -> int:
    app_data = _prepare_runtime_paths()
    app_file = _bundled_path("app.py")
    if not app_file.exists():
        raise FileNotFoundError(f"Could not find bundled app.py at {app_file}")

    port = _find_available_port()
    if os.environ.get(NO_BROWSER_ENV_VAR) != "1":
        threading.Thread(target=_open_browser_when_ready, args=(port,), daemon=True).start()

    from streamlit.web import cli as streamlit_cli

    sys.argv = [
        "streamlit",
        "run",
        str(app_file),
        "--server.address",
        HOST,
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--server.enableCORS",
        "false",
        "--server.enableXsrfProtection",
        "false",
        "--server.fileWatcherType",
        "none",
        "--browser.gatherUsageStats",
        "false",
        "--global.developmentMode",
        "false",
    ]
    print(f"Starting {APP_NAME} on http://{HOST}:{port}")
    print("MongoDB storage is required; SQLite fallback is disabled for the EXE.")
    print(f"App data folder: {app_data}")
    return int(streamlit_cli.main() or 0)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        detail = (
            f"{APP_NAME} could not start.\n\n"
            f"{exc}\n\n"
            f"Check logs in: {_app_data_dir() / 'logs'}"
        )
        _show_startup_error(detail)
        raise
