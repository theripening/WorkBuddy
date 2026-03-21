"""
WorkBuddy launcher.
Starts the Django server and opens a browser tab.
This is the PyInstaller entry point — also works as a normal Python script.
"""
import os
import sys
import threading
import time
import webbrowser

# When frozen by PyInstaller:
#   sys.frozen = True
#   sys._MEIPASS = temp dir where bundled files are extracted (read-only)
#   sys.executable = path to WorkBuddy.exe
#
# User data (db.sqlite3) lives next to the .exe so it persists across updates.
# App code (templates, migrations, etc.) lives in _MEIPASS.

_FROZEN = getattr(sys, "frozen", False)

if _FROZEN:
    _app_dir = sys._MEIPASS
    _data_dir = os.path.dirname(sys.executable)
    sys.path.insert(0, _app_dir)
else:
    _app_dir = os.path.dirname(os.path.abspath(__file__))
    _data_dir = _app_dir

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "workbuddy.settings")
os.environ["WORKBUDDY_DATA_DIR"] = _data_dir

import django  # noqa: E402 — must come after sys.path and env setup
from django.core.management import call_command  # noqa: E402

PORT = 8000


def _run_server():
    call_command("runserver", f"127.0.0.1:{PORT}", "--noreload")


if __name__ == "__main__":
    django.setup()

    # Auto-run migrations on startup (no-op if already up to date)
    call_command("migrate", verbosity=0)

    server_thread = threading.Thread(target=_run_server, daemon=True)
    server_thread.start()

    # Brief pause so the server is ready before the browser opens
    time.sleep(2)
    webbrowser.open(f"http://127.0.0.1:{PORT}/")

    print(f"WorkBuddy is running at http://127.0.0.1:{PORT}/")
    print("Close this window to stop the server.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
