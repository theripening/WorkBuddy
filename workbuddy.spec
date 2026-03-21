# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for WorkBuddy.
Build with:  pyinstaller workbuddy.spec
Output:      dist/WorkBuddy.exe
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Data files to bundle (source, dest-inside-bundle)
datas = [
    # App templates and migrations
    ("tickets/templates",   "tickets/templates"),
    ("tickets/migrations",  "tickets/migrations"),
    # Pre-collected static files (run collectstatic before building)
    ("staticfiles",         "staticfiles"),
    # Django's own templates (admin, auth, etc.)
    *collect_data_files("django", includes=["**/*.html"]),
]

hidden_imports = [
    # Django internals that PyInstaller misses
    "django.template.defaulttags",
    "django.template.defaultfilters",
    "django.template.loader_tags",
    "django.contrib.admin.apps",
    "django.contrib.auth.apps",
    "django.contrib.contenttypes.apps",
    "django.contrib.sessions.apps",
    "django.contrib.messages.apps",
    "django.contrib.staticfiles.apps",
    "django.core.management.commands.migrate",
    "django.core.management.commands.runserver",
    # App modules
    "workbuddy.settings",
    "workbuddy.urls",
    "workbuddy.wsgi",
    "tickets",
    "tickets.views",
    "tickets.models",
    "tickets.urls",
    "tickets.sync",
    "tickets.admin",
    "tickets.apps",
    # Whitenoise
    "whitenoise",
    "whitenoise.middleware",
    "whitenoise.storage",
    # All migration files
    *collect_submodules("tickets.migrations"),
    *collect_submodules("django.contrib.admin"),
    *collect_submodules("django.contrib.auth"),
    *collect_submodules("django.contrib.contenttypes"),
    *collect_submodules("django.contrib.sessions"),
    *collect_submodules("django.db.backends.sqlite3"),
    # pywin32 COM for Outlook integration
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="WorkBuddy",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,        # console=True so the server log is visible
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
