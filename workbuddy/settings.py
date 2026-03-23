import os
import sys
from pathlib import Path

# When running as a PyInstaller bundle:
#   - App files (templates, migrations) live in sys._MEIPASS (read-only temp dir)
#   - User data (db.sqlite3) lives in WORKBUDDY_DATA_DIR (next to the .exe)
_FROZEN = getattr(sys, "frozen", False)
BASE_DIR = Path(sys._MEIPASS) if _FROZEN else Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ["WORKBUDDY_DATA_DIR"]) if "WORKBUDDY_DATA_DIR" in os.environ else BASE_DIR

SECRET_KEY = "workbuddy-local-dev-key-replace-before-sharing"

DEBUG = True

ALLOWED_HOSTS = ["localhost", "127.0.0.1"]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "tickets",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "workbuddy.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

WSGI_APPLICATION = "workbuddy.wsgi.application"

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": DATA_DIR / "db.sqlite3",
    }
}

# Change to your local timezone, e.g. "America/Chicago", "America/Los_Angeles"
TIME_ZONE = "America/New_York"
LANGUAGE_CODE = "en-us"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"
STATIC_ROOT = BASE_DIR / "staticfiles"

# Whitenoise serves collected static files when running as a frozen bundle
if _FROZEN:
    MIDDLEWARE.insert(1, "whitenoise.middleware.WhiteNoiseMiddleware")

LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,
    "formatters": {
        "simple": {"format": "%(levelname)s %(name)s: %(message)s"},
    },
    "handlers": {
        "console": {
            "class": "logging.StreamHandler",
            "formatter": "simple",
        },
    },
    "loggers": {
        "tickets.sync": {
            "handlers": ["console"],
            "level": "DEBUG",
            "propagate": False,
        },
    },
}
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# Days before an open/in-progress ticket with no activity is considered stale
STALE_DAYS = 7

# WorkBuddyCloud shared ticket API (Phase 2 — optional).
# Set to your WorkBuddyCloud server URL to enable cloud sync.
# e.g. WORKBUDDY_CLOUD_URL = "http://192.168.1.50:8765"
# Leave as None to run in local-only mode (no cloud calls made).
WORKBUDDY_CLOUD_URL = None

