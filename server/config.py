import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
INSTANCE_DIR.mkdir(parents=True, exist_ok=True)

_default_sqlite_path = (INSTANCE_DIR / "app.db").as_posix()
_default_db_url = f"sqlite:///{_default_sqlite_path}"

DEFAULT_SECRET = "dev-only-not-for-production"
DEFAULT_ADMIN_PASSWORD = "changeme"


def _bool(env: str, default: bool = False) -> bool:
    return os.environ.get(env, str(default)).strip().lower() in ("1", "true", "yes", "on")


class Config:
    SECRET_KEY = os.environ.get("FLASK_SECRET_KEY", DEFAULT_SECRET)
    SQLALCHEMY_DATABASE_URI = os.environ.get("DATABASE_URL") or _default_db_url
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    SCRIPTS_DIR = Path(
        os.environ.get("SCRIPTS_DIR", BASE_DIR.parent)
    ).resolve()

    JOBS_DIR = BASE_DIR / "jobs"
    JOB_RETENTION_HOURS = int(os.environ.get("JOB_RETENTION_HOURS", "24"))
    SESSION_HOURS = int(os.environ.get("SESSION_HOURS", "8"))

    ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
    ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)

    MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # 200 MB upload cap

    # --- Security hardening ---
    # Secure cookies require HTTPS. Render provides HTTPS automatically;
    # set FORCE_HTTPS=true on production to enable Secure flag.
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    SESSION_COOKIE_SECURE = _bool("FORCE_HTTPS", False)
    WTF_CSRF_TIME_LIMIT = None  # CSRF token valid for session lifetime

    # Login brute-force defense
    LOGIN_RATE_LIMIT = os.environ.get("LOGIN_RATE_LIMIT", "5 per minute")
    MAX_FAILED_LOGINS = int(os.environ.get("MAX_FAILED_LOGINS", "5"))
    LOCKOUT_MINUTES = int(os.environ.get("LOCKOUT_MINUTES", "15"))

    # Postgres tweaks
    if SQLALCHEMY_DATABASE_URI.startswith("postgres://"):
        # Render/Heroku hand out the old scheme; SQLAlchemy wants postgresql://
        SQLALCHEMY_DATABASE_URI = SQLALCHEMY_DATABASE_URI.replace(
            "postgres://", "postgresql://", 1
        )

    # Connection-pool resilience for serverless Postgres (Neon).
    # Neon silently kills idle TCP/SSL connections after a few minutes; without
    # pre-ping, the pool hands out dead connections and the next query 500s with
    # "SSL connection has been closed unexpectedly" (seen 2026-04-30 in prod).
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,   # SELECT 1 before checkout; replace dead connections silently
        "pool_recycle": 280,     # recycle every <5 min, beats Neon's idle-timeout
        "pool_size": 5,
        "max_overflow": 5,
    }
