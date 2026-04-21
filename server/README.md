# StructAuto Server

Flask web app that runs the DXF + Node pipelines on a shared server so clients only need a browser.

## Local quickstart (Windows, PowerShell or CMD)

```bash
cd "D:\autocad automation\server"

# 1. Create venv and install deps
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# 2. Create .env (edit values)
copy .env.example .env
notepad .env

# 3. Run
python app.py
```

Open http://localhost:8000 — login with `admin` / `changeme` (from `.env`).

## .env settings

| Key | Purpose |
|---|---|
| `FLASK_SECRET_KEY` | Session cookie signing key. Generate any random string. |
| `SCRIPTS_DIR` | Path to the folder containing `dxf_columns_walls_pipeline.py`, `Unfiltered column coordinates generator.py`, `Appended node coordinate generator.py`. |
| `DATABASE_URL` | Default SQLite at `instance/app.db`. Switch to Postgres for production (`postgresql://user:pass@host/db`). |
| `ADMIN_USERNAME` / `ADMIN_PASSWORD` | First-run admin account. Change password via admin panel after login. |
| `SESSION_HOURS` | How long a login session stays active. |
| `JOB_RETENTION_HOURS` | Old `jobs/<id>/` folders get deleted after this many hours. |

## What it does

- **Login** — bcrypt-hashed passwords, session cookies.
- **DXF page** — upload 3 DXFs → runs `dxf_columns_walls_pipeline.py` → returns column + wall Excels.
- **Node page** — upload 5 Excels → runs Unfiltered + Appended → returns node_coordinates + other_coordinates + column_beam_pairs.
- **File System Access API** — on Chrome/Edge, outputs auto-save to a folder the user picks. Other browsers fall back to Downloads.
- **Admin panel** (`/admin`) — create / reset / disable / delete users. View + filter + export audit logs.

## Audit log actions

`login`, `logout`, `dxf_run`, `node_run`, `download`, `download_zip`,
`admin_create_user`, `admin_reset_password`, `admin_toggle_disabled`, `admin_delete_user`.

## Security hardening

Built-in defenses (active by default):

- **Bcrypt** password hashing.
- **CSRF** tokens on every POST form (Flask-WTF). File-upload API endpoints are session-authenticated and CSRF-exempt.
- **Rate limit** on `/login` — `LOGIN_RATE_LIMIT` (default `5 per minute` per IP).
- **Account lockout** — `MAX_FAILED_LOGINS` fails (default 5) → locked for `LOCKOUT_MINUTES` (default 15). Successful login resets the counter.
- **Session cookies** — `HttpOnly` + `SameSite=Lax`. `Secure` flag turns on when `FORCE_HTTPS=true`.
- **Startup warnings** — loud warning if `ADMIN_PASSWORD` or `FLASK_SECRET_KEY` is still the default.

Before deploying to production:

1. `FLASK_SECRET_KEY` — a long random string (`python -c "import secrets; print(secrets.token_urlsafe(48))"`).
2. `ADMIN_PASSWORD` — set a strong password.
3. `FORCE_HTTPS=true` — Render provides HTTPS automatically.
4. `DATABASE_URL=postgresql://...` — use Neon Postgres instead of SQLite (SQLite is wiped on each Render redeploy).
