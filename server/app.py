from __future__ import annotations

import csv
import io
import shutil
import time
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_login import (
    LoginManager,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from flask_wtf.csrf import CSRFProtect, CSRFError
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from sqlalchemy import desc

from config import Config, DEFAULT_ADMIN_PASSWORD, DEFAULT_SECRET
from models import AuditLog, User, db
from pipelines import PipelineError, run_dxf_pipeline, run_node_pipeline


csrf = CSRFProtect()
limiter = Limiter(key_func=get_remote_address, default_limits=[])


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)
    app.permanent_session_lifetime = timedelta(hours=Config.SESSION_HOURS)

    db.init_app(app)
    csrf.init_app(app)
    # CSRF is enforced on form POSTs; file-upload API endpoints are exempted
    # individually below (they authenticate via session and accept multipart).
    limiter.init_app(app)

    login_manager = LoginManager(app)
    login_manager.login_view = "login"

    @login_manager.user_loader
    def load_user(user_id: str):
        return db.session.get(User, int(user_id))

    Config.JOBS_DIR.mkdir(parents=True, exist_ok=True)

    with app.app_context():
        db.create_all()
        _seed_admin(app)
        _startup_warnings(app)

    _register_routes(app)
    return app


def _startup_warnings(app: Flask) -> None:
    if Config.SECRET_KEY == DEFAULT_SECRET:
        app.logger.warning(
            "!!! FLASK_SECRET_KEY is the default value — set a strong random value in production."
        )
    if Config.ADMIN_PASSWORD == DEFAULT_ADMIN_PASSWORD:
        app.logger.warning(
            "!!! ADMIN_PASSWORD is the default 'changeme' — change it in .env (local) or Render env vars (production)."
        )
    if not app.config.get("SESSION_COOKIE_SECURE"):
        app.logger.info(
            "SESSION_COOKIE_SECURE is OFF (HTTP-only dev mode). Set FORCE_HTTPS=true in production."
        )


def _seed_admin(app: Flask) -> None:
    if User.query.filter_by(username=Config.ADMIN_USERNAME).first():
        return
    admin = User(username=Config.ADMIN_USERNAME, role="admin")
    admin.set_password(Config.ADMIN_PASSWORD)
    db.session.add(admin)
    db.session.commit()
    app.logger.info("Seeded admin user '%s'", Config.ADMIN_USERNAME)


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not current_user.is_authenticated:
            return redirect(url_for("login"))
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return wrapper


def _log(
    action: str,
    success: bool = True,
    files: list[str] | None = None,
    duration: float | None = None,
    details: str = "",
) -> None:
    entry = AuditLog(
        username=(current_user.username if current_user.is_authenticated else None),
        action=action,
        files=",".join(files) if files else None,
        duration_seconds=duration,
        success=success,
        details=details[:2000] if details else None,
        ip_address=request.remote_addr,
    )
    db.session.add(entry)
    db.session.commit()


def _cleanup_old_jobs() -> None:
    cutoff = datetime.now(timezone.utc) - timedelta(hours=Config.JOB_RETENTION_HOURS)
    for p in Config.JOBS_DIR.iterdir():
        if not p.is_dir():
            continue
        mtime = datetime.fromtimestamp(p.stat().st_mtime, tz=timezone.utc)
        if mtime < cutoff:
            shutil.rmtree(p, ignore_errors=True)


def _new_job_dir() -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    d = Config.JOBS_DIR / f"{stamp}_{uuid.uuid4().hex[:8]}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _zip_outputs(outputs: list[Path], zip_name: str) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in outputs:
            zf.write(p, arcname=p.name)
    buf.seek(0)
    buf.name = zip_name  # hint for callers
    return buf


def _register_routes(app: Flask) -> None:

    @app.before_request
    def make_session_permanent():
        session.permanent = True

    @app.route("/login", methods=["GET", "POST"])
    @limiter.limit(Config.LOGIN_RATE_LIMIT, methods=["POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(url_for("home"))
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = User.query.filter_by(username=username).first()

            if user and user.is_locked_out():
                _log(
                    "login", success=False,
                    details=f"locked_out username={username!r}",
                )
                flash(
                    "This account is temporarily locked due to too many failed attempts. Try again later.",
                    "error",
                )
                return render_template("login.html"), 429

            if user and user.is_active and user.check_password(password):
                user.register_successful_login()
                db.session.commit()
                login_user(user, remember=False)
                _log("login", success=True)
                return redirect(url_for("home"))

            if user:
                user.register_failed_login(
                    Config.MAX_FAILED_LOGINS, Config.LOCKOUT_MINUTES
                )
                db.session.commit()
            _log("login", success=False, details=f"username={username!r}")
            flash("Invalid username or password.", "error")
        return render_template("login.html")

    @app.route("/logout")
    @login_required
    def logout():
        _log("logout")
        logout_user()
        return redirect(url_for("login"))

    @app.route("/")
    @login_required
    def home():
        return render_template("home.html")

    @app.route("/dxf", methods=["GET"])
    @login_required
    def dxf_page():
        return render_template("dxf.html")

    @app.route("/node", methods=["GET"])
    @login_required
    def node_page():
        return render_template("node.html")

    @app.route("/api/run/dxf", methods=["POST"])
    @login_required
    @csrf.exempt
    def api_run_dxf():
        _cleanup_old_jobs()
        files = request.files.getlist("files")
        dxfs = [f for f in files if f.filename.lower().endswith(".dxf")]
        if not dxfs:
            return jsonify(ok=False, error="No DXF files uploaded."), 400

        job_dir = _new_job_dir()
        saved = []
        for f in dxfs:
            dest = job_dir / Path(f.filename).name
            f.save(dest)
            saved.append(dest.name)

        started = time.time()
        try:
            result = run_dxf_pipeline(Config.SCRIPTS_DIR, job_dir)
        except PipelineError as e:
            _log(
                "dxf_run", success=False, files=saved,
                duration=time.time() - started,
                details=f"{e}\nSTDERR:\n{e.stderr[-2000:]}",
            )
            return jsonify(
                ok=False, error=str(e), stdout=e.stdout[-4000:], stderr=e.stderr[-4000:]
            ), 500

        output_names = [p.name for p in result["outputs"]]
        _log(
            "dxf_run", success=True, files=saved,
            duration=time.time() - started,
            details=f"outputs={output_names}",
        )
        return jsonify(
            ok=True,
            job_id=job_dir.name,
            outputs=output_names,
            stdout=result["stdout"][-4000:],
        )

    @app.route("/api/run/node", methods=["POST"])
    @login_required
    @csrf.exempt
    def api_run_node():
        _cleanup_old_jobs()
        files = request.files.getlist("files")
        xlsxs = [f for f in files if f.filename.lower().endswith(".xlsx")]
        if not xlsxs:
            return jsonify(ok=False, error="No Excel files uploaded."), 400

        job_dir = _new_job_dir()
        saved = []
        for f in xlsxs:
            dest = job_dir / Path(f.filename).name
            f.save(dest)
            saved.append(dest.name)

        started = time.time()
        try:
            result = run_node_pipeline(Config.SCRIPTS_DIR, job_dir)
        except PipelineError as e:
            _log(
                "node_run", success=False, files=saved,
                duration=time.time() - started,
                details=f"{e}\nSTDERR:\n{e.stderr[-2000:]}",
            )
            return jsonify(
                ok=False, error=str(e), stdout=e.stdout[-4000:], stderr=e.stderr[-4000:]
            ), 500

        output_names = [p.name for p in result["outputs"]]
        _log(
            "node_run", success=True, files=saved,
            duration=time.time() - started,
            details=f"outputs={output_names}",
        )
        return jsonify(
            ok=True,
            job_id=job_dir.name,
            outputs=output_names,
            stdout=result["stdout"][-4000:],
        )

    @app.route("/api/download/<job_id>/<path:filename>")
    @login_required
    def api_download_file(job_id: str, filename: str):
        job_dir = (Config.JOBS_DIR / job_id).resolve()
        if not str(job_dir).startswith(str(Config.JOBS_DIR.resolve())):
            abort(403)
        target = (job_dir / filename).resolve()
        if not str(target).startswith(str(job_dir)):
            abort(403)
        if not target.exists():
            abort(404)
        _log("download", files=[filename])
        return send_file(target, as_attachment=True, download_name=target.name)

    @app.route("/api/download-zip/<job_id>")
    @login_required
    def api_download_zip(job_id: str):
        job_dir = (Config.JOBS_DIR / job_id).resolve()
        if not str(job_dir).startswith(str(Config.JOBS_DIR.resolve())):
            abort(403)
        if not job_dir.exists():
            abort(404)
        outputs = sorted(job_dir.glob("*.xlsx"))
        buf = _zip_outputs(outputs, f"{job_id}.zip")
        _log("download_zip", files=[p.name for p in outputs])
        return send_file(
            buf, mimetype="application/zip",
            as_attachment=True, download_name=f"outputs_{job_id}.zip"
        )

    # ---------- Admin ----------

    @app.route("/admin")
    @admin_required
    def admin_home():
        return redirect(url_for("admin_users"))

    @app.route("/admin/users", methods=["GET", "POST"])
    @admin_required
    def admin_users():
        if request.method == "POST":
            action = request.form.get("action")
            if action == "create":
                uname = request.form.get("username", "").strip()
                pw = request.form.get("password", "")
                role = request.form.get("role", "user")
                if not uname or not pw:
                    flash("Username and password required.", "error")
                elif User.query.filter_by(username=uname).first():
                    flash("Username already exists.", "error")
                else:
                    u = User(username=uname, role=role)
                    u.set_password(pw)
                    db.session.add(u)
                    db.session.commit()
                    _log("admin_create_user", files=[uname])
                    flash(f"User '{uname}' created.", "success")
            elif action == "reset":
                uid = int(request.form.get("user_id"))
                pw = request.form.get("new_password", "")
                u = db.session.get(User, uid)
                if u and pw:
                    u.set_password(pw)
                    db.session.commit()
                    _log("admin_reset_password", files=[u.username])
                    flash(f"Password reset for '{u.username}'.", "success")
            elif action == "toggle_disabled":
                uid = int(request.form.get("user_id"))
                u = db.session.get(User, uid)
                if u and u.id != current_user.id:
                    u.disabled = not u.disabled
                    db.session.commit()
                    _log(
                        "admin_toggle_disabled", files=[u.username],
                        details=f"disabled={u.disabled}"
                    )
            elif action == "delete":
                uid = int(request.form.get("user_id"))
                u = db.session.get(User, uid)
                if u and u.id != current_user.id:
                    uname = u.username
                    db.session.delete(u)
                    db.session.commit()
                    _log("admin_delete_user", files=[uname])
                    flash(f"User '{uname}' deleted.", "success")
            return redirect(url_for("admin_users"))

        users = User.query.order_by(User.username).all()
        return render_template("admin/users.html", users=users)

    @app.route("/admin/logs")
    @admin_required
    def admin_logs():
        q = AuditLog.query
        username = request.args.get("username", "").strip()
        action = request.args.get("action", "").strip()
        if username:
            q = q.filter(AuditLog.username == username)
        if action:
            q = q.filter(AuditLog.action == action)
        logs = q.order_by(desc(AuditLog.timestamp)).limit(500).all()
        users = [u.username for u in User.query.order_by(User.username).all()]
        actions = [
            "login", "logout", "dxf_run", "node_run", "download", "download_zip",
            "admin_create_user", "admin_reset_password",
            "admin_toggle_disabled", "admin_delete_user",
        ]
        return render_template(
            "admin/logs.html",
            logs=logs, users=users, actions=actions,
            username=username, action=action,
        )

    @app.route("/admin/logs.csv")
    @admin_required
    def admin_logs_csv():
        logs = AuditLog.query.order_by(desc(AuditLog.timestamp)).limit(10000).all()
        out = io.StringIO()
        w = csv.writer(out)
        w.writerow([
            "timestamp_utc", "username", "action", "success",
            "duration_seconds", "files", "ip_address", "details"
        ])
        for log in logs:
            w.writerow([
                log.timestamp.isoformat() if log.timestamp else "",
                log.username or "",
                log.action,
                log.success,
                log.duration_seconds if log.duration_seconds is not None else "",
                log.files or "",
                log.ip_address or "",
                (log.details or "").replace("\n", " "),
            ])
        buf = io.BytesIO(out.getvalue().encode("utf-8"))
        buf.seek(0)
        return send_file(
            buf, mimetype="text/csv",
            as_attachment=True, download_name="audit_logs.csv"
        )

    @app.errorhandler(403)
    def forbidden(_):
        return render_template("error.html", code=403, message="Forbidden"), 403

    @app.errorhandler(404)
    def not_found(_):
        return render_template("error.html", code=404, message="Not found"), 404

    @app.errorhandler(429)
    def too_many_requests(e):
        return render_template(
            "error.html", code=429,
            message="Too many attempts. Please wait a minute and try again.",
        ), 429

    @app.errorhandler(CSRFError)
    def csrf_error(e):
        return render_template(
            "error.html", code=400,
            message="Security token expired or invalid. Refresh the page and try again.",
        ), 400


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
