from datetime import datetime, timezone
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
import bcrypt

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(128), nullable=False)
    role = db.Column(db.String(16), nullable=False, default="user")  # 'admin' | 'user'
    disabled = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))
    last_login_at = db.Column(db.DateTime)
    failed_login_count = db.Column(db.Integer, default=0, nullable=False)
    locked_until = db.Column(db.DateTime)

    def set_password(self, plaintext: str) -> None:
        self.password_hash = bcrypt.hashpw(
            plaintext.encode("utf-8"), bcrypt.gensalt()
        ).decode("utf-8")

    def check_password(self, plaintext: str) -> bool:
        try:
            return bcrypt.checkpw(
                plaintext.encode("utf-8"), self.password_hash.encode("utf-8")
            )
        except ValueError:
            return False

    @property
    def is_active(self) -> bool:
        return not self.disabled

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"

    def is_locked_out(self) -> bool:
        if not self.locked_until:
            return False
        return datetime.now(timezone.utc) < self.locked_until.replace(
            tzinfo=timezone.utc
        ) if self.locked_until.tzinfo is None else datetime.now(timezone.utc) < self.locked_until

    def register_failed_login(self, max_fails: int, lockout_minutes: int) -> None:
        self.failed_login_count = (self.failed_login_count or 0) + 1
        if self.failed_login_count >= max_fails:
            from datetime import timedelta
            self.locked_until = datetime.now(timezone.utc) + timedelta(
                minutes=lockout_minutes
            )

    def register_successful_login(self) -> None:
        self.failed_login_count = 0
        self.locked_until = None
        self.last_login_at = datetime.now(timezone.utc)


class AuditLog(db.Model):
    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), index=True)
    action = db.Column(db.String(32), nullable=False, index=True)
    timestamp = db.Column(
        db.DateTime, default=lambda: datetime.now(timezone.utc), index=True
    )
    files = db.Column(db.Text)  # comma-joined filenames
    duration_seconds = db.Column(db.Float)
    success = db.Column(db.Boolean, default=True)
    details = db.Column(db.Text)
    ip_address = db.Column(db.String(64))
