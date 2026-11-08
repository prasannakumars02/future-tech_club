import os
import secrets
from flask import Flask, current_app
from sqlalchemy import inspect, text
from .extensions import csrf, db, login_manager, migrate
from config import Config


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.from_object(Config)
    if test_config:
        app.config.update(test_config)
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(os.path.join(app.config["UPLOAD_FOLDER"], "videos"), exist_ok=True)
    os.makedirs(os.path.join(app.config["UPLOAD_FOLDER"], "chat-files"), exist_ok=True)
    os.makedirs(os.path.join(app.config["UPLOAD_FOLDER"], "task-files"), exist_ok=True)
    os.makedirs(os.path.join(app.config["UPLOAD_FOLDER"], "submissions"), exist_ok=True)
    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)
    migrate.init_app(app, db)

    from . import models
    from .routes import main
    app.register_blueprint(main)
    register_template_context(app)

    with app.app_context():
        db.create_all()
        ensure_member_schema()
        ensure_user_profile_schema()
        ensure_chat_schema()
        ensure_certificate_dispatch_schema()
        seed_initial_data()

    return app


def ensure_chat_schema():
    message_columns = {
        column["name"] for column in inspect(db.engine).get_columns("team_chat_message")
    }
    additions = {
        "attachment_storage_name": "VARCHAR(255)",
        "attachment_original_name": "VARCHAR(255)",
        "edited_at": "TIMESTAMP",
    }
    for column, definition in additions.items():
        if column not in message_columns:
            db.session.execute(
                text(f"ALTER TABLE team_chat_message ADD COLUMN {column} {definition}")
            )


def ensure_member_schema():
    member_columns = {
        column["name"] for column in inspect(db.engine).get_columns("team_member")
    }
    additions = {
        "department": "VARCHAR(100) NOT NULL DEFAULT ''",
        "year": "VARCHAR(30) NOT NULL DEFAULT ''",
        "user_id": "INTEGER REFERENCES user(id)",
    }
    for column, definition in additions.items():
        if column not in member_columns:
            db.session.execute(
                text(f"ALTER TABLE team_member ADD COLUMN {column} {definition}")
            )
    evaluation_columns = {
        column["name"] for column in inspect(db.engine).get_columns("evaluation")
    }
    if "approved" not in evaluation_columns:
        db.session.execute(
            text("ALTER TABLE evaluation ADD COLUMN approved BOOLEAN NOT NULL DEFAULT FALSE")
        )
    if "approved_by" not in evaluation_columns:
        db.session.execute(
            text("ALTER TABLE evaluation ADD COLUMN approved_by INTEGER")
        )
    if "approved_at" not in evaluation_columns:
        db.session.execute(
            text("ALTER TABLE evaluation ADD COLUMN approved_at TIMESTAMP")
        )
    from .models import Evaluation

    unattributed_approvals = Evaluation.query.filter_by(approved=True).filter(
        Evaluation.approved_by.is_(None),
        Evaluation.approved_at.is_(None),
    )
    for evaluation in unattributed_approvals:
        evaluation.approved = False
    db.session.commit()


def ensure_user_profile_schema():
    user_columns = {
        column["name"] for column in inspect(db.engine).get_columns("user")
    }
    additions = {
        "email": "VARCHAR(150) NOT NULL DEFAULT ''",
        "phone": "VARCHAR(40) NOT NULL DEFAULT ''",
    }
    for column, definition in additions.items():
        if column not in user_columns:
            db.session.execute(
                text(f"ALTER TABLE user ADD COLUMN {column} {definition}")
            )
    db.session.commit()


def ensure_certificate_dispatch_schema():
    certificate_columns = {
        column["name"] for column in inspect(db.engine).get_columns("certificate")
    }
    if "dispatch_status" not in certificate_columns:
        db.session.execute(
            text(
                "ALTER TABLE certificate ADD COLUMN dispatch_status "
                "VARCHAR(20) NOT NULL DEFAULT 'dispatched'"
            )
        )
    if "dispatched_at" not in certificate_columns:
        db.session.execute(
            text("ALTER TABLE certificate ADD COLUMN dispatched_at TIMESTAMP")
        )
    db.session.commit()


def register_template_context(app):
    from .rich_text import sanitize_rich_text

    app.add_template_filter(sanitize_rich_text, "rich_text")

    @app.context_processor
    def inject_site_branding():
        from flask_login import current_user
        from .models import SiteSetting
        settings = {setting.key: setting.value for setting in SiteSetting.query.all()}

        def branding_asset(setting_key, default_filename):
            filename = settings.get(setting_key, "")
            uploaded_path = os.path.join(
                app.config["UPLOAD_FOLDER"], "branding", filename
            )
            return (
                filename
                if filename and os.path.isfile(uploaded_path)
                else default_filename
            )

        unread_notifications = 0
        recent_notifications = []
        if current_user.is_authenticated:
            from .models import UserNotification
            unread_query = UserNotification.query.filter_by(
                user_id=current_user.id, read=False
            )
            unread_notifications = unread_query.count()
            recent_notifications = unread_query.order_by(
                UserNotification.created_at.desc()
            ).limit(3).all()
        return {
            "college_header": settings.get(
                "college_header", "N K Orchid College of Engineering and Technology"
            ) or "N K Orchid College of Engineering and Technology",
            "college_subtitle": settings.get("college_subtitle") or "Solapur",
            "club_logo": branding_asset("club_logo", "club_logo.jpg"),
            "college_logo": branding_asset("college_logo", "college_logo.png"),
            "college_banner_image": branding_asset(
                "college_banner_image", "header.png"
            ),
            "unread_notification_count": unread_notifications,
            "recent_notifications": recent_notifications,
            "notification_count_label": (
                "notification" if unread_notifications == 1 else "notifications"
            ),
        }


def seed_initial_data():
    from .models import Domain, Team, User
    from .extensions import db
    domains = [
        ("D01", "IoT & Robotics / Multimedia Group"),
        ("D02", "CodeWave Group"),
        ("D03", "Cyber Security Group"),
        ("D04", "Data Science Group"),
        ("D05", "Neural Techgenix (AI) Group"),
    ]
    for code, name in domains:
        if not Domain.query.filter_by(code=code).first():
            db.session.add(Domain(code=code, name=name))
    db.session.commit()
    domain_ids = {d.code: d.id for d in Domain.query.all()}
    for i in range(1, 28):
        code = f"FT_{i:02d}"
        team = Team.query.filter_by(team_code=code).first()
        if not team:
            team = Team(team_code=code, name=code, domain_id=None)
            db.session.add(team)
            db.session.flush()
        if not User.query.filter_by(username=code).first() and not User.query.filter_by(
            role="team_leader", team_id=team.id
        ).first():
            user = User(username=code, role="team_leader", display_name=code, team_id=team.id)
            user.set_password(
                current_app.config["TEAM_INITIAL_PASSWORD"] or secrets.token_urlsafe(9)
            )
            db.session.add(user)
    admin_username = current_app.config["BOOTSTRAP_ADMIN_USERNAME"]
    admin = User.query.filter_by(username=admin_username).first()
    if admin and admin.role != "admin":
        raise RuntimeError(
            f"Configured admin username {admin_username!r} belongs to a non-admin account."
        )

    legacy_admin = (
        User.query.filter_by(username="admin").first()
        if admin_username != "admin"
        else None
    )
    if admin is None and legacy_admin:
        bootstrap_password = current_app.config["BOOTSTRAP_ADMIN_PASSWORD"]
        if not bootstrap_password:
            raise RuntimeError(
                "Set BOOTSTRAP_ADMIN_PASSWORD before migrating the legacy admin account."
            )
        admin = legacy_admin
        admin.username = admin_username
        admin.role = "admin"
        admin.set_password(bootstrap_password)
    elif admin is None:
        bootstrap_password = current_app.config["BOOTSTRAP_ADMIN_PASSWORD"]
        if not bootstrap_password:
            raise RuntimeError(
                "Set BOOTSTRAP_ADMIN_PASSWORD before creating the admin account."
            )
        admin = User(
            username=admin_username,
            role="admin",
            display_name="Super Admin",
        )
        admin.set_password(bootstrap_password)
        db.session.add(admin)
    elif legacy_admin:
        legacy_admin.active = False
    from .models import CertificateType

    for name in (
        "Certificate of Participation",
        "Certificate of Winner",
        "Certificate of Appreciation - Coordinator",
        "Certificate of Appreciation - Mentor",
        "Certificate of Appreciation - Club Head and Judges",
    ):
        if not CertificateType.query.filter_by(name=name).first():
            db.session.add(CertificateType(name=name))
    db.session.commit()
