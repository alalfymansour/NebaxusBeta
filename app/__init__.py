from __future__ import annotations

import os
from datetime import datetime

import pytz
from flask import Flask, g, request, session
from flask_login import LoginManager
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text

from config import ConfigurationError, config

db = SQLAlchemy()
migrate = Migrate()
login_manager = LoginManager()


def create_app(config_name: str | None = None) -> Flask:
    """
    Application factory.

    Args:
        config_name: Key from ConfigRegistry. Defaults to FLASK_ENV env var,
                     falling back to 'default'.
    """
    app = Flask(__name__)
    _load_config(app, config_name)
    _init_extensions(app)
    _check_db_connection(app)
    _run_startup_data_fixes(app)
    _seed_admin_user(app)
    _register_context_processors(app)
    _register_hooks(app)
    _register_blueprints(app)
    return app


def _load_config(app: Flask, config_name: str | None) -> None:
    """Load and validate configuration from ConfigRegistry."""
    resolved_config_name = config_name or os.environ.get("FLASK_ENV") or "default"

    try:
        config_class = config[resolved_config_name]
    except KeyError as error:
        available_keys = ", ".join(sorted(config.keys()))
        message = (
            f"Application startup failed: unknown config '{resolved_config_name}'. "
            f"Available keys: {available_keys}"
        )
        app.logger.critical(message, exc_info=True)
        raise ConfigurationError(message) from error
    except ConfigurationError as error:
        message = (
            f"Application startup failed while loading '{resolved_config_name}' "
            f"configuration: {error}"
        )
        app.logger.critical(message, exc_info=True)
        raise ConfigurationError(message) from error

    app.config.from_object(config_class)

    # Preserve legacy static file behavior served from repository-level ./static.
    app.static_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "static"))

    # Keep historical engine defaults in case any config class omits an option.
    engine_options = dict(app.config.get("SQLALCHEMY_ENGINE_OPTIONS", {}))
    engine_defaults = {
        "pool_pre_ping": True,
        "pool_recycle": 280,
        "pool_size": 5,
        "max_overflow": 5,
    }
    for option, default_value in engine_defaults.items():
        engine_options.setdefault(option, default_value)
    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = engine_options

    app.logger.info("Loaded configuration class: %s", config_class.__name__)


def _init_extensions(app: Flask) -> None:
    """Initialize all Flask extensions."""
    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = "main.login"
    login_manager.login_message_category = "warning"

    @login_manager.user_loader
    def load_user(user_id: str):
        from .models import Employee

        return Employee.query.get(int(user_id))

    # Flask-Migrate depends on SQLAlchemy metadata, so db must be initialized first.
    migrate.init_app(app, db)


def _check_db_connection(app: Flask) -> None:
    """Verify database is reachable before serving requests."""
    with app.app_context():
        try:
            db.session.execute(text("SELECT 1"))
            app.logger.info("Database connection established")
        except Exception as e:
            app.logger.critical("Database unreachable: %s", e, exc_info=True)
            raise


def _run_startup_data_fixes(app: Flask) -> None:
    """
    Apply one-time data corrections that run on every boot.

    NOTE: Each fix here must have a comment explaining:
    - WHY it cannot be a migration
    - WHAT data condition it corrects
    - WHEN it can be safely removed
    """
    with app.app_context():
        try:
            # WHY: legacy plaintext values may reappear after manual data restores,
            # so we enforce cleanup at boot rather than relying on migration history.
            # WHAT: removes non-null values from employee.password_plain.
            # WHEN: safe to remove once password_plain is permanently dropped everywhere.
            scrub_result = db.session.execute(
                text("UPDATE employee SET password_plain = NULL WHERE password_plain IS NOT NULL")
            )
            db.session.commit()
            if (scrub_result.rowcount or 0) > 0:
                app.logger.warning(
                    "Cleared legacy plaintext passwords for %s employee rows.",
                    scrub_result.rowcount,
                )
        # ponytail: column may not exist (sqlite), just skip
        except Exception:
            db.session.rollback()


def _seed_admin_user(app: Flask) -> None:
    """Create the default admin user if no admin exists."""
    with app.app_context():
        if os.environ.get("SKIP_STARTUP_SEED") == "1":
            app.logger.warning("⏭️ Skipping default user creation (SKIP_STARTUP_SEED=1)")
            return

        if not _ensure_employee_schema(app):
            return

        from .models import Employee

        try:
            existing_admin = Employee.query.filter_by(is_admin=True).first()
        except Exception:
            app.logger.warning(
                "Could not query Employee table (migration pending?). "
                "Skipping admin seed."
            )
            return

        if existing_admin:
            app.logger.debug("Admin user already exists, skipping seed")
            return

        admin_username = os.environ.get("ADMIN_USERNAME")
        admin_password = os.environ.get("ADMIN_PASSWORD")
        missing = [
            var_name
            for var_name, value in {
                "ADMIN_USERNAME": admin_username,
                "ADMIN_PASSWORD": admin_password,
            }.items()
            if not value
        ]

        if missing:
            if _is_development_environment(app):
                app.logger.warning(
                    "Skipping admin seed in development due to missing env var(s): %s",
                    ", ".join(missing),
                )
                return
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing)}"
            )

        try:
            admin = Employee(
                name="مدير النظام",
                username=admin_username,
                position="مدير",
                phone="01000000000",
                salary=5000,
                is_admin=True,
            )
            admin.set_password(admin_password)
            db.session.add(admin)
            db.session.commit()
            app.logger.info("Admin user seeded successfully")
        except Exception as e:
            db.session.rollback()
            app.logger.error("Failed to seed admin user: %s", e, exc_info=True)
            raise


def _is_development_environment(app: Flask) -> bool:
    """Determine whether the current runtime should be treated as development."""
    env_name = str(app.config.get("ENV") or "").strip().lower()
    return env_name == "development" or bool(app.debug)


def _ensure_employee_schema(app: Flask) -> bool:
    """Validate employee schema and attempt auto-upgrade when required."""
    # Auto-migration is unsafe under multi-worker servers (e.g. gunicorn with 4 workers)
    # because all workers run this check concurrently on startup, risking race conditions.
    # Production deployments should set DISABLE_AUTO_MIGRATE=1 and run flask db upgrade manually.
    if os.environ.get("DISABLE_AUTO_MIGRATE") == "1":
        app.logger.debug(
            "Auto-migrate disabled (DISABLE_AUTO_MIGRATE=1). Skipping schema drift check."
        )
        return True

    inspector = inspect(db.engine)
    table_names = inspector.get_table_names()
    if "employee" not in table_names:
        app.logger.warning("⚠️ Employee table missing — creating tables...")
        try:
            from .models import Employee
            db.create_all()
            app.logger.info("✅ Tables created.")
        except Exception as e:
            app.logger.error("❌ Table creation failed: %s", e, exc_info=True)
            return False
        return True



def _register_context_processors(app: Flask) -> None:
    """Register Jinja2 context processors for template globals."""

    @app.context_processor
    def inject_permission_helpers():
        """Inject permission helper callables into all templates."""
        try:
            from .permissions import has_all, has_any, has_permission
        except ImportError as e:
            app.logger.error("Failed to load permission context: %s", e, exc_info=True)
            raise
        return {
            "has_permission": has_permission,
            "has_any": has_any,
            "has_all": has_all,
        }

    @app.context_processor
    def inject_current_year():
        """Inject current year and datetime for templates."""
        return {"current_year": datetime.now().year, "datetime": datetime}

    @app.context_processor
    def inject_trial_info():
        if os.environ.get("NEBAXUS_MODE") == "trial":
            from .trial import inject_trial_context
            return inject_trial_context()
        return {}

    @app.context_processor
    def inject_pending_followups():
        try:
            from .models import FollowUp, db
            from sqlalchemy import func
            count = (db.session.query(func.count(FollowUp.id))
                     .filter(FollowUp.status == 'قائمة')
                     .scalar() or 0)
        except Exception:
            count = 0
        return {"pending_followups_count": count}

    @app.template_filter("localtime")
    def localtime_filter(value, tz_name: str = "Africa/Cairo"):
        """Convert naive/UTC datetimes into the requested local timezone."""
        if value is None:
            return ""
        utc = pytz.utc
        local_tz = pytz.timezone(tz_name)
        if value.tzinfo is None:
            # Existing rows may contain naive UTC timestamps from historical inserts.
            value = utc.localize(value)
        return value.astimezone(local_tz)

    @app.context_processor
    def inject_breadcrumbs():
        return {'breadcrumbs': session.get('_nav', [])}


def _register_hooks(app: Flask) -> None:
    @app.before_request
    def track_nav_history():
        if request.method != 'GET':
            return
        if 'employee_id' not in session:
            return
        endpoint = request.endpoint
        if not endpoint:
            return
        skip = ('main.login', 'main.logout', 'main.splash', 'main.splash_only',
                'main.trigger_backup', 'main.trial_stats')
        if endpoint in skip:
            return
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return
        if request.is_json:
            return

        from .breadcrumbs import page_label
        label, icon = page_label(endpoint, **request.view_args or {})
        if not label:
            return

        history = session.get('_nav', [])
        path = request.path

        # لو المسار موجود مسبقاً — نشيل كل اللي بعده (زي browser back)
        for i, (_, p, _) in enumerate(history):
            if p == path:
                history = history[:i + 1]
                session['_nav'] = history
                return

        history.append((label, path, icon))
        if len(history) > 10:
            history = history[-10:]

        session['_nav'] = history



def _register_blueprints(app: Flask) -> None:
    """Register all application blueprints."""
    from .routes import main

    app.register_blueprint(main)  # Main domain: auth, orders, products, suppliers, and reporting.


