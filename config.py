import logging
import os
from urllib.parse import quote_plus

from dotenv import load_dotenv

load_dotenv()
logger = logging.getLogger(__name__)

# Number of persistent PostgreSQL connections kept open in the pool.
SQLALCHEMY_POOL_SIZE = 5
# Number of temporary overflow connections allowed beyond pool_size.
SQLALCHEMY_MAX_OVERFLOW = 5
# Seconds before recycling pooled connections to avoid stale DB sessions.
SQLALCHEMY_POOL_RECYCLE = 280


class ConfigurationError(RuntimeError):
    """Raised when required environment configuration is missing."""


def _missing_env_vars(*variable_names):
    return [name for name in variable_names if not os.environ.get(name)]


class BaseConfig:
    DEBUG = False
    TESTING = False
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    SECRET_KEY = None
    SQLALCHEMY_DATABASE_URI = None
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_size": SQLALCHEMY_POOL_SIZE,
        "max_overflow": SQLALCHEMY_MAX_OVERFLOW,
        "pool_recycle": SQLALCHEMY_POOL_RECYCLE,
    }

    @classmethod
    def _build_database_uri(cls, require_database_url: bool = False):
        database_url = os.environ.get("DATABASE_URL")

        if database_url:
            if database_url.startswith("postgres://"):
                database_url = database_url.replace("postgres://", "postgresql://", 1)
                logger.debug("Normalized DATABASE_URL scheme from postgres:// to postgresql://")

            logger.debug("Using DATABASE_URL for SQLAlchemy connection")
            return database_url

        if require_database_url:
            raise ConfigurationError("Missing required environment variable(s): DATABASE_URL")

        logger.warning(
            "DATABASE_URL is not set; using DB_HOST/DB_PORT/DB_NAME/DB_USER/DB_PASSWORD instead."
        )

        required_db_vars = ("DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD")
        missing_db_vars = _missing_env_vars(*required_db_vars)
        if missing_db_vars:
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing_db_vars)}"
            )

        db_host = os.environ["DB_HOST"]
        db_port = os.environ["DB_PORT"]
        db_name = os.environ["DB_NAME"]
        db_user = os.environ["DB_USER"]
        db_password = quote_plus(os.environ["DB_PASSWORD"])

        logger.debug("Using DB_* environment variables for SQLAlchemy connection")

        return f"postgresql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}"

    @classmethod
    def validate(cls):
        missing = _missing_env_vars("SECRET_KEY")
        if missing:
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing)}"
            )

        secret_key = os.environ["SECRET_KEY"]
        database_uri = cls._build_database_uri()
        cls.SECRET_KEY = secret_key
        cls.SQLALCHEMY_DATABASE_URI = database_uri


class DevelopmentConfig(BaseConfig):
    DEBUG = True

    @classmethod
    def validate(cls):
        super().validate()


class ProductionConfig(BaseConfig):
    DEBUG = False

    @classmethod
    def validate(cls):
        missing = _missing_env_vars("SECRET_KEY", "DATABASE_URL")
        if missing:
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing)}"
            )

        if os.environ.get("FLASK_DEBUG", "0") == "1":
            raise ConfigurationError("FLASK_DEBUG must be '0' when using ProductionConfig")

        secret_key = os.environ["SECRET_KEY"]
        database_uri = cls._build_database_uri(require_database_url=True)
        cls.SECRET_KEY = secret_key
        cls.SQLALCHEMY_DATABASE_URI = database_uri


class TestingConfig(BaseConfig):
    TESTING = True
    SQLALCHEMY_ENGINE_OPTIONS = {}

    @classmethod
    def validate(cls):
        missing = _missing_env_vars("SECRET_KEY")
        if missing:
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing)}"
            )

        secret_key = os.environ["SECRET_KEY"]
        database_uri = "sqlite:///:memory:"
        cls.SECRET_KEY = secret_key
        cls.SQLALCHEMY_DATABASE_URI = database_uri


class SQLiteConfig(BaseConfig):
    DEBUG = True
    basedir = os.path.abspath(os.path.dirname(__file__))

    @classmethod
    def validate(cls):
        missing = _missing_env_vars("SECRET_KEY")
        if missing:
            raise ConfigurationError(
                f"Missing required environment variable(s): {', '.join(missing)}"
            )

        secret_key = os.environ["SECRET_KEY"]
        instance_path = os.environ.get("NEBAXUS_INSTANCE_PATH") or os.path.join(
            cls.basedir, "app", "instance"
        )
        os.makedirs(instance_path, exist_ok=True)
        database_uri = "sqlite:///" + os.path.join(instance_path, "dukana.db")
        cls.SECRET_KEY = secret_key
        cls.SQLALCHEMY_DATABASE_URI = database_uri
        logger.warning("SQLiteConfig enabled — %s", database_uri)


# Backward-compatible alias for any legacy imports.
Config = BaseConfig


class ConfigRegistry(dict):
    """Config mapping that validates classes before returning them."""

    def __getitem__(self, key):
        config_class = super().__getitem__(key)
        config_class.validate()
        return config_class

    def get(self, key, default=None):
        config_class = super().get(key, default)
        if config_class is None:
            return None
        config_class.validate()
        return config_class


config = ConfigRegistry(
    {
        "development": DevelopmentConfig,
        "production": ProductionConfig,
        "testing": TestingConfig,
        "sqlite": SQLiteConfig,
        "default": DevelopmentConfig,
    }
)
