import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent


def load_env_file(path: Path = PROJECT_DIR / ".env") -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def path_from_env(name: str, default: Path) -> Path:
    value = os.environ.get(name)
    return Path(value).expanduser() if value else default


load_env_file()

DATA_DIR = path_from_env("DATA_DIR", BASE_DIR / "data")
UPLOAD_DIR = path_from_env("UPLOAD_DIR", BASE_DIR / "uploads")
GENERATED_DIR = path_from_env("GENERATED_DIR", BASE_DIR / "generated")
DOC_TEMPLATE_DIR = path_from_env("DOC_TEMPLATE_DIR", BASE_DIR / "doc_templates")
IMPORT_TEMPLATE_DIR = path_from_env("IMPORT_TEMPLATE_DIR", PROJECT_DIR / "templates" / "import")
OUTPUTS_DIR = path_from_env("OUTPUTS_DIR", PROJECT_DIR / "outputs")
DB_PATH = path_from_env("DB_PATH", DATA_DIR / "secretary_files.db")

APP_NAME = "秘书文件生成器"
P1_VERSION = "P1_v1.0"
SESSION_COOKIE = "sfg_session"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
UPLOAD_RETENTION_DAYS = 30
DEFAULT_ADMIN_USERNAME = os.environ.get("DEFAULT_ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = os.environ.get("DEFAULT_ADMIN_PASSWORD", "admin123")
