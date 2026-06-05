from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated"
DOC_TEMPLATE_DIR = BASE_DIR / "doc_templates"
OUTPUTS_DIR = PROJECT_DIR / "outputs"
DB_PATH = DATA_DIR / "secretary_files.db"

APP_NAME = "秘书文件生成器"
P1_VERSION = "P1_v1.0"
SESSION_COOKIE = "sfg_session"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
UPLOAD_RETENTION_DAYS = 30
