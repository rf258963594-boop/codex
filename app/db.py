import hashlib
import secrets
import sqlite3
from datetime import UTC, datetime

from config import DATA_DIR, DB_PATH


def connect():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000)
    return f"{salt}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        salt, _ = stored.split("$", 1)
    except ValueError:
        return False
    return secrets.compare_digest(hash_password(password, salt), stored)


def init_db():
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              username TEXT UNIQUE NOT NULL,
              password_hash TEXT NOT NULL,
              role TEXT NOT NULL DEFAULT 'staff',
              created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS sessions (
              token TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL,
              created_at TEXT NOT NULL,
              expires_at TEXT NOT NULL,
              FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS common_people (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              display_name TEXT UNIQUE NOT NULL,
              aliases TEXT,
              default_role TEXT,
              id_type TEXT,
              id_number TEXT,
              nationality TEXT,
              residential_address TEXT,
              email TEXT,
              phone TEXT,
              is_local_resident_director INTEGER NOT NULL DEFAULT 0,
              active INTEGER NOT NULL DEFAULT 1,
              notes TEXT
            );

            CREATE TABLE IF NOT EXISTS template_rules (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              rule_key TEXT UNIQUE NOT NULL,
              task_type TEXT NOT NULL,
              change_item TEXT NOT NULL,
              suggested_file TEXT NOT NULL,
              signing_mode TEXT NOT NULL DEFAULT 'default',
              signer_source TEXT NOT NULL DEFAULT '',
              can_merge_dr INTEGER NOT NULL DEFAULT 0,
              manual_review INTEGER NOT NULL DEFAULT 0,
              active INTEGER NOT NULL DEFAULT 1,
              notes TEXT
            );

            CREATE TABLE IF NOT EXISTS field_mappings (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              template_name TEXT NOT NULL,
              template_field TEXT NOT NULL,
              canonical_field TEXT NOT NULL,
              transform_hint TEXT,
              required INTEGER NOT NULL DEFAULT 0,
              UNIQUE(template_name, template_field)
            );

            CREATE TABLE IF NOT EXISTS generation_jobs (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              job_code TEXT UNIQUE NOT NULL,
              task_type TEXT NOT NULL,
              company_name TEXT,
              source_filename TEXT,
              upload_path TEXT NOT NULL,
              parsed_json TEXT NOT NULL,
              suggestions_json TEXT NOT NULL,
              status TEXT NOT NULL DEFAULT 'parsed',
              created_by INTEGER,
              created_at TEXT NOT NULL,
              FOREIGN KEY(created_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS audit_logs (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              user_id INTEGER,
              action TEXT NOT NULL,
              detail TEXT,
              created_at TEXT NOT NULL
            );
            """
        )
        ensure_columns(conn)
        user = conn.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
        if not user:
            conn.execute(
                "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
                ("admin", hash_password("admin123"), "admin", now()),
            )
        else:
            conn.execute("UPDATE users SET active = 1 WHERE username = ?", ("admin",))
        seed_defaults(conn)


def ensure_columns(conn: sqlite3.Connection):
    user_cols = {row["name"] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
    if "active" not in user_cols:
        conn.execute("ALTER TABLE users ADD COLUMN active INTEGER NOT NULL DEFAULT 1")
    if "last_login_at" not in user_cols:
        conn.execute("ALTER TABLE users ADD COLUMN last_login_at TEXT")

    common_cols = {row["name"] for row in conn.execute("PRAGMA table_info(common_people)").fetchall()}
    common_person_columns = {
        "aliases": "TEXT",
        "signature_text": "TEXT",
        "signature_image_path": "TEXT",
        "auto_signature_enabled": "INTEGER NOT NULL DEFAULT 0",
    }
    for column, column_type in common_person_columns.items():
        if column not in common_cols:
            conn.execute(f"ALTER TABLE common_people ADD COLUMN {column} {column_type}")

    job_cols = {row["name"] for row in conn.execute("PRAGMA table_info(generation_jobs)").fetchall()}
    reserved_job_columns = {
        "case_id": "TEXT",
        "business_order_id": "TEXT",
        "source_type": "TEXT",
        "source_file_id": "TEXT",
        "contact_person_id": "TEXT",
        "agent_person_id": "TEXT",
        "client_signatory_person_id": "TEXT",
        "authorized_representative_person_id": "TEXT",
        "prepared_by": "TEXT",
        "snapshot_version": "TEXT",
    }
    for column, column_type in reserved_job_columns.items():
        if column not in job_cols:
            conn.execute(f"ALTER TABLE generation_jobs ADD COLUMN {column} {column_type}")


def now() -> str:
    return datetime.now(UTC).replace(tzinfo=None).isoformat(timespec="seconds") + "Z"


def seed_defaults(conn: sqlite3.Connection):
    people = [
        ("挂名董事 A", "Nominee Director", "NRIC", "S1234567A", "Singaporean", "Singapore", "nd-a@example.com", "+65 90000001", 1, "示例，替换成真实资料"),
        ("挂名董事 B", "Nominee Director", "NRIC", "S2345678B", "Singaporean", "Singapore", "nd-b@example.com", "+65 90000002", 1, "示例，替换成真实资料"),
        ("公司秘书 A", "Secretary", "NRIC", "S3456789C", "Singaporean", "Singapore", "sec-a@example.com", "+65 90000003", 0, "示例，替换成真实资料"),
        ("公司秘书 B", "Secretary", "NRIC", "S4567890D", "Singaporean", "Singapore", "sec-b@example.com", "+65 90000004", 0, "示例，替换成真实资料"),
    ]
    for row in people:
        conn.execute(
            """
            INSERT OR IGNORE INTO common_people
            (display_name, default_role, id_type, id_number, nationality, residential_address, email, phone, is_local_resident_director, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            row,
        )

    rules = [
        ("inc_basic_dr", "incorporation", "incorporation", "First Directors Resolution", "together", "directors", 1, 0, "新注册基础董事决议"),
        ("inc_consent_director", "incorporation", "director_appointment", "Consent to Act as Director", "per_person", "new_directors", 0, 0, "每个董事一份"),
        ("inc_share_certificate", "incorporation", "share_issuance", "Share Certificate", "per_shareholder", "shareholders", 0, 0, "每个股东一份"),
        ("chg_company_name", "change", "company_name", "Members Resolution / Notice of Resolution", "together", "shareholders", 0, 1, "规则可后续调整"),
        ("chg_business", "change", "business_activity", "Directors Resolution", "together", "directors", 1, 0, "业务范围/SSIC"),
        ("chg_address", "change", "registered_office_address", "Directors Resolution", "together", "directors", 1, 0, "注册地址"),
        ("chg_fye", "change", "fye", "Directors Resolution / Manual Review", "together", "directors", 1, 1, "FYE 变更需复核"),
        ("chg_director_appoint", "change", "director_appointment", "Consent to Act + Directors Resolution", "mixed", "new_directors/directors", 1, 0, ""),
        ("chg_director_resign", "change", "director_resignation", "Resignation/Removal Docs + Checklist", "mixed", "resigning_directors/directors", 1, 1, "检查本地董事"),
        ("chg_secretary", "change", "secretary_change", "Directors Resolution + Checklist", "together", "directors", 1, 0, ""),
        ("chg_share_transfer", "change", "share_transfer", "Share Transfer Instrument + Share Update Checklist", "transfer_parties", "transferor/transferee", 1, 0, ""),
        ("chg_personal", "change", "personal_particulars", "Personal Particulars Filing Checklist", "none", "none", 0, 0, ""),
        ("bizfile_checklist", "change", "any_change", "BizFile Filing Checklist", "none", "none", 0, 0, "所有变更任务默认建议"),
    ]
    for row in rules:
        conn.execute(
            """
            INSERT OR IGNORE INTO template_rules
            (rule_key, task_type, change_item, suggested_file, signing_mode, signer_source, can_merge_dr, manual_review, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            row,
        )


def log_action(user_id: int | None, action: str, detail: str = ""):
    with connect() as conn:
        conn.execute(
            "INSERT INTO audit_logs (user_id, action, detail, created_at) VALUES (?, ?, ?, ?)",
            (user_id, action, detail, now()),
        )
