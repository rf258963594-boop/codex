from __future__ import annotations

import sys
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode, urljoin
from urllib.request import HTTPCookieProcessor, Request, build_opener


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "app"
BASE_URL = "http://127.0.0.1:8088/"

sys.path.insert(0, str(APP_DIR))

from db import connect, hash_password, init_db, now  # noqa: E402


ADMIN_USER = "regression_admin"
ADMIN_PASSWORD = "regression_admin123"
STAFF_USER = "regression_staff"
STAFF_PASSWORD = "regression123"


def upsert_user(username: str, password: str, role: str) -> None:
    init_db()
    with connect() as conn:
        existing = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE users SET password_hash = ?, role = ?, active = 1 WHERE username = ?",
                (hash_password(password), role, username),
            )
        else:
            conn.execute(
                "INSERT INTO users (username, password_hash, role, active, created_at) VALUES (?, ?, ?, 1, ?)",
                (username, hash_password(password), role, now()),
            )


def login(username: str, password: str):
    opener = build_opener(HTTPCookieProcessor(CookieJar()))
    body = urlencode({"username": username, "password": password}).encode("utf-8")
    request = Request(
        urljoin(BASE_URL, "login"),
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    opener.open(request, timeout=30).read()
    return opener


def main() -> None:
    upsert_user(ADMIN_USER, ADMIN_PASSWORD, "admin")
    upsert_user(STAFF_USER, STAFF_PASSWORD, "staff")

    admin = login(ADMIN_USER, ADMIN_PASSWORD)
    html = admin.open(urljoin(BASE_URL, "settings"), timeout=30).read().decode("utf-8", "ignore")
    required = [
        "管理员后台",
        "后台总览",
        "用户账户",
        "常用人员",
        "模板库",
        "正式签字文件模板",
        "Excel 导入表和辅助文件",
        "文件判断规则",
        "最近系统操作",
        "新增用户 / 重置密码",
        "新增常用人员",
        "内部编号 M01",
        "内部编号 M05",
    ]
    missing = [text for text in required if text not in html]
    if missing:
        raise RuntimeError(f"Admin settings page missing expected text: {missing}")
    structure_checks = {
        "admin-kpi": html.count('class="admin-kpi"'),
        "template-group": html.count('class="template-group"'),
        "admin-drawer": html.count('class="admin-drawer"'),
    }
    if structure_checks["admin-kpi"] < 6:
        raise RuntimeError(f"Admin overview cards missing: {structure_checks}")
    if structure_checks["template-group"] < 2:
        raise RuntimeError(f"Document template groups missing: {structure_checks}")
    if structure_checks["admin-drawer"] < 2:
        raise RuntimeError(f"Admin maintenance drawers missing: {structure_checks}")

    staff = login(STAFF_USER, STAFF_PASSWORD)
    try:
        staff.open(urljoin(BASE_URL, "settings"), timeout=30).read()
    except HTTPError as exc:
        if exc.code != 403:
            raise
        body = exc.read().decode("utf-8", "ignore")
        if "只有管理员可以打开后台设置" not in body:
            raise RuntimeError("Staff settings denial page did not show the expected message.")
    else:
        raise RuntimeError("Staff user could open admin settings page.")

    print("admin_settings_smoke_test=pass")


if __name__ == "__main__":
    main()
