from __future__ import annotations

import json
import re
import sys
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.parse import urlencode, urljoin
from urllib.request import HTTPCookieProcessor, Request, build_opener


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "app"
FIXTURE_DIR = ROOT / "tests" / "fixtures" / "defaults"
BASE_URL = "http://127.0.0.1:8088/"

sys.path.insert(0, str(APP_DIR))

from db import connect, hash_password, init_db, now  # noqa: E402


USER = "regression_staff"
PASSWORD = "regression123"

CASES = [
    ("P1_sparse_defaults.xlsx", "生成注册文件 PDF 包", "generate-p1"),
    ("ordinary_dr_sparse_defaults.xlsx", "生成普通变更董事决议 PDF 包", "generate-p2-m01"),
    ("transfer_in_sparse_defaults.xlsx", "生成转入文件 PDF 包", "generate-p2-m02"),
    ("share_transfer_sparse_defaults.xlsx", "生成股份转让 PDF 包", "generate-p2-m03"),
    ("share_allotment_sparse_defaults.xlsx", "生成增资配股 PDF 包", "generate-p2-m04"),
    ("annual_review_sparse_defaults.xlsx", "生成年审 PDF 包", "generate-p2-m05"),
]

INTERNAL_CODES = ("M01", "M02", "M03", "M04", "M05")


def ensure_staff_user() -> None:
    init_db()
    with connect() as conn:
        existing = conn.execute("SELECT id FROM users WHERE username = ?", (USER,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE users SET password_hash = ?, role = 'staff', active = 1 WHERE username = ?",
                (hash_password(PASSWORD), USER),
            )
        else:
            conn.execute(
                "INSERT INTO users (username, password_hash, role, active, created_at) VALUES (?, ?, 'staff', 1, ?)",
                (USER, hash_password(PASSWORD), now()),
            )


def post_form(opener, path: str, data: dict[str, str], timeout: int = 30):
    body = urlencode(data).encode("utf-8")
    request = Request(
        urljoin(BASE_URL, path),
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    return opener.open(request, timeout=timeout)


def post_file(opener, path: str, file_path: Path, timeout: int = 60):
    boundary = "----codex-default-regression-boundary"
    payload = b"".join(
        [
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="file"; filename="{file_path.name}"\r\n'.encode(),
            b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
            file_path.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode(),
        ]
    )
    request = Request(
        urljoin(BASE_URL, path),
        data=payload,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    return opener.open(request, timeout=timeout)


def job_id_from_url(url: str) -> str:
    match = re.search(r"[?&]id=(\d+)", url)
    if not match:
        raise RuntimeError(f"Could not read job id from {url}")
    return match.group(1)


def assert_staff_text_is_clean(html: str, context: str) -> None:
    for code in INTERNAL_CODES:
        if code in html:
            raise RuntimeError(f"{context}: staff page should not expose internal code {code}.")


def main() -> None:
    ensure_staff_user()
    opener = build_opener(HTTPCookieProcessor(CookieJar()))
    post_form(opener, "login", {"username": USER, "password": PASSWORD}).read()

    home_html = opener.open(BASE_URL, timeout=30).read().decode("utf-8", "ignore")
    for required in ["文件生成工作台", "注册文件生成", "变更 / 年审文件生成", "默认填写规则"]:
        if required not in home_html:
            raise RuntimeError(f"Home page missing text: {required}")
    assert_staff_text_is_clean(home_html, "home")

    results = []
    for filename, button_label, endpoint in CASES:
        fixture = FIXTURE_DIR / filename
        upload = post_file(opener, "upload", fixture)
        html = upload.read().decode("utf-8", "ignore")
        job_url = upload.geturl()
        if button_label not in html:
            raise RuntimeError(f"{filename}: generate button not found: {button_label}")
        if "管理员调试数据" in html or "管理员文件规则明细" in html:
            raise RuntimeError(f"{filename}: staff user can see admin-only debug content.")
        assert_staff_text_is_clean(html, filename)
        if filename.startswith(("ordinary_dr", "share_transfer", "share_allotment", "annual_review")) and "缺少董事签字人" not in html:
            raise RuntimeError(f"{filename}: missing blank director signer warning.")
        if filename.startswith(("transfer_in", "share_allotment", "annual_review")) and "缺少股东/客户授权签字人" not in html:
            raise RuntimeError(f"{filename}: missing blank member/client signer warning.")
        job_id = job_id_from_url(job_url)
        generated = post_form(opener, endpoint, {"job_id": job_id}, timeout=160)
        package = generated.read()
        if package[:2] != b"PK":
            raise RuntimeError(f"{filename}: generated response is not a zip file.")
        results.append(
            {
                "fixture": filename,
                "job_url": job_url,
                "button_label": button_label,
                "generated_url": generated.geturl(),
                "generated_bytes": len(package),
            }
        )

    print(json.dumps({"status": "pass", "cases": results}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
