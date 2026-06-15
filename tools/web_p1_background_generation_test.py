from __future__ import annotations

import http.cookiejar
import json
import os
import sqlite3
import time
import urllib.parse
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "app" / "data" / "secretary_files.db"
GENERATED_DIR = ROOT / "app" / "generated"
BASE_URL = os.environ.get("BASE_URL", "http://127.0.0.1:8088")
INPUT_XLSX = Path(
    os.environ.get("P1_TEST_XLSX", ROOT / "tests" / "fixtures" / "defaults" / "P1_sparse_defaults.xlsx")
)


def post_form(opener: urllib.request.OpenerDirector, path: str, data: dict[str, str]) -> str:
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    request = urllib.request.Request(
        BASE_URL + path,
        data=encoded,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with opener.open(request, timeout=30) as response:
        return response.geturl()


def upload_file(opener: urllib.request.OpenerDirector, path: Path) -> str:
    boundary = "----codex-local-test-boundary"
    payload = b"".join(
        [
            f"--{boundary}\r\n".encode("utf-8"),
            f'Content-Disposition: form-data; name="file"; filename="{path.name}"\r\n'.encode("utf-8"),
            b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
            path.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode("utf-8"),
        ]
    )
    request = urllib.request.Request(
        BASE_URL + "/upload",
        data=payload,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    with opener.open(request, timeout=60) as response:
        return response.geturl()


def latest_job(job_id: str):
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        return conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()


def main() -> None:
    cookie_jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))
    post_form(opener, "/login", {"username": "admin", "password": "admin123"})
    final_url = upload_file(opener, INPUT_XLSX)
    parsed_url = urllib.parse.urlparse(final_url)
    job_id = urllib.parse.parse_qs(parsed_url.query).get("id", [""])[0]
    if not job_id:
        raise RuntimeError(f"Cannot find job id from upload redirect: {final_url}")

    row = latest_job(job_id)
    if not row:
        raise RuntimeError(f"Cannot find job {job_id}")
    suggestions = json.loads(row["suggestions_json"])
    blocking = suggestions.get("summary", {}).get("blocking_errors", [])
    if blocking:
        raise RuntimeError(f"Uploaded job has blocking errors: {blocking}")

    post_form(opener, "/generate-p1", {"job_id": job_id})
    first_status = latest_job(job_id)["status"]
    if first_status not in {"generating_pdf", "pdf_generated"}:
        raise RuntimeError(f"Unexpected first generation status: {first_status}")

    for _ in range(80):
        row = latest_job(job_id)
        if row["status"] == "pdf_generated":
            zip_path = GENERATED_DIR / f"{row['job_code']}_P1_pdf_package.zip"
            if not zip_path.exists():
                raise RuntimeError(f"PDF status done but zip missing: {zip_path}")
            print(f"job_id={job_id}")
            print(f"job_code={row['job_code']}")
            print(f"status={row['status']}")
            print(f"zip={zip_path}")
            return
        if row["status"] == "generation_failed":
            raise RuntimeError(f"Generation failed for job {job_id}")
        time.sleep(3)
    raise TimeoutError(f"Timed out waiting for PDF generation for job {job_id}")


if __name__ == "__main__":
    main()
