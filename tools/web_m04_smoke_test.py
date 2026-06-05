from __future__ import annotations

import json
import re
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.parse import urlencode, urljoin
from urllib.request import HTTPCookieProcessor, Request, build_opener


BASE_URL = "http://127.0.0.1:8088/"
INPUT = Path("outputs/M04_share_allotment_stress_input.xlsx")


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
    boundary = "----codex-m04-smoke-boundary"
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


def main() -> None:
    opener = build_opener(HTTPCookieProcessor(CookieJar()))
    login = post_form(opener, "login", {"username": "admin", "password": "admin123"})
    login.read()

    upload = post_file(opener, "upload", INPUT)
    html = upload.read().decode("utf-8", "ignore")
    job_url = upload.geturl()
    if "生成增资配股 M04 PDF 包" not in html:
        raise RuntimeError("M04 generate button was not found on the job page.")
    if "生成普通董事决议 M01 PDF 包" in html:
        raise RuntimeError("M01 button should not appear for an allotment-only M04 job.")
    match = re.search(r"[?&]id=(\d+)", job_url)
    if not match:
        raise RuntimeError(f"Could not read job id from {job_url}")
    job_id = match.group(1)

    generated = post_form(opener, "generate-p2-m04", {"job_id": job_id}, timeout=120)
    package_bytes = generated.read()
    result = {
        "job_url": job_url,
        "has_m04_button": True,
        "has_m01_button": False,
        "has_form24_preview": "Form 24" in html or "Return of Allotment" in html,
        "generated_url": generated.geturl(),
        "generated_is_zip": package_bytes[:2] == b"PK",
        "generated_bytes": len(package_bytes),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
