from __future__ import annotations

import http.cookiejar
import json
import re
import urllib.parse
import urllib.request
import uuid
from pathlib import Path


BASE_URL = "http://127.0.0.1:8088"
INPUT = Path("outputs/M02_M01_transfer_in_registered_office_secretary_director_stress.xlsx")


def main() -> None:
    cookies = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookies))

    login_data = urllib.parse.urlencode({"username": "admin", "password": "admin123"}).encode()
    opener.open(
        urllib.request.Request(
            BASE_URL + "/login",
            data=login_data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        ),
        timeout=20,
    ).read()

    boundary = "----codex" + uuid.uuid4().hex
    body = b"".join(
        [
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="file"; filename="{INPUT.name}"\r\n'.encode(),
            b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
            INPUT.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode(),
        ]
    )
    upload_response = opener.open(
        urllib.request.Request(
            BASE_URL + "/upload",
            data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        ),
        timeout=60,
    )
    html = upload_response.read().decode("utf-8", "ignore")
    final_url = upload_response.geturl()
    match = re.search(r"id=(\d+)", final_url)
    job_id = match.group(1) if match else ""

    checks = {
        "final_url": final_url,
        "job_id": job_id,
        "has_m01_button": "generate-p2-m01" in html,
        "has_m02_button": "generate-p2-m02" in html,
        "has_address_change": "注册地址变更" in html or "Change of Registered Office" in html,
        "has_two_pdf_text": "两份 PDF" in html,
    }

    for endpoint in ["generate-p2-m01", "generate-p2-m02"]:
        generate_data = urllib.parse.urlencode({"job_id": job_id}).encode()
        response = opener.open(
            urllib.request.Request(
                BASE_URL + "/" + endpoint,
                data=generate_data,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
            ),
            timeout=160,
        )
        prefix = response.read(64)
        checks[endpoint] = {
            "generated_url": response.geturl(),
            "generated_is_zip": prefix.startswith(b"PK"),
        }

    print(json.dumps(checks, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
