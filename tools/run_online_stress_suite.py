from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from http.cookiejar import CookieJar
from pathlib import Path
from zipfile import ZipFile


DEFAULT_BASE_URL = "http://47.236.119.46:8088/"


def as_url(base_url: str, path: str) -> str:
    return urllib.parse.urljoin(base_url.rstrip("/") + "/", path.lstrip("/"))


def read_text(response) -> str:
    data = response.read()
    charset = response.headers.get_content_charset() or "utf-8"
    return data.decode(charset, errors="replace")


def encode_multipart(fields: dict[str, str], files: dict[str, Path]) -> tuple[bytes, str]:
    boundary = "----codex-online-stress-boundary"
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode())
        chunks.append(str(value).encode("utf-8"))
        chunks.append(b"\r\n")
    for name, path in files.items():
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        chunks.append(f"--{boundary}\r\n".encode())
        chunks.append(
            f'Content-Disposition: form-data; name="{name}"; filename="{path.name}"\r\n'.encode()
        )
        chunks.append(f"Content-Type: {content_type}\r\n\r\n".encode())
        chunks.append(path.read_bytes())
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode())
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


class OnlineClient:
    def __init__(self, base_url: str) -> None:
        self.base_url = base_url
        self.cookies = CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cookies))

    def get(self, path: str):
        return self.opener.open(as_url(self.base_url, path), timeout=60)

    def post_form(self, path: str, fields: dict[str, str]):
        data = urllib.parse.urlencode(fields).encode()
        req = urllib.request.Request(as_url(self.base_url, path), data=data, method="POST")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        return self.opener.open(req, timeout=60)

    def post_file(self, path: str, file_path: Path):
        body, content_type = encode_multipart({}, {"file": file_path})
        req = urllib.request.Request(as_url(self.base_url, path), data=body, method="POST")
        req.add_header("Content-Type", content_type)
        req.add_header("Content-Length", str(len(body)))
        return self.opener.open(req, timeout=180)

    def download(self, path: str) -> bytes:
        return self.get(path).read()


def login(client: OnlineClient, username: str, password: str) -> None:
    response = client.post_form("/login", {"username": username, "password": password})
    final_url = response.geturl()
    body = read_text(response)
    if final_url.endswith("/login") or "login_failed" in body.lower() or "密码错误" in body:
        raise RuntimeError("Login failed")


def extract_job_id(url: str, body: str) -> str:
    parsed = urllib.parse.urlparse(url)
    query_id = urllib.parse.parse_qs(parsed.query).get("id", [""])[0]
    if query_id:
        return query_id
    match = re.search(r"/job\?id=(\d+)", body)
    return match.group(1) if match else ""


def extract_generated_links(body: str) -> list[str]:
    links = re.findall(r"href=['\"](/generated/[^'\"]+\.zip)['\"]", body)
    return sorted(set(html.unescape(link) for link in links))


def inspect_zip(path: Path) -> dict[str, object]:
    with ZipFile(path) as zf:
        entries = zf.infolist()
        pdfs = [item.filename for item in entries if item.filename.lower().endswith(".pdf")]
        return {
            "entry_count": len(entries),
            "pdf_count": len(pdfs),
            "pdf_files": pdfs,
            "total_bytes": sum(item.file_size for item in entries),
        }


def upload_and_wait(
    client: OnlineClient,
    workbook: Path,
    download_dir: Path,
    poll_seconds: int,
    timeout_seconds: int,
) -> dict[str, object]:
    start = time.time()
    upload_response = client.post_file("/upload", workbook)
    body = read_text(upload_response)
    final_url = upload_response.geturl()
    job_id = extract_job_id(final_url, body)
    if not job_id:
        raise RuntimeError(f"Could not locate job id after upload: {workbook.name}")
    job_path = f"/job?id={job_id}"

    last_body = body
    links: list[str] = []
    status = "waiting"
    while time.time() - start <= timeout_seconds:
        response = client.get(job_path)
        last_body = read_text(response)
        links = extract_generated_links(last_body)
        if links:
            status = "generated"
            break
        if "generation_failed" in last_body or "生成失败" in last_body:
            status = "generation_failed"
            break
        time.sleep(poll_seconds)
    else:
        status = "timeout"

    download_dir.mkdir(parents=True, exist_ok=True)
    downloaded = []
    for link in links:
        target = download_dir / Path(urllib.parse.urlparse(link).path).name
        target.write_bytes(client.download(link))
        downloaded.append({"path": str(target), "zip": inspect_zip(target)})

    return {
        "workbook": str(workbook),
        "job_id": job_id,
        "job_url": as_url(client.base_url, job_path),
        "status": status,
        "generated_links": [as_url(client.base_url, link) for link in links],
        "downloaded": downloaded,
        "elapsed_seconds": round(time.time() - start, 2),
        "page_excerpt": re.sub(r"\s+", " ", last_body[:1200]),
    }


def write_report(results: list[dict[str, object]], out_dir: Path, base_url: str) -> Path:
    report_path = out_dir / "online_stress_report.md"
    lines = [
        "# Online Stress Test Report",
        "",
        f"Base URL: {base_url}",
        f"Generated at: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Summary",
        "",
    ]
    ok = sum(1 for item in results if item["status"] == "generated")
    lines.append(f"- Workbooks tested: {len(results)}")
    lines.append(f"- Generated successfully: {ok}")
    lines.append(f"- Failed or timeout: {len(results) - ok}")
    lines.append("")
    for item in results:
        lines.append(f"## {Path(str(item['workbook'])).name}")
        lines.append("")
        lines.append(f"- Status: {item['status']}")
        lines.append(f"- Job: {item['job_url']}")
        lines.append(f"- Time: {item['elapsed_seconds']} seconds")
        downloaded = item.get("downloaded", [])
        if downloaded:
            lines.append("- Packages:")
            for package in downloaded:
                zip_info = package["zip"]
                lines.append(
                    f"  - {Path(str(package['path'])).name}: {zip_info['pdf_count']} PDFs, {zip_info['entry_count']} entries"
                )
                for pdf_name in zip_info["pdf_files"]:
                    lines.append(f"    - {pdf_name}")
        else:
            lines.append("- Packages: none")
        lines.append("")
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Run online upload/download stress test.")
    parser.add_argument("--base-url", default=os.environ.get("RBIZ_BASE_URL", DEFAULT_BASE_URL))
    parser.add_argument("--username", default=os.environ.get("RBIZ_USERNAME", ""))
    parser.add_argument("--password", default=os.environ.get("RBIZ_PASSWORD", ""))
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--poll-seconds", type=int, default=5)
    parser.add_argument("--timeout-seconds", type=int, default=420)
    args = parser.parse_args()

    if not args.username or not args.password:
        raise SystemExit("Provide --username/--password or RBIZ_USERNAME/RBIZ_PASSWORD.")

    workbooks = sorted(args.input_dir.glob("*.xlsx"))
    if not workbooks:
        raise SystemExit(f"No .xlsx files found in {args.input_dir}")

    args.out_dir.mkdir(parents=True, exist_ok=True)
    client = OnlineClient(args.base_url)
    login(client, args.username, args.password)

    results: list[dict[str, object]] = []
    for workbook in workbooks:
        print(f"Uploading {workbook.name}...")
        try:
            result = upload_and_wait(
                client,
                workbook,
                args.out_dir / workbook.stem,
                args.poll_seconds,
                args.timeout_seconds,
            )
        except Exception as exc:
            result = {
                "workbook": str(workbook),
                "status": "error",
                "error": str(exc),
                "downloaded": [],
            }
        print(f"  {result['status']}")
        results.append(result)

    (args.out_dir / "online_stress_report.json").write_text(
        json.dumps(results, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    report = write_report(results, args.out_dir, args.base_url)
    print(report)


if __name__ == "__main__":
    main()
