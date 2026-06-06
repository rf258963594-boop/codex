from __future__ import annotations

import json
import sys
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "app"
OUT_DIR = ROOT / "outputs"
FIXTURE_DIR = ROOT / "tests" / "fixtures" / "defaults"
REPORT_PATH = OUT_DIR / "default_value_regression_report.json"

sys.path.insert(0, str(APP_DIR))

from doc_generator import (  # noqa: E402
    generate_p1_pdf_package,
    generate_p2_m01_pdf_package,
    generate_p2_m02_pdf_package,
    generate_p2_m03_pdf_package,
    generate_p2_m04_pdf_package,
    generate_p2_m05_pdf_package,
)
from excel_parser import parse_excel  # noqa: E402
from rules import suggest_files  # noqa: E402


BAD_TOKENS = ["{{", "}}", "[[REPEAT", "${", "None"]


CASES = [
    {
        "case": "P1_sparse_defaults",
        "path": FIXTURE_DIR / "P1_sparse_defaults.xlsx",
        "generators": [("P1", generate_p1_pdf_package)],
        "expected_yes": [],
        "expected_no": [],
    },
    {
        "case": "ordinary_dr_sparse_defaults",
        "path": FIXTURE_DIR / "ordinary_dr_sparse_defaults.xlsx",
        "generators": [("M01", generate_p2_m01_pdf_package)],
        "expected_yes": ["m01_available"],
        "expected_no": ["m02_available", "m03_available", "m04_available", "m05_available"],
    },
    {
        "case": "transfer_in_sparse_defaults",
        "path": FIXTURE_DIR / "transfer_in_sparse_defaults.xlsx",
        "generators": [("M02", generate_p2_m02_pdf_package)],
        "expected_yes": ["m02_available"],
        "expected_no": ["m01_available", "m03_available", "m04_available", "m05_available"],
    },
    {
        "case": "share_transfer_sparse_defaults",
        "path": FIXTURE_DIR / "share_transfer_sparse_defaults.xlsx",
        "generators": [("M03", generate_p2_m03_pdf_package)],
        "expected_yes": ["m03_available"],
        "expected_no": ["m01_available", "m02_available", "m04_available", "m05_available"],
    },
    {
        "case": "share_allotment_sparse_defaults",
        "path": FIXTURE_DIR / "share_allotment_sparse_defaults.xlsx",
        "generators": [("M04", generate_p2_m04_pdf_package)],
        "expected_yes": ["m04_available"],
        "expected_no": ["m01_available", "m02_available", "m03_available", "m05_available"],
    },
    {
        "case": "annual_review_sparse_defaults",
        "path": FIXTURE_DIR / "annual_review_sparse_defaults.xlsx",
        "generators": [("M05", generate_p2_m05_pdf_package)],
        "expected_yes": ["m05_available"],
        "expected_no": ["m01_available", "m02_available", "m03_available", "m04_available"],
    },
]


def pdf_text_and_pages(pdf_bytes: bytes) -> tuple[int, str]:
    from pypdf import PdfReader

    reader = PdfReader(BytesIO(pdf_bytes))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    return len(reader.pages), text


def check_zip(zip_path: Path) -> dict[str, object]:
    result: dict[str, object] = {
        "zip": str(zip_path.relative_to(ROOT)),
        "exists": zip_path.exists(),
        "pdfs": [],
        "errors": [],
    }
    if not zip_path.exists():
        result["errors"].append("zip_not_found")
        return result
    with ZipFile(zip_path) as zf:
        pdf_names = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
        if not pdf_names:
            result["errors"].append("no_pdf_in_zip")
        for name in pdf_names:
            data = zf.read(name)
            pdf_result = {"name": name, "bytes": len(data), "pages": 0, "bad_tokens": []}
            if data[:4] != b"%PDF":
                pdf_result["bad_tokens"].append("not_pdf_header")
            try:
                pages, text = pdf_text_and_pages(data)
                pdf_result["pages"] = pages
                pdf_result["bad_tokens"].extend(token for token in BAD_TOKENS if token in text)
                if pages <= 0:
                    pdf_result["bad_tokens"].append("zero_pages")
            except Exception as exc:  # pragma: no cover - diagnostic script
                pdf_result["bad_tokens"].append(f"pdf_read_error:{exc}")
            result["pdfs"].append(pdf_result)
            if pdf_result["bad_tokens"]:
                result["errors"].append({"pdf": name, "bad_tokens": pdf_result["bad_tokens"]})
    return result


def run_case(case: dict[str, object]) -> dict[str, object]:
    fixture = Path(case["path"])
    parsed = parse_excel(fixture)
    suggestions = suggest_files(parsed)
    summary = suggestions.get("summary", {})
    result: dict[str, object] = {
        "case": case["case"],
        "fixture": str(fixture.relative_to(ROOT)),
        "task_type": parsed.get("task_type"),
        "summary": summary,
        "expected_checks": [],
        "generated": [],
        "errors": [],
    }

    if summary.get("blocking_errors"):
        result["errors"].append({"blocking_errors": summary.get("blocking_errors")})

    for flag in case["expected_yes"]:
        ok = summary.get(flag) == "Yes"
        result["expected_checks"].append({"flag": flag, "expected": "Yes", "actual": summary.get(flag), "ok": ok})
        if not ok:
            result["errors"].append(f"{flag}_not_yes")
    for flag in case["expected_no"]:
        ok = summary.get(flag) == "No"
        result["expected_checks"].append({"flag": flag, "expected": "No", "actual": summary.get(flag), "ok": ok})
        if not ok:
            result["errors"].append(f"{flag}_not_no")

    for package, generator in case["generators"]:
        job_code = f"REGRESSION-{case['case']}-{package}"
        try:
            zip_path = generator(parsed, job_code)
            result["generated"].append(check_zip(zip_path))
        except Exception as exc:  # pragma: no cover - diagnostic script
            result["errors"].append({package: str(exc)})

    return result


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    results = [run_case(case) for case in CASES]
    report = {
        "status": "pass" if not any(item["errors"] for item in results) else "fail",
        "case_count": len(results),
        "results": results,
    }
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["status"] != "pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
