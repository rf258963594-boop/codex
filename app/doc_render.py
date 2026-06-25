from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
from pathlib import Path

from pdf2image import convert_from_path
from pypdf import PdfReader

from render_settings import LIBREOFFICE_PROFILE_DIR, POPPLER_BIN, SOFFICE_PATH


def file_uri(path: Path) -> str:
    return path.resolve().as_uri()


def libreoffice_env() -> dict[str, str]:
    env = os.environ.copy()
    # Use LibreOffice's headless VCL plugin and disable GPU probing. This makes
    # conversion less likely to block on desktop/printer integration.
    env.setdefault("SAL_USE_VCLPLUGIN", "svp")
    env.setdefault("SAL_DISABLE_OPENCL", "1")
    return env


def kill_process_tree(proc: subprocess.Popen) -> None:
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        return
    proc.kill()


def run_soffice(cmd: list[str], *, timeout: int) -> subprocess.CompletedProcess[str]:
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=libreoffice_env(),
        creationflags=creationflags,
    )
    try:
        stdout, stderr = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        kill_process_tree(proc)
        stdout, stderr = proc.communicate()
        raise TimeoutError(
            "LibreOffice conversion timed out. On Windows this is often caused "
            "by an unavailable default/network printer; set a local printer such "
            "as Microsoft Print to PDF as the default printer, then retry."
        ) from exc
    return subprocess.CompletedProcess(cmd, proc.returncode, stdout, stderr)


def ensure_libreoffice() -> Path:
    if not SOFFICE_PATH.exists():
        raise FileNotFoundError(
            f"LibreOffice executable not found: {SOFFICE_PATH}. "
            "Set SOFFICE_PATH to the soffice.com/soffice.exe path."
        )
    LIBREOFFICE_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    init_cmd = [
        str(SOFFICE_PATH),
        f"-env:UserInstallation={file_uri(LIBREOFFICE_PROFILE_DIR)}",
        "--headless",
        "--invisible",
        "--nologo",
        "--nodefault",
        "--nocrashreport",
        "--terminate_after_init",
    ]
    try:
        run_soffice(init_cmd, timeout=30)
    except TimeoutError:
        # The actual conversion path will surface a more specific error. Do not
        # fail application startup because the profile warm-up was blocked.
        pass
    return SOFFICE_PATH


def convert_docx_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    ensure_libreoffice()
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(SOFFICE_PATH),
        f"-env:UserInstallation={file_uri(LIBREOFFICE_PROFILE_DIR)}",
        "--headless",
        "--invisible",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        "--nodefault",
        "--nocrashreport",
        "--nolockcheck",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(out_dir.resolve()),
        str(docx_path.resolve()),
    ]
    proc = run_soffice(cmd, timeout=120)
    pdf_path = out_dir / f"{docx_path.stem}.pdf"
    if proc.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"LibreOffice conversion failed for {docx_path.name}\n"
            f"exit={proc.returncode}\nstdout={proc.stdout}\nstderr={proc.stderr}"
        )
    return pdf_path


def pdf_info(pdf_path: Path) -> dict:
    reader = PdfReader(str(pdf_path))
    page_text_lengths = []
    for page in reader.pages:
        page_text_lengths.append(len(page.extract_text() or ""))
    return {
        "file": pdf_path.name,
        "pages": len(reader.pages),
        "bytes": pdf_path.stat().st_size,
        "page_text_lengths": page_text_lengths,
    }


def poppler_ready() -> bool:
    if os.environ.get("POPPLER_BIN"):
        exe_suffix = ".exe" if os.name == "nt" else ""
        return bool(
            POPPLER_BIN
            and (POPPLER_BIN / f"pdfinfo{exe_suffix}").exists()
            and (POPPLER_BIN / f"pdftoppm{exe_suffix}").exists()
        )
    if POPPLER_BIN and POPPLER_BIN.exists():
        return (POPPLER_BIN / "pdfinfo.exe").exists() and (POPPLER_BIN / "pdftoppm.exe").exists()
    return bool(shutil.which("pdfinfo") and shutil.which("pdftoppm"))


def render_pdf_pages_to_png(pdf_path: Path, out_dir: Path, dpi: int = 120) -> list[Path]:
    if not poppler_ready():
        raise FileNotFoundError(
            "Poppler is not ready. POPPLER_BIN must contain pdfinfo.exe and pdftoppm.exe."
        )
    stem_dir = out_dir / pdf_path.stem
    if stem_dir.exists():
        for old in stem_dir.glob("*.png"):
            old.unlink()
    stem_dir.mkdir(parents=True, exist_ok=True)
    kwargs = {}
    if os.environ.get("POPPLER_BIN") or (POPPLER_BIN and POPPLER_BIN.exists()):
        kwargs["poppler_path"] = str(POPPLER_BIN)
    paths = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        fmt="png",
        output_folder=str(stem_dir),
        paths_only=True,
        **kwargs,
    )
    normalized: list[Path] = []
    for index, raw in enumerate(paths, start=1):
        src = Path(raw)
        dst = stem_dir / f"page-{index}.png"
        if src.resolve() != dst.resolve():
            if dst.exists():
                dst.unlink()
            shutil.move(str(src), str(dst))
        normalized.append(dst)
    return normalized


def render_folder_to_pdf(input_dir: Path, out_dir: Path) -> list[dict]:
    reports = []
    for docx_path in sorted(input_dir.glob("*.docx")):
        pdf_path = convert_docx_to_pdf(docx_path, out_dir)
        info = pdf_info(pdf_path)
        info["source"] = docx_path.name
        reports.append(info)
    return reports


def render_folder_to_png(pdf_dir: Path, out_dir: Path, dpi: int = 120) -> list[dict]:
    reports = []
    out_dir.mkdir(parents=True, exist_ok=True)
    for pdf_path in sorted(pdf_dir.glob("*.pdf")):
        pages = render_pdf_pages_to_png(pdf_path, out_dir, dpi=dpi)
        reports.append(
            {
                "source": pdf_path.name,
                "png_pages": len(pages),
                "folder": str((out_dir / pdf_path.stem).resolve()),
                "files": [p.name for p in pages],
            }
        )
    return reports


def environment_report() -> dict:
    return {
        "soffice_path": str(SOFFICE_PATH),
        "soffice_exists": SOFFICE_PATH.exists(),
        "libreoffice_profile_dir": str(LIBREOFFICE_PROFILE_DIR),
        "poppler_bin": str(POPPLER_BIN) if POPPLER_BIN else "",
        "poppler_ready": poppler_ready(),
        "notes": [
            "DOCX to PDF uses local LibreOffice and does not need internet.",
            "Full per-page PNG QA needs Poppler. Without it, use the generated PDFs for visual review.",
            "Python package installation from the internet may be blocked by the Codex/network policy.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Render DOCX templates to PDF using local LibreOffice.")
    parser.add_argument("--input-dir", default="outputs/P1_preserved_templates_v2")
    parser.add_argument("--out-dir", default="outputs/P1_preserved_templates_v2_pdf_check")
    parser.add_argument("--png-dir", default="outputs/P1_preserved_templates_v2_png_check")
    parser.add_argument("--png", action="store_true", help="Also render generated PDFs into per-page PNGs.")
    parser.add_argument("--dpi", type=int, default=120)
    parser.add_argument("--env", action="store_true", help="Print renderer environment only.")
    args = parser.parse_args()

    if args.env:
        print(json.dumps(environment_report(), ensure_ascii=False, indent=2))
        return

    reports = render_folder_to_pdf(Path(args.input_dir), Path(args.out_dir))
    png_reports = render_folder_to_png(Path(args.out_dir), Path(args.png_dir), dpi=args.dpi) if args.png else []
    report_path = Path(args.out_dir) / "render_report.json"
    report_path.write_text(
        json.dumps(
            {"environment": environment_report(), "files": reports, "png": png_reports},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(json.dumps({"report": str(report_path), "files": reports, "png": png_reports}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
