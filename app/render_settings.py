from __future__ import annotations

import os
import shutil
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = BASE_DIR / "outputs"

def default_soffice_path() -> Path:
    configured = os.environ.get("SOFFICE_PATH")
    if configured:
        return Path(configured)
    discovered = shutil.which("soffice") or shutil.which("libreoffice")
    if discovered:
        return Path(discovered)
    return Path(r"D:\Program Files\program\soffice.com")


SOFFICE_PATH = default_soffice_path()

# LibreOffice can hang if it tries to use a locked/default user profile. Keep a
# project-local profile for headless conversion.
LIBREOFFICE_PROFILE_DIR = Path(
    os.environ.get("LIBREOFFICE_PROFILE_DIR", str(OUTPUT_DIR / ".lo-profile-codex"))
)

DEFAULT_POPPLER_BIN = BASE_DIR / "tools" / "poppler_extract" / "poppler-26.02.0" / "Library" / "bin"
POPPLER_BIN = Path(os.environ["POPPLER_BIN"]) if os.environ.get("POPPLER_BIN") else DEFAULT_POPPLER_BIN
