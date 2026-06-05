# Document Rendering

DOCX to PDF conversion is done by LibreOffice. In local Windows testing this can use your installed LibreOffice; in Docker/cloud it uses the LibreOffice installed inside the container.

- Windows default fallback: `D:\Program Files\program\soffice.com`
- Docker default: `/usr/bin/soffice`
- Override anywhere with: `SOFFICE_PATH`
- Local LibreOffice profile: `outputs\.lo-profile-codex`
- Docker profile: `/tmp/libreoffice-profile`

Run a render check:

```powershell
& "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app\doc_render.py --input-dir outputs\P1_preserved_templates_v2 --out-dir outputs\P1_preserved_templates_v2_pdf_check
```

Full per-page PNG QA uses the repository-local Poppler zip extraction:

- Default path: `tools\poppler_extract\poppler-26.02.0\Library\bin`
- Override with: `POPPLER_BIN`

Run PDF plus PNG render check:

```powershell
& "C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app\doc_render.py --input-dir outputs\P1_preserved_templates_v2 --out-dir outputs\P1_preserved_templates_v2_pdf_check --png
```

For cloud deployment, prefer the Docker runtime in `Dockerfile`; it installs LibreOffice, Noto CJK fonts, English fonts, and Poppler together so generated PDFs are more stable across machines.
