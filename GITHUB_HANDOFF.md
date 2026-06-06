# Codex Handoff Notes

This repository is the internal RSIN corporate secretary document generator MVP.

For a fuller Chinese handoff that lets another Codex continue without reading the whole chat history, start with `CODEX_CONTEXT.md`.

## Current stage

- Local website entry: `http://127.0.0.1:8088/`
- Start helper: `OPEN_WEBSITE.cmd`
- Admin smoke-test login: `admin` / `admin123`
- Main flow: upload Excel input, review suggested file groups, generate PDF signing packages.

## Connected document groups

- P1: new company incorporation package.
- P2 M01: ordinary directors' resolution package.
- P2 M02: transfer-in package.
- P2 M03: share transfer package.
- P2 M04: share allotment package.
- P2 M05: annual review package.

## Important directories

- `app/`: local web app, parser, rules, document generation, render logic.
- `app/doc_templates/`: official DOCX templates used by the generator.
- `tools/`: template builders, website control scripts, smoke-test scripts.
- `templates/import/`: clean Excel import templates copied from the latest official outputs.
- `tests/fixtures/`: fake stress-test Excel inputs for P1/P2 document generation.
- `outputs/`: intentionally ignored. It contains generated files, screenshots, render caches, and older scratch outputs.
- `app/generated/`: intentionally ignored. It contains generated PDF/ZIP output.
- `app/uploads/`: intentionally ignored. It contains uploaded Excel files.
- `app/data/`: runtime database/state is ignored. The app can recreate local runtime files.

## Deployment notes

PDF generation depends on LibreOffice and fonts. For cloud deployment, use Docker or a fixed server environment with:

- LibreOffice
- English fonts
- Chinese fonts
- Poppler, if visual PDF checks are needed

See `DEPLOYMENT.md` and `RENDERING.md`.

## Review notes

The document wording is for internal workflow support and should still be reviewed by a qualified corporate secretary or legal reviewer before production use. M05 prepares annual-review signing and authorisation documents; it does not prove that ACRA filing has already been completed.

## Suggested next work

- Review M05 annual-review wording with real business users.
- Unify the website entry flow for M01-M05.
- Add cleaner admin template version management.
- Prepare a production database model for contacts, persons, companies, roles, and document snapshots.
