from __future__ import annotations

import argparse
import json
import shutil
import tempfile
import zipfile
from pathlib import Path

from config import DATA_DIR
from db import connect, init_db


COMMON_PERSON_COLUMNS = [
    "display_name",
    "aliases",
    "default_role",
    "id_type",
    "id_number",
    "nationality",
    "residential_address",
    "email",
    "phone",
    "is_local_resident_director",
    "active",
    "notes",
    "signature_text",
    "signature_image_path",
    "auto_signature_enabled",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import common people and signature images.")
    parser.add_argument("--input", required=True, type=Path, help="Migration zip created by export_common_people.py.")
    parser.add_argument(
        "--deactivate-missing",
        action="store_true",
        help="Deactivate existing common people not present in the import package.",
    )
    return parser.parse_args()


def safe_relative_path(value: str | None) -> str:
    if not value:
        return ""
    path = Path(str(value).replace("\\", "/"))
    if path.is_absolute():
        return ""
    parts = [part for part in path.parts if part not in {"", ".", ".."}]
    return "/".join(parts)


def extract_zip(input_path: Path, temp_dir: Path) -> dict:
    with zipfile.ZipFile(input_path, "r") as archive:
        for member in archive.infolist():
            target = (temp_dir / member.filename).resolve()
            if not str(target).startswith(str(temp_dir.resolve())):
                raise SystemExit(f"Unsafe path in migration package: {member.filename}")
            archive.extract(member, temp_dir)
    payload_path = temp_dir / "common_people.json"
    if not payload_path.exists():
        raise SystemExit("Invalid migration package: common_people.json is missing.")
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    if payload.get("schema") != "common_people_migration_v1":
        raise SystemExit("Invalid migration package schema.")
    return payload


def copy_signature_files(temp_dir: Path, people: list[dict]) -> None:
    for item in people:
        relative = safe_relative_path(item.get("signature_image_path"))
        item["signature_image_path"] = relative
        if not relative:
            continue
        source = temp_dir / "files" / relative
        if not source.exists() or not source.is_file():
            continue
        target = DATA_DIR / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)


def normalize_item(item: dict) -> dict:
    normalized = {column: item.get(column) for column in COMMON_PERSON_COLUMNS}
    normalized["display_name"] = str(normalized.get("display_name") or "").strip()
    normalized["is_local_resident_director"] = 1 if normalized.get("is_local_resident_director") else 0
    normalized["active"] = 1 if normalized.get("active", 1) else 0
    normalized["auto_signature_enabled"] = 1 if normalized.get("auto_signature_enabled") else 0
    normalized["signature_image_path"] = safe_relative_path(normalized.get("signature_image_path"))
    return normalized


def upsert_people(people: list[dict], deactivate_missing: bool) -> tuple[int, int, int]:
    imported_names: set[str] = set()
    inserted = 0
    updated = 0
    deactivated = 0

    with connect() as conn:
        for raw in people:
            item = normalize_item(raw)
            if not item["display_name"]:
                continue
            imported_names.add(item["display_name"])
            existing = conn.execute(
                "SELECT id FROM common_people WHERE display_name = ?",
                (item["display_name"],),
            ).fetchone()
            values = [item[column] for column in COMMON_PERSON_COLUMNS]
            if existing:
                assignments = ", ".join(f"{column} = ?" for column in COMMON_PERSON_COLUMNS[1:])
                conn.execute(
                    f"UPDATE common_people SET {assignments} WHERE display_name = ?",
                    [item[column] for column in COMMON_PERSON_COLUMNS[1:]] + [item["display_name"]],
                )
                updated += 1
            else:
                placeholders = ", ".join("?" for _ in COMMON_PERSON_COLUMNS)
                conn.execute(
                    f"INSERT INTO common_people ({', '.join(COMMON_PERSON_COLUMNS)}) VALUES ({placeholders})",
                    values,
                )
                inserted += 1

        if deactivate_missing and imported_names:
            placeholders = ", ".join("?" for _ in imported_names)
            result = conn.execute(
                f"UPDATE common_people SET active = 0 WHERE display_name NOT IN ({placeholders})",
                sorted(imported_names),
            )
            deactivated = result.rowcount if result.rowcount is not None else 0

    return inserted, updated, deactivated


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise SystemExit(f"Migration package not found: {args.input}")

    init_db()
    with tempfile.TemporaryDirectory(prefix="common_people_import_") as temp:
        temp_dir = Path(temp)
        payload = extract_zip(args.input, temp_dir)
        people = payload.get("people") or []
        copy_signature_files(temp_dir, people)
        inserted, updated, deactivated = upsert_people(people, args.deactivate_missing)

    print(f"Imported common people: inserted={inserted}, updated={updated}, deactivated={deactivated}")


if __name__ == "__main__":
    main()
