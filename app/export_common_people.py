from __future__ import annotations

import argparse
import json
import zipfile
from datetime import datetime
from pathlib import Path

from config import DATA_DIR, OUTPUTS_DIR
from db import connect


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
    parser = argparse.ArgumentParser(description="Export common people and signature images.")
    parser.add_argument("--output", type=Path, default=None, help="Output zip path.")
    parser.add_argument("--include-inactive", action="store_true", help="Also export inactive common people.")
    return parser.parse_args()


def clean_signature_path(value: str | None) -> str:
    if not value:
        return ""
    path = Path(str(value).replace("\\", "/"))
    if path.is_absolute():
        return ""
    parts = [part for part in path.parts if part not in {"", ".", ".."}]
    return "/".join(parts)


def default_output_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return OUTPUTS_DIR / f"common_people_migration_{stamp}.zip"


def main() -> None:
    args = parse_args()
    output_path = args.output or default_output_path()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    where = "" if args.include_inactive else "WHERE active = 1"
    with connect() as conn:
        rows = conn.execute(
            f"SELECT {', '.join(COMMON_PERSON_COLUMNS)} FROM common_people {where} ORDER BY display_name"
        ).fetchall()

    people: list[dict] = []
    signature_paths: set[str] = set()
    for row in rows:
        item = {column: row[column] for column in COMMON_PERSON_COLUMNS}
        signature_path = clean_signature_path(item.get("signature_image_path"))
        item["signature_image_path"] = signature_path
        if signature_path:
            signature_paths.add(signature_path)
        people.append(item)

    payload = {
        "schema": "common_people_migration_v1",
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "count": len(people),
        "people": people,
    }

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("common_people.json", json.dumps(payload, ensure_ascii=False, indent=2))
        for relative in sorted(signature_paths):
            source = DATA_DIR / relative
            if source.exists() and source.is_file():
                normalized = relative.replace("\\", "/")
                archive.write(source, f"files/{normalized}")

    print(output_path)
    print(f"Exported {len(people)} common people and {len(signature_paths)} signature references.")


if __name__ == "__main__":
    main()
