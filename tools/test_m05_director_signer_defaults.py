from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app"))

from doc_generator import build_m05_context  # noqa: E402


def base_parsed(annual_override: dict[str, str] | None = None) -> dict[str, object]:
    annual = {
        "annual_review_required": "Yes",
        "fye_date": "30/06/2026",
        "agm_date": "30/12/2026",
        "accounts_status": "active",
    }
    annual.update(annual_override or {})
    return {
        "company": {
            "company_name": "M05 DIRECTOR SIGNER DEFAULT TEST PTE. LTD.",
            "uen": "202600001A",
            "registered_office_address": "111 NORTH BRIDGE ROAD, SINGAPORE",
            "director_signer_names": "CLIENT DIRECTOR",
            "member_signer_names": "CLIENT DIRECTOR",
            "default_document_date": "30/12/2026",
        },
        "people": [
            {
                "person_id": "P001",
                "full_name": "CLIENT DIRECTOR",
                "is_director": "Yes",
                "is_nominee_director": "No",
                "is_shareholder": "Yes",
            },
            {
                "person_id": "P002",
                "full_name": "COMMON NOMINEE DIRECTOR",
                "is_director": "Yes",
                "is_nominee_director": "Yes",
                "is_local_resident_director": "Yes",
            },
        ],
        "annual_review": annual,
    }


def director_names(context: dict[str, object]) -> list[str]:
    return [row.get("full_name", "") for row in context["m05"]["director_signers"]]


def main() -> None:
    default_context = build_m05_context(base_parsed())
    default_names = director_names(default_context)
    assert default_names == ["CLIENT DIRECTOR", "COMMON NOMINEE DIRECTOR"], default_names

    override_context = build_m05_context(base_parsed({"director_signer_name": "CLIENT DIRECTOR"}))
    override_names = director_names(override_context)
    assert override_names == ["CLIENT DIRECTOR"], override_names

    print("PASS: M05 defaults to all current directors and respects annual-specific override.")


if __name__ == "__main__":
    main()
