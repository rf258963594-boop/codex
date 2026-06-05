from __future__ import annotations

import json
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from build_p2_m03_templates import (
    OUTPUT_DIR,
    TEMPLATE_DIR,
    add_company_header,
    add_para,
    add_table,
    clause_para,
    clause_title,
    configure_doc,
    marker,
    set_cell_margins,
    set_cell_text,
    set_row_cant_split,
    set_table_borders,
    set_table_width,
)


TEMPLATES = {
    "agm_package": "M05_agm_documents_package_standard.docx",
    "annual_return_package": "M05_annual_return_authorisation_package_standard.docx",
    "checklist": "M05_annual_review_checklist_standard.docx",
}


def ensure_dirs() -> None:
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for directory in (TEMPLATE_DIR, OUTPUT_DIR):
        for path in directory.glob("M05_*"):
            if path.is_file():
                path.unlink()


def add_signature_rows(doc: Document, expr: str, heading: str) -> None:
    add_para(doc, heading, bold=True, size=10.5, after=4, keep_with_next=True)
    marker(doc, f"[[REPEAT {expr}[]]]")
    table = doc.add_table(rows=1, cols=2)
    set_table_width(table, [4680, 4680])
    set_table_borders(table, "FFFFFF")
    set_cell_margins(table, top=40, bottom=60, start=0, end=160)
    set_row_cant_split(table.rows[0])
    set_cell_text(table.cell(0, 0), "{{sigrow.left_block}}", size=10.3)
    set_cell_text(table.cell(0, 1), "{{sigrow.right_block}}", size=10.3)


def add_attendance_table(doc: Document) -> None:
    marker(doc, "[[REPEAT m05.attendance_rows[]]]")
    add_table(
        doc,
        ["Name", "Capacity", "Signature"],
        [["{{attendee.full_name}}", "{{attendee.capacity}}", ""]],
        [3300, 2100, 3960],
    )


def add_checklist_table(doc: Document) -> None:
    marker(doc, "[[REPEAT m05.checklist_items[]]]")
    add_table(
        doc,
        ["Item", "Status", "Review note"],
        [["{{check.item}}", "{{check.status}}", "{{check.note}}"]],
        [2350, 1450, 5560],
    )


def build_agm_package() -> Document:
    doc = Document()
    configure_doc(doc, "M05 AGM Documents Package", "P2 annual review AGM documents")
    add_company_header(doc, "NOTICE OF AN ANNUAL GENERAL MEETING")
    add_table(
        doc,
        ["Meeting place", "Meeting time", "Meeting date"],
        [["{{m05.agm_place}}", "{{m05.agm_time}}", "{{m05.agm_date}}"]],
        [3900, 1800, 3060],
    )
    clause_para(doc, "NOTICE IS HEREBY GIVEN that an Annual General Meeting of the Company will be held at the place, time and date stated above, for the purpose of considering and, if thought fit, passing the ordinary business set out below.")
    clause_title(doc, "Business")
    clause_para(doc, "1. To receive and consider, and if thought fit, adopt the directors' statement and the financial statements of the Company for the financial year ended {{m05.fye_date_upper}}.")
    clause_para(doc, "2. To re-elect the director(s) retiring pursuant to the Constitution of the Company, where applicable.")
    clause_para(doc, "3. To approve the non-appointment of an auditor for the ensuing year where the Company qualifies for audit exemption.")
    clause_para(doc, "4. To confirm the directors' fees and remuneration for the financial year.")
    clause_para(doc, "5. To authorise the preparation and lodgement of the Annual Return and related statutory declarations with ACRA / BizFile.")
    clause_para(doc, "6. To transact any other ordinary business which may be properly transacted at an Annual General Meeting.")
    add_para(doc, "Dated this {{signature.day_ordinal}} day of {{signature.month_year}}", style="Signature Block", before=10, after=12)
    add_para(doc, "By Order of the Board", style="Signature Block", after=24)
    add_para(doc, "_____________________________________", style="Signature Block", after=2)
    add_para(doc, "{{m05.notice_issuer_name}}", style="Signature Block", after=0)
    add_para(doc, "{{m05.notice_issuer_capacity}}", style="Signature Block", after=0)
    add_para(doc, "A member entitled to attend and vote at this meeting is entitled to appoint a proxy to attend and vote instead of him/her. A proxy need not be a member of the Company.", style="Small Note", before=10, after=0)

    doc.add_page_break()
    add_company_header(doc, "AGREEMENT BY MEMBERS TO SHORTER NOTICE")
    clause_para(doc, "Pursuant to Section 177(3)(a) of the Companies Act 1967, the undersigned member(s) of the Company hereby agree to the Annual General Meeting of the Company being held on {{m05.agm_date}} at {{m05.agm_time}}, notwithstanding that it may be called by notice shorter than the period of notice otherwise required.")
    clause_para(doc, "The undersigned member(s) also consent, where applicable, to the receipt of the financial statements and related documents less than the prescribed period before the Annual General Meeting.")
    add_para(doc, "Dated this {{signature.day_ordinal}} day of {{signature.month_year}}", style="Signature Block", before=10, after=10)
    add_signature_rows(doc, "m05.member_signature_rows", "MEMBER(S)")

    doc.add_page_break()
    add_company_header(doc, "ATTENDANCE SHEET")
    clause_para(doc, "Attendance at the Annual General Meeting of the Company held on {{m05.agm_date}} at {{m05.agm_time}} at {{m05.agm_place}}.")
    add_attendance_table(doc)

    doc.add_page_break()
    add_company_header(doc, "MINUTES OF AN ANNUAL GENERAL MEETING")
    add_table(
        doc,
        ["Date and time", "Place", "Chairperson"],
        [["{{m05.agm_date}} at {{m05.agm_time}}", "{{m05.agm_place}}", "{{m05.chairperson_name}}"]],
        [2500, 5160, 1700],
    )
    clause_para(doc, "The Chairperson noted that a quorum was present and declared the meeting open. The notice convening the meeting was taken as read with the consent of the meeting.")
    clause_title(doc, "Directors' Statement and Financial Statements")
    clause_para(doc, "It was resolved that the directors' statement and the financial statements of the Company for the financial year ended {{m05.fye_date_upper}} be and are hereby received and adopted.")
    clause_title(doc, "Audit Exemption")
    clause_para(doc, "It was noted that the Company qualifies, or is expected to qualify, for audit exemption as a small company / exempt private company for the relevant financial year, subject to the final confirmation of the Company's records and applicable statutory requirements.")
    clause_title(doc, "Directors' Fees and Remuneration")
    clause_para(doc, "{{m05.directors_fee_text}}")
    clause_para(doc, "{{m05.directors_remuneration_text}}")
    clause_title(doc, "Non-Appointment of Auditor")
    clause_para(doc, "It was resolved that an auditor need not be appointed for the ensuing year if the Company continues to satisfy the applicable audit exemption requirements.")
    clause_title(doc, "Annual Return")
    clause_para(doc, "It was resolved that the authorised director and/or authorised representative be authorised to sign and lodge the Annual Return and related declarations with ACRA / BizFile for the financial year ended {{m05.fye_date_upper}}.")
    clause_title(doc, "Closure")
    clause_para(doc, "There being no further business, the meeting was closed.")
    add_para(doc, "", after=18)
    add_para(doc, "_____________________________________", style="Signature Block", after=2)
    add_para(doc, "{{m05.chairperson_name}}", style="Signature Block", after=0)
    add_para(doc, "Chairperson", style="Signature Block", after=0)
    return doc


def build_annual_return_package() -> Document:
    doc = Document()
    configure_doc(doc, "M05 Annual Return Authorisation Package", "P2 annual return authorisation and declarations")
    add_company_header(doc, "CERTIFICATE BY A COMPANY LIMITED BY SHARES UNDER SECTION 197(1)")
    clause_para(doc, "I, the undermentioned officer of the abovementioned Company, hereby certify that:")
    clause_para(doc, "a) I have verified that the summary of return by a company having a share capital, as reflected in the records of the Registrar, is accurate and up to date as at {{m05.ar_as_at_date}};")
    clause_para(doc, "b) I have made or caused to be made an inspection of the Register of Members and confirm that the relevant share transfer and share capital records have been reviewed for the purpose of the Annual Return;")
    clause_para(doc, "c) the Company is a private company and has not issued any invitation to the public to subscribe for any shares or debentures of the Company; and")
    clause_para(doc, "d) the number of members of the Company is within the limit applicable to a private company, subject to final verification against the Company's statutory records.")
    add_para(doc, "Dated this {{signature.day_ordinal}} day of {{signature.month_year}}", style="Signature Block", before=10, after=18)
    add_para(doc, "_____________________________________", style="Signature Block", after=2)
    add_para(doc, "{{m05.ar_signer_name}}", style="Signature Block", after=0)
    add_para(doc, "{{m05.ar_signer_capacity}}", style="Signature Block", after=0)

    doc.add_page_break()
    add_company_header(doc, "STATEMENT BY A SMALL COMPANY EXEMPT FROM AUDIT REQUIREMENTS")
    clause_para(doc, "I/We, the undermentioned director(s) of the Company, hereby declare that, on behalf of the Board of Directors:")
    clause_para(doc, "a) for the financial period {{m05.financial_period_text}}, the Company qualifies, or is expected to qualify based on the information provided, as a small company under Section 205C of the Companies Act 1967 read with the Thirteenth Schedule;")
    clause_para(doc, "b) no notice has been received from any member requiring the Company to obtain an audit of its financial statements in relation to the financial year; and")
    clause_para(doc, "c) the accounting and other records required to be kept by the Company under Section 199 of the Companies Act 1967 have been kept.")
    add_para(doc, "Dated this {{signature.day_ordinal}} day of {{signature.month_year}}", style="Signature Block", before=10, after=10)
    add_signature_rows(doc, "m05.director_signature_rows", "DIRECTOR(S)")

    doc.add_page_break()
    add_para(doc, "{{m05.agm_date}}", align=WD_ALIGN_PARAGRAPH.RIGHT, after=14)
    add_para(doc, "{{provider.name}}", bold=True, after=2)
    add_para(doc, "{{provider.registered_address}}", after=14)
    add_para(doc, "Dear Sirs", after=10)
    add_para(doc, "FILING OF ANNUAL RETURN - FINANCIAL YEAR ENDED {{m05.fye_date_upper}}", bold=True, after=10)
    clause_para(doc, "I/We, the director(s) and/or authorised signatory of {{company.company_name}} (UEN: {{company.uen}}), hereby authorise {{provider.name}} and its designated agents to prepare, initiate and lodge the Annual Return and related electronic filings with ACRA / BizFile for the financial year ended {{m05.fye_date_upper}}.")
    clause_para(doc, "I/We declare that the particulars of the Company and the information provided for the Annual Return are, to the best of my/our knowledge and belief, true, complete and up to date.")
    clause_para(doc, "I/We further confirm that the financial statements and related information provided for the purpose of the Annual Return have been prepared in accordance with the applicable requirements of the Companies Act 1967, subject to any final review required by the directors and the Company's appointed advisors.")
    clause_para(doc, "This authorisation does not remove the directors' responsibility for ensuring that the Company's statutory records, accounting records and Annual Return information are accurate and complete.")
    add_para(doc, "Yours faithfully,", after=18)
    add_para(doc, "_____________________________________", style="Signature Block", after=2)
    add_para(doc, "{{m05.ar_signer_name}}", style="Signature Block", after=0)
    add_para(doc, "{{m05.ar_signer_capacity}}", style="Signature Block", after=0)

    doc.add_page_break()
    add_company_header(doc, "MANAGEMENT REPRESENTATION")
    clause_para(doc, "We confirm, to the best of our knowledge and belief, the following matters for the financial period ended {{m05.fye_date_upper}}:")
    clause_para(doc, "1. We have not received any notice from a shareholder requiring the Company to obtain an audit of its financial statements in relation to the above period.")
    clause_para(doc, "2. All liabilities, contingent liabilities, guarantees and material commitments have been recorded or disclosed to the preparer of the financial statements.")
    clause_para(doc, "3. The Company's accounting and other records required under the Companies Act 1967 have been maintained.")
    clause_para(doc, "4. The financial statements are, to the best of our knowledge and belief, free from material misstatement, including omissions.")
    clause_para(doc, "5. There have been no subsequent events requiring adjustment or disclosure in the financial statements, except as already disclosed.")
    clause_para(doc, "6. We confirm that the Company is able to meet its liabilities as and when they fall due, or that appropriate financial support arrangements have been considered and disclosed where required.")
    add_para(doc, "Yours faithfully,", after=18)
    add_signature_rows(doc, "m05.director_signature_rows", "DIRECTOR(S)")
    return doc


def build_checklist() -> Document:
    doc = Document()
    configure_doc(doc, "M05 Annual Review Checklist", "P2 annual review internal checklist")
    add_company_header(doc, "ANNUAL REVIEW INTERNAL CHECKLIST")
    add_table(
        doc,
        ["FYE", "AGM date", "AGM route", "Accounts status"],
        [["{{m05.fye_date}}", "{{m05.agm_date}}", "{{m05.agm_route}}", "{{m05.accounts_status}}"]],
        [1800, 1800, 2500, 2260],
    )
    add_checklist_table(doc)
    add_para(doc, "This checklist is for internal review only and should not be treated as confirmation that ACRA / BizFile filing has been completed.", style="Small Note", before=8, after=0)
    return doc


def field_map_text() -> str:
    return """# M05 Annual Review Package - Field Map

M05 is the annual review / AGM / Annual Return authorisation package for an existing Singapore company.

## Generated PDFs

| File | Content | Main signer |
|---|---|---|
| AGM documents package | Notice of AGM, shorter notice consent, attendance sheet, AGM minutes | Member(s), chairperson / authorised signer |
| Annual Return authorisation package | Section 197 certificate, Section 205C statement, AR filing authorisation, management representation | Director(s) / authorised signer |
| Internal checklist | Filing and review checklist | Internal only |

## Primary fields

| Field | Source |
|---|---|
| `annual_review_required` | P2 one-page sheet or Quick Annual Review sheet |
| `fye_date` | Quick Annual Review |
| `agm_date` | Quick Annual Review; defaults to document date if blank |
| `agm_time` | Quick Annual Review; defaults to 10.00 a.m. |
| `agm_place` | Quick Annual Review; defaults to registered office |
| `financial_statement_date` | Quick Annual Review; defaults to FYE |
| `director_signer_name` / `director_signer_names` | Quick Annual Review or company page |
| `shareholder_signer_name` / `member_signer_names` | Quick Annual Review or company page |
| `ar_authorized_signer_name` | Quick Annual Review; defaults to first director signer |
| `directors_fee` / `directors_remuneration` | Quick Annual Review |
| `management_rep_letter` | Quick Annual Review; default Yes in v0.1 |

## Guardrails

The package prepares signing documents and authorisations. It does not represent that ACRA / BizFile filing has already been lodged.
"""


def manifest() -> dict[str, object]:
    return {
        "version": "P2_M05_v0.1",
        "package": "M05 Annual Review Package",
        "templates": TEMPLATES,
    }


def save_template(doc: Document, filename: str) -> None:
    output_path = OUTPUT_DIR / filename
    doc.save(output_path)
    shutil.copy2(output_path, TEMPLATE_DIR / filename)


def main() -> None:
    ensure_dirs()
    save_template(build_agm_package(), TEMPLATES["agm_package"])
    save_template(build_annual_return_package(), TEMPLATES["annual_return_package"])
    save_template(build_checklist(), TEMPLATES["checklist"])
    (OUTPUT_DIR / "M05_field_map.md").write_text(field_map_text(), encoding="utf-8")
    (OUTPUT_DIR / "M05_manifest.json").write_text(json.dumps(manifest(), ensure_ascii=False, indent=2), encoding="utf-8")
    for filename in TEMPLATES.values():
        print(TEMPLATE_DIR / filename)
    print(OUTPUT_DIR / "M05_field_map.md")


if __name__ == "__main__":
    main()
