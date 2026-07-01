# M06 Strike-off Package - Field Map

M06 is the P2 strike-off / company closure package.

## Generated files

| File | Trigger | Signer | Notes |
|---|---|---|---|
| Directors' Resolution for Strike-off | `strike_off_required = Yes` | Director signer(s) | Approves the strike-off application and authorises filing. |
| Shareholder Consent to Strike-off | `strike_off_required = Yes` | Each member/shareholder signer | One consent file per member/shareholder signer. |
| Director Declaration | `strike_off_required = Yes` | Declaration signer, default first director signer | Confirms no assets, liabilities, charges, proceedings and shareholder consent. |

## Primary input fields

| Field | Source | Used in |
|---|---|---|
| `company.company_name` | P2 one-page sheet / web form | All M06 documents |
| `company.uen` | P2 one-page sheet / web form | All M06 documents |
| `company.default_document_date` | P2 one-page sheet / web form | Document and declaration date |
| `company.director_signer_names` | P2 one-page sheet / web form | Directors' resolution and default declaration signer |
| `company.member_signer_names` | P2 one-page sheet / web form | Shareholder consent letters |
| `shareholdings.shareholder_name` / `shareholdings.shares` | P2 shareholder status sheet, if available | Shareholder consent shareholding line |
| `strike_off_required` | P2 one-page sheet / web form | Enables M06 |
| `strike_off_cessation_date` | P2 one-page sheet / web form | If blank, wording says the company has not commenced business since incorporation. If filled, wording says the company ceased business from this date. |
| `strike_off_declaration_signer_name` | P2 one-page sheet / web form | Optional override for declaration signer |

## Review guardrail

M06 prepares signing documents only. It does not verify ACRA, IRAS, CPF, charge register, bank account, creditor or litigation status. These remain manual review points before filing.
