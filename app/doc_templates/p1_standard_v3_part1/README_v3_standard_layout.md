# P1 Standard Template Pack v3

This pack is the standardized-layout version. It uses v2/source content but rebuilds the Word layout so the files are easier to batch-generate and visually review.

## Rules

- Preserve clauses and signed-document content; do not summarize agreement text.
- Use clean Word paragraphs/tables instead of manual spacing, floating objects, hidden text or old sample artifacts.
- Keep placeholders stable and readable.
- Treat unresolved legal or commercial judgment as review notes, not silent deletions.

## Files

- `01_first_directors_resolution_standard.docx` - First Directors Resolution
- `02_director_consent_form45_standard.docx` - Consent to Act as Director / Form 45
- `03_secretary_consent_form45b_standard.docx` - Consent to Act as Secretary / Form 45B
- `04_share_certificate_standard.docx` - Share Certificate
- `05_secretary_service_agreement_standard.docx` - Secretary Service Agreement
- `06_nominee_director_agreement_standard.docx` - Nominee Director Agreement
- `07_return_of_allotment_form24_standard.docx` - Return of Allotment of Shares / Form 24
- `08_rorc_notice_controller_standard.docx` - RORC Notice to Registrable Controller

## Deferred To P2

- `signature_record_attachment` - Online e-signature audit attachment; keep for P2 when e-signing is enabled.
- `statutory_registers` - Registers pack is useful later but not complete/mandatory for the first internal file generator.

## Main Placeholders

- `{{company.company_name}}`: Company legal name.
- `{{company.uen}}`: Company registration number/UEN; can be blank before incorporation.
- `{{company.incorporation_date}}`: Date of incorporation or appointment effective date.
- `{{company.registered_office_address}}`: Registered office address.
- `{{company.register_location}}`: Place where register of members and index is kept; default can be registered office.
- `{{company.first_directors_names}}`: Derived list of first director names for old resolution sentence.
- `{{company.director_signature_blocks}}`: Derived signature lines for all directors in the first directors resolution.
- `{{company.subscriber_share_lines}}`: Derived multi-line subscriber/share table text for old resolution layout.
- `{{company.share_par_value}}`: Legacy compatibility value; current templates prefer issued/paid-up share capital fields.
- `{{company.fye_month}}`: Financial year end month, auto-derived from incorporation date unless overridden.
- `{{company.first_financial_period_start}}`: First accounts period start, defaults to incorporation date and renders as a long date.
- `{{company.first_financial_period_end}}`: First accounts period end, auto-derived as the last day of the month before the incorporation month unless overridden.
- `{{director.*}}`: One director context for per-director documents.
- `{{secretary.*}}`: One secretary context for per-secretary documents.
- `{{shareholder.*}}`: One shareholder/share certificate context.
- `{{provider.*}}`: Website setting for RSIN/service provider details.
- `{{client_signatory.*}}`: Client authorised signatory context.
- `{{nominee_director.*}}`: Nominee director context.
- `{{signature.*}}`: Signature date fields, including day/month/year split for old legal wording.
- `{{company.share_currency}}`: Share currency; default usually SGD.
- `{{company.subscriber_share_totals}}`: Derived total/share-count cell used by the v3 subscriber shares table.
- `{{company.issued_share_capital}}`: Derived total issued share capital from shareholder rows; defaults to shares for normal 1:1 ordinary shares.
- `{{company.paid_up_capital}}`: Derived total paid-up share capital from shareholder rows; defaults to issued share capital when blank.
- `{{company.unpaid_share_capital}}`: Derived unpaid share capital; used as a review flag when issued capital is higher than paid-up capital.
- `{{company.amount_paid_per_share}}`: Derived paid amount per share for Form 24.
- `{{company.amount_due_per_share}}`: Derived amount due per share for Form 24; `-` when fully paid.
- `{{shareholders[]}}`: Repeating shareholder rows used by Form 24 and RORC; Form 24 allottees and shareholder signature blocks repeat from this list.
- `{{shareholder.date_of_birth}}`: Derived from the matched People row when the shareholder is an individual.
- `{{company.shareholder_signature_blocks}}`: Derived Form 24 signature lines, one block for each shareholder/allottee.
- `{{shareholder.issued_share_capital}}`: Shareholder-level issued share capital.
- `{{shareholder.paid_up_share_capital}}`: Shareholder-level paid-up share capital.
- `{{shareholder.unpaid_share_capital}}`: Shareholder-level unpaid share capital.
- `{{shareholder.paid_status_text}}`: Share certificate wording, e.g. fully paid or partly paid details.
- `{{shareholder.form24_allotment_text}}`: Form 24 allotment summary with shares, issued capital, paid-up capital and unpaid capital when applicable.
