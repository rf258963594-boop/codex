# P1 Preserved Template Pack v2

This version replaces the first simplified rebuild. It keeps the old template content and order as much as possible, while replacing sample data with clear placeholders.

## Files

- `01_first_directors_resolution_preserved.docx` - First Directors Resolution (source_docx_fieldized)
- `02_director_consent_form45_preserved.docx` - Consent to Act as Director / Form 45 (pdf_reconstructed_with_current_acra_wording)
- `03_secretary_consent_form45b_preserved.docx` - Consent to Act as Secretary / Form 45B (pdf_reconstructed_with_current_acra_wording)
- `04_share_certificate_preserved.docx` - Share Certificate (source_docx_fieldized)
- `05_secretary_service_agreement_preserved.docx` - Secretary Service Agreement (source_docx_fieldized_minor_text_cleanup)
- `06_nominee_director_agreement_preserved.docx` - Nominee Director Agreement (pdf_reconstructed_preserving_full_clause_structure)
- `07_signature_record_attachment_preserved.docx` - Signature Record Attachment (source_docx_fieldized_sample_data_removed)

## Upgrade Rules

- Do not summarize or shorten agreement clauses.
- Replace sample data with placeholders.
- Keep old display names such as Form 45/Form 45B, but use stable template IDs internally.
- Put legal/commercial uncertainty in review notes instead of silently changing signed terms.

## Main Placeholders

- `{{company.company_name}}`: Company legal name.
- `{{company.uen}}`: Company registration number/UEN; can be blank before incorporation.
- `{{company.incorporation_date}}`: Date of incorporation or appointment effective date.
- `{{company.registered_office_address}}`: Registered office address.
- `{{company.register_location}}`: Place where register of members and index is kept; default can be registered office.
- `{{company.first_directors_names}}`: Derived list of first director names for old resolution sentence.
- `{{company.subscriber_share_lines}}`: Derived multi-line subscriber/share table text for old resolution layout.
- `{{company.share_par_value}}`: Legacy wording value per share; default usually 1.
- `{{company.fye}}`: Financial year end.
- `{{company.first_financial_period_start}}`: First accounts period start.
- `{{company.first_financial_period_end}}`: First accounts period end.
- `{{director.*}}`: One director context for per-director documents.
- `{{secretary.*}}`: One secretary context for per-secretary documents.
- `{{shareholder.*}}`: One shareholder/share certificate context.
- `{{provider.*}}`: Website setting for RSIN/service provider details.
- `{{client_signatory.*}}`: Client authorised signatory context.
- `{{nominee_director.*}}`: Nominee director context.
- `{{signature.*}}`: Signature date fields, including day/month/year split for old legal wording.
- `{{signature_record.*}}`: Post-signing e-signature audit data.

## Review Notes

- These are operational templates, not legal advice. Final production use should be checked by a Singapore-qualified professional or your internal reviewer.
- Old templates still refer to Companies Act (Chapter 50) / Cap. 50 in places because the ACRA Form 45/Form 45B PDFs also display that legacy label. Do not bulk change this without legal review.
- The Nominee Director Agreement includes broad bilingual terms from the old PDF. Commercial terms such as fees, deposit, notice, local bank support and liability scope should be confirmed before use.