# P1 Rebuilt Template Pack

This folder contains rebuilt DOCX templates for the first internal website launch. The original template folder was read only; no original file was changed.

## Template files

- `01_first_directors_resolution.docx` - First Directors Resolution (one_per_company)
- `02_director_consent_legacy_form45.docx` - Consent to Act as Director / legacy Form 45 (one_per_director)
- `03_secretary_consent_legacy_form45b.docx` - Consent to Act as Secretary / legacy Form 45B (one_per_secretary)
- `04_share_certificate.docx` - Share Certificate (one_per_shareholder_or_certificate)
- `05_secretary_service_agreement.docx` - Secretary Service Agreement (one_per_company)
- `06_nominee_director_agreement.docx` - Nominee Director Agreement (one_per_nominee_director_relationship)
- `07_signature_record_attachment.docx` - Signature Record Attachment (append_to_each_signed_document)

## Placeholder convention

- Single value: `{{company.company_name}}`
- List loop: `{% for director in directors %}` ... `{% endfor %}`
- Keep internal rule names stable; old names such as Form 45/45B are display labels only.

## Fields

### Company

| Placeholder | Meaning | Source / default |
|---|---|---|
| `{{company.company_name}}` | Company legal name | Company sheet: company_name |
| `{{company.uen}}` | Company registration number/UEN | BizFile or post-incorporation value |
| `{{company.company_type}}` | Company type | Default: Private Company Limited by Shares |
| `{{company.incorporation_date}}` | Date of incorporation or document date | Company sheet |
| `{{company.registered_office_address}}` | Registered office | Company sheet |
| `{{company.register_location}}` | Place where registers are kept | Default: registered office |
| `{{company.office_hours}}` | Registered office hours | Default working hours |
| `{{company.fye}}` | Financial year end | Company sheet |
| `{{company.first_financial_period_start}}` | First financial period start | Calculated or manual |
| `{{company.first_financial_period_end}}` | First financial period end | Calculated or manual |
| `{{company.share_currency}}` | Share currency | Default: SGD |
| `{{company.share_class}}` | Share class | Default: Ordinary |
| `{{company.share_par_value}}` | Per-share value where used by legacy wording | Default: 1 |
| `{{company.paid_up_capital}}` | Total paid-up capital | Calculated from Shareholders |
| `{{company.has_common_seal}}` | Whether common seal wording should be kept | Default: No / review |

### People

| Placeholder | Meaning | Source / default |
|---|---|---|
| `{{director.full_name}}` | Director full legal name | People.full_name |
| `{{director.id_type}}` | Passport/NRIC/FIN | People.id_type |
| `{{director.id_number}}` | ID number | People.id_number |
| `{{director.nationality}}` | Nationality | People.nationality |
| `{{director.residential_address}}` | Residential address | People.residential_address |
| `{{director.email}}` | Email | People.email |
| `{{director.phone}}` | Phone | People.phone |
| `{{director.appointment_date}}` | Director appointment/effective date | People.appointment_date |
| `{{director.signature_date}}` | Director signature date | Generated or manual |
| `{{secretary.full_name}}` | Secretary full legal name | People.full_name |
| `{{secretary.id_type}}` | Passport/NRIC/FIN | People.id_type |
| `{{secretary.id_number}}` | ID number | People.id_number |
| `{{secretary.nationality}}` | Nationality | People.nationality |
| `{{secretary.residential_address}}` | Residential address | People.residential_address |
| `{{secretary.email}}` | Secretary email | People.email or common person setting |
| `{{secretary.phone}}` | Secretary phone | People.phone or common person setting |
| `{{secretary.appointment_date}}` | Secretary appointment/effective date | People.appointment_date |
| `{{secretary.signature_date}}` | Secretary signature date | Generated or manual |
| `{{nominee_director.full_name}}` | Nominee director name | Common people or People.full_name |
| `{{nominee_director.id_number}}` | Nominee director ID number | Common people or People.id_number |
| `{{nominee_director.signature_date}}` | Nominee director signature date | Generated or manual |
| `{{client_signatory.full_name}}` | Client signing person | People where signing_required=Yes |
| `{{client_signatory.signature_date}}` | Client signatory signature date | Generated or manual |
| `{{secretary_or_director.full_name}}` | Second certificate signer | Website chooses secretary or another director |

### Shareholders

| Placeholder | Meaning | Source / default |
|---|---|---|
| `{{shareholder.shareholder_name}}` | Person or corporate shareholder display name | Derived |
| `{{shareholder.person_full_name}}` | Individual shareholder name | Shareholders.person_full_name |
| `{{shareholder.person_id_number}}` | Individual ID number | Shareholders.person_id_number |
| `{{shareholder.corporate_name}}` | Corporate shareholder name | Shareholders.corporate_name |
| `{{shareholder.corporate_registration_number}}` | Corporate registration number | Shareholders.corporate_registration_number |
| `{{shareholder.corporate_registered_address}}` | Corporate registered address | Shareholders.corporate_registered_address |
| `{{shareholder.share_class}}` | Share class | Shareholders.share_class |
| `{{shareholder.shares}}` | Number of shares | Shareholders.shares |
| `{{shareholder.paid_amount}}` | Paid amount | Shareholders.paid_amount |
| `{{shareholder.currency}}` | Currency | Shareholders.currency |
| `{{shareholder.certificate_no}}` | Share certificate number | Generated by website |
| `{{shareholder.folio_no}}` | Folio number | Generated or manual |
| `{{shareholder.distinctive_numbers}}` | Distinctive share numbers | Generated or manual |
| `{{shareholder.allotment_transfer_no}}` | Allotment or transfer number | Generated or manual |
| `{{shareholder.shareholder_address}}` | Shareholder display address | Derived from individual/corporate address |
| `{{shareholder.remarks}}` | Certificate/counterfoil remarks | Generated or manual |

### Provider and signing

| Placeholder | Meaning | Source / default |
|---|---|---|
| `{{provider.name}}` | Corporate secretary service provider | Website setting |
| `{{provider.uen}}` | Provider UEN | Website setting |
| `{{provider.address}}` | Provider address | Website setting |
| `{{provider.email}}` | Provider email | Website setting |
| `{{provider.authorised_signatory_name}}` | Provider signing person | Website setting |
| `{{generation.prepared_by}}` | Internal preparer | Generation.prepared_by |
| `{{generation.signing_mode}}` | Signing mode | Generation.signing_mode |
| `{{signature.date}}` | Signature date | Generated or manual |
| `{{signature.day}}` | Signature day for legacy wording | Derived from signature.date |
| `{{signature.month_year}}` | Signature month/year for legacy wording | Derived from signature.date |
| `{{signature_record.document_title}}` | Signed document title | Post-signing data |
| `{{signature_record.envelope_id}}` | E-sign envelope/document ID | Post-signing data |
| `{{signature_record.generated_at}}` | Audit record generated timestamp | Post-signing data |
| `{{signature_record.checksum}}` | Digital fingerprint/checksum | Post-signing data |
| `{{signature_record.signers}}` | Signer audit table | Post-signing data |
| `{{signer.full_name}}` | Signer name inside signature audit table | Post-signing data |
| `{{signer.email}}` | Signer email inside signature audit table | Post-signing data |
| `{{signer.party_id}}` | Signer party ID | Post-signing data |
| `{{signer.status}}` | Signer status | Post-signing data |
| `{{signer.signed_at}}` | Signer completion timestamp | Post-signing data |
| `{{event.timestamp}}` | Audit event timestamp | Post-signing data |
| `{{event.action}}` | Audit event action | Post-signing data |
| `{{event.ip_address}}` | Audit event IP address | Post-signing data |
| `{{event.details}}` | Audit event details | Post-signing data |
