# P1 M06 Statutory Registers Pack

## Scope

M06 is an initial statutory registers package generated from the P1 incorporation data. It is included in the P1 package as:

- `11_statutory_registers_package.docx`

The template source is:

- `app/doc_templates/p1_standard_v3_part1/11_statutory_registers_package_standard.docx`

## Registers Covered

The first version includes:

- Register of Members
- Register of Directors
- Register of Secretaries
- Register of Registrable Controllers
- Register of Nominee Directors
- Register of Nominee Shareholders, only when nominee shareholder fields are filled
- Register of Share Certificates
- Preparation / review notes

## Data Logic

The package reuses existing P1 data where possible:

- Company data comes from the `公司信息` sheet.
- Officers come from the `人员信息` sheet and common people database resolution.
- Shareholder/member data comes from the `股东与股份` sheet.
- Share certificates use the same shareholder/certificate data as the P1 share certificate documents.

## Default Rules

- Register location defaults to registered office if `register_location` is blank.
- Controllers default to shareholders holding 25% or more of issued shares.
- `is_registrable_controller=Yes` forces a shareholder into RORC.
- `is_registrable_controller=No` excludes a shareholder from the auto controller list.
- Nominee director nominator defaults to the first non-nominee client director.
- If no client director exists, the fallback is first individual shareholder, then authorised representative of corporate shareholder, then client signatory.
- Nominee shareholder register is omitted unless `is_nominee_shareholder` or nominee shareholder nominator/beneficial owner fields are completed.

## New Optional P1 Fields

Company sheet:

- `nominee_director_nominator_name`
- `nominee_director_nominator_id_number`
- `nominee_director_nominator_address`

Shareholder sheet:

- `is_registrable_controller`
- `controller_basis`
- `controller_start_date`
- `is_nominee_shareholder`
- `nominee_shareholder_nominator_name`
- `nominee_shareholder_nominator_id_number`
- `nominee_shareholder_nominator_address`

All new fields are optional. Normal P1 registration data can remain blank and still generate M06.

## Verification

Local DOCX smoke test generated:

- `app/generated/M06_P1_REGISTER_SAMPLE_DOCX_P1_docs/11_statutory_registers_package.docx`

Checks performed:

- P1 DOCX package generation succeeded.
- M06 contains no unresolved template markers.
- M06 tables include members, directors, secretary, controllers, nominee director, certificates and review notes.

PDF conversion could not be completed in the current local environment because the local LibreOffice executable failed before conversion. This appears to be an environment issue, not M06 template logic.
