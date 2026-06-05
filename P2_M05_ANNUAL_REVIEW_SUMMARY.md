# P2 M05 年审包总结

更新时间：2026-06-05  
状态：已接入 M05 v0.1 PDF 生成器。

## 文件组合

M05 最终下载包包含 3 个 PDF：

- `01_M05_agm_documents_package_{company}.pdf`：AGM 文件包。
- `02_M05_annual_return_authorisation_package_{company}.pdf`：Annual Return 授权声明包。
- `99_M05_annual_review_internal_checklist_{company}.pdf`：内部复核清单。

## AGM 文件包内容

- Notice of Annual General Meeting
- Agreement by Members to Shorter Notice
- Attendance Sheet
- Minutes of Annual General Meeting

## Annual Return 授权声明包内容

- Certificate by a Company Limited by Shares under Section 197(1)
- Statement by a Small Company Exempt from Audit Requirements under Section 205C
- Filing of Annual Return authorisation letter to RSIN GROUP PTE. LTD.
- Management Representation

## 触发逻辑

P2 v7 一页式业务单或“快速年审”页中满足以下任一条件：

- `annual_review_required = Yes`
- 已填写 `fye_date`
- 已填写 `agm_date`

生成器状态字段：

- `m05_available = Yes`
- 不会因为纯年审误触发 M01 / M02 / M03 / M04。

## 核心字段

- `company.company_name`
- `company.uen`
- `company.registered_office_address`
- `company.default_document_date`
- `annual_review.fye_date`
- `annual_review.agm_date`
- `annual_review.agm_time`
- `annual_review.agm_place`
- `annual_review.financial_statement_date`
- `annual_review.director_signer_name`
- `annual_review.shareholder_signer_name`
- `annual_review.ar_authorized_signer_name`
- `annual_review.directors_fee`
- `annual_review.directors_remuneration`
- `annual_review.management_rep_letter`

## 签字逻辑

- AGM Notice：默认由第一名董事签发。
- Shorter Notice Consent：所有 member / shareholder signer 签字。
- Attendance Sheet：董事签字人 + 成员签字人列入出席名单。
- AGM Minutes：主席签字，默认第一名董事。
- Section 197 Certificate / AR Authorisation：AR authorised signer，留空时默认第一名董事。
- Section 205C Statement / MRL：所有 director signer 签字。
- 内部清单：不作为客户签署文件。

## 复核边界

- M05 只准备客户签署文件和授权文件。
- M05 不代表 ACRA / BizFile filing 已完成。
- 财报、audit exemption、小公司条件、Annual Return filing 数据仍需人工复核。
- audited / manual accounts status 后续应做更细分模板。

## 压力测试

- 测试表：`outputs/M05_annual_review_stress_input.xlsx`
- 本地生成包：`app/generated/M05-ANNUAL-REVIEW-STRESS_P2_M05_pdf_package.zip`
- 网页测试生成包：`app/generated/JOB-20260605-134926-A74D_P2_M05_pdf_package.zip`
- PDF 视觉检查目录：`outputs/M05_annual_review_pdf_review_v1`
