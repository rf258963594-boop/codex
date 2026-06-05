# P2 M02 转入包阶段总结

更新时间：2026-06-05

## 当前状态

- M02 转入包已接入网站生成器。
- 同一张 P2 v7 一页式业务单如果同时识别到普通 DR 和转入，任务页会同时显示 M01 和 M02 两个生成按钮。
- M02 生成入口：`/generate-p2-m02`
- M02 PDF 包命名：`{job_code}_P2_M02_pdf_package.zip`
- 当前 M02 内部版本：`P2_M02_v0.2`

## 触发逻辑

- `transfer_in_required = Yes` 时启用 M02。
- `transfer_in_mode` 只作为内部复核标记，不显示在正式文件中。
- `generate_resignation_letter = Yes` 只表示允许生成辞任信；系统仍需要在“人员任免”里识别到 `resign_director` 或 `resign_secretary` 及人员姓名，才会生成辞任信。

## 生成文件

- `01_M02_notice_and_resolutions_*.pdf`
  - Notice of Extraordinary General Meeting
  - Consent to Shorter Notice
  - Minutes / Written Resolution of the Members
  - Directors' Resolutions in Writing
- `02_M02_handover_and_resignation_package_*.pdf`
  - Handover / Termination Letter
  - Optional resignation letter sections appended in the same PDF when requested and valid.

这样做的目的，是让签字流更顺：第一份集中处理股东/成员授权和董事执行决议，第二份集中处理交接信和实际辞任文件。

## 字段来源

- 公司名称、UEN、注册地址、文件日期来自 P2 一页式业务单。
- 董事签字人来自 `director_signer_names`，支持换行、逗号、分号分隔。
- 股东/成员签字人来自 `member_signer_names` 或 `shareholder_signer_names`，支持多人。
- 交接信客户方签字人来自 `client_signatory_name`；留空时默认使用第一个股东/成员签字人。
- 旧秘书服务方来自 `old_secretary_company`。
- 新秘书服务方来自 `new_secretary_company`，默认可用 `RSIN GROUP PTE. LTD.`。

## 签字逻辑

- Notice：默认由客户方董事或授权人发出。
- Consent to Shorter Notice：所有股东/成员签字。
- Members' Resolution / Minutes：所有股东/成员签字。
- Directors' Resolution：所有董事签字。
- Handover / Termination Letter：客户方签字人签字。
- Optional Resignation Letter：对应离任董事/秘书本人签字。

## 模板路径

- `app/doc_templates/p2_standard_v1/M02_resolution_package_transfer_in_standard.docx`
- `app/doc_templates/p2_standard_v1/M02_handover_and_resignation_package_standard.docx`
- 字段说明：`outputs/P2_standard_templates_v1/M02_field_map.md`

## 测试记录

- 转入 + 换秘书董事 + 注册地址变更压力测试表：`outputs/M02_M01_transfer_in_registered_office_secretary_director_stress.xlsx`
- 底层生成测试 PDF 包：
  - `app/generated/M02-MULTI-SIGNER-ADDRESS-STRESS_P2_M01_pdf_package.zip`
  - `app/generated/M02-MULTI-SIGNER-ADDRESS-STRESS_P2_M02_pdf_package.zip`
- 网站上传测试任务：`http://127.0.0.1:8088/job?id=15`
- 网站测试 PDF 包：
  - `app/generated/JOB-20260605-014601-0F14_P2_M01_pdf_package.zip`
  - `app/generated/JOB-20260605-014601-0F14_P2_M02_pdf_package.zip`

已确认：
- M01 可显示注册地址变更。
- M02 文件标题已从内部化的 resolution package wording 改为正式的 `NOTICE OF EXTRAORDINARY GENERAL MEETING AND RESOLUTIONS`。
- M02 多股东/成员签字和多董事签字均可渲染。
- PDF 内未发现未替换占位符。
- 正式 PDF 内未发现 `outgoing`、`incoming`、`cooperative`、`non-cooperative`、`transfer-in mode` 等不适合正式文件的内部词。

## 下一步建议

1. 接入 M03 股份转让包：Instrument of Transfer、转股董事决议、更新股权证书、登记册/印花税复核清单。
2. 接入 M04 增资配股包：S161/股东授权、董事配股决议、Form 24、股权证书。
3. 整理并接入年审正式模板。
