# P2 M05 年审包总结

更新时间：2026-06-26  
状态：已升级到 M05 v0.4。

## 当前结论

M05 v0.4 已经把年审包简化成更适合内部操作的状态模型：人工和 AI 主要只需要判断公司是普通活跃、休眠，还是已审计。

现在不再用同一套普通 AGM 文本套所有公司。系统会根据年审字段自动切换：

- 普通活跃：正常 AGM 或书面年审文本，财报状态显示为 active / unaudited financial statements。
- 休眠：休眠公司文本，不生成普通 AGM 的 Proxy、Attendance Sheet、AGM Minutes。
- 审计财报：审计财报文本，避免误用 Section 205C 小公司审计豁免声明。
- AGM exempt / dispensed / written resolutions：不生成普通 AGM Notice / Proxy / Attendance / Minutes，改为成员书面确认/书面年审路线。

注意：M05 仍然是内部文件生成器，不代表 ACRA / BizFile filing 已完成，也不替代秘书/法律复核。

## 文件组合

M05 最终下载包包含 2 个 PDF：

- `01_M05_annual_review_signing_package_{company}.pdf`：AGM/书面年审文件 + Annual Return 授权及声明的合并签署包。
- `99_M05_annual_review_internal_checklist_{company}.pdf`：内部复核清单。

## 普通 AGM 路线

当 `agm_status = Held AGM` 或 `agm_route = ordinary_agm`，且不是休眠豁免路线时，AGM 文件包包含：

- Directors' Resolutions in Writing
- Notice of Annual General Meeting of Members
- Agenda / Ordinary Business
- Agreement by Members to Shorter Notice
- Instrument Appointing a Proxy
- Attendance Sheet
- Minutes of an Annual General Meeting

## 休眠 / AGM 豁免 / 书面路线

当字段显示 dormant、AGM exempt、AGM dispensed 或 written resolutions 时，AGM 文件包改为：

- Directors' Resolutions in Writing
- Members' Written Resolution / Consent

不会生成普通 AGM 的 Proxy、Attendance Sheet、AGM Minutes。

## Annual Return 授权及声明包

Annual Return 包包含：

- Annual Return Review Summary
- Certificate by Company Limited by Shares under Section 197(1)
- 动态声明页：
  - 普通小公司：Statement by Small Company Exempt from Audit Requirements
  - 休眠公司：Dormant Company / No Financial Statements Statement
  - 审计财报：Audited Accounts Review Statement
- Filing of Annual Return authorisation letter to RSIN GROUP PTE. LTD.
- Management Representation 或 Dormant Company Representation

## 关键字段

- `accounts_status`：核心字段，只使用 active / dormant / audited。留空或 Auto 按 active 处理。
- `agm_route`：ordinary_agm / written_resolutions / exempt_private_company / dispensed_with_agm / manual。
- `company_activity_status`：兼容/复核字段，通常留空。
- `financial_statements_type`：兼容字段，通常留空。
- `financial_statements_required`：兼容字段，通常留空。
- `audit_exemption_status`：兼容/复核字段，通常留空。
- `agm_status`：Auto / Held AGM / Dispensed with AGM / Exempt from AGM / Written resolutions
- `acra_dormant_relevant_company`：Auto / Yes / No
- `total_assets_under_500k`：Auto / Yes / No
- `iras_tax_status`：Auto / Active / Dormant / Dormant waiver granted / Manual review

默认策略：`accounts_status` 是控制字段；普通年审留空或填 active。只有明确休眠时填 dormant，明确审计时填 audited。其他字段只作兼容和人工复核，不应在普通年审里强行填写。

## 签字逻辑

- Directors' Resolutions：所有 `director_signer_name` 中列出的董事签字。
- 普通 AGM Notice / Minutes：默认第一名董事或主席签署。
- Shorter Notice / Proxy：所有 `shareholder_signer_name` 中列出的 member / shareholder 签字。
- 书面年审 / AGM 豁免路线：成员签字人签署 Members' Written Resolution / Consent。
- Section 197 Certificate / AR Authorisation：由 `ar_authorized_signer_name` 签署；留空时默认第一名董事。
- 动态声明页 / MRL：所有 director signer 签字。
- 内部清单：仅内部复核，不作为客户签署文件。

## 压力测试

- 普通年审测试表：`outputs/M05_annual_review_stress_input_v03.xlsx`
- 休眠年审测试表：`outputs/M05_dormant_annual_review_input.xlsx`
- 审计年审测试表：`outputs/M05_audited_annual_review_input.xlsx`
- 普通年审 PDF 包：`app/generated/M05-ANNUAL-REVIEW-STRESS-V03_P2_M05_pdf_package.zip`
- 休眠年审 PDF 包：`app/generated/M05-DORMANT-REVIEW-V03_P2_M05_pdf_package.zip`
- 审计年审 PDF 包：`app/generated/M05-AUDITED-REVIEW-V03_P2_M05_pdf_package.zip`
- 最新网页测试生成包：`app/generated/JOB-20260608-035209-85A0_P2_M05_pdf_package.zip`
- 最新网页休眠测试生成包：`app/generated/JOB-20260608-035402-96DA_P2_M05_pdf_package.zip`
- PDF 视觉检查目录：`outputs/m05_v03_render_check`

检查结果：

- 普通 AGM 路线：合并签署包 + 内部清单。
- 休眠书面路线：合并签署包 + 内部清单，不生成普通 AGM Proxy、Attendance Sheet、AGM Minutes。
- 审计财报路线：使用 Audited Accounts Review Statement，没有误用 Section 205C 小公司审计豁免声明。
- 文本检查未发现未替换占位符、None、NaN 或 IF 标记残留。
- 休眠包未出现 Proxy、Attendance Sheet、AGM Minutes，也未出现普通 MRL 的 “financial statements have been prepared” 文本。
