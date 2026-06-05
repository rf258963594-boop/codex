# P2 样本文件归并与母版模板方案

最后更新：2026-05-29

## 读取范围

来源文件夹：

`C:\Users\25896\OneDrive\Desktop\变更文件模版\变更文件模版`

本次读取到 15 个 PDF。用户原先提到 14 个，当前文件夹实际多出一份 FYE 变更样本：

- `DR - Change of FYE - LIGHT CONE PTE. LTD.(1).pdf`

这份可归入普通 DR 母版。

## 结论

这些样本不应做成 15 个独立模板。建议归并为 5 个母版模板/文件包：

1. `M01_Combined_DR`：普通董事决议母版
2. `M02_EGM_WR_Takeover_Pack`：EGM / 股东书面决议 / 转入交接母版
3. `M03_Share_Transaction_Pack`：股份转让 / 增资配股母版
4. `M04_Annual_Review_AGM_Pack`：年审 AGM 母版
5. `M05_Strike_Off_Special_Pack`：注销及高风险特殊事项母版

其中 `M01` 和 `M02` 是最核心的两个母版。多数变更不是新增文件，而是在这两个母版里按事件自动出现/隐藏条款。转入属于 `M02` 的一种场景，交接信和辞职信作为 `M02` 转入包里的可选输出，不再单独拆成母版。

## 样本到母版映射

| 样本文件 | 实际内容 | 归并母版 | 处理方式 |
|---|---|---|---|
| DR - Change of Registered Office Address - SUNSEA | 注册地址变更 DR | M01 Combined DR | 做成 `change_registered_office` 条款 |
| DR - Change of Business Activity - TODRINK | 营业范围/SSIC 变更 | M01 Combined DR | 样本混入 EGM/Special wording，重建时改成干净 DR 条款 |
| DR - Change of Officer Particular - TODRINK | 董事/股东资料更新 | M01 Combined DR | 做成 `update_officer_particulars` 条款 |
| DR - Change of FYE - LIGHT CONE | 财年年结日变更 | M01 Combined DR | 做成 `change_fye` 条款 |
| 转股TRANSFER OF SHARES_TRI-X | 股份转让董事批准 | M03 Share Transaction + M01 DR clause | DR 作为转股包中的一个模块，不单独做母版 |
| 增注册资本DR - Authority to Issue Shares - FQ AI | 授权发股、配股、EGM、认购申请 | M02 + M03 | 拆成股东授权、配股 DR、认购申请三个模块 |
| 增注册资本 Form 24 - FQ AI | Return of Allotment / Form 24 | M03 Share Transaction | 只在增资配股时生成，不用于普通转股 |
| EGM SUNSEA | 董事/秘书辞任任命、BizFile 授权、短通知 | M02 EGM/WR Takeover | 拆成 EGM 通知/短通知/会议记录 + 转入可选交接信/辞职信 |
| EGM TODRINK | 董事辞任、BizFile 授权、短通知 | M02 EGM/WR | 做成股东层面事项模块 |
| 转股EGM TRI-X | 文件名像转股，正文是董事辞任 EGM | M02 EGM/WR | 不归为转股；归为董事辞任/股东决议模块 |
| 改名 Change of Company Name - MIRANEX | 公司更名 Special Resolution | M02 EGM/WR | 做成 `change_company_name` special resolution 模块 |
| Letter of Termination and Handover - SUNSEA | 终止旧秘书服务、交接文件 | M02 EGM/WR Takeover | 转入包里的可选交接信 |
| 年审 SINGERIA AGM FYE 31 MARCH 2026 | 完整 AGM 签字包 | M04 Annual Review | 作为年审最终模板基础 |
| 年审 AGM FYE2024 Ah Sheng | ACRA Annual Return/申报资料输出 | M04 Annual Review reference | 不作为签字模板母版，只做字段校验参考 |
| 注销 SGRONGJI STRIKING OFF | 注销/Strike off | M05 Strike-Off | 高风险特殊包，必须人工复核 |

## M01 Combined DR：普通董事决议母版

### 适用事件

- `change_registered_office`
- `change_office_hours`
- `change_business_activity`
- `change_fye`
- `update_director_particulars`
- `update_secretary_particulars`
- `update_shareholder_particulars`
- `appoint_secretary`
- `resign_secretary`
- `appoint_director`，普通任命时
- `resign_director`，普通辞任记录时
- `share_transfer_approval`，作为转股包内 DR 条款
- `share_allotment_approval`，作为增资配股包内 DR 条款
- `bizfile_authorization`

### 模板结构

1. 公司标题
2. Directors' Resolution in Writing
3. 背景/依据：Company's Constitution
4. 自动条款区：
   - Registered Office
   - Office Hours
   - Business Activities / SSIC
   - Financial Year End
   - Officer Particulars
   - Appointment / Resignation of Director
   - Appointment / Resignation of Secretary
   - Share Transfer Approval
   - Share Allotment Approval
   - BizFile Authorization
5. 签字区：董事签署

### 关键字段

- company_name
- uen
- resolution_date
- document_date
- director_signers
- registered_office_old
- registered_office_new
- office_hours_old
- office_hours_new
- ssic_old_primary
- ssic_new_primary
- business_activity_old_primary
- business_activity_new_primary
- ssic_old_secondary
- ssic_new_secondary
- business_activity_old_secondary
- business_activity_new_secondary
- fye_old
- fye_new
- first_accounts_period_start
- first_accounts_period_end
- officer_person_id
- old_id_type
- old_id_number
- old_address
- new_id_type
- new_id_number
- new_address
- appointed_person_id
- resigned_person_id
- effective_date
- bizfile_authorized_firm

## M02 EGM / WR / Takeover Pack：股东层面与转入母版

### 适用事件

- `transfer_in_cooperative`
- `transfer_in_non_cooperative`
- `remove_director`
- `change_company_name`
- `constitution_change`
- `share_allotment_authority`
- `change_share_class`
- `strike_off_shareholder_consent`

### 模板结构

M02 不是单一页面，应作为一个 Pack，有两种输出模式：

- EGM 模式：
  - Notice of EGM
  - Agreement to Shorter Notice
  - Attendance Sheet，可选
  - Minutes of EGM
- WR 模式：
  - Members' Written Resolution
  - Shareholders' Consent / Agreement

转入场景统一归入 M02。转入默认走 EGM/WR 股东授权，再配合董事执行文件；交接信和辞职信不再作为独立母版，只作为 M02 转入包的可选文件：

- Termination / Handover Letter，可选，通常建议生成
- Director Resignation Letter，可选，仅配合辞任时生成
- Secretary Resignation Letter，可选，仅配合辞任时生成
- Records Request Checklist，可选

### 自动条款区

- Appointment / Resignation / Removal of Director
- Appointment / Resignation of Secretary，股东授权层面
- Transfer-in Authorization
- Termination / Handover request
- Optional Director / Secretary Resignation acknowledgement
- Change of Company Name，Special Resolution
- Authority to Issue Shares，S161
- Constitution Change
- Strike-off Shareholder Consent
- BizFile Authorization

### 关键字段

- company_name
- uen
- meeting_date
- meeting_time
- meeting_place
- notice_date
- chairman_name
- shareholders
- shareholder_signers
- ordinary_resolutions
- special_resolutions
- shorter_notice_required
- new_company_name
- old_company_name
- removed_director_person_id
- old_secretary_firm
- new_secretary_firm
- termination_effective_date
- handover_letter_required
- resignation_letter_required
- outgoing_director
- outgoing_secretary
- resignation_effective_date
- cooperative_status
- records_requested
- client_signer
- bizfile_authorized_firm
- share_issue_authority_date
- constitution_change_description

### 转入逻辑

- 配合转入：M02 生成 EGM/WR + 董事执行文件 + 交接信；辞职信按选项生成。
- 不配合转入：M02 生成 EGM/WR + 董事执行文件 + 交接信；不生成旧人员辞职信，改用 removal / replacement wording。
- 交接信由客户董事或授权人签署，但字段仍来自同一个转入事件。

## M03 Share Transaction Pack：股份交易母版

M03 分两种模式：`share_transfer` 和 `share_allotment`。

### share_transfer：股份转让

生成文件：

- Share Transfer Directors' Resolution
- Instrument of Transfer
- Stamp Duty Review Checklist
- Updated Share Certificate
- Internal Register Update Checklist

不生成：

- Form 24

关键字段：

- transferor_shareholder_id
- transferor_name
- transferee_shareholder_id
- transferee_name
- share_class
- shares_transferred
- transfer_date
- instrument_date
- consideration_basis
- cash_consideration
- stamp_duty_required
- nav_required
- old_certificate_no
- new_certificate_no
- director_signers
- transferor_signer
- transferee_signer

### share_allotment：增资配股

生成文件：

- S161 Authority / EGM / WR
- Allotment Directors' Resolution
- Share Application Letter
- Return of Allotment / Form 24
- Share Certificate
- Register Update Checklist

关键字段：

- allottee_person_id
- allottee_name
- share_class
- shares_allotted
- amount_paid_per_share
- total_paid
- currency
- allotment_date
- authority_date
- authority_type
- section_161_required
- post_allotment_total_shares
- post_allotment_paid_up_capital
- certificate_no

## M04 Annual Review / AGM Pack：年审母版

### 推荐样本

使用 `年审SINGERIA MINERAL RESOURCES TRADING...AGM FYE 31 MARCH 2026_signed.pdf` 作为基础升级。

`Ah Sheng` 那份是 Annual Return/申报输出，适合做字段校验参考，不适合作为签字模板基础。

### 生成文件

- Cover Letter / Document List
- Directors' Resolution to Convene AGM
- Notice of AGM
- Agreement to Shorter Notice
- Attendance Sheet
- Minutes of AGM
- Certificate by Company Limited by Shares
- Statement by Dormant Company Exempt from Audit，按条件
- Authorisation for BizFile Services – Annual Filing
- Management Representation Letter
- CDD Form，可后续拆到 KYC 模块

### 关键字段

- company_name
- uen
- registered_office_address
- fye_date
- financial_year_start
- financial_year_end
- agm_date
- agm_time
- agm_place
- agm_route
- accounts_status
- audit_exemption
- dormant_status
- directors_fee
- directors_remuneration
- ar_authorized_signer
- shareholder_signers
- director_signers
- management_rep_letter_required
- cdd_form_required

## M05 Strike-Off / Special Pack：注销和特殊事项母版

### 适用事件

- `strike_off`
- 未来可扩展：
  - share_capital_reduction
  - share_buyback
  - complex_constitution_change
  - compulsory_removal_dispute

### 生成文件

- Directors' Resolution for Strike Off
- Shareholder Consent
- Strike-off Declaration
- Assets/Liabilities/Tax/CPF/Court Proceedings Checklist

### 关键字段

- company_name
- uen
- strike_off_reason
- ceased_business_date
- no_assets_liabilities_confirmation
- shareholder_consent_obtained
- latest_accounts_status
- tax_clearance_status
- cpf_status
- charge_register_status
- court_proceedings_status
- director_signers
- shareholder_signers

## 样本质量问题

重建模板时应修正以下问题：

- 多个 DR 样本混入 EGM / Special Resolution wording，需拆清楚。
- 多个 EGM 样本日期缺失，例如 “on 2026 at 10.00 a.m.”。
- pronoun 错误，例如女性人员仍写 he/his。
- 旧 wording 包括 `Cap. 50`、`Articles of Association`、`Common Seal`，应按当前文件口径和内部法律复核决定是否更新。
- 文件名不完全可靠：`转股EGM TRI-X` 实际是董事辞任 EGM，不是转股文件。
- `Ah Sheng` 年审样本是 AR 输出，不适合作为签字文件模板。

## 与 v3 Excel 的关系

不要先把 15 个样本做成 15 个上传选项。v3 Excel 只需要记录事件和字段：

- Company
- People
- Shareholdings
- ChangeEvents
- ShareTransfers
- ShareAllotments
- AnnualReview
- Registers，待补
- Takeover 字段可并入 ChangeEvents/OutputOptions；不再作为独立母版
- OutputOptions

网站生成时根据事件自动组合：

- 普通事件 -> M01 DR
- 股东层事件 -> M02 EGM/WR
- 转入事件 -> M02 EGM/WR/Takeover + M01 directors' execution clauses
- 转股事件 -> M03 transfer mode + M01 transfer approval clause
- 增资事件 -> M02 S161 + M03 allotment mode + M01 allotment clause
- 年审事件 -> M04
- 注销事件 -> M05

## 建议开发顺序

1. `M01 Combined DR` 已建立第一版母版并接入网页生成：`outputs/P2_standard_templates_v1/M01_combined_directors_resolution_standard.docx`，字段说明见 `outputs/P2_standard_templates_v1/M01_field_map.md`。
2. 再重建 `M02 EGM/WR/Takeover Pack`，含交接信和辞职信可选输出。
3. 再做 `M03 Share Transaction Pack`，先 transfer，再 allotment/Form 24。
4. 再做 `M04 Annual Review / AGM Pack`。
5. 最后做 `M05 Strike-Off / Special Pack`。

这样能最大化复用字段和签字逻辑，避免每个 PDF 做一套孤立模板。
