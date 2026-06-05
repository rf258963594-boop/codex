# P2 公司维护、变更、年审逻辑归档

最后更新：2026-05-29

## 当前状态

- P1 注册模板已重建并接入网站生成。
- P2 目前完成了「导入表格 v3」「上传后自动判断预览」和 `M01 Combined DR` 第一版 DOCX 母版。
- `M01 Combined DR` 已接入网页生成：上传维护/变更年审 v3 表后，如果识别到普通 DR / 转股董事批准 / 配股董事批准，任务页会出现 M01 Word 包生成按钮。
- P2 的完整 Word/PDF 文件模板还没有全部重建；M02-M05 仍先保留为判断预览。
- `M01 Combined DR` 路径：`outputs/P2_standard_templates_v1/M01_combined_directors_resolution_standard.docx`；字段说明：`outputs/P2_standard_templates_v1/M01_field_map.md`。
- 现有 P2 样本来源包括年审、注册地址变更、营业范围变更、人员资料变更、秘书/董事变更、转入交接、转股、增资配股、公司更名、注销等 PDF 样本。

## 核心原则

系统不应先问用户生成 DR 还是 EGM，而应先读取表格中的业务事件：

`变更事件 -> 审批层级 -> 文件包 -> 字段 -> 签字人 -> 生成文件`

网页选项只作为确认或覆盖项，不作为主数据来源。

一个 Excel 应作为一张「业务单」，可以同时包含：

- 年审
- 普通变更
- 转入
- 董事/秘书任免
- 人员身份/地址资料更新
- 股份转让
- 增资配股
- 公司更名
- 注销/strike off 提示

系统上传后应自动预览多个文件包，不要求用户分多次操作。

## 审批层级分类

| 层级 | 典型事项 | 默认文件类型 | 说明 |
|---|---|---|---|
| 董事会事项 | 地址、营业范围、FYE、秘书任免、人员资料更新 | DR | 多个普通事项可合并为一份 Combined Directors' Resolution |
| 股东层面事项 | 公司更名、移除董事、修改 Constitution、重大资本事项 | EGM / WR | 不能用普通 DR 替代 |
| 独立交易文件 | 转股、配股申请、转入交接信/辞职信 | Instrument / Letter / Consent | 转股/配股文件必须单独出；交接信/辞职信作为转入 EGM 包里的可选输出 |
| 年审事项 | AGM、Annual Return 授权、财报呈交 | AGM Pack | 以 SINGERIA 年审样本作为升级基础 |
| 高风险事项 | 注销、强制移除、减资、复杂股权 | Manual Review Pack | 自动判断可以提示，但正式生成前需人工复核 |

## 完整事件类型清单

| event_type | 中文 | 默认文件包 | 默认审批 | 关键字段 |
|---|---|---|---|---|
| change_registered_office | 注册地址变更 | 普通 DR 包 | DR | 旧地址、新地址、生效日、签字董事 |
| change_office_hours | 办公时间变更 | 普通 DR 包 | DR | 旧办公时间、新办公时间、生效日 |
| change_business_activity | 营业范围/SSIC 变更 | 普通 DR 包 | DR | 旧 SSIC、新 SSIC、业务描述、生效日 |
| change_fye | 财年年结日变更 | 普通 DR 包 | DR | 旧 FYE、新 FYE、生效日、理由 |
| update_director_particulars | 董事资料更新 | 普通 DR/Checklist | DR/Checklist | 董事、旧证件/地址/电话、新资料 |
| update_secretary_particulars | 秘书资料更新 | 普通 DR/Checklist | DR/Checklist | 秘书、旧资料、新资料 |
| update_shareholder_particulars | 股东资料更新 | 普通 DR/Checklist | DR/Checklist | 股东、旧资料、新资料 |
| appoint_director | 委任董事 | 董事任免包 | DR 或 EGM/WR | 新董事资料、任命日、本地董事判断 |
| resign_director | 董事辞任 | 董事任免包 | DR | 离任董事、辞任日、辞职信选项 |
| remove_director | 移除董事 | 股东决议包 | EGM/WR | 被移除董事、股东签字、原因、复核 |
| appoint_secretary | 委任秘书 | 秘书任免包 | DR | 新秘书资料、任命日、Form45B/Consent |
| resign_secretary | 秘书辞任 | 秘书任免包 | DR | 离任秘书、辞任日、辞职信选项 |
| transfer_in_cooperative | 配合转入 | 转入包 | EGM/WR + DR | 旧秘书、新秘书、交接日、是否辞职信 |
| transfer_in_non_cooperative | 不配合转入 | 转入包 | EGM/WR + DR | 股东授权、移除/替换逻辑、人工复核 |
| change_company_name | 公司更名 | 公司更名包 | EGM/WR Special Resolution | 旧名、新名、决议日 |
| share_transfer | 股份转让 | 股份转让包 | DR + Instrument | 转让人、受让人、股数、日期、证书号 |
| share_allotment | 增资配股 | 增资配股包 | S161/EGM/WR + DR | 认购人、股数、每股金额、配股日 |
| change_share_class | 股份类别变更 | 股本变更包 | EGM/WR + DR | 旧类别、新类别、股数、权利说明 |
| constitution_change | Constitution 修改 | 股东决议包 | EGM/WR Special Resolution | 修改事项、特殊决议日期 |
| rorc_update | RORC 控制人更新 | Registers 包 | Notice/Register | 控制人、控制方式、控制比例、生效日 |
| rond_update | Nominee Director Register 更新 | Registers 包 | Notice/Register | 挂名董事、nominator、日期 |
| rons_update | Nominee Shareholder Register 更新 | Registers 包 | Notice/Register | 名义股东、实际控制人、日期 |
| annual_review | 年审 | 年审包 | AGM/WR/Exempt | FYE、AGM 日期、财报状态、AR 授权人 |
| strike_off | 注销 | 注销包 | DR + 股东同意 | 无资产负债声明、税务/CPF 状态、股东同意 |

## 文件包清单

### 普通 DR 包

- Combined Directors' Resolution
- BizFile Filing Checklist

适用事项：

- 注册地址变更
- 办公时间变更
- 营业范围/SSIC 变更
- FYE 变更
- 人员资料更新
- 部分董事/秘书任免

### 董事/秘书任免包

- Directors' Resolution
- Consent to Act as Director / Form 45
- Consent to Act as Secretary / Form 45B
- Director Resignation Letter，可选
- Secretary Resignation Letter，可选
- BizFile Filing Checklist

### 转入包

- EGM Notice 或 Members' Written Resolution
- EGM Minutes 或 Written Resolution record
- Directors' Resolution for takeover actions
- Termination / Handover Letter，可选但通常建议生成
- Resignation Letter，可选，仅配合辞任时生成
- Consent Form，新董事/新秘书上任时使用

逻辑：

- 转入统一作为 EGM/WR 的一部分处理，不再单独拆一个 Takeover 母版。
- 配合转入：可生成交接信和辞职信。
- 不配合转入：可生成交接信，但不应生成旧人员辞职信，应使用 removal / replacement wording，并由股东层面授权。
- 股东授权新秘书公司更稳妥，但执行 BizFile/任命秘书仍应由董事层面文件承接。

### 股份转让包

- Share Transfer Directors' Resolution
- Instrument of Transfer
- Stamp Duty Review Checklist
- Updated Share Certificate
- Internal Register Update Checklist

逻辑：

- 普通转股不生成 Form 24。
- Form 24 只用于新股发行/配股。
- EROM 是 ACRA 系统登记结果，不作为文件模板生成。

### 增资配股包

- S161 Authority / EGM 或 Members' Written Resolution
- Allotment Directors' Resolution
- Share Application Letter
- Return of Allotment / Form 24
- Share Certificate
- Register Update Checklist

### 公司更名包

- EGM Notice 或 Members' Written Resolution
- Special Resolution
- BizFile Authorization
- Internal Checklist

### 年审包

建议以 SINGERIA 年审样本为最终版基础，不以 Ah Sheng ACRA Annual Return 输出作为模板基础。

- Cover Letter / Document List
- Directors' Resolution to Convene AGM
- Notice of AGM
- Shorter Notice Consent
- Attendance Sheet
- Minutes of AGM
- Certificate / Statement by Company
- Dormant Company Statement，按条件
- Annual Return Authorization
- Management Representation Letter
- CDD Form，后续可作为 KYC 模块，不一定进入核心年审包

### Registers 包

- RORC Notice to Registrable Controller
- Register of Registrable Controllers update
- ROND Notice / Register of Nominee Directors update
- RONS Notice / Register of Nominee Shareholders update
- Internal Register Checklist

### 注销包

- Directors' Resolution
- Shareholder Consent
- Strike-off Declaration
- Assets/Liabilities/Tax/CPF/Court Proceedings Checklist

应标记为人工复核。

## 表格结构建议

P2 不应拆成年审表和变更表两个完全割裂的系统。建议保留一张「公司维护业务单」：

- Company
- People
- Shareholdings
- ChangeEvents
- ShareTransfers
- ShareAllotments
- AnnualReview
- Registers
- Takeover 字段可并入 ChangeEvents / OutputOptions，不作为独立母版
- OutputOptions
- Rules/Options

## 当前 v3 表格已有

- Company
- People
- Shareholdings
- ChangeEvents
- ShareTransfers
- ShareAllotments
- AnnualReview
- OutputOptions
- 规则说明_Rules
- 字段说明_Options

## 当前 v3 表格应补

### ChangeEvents

- event_subtype
- approval_route_auto
- approval_route_override
- requires_shareholder_approval
- requires_special_resolution
- requires_consent_form
- requires_bizfile_filing
- package_group
- signer_person_id

### People

- old_id_type
- old_id_number
- old_residential_address
- old_phone
- old_email
- new_id_type
- new_id_number
- new_residential_address
- new_phone
- new_email
- resignation_letter_required
- consent_required

### ShareTransfers

- instrument_date
- stamp_duty_required
- consideration_basis
- nav_required
- old_total_shares
- new_total_shares
- old_certificate_cancelled
- new_certificate_required
- transfer_document_signed_date

### ShareAllotments

- authority_type
- section_161_required
- post_allotment_total_shares
- post_allotment_paid_up_capital
- form24_required
- share_application_required
- certificate_required

### Registers

- register_type
- person_id
- controller_type
- control_percentage
- control_method
- notice_date
- effective_date
- register_update_required

### AnnualReview

- agm_exempt_reason
- filing_due_date
- ar_filing_date
- solvency_status
- financial_statement_type
- dormant_period
- director_fee_approval
- xbrl_required

### Transfer-in / Takeover 字段

- old_secretary_firm
- new_secretary_firm
- old_party_cooperative
- handover_required
- handover_letter_date
- records_requested
- old_nominee_director_present
- removal_required

这些字段服务于转入 EGM/WR 包，不单独形成一个 Takeover 模板母版。

## 明确排除或暂缓

- 审计师任免文件：暂不做，通常由外部审计师处理。
- EROM 文件：不生成，仅作为 ACRA/BizFile 完成提示。
- 清盘、法院、诉讼、复杂减资：暂不自动生成正式文件。
- 复杂资本重组、不同类别股份权利变更：可记录事件，但正式模板应后置并人工复核。

## P2 文件模板开发顺序

下一步应按 P1 的方法处理：

1. 读取现有 PDF/Word 样本。
2. 提取正文、签字块、字段、文件结构。
3. 判断每个文件属于哪个文件包。
4. 建立「模板 -> 字段 -> 签字人 -> 触发事件」矩阵。
5. 重建 DOCX 模板，尽量保留原文和专业格式，只升级排版和字段，不简化条款。
6. 渲染检查 PDF/PNG。
7. 再回头调整 Excel v3 字段，保证入口字段能覆盖模板需要。
8. 最后接入网站生成。

## 重要判断

- P2 不应现在直接写生成器。
- 应先重建模板和字段矩阵。
- 表格字段应服务于模板，不应凭空设计太多后续用不到的字段。
- 变更事项可以合并，但法律文件动作不能乱合并：普通 DR 可合并；股东事项要 EGM/WR；转股/配股必须保留独立交易文件；转入的交接信/辞职信作为 EGM/WR 转入包内可选文件。
