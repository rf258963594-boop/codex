# 项目阶段封版记录：P1 / M01 / M02

更新时间：2026-06-05  
状态：P1 / M01 / M02 已封版；M03 股份转让包、M04 增资配股包已接入 v0.1；可作为后续年审接入标准。

## 1. 当前网站状态

- 本地网站地址：`http://127.0.0.1:8088/`
- 推荐打开方式：双击项目根目录下的 `OPEN_WEBSITE.cmd`
- 推荐管理脚本：`tools/website_control.ps1 start|stop|restart|status`
- 当前只保留 `8088` 作为正式本地测试端口；旧测试入口不作为日常入口。
- 当前用户逻辑：管理员可看模板库、人员管理、调试编号和内部逻辑；普通用户主要操作上传表格、复核任务、下载 PDF 包。

## 2. 当前正式导入表

### P1 新公司注册

- 空白表：`outputs/AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_空白.xlsx`
- 示例表：`outputs/P1_v31_registration_sample_share_capital.xlsx`

关键规则：
- 日期输入建议 `DD/MM/YYYY`。
- 正式文件日期统一用注册日期。
- FYE 留空时由注册日期自动计算，特殊情况手动覆盖。
- 人员角色尽量留空 / Auto，由股东页、人员页和后台常用人员共同判断。
- 股本按 ACRA 页面逻辑保留 `Number of shares`、`Issued share capital`、`Paid-up share capital` 三个核心口径。
- 默认每股 1:1；如果 issued / paid-up 不等，系统按表格金额生成 partly paid / unpaid 信息。

### P2 维护 / 变更 / 年审

- 空白表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_空白.xlsx`
- 示例表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_示例.xlsx`

关键规则：
- 主入口为“一页式快速业务单”，目标是最少字段先生成文件。
- 常规公司变更用 M01；转入秘书公司用 M02；股份转让用 M03；增资配股用 M04；年审仍待正式接入。
- 个人资料更新、董事/秘书任免、注册地址、营业范围、FYE、办公时间等都可以合并进 M01。
- 董事签字人可在 `director_signer_names` 中用换行、逗号或分号分隔。
- M02 股东/成员签字人使用 `member_signer_names`，多人同样可换行、逗号或分号分隔。
- 交接信客户方签字人使用 `client_signatory_name`；留空时默认使用第一个股东/成员签字人。

## 3. 当前正式模板

### P1 注册文件模板

- 模板目录：`app/doc_templates/p1_standard_v3_part1`
- 已包含：
  - First Directors' Resolution
  - Form 45
  - Form 45B
  - Share Certificate
  - Secretary Service Agreement
  - Nominee Director Agreement
  - Return of Allotment / Form 24
  - RORC Notice

### P2 M01 普通董事决议

- 模板：`app/doc_templates/p2_standard_v1/M01_combined_directors_resolution_standard.docx`
- 当前生成包：`{job_code}_P2_M01_pdf_package.zip`
- 当前文件名：`01_M01_combined_directors_resolution_{company}.pdf`

M01 已支持合并：
- 注册地址变更
- 办公时间变更
- 营业范围 / SSIC 变更，支持 primary / secondary 两项
- FYE 变更
- 董事、秘书、股东或人员身份资料更新
- 董事委任 / 辞任
- 秘书委任 / 辞任
- 股份转让批准：历史能力保留，但默认由 M03 专属包处理，避免重复。
- 增资配股批准：历史能力保留，但默认由 M04 专属包处理，避免重复。
- ACRA / BizFile 授权提交

签字逻辑：
- 默认所有 `director_signer_names` 中列出的董事签字。
- 如果没有签字人，签字栏会保留空白供人工处理。

### P2 M02 转入包

- 决议包模板：`app/doc_templates/p2_standard_v1/M02_resolution_package_transfer_in_standard.docx`
- 交接辞任包模板：`app/doc_templates/p2_standard_v1/M02_handover_and_resignation_package_standard.docx`
- 当前生成包：`{job_code}_P2_M02_pdf_package.zip`

当前生成两份 PDF：
- `01_M02_notice_and_resolutions_{company}.pdf`
  - Notice of Extraordinary General Meeting
  - Consent to Shorter Notice
  - Minutes / Written Resolution of the Members
  - Directors' Resolutions in Writing
- `02_M02_handover_and_resignation_package_{company}.pdf`
  - Handover / Termination Letter
  - Optional resignation letters, only when `generate_resignation_letter = Yes` and personnel actions contain valid resigning director / secretary rows.

签字逻辑：
- Notice：默认由客户方董事或授权人发出。
- Consent to Shorter Notice：所有 `member_signer_names` 股东/成员签字。
- Members' Resolution / Minutes：所有 `member_signer_names` 股东/成员签字。
- Directors' Resolution：所有 `director_signer_names` 董事签字。
- Handover / Termination Letter：`client_signatory_name` 签字；留空时用第一个 member/shareholder signer。
- Resignation Letter：只由对应离任人员签字。

措辞规则：
- 正式文件不显示 `outgoing`、`incoming`、`cooperative`、`non-cooperative`、`transfer-in mode` 等内部词。
- 正式文件使用中性表述，例如 `Existing service provider` 和 `New service provider`。
- 注册地址变更仍属于 M01；M02 只把当前注册地址作为会议/通知地点背景。

## 4. 当前压力测试资料

- P1 股本特殊情况测试表：`outputs/P1_v31_share_capital_partly_paid_test.xlsx`
- P1 股本特殊情况 PDF 包：`app/generated/P1_V31_SHARE_CAPITAL_PARTLY_PAID_TEST_P1_pdf_package.zip`
- M01 v7 一页式压力测试表：`outputs/M01_v7_onepage_stress_input.xlsx`
- M01 v7 一页式压力测试包：`outputs/M01_v7_onepage_stress_delivery_package.zip`
- M01 + M02 转入、换秘书董事、注册地址变更压力测试表：`outputs/M02_M01_transfer_in_registered_office_secretary_director_stress.xlsx`
- 最新网站上传测试任务：`http://127.0.0.1:8088/job?id=15`
- 最新网站 M01 PDF 包：`app/generated/JOB-20260605-014601-0F14_P2_M01_pdf_package.zip`
- 最新网站 M02 PDF 包：`app/generated/JOB-20260605-014601-0F14_P2_M02_pdf_package.zip`

已确认：
- M01 可显示注册地址变更、人员任免和多董事签字。
- M02 第一份文件标题已改为 `NOTICE OF EXTRAORDINARY GENERAL MEETING AND RESOLUTIONS`。
- M02 多股东/成员签字、多董事签字、交接辞任包均可生成。
- PDF 文本未发现未替换占位符。

### P2 M03 股份转让包

- 转股董事决议模板：`app/doc_templates/p2_standard_v1/M03_share_transfer_directors_resolution_standard.docx`
- Instrument of Transfer 模板：`app/doc_templates/p2_standard_v1/M03_instrument_of_transfer_standard.docx`
- 更新股权证书模板：`app/doc_templates/p2_standard_v1/M03_updated_share_certificate_standard.docx`
- 内部复核清单模板：`app/doc_templates/p2_standard_v1/M03_register_and_stamp_duty_checklist_standard.docx`
- 当前生成包：`{job_code}_P2_M03_pdf_package.zip`

M03 已支持：
- 多笔股份转让
- 转让人 / 受让人为自然人或公司名称
- 多董事签字
- 每个受让人一份更新股权证书
- 新证书号留空时显示 `To be assigned`
- Stamp duty / NAV / corporate party / 证书编号复核提醒

签字逻辑：
- 转股董事决议：所有 `director_signer_names` 董事签字，两栏排列。
- Instrument of Transfer：转让人和受让人签字。
- 更新股权证书：保留 Director 和 Company Secretary 签字栏。
- 内部清单：内部复核，不作为客户签署文件。

### P2 M04 增资配股包

- S161 / 股东授权模板：`app/doc_templates/p2_standard_v1/M04_s161_members_authority_standard.docx`
- 配股董事决议模板：`app/doc_templates/p2_standard_v1/M04_allotment_directors_resolution_standard.docx`
- Share Application 模板：`app/doc_templates/p2_standard_v1/M04_share_application_standard.docx`
- 股权证书模板：`app/doc_templates/p2_standard_v1/M04_share_certificate_standard.docx`
- Form 24 模板：`app/doc_templates/p2_standard_v1/M04_return_of_allotment_form24_standard.docx`
- 内部复核清单模板：`app/doc_templates/p2_standard_v1/M04_register_update_checklist_standard.docx`
- 当前生成包：`{job_code}_P2_M04_pdf_package.zip`

M04 已支持：
- 多名认购人
- 全额实缴和部分实缴 / unpaid capital 显示
- S161 / 股东授权、董事配股决议、Share Application、Form 24、股权证书、内部清单
- 股权证书统一为横版全页证书正本，不恢复旧 counterfoil / receipt 页
- 资本汇总不足时生成文件但提示人工复核

签字逻辑：
- S161 / 股东授权：所有 `member_signer_names` 股东/成员签字。
- 配股董事决议：所有 `director_signer_names` 董事签字。
- Share Application：对应认购人签字。
- 股权证书：保留 Director 和 Company Secretary 签字栏。
- Form 24：董事 / 授权签字栏，最终 ACRA / BizFile 申报仍需人工复核。

## 5. 下一阶段优先级

建议顺序：
1. 年审包：快速年审页已预留，但正式模板仍需重新整理。
2. 数据库升级：先预留联系人、公司、人员、股东和任务快照字段；后续再做在线填写和电子签名。

## 6. 封版判断

当前阶段可以作为后续开发基准：
- 表格入口以“竖排、一页式、少填字段、Auto 默认”为标准。
- 网站任务页以“上传后自动判断、显示文件预览、已接入文件显示生成按钮”为标准。
- 模板文件以“Word 模板可维护，最终输出 PDF”为标准。
- 生成逻辑以“表格是主数据来源，网站后台常用人员补充数据”为标准。
