# 新加坡秘书公司文件生成器 MVP

本项目是本地运行的内部文件生成网站。当前阶段以“上传 Excel 资料表，生成签字 PDF 包”为主，不做客户后台和电子签名。

## 当前正式入口

- 本地网站端口：`http://127.0.0.1:8088/`
- 推荐启动：双击 `OPEN_WEBSITE.cmd`
- 保持网站运行：启动窗口需要保持打开；关闭窗口后网站会停止。
- 管理脚本：`tools/website_control.ps1 start|stop|restart|status`

目前只保留一个正式本地端口：`8088`。其他旧测试入口不作为日常使用入口。

## 当前可用功能

- P1 新公司注册：已接入 PDF 生成。
- P2 M01 普通董事决议：已接入 PDF 生成。
- P2 M02 转入包：已接入两份签字 PDF 生成。
- P2 M03 股份转让包：已接入 PDF 生成。
- P2 M04 增资配股包：已接入 PDF 生成。
- P2 其他文件组：年审先保留判断和预览逻辑，后续接入模板。
- 年审：v7 表格已预留快速年审页，正式文件模板仍待下一阶段接入。

## 当前正式导入表

- P1 空白表：`outputs/AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_空白.xlsx`
- P1 示例表：`outputs/P1_v31_registration_sample_share_capital.xlsx`
- P2/M01 空白表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_空白.xlsx`
- P2/M01 示例表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_示例.xlsx`

网站普通用户首页只显示以上正式导入表。管理员后台默认也只显示这些入口；旧版、测试和辅助入口需要点“显示旧版 / 测试 / 辅助入口”才会展开。

## 当前正式文件模板

- P1 注册文件模板目录：`app/doc_templates/p1_standard_v3_part1`
- M01 普通董事决议模板：`app/doc_templates/p2_standard_v1/M01_combined_directors_resolution_standard.docx`
- M02 转入决议包模板：`app/doc_templates/p2_standard_v1/M02_resolution_package_transfer_in_standard.docx`
- M02 交接辞任包模板：`app/doc_templates/p2_standard_v1/M02_handover_and_resignation_package_standard.docx`
- M03 转股董事决议模板：`app/doc_templates/p2_standard_v1/M03_share_transfer_directors_resolution_standard.docx`
- M03 Instrument of Transfer 模板：`app/doc_templates/p2_standard_v1/M03_instrument_of_transfer_standard.docx`
- M03 更新股权证书模板：`app/doc_templates/p2_standard_v1/M03_updated_share_certificate_standard.docx`
- M03 内部复核清单模板：`app/doc_templates/p2_standard_v1/M03_register_and_stamp_duty_checklist_standard.docx`
- M04 S161 / 股东授权模板：`app/doc_templates/p2_standard_v1/M04_s161_members_authority_standard.docx`
- M04 配股董事决议模板：`app/doc_templates/p2_standard_v1/M04_allotment_directors_resolution_standard.docx`
- M04 Share Application 模板：`app/doc_templates/p2_standard_v1/M04_share_application_standard.docx`
- M04 股权证书模板：`app/doc_templates/p2_standard_v1/M04_share_certificate_standard.docx`
- M04 Form 24 模板：`app/doc_templates/p2_standard_v1/M04_return_of_allotment_form24_standard.docx`
- M04 内部复核清单模板：`app/doc_templates/p2_standard_v1/M04_register_update_checklist_standard.docx`

P1 已包含 First DR、Form 45、Form 45B、Share Certificate、Secretary Agreement、Nominee Director Agreement、Form 24、RORC。

## 已归档/备用内容

`outputs` 里还有不少旧版模板、压力测试表、渲染检查目录和历史产物，例如 v2/v3/v4/v5 导入表、P1 旧模板包、M01 压力测试包等。它们暂时不删除，用于回溯和排查；但不作为日常操作入口。

## 上线提醒

本地 Word 转 PDF 依赖 LibreOffice。云端建议使用 Docker 固定安装 LibreOffice、中文字体、英文字体和 Poppler，避免服务器环境和本机不一致。

## M02 转入包更新

- P2 M02 转入包已接入 PDF 生成。
- 触发字段：P2 v7 一页式业务单中 `transfer_in_required = Yes`。
- 默认生成两份 PDF：`01_M02_notice_and_resolutions_*.pdf` 和 `02_M02_handover_and_resignation_package_*.pdf`。
- 第一份 PDF 合并 Notice of EGM、Consent to Shorter Notice、Members' Resolution / Minutes、Directors' Resolution。
- 第二份 PDF 合并 Handover / Termination Letter；当 `generate_resignation_letter = Yes` 且“人员任免”中存在离任董事/秘书时，辞任信附在同一 PDF 后面。
- 测试表：`outputs/M02_transfer_in_stress_input.xlsx`
- 测试 PDF 包：`app/generated/M02-V02-STRESS_P2_M02_pdf_package.zip`

## 阶段封版与 P2 文件包规划

- P1 / M01 / M02 当前封版记录：`PROJECT_STAGE_FREEZE_2026-06-05.md`
- M03 股份转让包分析：`P2_M03_SHARE_TRANSFER_ANALYSIS.md`
- M04 增资配股包总结：`P2_M04_SHARE_ALLOTMENT_SUMMARY.md`
- M03 字段逻辑表：`outputs/P2_M03_股份转让字段逻辑表_v0.1.xlsx`

## M03 股份转让包更新

- P2 M03 已接入 PDF 生成。
- 触发逻辑：P2 v7 一页式业务单中“股份转让”区存在有效转股行，且 `generate` 不是 No。
- 默认生成：转股董事决议、Instrument of Transfer、每个受让人的更新股权证书、内部 Register / Stamp Duty Checklist。
- 普通股份转让不生成 Form 24；Form 24 仍保留给 M04 增资配股。
- M03 / M04 股权证书统一为横版全页证书正本，不恢复旧版 counterfoil / receipt 页。
- 压力测试表：`outputs/M03_share_transfer_stress_input.xlsx`
- 压力测试 PDF 包：`app/generated/M03-SHARE-TRANSFER-STRESS_P2_M03_pdf_package.zip`

## M04 增资配股包更新

- P2 M04 已接入 PDF 生成。
- 触发逻辑：P2 v7 一页式业务单中“增资配股”区存在有效配股行，且 `generate` 不是 No。
- 默认生成：S161 / 股东授权、配股董事决议、每名认购人的 Share Application、每名认购人的横版股权证书、Form 24、内部 Register Update Checklist。
- Form 24 只在 M04 增资配股包生成，不在普通股份转让 M03 生成。
- 压力测试表：`outputs/M04_share_allotment_stress_input.xlsx`
- 压力测试 PDF 包：`app/generated/M04-SHARE-ALLOTMENT-STRESS_P2_M04_pdf_package.zip`
