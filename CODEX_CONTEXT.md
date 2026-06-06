# 给下一个 Codex 的项目上下文

这个仓库是 RSIN 内部使用的“新加坡秘书公司签字文件生成网站”MVP。用户是业务方，不是技术人员；开发时要优先保持操作逻辑简单、表格好填、文件专业严谨。

## 当前阶段结论

项目已经进入第一阶段可交接状态：

- 本地网站可运行，入口是 `http://127.0.0.1:8088/`。
- 已接入 P1 注册文件包。
- 已接入 P2 的 M01-M05 文件包。
- 已有 Docker / LibreOffice / PDF 转换部署说明。
- 已有 GitHub 私有仓库，用于后续 Codex 接力开发和审核。

当前仓库只包含代码、正式模板、干净导入表、假数据测试表和阶段说明。真实客户上传文件、生成文件、本地数据库、日志和缓存都已被 `.gitignore` 排除。

最新默认值和普通用户流程回归记录见 `DEFAULT_VALUE_REGRESSION_SUMMARY.md`。这轮测试专门覆盖“能留空就留空”的 P1 / P2 M01-M05 输入表，避免只用全字段压力表导致默认逻辑漏测。

## 用户业务目标

用户希望先上线一个内部网站：

1. 上传 Excel 表格。
2. 系统根据表格判断要生成哪些文件。
3. 生成签字用 PDF 文件包。
4. 后续再逐步升级为客户管理后台、在线填写资料、电子签名、手机 App 或云端系统。

用户更关注：

- 表格是否容易填写。
- 文件是否专业、严谨、不要过度简化。
- 多股东、多董事、多签字人的逻辑是否清楚。
- 后续能不能扩展数据库、联系人、公司、人、角色、文件快照。

## 本地运行

推荐入口：

- 双击 `OPEN_WEBSITE.cmd`
- 或使用 `tools/website_control.ps1 start|stop|restart|status`

本地网站：

- `http://127.0.0.1:8088/`

管理员测试账号：

- `admin`
- `admin123`

PDF 生成依赖 LibreOffice。云端部署建议使用 Docker 固定 LibreOffice、中文字体、英文字体。参考：

- `DEPLOYMENT.md`
- `RENDERING.md`

## 当前正式文件包

### P1 注册文件包

已接入 PDF 生成。正式模板目录：

- `app/doc_templates/p1_standard_v3_part1/`

包含：

- First Directors' Resolution
- Form 45
- Form 45B
- Share Certificate
- Secretary Service Agreement
- Nominee Director Agreement
- Form 24
- RORC Notice / Controller document

关键业务逻辑：

- 日期以注册日期为主。
- 输入可用 Excel 表格，网站负责计算财年相关默认值。
- 多董事、多股东、多签字人已考虑。
- Share Certificate 已调整为横版证书风格。
- 不要恢复旧的 counterfoil / receipt 页，除非用户明确要求。

### P2 M01 普通董事决议

模板：

- `app/doc_templates/p2_standard_v1/M01_combined_directors_resolution_standard.docx`

用途：

- 普通公司信息变更。
- 例如注册地址、营业范围、人员任免、个人信息变更等。

关键业务逻辑：

- M01 是普通 DR，适合多项普通变更组合到一份董事决议。
- 签字逻辑默认所有指定董事签字。
- 一页式 V7 表格是当前主要入口。
- 业务范围变更只显示新内容，不强制显示旧内容框。
- 人员任免模板目前以姓名和日期为主，不强制 IC 和地址。

### P2 M02 转入包

模板：

- `app/doc_templates/p2_standard_v1/M02_resolution_package_transfer_in_standard.docx`
- `app/doc_templates/p2_standard_v1/M02_handover_and_resignation_package_standard.docx`

用途：

- 新秘书公司接手客户。
- 更换秘书、董事、注册地址等组合场景。

关键业务逻辑：

- 不使用 `outgoing` 这类显得不专业或不尊重前秘书公司的词。
- 如果是强制移除前秘书/前董事，核心应由股东决议支持。
- 用户希望签字顺序和文件包逻辑更顺：决议类和交接/辞职类分成两个 PDF。
- 多股东、多董事签字已考虑。
- 公司注册地址变更可以和转入一起处理。

### P2 M03 股份转让包

模板：

- `app/doc_templates/p2_standard_v1/M03_share_transfer_directors_resolution_standard.docx`
- `app/doc_templates/p2_standard_v1/M03_instrument_of_transfer_standard.docx`
- `app/doc_templates/p2_standard_v1/M03_updated_share_certificate_standard.docx`
- `app/doc_templates/p2_standard_v1/M03_register_and_stamp_duty_checklist_standard.docx`

用途：

- 股份转让。

关键业务逻辑：

- 普通股份转让不生成 Form 24。
- 每个受让人可以生成更新后的 Share Certificate。
- Share Certificate 风格应与 P1 横版证书尽量一致。
- 内部 checklist 用于复核 register / stamp duty，不是客户主要签字文件。

### P2 M04 增资配股包

模板：

- `app/doc_templates/p2_standard_v1/M04_s161_members_authority_standard.docx`
- `app/doc_templates/p2_standard_v1/M04_allotment_directors_resolution_standard.docx`
- `app/doc_templates/p2_standard_v1/M04_share_application_standard.docx`
- `app/doc_templates/p2_standard_v1/M04_share_certificate_standard.docx`
- `app/doc_templates/p2_standard_v1/M04_return_of_allotment_form24_standard.docx`
- `app/doc_templates/p2_standard_v1/M04_register_update_checklist_standard.docx`

用途：

- 增资配股。

关键业务逻辑：

- Form 24 在 M04 生成，不在普通 M03 股份转让生成。
- 每名认购人可单独生成 Share Application。
- 每名认购人可生成对应 Share Certificate。
- Share Certificate 应与 P1 横版证书风格一致。

### P2 M05 年审包

模板：

- `app/doc_templates/p2_standard_v1/M05_agm_documents_package_standard.docx`
- `app/doc_templates/p2_standard_v1/M05_annual_return_authorisation_package_standard.docx`
- `app/doc_templates/p2_standard_v1/M05_annual_review_checklist_standard.docx`

用途：

- 年审签字文件包。

当前生成 3 份 PDF：

- AGM Documents Package
- Annual Return Authorisation Package
- Internal Annual Review Checklist

关键业务逻辑：

- M05 是年审签字/授权文件包，不是 ACRA 已申报完成证明。
- FYE、AGM 日期、董事签字人、成员签字人、AR 签字人从 V7 快速年审字段读取。
- 财报、审计豁免、AR filing 内容仍需要人工复核。

## 正式导入表

正式导入表已从 `outputs` 复制到仓库内，方便下一个 Codex 使用：

- `templates/import/P1_registration_blank_v3_1.xlsx`
- `templates/import/P1_registration_sample_v3_1.xlsx`
- `templates/import/P2_maintenance_annual_blank_v7.xlsx`
- `templates/import/P2_maintenance_annual_sample_v7.xlsx`

其中 P2 V7 是当前主入口，覆盖 M01-M05 的快速业务单和快速年审。

## 假数据压力测试表

测试表在：

- `tests/fixtures/M01_onepage_stress_input.xlsx`
- `tests/fixtures/M02_transfer_in_registered_office_secretary_director_stress.xlsx`
- `tests/fixtures/M03_share_transfer_stress_input.xlsx`
- `tests/fixtures/M04_share_allotment_stress_input.xlsx`
- `tests/fixtures/M05_annual_review_stress_input.xlsx`

这些是假的压力测试数据，可以用于另一个 Codex 快速生成 PDF 检查。

## 主要代码入口

- `app/server.py`：网站路由、上传、任务页、生成按钮。
- `app/excel_parser.py`：Excel 读取和字段解析。
- `app/rules.py`：根据表格判断 P1 / M01-M05 哪些文件包可生成。
- `app/doc_generator.py`：核心 DOCX/PDF 生成逻辑。
- `app/doc_render.py`：DOCX 转 PDF 渲染逻辑。
- `app/db.py`：本地 SQLite 数据库。
- `app/static/styles.css`：网站样式。

工具脚本：

- `tools/build_p2_m02_templates.py`
- `tools/build_p2_m03_templates.py`
- `tools/build_p2_m04_templates.py`
- `tools/build_p2_m05_templates.py`
- `tools/create_*_stress_test.py`
- `tools/web_*_smoke_test.py`
- `tools/website_control.ps1`

## 已排除、不应上传的内容

`.gitignore` 已排除：

- `outputs/`
- `app/uploads/`
- `app/generated/`
- `app/data/*`，但保留 `app/data/.gitkeep`
- `node_modules/`
- `tools/poppler_extract/`
- 日志文件
- Python 缓存

不要把真实客户 BizFile、护照、签字 PDF、上传缓存、生成 ZIP、SQLite 本地库推到 GitHub。

## 用户偏好的产品逻辑

用户不想要“开发者视角”的复杂按钮堆叠，而是希望：

- 普通用户只看到上传、复核、生成、下载。
- 管理员可以看到模板库、常用人员、后台逻辑、操作编号。
- Excel 表格尽量少填，能默认就默认。
- `Auto` 或空值应尽量代表系统自动判断，减少每格点选。
- 常用秘书、挂名董事应从后台维护和调用。
- 后续数据库要预留联系人、人员、公司、角色、文件快照。

## 重要业务判断记录

- 公司注册文件多数按注册日期出具。
- 新加坡正式文件正文日期倾向使用日/月/年或英文正式日期表达，输入端可用系统识别。
- Form 24 在注册和增资配股里有意义；普通股份转让不默认生成 Form 24。
- RORC 不应过度简化。
- 股份字段需要区分 Number of shares、Issued share capital、Paid-up share capital。
- 默认每股价值常见为 1:1，但也可能出现 issued capital 与 paid-up capital 不一致。
- 未来线上签名和客户后台需要数据库，不能只靠一次性 Excel。

## 下一步建议

优先级建议：

1. 让用户人工审一轮 M05 输出文本，确认年审措辞。
2. 统一 P2 M01-M05 的网页入口，让普通用户更容易理解。
3. 完善管理员模板库：下载模板、上传新版、启用版本、保留历史版本。
4. 优化常用人员管理：新增、编辑、停用、删除、默认身份。
5. 建立数据库草案：contact、person、company、role、shareholding、document_snapshot。
6. 再开始做更多变更类型，不要先把页面复杂化。

## 给下一个 Codex 的工作方式提醒

用户希望你不要做 yes man。遇到业务逻辑不顺时，要指出漏洞，并给出更稳的方案。

不要轻易简化法律/秘书文件正文。可以优化格式和字段，但不要把专业条款缩成一句话。

开发时每次改模板后，最好：

1. 生成假数据表。
2. 生成 PDF。
3. 渲染检查版面。
4. 检查是否有未替换占位符。
5. 再更新阶段总结。
