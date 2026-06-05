# P1 / M01 阶段总结

更新时间：2026-06-05

## 当前网站入口

- 本地测试地址：`http://127.0.0.1:8088/`
- 推荐启动：双击 `OPEN_WEBSITE.cmd`
- 管理脚本：`tools/website_control.ps1 start|stop|restart|status`
- 当前只保留 `8088` 作为正式本地测试端口。

## 当前正式导入表

- P1 空白表：`outputs/AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_空白.xlsx`
- P1 示例表：`outputs/P1_v31_registration_sample_share_capital.xlsx`
- P2/M01 空白表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_空白.xlsx`
- P2/M01 示例表：`outputs/AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_示例.xlsx`

## P1 新公司注册

### 表格逻辑

- 使用 v3.1 人性化竖排 Auto 表。
- 公司信息、人员信息、股东与股份、输出设置分开。
- 日期输入建议 `DD/MM/YYYY`。
- P1 全套文件日期统一使用注册日期。
- FYE 默认由注册日期自动计算，特殊情况可覆盖。
- 人员角色尽量用 Auto，系统按人员页、股东页、后台常用人员自动判断。
- 常用秘书、挂名董事可从后台常用人员调用。

### 股本逻辑

按 ACRA 页面简化为：
- `Number of shares`
- `Issued share capital`
- `Paid-up share capital`

默认只填股数即可，系统按 1:1 自动补 issued 和 paid-up。特殊情况可手动填 issued / paid-up：
- 100 万股、issued 100 万、paid-up 1：系统会计算 unpaid，并在 Form 24 和 Share Certificate 中显示未缴金额。
- 100 万股、issued 10 万、paid-up 10 万：系统会显示每股 0.10。

### P1 已生成文件

- `01_first_directors_resolution`：每家公司一份，所有董事签。
- `02_director_consent_form45`：每名董事一份。
- `03_secretary_consent_form45b`：每名秘书一份。
- `04_share_certificate`：每名股东一份，证书正文已改为证书样式；fully paid / partly paid 按股本计算。
- `05_secretary_service_agreement`：服务方 + 客户方代表签。
- `06_nominee_director_agreement`：每名挂名董事一份，挂名董事 + 客户方代表签。
- `07_return_of_allotment_form24`：公司层面一份，列出所有股东配股资料。
- `08_rorc_notice_controller`：每名控制人一份。

### P1 签字逻辑

- First DR：所有董事签。
- Form 45 / Form 45B：对应董事/秘书本人签。
- Form 24：股东信息来自股东页；多股东会列多个股东。
- 客户方签字人默认股东 1；可由 `client_signatory_person_id` 指定。
- Share Certificate 董事签字栏不强行写挂名董事；优先客户董事逻辑，具体签署仍建议人工复核。

## P2 M01 普通董事决议

### 表格逻辑

- 使用 v7 一页式快速业务单。
- 目标是“最少信息先生成文件”，不是完整客户资料库表。
- 适合普通 DR 相关事项：公司名称、UEN、注册地址、变更事项、董事签字人等。
- 快速年审页已预留，但年审正式文件仍待接入。

### M01 已支持的 DR 内容

- 注册地址变更
- 办公时间变更
- 营业范围变更，支持 primary / secondary 两项
- FYE 变更
- 董事/秘书/股东/人员身份信息更新
- 董事委任
- 董事辞任
- 秘书委任
- 秘书辞任
- 股份转让批准：历史模板能力保留，但当前默认由 M03 股份转让包处理，避免 M01/M03 重复生成同一项转股批准。
- 增资配股批准：历史模板能力保留，但当前默认由 M04 增资配股包处理，避免 M01/M04 重复生成同一项配股批准。
- 增资配股批准
- BizFile / ACRA 授权提交

### M01 签字逻辑

- 默认所有当前董事签。
- 一页式表格中可用逗号、分号或换行指定董事签字人。
- 当前生成结果是 M01 普通董事决议 PDF 包。

## 当前正式模板路径

- P1 模板目录：`app/doc_templates/p1_standard_v3_part1`
- M01 模板：`app/doc_templates/p2_standard_v1/M01_combined_directors_resolution_standard.docx`
- 模板库支持管理员下载 Word 模板、上传新版草稿、预览 PDF、启用新版和回滚历史版本。

## 测试与验证

- P1 多人员、多股东、长地址压力测试。
- P1 股本特殊情况测试：`outputs/P1_v31_share_capital_partly_paid_test.xlsx`
- P1 股本特殊情况 PDF 包：`app/generated/P1_V31_SHARE_CAPITAL_PARTLY_PAID_TEST_P1_pdf_package.zip`
- M01 v7 一页式压力测试：`outputs/M01_v7_onepage_stress_delivery_package.zip`

## 下一阶段建议

1. 整理年审正式模板并接入快速年审页。
2. 再考虑客户资料库、联系人 ID、在线填写和电子签名。
