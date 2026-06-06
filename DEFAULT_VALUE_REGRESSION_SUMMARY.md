# 默认值与普通用户流程回归总结

日期：2026-06-06

## 本轮目标

- 验证 P1 / P2 M01-M05 在“能留空就留空”的情况下是否仍能正确判断和生成 PDF。
- 优化普通员工视角的网站操作流程，避免页面像开发调试工具。
- 确保内部操作编号只在管理员视角保留，普通员工页面用业务名称。

## 发现并修复的问题

1. P2 V7 空白表格里部分生成开关默认是 `No`，会导致用户只填写新地址、业务范围、年审 FYE 时系统不自动触发文件包。
   - 已改为留空 / Auto 逻辑。
   - 普通 DR、股份转让、增资配股、年审均已覆盖。

2. P2 一页式表格的明细行在 `generate` 留空时曾被解析成 `No`。
   - 已改成空白默认为 `Auto`。
   - 只要对应核心字段有数据，系统会自动识别该事项。

3. 留空签字人时，系统可以生成文件，但签字栏可能为空。
   - 已加入复核提醒。
   - 董事签字人缺失、股东/客户授权签字人缺失都会在任务页提示。

4. 普通员工页面仍可能看到 M01/M02 等内部编号。
   - 已把首页、任务复核页、状态标签改成业务语言。
   - 管理员后台仍保留内部编号，便于开发和排查。

## 新增测试资料

留空默认值测试表位于：

- `tests/fixtures/defaults/P1_sparse_defaults.xlsx`
- `tests/fixtures/defaults/ordinary_dr_sparse_defaults.xlsx`
- `tests/fixtures/defaults/transfer_in_sparse_defaults.xlsx`
- `tests/fixtures/defaults/share_transfer_sparse_defaults.xlsx`
- `tests/fixtures/defaults/share_allotment_sparse_defaults.xlsx`
- `tests/fixtures/defaults/annual_review_sparse_defaults.xlsx`

这些表刻意保留大量默认字段为空，用来验证系统默认值和 Auto 判断。

## 新增测试工具

- `tools/create_default_regression_inputs.py`
  生成上述留空默认值测试表。

- `tools/run_default_regression.py`
  直接解析测试表并生成 P1 / M01-M05 PDF 包，检查 PDF 页数、占位符残留和关键生成判断。

- `tools/web_default_regression_smoke_test.py`
  用普通员工账号通过网页上传测试表，确认按钮、提醒、生成下载和员工视角页面都正常。

## 已通过测试

命令：

```powershell
$py='C:\Users\25896\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $py -m py_compile app\server.py app\rules.py app\excel_parser.py tools\create_default_regression_inputs.py tools\run_default_regression.py tools\web_default_regression_smoke_test.py
& $py tools\run_default_regression.py
& $py tools\web_default_regression_smoke_test.py
```

结果：

- P1 注册文件包：通过，生成 9 份 PDF。
- 普通董事决议：通过，留空签字人时有提醒。
- 转入文件包：通过，留空股东/客户授权签字人时有提醒。
- 股份转让包：通过，留空生成开关可自动识别。
- 增资配股包：通过，留空生成开关可自动识别。
- 年审文件包：通过，快速年审字段可触发生成。
- 员工网页视角：通过，看不到管理员调试数据，也看不到 M01-M05 内部编号。

最新机器报告写入：

- `outputs/default_value_regression_report.json`

`outputs` 目录被 `.gitignore` 排除，所以机器报告不提交；本文件作为可提交的人工摘要。
