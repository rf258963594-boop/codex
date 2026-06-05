# P2 M04 增资配股包总结

更新时间：2026-06-05  
状态：已接入 M04 v0.1 PDF 生成器，并按签署场景合并最终 PDF 包。

## 文件组合

模板仍然分开维护，最终下载包按签署逻辑合并：

- `01_M04_member_authority_package_{company}.pdf`：S161 / 股东授权包。
- `02_M04_directors_allotment_and_form24_package_{company}.pdf`：董事配股决议 + Return of Allotment / Form 24。
- `03_M04_share_application_{idx}_{allottee}.pdf`：每名认购人一份 Share Application。
- `04_M04_share_certificate_{idx}_{allottee}.pdf`：每名认购人一份 Share Certificate。
- `99_M04_internal_register_update_checklist_{company}.pdf`：内部复核清单，不作为客户签署文件。

## 触发逻辑

P2 v7 一页式业务单中“增资配股”区存在有效行，且 `generate` 不是 No。

核心字段：

- `allottee_name`
- `shares_allotted`
- `issued_share_capital`
- `amount_paid_per_share`
- `total_paid`
- `currency`
- `allotment_date`
- `authority_date`
- `certificate_no`

公司首页新增 / 使用：

- `total_issued_shares`
- `issued_share_capital`
- `paid_up_capital`
- `currency`

## 签字逻辑

- S161 / 股东授权：所有 `member_signer_names` 股东 / 成员签字。
- 配股董事决议：所有 `director_signer_names` 董事签字。
- Share Application：对应认购人签字。
- 股权证书：保留 Director 和 Company Secretary 签字栏，董事姓名默认留空。
- Form 24：董事 / 授权签字栏，最终 ACRA / BizFile 申报仍需人工复核。

## 证书规则

- M04 股权证书改为接近 P1 的横版全页证书正本。
- 不恢复旧版 counterfoil / receipt 页。
- 证书号留空时显示 `To be assigned (xx)`。
- 支持 fully paid 和 partly paid / unpaid capital 文案。

## 复核点

- 当前 issued share capital / paid-up share capital 不完整时，Form 24 资本汇总需要人工复核。
- 部分实缴 / unpaid capital 必须人工复核。
- 股权证书编号需要签署前复核。
- ACRA / BizFile lodgement 不是由本包自动完成。

## 压力测试

- 测试表：`outputs/M04_share_allotment_stress_input.xlsx`
- 本地生成包：`app/generated/M04-SHARE-ALLOTMENT-STRESS_P2_M04_pdf_package.zip`
- PDF 视觉检查目录：`outputs/M04_share_allotment_pdf_review_v4`
