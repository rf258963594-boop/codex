# P2 M03 股份转让包分析

更新时间：2026-06-05  
状态：已完成样本审计、字段逻辑规划，并已接入 M03 v0.1 生成器。

## 1. 已读取样本

样本目录：`C:\Users\25896\OneDrive\Desktop\变更文件模版\变更文件模版`

与 M03 最相关的样本：
- `转股TRANSFER OF SHARES_TRI-X KNOWLEDGE HUB 25052026.pdf`
- `转股EGM TRI-X KNOWLEDGE HUB 25052026.pdf`

审计结论：
- `转股TRANSFER OF SHARES...pdf` 实际是董事书面决议，内容为批准股份转让，并批准执行股权证书。
- `转股EGM...pdf` 文件名看起来像转股 EGM，但文本内容实际是董事辞任的 EGM Notice + Consent to Shorter Notice，没有真正的股份转让条款。这个文件不建议作为 M03 转股母版，只能作为 EGM Notice / Shorter Notice 的格式参考。

## 2. 官方规则参考

开发判断参考：
- ACRA Transfer of Shares 页面：转股准备文件包含 proper instrument of transfer；转股在 electronic Register of Members 更新后生效。参考：`https://www.acra.gov.sg/how-to-guides/shares-and-updating-share-information/transfer-of-shares`
- IRAS Buying or Acquiring Shares 页面：转让股份的印花税通常需要按实际价格或股份价值/NAV 等口径复核。参考：`https://www.iras.gov.sg/taxes/stamp-duty/for-shares/buying-or-acquiring-shares`

因此 M03 第一版应做到：
- 可以生成签字文件。
- 不把 Form 24 放进普通转股包。
- 印花税 / NAV / 交易对价不自动判断“合规完成”，只做复核提醒。

## 3. M03 v0.1 文件组合

### 第一份 PDF：M03 Share Transfer Directors' Resolution

建议包含：
- Directors' Resolutions in Writing - Approval of Transfer of Shares
- Optional ACRA / BizFile / Register update authorization

触发：
- `share_transfers` 至少一行 `generate = Yes`，或行内有转让人、受让人、转让股数。

签字：
- 所有 `director_signer_names` 董事签字。

说明：
- 这个文件可以复用 M01 已经有的“股份转让批准”逻辑，但 M03 应独立生成一份更聚焦的转股董事决议。
- 如果同一张表同时有注册地址、人员任免等普通变更，M01 仍负责普通 DR；M03 负责转股包。

### 第二份 PDF：M03 Instrument of Transfer

建议包含：
- Instrument of Transfer
- Transferor signature block
- Transferee signature block
- Optional witness / authorised representative block

触发：
- 每一条股份转让生成一份 Instrument of Transfer，或者多条同一转让人/受让人的交易可按后续规则合并。

签字：
- 转让人签字。
- 受让人签字。
- 如为公司股东/公司受让人，需要授权代表签字；第一版可先显示 `Authorised Signatory`。

说明：
- 这个是 M03 核心文件，不应仅用董事决议替代。
- 后续可加入公司股东的授权代表姓名、职务、公司注册号。

### 第三类 PDF：M03 Updated Share Certificate

建议包含：
- 新股东 / 受让人的 Share Certificate。
- 如果转让人仍有剩余股份，可选择给转让人生成更新后的 Share Certificate。
- 旧证书取消记录可放在内部 checklist 或证书脚注。

触发：
- 默认 Auto：有转股就生成新证书。

签字：
- 股权证书沿用 P1 当前逻辑：Director 栏不强行预填挂名董事姓名，保留董事签字栏；Company Secretary 签发。
- 如果要预填董事，优先客户董事，不默认挂名董事。

### 第四份 PDF：M03 Internal Checklist / Register Update Note

建议第一版先作为内部核对文件，不一定让客户签。

内容：
- Register of Members 更新
- Share certificate cancelled / issued
- Stamp duty review
- ACRA / BizFile / EROM 更新提醒
- 是否需要人工复核 NAV / consideration

签字：
- 内部核对，不作为客户签署文件。

## 4. 不建议第一版纳入 M03 的内容

- Form 24：普通股份转让不生成；Form 24 留给 M04 增资配股。
- EGM：普通转股通常先以董事批准 + Instrument of Transfer 为主；只有章程、股东协议、强制股东授权或特殊交易场景才触发 EGM / Members' Resolution。
- 自动计算印花税结论：第一版只提醒复核，不自动输出确定结论。
- 复杂 NAV / 财报估值：保留字段和复核提示，不作为自动判断核心。

## 5. M03 输入字段建议

### 一页式快速表新增 / 保留字段

当前 v7 已有股份转让区：
- `generate`
- `transferor_name`
- `transferee_name`
- `shares_transferred`
- `share_class`
- `transfer_date`
- `consideration_basis`
- `remarks`

建议 M03 第一版增加或预留：
- `transferor_id_number`
- `transferee_id_number`
- `transferor_address`
- `transferee_address`
- `consideration_amount`
- `currency`
- `old_certificate_no`
- `new_certificate_no`
- `transferor_remaining_shares`
- `generate_new_certificate`
- `stamp_duty_review`

其中真正必填可保持很少：
- 公司名称
- UEN
- 当前注册地址
- 文件日期 / 转让日期
- 董事签字人
- 转让人
- 受让人
- 转让股数
- 股份类别，留空默认 Ordinary

## 6. 系统判断逻辑

建议：
- `generate = No`：不生成。
- `generate = Yes`：生成。
- `generate` 留空但转让人、受让人、转让股数任意字段有值：系统按 Auto 识别为 Yes。
- `share_class` 留空：默认 Ordinary。
- `transfer_date` 留空：默认 `default_document_date`。
- `currency` 留空：默认公司币种；如果公司币种也空，默认 SGD。
- `generate_new_certificate` 留空：默认 Auto，即有转股就生成新证书。
- `stamp_duty_review` 留空：默认 Auto，只在预览/Checklist 提醒。

## 7. M03 与 M01 / M02 的关系

- M01：普通董事决议，可包含“批准股份转让”条款。
- M02：转入秘书公司包，处理股东/董事授权交接和新旧秘书董事任免逻辑。
- M03：股份转让专属包，处理 Instrument of Transfer、转股决议、股权证书和登记册更新。

如果同一业务单同时出现转入 + 股份转让：
- M02 生成转入相关两份 PDF。
- M03 生成转股相关 PDF。
- M01 只在还有注册地址、人员任免、营业范围等普通变更时生成；如果 M03 已生成独立转股董事决议，可避免 M01 重复生成同一条转股批准，后续接入时需要做去重。

## 8. 风险和人工复核点

必须提示人工复核：
- 转让人为公司股东或受让人为公司股东时，授权代表和公司内部批准是否齐全。
- 对价不是内部象征性转让，或涉及旧公司资产价值时，印花税 / NAV 口径要复核。
- 转让后是否仍满足公司章程、股东协议或投资人限制。
- 转让后股权证书编号、旧证书取消、新证书签发是否连续。
- 转让后 RORC / 控制人是否变化；如果 25% 以上控制人变化，后续应触发 RORC 更新文件。

## 9. 已完成接入项

- P2 v7 一页式业务单的股份转让区已加入 M03 所需字段。
- 已新建 M03 Word 模板。
- 已在 `doc_generator.py` 增加 `generate_p2_m03_pdf_package`。
- 已在网站任务页增加 M03 生成按钮，只在 `summary.m03_available = Yes` 时显示。
- 已做 M03 压力测试：公司股东、自然人股东、长地址、两笔转股、多董事签字、证书号留空、印花税/NAV 复核提醒。

## 10. 后续建议

1. 接入年审包：先做最常用 AGM / AR 签字文件，复杂年审场景保留人工复核。
2. 后续数据库化时，把转股和配股统一沉淀为股份交易记录，保留生成文件时的数据快照，避免以后公司资料更新后反向影响旧文件。
