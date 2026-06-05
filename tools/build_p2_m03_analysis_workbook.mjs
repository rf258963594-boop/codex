import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = "outputs";
const outPath = `${outDir}/P2_M03_股份转让字段逻辑表_v0.1.xlsx`;
const previewPath = `${outDir}/P2_M03_字段逻辑表预览.png`;

const workbook = Workbook.create();

function setTitle(sheet, title, subtitle) {
  sheet.showGridLines = false;
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#0F766E",
    font: { bold: true, color: "#FFFFFF", size: 16 },
  };
  sheet.getRange("A2:H2").merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    fill: "#ECFDF5",
    font: { color: "#064E3B", size: 10 },
  };
}

function writeTable(sheet, startCell, headers, rows, widths = []) {
  const startCol = startCell.match(/[A-Z]+/)[0];
  const startRow = Number(startCell.match(/\d+/)[0]);
  const colCount = headers.length;
  const rowCount = rows.length + 1;
  const endColCode = startCol.charCodeAt(0) + colCount - 1;
  const endCol = String.fromCharCode(endColCode);
  const range = `${startCell}:${endCol}${startRow + rowCount - 1}`;
  sheet.getRange(range).values = [headers, ...rows];
  const headerRange = sheet.getRange(`${startCell}:${endCol}${startRow}`);
  headerRange.format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
  const bodyRange = sheet.getRange(`${startCol}${startRow + 1}:${endCol}${startRow + rowCount - 1}`);
  bodyRange.format = {
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9E2EC" },
  };
  sheet.getRange(range).format.borders = { preset: "all", style: "thin", color: "#D9E2EC" };
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

const summary = workbook.worksheets.add("M03总览");
setTitle(summary, "P2 M03 股份转让包 - 总览", "用于确认 M03 第一版文件组合、触发条件、签字人和人工复核点。");
writeTable(
  summary,
  "A4",
  ["文件组", "建议文件名", "触发条件", "签字人", "是否第一版生成", "备注"],
  [
    ["M03-01", "Share Transfer Directors' Resolution", "股份转让区识别到有效转股", "所有董事签字", "Yes", "可复用 M01 转股批准逻辑，但 M03 建议独立生成。"],
    ["M03-02", "Instrument of Transfer", "每条有效转股", "转让人 + 受让人", "Yes", "M03 核心文件，不能只用董事决议替代。"],
    ["M03-03", "Updated Share Certificate", "有转股且 generate_new_certificate 不是 No", "董事签字栏 + Company Secretary", "Yes", "受让人新证书；转让人剩余股份证书可后续扩展。"],
    ["M03-04", "Internal Checklist / Register Update Note", "有转股", "内部核对", "Yes", "登记册、证书取消/签发、印花税复核提醒。"],
    ["M03-05", "EGM / Members' Resolution", "特殊章程/股东授权要求", "股东/成员", "No / Manual", "普通转股不默认生成，避免复杂化。"],
    ["M03-06", "Form 24", "增资配股", "N/A", "No", "普通转股不生成；保留给 M04 增资配股。"],
  ],
  [90, 210, 210, 150, 130, 280],
);

const fields = workbook.worksheets.add("输入字段");
setTitle(fields, "P2 M03 输入字段", "目标是最少字段可生成，复杂字段只作为可选复核。");
writeTable(
  fields,
  "A4",
  ["分类", "字段 key", "中文名称", "是否必填", "默认 / Auto", "用于哪些文件", "填写说明"],
  [
    ["公司", "company_name", "公司名称", "必填", "", "全部", "来自 P2 快速业务单。"],
    ["公司", "uen", "UEN", "建议必填", "", "全部", "现有公司文件建议填。"],
    ["公司", "registered_office_address", "当前注册地址", "建议必填", "", "Resolution / Instrument", "会议地址或公司背景信息。"],
    ["公司", "default_document_date", "文件日期", "建议必填", "留空可用当天", "全部", "DD/MM/YYYY。"],
    ["签字", "director_signer_names", "董事签字人", "建议必填", "留空则签字栏空白", "Resolution / Certificate", "多人用换行、逗号或分号。"],
    ["股份转让", "generate", "是否办理", "可空", "有核心字段则 Auto=Yes", "全部 M03", "No 表示不生成。"],
    ["股份转让", "transferor_name", "转让人", "必填", "", "Resolution / Instrument", "自然人或公司名称。"],
    ["股份转让", "transferee_name", "受让人", "必填", "", "Resolution / Instrument / Certificate", "自然人或公司名称。"],
    ["股份转让", "shares_transferred", "转让股数", "必填", "", "全部 M03", "数字。"],
    ["股份转让", "share_class", "股份类别", "可空", "Ordinary", "全部 M03", "默认 Ordinary。"],
    ["股份转让", "transfer_date", "转让日期", "可空", "default_document_date", "全部 M03", "DD/MM/YYYY。"],
    ["股份转让", "consideration_basis", "对价口径", "可空", "internal_paid_up_basis", "Instrument / Checklist", "内部转让可默认；涉及 NAV 需复核。"],
    ["股份转让", "consideration_amount", "转让对价", "可空", "", "Instrument / Checklist", "普通内部转让可空或 1。"],
    ["股份转让", "currency", "币种", "可空", "SGD 或公司币种", "Instrument / Certificate", "如 USD 需填写。"],
    ["股份转让", "old_certificate_no", "旧证书号", "可空", "", "Checklist / Certificate", "便于证书取消记录。"],
    ["股份转让", "new_certificate_no", "新证书号", "可空", "系统可自动建议", "Certificate", "第一版可人工填。"],
    ["股份转让", "transferor_remaining_shares", "转让人剩余股数", "可空", "", "Certificate / Checklist", "如需要给转让人更新证书时填。"],
    ["股份转让", "generate_new_certificate", "是否生成新证书", "可空", "Auto", "Certificate", "Auto / Yes / No。"],
    ["股份转让", "stamp_duty_review", "印花税复核", "可空", "Auto", "Checklist", "Auto 只提醒，不自动下结论。"],
    ["股份转让", "remarks", "备注", "可空", "", "内部", "不进入正式文件或只进 checklist。"],
  ],
  [92, 180, 150, 90, 160, 190, 310],
);

const logic = workbook.worksheets.add("系统判断");
setTitle(logic, "P2 M03 系统判断逻辑", "用于开发接入 generate_p2_m03_pdf_package 时直接参考。");
writeTable(
  logic,
  "A4",
  ["规则", "判断", "系统动作", "人工复核"],
  [
    ["触发", "generate=Yes 或 transferor/transferee/shares 有值", "显示 M03 生成按钮并加入预览", "无"],
    ["跳过", "generate=No", "不生成该行转股文件", "无"],
    ["股份类别", "share_class 留空", "使用 Ordinary", "低"],
    ["转让日期", "transfer_date 留空", "使用 default_document_date", "低"],
    ["股权证书", "generate_new_certificate 留空", "默认生成受让人新证书", "中：证书号需核对"],
    ["Form 24", "普通转股", "不生成 Form 24", "低"],
    ["公司股东/受让人", "名称看起来是公司或填写 company registration no.", "Instrument 显示 Authorised Signatory", "高：授权代表需核对"],
    ["对价/NAV", "consideration_basis=stamp_duty_higher_of_price_or_nav 或 stamp_duty_review=Yes", "Checklist 显示高风险复核", "高"],
    ["RORC", "转让后控制人可能变化", "Checklist 提醒后续 RORC 更新", "中"],
    ["M01 去重", "M03 已生成独立转股 DR", "M01 可避免重复输出转股批准条款", "中：接入时确认"],
  ],
  [160, 300, 300, 180],
);

const audit = workbook.worksheets.add("样本审计");
setTitle(audit, "已读取样本审计", "避免把文件名误导当成真实业务逻辑。");
writeTable(
  audit,
  "A4",
  ["样本文件", "实际内容", "可复用部分", "问题 / 注意"],
  [
    ["转股TRANSFER OF SHARES_TRI-X KNOWLEDGE HUB 25052026.pdf", "董事书面决议：批准 OOI KAI-LI 转让 10,000 股给 LIU KE，并批准执行股权证书。", "可作为 M03 Directors' Resolution 内容参考。", "缺少 Instrument of Transfer 的转让双方详细字段、证书编号、对价、印花税提醒。"],
    ["转股EGM TRI-X KNOWLEDGE HUB 25052026.pdf", "EGM Notice + Shorter Notice Consent，但条款是董事辞任，不是股份转让。", "只能参考 Notice / Shorter Notice 排版。", "不能作为 M03 转股母版。"],
  ],
  [300, 360, 260, 360],
);

for (const sheetName of ["M03总览", "输入字段", "系统判断", "样本审计"]) {
  const sheet = workbook.worksheets.getItem(sheetName);
  sheet.freezePanes.freezeRows(3);
}

await fs.mkdir(outDir, { recursive: true });

const sheetNames = ["M03总览", "输入字段", "系统判断", "样本审计"];
for (const sheetName of sheetNames) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const safeName = sheetName.replace(/[\\/:*?"<>|]/g, "_");
  const path = sheetName === "M03总览" ? previewPath : `${outDir}/P2_M03_${safeName}_预览.png`;
  await fs.writeFile(path, new Uint8Array(await preview.arrayBuffer()));
}

const scan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outPath);
console.log(JSON.stringify({ outPath, previewPath, formulaErrorScan: scan.ndjson }));
process.exit(0);
