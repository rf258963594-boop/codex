import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const ROOT = path.resolve(path.dirname(__filename), "..");
const OUT = path.join(ROOT, "outputs");

const COLORS = {
  title: "#0F766E",
  header: "#1F4D78",
  key: "#EEF2F7",
  input: "#FFF8D6",
  note: "#F8FAFC",
  border: "#D9E2EC",
  borderStrong: "#B7C6D8",
};

const companyFields = [
  ["公司名称", "Company name", "company_name", "", "Yes", "英文法定名称。"],
  ["业务单编号", "Business order ID", "business_order_id", "", "Optional", "可空；以后对接客户系统/电子签名时使用。"],
  ["数据来源", "Source type", "source_type", "Excel", "Optional", "Excel / BizFile / AI / Manual。"],
  ["来源文件编号", "Source file ID", "source_file_id", "", "Optional", "可填 BizFile/护照/旧表等来源文件编号或备注。"],
  ["注册地址", "Registered office address", "registered_office_address", "", "Yes", "英文地址；可用默认地址或客户地址。"],
  ["注册/文件日期", "Incorporation / document date", "incorporation_date", "", "Recommended", "DD/MM/YYYY；FYE 可由系统自动计算。"],
  ["首个财政年结日覆盖", "First FYE override", "first_fye", "", "Optional", "通常留空；特殊首个财年才填 DD/MM/YYYY。"],
  ["财政年度月份/年结日覆盖", "FYE override", "fye", "", "Optional", "通常留空；特殊 FYE 才填月份或 DD/MM/YYYY。"],
  ["主要业务描述", "Primary business activity", "business_activity_1", "", "Recommended", "建议填写英文描述。"],
  ["主要 SSIC Code", "Primary SSIC code", "ssic_code_1", "", "Recommended", "不知道可留空，复核时补。"],
  ["第二业务描述", "Secondary business activity", "business_activity_2", "", "Optional", ""],
  ["第二 SSIC Code", "Secondary SSIC code", "ssic_code_2", "", "Optional", ""],
  ["股本币种", "Share currency", "currency", "SGD", "Recommended", "留空默认 SGD。"],
  ["默认股份类别", "Default share class", "share_class_default", "Ordinary", "Recommended", "留空默认 Ordinary。"],
  ["已发行股本备用", "Issued share capital fallback", "issued_share_capital", "", "通常留空", "系统优先自动合计“股东与股份”页；这里只在特殊覆盖/核对时填写。"],
  ["已发行股数备用", "Issued shares fallback", "total_issued_shares", "", "通常留空", "系统优先自动合计“股东与股份”页；这里只在特殊覆盖/核对时填写。"],
  ["实缴股本备用", "Paid-up share capital fallback", "paid_up_capital", "", "通常留空", "系统优先自动合计“股东与股份”页；这里只在特殊覆盖/核对时填写。"],
  ["办公时间", "Office hours", "office_hours", "", "Optional", "留空默认 Monday to Friday, 9.00 a.m. to 5.00 p.m."],
  ["签字/文件日期", "Document date", "document_date", "", "System", "P1 全套文件统一使用注册日期；这里即使填写也会被注册日期覆盖。"],
  ["UEN/注册号", "UEN / Registration number", "uen", "", "Optional", "注册前通常留空。"],
  ["股东名册保存地址", "Register location", "register_location", "", "Optional", "留空默认注册地址。"],
  ["联系人", "Contact person_id", "contact_person_id", "", "Optional", "可空；代理注册/后续客户管理使用，引用人员信息里的 person_id。"],
  ["代理/代办人", "Agent person_id", "agent_person_id", "", "Optional", "可空；联系人不是董事/股东时使用。"],
  ["客户方签字人", "Client signer person_id", "client_signatory_person_id", "", "Optional", "填写人员编号时优先作为服务协议/挂名董事协议的客户方签字人；留空默认股东1。"],
  ["授权代表", "Authorised rep person_id", "authorized_representative_person_id", "", "Optional", "可空；企业股东或代理场景使用。"],
  ["任务类型", "Task type", "task_type", "incorporation", "System", "固定值，不要修改。"],
  ["内部备注", "Internal remarks", "remarks", "", "Optional", "不进入正式文件。"],
];

const peopleFields = [
  ["人员编号", "Person ID", "person_id", "建议 P001/P002；给联系人、代理人、签字人引用；不引用时可空"],
  ["来源", "Source", "source", "可空；空白/new=手填；common=调用后台常用人员"],
  ["常用人员名称", "Common person", "common_person_name", "source=common 时建议填写后台保存的名称；如留空，系统会尝试用英文姓名匹配，但精确填写更稳。"],
  ["英文姓名", "Full name", "full_name", "只填正式英文姓名；不要写 Director/Shareholder/客户董事等身份说明；调用常用人员时可空"],
  ["证件类型", "ID type", "id_type", "Passport / NRIC / FIN"],
  ["证件号", "ID number", "id_number", ""],
  ["国籍", "Nationality", "nationality", ""],
  ["出生日期", "Date of birth", "date_of_birth", "DD/MM/YYYY，可空"],
  ["住址", "Residential address", "residential_address", "长地址可换行"],
  ["邮箱", "Email", "email", ""],
  ["电话", "Phone", "phone", ""],
  ["是否董事", "Director", "is_director", "可空=Auto；客户董事建议填 Yes；常用挂名董事可留空由后台判断"],
  ["是否本地居民董事", "Local resident director", "is_local_resident_director", "通常留空=Auto；常用挂名董事由后台判断，只有手填本地董事/覆盖时选 Yes/No"],
  ["是否挂名董事", "Nominee director", "is_nominee_director", "通常留空=Auto；常用挂名董事由后台判断，客户董事不用点 No"],
  ["是否秘书", "Secretary", "is_secretary", "通常留空=Auto；常用秘书由后台判断，手填秘书才选 Yes"],
  ["是否股东", "Shareholder", "is_shareholder", "通常留空=Auto；系统按“股东与股份”页自动判断"],
  ["是否授权代表", "Authorised rep", "is_authorized_rep", "通常留空；企业股东授权代表请优先在“股东与股份”页填写名称。只有需要把此人作为通用客户授权签字人时才填 Yes。"],
  ["是否需要签字", "Signing required", "signing_required", "通常留空=Auto；系统按文件类型和人员身份判断；客户协议指定签字人请填公司信息页 client_signatory_person_id"],
  ["任命/文件日期", "Appointment/document date (system)", "appointment_date", "可空；P1 生成时统一使用注册日期，通常不用单独填写"],
  ["备注", "Remarks", "remarks", ""],
];

const shareholderFields = [
  ["股东类型", "Shareholder type", "shareholder_type", "person / corporate"],
  ["自然人股东姓名", "Person shareholder name", "person_full_name", "自然人股东填写，应和人员信息匹配"],
  ["自然人证件号", "Person ID number", "person_id_number", "用于匹配人员信息"],
  ["企业股东名称", "Corporate shareholder name", "corporate_name", "企业股东填写"],
  ["企业注册地", "Corporate registration country", "corporate_registration_country", ""],
  ["企业注册号", "Corporate registration no.", "corporate_registration_number", ""],
  ["企业注册地址", "Corporate registered address", "corporate_registered_address", ""],
  ["企业授权代表", "Authorised representative", "authorized_rep_full_name", "企业股东填写代表签字人姓名；人员页“是否授权代表”可留空。"],
  ["股份类别", "Share class", "share_class", "留空默认 Ordinary"],
  ["股数", "Shares", "shares", "必填"],
  ["已发行股本", "Issued share capital", "issued_share_capital", "通常留空=股数 1:1；如 100万股每股1，填 1,000,000"],
  ["实缴股本", "Paid-up share capital", "paid_amount", "通常留空=已发行股本；如 issued 1,000,000 但实缴 1，则填 1"],
  ["币种", "Currency", "currency", "留空默认 SGD"],
  ["证书编号覆盖", "Certificate no. override", "certificate_no", "通常留空，系统自动编号"],
  ["签署确认/系统保留", "Signing acknowledgement / system", "signing_required", "通常留空=Auto；系统按股东表生成股权证书、Form 24、RORC"],
  ["备注", "Remarks", "remarks", ""],
];

const outputFields = [
  ["文件包", "Package", "package", "INCORPORATION_BASIC", "系统使用"],
  ["输出格式", "Output format", "output_format", "docx_pdf_zip", "后续可统一只出 PDF"],
  ["经办人", "Prepared by", "prepared_by", "", "可空"],
  ["签字模式", "Signing mode", "signing_mode", "default", "默认即可"],
  ["客户方签字默认", "Client signer default", "client_signatory_logic", "shareholder_1", "默认股东1；如公司信息填写 client_signatory_person_id，则优先用指定人员"],
  ["Form 24 生成方式", "Form 24 mode", "form24_mode", "one_company_file_all_shareholders", "一家公司一份，列全部股东"],
  ["股权证书董事签字", "Certificate director signer", "share_certificate_signing_director", "director_1", "默认第一个客户董事；后续可网页覆盖"],
];

function colName(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function range(sheet, a1) {
  return sheet.getRange(a1);
}

function title(sheet, a1) {
  range(sheet, a1).format = {
    fill: COLORS.title,
    font: { bold: true, color: "#FFFFFF", size: 16 },
    wrapText: true,
    borders: { preset: "all", style: "thin", color: COLORS.borderStrong },
  };
}

function header(sheet, a1) {
  range(sheet, a1).format = {
    fill: COLORS.header,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
}

function keyStyle(sheet, a1) {
  range(sheet, a1).format = {
    fill: COLORS.key,
    font: { color: "#475569", size: 9 },
    wrapText: true,
  };
}

function inputStyle(sheet, a1) {
  range(sheet, a1).format = {
    fill: COLORS.input,
    wrapText: true,
    verticalAlignment: "top",
  };
}

function noteStyle(sheet, a1) {
  range(sheet, a1).format = {
    fill: COLORS.note,
    font: { color: "#334155" },
    wrapText: true,
  };
}

function grid(sheet, a1, color = COLORS.border) {
  range(sheet, a1).format.borders = { preset: "all", style: "thin", color };
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

function addList(sheet, a1, values) {
  sheet.dataValidations.add({ range: a1, rule: { type: "list", values } });
}

function buildReadme(wb) {
  const ws = wb.worksheets.add("填写说明");
  ws.showGridLines = false;
  range(ws, "A1:F1").merge();
  range(ws, "A1").values = [["新公司注册资料导入模板 v3.1 - 字段竖排版"]];
  title(ws, "A1:F1");
  range(ws, "A3:C20").values = [
    ["主题", "中文说明", "English notes"],
    ["核心逻辑", "字段放左边竖排；公司/输出只有一个填写列；人员和股东按列横向展开。", "Fields are vertical; people/shareholders are one column per object."],
    ["空白规则", "空白=Auto。除公司名称、注册地址、注册日期、人员基本资料、股东股份外，大部分选项可以不点。", "Blank means Auto. Most option cells do not need to be clicked."],
    ["身份默认值", "人员身份字段支持 Auto/Yes/No：空白或 Auto 表示系统自动判断；Yes 表示强制是；No 表示明确不是。", "Role fields support Auto/Yes/No. Blank or Auto lets the system decide."],
    ["默认判断", "常用人员按后台默认身份自动判断；自然人股东按股东与股份表自动判断 is_shareholder。客户董事建议明确填 Yes。", "Common people use saved roles; shareholders are inferred from the shareholder sheet. Client director should usually be marked Yes."],
    ["冲突规则", "如果后台/股东表认为应当是，但你填 No，系统会尊重 No，当作明确排除。", "If you enter No, the system treats it as an explicit override."],
    ["签字字段", "人员里的“是否需要签字”一般保持 Auto，由系统按文件类型和人员身份判断。秘书协议/挂名董事协议的客户方签字人默认股东1；如需指定，填写公司信息页 client_signatory_person_id。", "Keep signing_required as Auto in most cases. Statutory signing follows roles/shareholder rows. Client-side agreements default to shareholder 1 unless client_signatory_person_id is set."],
    ["股份核对", "公司信息页的总股数/已发行股本/实缴股本通常不用填，系统会优先合计“股东与股份”页；自动核对页只用于人工复核。", "Share count, issued share capital and paid-up capital are calculated from the shareholder sheet. The check sheet is for review."],
    ["为什么这样做", "它和“每人一行”的数据库表是同一个逻辑，只是转置后更适合人工检查。", "It is the transposed version of one-row-per-person data."],
    ["AI 使用", "可以把 BizFile、护照、旧注册表交给 AI，让 AI 按本模板填入。不要改 sheet 名称和 field_key。", "AI can fill this template. Do not rename sheets or field_key values."],
    ["黄色格", "黄色区域是填写区。灰色 field_key 给系统读取，也方便 AI 对齐字段。", "Yellow cells are inputs. Grey field_key cells are for the system."],
    ["常用人员", "秘书/挂名董事可填 source=common，再填 common_person_name，网站会从后台常用人员库补资料。", "Use source=common for saved nominee directors/secretaries."],
    ["默认值", "FYE、办公时间、证书编号、总股数、实缴金额等可由系统默认或计算。", "FYE, office hours, certificate numbers and totals can be defaulted/calculated."],
    ["日期格式", "表格建议用新加坡常用 DD/MM/YYYY；系统仍兼容 YYYY-MM-DD。正式文书会自动转成 1st JUNE 2026 这类日月年英文表述。", "Use DD/MM/YYYY in the sheet. The system also accepts YYYY-MM-DD and renders formal documents as day-month-year wording."],
    ["P1 日期规则", "P1 全套文件统一使用注册日期 incorporation_date。即使填写 document_date、signature_date 或 appointment_date，系统也会按注册日期生成。", "All P1 documents use incorporation_date. document_date, signature_date and appointment_date are ignored for P1."],
    ["复核重点", "上传后检查董事、秘书、股东股份合计、RORC 控制人、Form 24 和股权证书签字逻辑。", "Review directors, secretary, share totals, RORC, Form 24 and certificate signing."],
    ["兼容性", "网站仍支持旧 v2 横排模板；v3 是推荐主入口。", "The website still supports v2. v3 is the recommended entry."],
    ["注意", "不要删除 field_key 行列；不要合并填写区；正式英文姓名字段不要填中文名，也不要加 Director/Shareholder/客户董事等身份说明。", "Do not delete field_key rows/columns, merge input cells, or add role descriptions to legal names."],
  ];
  header(ws, "A3:C3");
  noteStyle(ws, "A4:C20");
  grid(ws, "A3:C20");
  setWidths(ws, [150, 560, 440, 20, 20, 20]);
}

function buildVerticalSingleSheet(wb, name, titleText, fields, sampleValues = {}) {
  const ws = wb.worksheets.add(name);
  ws.showGridLines = false;
  range(ws, "A1:F1").merge();
  range(ws, "A1").values = [[titleText]];
  title(ws, "A1:F1");
  range(ws, "A3:F3").values = [["项目", "English", "field_key", "填写内容 / Value", "必填", "说明"]];
  header(ws, "A3:F3");
  const data = fields.map(([cn, en, key, defaultValue, required, note]) => [
    cn,
    en,
    key,
    Object.hasOwn(sampleValues, key) ? sampleValues[key] : defaultValue,
    required,
    note,
  ]);
  range(ws, `A4:F${data.length + 3}`).values = data;
  keyStyle(ws, `C4:C${data.length + 3}`);
  inputStyle(ws, `D4:D${data.length + 3}`);
  noteStyle(ws, `F4:F${data.length + 3}`);
  grid(ws, `A3:F${data.length + 3}`);
  ws.freezePanes.freezeRows(3);
  setWidths(ws, [190, 220, 210, 340, 110, 520]);
}

function buildTransposedSheet(wb, name, titleText, fields, objectPrefix, objectCount, sampleObjects = []) {
  const ws = wb.worksheets.add(name);
  ws.showGridLines = false;
  const lastCol = colName(4 + objectCount);
  range(ws, `A1:${lastCol}1`).merge();
  range(ws, "A1").values = [[titleText]];
  title(ws, `A1:${lastCol}1`);
  const headerRow = ["项目", "English", "field_key", "说明"];
  for (let i = 1; i <= objectCount; i += 1) headerRow.push(`${objectPrefix}${i}`);
  range(ws, `A3:${lastCol}3`).values = [headerRow];
  header(ws, `A3:${lastCol}3`);
  const rows = fields.map(([cn, en, key, note]) => {
    const values = sampleObjects.map((obj) => obj[key] ?? "");
    while (values.length < objectCount) values.push("");
    return [cn, en, key, note, ...values];
  });
  range(ws, `A4:${lastCol}${rows.length + 3}`).values = rows;
  keyStyle(ws, `C4:C${rows.length + 3}`);
  noteStyle(ws, `D4:D${rows.length + 3}`);
  inputStyle(ws, `E4:${lastCol}${rows.length + 3}`);
  grid(ws, `A3:${lastCol}${rows.length + 3}`);
  ws.freezePanes.freezeRows(3);
  ws.freezePanes.freezeColumns(4);
  setWidths(ws, [180, 210, 200, 300, ...Array(objectCount).fill(190)]);
  return ws;
}

function sampleCompany() {
  return {
    company_name: "LONGFIELD TEST HOLDINGS PTE. LTD.",
    business_order_id: "INC-TEST-001",
    source_type: "Excel",
    registered_office_address: "111 NORTH BRIDGE ROAD, #29-06A PENINSULA PLAZA, SINGAPORE 179098",
    incorporation_date: "27/05/2026",
    business_activity_1: "Management consultancy services",
    ssic_code_1: "70201",
    business_activity_2: "Development of software and applications",
    ssic_code_2: "62011",
    contact_person_id: "P003",
    agent_person_id: "",
    client_signatory_person_id: "P004",
    authorized_representative_person_id: "",
    remarks: "示例数据：客户协议签字人指定 P004；正式填写时可留空默认股东1。",
  };
}

function samplePeople() {
  return [
    { person_id: "P001", source: "common", common_person_name: "挂名董事 A", is_director: "Auto", is_local_resident_director: "Auto", is_nominee_director: "Auto", is_secretary: "Auto", is_shareholder: "Auto", is_authorized_rep: "", signing_required: "Auto", appointment_date: "27/05/2026", remarks: "调用后台常用人员资料；身份可保持 Auto" },
    { person_id: "P002", source: "common", common_person_name: "公司秘书 A", is_director: "Auto", is_local_resident_director: "Auto", is_nominee_director: "Auto", is_secretary: "Auto", is_shareholder: "Auto", is_authorized_rep: "", signing_required: "Auto", appointment_date: "27/05/2026", remarks: "调用后台常用人员资料；身份可保持 Auto" },
    { person_id: "P003", source: "new", full_name: "ZHANG YI", id_type: "Passport", id_number: "P1000001A", nationality: "Chinese", date_of_birth: "03/03/1988", residential_address: "ROOM 1808, INTERNATIONAL COMMERCE CENTRE, BEIJING, CHINA", email: "zhang.yi@example.com", phone: "+86 13800000001", is_director: "Yes", is_local_resident_director: "Auto", is_nominee_director: "Auto", is_secretary: "Auto", is_shareholder: "Auto", is_authorized_rep: "", signing_required: "Auto", appointment_date: "27/05/2026", remarks: "客户董事兼股东；股东身份可由股东与股份表判断" },
    { person_id: "P004", source: "new", full_name: "ZHANG ER", id_type: "Passport", id_number: "P2000002B", nationality: "Chinese", date_of_birth: "04/04/1990", residential_address: "SHENZHEN BAY TECHNOLOGY ECOLOGY PARK, SHENZHEN, CHINA", email: "zhang.er@example.com", phone: "+86 13900000002", is_director: "Yes", is_local_resident_director: "Auto", is_nominee_director: "Auto", is_secretary: "Auto", is_shareholder: "Auto", is_authorized_rep: "", signing_required: "Auto", appointment_date: "27/05/2026", remarks: "客户董事兼股东；股东身份可由股东与股份表判断" },
  ];
}

function sampleShareholders() {
  return [
    { shareholder_type: "person", person_full_name: "ZHANG YI", person_id_number: "P1000001A", share_class: "Ordinary", shares: 700, issued_share_capital: 700, paid_amount: 700, currency: "SGD", signing_required: "Yes", remarks: "70% 股东，RORC 控制人" },
    { shareholder_type: "person", person_full_name: "ZHANG ER", person_id_number: "P2000002B", share_class: "Ordinary", shares: 300, issued_share_capital: 300, paid_amount: 300, currency: "SGD", signing_required: "Yes", remarks: "30% 股东，RORC 控制人" },
  ];
}

function buildMappingSheet(wb) {
  const ws = wb.worksheets.add("字段映射");
  ws.showGridLines = false;
  range(ws, "A1:E1").merge();
  range(ws, "A1").values = [["字段映射 Field Mapping"]];
  title(ws, "A1:E1");
  range(ws, "A3:E8").values = [
    ["表格", "转换成系统数据", "读取方式", "用户是否填写", "备注"],
    ["公司信息", "Company", "field_key + 填写内容", "Yes", "单一公司信息。"],
    ["人员信息", "People", "field_key + 人员列；每一列是一名人员", "Yes", "转置后等同于旧版每人一行。"],
    ["股东与股份", "Shareholders", "field_key + 股东列；每一列是一名股东", "Yes", "转置后等同于旧版每股东一行。"],
    ["输出设置", "Generation", "field_key + 填写内容", "Optional", "大多数字段可不改。"],
    ["旧版兼容", "Company / People / Shareholders / Generation", "旧横排标准表", "Yes", "仍然可上传。"],
  ];
  header(ws, "A3:E3");
  noteStyle(ws, "A4:E8");
  grid(ws, "A3:E8");
  setWidths(ws, [170, 250, 420, 150, 360]);
}

function buildCheckSheet(wb) {
  const ws = wb.worksheets.add("自动核对");
  ws.showGridLines = false;
  range(ws, "A1:D1").merge();
  range(ws, "A1").values = [["自动核对 Auto Checks"]];
  title(ws, "A1:D1");
  range(ws, "A3:D16").values = [
    ["核对项", "系统读取/计算", "状态", "说明"],
    ["股东页合计股数", "=SUM('股东与股份'!E13:N13)", "", "来自“股东与股份”页的 Shares。"],
    ["股东页合计已发行股本", "=SUM('股东与股份'!E14:N14)", "", "来自“股东与股份”页的 Issued share capital。留空时网站默认按股数 1:1。"],
    ["股东页合计实缴股本", "=SUM('股东与股份'!E15:N15)", "", "来自“股东与股份”页的 Paid-up share capital。留空时网站默认等于已发行股本。"],
    ["公司页备用股数", "='公司信息'!D19", "", "通常留空；留空时系统使用股东页合计。"],
    ["公司页备用已发行股本", "='公司信息'!D18", "", "通常留空；留空时系统使用股东页合计。"],
    ["公司页备用实缴股本", "='公司信息'!D20", "", "通常留空；留空时系统使用股东页合计。"],
    ["未缴股本提示", "=MAX(0,SUM('股东与股份'!E14:N14)-SUM('股东与股份'!E15:N15))", "=IF(B10>0,\"需要复核：存在未缴股本\",\"OK\")", "例如 issued 1,000,000 / paid-up 1，会出现未缴股本。"],
    ["股数一致性", "", "=IF('公司信息'!D19=\"\",\"OK - 公司页留空，使用股东页合计\",IF('公司信息'!D19=SUM('股东与股份'!E13:N13),\"OK\",\"请检查：公司页备用股数与股东页合计不一致\"))", "只有填写公司页备用股数时才需要一致。"],
    ["已发行股本一致性", "", "=IF('公司信息'!D18=\"\",\"OK - 公司页留空，使用股东页合计\",IF('公司信息'!D18=SUM('股东与股份'!E14:N14),\"OK\",\"请检查：公司页备用已发行股本与股东页合计不一致\"))", "只有填写公司页备用已发行股本时才需要一致。"],
    ["实缴一致性", "", "=IF('公司信息'!D20=\"\",\"OK - 公司页留空，使用股东页合计\",IF('公司信息'!D20=SUM('股东与股份'!E15:N15),\"OK\",\"请检查：公司页备用实缴股本与股东页合计不一致\"))", "只有填写公司页备用实缴股本时才需要一致。"],
    ["客户协议签字人", "='公司信息'!D27", "=IF('公司信息'!D27=\"\",\"默认股东1\",\"使用指定人员编号\")", "只影响秘书协议/挂名董事协议的客户方签字人。"],
    ["注册日期", "='公司信息'!D9", "=IF('公司信息'!D9=\"\",\"建议填写注册日期\",\"OK\")", "P1 全套文件日期统一用注册日期。"],
    ["备注", "", "", "本页只用于人工复核，不影响系统读取。"],
  ];
  header(ws, "A3:D3");
  noteStyle(ws, "A4:D16");
  grid(ws, "A3:D16");
  setWidths(ws, [190, 260, 380, 520]);
}

async function buildWorkbook(sample) {
  const wb = Workbook.create();
  buildReadme(wb);
  buildVerticalSingleSheet(wb, "公司信息", "公司信息 Company Information", companyFields, sample ? sampleCompany() : {});
  const peopleSheet = buildTransposedSheet(wb, "人员信息", "人员信息 People", peopleFields, "人员", 10, sample ? samplePeople() : []);
  const shareholderSheet = buildTransposedSheet(wb, "股东与股份", "股东与股份 Shareholders", shareholderFields, "股东", 10, sample ? sampleShareholders() : []);
  buildVerticalSingleSheet(wb, "输出设置", "输出设置 Output Options", outputFields, {});
  buildCheckSheet(wb);
  buildMappingSheet(wb);

  addList(peopleSheet, "E5:N5", ["new", "common"]);
  addList(peopleSheet, "E15:N21", ["Auto", "Yes", "No"]);
  addList(shareholderSheet, "E4:N4", ["person", "corporate"]);

  await wb.render({ sheetName: "公司信息", range: "A1:F24", scale: 1, format: "png" });
  await wb.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "formula error scan",
  });
  return wb;
}

async function saveWorkbook(wb, fileName) {
  await fs.mkdir(OUT, { recursive: true });
  const xlsx = await SpreadsheetFile.exportXlsx(wb);
  const fullPath = path.join(OUT, fileName);
  await xlsx.save(fullPath);
  console.log(fullPath);
}

await saveWorkbook(await buildWorkbook(false), "AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_空白.xlsx");
await saveWorkbook(await buildWorkbook(true), "AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_示例.xlsx");
await saveWorkbook(await buildWorkbook(false), "AI适配_新公司注册资料模板_v3_人性化竖排_Auto版_空白.xlsx");
await saveWorkbook(await buildWorkbook(true), "AI适配_新公司注册资料模板_v3_人性化竖排_Auto版_示例.xlsx");
process.exit(0);
