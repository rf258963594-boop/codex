from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path


OUT = Path(r"C:\Users\25896\Documents\Codex\2026-05-25\new-chat\outputs")


def editable_rank(record: dict) -> tuple[int, bool, str]:
    return {".docx": 0, ".doc": 1, ".pdf": 2, ".xls": 3}.get(record["extension"], 9), "没有备注" in record["relative_path"], record["name"]


def text_rank(record: dict) -> tuple[int, int, bool, int]:
    has_text = 0 if record["text_length"] else 1
    ext = {".docx": 0, ".pdf": 1, ".doc": 2, ".xls": 3}.get(record["extension"], 9)
    return has_text, ext, "没有备注" in record["relative_path"], -record["text_length"]


def main() -> None:
    json_path = OUT / "签字文件模板审计数据.json"
    data = json.loads(json_path.read_text(encoding="utf-8"))
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for record in data["records"]:
        if record["role"] not in {"unknown", "signature_file_list"}:
            groups[(record["category"], record["role"])].append(record)

    labels = {
        "first_directors_resolution": "注册基础董事决议",
        "allotment_or_share_statement_form24": "股份/配股声明 legacy Form 24",
        "director_consent_form45": "董事同意/声明 legacy Form 45",
        "secretary_consent_form45b": "秘书同意/声明 legacy Form45B",
        "share_certificate": "股票证书",
        "rorc": "RORC 控制人登记",
        "secretary_agreement": "秘书服务协议",
        "nominee_director_agreement": "挂名董事协议",
        "client_acceptance_form": "客户接受/KYC表",
        "background_check": "背景调查/KYC",
        "agm_package": "AGM/年审包",
        "signature_record_attachment": "签名记录附件",
    }
    priorities = {
        "first_directors_resolution": "P1",
        "director_consent_form45": "P1",
        "secretary_consent_form45b": "P1",
        "share_certificate": "P1",
        "secretary_agreement": "P1",
        "nominee_director_agreement": "P1",
        "signature_record_attachment": "P1",
        "allotment_or_share_statement_form24": "P2",
        "rorc": "P2",
        "client_acceptance_form": "P2",
        "background_check": "P2",
        "agm_package": "P2",
    }
    remarks = {
        "first_directors_resolution": "一家公司一份，董事/挂名董事同签；可作为注册包主文件。",
        "director_consent_form45": "每个董事一份；旧名可保留显示但内部规则建议叫 director_consent。",
        "secretary_consent_form45b": "每个秘书一份或默认秘书一份；内部规则叫 secretary_consent。",
        "share_certificate": "按股东重复生成；要支持多个股东和公司股东。",
        "secretary_agreement": "服务协议，不是 ACRA 文件；签署方需要确认是客户个人、公司还是授权人。",
        "nominee_director_agreement": "依赖常用挂名董事库；按挂名董事/客户关系生成。",
        "allotment_or_share_statement_form24": "旧命名，是否继续出需要业务确认；不要作为强制法定逻辑写死。",
        "rorc": "需要在注册数据里增加 controller/UBO 字段，否则无法可靠生成。",
        "client_acceptance_form": "KYC 数据需求比注册主表多，建议作为补充文件包。",
        "background_check": "KYC/背景调查，字段独立，不宜混进核心注册流程。",
        "agm_package": "年审单独文件包；需要财报/AGM 日期/股东董事出席签字信息。",
        "signature_record_attachment": "通用附件，生成每个签字文件后追加。",
    }

    lines = [
        "# 签字文件模板逻辑矩阵",
        "",
        "说明：接入源文件优先选可编辑 `.docx/.doc`；内容判断来源可使用同名 PDF 或 docx 抽取文字。原目录未修改。",
        "",
        "| 优先级 | 文件逻辑 | 接入源文件 | 内容判断来源 | 字段/关键词 | 签字与生成逻辑 |",
        "|---|---|---|---|---|---|",
    ]
    console_rows = []
    for key in sorted(groups):
        role = key[1]
        group = groups[key]
        editable = sorted(group, key=editable_rank)[0]
        text_src = sorted(group, key=text_rank)[0]
        fields = ", ".join(text_src.get("field_hits", [])[:10]) or "未抽到文字，需人工看模板"
        lines.append(
            f"| {priorities.get(role, 'P3')} | {labels.get(role, role)} | `{editable['relative_path']}` | `{text_src['relative_path']}` | {fields} | {remarks.get(role, '需人工判断')} |"
        )
        console_rows.append((role, editable["extension"], text_src["extension"], text_src["text_length"]))

    lines.extend(
        [
            "",
            "## 我对逻辑的审核",
            "",
            "1. 第一版注册包可以接 P1：First Directors Resolution、Director Consent、Secretary Consent、Share Certificate、Secretary Agreement、Nominee Director Agreement、签名记录附件。",
            "2. P2 文件先不要挡住上线：Form 24、RORC、Client Acceptance、背景调查、AGM。它们需要更多字段或业务确认。",
            "3. `FORM 24 / 45 / 45B` 这些旧名称适合做“模板显示名”，不适合做系统内部规则名。内部规则应该用稳定名字：`director_consent`、`secretary_consent`、`share_certificate`。",
            "4. 现有注册资料 Excel 还缺 `controllers/UBO` 结构；如果接 RORC 和背景调查，必须加一个 KYC/Controller sheet，不能硬从 shareholder/director 猜。",
            "5. 年审 AGM 应该拆成独立入口或文件包，不要混在注册包里。",
            "6. “没有备注”PDF 是输出参考，不是第二套模板。",
        ]
    )

    md = OUT / "签字文件模板逻辑矩阵.md"
    md.write_text("\n".join(lines), encoding="utf-8")
    print(md)
    for row in console_rows:
        print(*row)


if __name__ == "__main__":
    main()

