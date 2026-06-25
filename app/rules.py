from __future__ import annotations

from typing import Any

from config import P1_VERSION
from doc_generator import (
    build_context,
    default_first_fye_date,
    fye_month_text,
    format_money_number,
    format_number,
    short_date_text,
    to_number,
)
from excel_parser import is_yes


def text(value: Any) -> str:
    return str(value or "").strip()


def role_has(person: dict[str, Any], key: str) -> bool:
    return is_yes(person.get(key, ""))


def suggest_files(parsed: dict[str, Any], db_rules: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    task_type = parsed.get("task_type", "unknown")
    if task_type == "incorporation":
        return suggest_incorporation(parsed)
    if task_type == "maintenance":
        return suggest_maintenance(parsed)
    if task_type == "change":
        return suggest_change(parsed)
    return {
        "summary": {
            "task_type": task_type,
            "blocking_errors": ["无法识别模板类型，请使用注册或变更 v2 模板。"],
            "warnings": [],
            "info": [],
        },
        "detected_changes": [],
        "preview": [],
        "files": [],
    }


def suggest_incorporation(parsed: dict[str, Any]) -> dict[str, Any]:
    company = parsed.get("company", {})
    people = parsed.get("people", [])
    shareholders_raw = parsed.get("shareholders", [])
    directors = [p for p in people if role_has(p, "is_director")]
    secretaries = [p for p in people if role_has(p, "is_secretary")]
    nominee_directors = [p for p in people if role_has(p, "is_nominee_director")]
    client_directors = [p for p in directors if not role_has(p, "is_nominee_director")]

    blocking_errors: list[str] = []
    warnings: list[str] = []
    info: list[str] = []

    if not text(company.get("company_name")):
        blocking_errors.append("缺少公司名称。")
    if not directors:
        blocking_errors.append("没有识别到董事，无法生成董事同意书和董事决议。")
    if not secretaries:
        blocking_errors.append("没有识别到公司秘书，当前 P1 包无法完整生成。")
    if not shareholders_raw:
        blocking_errors.append("没有识别到股东，无法生成股权证书、Form 24 和 RORC。")
    if directors and not client_directors:
        warnings.append("没有识别到客户方董事。股权证书 Director 栏会保持空白，请人工确认实际签署董事。")

    person_by_key = {
        (text(p.get("full_name")).lower(), text(p.get("id_number")).lower()): p
        for p in people
        if text(p.get("full_name")) or text(p.get("id_number"))
    }
    person_key_set = set(person_by_key)
    for idx, shareholder in enumerate(shareholders_raw, start=1):
        shareholder_type = text(shareholder.get("shareholder_type")).lower()
        shares = to_number(shareholder.get("shares"))
        if shares <= 0:
            blocking_errors.append(f"Shareholders 第 {idx} 行股份数为空或小于等于 0。")
        if shareholder_type != "corporate":
            key = (text(shareholder.get("person_full_name")).lower(), text(shareholder.get("person_id_number")).lower())
            if key not in person_key_set:
                blocking_errors.append(f"自然人股东第 {idx} 行没有匹配到 People 里的姓名 + 证件号。")
            elif not role_has(person_by_key[key], "is_shareholder"):
                warnings.append(f"自然人股东第 {idx} 行在股东表存在，但人员信息里不是股东；系统仍会按股东表生成股权文件，请确认是否故意填 No。")
        else:
            if not text(shareholder.get("corporate_name")) or not text(shareholder.get("corporate_registration_number")):
                blocking_errors.append(f"公司股东第 {idx} 行缺少公司名称或注册号。")
            if not text(shareholder.get("authorized_rep_full_name")):
                warnings.append(f"公司股东第 {idx} 行没有授权代表，客户方签字可能需要人工指定。")
        if shares > 0 and to_number(shareholder.get("paid_amount")) <= 0:
            warnings.append(f"Shareholders 第 {idx} 行 paid_amount 为空或为 0，请确认是否符合实际实缴安排。")

    total_shares_from_rows = sum(to_number(row.get("shares")) for row in shareholders_raw)
    total_paid_from_rows = sum(to_number(row.get("paid_amount")) for row in shareholders_raw)
    company_total_shares = to_number(company.get("total_issued_shares"))
    company_paid = to_number(company.get("paid_up_capital"))
    if company_total_shares and total_shares_from_rows and company_total_shares != total_shares_from_rows:
        warnings.append(
            f"Company.total_issued_shares={format_number(company_total_shares)} 与股东行合计 {format_number(total_shares_from_rows)} 不一致；系统会优先使用股东行合计。"
        )
    if company_paid and total_paid_from_rows and company_paid != total_paid_from_rows:
        warnings.append(
            f"Company.paid_up_capital={format_money_number(company_paid)} 与股东行 paid_amount 合计 {format_money_number(total_paid_from_rows)} 不一致；系统会优先使用股东行合计。"
        )

    if text(company.get("incorporation_date")):
        auto_first_fye = default_first_fye_date(company.get("incorporation_date"))
        info.append(f"FYE 留空时，首个财年结束日将按注册日期自动计算为 {short_date_text(auto_first_fye)}，财年月为 {fye_month_text(auto_first_fye)}。")
    else:
        warnings.append("建议填写 incorporation_date；否则网页端会使用今天日期推算 FYE 和签字日期。")

    try:
        context = build_context(parsed)
    except Exception as exc:
        context = {}
        blocking_errors.append(f"生成上下文失败：{exc}")

    normalized_shareholders = context.get("shareholders", []) if context else []
    controllers = context.get("registrable_controllers", []) if context else []
    if normalized_shareholders:
        controller_names = ", ".join(row.get("shareholder_name", "") for row in controllers)
        info.append(f"RORC 将按 25% 以上持股初步判断控制人：{controller_names or '无'}。")

    files = [
        file_item("First Directors Resolution", "P1 注册董事决议", "所有董事同签一份", "Yes"),
        file_item("Consent to Act as Director / Form 45", f"识别到 {len(directors)} 名董事", "每名董事一份", "Yes" if directors else "No"),
        file_item("Secretary Consent / Form 45B", f"识别到 {len(secretaries)} 名秘书", "每名秘书一份", "Yes" if secretaries else "No"),
        file_item("Share Certificate", f"识别到 {len(shareholders_raw)} 名股东", "每名股东一份；Director 栏不预填姓名，秘书签发", "Yes" if shareholders_raw else "No"),
        file_item("Secretary Agreement", "秘书服务协议", "服务方 + 客户方代表；客户方默认股东 1", "Yes", manual_review=True),
        file_item("Nominee Director Agreement", f"识别到 {len(nominee_directors)} 名挂名董事", "挂名董事 + 客户方代表", "Yes" if nominee_directors else "No", manual_review=True),
        file_item("Return of Allotment of Shares / Form 24", "注册配股资料", "当前为一家公司一份，列出全部股东签字块", "Yes" if shareholders_raw else "No", manual_review=True),
        file_item("RORC Notice to Registrable Controller", "按 25% 以上持股初步筛选", "每名控制人一份", "Auto", manual_review=True),
        file_item("Register of Members", "注册后的初始股东名册", "不签；内部存档/交接", "Yes", manual_review=True),
        file_item("Signature Record Attachment", "P2：电子签名审计附件", "暂不生成", "No"),
        file_item("BizFile Filing Checklist", "注册资料检查", "不签", "Yes"),
    ]

    return {
        "summary": {
            "task_type": "incorporation",
            "p1_version": P1_VERSION,
            "company_name": text(company.get("company_name")),
            "directors": len(directors),
            "secretaries": len(secretaries),
            "shareholders": len(shareholders_raw),
            "registrable_controllers": len(controllers),
            "blocking_errors": blocking_errors,
            "warnings": warnings,
            "info": info,
        },
        "detected_changes": ["incorporation"],
        "preview": build_incorporation_preview(context) if context else [],
        "files": files,
    }


def build_incorporation_preview(context: dict[str, Any]) -> list[dict[str, Any]]:
    directors = context.get("directors", [])
    secretaries = context.get("secretaries", [])
    shareholders = context.get("shareholders", [])
    nominees = context.get("nominee_directors", []) or context.get("local_directors", [])
    controllers = context.get("registrable_controllers", [])
    client_signatory = context.get("client_signatory", {})
    secretary = context.get("secretary", {})
    return [
        preview_item("First Directors Resolution", 1, "所有董事", names(directors)),
        preview_item("Director Consent / Form 45", len(directors), "每名董事", names(directors)),
        preview_item("Secretary Consent / Form 45B", len(secretaries), "每名秘书", names(secretaries)),
        preview_item("Share Certificate", len(shareholders), "公司签发：Director 空白签字栏 + Company Secretary", f"秘书：{secretary.get('full_name', '')}；股东：{names(shareholders, 'shareholder_name')}"),
        preview_item("Secretary Service Agreement", 1, "服务方 + 客户方代表", client_signatory.get("full_name", "")),
        preview_item("Nominee Director Agreement", len(nominees), "挂名董事 + 客户方代表", names(nominees)),
        preview_item("Return of Allotment / Form 24", 1, "一家公司一份；列全部股东", names(shareholders, "shareholder_name")),
        preview_item("RORC Notice", len(controllers), "每名控制人一份", names(controllers, "shareholder_name")),
        preview_item("Register of Members", 1, "初始股东名册；不签", names(shareholders, "shareholder_name")),
    ]


def suggest_change(parsed: dict[str, Any]) -> dict[str, Any]:
    changes: list[str] = []
    warnings: list[str] = []
    info: list[str] = ["当前第一阶段只做变更判断预览，不生成变更文件。"]
    dr_groups: dict[str, list[str]] = {}

    for row in parsed.get("changes", []):
        required = is_yes(row.get("change_required")) or bool(text(row.get("new_value")))
        if required:
            change_item = text(row.get("change_item"))
            changes.append(change_item)
            if is_yes(row.get("combine_in_same_dr")):
                group = text(row.get("resolution_group")) or "DR-001"
                dr_groups.setdefault(group, []).append(change_item)
            if is_yes(row.get("manual_review_required")):
                warnings.append(f"{change_item} 标记为需要人工复核。")

    people = parsed.get("people", [])
    if any(text(p.get("position_action")) == "appoint" and "Director" in text(p.get("new_roles")) for p in people):
        changes.append("director_appointment")
    if any(text(p.get("position_action")) in {"resign", "remove"} and "Director" in text(p.get("current_roles")) for p in people):
        changes.append("director_resignation")
        warnings.append("董事辞任/移除后请人工确认仍有至少一名本地居民董事。")
    if any(text(p.get("position_action")) in {"appoint", "resign", "remove", "update"} and ("Secretary" in text(p.get("current_roles")) or "Secretary" in text(p.get("new_roles"))) for p in people):
        changes.append("secretary_change")

    if any(is_yes(row.get("change_required")) for row in parsed.get("personal_changes", [])):
        changes.append("personal_particulars")
    for row in parsed.get("share_changes", []):
        if is_yes(row.get("change_required")):
            changes.append("share_transfer")
            basis = text(row.get("value_basis"))
            if basis == "stamp_duty_higher_of_price_or_nav":
                warnings.append("股权转让选择了印花税/交易价值口径，请按实际价格或 NAV/股份价值中较高者人工复核。")

    changes = dedupe([c for c in changes if c])
    files = []
    if any(c in changes for c in ["business_activity_1", "ssic_code_1", "registered_office_address", "fye", "director_appointment", "director_resignation", "secretary_change", "share_transfer"]):
        files.append(file_item("Directors Resolution", "可合并的董事批准事项", "所有需签董事同签一份", "Yes"))
    if "company_name" in changes:
        files.append(file_item("Members Resolution / Notice of Resolution", "公司名称变更", "股东同签一份", "Yes", manual_review=True))
    if "director_appointment" in changes:
        files.append(file_item("Consent to Act as Director", "新增董事", "每名新董事一份", "Yes"))
    if "director_resignation" in changes:
        files.append(file_item("Director Resignation / Removal Documents", "董事辞任/移除", "相关人员签署", "Yes", manual_review=True))
    if "share_transfer" in changes:
        files.append(file_item("Share Transfer Instrument", "股份转让/股权变化", "转出方 + 转入方", "Yes", manual_review=True))
        files.append(file_item("Updated Share Certificate", "股权变化后更新证书", "公司签发", "Auto", manual_review=True))
    if "personal_particulars" in changes:
        files.append(file_item("Personal Particulars Filing Checklist", "个人资料变更", "通常不签", "Yes"))
    if changes:
        files.append(file_item("BizFile Filing Checklist", "所有变更任务", "不签", "Yes"))
    else:
        warnings.append("没有识别到明确变更项目，请检查 change_required 或 new_value。")

    preview = [
        preview_item("变更文件预览", len(files), "暂不生成，仅判断", "；".join(changes) or "无"),
    ]
    for group, items in dr_groups.items():
        preview.append(preview_item(f"DR 合并组 {group}", 1, "同一份 Directors' Resolution", "、".join(items)))

    return {
        "summary": {
            "task_type": "change",
            "company_name": parsed.get("company", {}).get("company_name", ""),
            "detected_change_count": len(changes),
            "blocking_errors": [],
            "warnings": warnings,
            "info": info,
        },
        "detected_changes": changes,
        "preview": preview,
        "files": files,
    }


NO_VALUES = {"no", "n", "false", "0", "否", "不", "无需", "不用"}

DR_EVENT_LABELS = {
    "change_registered_office": "注册地址变更",
    "change_office_hours": "办公时间变更",
    "change_business_activity": "营业范围/SSIC 变更",
    "change_fye": "财政年结日变更",
    "update_officer_particulars": "人员身份/地址资料更新",
    "appoint_director": "委任董事",
    "resign_director": "董事辞任",
    "appoint_secretary": "委任秘书",
    "resign_secretary": "秘书辞任",
}

EGM_EVENT_LABELS = {
    "transfer_in": "转入",
    "transfer_in_cooperative": "转入",
    "transfer_in_non_cooperative": "转入",
    "remove_director": "移除董事",
    "change_company_name": "公司更名",
    "strike_off": "注销/Strike off",
}


def suggest_maintenance(parsed: dict[str, Any]) -> dict[str, Any]:
    company = parsed.get("company", {})
    people = parsed.get("people", [])
    events = [row for row in parsed.get("change_events", []) if active_change_event(row)]
    transfers = [row for row in parsed.get("share_transfers", []) if active_data_row(row, "generate", ["transferor_name", "transferee_name", "shares_transferred"])]
    allotments = [row for row in parsed.get("share_allotments", []) if active_data_row(row, "generate", ["allottee_name", "shares_allotted", "total_paid"])]
    annual = parsed.get("annual_review", {})
    annual_required = annual_review_required(annual)

    blocking_errors: list[str] = []
    warnings: list[str] = []
    info: list[str] = ["公司维护/变更/年审表已支持普通董事决议、转入文件、股份转让、增资配股和年审文件包生成。"]
    files: list[dict[str, Any]] = []
    preview: list[dict[str, Any]] = []
    detected: list[str] = []

    if not text(company.get("company_name")):
        blocking_errors.append("缺少公司名称。")
    if not text(company.get("uen")):
        warnings.append("建议填写 UEN；现有公司变更和年审通常需要用 UEN 核对。")

    dr_groups: dict[str, list[str]] = {}
    egm_items: list[str] = []
    resignation_letters: list[str] = []
    transfer_in_items: list[str] = []
    high_risk_items: list[str] = []

    for row in events:
        event_type = text(row.get("event_type"))
        label = text(row.get("event_name_cn")) or DR_EVENT_LABELS.get(event_type) or EGM_EVENT_LABELS.get(event_type) or event_type
        detected.append(event_type)
        if event_type in DR_EVENT_LABELS or is_yes(row.get("combine_in_dr")):
            group = text(row.get("document_group")) or "DR-001"
            dr_groups.setdefault(group, []).append(label)
        if event_type in EGM_EVENT_LABELS or text(row.get("approval_route")).upper() in {"EGM", "WR", "EGM+DR"}:
            egm_items.append(label)
        if event_type.startswith("transfer_in"):
            transfer_in_items.append(label)
        if event_type in {"resign_director", "resign_secretary"} and is_yes(row.get("resignation_letter")):
            resignation_letters.append(label)
        if is_yes(row.get("manual_review_required")) or event_type in {"remove_director", "change_company_name", "strike_off"}:
            high_risk_items.append(label)

    director_appointments = [row for row in events if text(row.get("event_type")) == "appoint_director"]
    secretary_appointments = [row for row in events if text(row.get("event_type")) == "appoint_secretary"]

    director_signer_hint = text(
        company.get("director_signer_names")
        or company.get("director_signer_name")
        or annual.get("director_signer_name")
        or annual.get("director_signer_names")
    )
    member_signer_hint = text(
        company.get("member_signer_names")
        or company.get("shareholder_signer_names")
        or company.get("client_signatory_name")
        or annual.get("shareholder_signer_name")
        or annual.get("shareholder_signer_names")
    )
    has_director_candidate = any(role_has(person, "is_director") for person in people)
    has_client_candidate = any(
        role_has(person, "is_shareholder") or role_has(person, "is_client_signatory")
        for person in people
    )
    if (dr_groups or transfers or allotments or annual_required) and not director_signer_hint and not has_director_candidate:
        warnings.append("缺少董事签字人。系统可以生成文件，但董事签字栏可能为空，请在表格填写 director_signer_names 或在人员资料中标记董事。")
    if (transfer_in_items or allotments or annual_required) and not member_signer_hint and not has_client_candidate:
        warnings.append("缺少股东/客户授权签字人。系统可以生成文件，但股东或客户授权签字栏可能为空，请填写 member_signer_names 或 client_signatory_name。")

    if dr_groups:
        for group, labels in dr_groups.items():
            files.append(file_item("Combined Directors' Resolution", f"{group}: {'、'.join(labels)}", "董事签署", "Yes", package="普通变更 DR 包", doc_type="DR"))
            preview.append(preview_item(f"普通变更 DR 包 {group}", 1, "同组事项合并为一份董事决议", "、".join(labels)))
        if director_appointments:
            files.append(file_item("Director Consent / Form 45", f"识别到 {len(director_appointments)} 名新董事", "每名新董事一份；随 M01 包生成", "Yes", package="普通变更 DR 包", doc_type="Form 45"))
            preview.append(preview_item("Director Consent / Form 45", len(director_appointments), "每名新董事一份；随 M01 包生成", event_target_summary(director_appointments)))
        if secretary_appointments:
            files.append(file_item("Secretary Consent / Form 45B", f"识别到 {len(secretary_appointments)} 名新秘书", "每名新秘书一份；随 M01 包生成", "Yes", package="普通变更 DR 包", doc_type="Form 45B"))
            preview.append(preview_item("Secretary Consent / Form 45B", len(secretary_appointments), "每名新秘书一份；随 M01 包生成", event_target_summary(secretary_appointments)))

    if transfer_in_items:
        files.extend(
            [
                file_item("Transfer-In Resolution Package", "Notice / Shorter Notice / 股东决议 / 董事执行决议合并", "股东/授权人 + 董事签署", "Yes", manual_review=True, package="转入包", doc_type="Resolution Pack"),
                file_item("Handover and Resignation Package", "交接/终止服务信；如选择辞任信则附在同一 PDF 后面", "客户授权人；离任人员如适用", "Yes", manual_review=True, package="转入包", doc_type="Letter Pack"),
            ]
        )
        preview.append(preview_item("转入包", 2, "两份 PDF：决议包 + 交接辞任包", "、".join(transfer_in_items)))

    if resignation_letters:
        if transfer_in_items:
            preview.append(preview_item("转入包可选辞职信", len(resignation_letters), "附在第二份交接辞任包 PDF 后面", "、".join(resignation_letters)))
        else:
            files.append(file_item("Resignation Letter", "表格选择了 resignation_letter=Yes", "离任董事/秘书签署", "Yes", manual_review=True, package="董事/秘书任免包", doc_type="Letter"))
            preview.append(preview_item("董事/秘书任免包可选辞职信", len(resignation_letters), "按离任人员分别生成", "、".join(resignation_letters)))

    if egm_items and not transfer_in_items:
        files.append(file_item("EGM / Members' Written Resolution Pack", "股东层面事项： " + "、".join(egm_items), "股东签署", "Yes", manual_review=True, package="股东决议包", doc_type="EGM/WR"))
        preview.append(preview_item("股东决议包", 1, "EGM 或股东书面决议", "、".join(egm_items)))

    if transfers:
        detected.append("share_transfer")
        files.extend(
            [
                file_item("Share Transfer Directors' Resolution", f"识别到 {len(transfers)} 条股份转让", "董事签署", "Yes", package="股份转让包", doc_type="DR"),
                file_item("Instrument of Transfer", "股份转让核心交易文件", "转让人 + 受让人签署", "Yes", manual_review=True, package="股份转让包", doc_type="Instrument"),
                file_item("Updated Share Certificate", "转股后更新股权证书", "公司签发", "Auto", manual_review=True, package="股份转让包", doc_type="Certificate"),
                file_item("Register of Members Update Record", "股份转让后的股东名册更新记录", "不签；内部存档/交接", "Yes", manual_review=True, package="股份转让包", doc_type="Register"),
                file_item("Stamp Duty Review Checklist", "涉及转股印花税/交易价值复核", "内部核对", "Auto", manual_review=True, package="股份转让包", doc_type="Checklist"),
            ]
        )
        preview.append(preview_item("股份转让包", len(transfers), "不生成 Form 24；含 Register 更新记录", transfer_summary(transfers)))
        info.append("股份转让不生成 Form 24；Form 24 只用于增资/配股。")
        if any(is_yes(row.get("stamp_duty_review")) or text(row.get("consideration_basis")) == "stamp_duty_higher_of_price_or_nav" for row in transfers):
            warnings.append("股份转让涉及印花税或公司资产价值口径，请人工复核对价/NAV。")

    if allotments:
        detected.append("share_allotment")
        files.extend(
            [
                file_item("S161 Authority / EGM or Members' Written Resolution", "股东授权董事发行股份", "股东签署", "Yes", manual_review=True, package="增资配股包", doc_type="EGM/WR"),
                file_item("Allotment Directors' Resolution", f"识别到 {len(allotments)} 条配股", "董事签署", "Yes", package="增资配股包", doc_type="DR"),
                file_item("Share Application Letter", "认购人申请认购新股", "认购人签署", "Yes", package="增资配股包", doc_type="Application"),
                file_item("Return of Allotment / Form 24", "增资/配股申报资料", "董事/秘书申报", "Yes", manual_review=True, package="增资配股包", doc_type="Form 24"),
                file_item("Share Certificate", "新股发行后签发证书", "公司签发", "Auto", manual_review=True, package="增资配股包", doc_type="Certificate"),
                file_item("Register of Members Update Record", "增资配股后的股东名册更新记录", "不签；内部存档/交接", "Yes", manual_review=True, package="增资配股包", doc_type="Register"),
            ]
        )
        preview.append(preview_item("增资配股包", len(allotments), "S161/股东授权 + 董事配股 + Form 24 + Register 更新", allotment_summary(allotments)))

    if annual_required:
        detected.append("annual_review")
        fye = text(annual.get("fye_date"))
        if not fye:
            warnings.append("年审已启用，但缺少 fye_date。")
        status_values = " ".join(
            text(annual.get(key)).lower()
            for key in [
                "accounts_status",
                "company_activity_status",
                "financial_statements_type",
                "audit_exemption_status",
                "agm_status",
                "agm_route",
            ]
        )
        is_dormant_route = "dormant" in status_values
        is_audited_route = any(value in status_values.split() for value in ["audited", "audit_required"])
        agm_route_values = " ".join(text(annual.get(key)).lower() for key in ["agm_status", "agm_route"])
        no_agm_route = any(
            token in status_values
            for token in ["dormant_company", "exempt", "dispensed", "written_resolution", "written_resolutions", "no_agm"]
        ) and any(token in agm_route_values for token in ["dormant_company", "exempt", "dispensed", "written_resolution", "written_resolutions", "no_agm"])
        files.extend(
            [
                file_item("Annual Review Document List", "年审文件清单", "不签或随包", "Yes", package="年审包", doc_type="Cover"),
                file_item("Annual Review Directors' Resolution", "批准年审、财报/休眠状态和 AR 授权", "董事签署", "Yes", package="年审包", doc_type="DR"),
            ]
        )
        if no_agm_route:
            files.append(file_item("Members' Written Resolution / Consent", "AGM 豁免、书面决议或休眠年审路线", "成员/股东签署", "Yes", manual_review=True, package="年审包", doc_type="Consent"))
        else:
            files.extend(
                [
                    file_item("Notice of AGM", "通知股东 AGM 事项", "董事/秘书发出", "Yes", package="年审包", doc_type="Notice"),
                    file_item("Shorter Notice Consent", "短通知同意", "股东签署", "Auto", manual_review=True, package="年审包", doc_type="Consent"),
                    file_item("Proxy Form", "AGM proxy appointing instrument", "成员/股东签署", "Auto", manual_review=True, package="年审包", doc_type="Proxy"),
                    file_item("Attendance Sheet and AGM Minutes", "AGM 出席和会议记录", "主席/股东", "Yes", package="年审包", doc_type="Minutes"),
                ]
            )
        statement_label = "Dormant / no financial statements statement" if is_dormant_route else "Audited accounts review statement" if is_audited_route else "Section 205C small company audit exemption statement"
        files.extend(
            [
                file_item("Annual Return Authorization", "授权 BizFile 年报申报", "董事/客户授权人", "Yes", package="年审包", doc_type="Authorization"),
                file_item(statement_label, "财报方式/休眠/审计状态声明", "董事签署", "Yes", manual_review=True, package="年审包", doc_type="Statement"),
                file_item("Management Representation Letter", "管理层声明或休眠公司声明", "董事签署", text(annual.get("management_rep_letter")) or "Auto", manual_review=True, package="年审包", doc_type="MRL"),
            ]
        )
        preview.append(preview_item("年审包", 1, "合并签署 PDF + 独立内部复核清单", f"FYE: {fye or '-'}；AGM: {text(annual.get('agm_date')) or '-'}；状态: {text(annual.get('accounts_status')) or 'default'} / {text(annual.get('agm_status')) or text(annual.get('agm_route')) or 'ordinary_agm'}"))

    if high_risk_items:
        warnings.append("以下事项建议人工复核后再生成正式文件：" + "、".join(dedupe(high_risk_items)))
    if not (events or transfers or allotments or annual_required):
        warnings.append("没有识别到要生成的年审或变更事项，请检查 ChangeEvents.generate、ShareTransfers.generate、ShareAllotments.generate 或 AnnualReview.annual_review_required。")

    files.append(file_item("Internal Filing Checklist", "汇总本次业务单所有事项", "内部核对", "Yes", package="内部核对", doc_type="Checklist"))

    detected = dedupe([item for item in detected if item])
    package_names = dedupe([text(item.get("package")) for item in files if text(item.get("package"))])
    return {
        "summary": {
            "task_type": "maintenance",
            "company_name": text(company.get("company_name")),
            "detected_change_count": len(detected),
            "package_count": len(package_names),
            "ordinary_dr_groups": len(dr_groups),
            "m01_available": "Yes" if dr_groups else "No",
            "m02_available": "Yes" if transfer_in_items else "No",
            "m03_available": "Yes" if transfers else "No",
            "m04_available": "Yes" if allotments else "No",
            "m05_available": "Yes" if annual_required else "No",
            "transfer_in": "Yes" if transfer_in_items else "No",
            "share_transfers": len(transfers),
            "share_allotments": len(allotments),
            "annual_review": "Yes" if annual_required else "No",
            "manual_review_items": len(high_risk_items),
            "blocking_errors": blocking_errors,
            "warnings": warnings,
            "info": info,
        },
        "detected_changes": detected,
        "preview": preview,
        "files": files,
    }


def active_change_event(row: dict[str, Any]) -> bool:
    if not text(row.get("event_type")):
        return False
    return active_data_row(
        row,
        "generate",
        [
            "effective_date",
            "target_person_id",
            "target_name",
            "old_value",
            "new_value",
            "event_name_cn",
            "new_registered_office_address",
            "new_primary_activity",
            "new_primary_ssic",
            "new_secondary_activity",
            "new_secondary_ssic",
            "new_fye",
            "new_office_hours",
            "field_label",
        ],
    )


def active_data_row(row: dict[str, Any], flag_key: str, trigger_keys: list[str]) -> bool:
    flag = text(row.get(flag_key)).lower()
    if flag in NO_VALUES:
        return False
    if is_yes(row.get(flag_key)):
        return True
    return any(text(row.get(key)) for key in trigger_keys)


def annual_review_required(annual: dict[str, Any]) -> bool:
    if not annual:
        return False
    if text(annual.get("annual_review_required")).lower() in NO_VALUES:
        return False
    return is_yes(annual.get("annual_review_required")) or any(text(annual.get(key)) for key in ["fye_date", "agm_date"])


def transfer_summary(rows: list[dict[str, Any]]) -> str:
    parts = []
    for row in rows:
        parts.append(f"{text(row.get('transferor_name')) or text(row.get('transferor_shareholder_id'))} → {text(row.get('transferee_name')) or text(row.get('transferee_shareholder_id'))}: {text(row.get('shares_transferred')) or '-'} shares")
    return "；".join(parts) or "-"


def allotment_summary(rows: list[dict[str, Any]]) -> str:
    parts = []
    for row in rows:
        parts.append(f"{text(row.get('allottee_name')) or text(row.get('allottee_person_id'))}: {text(row.get('shares_allotted')) or '-'} shares")
    return "；".join(parts) or "-"


def event_target_summary(rows: list[dict[str, Any]]) -> str:
    parts = []
    for row in rows:
        parts.append(
            text(row.get("target_name"))
            or text(row.get("target_person_id"))
            or text(row.get("new_value"))
            or text(row.get("event_name_cn"))
        )
    return "、".join([part for part in parts if part]) or "-"


def file_item(
    name: str,
    reason: str,
    signing: str,
    suggested: str,
    manual_review: bool = False,
    package: str = "",
    doc_type: str = "",
) -> dict[str, Any]:
    return {
        "package": package,
        "doc_type": doc_type,
        "name": name,
        "reason": reason,
        "signing": signing,
        "suggested": suggested,
        "manual_review": manual_review,
    }


def preview_item(file_name: str, count: int, signing_logic: str, signer_detail: str) -> dict[str, Any]:
    return {
        "file": file_name,
        "count": count,
        "signing_logic": signing_logic,
        "signer_detail": signer_detail,
    }


def names(rows: list[dict[str, Any]], key: str = "full_name") -> str:
    return ", ".join(text(row.get(key)) for row in rows if text(row.get(key))) or "-"


def dedupe(items: list[str]) -> list[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            out.append(item)
            seen.add(item)
    return out
