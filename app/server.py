from __future__ import annotations

import html
import json
import os
import re
import secrets
import shutil
import sys
import threading
import traceback
import zipfile
from datetime import UTC, datetime, timedelta, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from PIL import Image, UnidentifiedImageError

from config import (
    APP_NAME,
    DATA_DIR,
    DOC_TEMPLATE_DIR,
    GENERATED_DIR,
    IMPORT_TEMPLATE_DIR,
    MAX_UPLOAD_BYTES,
    OUTPUTS_DIR,
    P1_VERSION,
    SESSION_COOKIE,
    UPLOAD_DIR,
    UPLOAD_RETENTION_DAYS,
)
from db import connect, hash_password, init_db, log_action, now, verify_password
from doc_generator import (
    generate_p1_pdf_package,
    generate_p2_m01_pdf_package,
    generate_p2_m02_pdf_package,
    generate_p2_m03_pdf_package,
    generate_p2_m04_pdf_package,
    generate_p2_m05_pdf_package,
    safe_filename,
)
from doc_render import convert_docx_to_pdf
from excel_parser import parse_excel, resolve_people, to_json
from rules import suggest_files
from signatures import SIGNATURE_DIR, default_signature_text, ensure_signature_image, safe_slug, signature_relative_path


SIGNATURE_UPLOAD_MAX_BYTES = 2 * 1024 * 1024
DISPLAY_TIMEZONE = timezone(timedelta(hours=8), name="SGT")


TEMPLATE_DOWNLOADS = {
    "registration_human_blank": ("新公司注册 v3.1 人性化竖排 Auto 空白模板", OUTPUTS_DIR / "AI适配_新公司注册资料模板_v3.1_人性化竖排_Auto版_空白.xlsx"),
    "registration_human_sample": ("新公司注册 v3.1 人性化竖排 Auto 示例模板", OUTPUTS_DIR / "P1_v31_registration_sample_share_capital.xlsx"),
    "registration_blank": ("新公司注册 v2 横排空白模板", OUTPUTS_DIR / "AI适配_新公司注册资料模板_v2_空白.xlsx"),
    "registration_sample": ("新公司注册 v2 横排示例模板", OUTPUTS_DIR / "AI适配_新公司注册资料模板_v2_示例.xlsx"),
    "maintenance_blank": ("公司维护/变更年审 v7 一页式快速业务单含快速年审空白模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_空白.xlsx"),
    "maintenance_sample": ("公司维护/变更年审 v7 一页式快速业务单含快速年审示例模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v7_一页式快速业务单含快速年审_示例.xlsx"),
    "maintenance_v5_blank": ("公司维护/变更年审 v5 快速生成详情开关版空白模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v5_快速生成详情开关版_空白.xlsx"),
    "maintenance_v5_sample": ("公司维护/变更年审 v5 快速生成详情开关版示例模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v5_快速生成详情开关版_示例.xlsx"),
    "maintenance_v4_blank": ("公司维护/变更年审 v4 竖排 Auto 预留字段空白模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v4_竖排_Auto预留字段版_空白.xlsx"),
    "maintenance_v4_sample": ("公司维护/变更年审 v4 竖排 Auto 预留字段示例模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v4_竖排_Auto预留字段版_示例.xlsx"),
    "maintenance_legacy_blank": ("公司维护/变更年审 v3 横排空白模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v3_空白.xlsx"),
    "maintenance_legacy_sample": ("公司维护/变更年审 v3 横排示例模板", OUTPUTS_DIR / "AI适配_公司维护变更年审资料模板_v3_示例.xlsx"),
    "p2_m01_dr": ("P2 M01 普通董事决议母版", OUTPUTS_DIR / "P2_standard_templates_v1" / "M01_combined_directors_resolution_standard.docx"),
    "p2_m01_field_map": ("P2 M01 字段图谱", OUTPUTS_DIR / "P2_standard_templates_v1" / "M01_field_map.md"),
    "p2_m03_field_map": ("P2 M03 股份转让字段图谱", OUTPUTS_DIR / "P2_standard_templates_v1" / "M03_field_map.md"),
    "p2_m04_field_map": ("P2 M04 增资配股字段图谱", OUTPUTS_DIR / "P2_standard_templates_v1" / "M04_field_map.md"),
    "p2_m05_field_map": ("P2 M05 年审字段图谱", OUTPUTS_DIR / "P2_standard_templates_v1" / "M05_field_map.md"),
    "change_blank": ("现有公司变更 v2 空白模板", OUTPUTS_DIR / "AI适配_现有公司变更资料模板_v2_空白.xlsx"),
    "change_sample": ("现有公司变更 v2 示例模板", OUTPUTS_DIR / "AI适配_现有公司变更资料模板_v2_示例.xlsx"),
}

TEMPLATE_DOWNLOADS.update(
    {
        "registration_human_blank": (
            TEMPLATE_DOWNLOADS["registration_human_blank"][0],
            IMPORT_TEMPLATE_DIR / "P1_registration_blank_v3_1.xlsx",
        ),
        "registration_human_sample": (
            TEMPLATE_DOWNLOADS["registration_human_sample"][0],
            IMPORT_TEMPLATE_DIR / "P1_registration_sample_v3_1.xlsx",
        ),
        "maintenance_blank": (
            TEMPLATE_DOWNLOADS["maintenance_blank"][0],
            IMPORT_TEMPLATE_DIR / "P2_maintenance_annual_blank_v7.xlsx",
        ),
        "maintenance_sample": (
            TEMPLATE_DOWNLOADS["maintenance_sample"][0],
            IMPORT_TEMPLATE_DIR / "P2_maintenance_annual_sample_v7.xlsx",
        ),
    }
)

PUBLIC_TEMPLATE_KEYS = {
    "registration_human_blank",
    "registration_human_sample",
    "maintenance_blank",
    "maintenance_sample",
}

ACTIVE_IMPORT_TEMPLATE_KEYS = {
    "registration_human_blank",
    "registration_human_sample",
    "maintenance_blank",
    "maintenance_sample",
}

TEMPLATE_NOTES = {
    "registration_human_blank": ("新公司注册 v3 竖排填写表", "推荐：字段竖排，人员/股东横向展开"),
    "registration_human_sample": ("新公司注册 v3 示例", "给 AI 或人工参考填写方式"),
    "registration_blank": ("旧版注册 v2 横排表", "管理员保留，普通入口已隐藏"),
    "registration_sample": ("旧版注册 v2 示例", "管理员保留，普通入口已隐藏"),
    "maintenance_blank": ("维护/变更/年审 v7 一页式业务单", "推荐：最少字段先生成文件，另含快速年审页；普通董事决议签字人支持多个董事"),
    "maintenance_sample": ("维护/变更/年审 v7 示例", "示例含多个董事签字人、董事辞任、委任、可选转股和快速年审默认字段"),
    "maintenance_v5_blank": ("旧版 P2 v5 详情开关表", "管理员保留，适合需要更多结构化字段时使用"),
    "maintenance_v5_sample": ("旧版 P2 v5 示例", "管理员保留，普通入口已改用 v6"),
    "maintenance_v4_blank": ("旧版 P2 v4 竖排表", "管理员保留，适合完整资料库式导入"),
    "maintenance_v4_sample": ("旧版 P2 v4 示例", "管理员保留，普通入口已改用 v5"),
    "maintenance_legacy_blank": ("旧版 P2 v3 横排表", "管理员保留，普通入口已隐藏"),
    "maintenance_legacy_sample": ("旧版 P2 v3 横排示例", "管理员保留，普通入口已隐藏"),
    "p2_m01_dr": ("普通董事决议母版", "管理员维护；内部编号 M01"),
    "p2_m01_field_map": ("普通董事决议字段图谱", "开发和管理员排查字段用"),
    "p2_m03_field_map": ("股份转让字段图谱", "开发和管理员排查 M03 字段用"),
    "p2_m04_field_map": ("增资配股字段图谱", "开发和管理员排查 M04 字段用"),
    "p2_m05_field_map": ("年审字段图谱", "开发和管理员排查 M05 字段用"),
    "change_blank": ("旧变更导入表", "保留参考，建议改用 v3"),
    "change_sample": ("旧变更示例", "保留参考，建议改用 v3"),
}

DOCUMENT_TEMPLATES = {
    "p1_first_directors_resolution": {
        "category": "注册 P1",
        "name": "注册董事决议",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "01_first_directors_resolution_standard.docx",
        "note": "所有董事同签一份",
    },
    "p1_form45_director": {
        "category": "注册 P1",
        "name": "Form 45 董事同意书",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "02_director_consent_form45_standard.docx",
        "note": "每名董事一份",
    },
    "p1_form45b_secretary": {
        "category": "注册 P1",
        "name": "Form 45B 秘书同意书",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "03_secretary_consent_form45b_standard.docx",
        "note": "每名秘书一份",
    },
    "p1_share_certificate": {
        "category": "注册 P1",
        "name": "股权证书",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "04_share_certificate_standard.docx",
        "note": "每名股东一份",
    },
    "p1_secretary_agreement": {
        "category": "注册 P1",
        "name": "秘书服务协议",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "05_secretary_service_agreement_standard.docx",
        "note": "服务方 + 客户方代表",
    },
    "p1_nominee_director_agreement": {
        "category": "注册 P1",
        "name": "挂名董事协议",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "06_nominee_director_agreement_standard.docx",
        "note": "每名挂名董事一份",
    },
    "p1_form24": {
        "category": "注册 P1",
        "name": "Form 24 / Return of Allotment",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "07_return_of_allotment_form24_standard.docx",
        "note": "注册配股资料",
    },
    "p1_rorc": {
        "category": "注册 P1",
        "name": "RORC Notice",
        "version": "v1.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "08_rorc_notice_controller_standard.docx",
        "note": "每名控制人一份",
    },
    "p1_register_members": {
        "category": "注册 P1",
        "name": "Register of Members",
        "version": "v1.0",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p1_standard_v3_part1" / "09_register_of_members_standard.docx",
        "note": "注册后的初始股东名册",
    },
    "p2_m01_ordinary_dr": {
        "category": "变更 P2",
        "name": "普通董事决议",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M01_combined_directors_resolution_standard.docx",
        "note": "普通 DR，内部编号 M01",
    },
    "p2_m02_resolution_package": {
        "category": "变更 P2",
        "name": "M02 EGM Notice / Resolutions",
        "version": "v0.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M02_resolution_package_transfer_in_standard.docx",
        "note": "Notice / Shorter Notice / Members' Resolution / Directors' Resolution，内部编号 M02",
    },
    "p2_m02_handover_resignation_package": {
        "category": "变更 P2",
        "name": "M02 交接辞任包",
        "version": "v0.2",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M02_handover_and_resignation_package_standard.docx",
        "note": "交接/终止服务信，辞任信按表格规则自动附在同一 PDF 后面，内部编号 M02",
    },
    "p2_m03_resolution": {
        "category": "变更 P2",
        "name": "M03 转股董事决议",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M03_share_transfer_directors_resolution_standard.docx",
        "note": "股份转让批准，内部编号 M03",
    },
    "p2_m03_instrument": {
        "category": "变更 P2",
        "name": "M03 Instrument of Transfer",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M03_instrument_of_transfer_standard.docx",
        "note": "转让人 + 受让人签署，内部编号 M03",
    },
    "p2_m03_certificate": {
        "category": "变更 P2",
        "name": "M03 更新股权证书",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M03_updated_share_certificate_standard.docx",
        "note": "每个受让人一份，内部编号 M03",
    },
    "p2_m03_checklist": {
        "category": "变更 P2",
        "name": "M03 Register / Stamp Duty Checklist",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M03_register_and_stamp_duty_checklist_standard.docx",
        "note": "内部复核清单，内部编号 M03",
    },
    "p2_m03_register_members": {
        "category": "变更 P2",
        "name": "M03 Register of Members Update",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M03_register_of_members_update_standard.docx",
        "note": "股份转让后的成员名册更新记录，内部编号 M03",
    },
    "p2_m04_authority": {
        "category": "变更 P2",
        "name": "M04 S161 / 股东授权",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_s161_members_authority_standard.docx",
        "note": "股东授权董事发行股份，内部编号 M04",
    },
    "p2_m04_resolution": {
        "category": "变更 P2",
        "name": "M04 配股董事决议",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_allotment_directors_resolution_standard.docx",
        "note": "董事批准具体配股，内部编号 M04",
    },
    "p2_m04_application": {
        "category": "变更 P2",
        "name": "M04 Share Application",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_share_application_standard.docx",
        "note": "每名认购人一份，内部编号 M04",
    },
    "p2_m04_certificate": {
        "category": "变更 P2",
        "name": "M04 股权证书",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_share_certificate_standard.docx",
        "note": "每名认购人一份，内部编号 M04",
    },
    "p2_m04_form24": {
        "category": "变更 P2",
        "name": "M04 Form 24 / Return of Allotment",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_return_of_allotment_form24_standard.docx",
        "note": "增资配股才生成，内部编号 M04",
    },
    "p2_m04_checklist": {
        "category": "变更 P2",
        "name": "M04 Register Update Checklist",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_register_update_checklist_standard.docx",
        "note": "内部复核清单，内部编号 M04",
    },
    "p2_m04_register_members": {
        "category": "变更 P2",
        "name": "M04 Register of Members Update",
        "version": "v0.1",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M04_register_of_members_update_standard.docx",
        "note": "增资配股后的成员名册更新记录，内部编号 M04",
    },
    "p2_m05_agm_package": {
        "category": "年审 P2",
        "name": "M05 AGM 文件包",
        "version": "v0.2",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M05_agm_documents_package_standard.docx",
        "note": "DR + 普通 AGM 或书面年审/AGM 豁免路线，内部编号 M05",
    },
    "p2_m05_annual_return_package": {
        "category": "年审 P2",
        "name": "M05 Annual Return 授权声明包",
        "version": "v0.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M05_annual_return_authorisation_package_standard.docx",
        "note": "AR review / Section 197 / 动态审计或休眠声明 / AR 授权 / MRL，内部编号 M05",
    },
    "p2_m05_checklist": {
        "category": "年审 P2",
        "name": "M05 年审内部复核清单",
        "version": "v0.3",
        "status": "启用",
        "path": DOC_TEMPLATE_DIR / "p2_standard_v1" / "M05_annual_review_checklist_standard.docx",
        "note": "内部复核清单，内部编号 M05",
    },
}

TEMPLATE_REGISTRY_PATH = DATA_DIR / "template_registry.json"
TEMPLATE_BACKUP_DIR = DATA_DIR / "template_versions"


def h(value: object) -> str:
    return html.escape(str(value or ""))


def display_time(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        if text.endswith("Z"):
            dt = datetime.fromisoformat(text[:-1]).replace(tzinfo=UTC)
        else:
            dt = datetime.fromisoformat(text)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=UTC)
        return dt.astimezone(DISPLAY_TIMEZONE).strftime("%Y-%m-%d %H:%M:%S SGT")
    except ValueError:
        return text


def render_page(title: str, body: str, user: dict | None = None) -> bytes:
    nav = ""
    if user:
        admin_link = '<a href="/settings">管理员后台</a>' if user.get("role") == "admin" else ""
        nav = f"""
        <nav>
          <a href="/">上传</a>
          <a href="/jobs">任务记录</a>
          {admin_link}
          <span class="spacer"></span>
          <span>{h(user['username'])} · {h(user['role'])}</span>
          <a href="/logout">退出</a>
        </nav>
        """
    page = f"""<!doctype html>
    <html lang="zh-CN">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>{h(title)} · {h(APP_NAME)}</title>
      <link rel="stylesheet" href="/static/styles.css">
    </head>
    <body>
      <header><h1>{h(APP_NAME)}</h1>{nav}</header>
      <main>{body}</main>
    </body>
    </html>"""
    return page.encode("utf-8")


class App(BaseHTTPRequestHandler):
    server_version = "SecretaryFileServer/0.2"

    def log_message(self, format, *args):
        return

    def do_GET(self):
        try:
            self.route_get()
        except Exception as exc:
            self.error_page(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))

    def do_POST(self):
        try:
            self.route_post()
        except Exception as exc:
            self.error_page(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))

    def route_get(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/static/styles.css":
            return self.serve_file(Path(__file__).parent / "static" / "styles.css", "text/css; charset=utf-8", inline=True)
        if path.startswith("/generated/"):
            user = self.current_user()
            if not user:
                return self.redirect("/login")
            file_path = GENERATED_DIR / Path(path).name
            if file_path.exists():
                log_action(user["id"], "download_generated", file_path.name)
            return self.serve_file(file_path, content_type_for(file_path))
        if path.startswith("/templates/"):
            user = self.current_user()
            if not user:
                return self.redirect("/login")
            key = Path(path).name
            if key not in PUBLIC_TEMPLATE_KEYS and user.get("role") != "admin":
                return self.error_page(HTTPStatus.FORBIDDEN, "这个模板只开放给管理员。")
            item = TEMPLATE_DOWNLOADS.get(key)
            if not item:
                return self.error_page(HTTPStatus.NOT_FOUND, "模板不存在")
            if item[1].exists():
                log_action(user["id"], "download_import_template", key)
            return self.serve_file(item[1], content_type_for(item[1]))
        if path.startswith("/document-template/"):
            user = self.current_user()
            if not user:
                return self.redirect("/login")
            if user.get("role") != "admin":
                return self.error_page(HTTPStatus.FORBIDDEN, "正式文件模板只开放给管理员。")
            key = Path(path).name
            item = DOCUMENT_TEMPLATES.get(key)
            if not item:
                return self.error_page(HTTPStatus.NOT_FOUND, "文件模板不存在")
            if item["path"].exists():
                log_action(user["id"], "download_document_template", key)
            return self.serve_file(item["path"], content_type_for(item["path"]))
        if path == "/login":
            return self.login_page()
        if path == "/logout":
            self.clear_session()
            return self.redirect("/login")

        user = self.require_user()
        if not user:
            return
        if path == "/":
            return self.home_page(user)
        if path == "/p2-form":
            return self.p2_form_page(user)
        if path == "/jobs":
            return self.jobs_page(user)
        if path == "/job":
            job_id = parse_qs(parsed.query).get("id", [""])[0]
            return self.job_page(user, job_id)
        if path == "/settings":
            return self.settings_page(user, parse_qs(parsed.query))
        self.error_page(HTTPStatus.NOT_FOUND, "页面不存在")

    def route_post(self):
        parsed = urlparse(self.path)
        if parsed.path == "/login":
            return self.handle_login()
        user = self.require_user()
        if not user:
            return
        if parsed.path == "/upload":
            return self.handle_upload(user)
        if parsed.path == "/p2-form":
            return self.handle_p2_form(user)
        if parsed.path in {"/generate-placeholder", "/generate-p1"}:
            return self.handle_generate_p1(user)
        if parsed.path == "/generate-p2-m01":
            return self.handle_generate_p2_m01(user)
        if parsed.path == "/generate-p2-m02":
            return self.handle_generate_p2_m02(user)
        if parsed.path == "/generate-p2-m03":
            return self.handle_generate_p2_m03(user)
        if parsed.path == "/generate-p2-m04":
            return self.handle_generate_p2_m04(user)
        if parsed.path == "/generate-p2-m05":
            return self.handle_generate_p2_m05(user)
        if parsed.path == "/settings/common-people/save":
            return self.handle_save_common_person(user)
        if parsed.path == "/settings/users/save":
            return self.handle_save_user(user)
        if parsed.path == "/settings/users/toggle":
            return self.handle_toggle_user(user)
        if parsed.path == "/settings/template/upload":
            return self.handle_upload_document_template(user)
        if parsed.path == "/settings/template/activate-draft":
            return self.handle_activate_document_template_draft(user)
        if parsed.path == "/settings/template/rollback":
            return self.handle_rollback_document_template(user)
        if parsed.path == "/settings/common-people/toggle":
            return self.handle_toggle_common_person(user)
        if parsed.path == "/settings/common-people/delete":
            return self.handle_delete_common_person(user)
        if parsed.path == "/settings/common-people/disable-samples":
            return self.handle_disable_sample_people(user)
        if parsed.path == "/job/delete":
            return self.handle_delete_job(user)
        if parsed.path == "/jobs/delete-selected":
            return self.handle_delete_selected_jobs(user)
        if parsed.path == "/jobs/cleanup-old":
            return self.handle_cleanup_old_jobs(user)
        self.error_page(HTTPStatus.NOT_FOUND, "接口不存在")

    def login_page(self, message: str = ""):
        body = f"""
        <section class="panel narrow">
          <h2>登录</h2>
          <p class="muted">本地内部测试环境。上线或多人使用前，请修改默认账号密码。</p>
          {'<p class="error">' + h(message) + '</p>' if message else ''}
          <form method="post" action="/login" class="stack">
            <label>用户名<input name="username" autocomplete="username"></label>
            <label>密码<input name="password" type="password" autocomplete="current-password"></label>
            <button type="submit">登录</button>
          </form>
        </section>
        """
        self.send_html(render_page("登录", body))

    def home_page(self, user):
        p1_cards = public_template_cards(["registration_human_blank", "registration_human_sample"])
        p2_cards = public_template_cards(["maintenance_blank", "maintenance_sample"])
        body = f"""
        <section class="panel">
          <div class="toolbar">
            <div>
              <h2>文件生成工作台</h2>
              <p class="muted">选择对应业务表，拖拽或点击上传 Excel 后，系统会自动判断并生成 PDF 包。</p>
            </div>
            <span class="badge">内部试运行</span>
          </div>
          <form method="post" action="/upload" enctype="multipart/form-data" class="upload upload-drop-form" id="main-upload-form">
            <label class="upload-dropzone" id="main-upload-dropzone">
              <input class="upload-input" id="main-upload-input" type="file" name="file" accept=".xlsx" required>
              <span class="upload-icon">XLSX</span>
              <span class="upload-title">拖拽 Excel 到这里，或点击选择文件</span>
              <span class="upload-hint" id="main-upload-hint">支持 .xlsx；选择后会直接上传并生成任务。</span>
            </label>
            <button type="submit">上传并分析</button>
          </form>
          <script>
            (() => {{
              const form = document.getElementById("main-upload-form");
              const dropzone = document.getElementById("main-upload-dropzone");
              const input = document.getElementById("main-upload-input");
              const hint = document.getElementById("main-upload-hint");
              if (!form || !dropzone || !input || !hint) return;

              const submitIfReady = () => {{
                if (!input.files || input.files.length === 0) return;
                const file = input.files[0];
                if (!file.name.toLowerCase().endsWith(".xlsx")) {{
                  hint.textContent = "请上传 .xlsx 格式的业务表。";
                  dropzone.classList.add("upload-invalid");
                  return;
                }}
                hint.textContent = "正在上传：" + file.name;
                dropzone.classList.remove("upload-invalid");
                dropzone.classList.add("upload-selected");
                form.requestSubmit();
              }};

              input.addEventListener("change", submitIfReady);
              ["dragenter", "dragover"].forEach((eventName) => {{
                dropzone.addEventListener(eventName, (event) => {{
                  event.preventDefault();
                  dropzone.classList.add("is-dragover");
                }});
              }});
              ["dragleave", "drop"].forEach((eventName) => {{
                dropzone.addEventListener(eventName, (event) => {{
                  event.preventDefault();
                  dropzone.classList.remove("is-dragover");
                }});
              }});
              dropzone.addEventListener("drop", (event) => {{
                const files = event.dataTransfer && event.dataTransfer.files;
                if (!files || files.length === 0) return;
                input.files = files;
                submitIfReady();
              }});
            }})();
          </script>
        </section>
        <section class="panel">
          <h3>选择业务入口</h3>
          <div class="operation-grid">
            <div class="operation-card">
              <strong>注册文件生成</strong>
              <p>用于新加坡新公司注册后整套签字文件。适合董事、秘书、股东、Form 24、RORC、股权证书和初始股东名册。</p>
              <div class="download-grid compact">{p1_cards or '<p class="muted">未找到注册导入模板。</p>'}</div>
            </div>
            <div class="operation-card">
              <strong>变更 / 年审文件生成</strong>
              <p>用于普通董事决议、转入、股份转让、增资配股和年审包。系统会根据表格内容自动判断要生成哪些文件。</p>
              <div class="button-row"><a class="button-link secondary" href="/p2-form">打开 P2 快速表单</a></div>
              <div class="download-grid compact">{p2_cards or '<p class="muted">未找到变更/年审导入模板。</p>'}</div>
            </div>
          </div>
        </section>
        <section class="grid">
          <div class="panel">
            <h3>日常操作流程</h3>
            <ol>
              <li>下载对应业务表，人工或 AI 辅助整理资料。</li>
              <li>上传 Excel，先看系统检查和生成预览。</li>
              <li>确认文件清单、签字人和生成份数。</li>
              <li>生成 PDF 文件包，发给客户签字或内部复核。</li>
            </ol>
          </div>
          <div class="panel">
            <h3>当前可生成</h3>
            <p>新公司注册、普通董事决议、转入、股份转让、增资配股和年审包都已接入 PDF 生成。注册、转股和配股会同步生成 Register of Members 相关文件。</p>
          </div>
          <div class="panel">
            <h3>默认填写规则</h3>
            <p>空白或 Auto 通常表示由系统自动判断；明确填写 No 才会关闭对应事项。公司资料变更填写新值即可触发；人员任免、转股和配股填写核心行信息即可触发。签字人留空时系统会提醒人工复核。</p>
          </div>
        </section>
        <section class="panel warning">
          <strong>上线前安全提醒</strong>
          <p>当前是本地内部工具。云端部署前需要更换默认密码、启用 HTTPS、限制访问范围，并整理正式用户权限。</p>
        </section>
        """
        self.send_html(render_page("上传", body, user))

    def p2_form_page(self, user):
        common_people = active_common_people_map()
        common_options = common_people_options(common_people)
        today = singapore_date_input()
        body = f"""
        <section class="panel">
          <div class="toolbar">
            <div>
              <h2>P2 快速表单</h2>
              <p class="muted">按业务事项填写；提交后系统会直接创建任务并自动生成可用的 PDF 包。</p>
            </div>
            <a class="button-link secondary" href="/">返回上传入口</a>
          </div>
        </section>
        <form method="post" action="/p2-form" class="p2-form">
          <datalist id="common-people">{common_options}</datalist>
          <section class="panel">
            <h3>公司基本信息</h3>
            <div class="form-grid">
              <label>公司名称<input name="company_name" required placeholder="EXAMPLE PTE. LTD."></label>
              <label>UEN<input name="uen" placeholder="202400000A"></label>
              <label>文件日期<input name="default_document_date" type="date" value="{h(today)}"></label>
              <label>币种<input name="currency" value="SGD"></label>
              <label class="wide">当前注册地址<textarea name="registered_office_address" rows="2"></textarea></label>
              <label class="wide">新注册地址<textarea name="new_registered_office_address" rows="2" placeholder="留空则不触发注册地址变更；转入时可默认填 RSIN 地址"></textarea></label>
              <label>董事签字人<input name="director_signer_names" list="common-people" placeholder="多个用逗号隔开"></label>
              <label>股东/客户签字人<input name="member_signer_names" placeholder="多个用逗号隔开"></label>
              <label>当前已发行股本<input name="issued_share_capital" placeholder="例如 80000"></label>
              <label>当前实缴股本<input name="paid_up_capital" placeholder="例如 80000"></label>
            </div>
          </section>

          <section class="panel">
            <h3>M01 普通变更 / 董事决议</h3>
            <div class="form-grid">
              <label class="check"><input type="checkbox" name="change_registered_office" value="1"> 注册地址变更</label>
              <label class="check"><input type="checkbox" name="change_business_activity" value="1"> 营业范围 / SSIC 变更</label>
              <label class="check"><input type="checkbox" name="change_fye" value="1"> 财年日变更</label>
              <label>新主 SSIC<input name="primary_ssic_new"></label>
              <label>新主营业务<input name="primary_activity_new"></label>
              <label>新副 SSIC<input name="secondary_ssic_new"></label>
              <label>新副营业务<input name="secondary_activity_new"></label>
              <label>新财年日<input name="new_fye" type="date"></label>
              <label>现任董事辞任<input name="resign_director_names" placeholder="多个用逗号隔开"></label>
              <label>新董事委任<input name="appoint_director_names" list="common-people" placeholder="多个用逗号隔开"></label>
              <label>前秘书辞任<input name="resign_secretary_names" placeholder="多个用逗号隔开"></label>
              <label>新秘书委任<input name="appoint_secretary_names" list="common-people" placeholder="多个用逗号隔开"></label>
              <label class="check"><input type="checkbox" name="resignation_letter" value="1"> 为辞任人员生成辞职信</label>
            </div>
          </section>

          <section class="panel">
            <h3>M02 转入</h3>
            <div class="form-grid">
              <label class="check wide"><input type="checkbox" name="transfer_in_required" value="1"> 生成转入 EGM / 交接文件包</label>
              <label>前秘书公司名称<input name="old_secretary_company" placeholder="不知道可留空"></label>
              <label>新秘书公司<input name="new_secretary_company" value="RSIN GROUP PTE. LTD."></label>
              <label>转入新挂名董事<input name="transfer_in_nominee_director_names" list="common-people" placeholder="例如 LE THI NGOC TRANG"></label>
              <label>转入新秘书<input name="transfer_in_secretary_names" list="common-people" placeholder="例如 FENDI CHANDRA TING S ING EE"></label>
              <label class="check wide"><input type="checkbox" name="transfer_in_resignation_letter" value="1"> 如有辞任人员，同包生成辞职信</label>
            </div>
          </section>

          <section class="panel">
            <h3>M03 股份转让</h3>
            <div class="form-grid">
              <label class="check wide"><input type="checkbox" name="share_transfer_required" value="1"> 生成股份转让包</label>
              <label>转让人<input name="transferor_name"></label>
              <label>受让人<input name="transferee_name"></label>
              <label>股份数量<input name="shares_transferred"></label>
              <label>股份类别<input name="transfer_share_class" value="Ordinary"></label>
              <label>转让日期<input name="transfer_date" type="date"></label>
              <label>对价金额<input name="consideration_amount" placeholder="通常可留空"></label>
              <label>对价逻辑<select name="consideration_basis">
                <option value="internal_paid_up_basis">内部转让 / 实缴比例</option>
                <option value="acra_paid_up_capital_basis">ACRA 登记实缴资本</option>
                <option value="stamp_duty_higher_of_price_or_nav">需按印花税 / NAV 复核</option>
              </select></label>
            </div>
          </section>

          <section class="panel">
            <h3>M04 增资配股</h3>
            <div class="form-grid">
              <label class="check wide"><input type="checkbox" name="share_allotment_required" value="1"> 生成增资配股包</label>
              <label>认购人 / 获配人<input name="allottee_name"></label>
              <label>配发股份数量<input name="shares_allotted"></label>
              <label>股份类别<input name="allotment_share_class" value="Ordinary"></label>
              <label>新增已发行股本<input name="allotment_issued_share_capital"></label>
              <label>新增实缴股本<input name="allotment_paid_up_share_capital"></label>
              <label>配股日期<input name="allotment_date" type="date"></label>
            </div>
          </section>

          <section class="panel">
            <h3>M05 年审</h3>
            <div class="form-grid">
              <label class="check wide"><input type="checkbox" name="annual_review_required" value="1"> 生成年审包</label>
              <label>FYE<input name="fye_date" type="date"></label>
              <label>AGM / 年审文件日期<input name="agm_date" type="date"></label>
              <label>AGM 路线<select name="agm_route">
                <option value="ordinary_agm">普通 AGM</option>
                <option value="written_resolution">书面年审 / 股东书面决议</option>
                <option value="exempt_or_dispensed">AGM 豁免 / dispense</option>
              </select></label>
              <label>财报状态<select name="accounts_status">
                <option value="non_dormant">非休眠 / 常规小公司</option>
                <option value="dormant">休眠</option>
                <option value="audited">已审计</option>
              </select></label>
              <label>董事费<input name="directors_fee" value="0"></label>
              <label>董事薪酬<input name="directors_remuneration" value="0"></label>
            </div>
          </section>

          <section class="panel">
            <h3>个人资料变更</h3>
            <div class="form-grid">
              <label>人员姓名<input name="particular_person_name" list="common-people"></label>
              <label>变更项目<select name="particular_field_label">
                <option value="">不生成</option>
                <option value="Residential address">地址</option>
                <option value="ID number">证件号</option>
                <option value="Email">邮箱</option>
                <option value="Phone">电话</option>
                <option value="Nationality">国籍</option>
              </select></label>
              <label class="wide">旧信息<textarea name="particular_old_value" rows="2"></textarea></label>
              <label class="wide">新信息<textarea name="particular_new_value" rows="2"></textarea></label>
            </div>
          </section>

          <section class="panel">
            <button type="submit">创建任务并生成 PDF</button>
            <a class="button-link secondary" href="/">取消</a>
          </section>
        </form>
        """
        self.send_html(render_page("P2 快速表单", body, user))

    def handle_p2_form(self, user):
        fields = self.read_form_urlencoded()
        common = active_common_people_map()
        parsed = build_p2_form_parsed(fields, common)
        suggestions = suggest_files(parsed)
        company_name = parsed.get("company", {}).get("company_name", "")
        summary = suggestions.get("summary", {})
        status = "blocked" if summary.get("blocking_errors") else "needs_review"
        job_code = f"JOB-{datetime.now(UTC).strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(2).upper()}"
        source_filename = f"P2_WEB_FORM_{safe_filename(company_name or 'company')}.form"
        metadata = case_metadata_from_parsed(parsed, source_filename, job_code)
        with connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO generation_jobs
                (job_code, case_id, business_order_id, source_type, source_file_id, contact_person_id, agent_person_id,
                 client_signatory_person_id, authorized_representative_person_id, prepared_by, snapshot_version,
                 task_type, company_name, source_filename, upload_path, parsed_json, suggestions_json, status, created_by, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    job_code,
                    metadata["case_id"],
                    metadata["business_order_id"],
                    metadata["source_type"],
                    metadata["source_file_id"],
                    metadata["contact_person_id"],
                    metadata["agent_person_id"],
                    metadata["client_signatory_person_id"],
                    metadata["authorized_representative_person_id"],
                    metadata["prepared_by"],
                    metadata["snapshot_version"],
                    parsed.get("task_type", "maintenance"),
                    company_name,
                    source_filename,
                    "web-form:p2",
                    to_json(parsed),
                    json.dumps(suggestions, ensure_ascii=False, indent=2),
                    status,
                    user["id"],
                    now(),
                ),
            )
            job_id = cur.lastrowid
        log_action(user["id"], "submit_p2_form", source_filename)
        self.queue_auto_pdf_generation(user["id"], job_id, job_code, parsed, suggestions)
        self.redirect(f"/job?id={job_id}")

    def jobs_page(self, user):
        with connect() as conn:
            rows = conn.execute(
                "SELECT id, job_code, case_id, task_type, company_name, source_filename, status, created_at FROM generation_jobs ORDER BY id DESC LIMIT 80"
            ).fetchall()
        is_admin = user.get("role") == "admin"
        trs = "".join(job_table_row(r, is_admin) for r in rows)
        first_head = "任务编号" if is_admin else "任务"
        select_head = "<th>选择</th>" if is_admin else ""
        empty_colspan = "7" if is_admin else "6"
        table = f"""
          <table>
            <thead><tr>{select_head}<th>{first_head}</th><th>类型</th><th>公司</th><th>来源文件</th><th>状态</th><th>时间</th></tr></thead>
            <tbody>{trs or f'<tr><td colspan="{empty_colspan}">暂无任务</td></tr>'}</tbody>
          </table>
        """
        if is_admin:
            table = f"""
          <form method="post" action="/jobs/delete-selected">
            {table}
            <div class="button-row">
              <button class="danger" type="submit">删除选中任务</button>
            </div>
          </form>
          <form method="post" action="/jobs/cleanup-old" class="inline-form cleanup-form">
            <label>清理超过 <input type="number" name="days" min="7" max="365" value="30"> 天的旧任务</label>
            <button class="secondary-action" type="submit">清理旧任务</button>
          </form>
          <p class="muted">删除任务会同时清理上传源文件和已生成的 PDF 包；下载 PDF、模板和导入表会记录在后台日志里。</p>
            """
        body = f"""
        <section class="panel">
          <h2>任务记录</h2>
          {table}
        </section>
        """
        self.send_html(render_page("任务记录", body, user))

    def job_page(self, user, job_id: str):
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在")

        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        summary = suggestions.get("summary", {})
        blocking_errors = summary.get("blocking_errors", [])
        warnings = summary.get("warnings", [])
        info = summary.get("info", [])
        is_generating = is_generation_status(row["status"])
        can_generate = row["task_type"] == "incorporation" and not blocking_errors and not is_generating
        can_generate_m01 = row["task_type"] == "maintenance" and not blocking_errors and not is_generating and summary.get("m01_available") == "Yes"
        can_generate_m02 = row["task_type"] == "maintenance" and not blocking_errors and not is_generating and summary.get("m02_available") == "Yes"
        can_generate_m03 = row["task_type"] == "maintenance" and not blocking_errors and not is_generating and summary.get("m03_available") == "Yes"
        can_generate_m04 = row["task_type"] == "maintenance" and not blocking_errors and not is_generating and summary.get("m04_available") == "Yes"
        can_generate_m05 = row["task_type"] == "maintenance" and not blocking_errors and not is_generating and summary.get("m05_available") == "Yes"
        is_admin = user.get("role") == "admin"

        generated_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P1_pdf_package.zip"
        generated_m01_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P2_M01_pdf_package.zip"
        generated_m02_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P2_M02_pdf_package.zip"
        generated_m03_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P2_M03_pdf_package.zip"
        generated_m04_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P2_M04_pdf_package.zip"
        generated_m05_zip = GENERATED_DIR / f"{safe_filename(row['job_code'])}_P2_M05_pdf_package.zip"
        generated_packages = [
            ("P1 注册文件包", generated_zip),
            ("M01 普通变更 DR 包", generated_m01_zip),
            ("M02 转入文件包", generated_m02_zip),
            ("M03 股份转让包", generated_m03_zip),
            ("M04 增资配股包", generated_m04_zip),
            ("M05 年审文件包", generated_m05_zip),
        ]
        has_download = any(path.exists() for _, path in generated_packages)
        generated_link = (
            f"<a class='button-link' href='/generated/{h(generated_zip.name)}'>{h(download_label('P1', user))}</a>"
            if generated_zip.exists()
            else ""
        )
        generated_m01_link = (
            f"<a class='button-link secondary' href='/generated/{h(generated_m01_zip.name)}'>{h(download_label('M01', user))}</a>"
            if generated_m01_zip.exists()
            else ""
        )
        generated_m02_link = (
            f"<a class='button-link secondary' href='/generated/{h(generated_m02_zip.name)}'>{h(download_label('M02', user))}</a>"
            if generated_m02_zip.exists()
            else ""
        )
        generated_m03_link = (
            f"<a class='button-link secondary' href='/generated/{h(generated_m03_zip.name)}'>{h(download_label('M03', user))}</a>"
            if generated_m03_zip.exists()
            else ""
        )
        generated_m04_link = (
            f"<a class='button-link secondary' href='/generated/{h(generated_m04_zip.name)}'>{h(download_label('M04', user))}</a>"
            if generated_m04_zip.exists()
            else ""
        )
        generated_m05_link = (
            f"<a class='button-link secondary' href='/generated/{h(generated_m05_zip.name)}'>{h(download_label('M05', user))}</a>"
            if generated_m05_zip.exists()
            else ""
        )
        download_links = f"<div class='button-row'>{generated_link}{generated_m01_link}{generated_m02_link}{generated_m03_link}{generated_m04_link}{generated_m05_link}</div>" if has_download else ""
        generation_notice = (
            """
            <div class="info-box">
              <strong>PDF 正在后台生成</strong>
              <p>页面会自动刷新；生成完成后下载按钮会出现在这里。你也可以稍后回到任务列表打开。</p>
            </div>
            <script>setTimeout(() => window.location.reload(), 5000);</script>
            """
            if is_generating
            else ""
        )
        generation_notice = auto_generation_notice(is_generating)
        generate_control = (
            "<p class='muted'>PDF 正在后台生成，请稍等。</p>"
            if is_generating
            else f"""
            <form method="post" action="/generate-p1">
              <input type="hidden" name="job_id" value="{h(row['id'])}">
              <button type="submit">{h(generate_button_label('P1', user))}</button>
              {dev_code('P1', user)}
            </form>
            """
            if can_generate
            else "<p class='error'>当前存在严重错误或不是注册任务，不能生成注册文件包。</p>"
        )
        if row["task_type"] in {"change", "maintenance"}:
            if is_generating:
                generate_control = "<p class='muted'>PDF 正在后台生成，请稍等。</p>"
            elif can_generate_m01 or can_generate_m02 or can_generate_m03 or can_generate_m04 or can_generate_m05:
                maintenance_forms = []
                if can_generate_m01:
                    maintenance_forms.append(
                        f"""
                        <form method="post" action="/generate-p2-m01">
                          <input type="hidden" name="job_id" value="{h(row['id'])}">
                          <button type="submit">{h(generate_button_label('M01', user))}</button>
                          {dev_code('M01', user)}
                        </form>
                        """
                    )
                if can_generate_m02:
                    maintenance_forms.append(
                        f"""
                        <form method="post" action="/generate-p2-m02">
                          <input type="hidden" name="job_id" value="{h(row['id'])}">
                          <button type="submit">{h(generate_button_label('M02', user))}</button>
                          {dev_code('M02', user)}
                        </form>
                        """
                    )
                if can_generate_m03:
                    maintenance_forms.append(
                        f"""
                        <form method="post" action="/generate-p2-m03">
                          <input type="hidden" name="job_id" value="{h(row['id'])}">
                          <button type="submit">{h(generate_button_label('M03', user))}</button>
                          {dev_code('M03', user)}
                        </form>
                        """
                    )
                if can_generate_m04:
                    maintenance_forms.append(
                        f"""
                        <form method="post" action="/generate-p2-m04">
                          <input type="hidden" name="job_id" value="{h(row['id'])}">
                          <button type="submit">{h(generate_button_label('M04', user))}</button>
                          {dev_code('M04', user)}
                        </form>
                        """
                    )
                if can_generate_m05:
                    maintenance_forms.append(
                        f"""
                        <form method="post" action="/generate-p2-m05">
                          <input type="hidden" name="job_id" value="{h(row['id'])}">
                          <button type="submit">{h(generate_button_label('M05', user))}</button>
                          {dev_code('M05', user)}
                        </form>
                        """
                    )
                generate_control = f"""
                <div class="stack">
                  {''.join(maintenance_forms)}
                </div>
                <p class="muted">已接入普通董事决议、转入文件、股份转让、增资配股和年审文件包。</p>
                """
            elif row["task_type"] == "maintenance":
                generate_control = "<p class='muted'>这张表没有识别到已接入生成器的事项；文件包先保留为预览。</p>"
            else:
                generate_control = "<p class='muted'>旧变更 v2 表当前只做判断预览；请使用公司维护/变更年审 v3 表生成正式文件。</p>"

        can_generate_any = bool(can_generate or can_generate_m01 or can_generate_m02 or can_generate_m03 or can_generate_m04 or can_generate_m05)
        generate_control = auto_generation_control(row, has_download, can_generate_any, bool(blocking_errors))

        admin_details = ""
        if is_admin:
            admin_details = f"""
            <section class="panel">
              <details>
                <summary>管理员调试数据</summary>
                <p class="muted">任务编号：{h(row['job_code'])}；内部操作编号会保留在管理员视角，方便开发和排查。</p>
                <h4>解析数据</h4>
                <pre>{h(json.dumps(parsed, ensure_ascii=False, indent=2))}</pre>
                <h4>判断结果</h4>
                <pre>{h(json.dumps(suggestions, ensure_ascii=False, indent=2))}</pre>
              </details>
            </section>
            """

        body = f"""
        <section class="panel">
          <div class="toolbar">
            <div>
              <h2>{h(row['company_name']) or h(row['job_code'])}</h2>
              <p class="muted">{h(task_label(row['task_type']))} · {h(row['source_filename'])}{' · 管理员任务编号 ' + h(row['job_code']) if is_admin else ''}</p>
            </div>
            <span class="badge">{h(status_label(row['status']))}</span>
          </div>
          {generation_notice}
          {download_links}
          {generated_files_panel(generated_packages)}
          {review_progress(row, summary, can_generate, can_generate_m01 or can_generate_m02 or can_generate_m03 or can_generate_m04 or can_generate_m05, bool(generated_link or generated_m01_link or generated_m02_link or generated_m03_link or generated_m04_link or generated_m05_link))}
          {alert_block("严重错误", blocking_errors, "error-box")}
          {review_workflow(summary, warnings, blocking_errors, user)}
          {alert_block("系统说明", info, "info-box")}
          {summary_cards(summary)}
          {case_metadata_panel(row, user)}
          {package_status_cards(suggestions.get("files", []), summary, user)}
        </section>
        <section class="panel">
          <h3>下一步操作</h3>
          {generation_steps(row, summary, can_generate, can_generate_m01 or can_generate_m02 or can_generate_m03 or can_generate_m04 or can_generate_m05, user)}
        </section>
        <section class="panel">
          <h3>文件预览</h3>
          {preview_table(suggestions.get("preview", []))}
          {generate_control}
        </section>
        <section class="panel">
          <h3>{'管理员文件规则明细' if is_admin else '文件清单'}</h3>
          {files_table(suggestions.get("files", []), user)}
        </section>
        {admin_details}
        {delete_job_form(row, user)}
        """
        self.send_html(render_page("任务详情", body, user))

    def settings_page(self, user, query: dict[str, list[str]] | None = None):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以打开后台设置。")
        query = query or {}
        show_archived_templates = query.get("show_archived_templates", [""])[0] == "1"
        edit_user_id = query.get("edit_user", [""])[0]
        edit_person_id = query.get("edit_person", [""])[0]
        with connect() as conn:
            users = conn.execute("SELECT id, username, role, active, last_login_at, created_at FROM users ORDER BY id").fetchall()
            people = conn.execute("SELECT * FROM common_people ORDER BY display_name").fetchall()
            rules = conn.execute("SELECT * FROM template_rules ORDER BY task_type, change_item").fetchall()
            job_stats = conn.execute(
                """
                SELECT
                  COUNT(*) AS total_jobs,
                  SUM(CASE WHEN status = 'blocked' THEN 1 ELSE 0 END) AS blocked_jobs,
                  SUM(CASE WHEN status LIKE '%pdf_generated' OR status = 'pdf_generated' THEN 1 ELSE 0 END) AS pdf_jobs
                FROM generation_jobs
                """
            ).fetchone()
            login_logs = conn.execute(
                """
                SELECT audit_logs.*, users.username
                FROM audit_logs LEFT JOIN users ON users.id = audit_logs.user_id
                WHERE audit_logs.action IN (
                  'login', 'login_failed', 'save_user', 'toggle_user',
                  'save_common_person', 'toggle_common_person', 'delete_common_person', 'disable_sample_people',
                  'upload_document_template_draft', 'activate_document_template_draft', 'rollback_document_template',
                  'download_generated', 'download_import_template', 'download_document_template',
                  'delete_job', 'delete_jobs', 'cleanup_old_jobs'
                )
                ORDER BY audit_logs.id DESC
                LIMIT 30
                """
            ).fetchall()
        edit_user = next((row for row in users if str(row["id"]) == edit_user_id), None)
        edit_person = next((row for row in people if str(row["id"]) == edit_person_id), None)
        user_rows = "".join(
            user_table_row(u, user)
            for u in users
        )
        people_rows = "".join(
            common_person_table_row(p)
            for p in people
        )
        login_rows = "".join(login_log_row(row) for row in login_logs)
        rule_rows = "".join(
            f"<tr><td>{h(r['rule_key'])}</td><td>{h(task_label(r['task_type']))}</td><td>{h(r['change_item'])}</td><td>{h(r['suggested_file'])}</td><td>{h(r['signing_mode'])}</td><td>{h(r['signer_source'])}</td><td>{'可合并' if r['can_merge_dr'] else ''}</td><td>{'需复核' if r['manual_review'] else ''}</td></tr>"
            for r in rules
        )
        overview_html = admin_overview_html(users, people, rules, job_stats)
        template_rows = template_library_rows_html(show_archived_templates)
        import_template_toggle = (
            "<p class='muted template-toggle'>当前只显示正式入口。<a href='/settings?show_archived_templates=1#templates'>显示旧版 / 测试 / 辅助入口</a></p>"
            if not show_archived_templates
            else "<p class='muted template-toggle'>已显示旧版 / 测试 / 辅助入口。<a href='/settings#templates'>收起旧入口</a></p>"
        )
        document_template_sections = document_template_sections_html()
        user_form_open = " open" if edit_user else ""
        person_form_open = " open" if edit_person else ""
        body = f"""
        <section class="panel">
          <div class="toolbar">
            <div>
              <h2>管理员后台</h2>
              <p class="muted">这里保留内部操作编号、模板母版和规则明细。普通用户只看到业务流程和可生成文件。</p>
            </div>
            <span class="badge">Admin</span>
          </div>
          <div class="settings-nav">
            <a href="#users">用户账户</a>
            <a href="#people">常用人员</a>
            <a href="#templates">模板库</a>
            <a href="#rules">文件规则</a>
            <a href="#logs">操作记录</a>
          </div>
        </section>
        {overview_html}
        <details class="admin-overview-drawer admin-logic-drawer">
          <summary>
            <span>系统逻辑说明</span>
            <small>字段来源、默认值和当前文件包能力</small>
          </summary>
          <section class="grid">
            <div class="panel">
              <h3>字段来源</h3>
              <div class="source-grid">
                <div><strong>Excel 资料表</strong><span>公司、人员、股份、变更事项和年审数据的主来源。</span></div>
                <div><strong>后台默认值</strong><span>常用秘书、挂名董事、默认地址或联系方式，可被 Excel 覆盖。</span></div>
                <div><strong>系统自动计算</strong><span>例如首个 FYE、文件份数、签字人和可合并的 DR 事项。</span></div>
              </div>
            </div>
            <div class="panel">
              <h3>当前生成能力</h3>
              <div class="source-grid">
                <div><strong>注册文件包</strong><span>内部编号 P1，已生成 PDF 包。</span></div>
                <div><strong>普通董事决议</strong><span>内部编号 M01，已生成 PDF 包。</span></div>
                <div><strong>转入文件包</strong><span>内部编号 M02，已按两份签字 PDF 生成。</span></div>
                <div><strong>股份转让包</strong><span>内部编号 M03，已生成转股决议、Instrument、股权证书、Register 更新和内部清单。</span></div>
                <div><strong>增资配股包</strong><span>内部编号 M04，已生成授权、配股、申请书、证书、Form 24 和 Register 更新。</span></div>
                <div><strong>年审文件包</strong><span>内部编号 M05，已生成 AGM、AR 授权和复核清单。</span></div>
              </div>
            </div>
          </section>
        </details>
        <section class="panel" id="users">
          <div class="section-header">
            <div>
              <h3>用户账户</h3>
              <p class="muted">管理员维护账号和权限；普通用户只能上传、复核、生成和下载文件。</p>
            </div>
            <span class="badge">{len(users)} 个账号</span>
          </div>
          <details class="admin-drawer"{user_form_open}>
            <summary>{'编辑用户' if edit_user else '新增用户 / 重置密码'}</summary>
            {user_form_html(edit_user)}
          </details>
          <div class="table-wrap"><table><thead><tr><th>用户名</th><th>角色</th><th>状态</th><th>最后登录</th><th>创建时间</th><th>操作</th></tr></thead><tbody>{user_rows or '<tr><td colspan="6">暂无用户</td></tr>'}</tbody></table></div>
        </section>
        <section class="panel" id="people">
          <div class="section-header">
            <div>
              <h3>常用人员</h3>
              <p class="muted">用于 Excel 里调用常用秘书、挂名董事、本地董事或客户签字人。示例数据可以停用，也可以直接改成真实资料。</p>
            </div>
            <form method="post" action="/settings/common-people/disable-samples" class="inline-form">
              <button class="secondary-action" type="submit">停用全部示例人员</button>
            </form>
          </div>
          <details class="admin-drawer"{person_form_open}>
            <summary>{'编辑常用人员' if edit_person else '新增常用人员'}</summary>
            {common_person_form_html(edit_person)}
          </details>
          <div class="table-wrap"><table><thead><tr><th>名称</th><th>匹配简称</th><th>身份/职务</th><th>证件类型</th><th>证件号</th><th>国籍</th><th>电话</th><th>状态</th><th>签名</th><th>备注</th><th>操作</th></tr></thead><tbody>{people_rows or '<tr><td colspan="11">暂无常用人员</td></tr>'}</tbody></table></div>
        </section>
        <section class="panel" id="templates">
          <div class="section-header">
            <div>
              <h3>模板库</h3>
              <p class="muted">Excel 导入模板给员工下载；正式签字文件模板只在管理员后台维护。上传新版会先生成草稿 PDF，确认后再启用。</p>
            </div>
            <span class="badge">正式模板 {len(DOCUMENT_TEMPLATES)} 个</span>
          </div>
          <h4>正式签字文件模板</h4>
          {document_template_sections}
          <h4>Excel 导入表和辅助文件</h4>
          {import_template_toggle}
          <div class="table-wrap"><table><thead><tr><th>内部键</th><th>业务名称</th><th>说明</th><th>可见范围</th><th>状态</th><th>下载</th></tr></thead><tbody>{template_rows}</tbody></table></div>
        </section>
        <section class="panel" id="rules">
          <div class="section-header">
            <div>
              <h3>文件判断规则</h3>
              <p class="muted">这里记录后台规则和内部编号，供管理员和开发排查。普通用户只看到业务名称。</p>
            </div>
            <span class="badge">{len(rules)} 条规则</span>
          </div>
          <div class="table-wrap"><table><thead><tr><th>规则编号</th><th>任务类型</th><th>事项</th><th>建议文件</th><th>签字方式</th><th>签字来源</th><th>合并</th><th>复核</th></tr></thead><tbody>{rule_rows or '<tr><td colspan="8">暂无规则</td></tr>'}</tbody></table></div>
        </section>
        <section class="panel" id="logs">
          <div class="section-header">
            <div>
              <h3>最近系统操作</h3>
              <p class="muted">保留登录、下载、模板变更和人员维护记录，方便内部追踪。</p>
            </div>
            <span class="badge">最近 30 条</span>
          </div>
          <div class="table-wrap"><table><thead><tr><th>时间</th><th>用户</th><th>动作</th><th>详情</th></tr></thead><tbody>{login_rows or '<tr><td colspan="4">暂无记录</td></tr>'}</tbody></table></div>
        </section>
        """
        self.send_html(render_page("管理员后台", body, user))

    def handle_login(self):
        fields = self.read_form_urlencoded()
        username = fields.get("username", [""])[0]
        password = fields.get("password", [""])[0]
        with connect() as conn:
            user = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
            if not user or not verify_password(password, user["password_hash"]):
                log_action(None, "login_failed", username)
                return self.login_page("用户名或密码错误")
            if not user["active"]:
                log_action(user["id"], "login_failed", f"{username}:disabled")
                return self.login_page("这个账号已停用，请联系管理员。")
            token = secrets.token_urlsafe(32)
            expires = (datetime.now(UTC) + timedelta(hours=10)).replace(tzinfo=None).isoformat(timespec="seconds") + "Z"
            conn.execute(
                "INSERT INTO sessions (token, user_id, created_at, expires_at) VALUES (?, ?, ?, ?)",
                (token, user["id"], now(), expires),
            )
            conn.execute("UPDATE users SET last_login_at = ? WHERE id = ?", (now(), user["id"]))
        log_action(user["id"], "login", username)
        self.send_response(302)
        self.send_header("Location", "/")
        self.send_header("Set-Cookie", f"{SESSION_COOKIE}={token}; HttpOnly; SameSite=Lax; Path=/")
        self.end_headers()

    def handle_upload(self, user):
        cleanup_old_uploads()
        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD_BYTES:
            return self.error_page(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "文件太大")
        body = self.rfile.read(length)
        filename, file_bytes = parse_multipart_file(content_type, body, "file")
        if not filename or not file_bytes:
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有收到文件")
        if not filename.lower().endswith(".xlsx"):
            return self.error_page(HTTPStatus.BAD_REQUEST, "请上传 .xlsx 文件")
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = f"{datetime.now(UTC).strftime('%Y%m%d%H%M%S')}_{secrets.token_hex(4)}_{Path(filename).name}"
        upload_path = UPLOAD_DIR / safe_name
        upload_path.write_bytes(file_bytes)

        common = active_common_people_map()
        parsed = parse_excel(upload_path, common)
        suggestions = suggest_files(parsed)
        company_name = parsed.get("company", {}).get("company_name", "")
        summary = suggestions.get("summary", {})
        status = "blocked" if summary.get("blocking_errors") else "needs_review"
        job_code = f"JOB-{datetime.now(UTC).strftime('%Y%m%d-%H%M%S')}-{secrets.token_hex(2).upper()}"
        metadata = case_metadata_from_parsed(parsed, filename, job_code)
        with connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO generation_jobs
                (job_code, case_id, business_order_id, source_type, source_file_id, contact_person_id, agent_person_id,
                 client_signatory_person_id, authorized_representative_person_id, prepared_by, snapshot_version,
                 task_type, company_name, source_filename, upload_path, parsed_json, suggestions_json, status, created_by, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    job_code,
                    metadata["case_id"],
                    metadata["business_order_id"],
                    metadata["source_type"],
                    metadata["source_file_id"],
                    metadata["contact_person_id"],
                    metadata["agent_person_id"],
                    metadata["client_signatory_person_id"],
                    metadata["authorized_representative_person_id"],
                    metadata["prepared_by"],
                    metadata["snapshot_version"],
                    parsed.get("task_type", "unknown"),
                    company_name,
                    filename,
                    str(upload_path),
                    to_json(parsed),
                    json.dumps(suggestions, ensure_ascii=False, indent=2),
                    status,
                    user["id"],
                    now(),
                ),
            )
            job_id = cur.lastrowid
        log_action(user["id"], "upload", filename)
        self.queue_auto_pdf_generation(user["id"], job_id, job_code, parsed, suggestions)
        self.redirect(f"/job?id={job_id}")

    def queue_pdf_generation(self, user_id: int, row, parsed: dict, generator, generating_status: str, done_status: str, action: str) -> None:
        job_id = row["id"]
        job_code = row["job_code"]
        with connect() as conn:
            conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", (generating_status, job_id))

        def worker() -> None:
            try:
                generator(parsed, job_code)
                with connect() as conn:
                    conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", (done_status, job_id))
                log_action(user_id, action, job_code)
            except Exception:
                detail = traceback.format_exc(limit=8)
                with connect() as conn:
                    conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", ("generation_failed", job_id))
                log_action(user_id, f"{action}_failed", f"{job_code}: {detail[-1800:]}")

        threading.Thread(target=worker, name=f"pdf-generation-{job_code}", daemon=True).start()

    def queue_auto_pdf_generation(self, user_id: int, job_id: int, job_code: str, parsed: dict, suggestions: dict) -> None:
        tasks = auto_generation_tasks(parsed, suggestions)
        if not tasks:
            return
        with connect() as conn:
            conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", ("generating_pdf", job_id))

        def worker() -> None:
            generated: list[str] = []
            try:
                for code, generator in tasks:
                    generator(parsed, job_code)
                    generated.append(code)
                with connect() as conn:
                    conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", ("pdf_generated", job_id))
                log_action(user_id, "auto_generate_pdf_packages", f"{job_code}: {', '.join(generated)}")
            except Exception:
                detail = traceback.format_exc(limit=8)
                with connect() as conn:
                    conn.execute("UPDATE generation_jobs SET status = ? WHERE id = ?", ("generation_failed", job_id))
                log_action(user_id, "auto_generate_pdf_packages_failed", f"{job_code}: {detail[-1800:]}")

        threading.Thread(target=worker, name=f"auto-pdf-generation-{job_code}", daemon=True).start()

    def handle_generate_p1(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "incorporation":
            return self.error_page(HTTPStatus.BAD_REQUEST, "当前第一阶段只支持注册 P1 文件包生成。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修正 Excel 后重新上传。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p1_pdf_package, "generating_pdf", "pdf_generated", "generate_p1_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_generate_p2_m01(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "maintenance":
            return self.error_page(HTTPStatus.BAD_REQUEST, "普通董事决议生成只支持公司维护/变更/年审业务单。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修正 Excel 后重新上传。")
        if suggestions.get("summary", {}).get("m01_available") != "Yes":
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有识别到可生成普通董事决议的事项。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p2_m01_pdf_package, "p2_m01_generating_pdf", "p2_m01_pdf_generated", "generate_p2_m01_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_generate_p2_m02(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在。")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "maintenance":
            return self.error_page(HTTPStatus.BAD_REQUEST, "转入包生成只支持公司维护/变更年审表。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修改 Excel 后重新上传。")
        if suggestions.get("summary", {}).get("m02_available") != "Yes":
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有识别到可生成转入包的事项。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p2_m02_pdf_package, "p2_m02_generating_pdf", "p2_m02_pdf_generated", "generate_p2_m02_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_generate_p2_m03(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在。")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "maintenance":
            return self.error_page(HTTPStatus.BAD_REQUEST, "股份转让包生成只支持公司维护/变更年审表。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修改 Excel 后重新上传。")
        if suggestions.get("summary", {}).get("m03_available") != "Yes":
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有识别到可生成股份转让包的事项。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p2_m03_pdf_package, "p2_m03_generating_pdf", "p2_m03_pdf_generated", "generate_p2_m03_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_generate_p2_m04(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在。")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "maintenance":
            return self.error_page(HTTPStatus.BAD_REQUEST, "增资配股包生成只支持公司维护/变更年审表。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修改 Excel 后重新上传。")
        if suggestions.get("summary", {}).get("m04_available") != "Yes":
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有识别到可生成增资配股包的事项。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p2_m04_pdf_package, "p2_m04_generating_pdf", "p2_m04_pdf_generated", "generate_p2_m04_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_generate_p2_m05(self, user):
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在。")
        parsed = json.loads(row["parsed_json"])
        suggestions = json.loads(row["suggestions_json"])
        errors = suggestions.get("summary", {}).get("blocking_errors", [])
        if row["task_type"] != "maintenance":
            return self.error_page(HTTPStatus.BAD_REQUEST, "年审包生成只支持公司维护/变更/年审业务单。")
        if errors:
            return self.error_page(HTTPStatus.BAD_REQUEST, "存在严重错误，请先修改 Excel 后重新上传。")
        if suggestions.get("summary", {}).get("m05_available") != "Yes":
            return self.error_page(HTTPStatus.BAD_REQUEST, "没有识别到可生成年审文件包的事项。")
        if is_generation_status(row["status"]):
            return self.redirect(f"/job?id={job_id}")
        self.queue_pdf_generation(user["id"], row, parsed, generate_p2_m05_pdf_package, "p2_m05_generating_pdf", "p2_m05_pdf_generated", "generate_p2_m05_pdf_package")
        self.redirect(f"/job?id={job_id}")

    def handle_save_common_person(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以保存常用人员。")
        files: dict[str, tuple[str, bytes]] = {}
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" in content_type.lower():
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            raw_fields, files = parse_multipart_form(content_type, body)
            fields = normalize_form_fields(raw_fields)
        else:
            fields = self.read_form_urlencoded()
        person_id = fields.get("id", [""])[0].strip()
        display_name = fields.get("display_name", [""])[0].strip()
        if not display_name:
            return self.error_page(HTTPStatus.BAD_REQUEST, "请填写显示名称。")

        values = {
            "display_name": display_name,
            "aliases": fields.get("aliases", [""])[0].strip(),
            "default_role": person_roles_from_fields(fields),
            "id_type": fields.get("id_type", [""])[0].strip(),
            "id_number": fields.get("id_number", [""])[0].strip(),
            "nationality": fields.get("nationality", [""])[0].strip(),
            "residential_address": fields.get("residential_address", [""])[0].strip(),
            "email": fields.get("email", [""])[0].strip(),
            "phone": fields.get("phone", [""])[0].strip(),
            "is_local_resident_director": 1 if (
                fields.get("role_local_director", [""])[0] == "1"
                or fields.get("role_nominee_director", [""])[0] == "1"
                or fields.get("is_local_resident_director", [""])[0] == "1"
            ) else 0,
            "active": 1 if fields.get("active", [""])[0] == "1" else 0,
            "notes": fields.get("notes", [""])[0].strip(),
            "signature_text": fields.get("signature_text", [""])[0].strip() or default_signature_text(display_name),
            "auto_signature_enabled": 1 if fields.get("auto_signature_enabled", [""])[0] == "1" else 0,
            "signature_image_path": fields.get("signature_image_path", [""])[0].strip(),
        }
        uploaded_signature = files.get("signature_file")
        if uploaded_signature and uploaded_signature[1]:
            try:
                values["signature_image_path"] = save_uploaded_signature_png(display_name, uploaded_signature[0], uploaded_signature[1])
            except ValueError as exc:
                return self.error_page(HTTPStatus.BAD_REQUEST, str(exc))
        elif values["auto_signature_enabled"] and (
            not values["signature_image_path"] or fields.get("regenerate_signature", [""])[0] == "1"
        ):
            values["signature_image_path"] = ensure_signature_image(values["signature_text"], display_name)
        with connect() as conn:
            if person_id:
                conn.execute(
                    """
                    UPDATE common_people
                    SET display_name = ?, aliases = ?, default_role = ?, id_type = ?, id_number = ?, nationality = ?,
                        residential_address = ?, email = ?, phone = ?, is_local_resident_director = ?,
                        active = ?, notes = ?, signature_text = ?, signature_image_path = ?,
                        auto_signature_enabled = ?
                    WHERE id = ?
                    """,
                    (
                        values["display_name"],
                        values["aliases"],
                        values["default_role"],
                        values["id_type"],
                        values["id_number"],
                        values["nationality"],
                        values["residential_address"],
                        values["email"],
                        values["phone"],
                        values["is_local_resident_director"],
                        values["active"],
                        values["notes"],
                        values["signature_text"],
                        values["signature_image_path"],
                        values["auto_signature_enabled"],
                        person_id,
                    ),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO common_people
                    (display_name, aliases, default_role, id_type, id_number, nationality, residential_address, email, phone,
                     is_local_resident_director, active, notes, signature_text, signature_image_path, auto_signature_enabled)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(display_name) DO UPDATE SET
                      aliases = excluded.aliases,
                      default_role = excluded.default_role,
                      id_type = excluded.id_type,
                      id_number = excluded.id_number,
                      nationality = excluded.nationality,
                      residential_address = excluded.residential_address,
                      email = excluded.email,
                      phone = excluded.phone,
                      is_local_resident_director = excluded.is_local_resident_director,
                      active = excluded.active,
                      notes = excluded.notes,
                      signature_text = excluded.signature_text,
                      signature_image_path = excluded.signature_image_path,
                      auto_signature_enabled = excluded.auto_signature_enabled
                    """,
                    (
                        values["display_name"],
                        values["aliases"],
                        values["default_role"],
                        values["id_type"],
                        values["id_number"],
                        values["nationality"],
                        values["residential_address"],
                        values["email"],
                        values["phone"],
                        values["is_local_resident_director"],
                        values["active"],
                        values["notes"],
                        values["signature_text"],
                        values["signature_image_path"],
                        values["auto_signature_enabled"],
                    ),
                )
        log_action(user["id"], "save_common_person", display_name)
        self.redirect("/settings#people")

    def handle_toggle_common_person(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以停用或启用常用人员。")
        fields = self.read_form_urlencoded()
        person_id = fields.get("id", [""])[0].strip()
        active = 1 if fields.get("active", [""])[0] == "1" else 0
        if not person_id:
            return self.error_page(HTTPStatus.BAD_REQUEST, "缺少人员 ID。")
        with connect() as conn:
            row = conn.execute("SELECT display_name FROM common_people WHERE id = ?", (person_id,)).fetchone()
            if not row:
                return self.error_page(HTTPStatus.NOT_FOUND, "常用人员不存在。")
            conn.execute("UPDATE common_people SET active = ? WHERE id = ?", (active, person_id))
        log_action(user["id"], "toggle_common_person", f"{row['display_name']}:{active}")
        self.redirect("/settings#people")

    def handle_delete_common_person(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以删除常用人员。")
        fields = self.read_form_urlencoded()
        person_id = fields.get("id", [""])[0].strip()
        if not person_id:
            return self.error_page(HTTPStatus.BAD_REQUEST, "缺少人员 ID。")
        with connect() as conn:
            row = conn.execute("SELECT display_name FROM common_people WHERE id = ?", (person_id,)).fetchone()
            if not row:
                return self.error_page(HTTPStatus.NOT_FOUND, "常用人员不存在。")
            conn.execute("DELETE FROM common_people WHERE id = ?", (person_id,))
        log_action(user["id"], "delete_common_person", row["display_name"])
        self.redirect("/settings#people")

    def handle_disable_sample_people(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以停用示例人员。")
        sample_names = ("挂名董事 A", "挂名董事 B", "公司秘书 A", "公司秘书 B")
        with connect() as conn:
            placeholders = ",".join("?" for _ in sample_names)
            conn.execute(
                f"UPDATE common_people SET active = 0 WHERE notes LIKE '示例%' OR display_name IN ({placeholders})",
                sample_names,
            )
        log_action(user["id"], "disable_sample_people", "common_people")
        self.redirect("/settings#people")

    def handle_save_user(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以保存用户。")
        fields = self.read_form_urlencoded()
        account_id = fields.get("id", [""])[0].strip()
        username = fields.get("username", [""])[0].strip()
        role = fields.get("role", ["staff"])[0].strip()
        password = fields.get("password", [""])[0]
        active = 1 if fields.get("active", [""])[0] == "1" else 0
        if not username:
            return self.error_page(HTTPStatus.BAD_REQUEST, "请填写用户名。")
        if role not in {"admin", "staff"}:
            return self.error_page(HTTPStatus.BAD_REQUEST, "角色只能是管理员或普通用户。")
        if account_id and str(user["id"]) == account_id and role != "admin":
            return self.error_page(HTTPStatus.BAD_REQUEST, "不能把当前登录的管理员账号改成普通用户。")
        if account_id and str(user["id"]) == account_id and not active:
            return self.error_page(HTTPStatus.BAD_REQUEST, "不能停用当前登录的管理员账号。")

        with connect() as conn:
            if account_id:
                existing = conn.execute("SELECT id FROM users WHERE id = ?", (account_id,)).fetchone()
                if not existing:
                    return self.error_page(HTTPStatus.NOT_FOUND, "用户不存在。")
                duplicate = conn.execute("SELECT id FROM users WHERE username = ? AND id <> ?", (username, account_id)).fetchone()
                if duplicate:
                    return self.error_page(HTTPStatus.BAD_REQUEST, "用户名已存在。")
                if password:
                    conn.execute(
                        "UPDATE users SET username = ?, role = ?, active = ?, password_hash = ? WHERE id = ?",
                        (username, role, active, hash_password(password), account_id),
                    )
                    conn.execute("DELETE FROM sessions WHERE user_id = ?", (account_id,))
                else:
                    conn.execute("UPDATE users SET username = ?, role = ?, active = ? WHERE id = ?", (username, role, active, account_id))
                    if not active:
                        conn.execute("DELETE FROM sessions WHERE user_id = ?", (account_id,))
            else:
                if not password:
                    return self.error_page(HTTPStatus.BAD_REQUEST, "新增用户必须填写密码。")
                existing = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
                if existing:
                    return self.error_page(HTTPStatus.BAD_REQUEST, "用户名已存在。")
                conn.execute(
                    "INSERT INTO users (username, password_hash, role, active, created_at) VALUES (?, ?, ?, ?, ?)",
                    (username, hash_password(password), role, active, now()),
                )
        log_action(user["id"], "save_user", f"{username}:{role}")
        self.redirect("/settings#users")

    def handle_toggle_user(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以停用或启用用户。")
        fields = self.read_form_urlencoded()
        account_id = fields.get("id", [""])[0].strip()
        active = 1 if fields.get("active", [""])[0] == "1" else 0
        if not account_id:
            return self.error_page(HTTPStatus.BAD_REQUEST, "缺少用户 ID。")
        if str(user["id"]) == account_id and not active:
            return self.error_page(HTTPStatus.BAD_REQUEST, "不能停用当前登录的管理员账号。")
        with connect() as conn:
            row = conn.execute("SELECT username FROM users WHERE id = ?", (account_id,)).fetchone()
            if not row:
                return self.error_page(HTTPStatus.NOT_FOUND, "用户不存在。")
            conn.execute("UPDATE users SET active = ? WHERE id = ?", (active, account_id))
            if not active:
                conn.execute("DELETE FROM sessions WHERE user_id = ?", (account_id,))
        log_action(user["id"], "toggle_user", f"{row['username']}:{active}")
        self.redirect("/settings#users")

    def handle_upload_document_template(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以上传正式文件模板。")
        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", "0"))
        if length > MAX_UPLOAD_BYTES:
            return self.error_page(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "文件太大")
        body = self.rfile.read(length)
        fields, files = parse_multipart_form(content_type, body)
        key = fields.get("template_key", "").strip()
        version = fields.get("version", "").strip()
        file_item = files.get("file")
        if key not in DOCUMENT_TEMPLATES:
            return self.error_page(HTTPStatus.BAD_REQUEST, "无法识别要更新的模板。")
        if not file_item:
            return self.error_page(HTTPStatus.BAD_REQUEST, "请上传 .docx 模板文件。")
        filename, file_bytes = file_item
        if not filename.lower().endswith(".docx"):
            return self.error_page(HTTPStatus.BAD_REQUEST, "正式文件模板请上传 .docx 文件。")

        target = DOCUMENT_TEMPLATES[key]["path"]
        resolved_target = target.resolve()
        if not resolved_target.is_relative_to(DOC_TEMPLATE_DIR.resolve()):
            return self.error_page(HTTPStatus.BAD_REQUEST, "模板路径不安全，已拒绝。")
        if not target.exists():
            return self.error_page(HTTPStatus.NOT_FOUND, "当前模板文件不存在，不能直接覆盖。")

        timestamp = datetime.now(UTC).strftime("%Y%m%d-%H%M%S")
        registry = load_template_registry()
        draft_dir = TEMPLATE_BACKUP_DIR / safe_filename(key) / "drafts"
        draft_dir.mkdir(parents=True, exist_ok=True)
        draft_path = draft_dir / f"{timestamp}_{safe_filename(filename)}"
        draft_path.write_bytes(file_bytes)

        try:
            preview_pdf = render_template_preview_pdf(key, draft_path, timestamp)
        except Exception as exc:
            draft_path.unlink(missing_ok=True)
            return self.error_page(HTTPStatus.BAD_REQUEST, f"新版模板无法转换为 PDF，没有启用。错误：{exc}")

        state = registry.get(key, {})
        registry[key] = {
            **state,
            "draft": {
                "version": version or f"draft-{timestamp}",
                "created_at": now(),
                "uploaded_filename": filename,
                "draft_path": str(draft_path),
                "preview_pdf": preview_pdf.name,
            },
        }
        save_template_registry(registry)
        log_action(user["id"], "upload_document_template_draft", f"{key}:{filename}")
        self.redirect("/settings#templates")

    def handle_activate_document_template_draft(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以启用模板草稿。")
        fields = self.read_form_urlencoded()
        key = fields.get("template_key", [""])[0].strip()
        if key not in DOCUMENT_TEMPLATES:
            return self.error_page(HTTPStatus.BAD_REQUEST, "无法识别要启用的模板。")
        registry = load_template_registry()
        state = registry.get(key, {})
        draft = state.get("draft") if isinstance(state.get("draft"), dict) else None
        if not draft:
            return self.error_page(HTTPStatus.BAD_REQUEST, "这个模板没有待启用草稿。")
        draft_path = Path(str(draft.get("draft_path") or ""))
        if not draft_path.exists():
            return self.error_page(HTTPStatus.NOT_FOUND, "草稿模板文件不存在。")
        target = DOCUMENT_TEMPLATES[key]["path"]
        timestamp = datetime.now(UTC).strftime("%Y%m%d-%H%M%S")
        backup_dir = TEMPLATE_BACKUP_DIR / safe_filename(key)
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"{timestamp}_{target.name}"
        shutil.copy2(target, backup_path)
        shutil.copy2(draft_path, target)
        history = list(state.get("history", []))
        history.append(
            {
                "version": state.get("version") or DOCUMENT_TEMPLATES[key]["version"],
                "backup_path": str(backup_path),
                "backed_up_at": now(),
                "note": "启用草稿前自动备份",
            }
        )
        registry[key] = {
            "version": draft.get("version") or f"custom-{timestamp}",
            "status": "启用",
            "updated_at": now(),
            "uploaded_filename": draft.get("uploaded_filename", ""),
            "backup_path": str(backup_path),
            "preview_pdf": draft.get("preview_pdf", ""),
            "history": history,
        }
        save_template_registry(registry)
        log_action(user["id"], "activate_document_template_draft", key)
        self.redirect("/settings#templates")

    def handle_rollback_document_template(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以回滚正式文件模板。")
        fields = self.read_form_urlencoded()
        key = fields.get("template_key", [""])[0].strip()
        backup_path_raw = fields.get("backup_path", [""])[0].strip()
        if key not in DOCUMENT_TEMPLATES or not backup_path_raw:
            return self.error_page(HTTPStatus.BAD_REQUEST, "缺少模板或历史版本。")
        target = DOCUMENT_TEMPLATES[key]["path"]
        backup_path = Path(backup_path_raw)
        try:
            if not backup_path.resolve().is_relative_to(TEMPLATE_BACKUP_DIR.resolve()):
                return self.error_page(HTTPStatus.BAD_REQUEST, "历史模板路径不安全，已拒绝。")
        except OSError:
            return self.error_page(HTTPStatus.BAD_REQUEST, "历史模板路径无效。")
        if not backup_path.exists() or not backup_path.is_file():
            return self.error_page(HTTPStatus.NOT_FOUND, "历史模板文件不存在。")

        timestamp = datetime.now(UTC).strftime("%Y%m%d-%H%M%S")
        backup_dir = TEMPLATE_BACKUP_DIR / safe_filename(key)
        backup_dir.mkdir(parents=True, exist_ok=True)
        current_backup = backup_dir / f"{timestamp}_before_rollback_{target.name}"
        shutil.copy2(target, current_backup)
        shutil.copy2(backup_path, target)
        try:
            preview_pdf = render_template_preview_pdf(key, target, timestamp)
        except Exception as exc:
            shutil.copy2(current_backup, target)
            return self.error_page(HTTPStatus.BAD_REQUEST, f"历史模板无法转换为 PDF，已恢复当前模板。错误：{exc}")

        registry = load_template_registry()
        state = registry.get(key, {})
        history = list(state.get("history", []))
        history.append(
            {
                "version": state.get("version") or DOCUMENT_TEMPLATES[key]["version"],
                "backup_path": str(current_backup),
                "backed_up_at": now(),
                "note": "回滚前自动备份",
            }
        )
        registry[key] = {
            **state,
            "version": f"rollback-{timestamp}",
            "status": "启用",
            "updated_at": now(),
            "backup_path": str(current_backup),
            "preview_pdf": preview_pdf.name,
            "history": history,
        }
        save_template_registry(registry)
        log_action(user["id"], "rollback_document_template", key)
        self.redirect("/settings#templates")

    def handle_delete_job(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以删除任务。")
        fields = self.read_form_urlencoded()
        job_id = fields.get("job_id", [""])[0]
        with connect() as conn:
            row = conn.execute("SELECT * FROM generation_jobs WHERE id = ?", (job_id,)).fetchone()
            if not row:
                return self.error_page(HTTPStatus.NOT_FOUND, "任务不存在")
            conn.execute("DELETE FROM generation_jobs WHERE id = ?", (job_id,))
        cleanup_job_files(row)
        log_action(user["id"], "delete_job", row["job_code"])
        self.redirect("/jobs")

    def handle_delete_selected_jobs(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以批量删除任务。")
        fields = self.read_form_urlencoded()
        raw_ids = fields.get("job_id", [])
        job_ids = [item for item in raw_ids if item.isdigit()]
        if not job_ids:
            return self.error_page(HTTPStatus.BAD_REQUEST, "请先勾选要删除的任务。")
        placeholders = ",".join("?" for _ in job_ids)
        with connect() as conn:
            rows = conn.execute(f"SELECT * FROM generation_jobs WHERE id IN ({placeholders})", job_ids).fetchall()
        if not rows:
            return self.error_page(HTTPStatus.NOT_FOUND, "没有找到可删除的任务。")
        for row in rows:
            cleanup_job_files(row)
        with connect() as conn:
            conn.execute(f"DELETE FROM generation_jobs WHERE id IN ({placeholders})", job_ids)
        log_action(user["id"], "delete_jobs", f"{len(rows)} tasks")
        self.redirect("/jobs")

    def handle_cleanup_old_jobs(self, user):
        if user.get("role") != "admin":
            return self.error_page(HTTPStatus.FORBIDDEN, "只有管理员可以清理旧任务。")
        fields = self.read_form_urlencoded()
        raw_days = fields.get("days", ["30"])[0]
        try:
            days = max(7, min(365, int(raw_days)))
        except ValueError:
            days = 30
        cutoff = (datetime.now(UTC) - timedelta(days=days)).replace(tzinfo=None).isoformat(timespec="seconds") + "Z"
        with connect() as conn:
            rows = conn.execute("SELECT * FROM generation_jobs WHERE created_at < ?", (cutoff,)).fetchall()
        for row in rows:
            cleanup_job_files(row)
        if rows:
            ids = [str(row["id"]) for row in rows]
            placeholders = ",".join("?" for _ in ids)
            with connect() as conn:
                conn.execute(f"DELETE FROM generation_jobs WHERE id IN ({placeholders})", ids)
        log_action(user["id"], "cleanup_old_jobs", f"{len(rows)} tasks older than {days} days")
        self.redirect("/jobs")

    def current_user(self):
        token = self.cookie_value(SESSION_COOKIE)
        if not token:
            return None
        with connect() as conn:
            row = conn.execute(
                """
                SELECT users.id, users.username, users.role
                FROM sessions JOIN users ON users.id = sessions.user_id
                WHERE sessions.token = ? AND sessions.expires_at > ? AND users.active = 1
                """,
                (token, now()),
            ).fetchone()
        return dict(row) if row else None

    def require_user(self):
        user = self.current_user()
        if not user:
            self.redirect("/login")
            return None
        return user

    def clear_session(self):
        token = self.cookie_value(SESSION_COOKIE)
        if token:
            with connect() as conn:
                conn.execute("DELETE FROM sessions WHERE token = ?", (token,))
        self.send_response(302)
        self.send_header("Location", "/login")
        self.send_header("Set-Cookie", f"{SESSION_COOKIE}=; Max-Age=0; Path=/")
        self.end_headers()

    def cookie_value(self, key: str) -> str:
        raw = self.headers.get("Cookie", "")
        for part in raw.split(";"):
            if "=" in part:
                k, v = part.strip().split("=", 1)
                if k == key:
                    return v
        return ""

    def read_form_urlencoded(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8")
        return parse_qs(raw)

    def send_html(self, data: bytes, status=HTTPStatus.OK):
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def serve_file(self, path: Path, content_type: str, inline: bool = False):
        if not path.exists() or not path.is_file():
            return self.error_page(HTTPStatus.NOT_FOUND, "文件不存在")
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        disposition = "inline" if inline else "attachment"
        self.send_header("Content-Disposition", f"{disposition}; filename*=UTF-8''{quote(path.name)}")
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, location: str):
        self.send_response(302)
        self.send_header("Location", location)
        self.end_headers()

    def error_page(self, status: HTTPStatus, message: str):
        body = f"<section class='panel'><h2>{status.value}</h2><p>{h(message)}</p><p><a href='/'>返回首页</a></p></section>"
        self.send_html(render_page(status.phrase, body, self.current_user()), status)


def singapore_date_input() -> str:
    return datetime.now(DISPLAY_TIMEZONE).strftime("%Y-%m-%d")


def active_common_people_map() -> dict[str, dict[str, object]]:
    with connect() as conn:
        common_rows = conn.execute("SELECT * FROM common_people WHERE active = 1").fetchall()
    return {
        r["display_name"]: {
            "id": r["id"],
            "display_name": r["display_name"],
            "full_name": r["display_name"],
            "aliases": r["aliases"] if "aliases" in r.keys() else "",
            "default_role": r["default_role"],
            "id_type": r["id_type"],
            "id_number": r["id_number"],
            "nationality": r["nationality"],
            "residential_address": r["residential_address"],
            "email": r["email"],
            "phone": r["phone"],
            "signature_text": r["signature_text"] if "signature_text" in r.keys() else "",
            "signature_image_path": r["signature_image_path"] if "signature_image_path" in r.keys() else "",
            "auto_signature_enabled": r["auto_signature_enabled"] if "auto_signature_enabled" in r.keys() else 0,
            "is_local_resident_director": bool(r["is_local_resident_director"]),
        }
        for r in common_rows
    }


def common_people_options(common_people: dict[str, dict[str, object]]) -> str:
    values: list[str] = []
    for person in common_people.values():
        display_name = str(person.get("display_name") or "").strip()
        if display_name:
            values.append(display_name)
        aliases = str(person.get("aliases") or "")
        values.extend(part.strip() for part in re.split(r"[,;/\n]+", aliases) if part.strip())
    return "".join(f'<option value="{h(value)}"></option>' for value in dedupe_strings(values))


def form_value(fields: dict[str, list[str]], key: str, default: str = "") -> str:
    return (fields.get(key, [default])[0] or "").strip()


def form_checked(fields: dict[str, list[str]], key: str) -> bool:
    return form_value(fields, key).lower() in {"1", "yes", "true", "on"}


def split_form_names(value: object) -> list[str]:
    return [part.strip() for part in re.split(r"[,;/\n]+", str(value or "")) if part.strip()]


def normalized_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def common_name_lookup(common_people: dict[str, dict[str, object]]) -> set[str]:
    keys: set[str] = set()
    for person in common_people.values():
        keys.add(normalized_key(person.get("display_name")))
        for alias in re.split(r"[,;/\n]+", str(person.get("aliases") or "")):
            keys.add(normalized_key(alias))
    return {key for key in keys if key}


def dedupe_strings(values: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value or "").strip()
        key = text.lower()
        if text and key not in seen:
            seen.add(key)
            out.append(text)
    return out


def build_p2_form_parsed(fields: dict[str, list[str]], common_people: dict[str, dict[str, object]]) -> dict[str, object]:
    document_date = form_value(fields, "default_document_date") or singapore_date_input()
    current_address = form_value(fields, "registered_office_address")
    new_address = form_value(fields, "new_registered_office_address")
    company = {
        "task_type": "maintenance",
        "template_version": "p2_web_form_v1",
        "source_type": "WebForm",
        "company_name": form_value(fields, "company_name"),
        "uen": form_value(fields, "uen"),
        "registered_office_address": current_address,
        "default_document_date": document_date,
        "currency": form_value(fields, "currency", "SGD") or "SGD",
        "issued_share_capital": form_value(fields, "issued_share_capital"),
        "paid_up_capital": form_value(fields, "paid_up_capital"),
        "director_signer_names": form_value(fields, "director_signer_names"),
        "member_signer_names": form_value(fields, "member_signer_names"),
        "client_signatory_name": form_value(fields, "member_signer_names"),
        "new_secretary_company": form_value(fields, "new_secretary_company") or "RSIN GROUP PTE. LTD.",
        "old_secretary_company": form_value(fields, "old_secretary_company"),
        "new_registered_office_address": new_address,
        "egm_meeting_place": new_address,
    }
    lookup = common_name_lookup(common_people)
    people: list[dict[str, object]] = []
    person_index = 1

    def add_person(name: str, **roles: str) -> str:
        nonlocal person_index
        name = name.strip()
        if not name:
            return ""
        for person in people:
            if str(person.get("_input_name", "")).lower() == name.lower():
                for key, value in roles.items():
                    if value:
                        person[key] = value
                return str(person["person_id"])
        person_id = f"WEBP{person_index:03d}"
        person_index += 1
        is_common = normalized_key(name) in lookup
        item: dict[str, object] = {
            "person_id": person_id,
            "_input_name": name,
            "source": "common" if is_common else "new",
            "common_person_name": name if is_common else "",
            "full_name": "" if is_common else name,
            "is_director": "No",
            "is_secretary": "No",
            "is_shareholder": "No",
            "is_client_signatory": "No",
            "is_local_resident_director": "Auto",
            "is_nominee_director": "Auto",
        }
        item.update({key: value for key, value in roles.items() if value})
        people.append(item)
        return person_id

    for idx, name in enumerate(split_form_names(company["director_signer_names"])):
        add_person(name, is_director="Yes", is_client_signatory="Yes" if idx == 0 else "No")
    for name in split_form_names(company["member_signer_names"]):
        add_person(name, is_shareholder="Yes", is_client_signatory="Yes")

    events: list[dict[str, object]] = []

    def event(event_type: str, **extra: object) -> dict[str, object]:
        row = {
            "event_id": f"EV{len(events) + 1:03d}",
            "event_type": event_type,
            "event_name_cn": P2_EVENT_CN.get(event_type, event_type),
            "generate": "Yes",
            "effective_date": extra.pop("effective_date", document_date),
            "approval_route": extra.pop("approval_route", "DR"),
            "document_group": extra.pop("document_group", "DR-001"),
            "combine_in_dr": extra.pop("combine_in_dr", "Yes"),
            "resignation_letter": extra.pop("resignation_letter", "No"),
            "manual_review_required": extra.pop("manual_review_required", "No"),
        }
        row.update(extra)
        events.append(row)
        return row

    if form_checked(fields, "change_registered_office") or new_address:
        event(
            "change_registered_office",
            old_registered_office_address=current_address,
            new_registered_office_address=new_address,
            old_value=current_address,
            new_value=new_address,
        )
    if form_checked(fields, "change_business_activity") or any(
        form_value(fields, key)
        for key in ["primary_ssic_new", "primary_activity_new", "secondary_ssic_new", "secondary_activity_new"]
    ):
        event(
            "change_business_activity",
            primary_ssic_new=form_value(fields, "primary_ssic_new"),
            primary_activity_new=form_value(fields, "primary_activity_new"),
            secondary_ssic_new=form_value(fields, "secondary_ssic_new"),
            secondary_activity_new=form_value(fields, "secondary_activity_new"),
            new_value=form_value(fields, "primary_activity_new"),
        )
    if form_checked(fields, "change_fye") or form_value(fields, "new_fye"):
        event("change_fye", new_value=form_value(fields, "new_fye"))

    resignation_letter = "Yes" if form_checked(fields, "resignation_letter") or form_checked(fields, "transfer_in_resignation_letter") else "No"
    for name in split_form_names(form_value(fields, "resign_director_names")):
        person_id = add_person(name, is_director="Yes")
        event("resign_director", target_person_id=person_id, target_name=name, resignation_letter=resignation_letter)
    for name in split_form_names(form_value(fields, "appoint_director_names")) + split_form_names(form_value(fields, "transfer_in_nominee_director_names")):
        person_id = add_person(name, is_director="Yes", is_local_resident_director="Auto", is_nominee_director="Auto")
        event("appoint_director", target_person_id=person_id, target_name=name)
    for name in split_form_names(form_value(fields, "resign_secretary_names")):
        person_id = add_person(name, is_secretary="Yes")
        event("resign_secretary", target_person_id=person_id, target_name=name, resignation_letter=resignation_letter)
    for name in split_form_names(form_value(fields, "appoint_secretary_names")) + split_form_names(form_value(fields, "transfer_in_secretary_names")):
        person_id = add_person(name, is_secretary="Yes")
        event("appoint_secretary", target_person_id=person_id, target_name=name)

    if form_checked(fields, "transfer_in_required"):
        event(
            "transfer_in",
            approval_route="EGM+DR",
            document_group="TAKEOVER-001",
            combine_in_dr="No",
            old_secretary_company=company["old_secretary_company"],
            new_secretary_company=company["new_secretary_company"],
            old_value=company["old_secretary_company"],
            new_value=company["new_secretary_company"],
            resignation_letter=resignation_letter,
        )

    particular_name = form_value(fields, "particular_person_name")
    particular_field = form_value(fields, "particular_field_label")
    particular_new = form_value(fields, "particular_new_value")
    if particular_name and particular_field and particular_new:
        person_id = add_person(particular_name)
        event(
            "update_officer_particulars",
            target_person_id=person_id,
            target_name=particular_name,
            field_label=particular_field,
            old_value=form_value(fields, "particular_old_value"),
            new_value=particular_new,
        )

    share_transfers: list[dict[str, object]] = []
    if form_checked(fields, "share_transfer_required") or all(
        form_value(fields, key) for key in ["transferor_name", "transferee_name", "shares_transferred"]
    ):
        share_transfers.append(
            {
                "generate": "Yes",
                "transferor_name": form_value(fields, "transferor_name"),
                "transferee_name": form_value(fields, "transferee_name"),
                "share_class": form_value(fields, "transfer_share_class", "Ordinary") or "Ordinary",
                "shares_transferred": form_value(fields, "shares_transferred"),
                "transfer_date": form_value(fields, "transfer_date") or document_date,
                "consideration_amount": form_value(fields, "consideration_amount"),
                "consideration_basis": form_value(fields, "consideration_basis") or "internal_paid_up_basis",
                "stamp_duty_review": "Yes" if form_value(fields, "consideration_basis") == "stamp_duty_higher_of_price_or_nav" else "No",
            }
        )

    share_allotments: list[dict[str, object]] = []
    if form_checked(fields, "share_allotment_required") or all(
        form_value(fields, key) for key in ["allottee_name", "shares_allotted"]
    ):
        share_allotments.append(
            {
                "generate": "Yes",
                "allottee_name": form_value(fields, "allottee_name"),
                "share_class": form_value(fields, "allotment_share_class", "Ordinary") or "Ordinary",
                "shares_allotted": form_value(fields, "shares_allotted"),
                "issued_share_capital": form_value(fields, "allotment_issued_share_capital"),
                "paid_up_share_capital": form_value(fields, "allotment_paid_up_share_capital"),
                "total_paid": form_value(fields, "allotment_paid_up_share_capital"),
                "allotment_date": form_value(fields, "allotment_date") or document_date,
                "authority_date": document_date,
                "form24_required": "Yes",
                "generate_certificate": "Yes",
            }
        )

    annual_enabled = form_checked(fields, "annual_review_required") or bool(form_value(fields, "fye_date") or form_value(fields, "agm_date"))
    annual_review = {
        "annual_review_required": "Yes" if annual_enabled else "No",
        "fye_date": form_value(fields, "fye_date"),
        "agm_date": form_value(fields, "agm_date") or document_date,
        "agm_time": "10.00 a.m.",
        "agm_place": new_address or current_address,
        "agm_route": form_value(fields, "agm_route") or "ordinary_agm",
        "accounts_status": form_value(fields, "accounts_status") or "non_dormant",
        "director_signer_name": company["director_signer_names"],
        "shareholder_signer_name": company["member_signer_names"],
        "ar_authorized_signer_name": company["director_signer_names"],
        "directors_fee": form_value(fields, "directors_fee", "0") or "0",
        "directors_remuneration": form_value(fields, "directors_remuneration", "0") or "0",
        "management_rep_letter": "Yes",
    }
    resolved_people = resolve_people(people, common_people)
    for person in resolved_people:
        if not person.get("full_name"):
            person["full_name"] = person.get("common_person_name") or person.get("_input_name") or ""

    output_options = {
        "package": "P2_WEB_FORM",
        "output_format": "pdf_zip",
        "notes": "Created from P2 quick web form.",
    }
    return {
        "task_type": "maintenance",
        "company": company,
        "people": resolved_people,
        "shareholders": [],
        "shareholdings": [],
        "generation": output_options,
        "output_options": output_options,
        "changes": events,
        "change_events": events,
        "personal_changes": [],
        "share_changes": share_transfers,
        "share_transfers": share_transfers,
        "share_allotments": share_allotments,
        "annual_review": annual_review,
    }


P2_EVENT_CN = {
    "change_registered_office": "注册地址变更",
    "change_business_activity": "营业范围/SSIC 变更",
    "change_fye": "财年日变更",
    "appoint_director": "委任董事",
    "resign_director": "董事辞任",
    "appoint_secretary": "委任秘书",
    "resign_secretary": "秘书辞任",
    "transfer_in": "秘书公司转入",
    "update_officer_particulars": "人员资料变更",
}


def alert_block(title: str, items: list[str], class_name: str) -> str:
    if not items:
        return ""
    lis = "".join(f"<li>{h(item)}</li>" for item in items)
    return f"<div class='{class_name}'><strong>{h(title)}</strong><ul>{lis}</ul></div>"


def auto_generation_tasks(parsed: dict, suggestions: dict) -> list[tuple[str, object]]:
    summary = suggestions.get("summary", {}) if isinstance(suggestions, dict) else {}
    if summary.get("blocking_errors"):
        return []
    task_type = parsed.get("task_type")
    if task_type == "incorporation":
        return [("P1", generate_p1_pdf_package)]
    if task_type != "maintenance":
        return []
    options = [
        ("M01", "m01_available", generate_p2_m01_pdf_package),
        ("M02", "m02_available", generate_p2_m02_pdf_package),
        ("M03", "m03_available", generate_p2_m03_pdf_package),
        ("M04", "m04_available", generate_p2_m04_pdf_package),
        ("M05", "m05_available", generate_p2_m05_pdf_package),
    ]
    return [(code, generator) for code, key, generator in options if summary.get(key) == "Yes"]


def dev_code(code: str, user: dict | None) -> str:
    if not user or user.get("role") != "admin":
        return ""
    return f"<span class='dev-code'>内部编号 {h(code)}</span>"


def case_metadata_from_parsed(parsed: dict[str, object], source_filename: str, job_code: str) -> dict[str, str]:
    company = parsed.get("company", {}) if isinstance(parsed.get("company"), dict) else {}
    generation = parsed.get("generation", {}) if isinstance(parsed.get("generation"), dict) else {}
    output_options = parsed.get("output_options", {}) if isinstance(parsed.get("output_options"), dict) else {}
    sources = [company, generation, output_options]

    def first_value(*keys: str) -> str:
        for source in sources:
            for key in keys:
                value = source.get(key)
                if value not in (None, ""):
                    return str(value).strip()
        return ""

    business_order_id = first_value("business_order_id", "case_id", "order_id")
    return {
        "case_id": business_order_id or job_code,
        "business_order_id": business_order_id,
        "source_type": first_value("source_type") or "Excel",
        "source_file_id": first_value("source_file_id") or source_filename,
        "contact_person_id": first_value("contact_person_id"),
        "agent_person_id": first_value("agent_person_id"),
        "client_signatory_person_id": first_value("client_signatory_person_id", "client_signer_person_id"),
        "authorized_representative_person_id": first_value("authorized_representative_person_id", "authorized_rep_person_id"),
        "prepared_by": first_value("prepared_by"),
        "snapshot_version": "parsed-v1",
    }


def case_metadata_panel(row, user: dict | None) -> str:
    fields = [
        ("业务单编号", row["case_id"] or row["job_code"]),
        ("数据来源", row["source_type"] or "Excel"),
        ("来源文件编号", row["source_file_id"] or row["source_filename"]),
        ("联系人", row["contact_person_id"] or "-"),
        ("代理/代办人", row["agent_person_id"] or "-"),
        ("客户方签字人", row["client_signatory_person_id"] or "-"),
        ("授权代表", row["authorized_representative_person_id"] or "-"),
        ("经办人", row["prepared_by"] or "-"),
    ]
    rows = "".join(f"<div><span>{h(label)}</span><strong>{h(value)}</strong></div>" for label, value in fields)
    admin_note = ""
    if user and user.get("role") == "admin":
        admin_note = f"<p class='muted'>数据快照：本次上传解析后的完整资料已保存为任务快照，版本 {h(row['snapshot_version'] or 'parsed-v1')}。以后人员主档变化，不会自动改动这次已生成文件的历史依据。</p>"
    return f"""
    <div class="case-meta">
      <div class="case-meta-grid">{rows}</div>
      {admin_note}
    </div>
    """


def public_template_cards(keys: list[str]) -> str:
    cards = []
    for key in keys:
        item = TEMPLATE_DOWNLOADS.get(key)
        if not item:
            continue
        label, path = item
        if not path.exists():
            continue
        title, note = TEMPLATE_NOTES.get(key, (label, ""))
        cards.append(
            f"""
            <a class="download-card" href="/templates/{h(key)}">
              <strong>{h(title)}</strong>
              <span>{h(note)}</span>
            </a>
            """
        )
    return "".join(cards)


def admin_overview_html(users, people, rules, job_stats) -> str:
    active_users = sum(1 for row in users if row["active"])
    staff_users = sum(1 for row in users if row["role"] == "staff")
    active_people = sum(1 for row in people if row["active"])
    sample_people = sum(1 for row in people if str(row["notes"] or "").startswith("示例"))
    active_imports = sum(1 for key in ACTIVE_IMPORT_TEMPLATE_KEYS if TEMPLATE_DOWNLOADS.get(key, ("", Path()))[1].exists())
    missing_imports = sum(1 for key in ACTIVE_IMPORT_TEMPLATE_KEYS if not TEMPLATE_DOWNLOADS.get(key, ("", Path()))[1].exists())
    registry = load_template_registry()
    draft_templates = sum(
        1
        for state in registry.values()
        if isinstance(state, dict) and isinstance(state.get("draft"), dict)
    )
    missing_doc_templates = sum(1 for meta in DOCUMENT_TEMPLATES.values() if not meta["path"].exists())
    total_jobs = int(job_stats["total_jobs"] or 0) if job_stats else 0
    blocked_jobs = int(job_stats["blocked_jobs"] or 0) if job_stats else 0
    pdf_jobs = int(job_stats["pdf_jobs"] or 0) if job_stats else 0
    cards = [
        ("用户账户", f"{active_users}/{len(users)} 启用", f"普通用户 {staff_users} 个"),
        ("常用人员", f"{active_people}/{len(people)} 启用", f"示例资料 {sample_people} 个"),
        ("导入表", f"{active_imports} 个正式入口", f"缺失 {missing_imports} 个"),
        ("正式模板", f"{len(DOCUMENT_TEMPLATES) - missing_doc_templates}/{len(DOCUMENT_TEMPLATES)} 可用", f"待启用草稿 {draft_templates} 个"),
        ("文件规则", f"{len(rules)} 条", "用于判断文件包和签字逻辑"),
        ("生成任务", f"{total_jobs} 个历史任务", f"已生成 PDF {pdf_jobs} 个；需修正 {blocked_jobs} 个"),
    ]
    card_html = "".join(
        f"""
        <div class="admin-kpi">
          <span>{h(label)}</span>
          <strong>{h(value)}</strong>
          <small>{h(note)}</small>
        </div>
        """
        for label, value, note in cards
    )
    return f"""
    <section class="panel" id="overview">
      <details class="admin-overview-drawer">
        <summary>
          <span>后台总览</span>
          <small>账号、常用人员、模板、规则和任务状态</small>
        </summary>
        <div class="section-header compact">
          <div>
            <h3>后台总览</h3>
            <p class="muted">先看账号、常用人员、模板和任务状态；需要维护时再进入下面对应模块。</p>
          </div>
          <span class="badge">维护面板</span>
        </div>
        <div class="admin-kpi-grid">{card_html}</div>
      </details>
    </section>
    """


def template_library_rows_html(show_archived: bool = False) -> str:
    keys = [key for key in TEMPLATE_DOWNLOADS if key in ACTIVE_IMPORT_TEMPLATE_KEYS]
    if show_archived:
        keys.extend(key for key in TEMPLATE_DOWNLOADS if key not in ACTIVE_IMPORT_TEMPLATE_KEYS)
    rows = [
        template_library_row(key, *TEMPLATE_DOWNLOADS[key], archived=key not in ACTIVE_IMPORT_TEMPLATE_KEYS)
        for key in keys
    ]
    return "".join(rows) or '<tr><td colspan="6">暂无导入表</td></tr>'


def template_library_row(key: str, label: str, path: Path, archived: bool = False) -> str:
    title, note = TEMPLATE_NOTES.get(key, (label, ""))
    status = "已找到" if path.exists() else "缺失"
    if archived:
        status = f"{status} / 默认隐藏"
    link = f"<a href='/templates/{h(key)}'>下载</a>" if path.exists() else ""
    visibility = "普通用户可见" if key in PUBLIC_TEMPLATE_KEYS else "仅管理员"
    row_class = " class='archived-row'" if archived else ""
    return f"<tr{row_class}><td>{h(key)}</td><td>{h(title)}</td><td>{h(note)}</td><td>{h(visibility)}</td><td>{h(status)}</td><td>{link}</td></tr>"


def document_template_sections_html() -> str:
    registry = load_template_registry()
    grouped: dict[str, list[str]] = {}
    for key, meta in DOCUMENT_TEMPLATES.items():
        category = str(meta["category"])
        grouped.setdefault(category, []).append(document_template_row_html(key, meta, registry))
    sections = []
    for category, rows in grouped.items():
        sections.append(
            f"""
            <div class="template-group">
              <div class="template-group-title">
                <strong>{h(category)}</strong>
                <span>{len(rows)} 个模板</span>
              </div>
              <div class="table-wrap">
                <table>
                  <thead><tr><th>文件类型</th><th>内部键</th><th>当前版本</th><th>状态</th><th>说明</th><th>操作</th></tr></thead>
                  <tbody>{''.join(rows)}</tbody>
                </table>
              </div>
            </div>
            """
        )
    return "".join(sections) or "<p class='muted'>暂无正式文件模板</p>"


def document_template_rows_html() -> str:
    registry = load_template_registry()
    rows = []
    for key, meta in DOCUMENT_TEMPLATES.items():
        rows.append(document_template_row_html(key, meta, registry, include_category=True))
    return "".join(rows) or '<tr><td colspan="6">暂无正式文件模板</td></tr>'


def document_template_row_html(key: str, meta: dict[str, object], registry: dict[str, dict[str, object]], include_category: bool = False) -> str:
    state = registry.get(key, {})
    draft = state.get("draft") if isinstance(state.get("draft"), dict) else None
    path = meta["path"]
    version = state.get("version") or meta["version"]
    status = state.get("status") or meta["status"]
    updated = state.get("updated_at")
    status_text = f"{status}；{updated}" if updated else status
    status_kind = "ok"
    if draft:
        status_text = f"{status_text}；有草稿待启用"
        status_kind = "draft"
    if not path.exists():
        status_text = "缺失"
        status_kind = "danger"
    download = f"<a href='/document-template/{h(key)}'>下载模板</a>" if path.exists() else ""
    preview = ""
    preview_name = state.get("preview_pdf", "")
    if preview_name and (GENERATED_DIR / Path(preview_name).name).exists():
        preview = f"<a href='/generated/{h(Path(preview_name).name)}'>预览 PDF</a>"
    draft_html = template_draft_panel(key, draft)
    history_form = template_history_form(key, state)
    upload_form = f"""
    <form method="post" action="/settings/template/upload" enctype="multipart/form-data" class="inline-upload">
      <input type="hidden" name="template_key" value="{h(key)}">
      <input name="version" placeholder="新版号，如 v1.4">
      <input type="file" name="file" accept=".docx" required>
      <button type="submit">上传新版草稿</button>
    </form>
    """
    category_cell = f"<td>{h(meta['category'])}</td>" if include_category else ""
    return (
        f"<tr><td>{h(meta['name'])}</td>{category_cell}<td><span class='dev-code'>{h(key)}</span></td>"
        f"<td>{h(version)}</td><td><span class='status-pill {status_kind}'>{h(status_text)}</span></td>"
        f"<td>{h(meta['note'])}</td><td><div class='template-actions'>{download}{preview}{draft_html}{upload_form}{history_form}</div></td></tr>"
    )


def template_history_form(key: str, state: dict[str, object]) -> str:
    history = state.get("history", [])
    if not isinstance(history, list) or not history:
        return "<span class='muted'>暂无历史版本</span>"
    options = []
    for item in reversed(history[-8:]):
        if not isinstance(item, dict):
            continue
        backup_path = str(item.get("backup_path") or "")
        label = f"{item.get('version') or '历史版本'} · {item.get('backed_up_at') or ''}"
        if backup_path and Path(backup_path).exists():
            options.append(f"<option value='{h(backup_path)}'>{h(label)}</option>")
    if not options:
        return "<span class='muted'>历史文件缺失</span>"
    return f"""
    <form method="post" action="/settings/template/rollback" class="inline-upload">
      <input type="hidden" name="template_key" value="{h(key)}">
      <select name="backup_path">{''.join(options)}</select>
      <button class="secondary-action" type="submit">回滚启用</button>
    </form>
    """


def template_draft_panel(key: str, draft: dict[str, object] | None) -> str:
    if not draft:
        return ""
    preview_name = Path(str(draft.get("preview_pdf") or "")).name
    preview = f"<a href='/generated/{h(preview_name)}'>查看草稿 PDF</a>" if preview_name and (GENERATED_DIR / preview_name).exists() else "<span class='muted'>草稿 PDF 缺失</span>"
    return f"""
    <div class="draft-box">
      <strong>有待启用草稿</strong>
      <span>{h(draft.get('version') or '')} · {h(display_time(draft.get('created_at')))}</span>
      <div class="row-actions">
        {preview}
        <form method="post" action="/settings/template/activate-draft" class="inline-form">
          <input type="hidden" name="template_key" value="{h(key)}">
          <button type="submit">确认启用草稿</button>
        </form>
      </div>
    </div>
    """


def render_template_preview_pdf(key: str, target: Path, timestamp: str) -> Path:
    preview_work_dir = GENERATED_DIR / "_template_preview_work"
    pdf_path = convert_docx_to_pdf(target, preview_work_dir)
    preview_name = f"template_preview_{safe_filename(key)}_{timestamp}.pdf"
    preview_path = GENERATED_DIR / preview_name
    if preview_path.exists():
        preview_path.unlink()
    shutil.copy2(pdf_path, preview_path)
    return preview_path


def row_value(row, key: str, default: str = "") -> str:
    if not row:
        return default
    return str(row[key] if row[key] is not None else default)


PERSON_ROLE_OPTIONS = [
    ("role_secretary", "Secretary", "公司秘书"),
    ("role_nominee_director", "Nominee Director", "挂名董事"),
    ("role_local_director", "Local Resident Director", "本地董事"),
    ("role_client_director", "Client Director", "客户董事"),
    ("role_client_signatory", "Client Signatory", "客户签字人"),
]


def split_person_roles(raw: object) -> list[str]:
    text_value = str(raw or "")
    roles = []
    for part in text_value.replace("；", ",").replace("，", ",").split(","):
        role = part.strip()
        if role:
            roles.append(role)
    return roles


def person_roles_from_fields(fields: dict[str, list[str]]) -> str:
    roles = [role for field, role, _label in PERSON_ROLE_OPTIONS if fields.get(field, [""])[0] == "1"]
    legacy_role = fields.get("default_role", [""])[0].strip()
    if legacy_role:
        roles.extend(split_person_roles(legacy_role))
    deduped = []
    for role in roles:
        if role not in deduped:
            deduped.append(role)
    return ", ".join(deduped)


def role_checkboxes(selected_roles: list[str]) -> str:
    return "".join(
        f"<label class='check'><input type='checkbox' name='{field}' value='1' {'checked' if role in selected_roles else ''}> {label}</label>"
        for field, role, label in PERSON_ROLE_OPTIONS
    )


def role_badges(raw_roles: object) -> str:
    roles = split_person_roles(raw_roles)
    if not roles:
        return "-"
    labels = {role: label for _field, role, label in PERSON_ROLE_OPTIONS}
    return "".join(f"<span class='role-badge'>{h(labels.get(role, role))}</span>" for role in roles)


def checked_attr(value: object) -> str:
    return "checked" if value else ""


def role_options(selected: str) -> str:
    return "".join(
        f"<option value='{value}' {'selected' if selected == value else ''}>{label}</option>"
        for value, label in [("staff", "普通用户"), ("admin", "管理员")]
    )


def user_form_html(edit_user) -> str:
    editing = bool(edit_user)
    active = True if not edit_user else bool(edit_user["active"])
    return f"""
    <form method="post" action="/settings/users/save" class="admin-form">
      <input type="hidden" name="id" value="{h(row_value(edit_user, 'id'))}">
      <label>用户名<input name="username" value="{h(row_value(edit_user, 'username'))}" required placeholder="例如 staff01"></label>
      <label>角色<select name="role">{role_options(row_value(edit_user, 'role', 'staff'))}</select></label>
      <label>密码<input name="password" type="password" placeholder="新增必填；编辑时填入即重置密码"></label>
      <label class="check"><input type="checkbox" name="active" value="1" {checked_attr(active)}> 启用</label>
      <button type="submit">{'保存修改' if editing else '新增用户'}</button>
      {'<a class="button-link secondary" href="/settings">取消编辑</a>' if editing else ''}
    </form>
    """


def user_table_row(row, current_user: dict) -> str:
    active = bool(row["active"])
    toggle_label = "停用" if active else "启用"
    toggle_value = "0" if active else "1"
    status = "<span class='status-pill ok'>启用</span>" if active else "<span class='status-pill muted'>停用</span>"
    disable_self = str(row["id"]) == str(current_user["id"]) and active
    toggle_form = (
        "<span class='muted'>当前账号</span>"
        if disable_self
        else f"""
        <form method="post" action="/settings/users/toggle" class="inline-form">
          <input type="hidden" name="id" value="{h(row['id'])}">
          <input type="hidden" name="active" value="{toggle_value}">
          <button class="secondary-action" type="submit">{toggle_label}</button>
        </form>
        """
    )
    return f"""
    <tr>
      <td>{h(row['username'])}</td>
      <td>{h(role_label(row['role']))}</td>
      <td>{status}</td>
      <td>{h(display_time(row['last_login_at']))}</td>
      <td>{h(display_time(row['created_at']))}</td>
      <td><div class="row-actions"><a href="/settings?edit_user={h(row['id'])}#users">编辑</a>{toggle_form}</div></td>
    </tr>
    """


def login_log_row(row) -> str:
    actor = row["username"] or row["detail"] or "-"
    action = {
        "login": "登录成功",
        "login_failed": "登录失败",
        "save_user": "保存用户",
        "toggle_user": "启用/停用用户",
        "save_common_person": "保存常用人员",
        "toggle_common_person": "启用/停用常用人员",
        "delete_common_person": "删除常用人员",
        "disable_sample_people": "停用示例人员",
        "upload_document_template_draft": "上传模板草稿",
        "activate_document_template_draft": "启用模板草稿",
        "rollback_document_template": "回滚模板",
        "download_generated": "下载生成文件",
        "download_import_template": "下载导入表",
        "download_document_template": "下载文件模板",
        "delete_job": "删除任务",
        "delete_jobs": "批量删除任务",
        "cleanup_old_jobs": "清理旧任务",
    }.get(row["action"], row["action"])
    return f"<tr><td>{h(display_time(row['created_at']))}</td><td>{h(actor)}</td><td>{h(action)}</td><td>{h(row['detail'])}</td></tr>"


def job_table_row(row, is_admin: bool) -> str:
    label = (row["case_id"] or row["job_code"]) if is_admin else (row["company_name"] or row["source_filename"])
    checkbox = f"<td><input type='checkbox' name='job_id' value='{h(row['id'])}'></td>" if is_admin else ""
    return f"""
    <tr>
      {checkbox}
      <td><a href="/job?id={h(row['id'])}">{h(label)}</a></td>
      <td>{h(task_label(row['task_type']))}</td>
      <td>{h(row['company_name'])}</td>
      <td>{h(row['source_filename'])}</td>
      <td>{h(status_label(row['status']))}</td>
      <td>{h(display_time(row['created_at']))}</td>
    </tr>
    """


def common_person_form_html(edit_person) -> str:
    editing = bool(edit_person)
    active = True if not edit_person else bool(edit_person["active"])
    roles = split_person_roles(row_value(edit_person, "default_role"))
    if edit_person and bool(edit_person["is_local_resident_director"]) and "Local Resident Director" not in roles:
        roles.append("Local Resident Director")
    return f"""
    <form method="post" action="/settings/common-people/save" enctype="multipart/form-data" class="admin-form">
      <input type="hidden" name="id" value="{h(row_value(edit_person, 'id'))}">
      <label>显示名称 / 正式全名<input name="display_name" value="{h(row_value(edit_person, 'display_name'))}" required></label>
      <label>匹配简称 / 别名<input name="aliases" value="{h(row_value(edit_person, 'aliases'))}" placeholder="例如 fendi, trang；多个用逗号分开"></label>
      <div class="wide role-picker"><strong>身份/职务</strong>{role_checkboxes(roles)}</div>
      <label>证件类型<input name="id_type" value="{h(row_value(edit_person, 'id_type'))}" placeholder="NRIC / Passport / FIN"></label>
      <label>证件号码<input name="id_number" value="{h(row_value(edit_person, 'id_number'))}"></label>
      <label>国籍<input name="nationality" value="{h(row_value(edit_person, 'nationality'))}"></label>
      <label>电话<input name="phone" value="{h(row_value(edit_person, 'phone'))}"></label>
      <label>Email<input name="email" value="{h(row_value(edit_person, 'email'))}"></label>
      <label class="wide">住址<input name="residential_address" value="{h(row_value(edit_person, 'residential_address'))}"></label>
      <label>签名简称<input name="signature_text" value="{h(row_value(edit_person, 'signature_text'))}" placeholder="例如 Fendi / L.T.N. Trang"></label>
      <label class="wide">签名 PNG<input name="signature_file" type="file" accept="image/png,.png"></label>
      <input type="hidden" name="signature_image_path" value="{h(row_value(edit_person, 'signature_image_path'))}">
      <input type="hidden" name="is_local_resident_director" value="{'1' if 'Local Resident Director' in roles or 'Nominee Director' in roles else '0'}">
      <label class="check"><input type="checkbox" name="auto_signature_enabled" value="1" {checked_attr(bool(row_value(edit_person, 'auto_signature_enabled')))}> 生成文件时自动套用签名</label>
      <label class="check"><input type="checkbox" name="regenerate_signature" value="1"> 重新生成文字签名图片</label>
      <label class="check"><input type="checkbox" name="active" value="1" {checked_attr(active)}> 启用</label>
      <label class="wide">备注<input name="notes" value="{h(row_value(edit_person, 'notes'))}"></label>
      <button type="submit">{'保存修改' if editing else '新增常用人员'}</button>
      {'<a class="button-link secondary" href="/settings">取消编辑</a>' if editing else ''}
    </form>
    """


def common_person_table_row(row) -> str:
    active = bool(row["active"])
    toggle_label = "停用" if active else "启用"
    toggle_value = "0" if active else "1"
    status = "<span class='status-pill ok'>启用</span>" if active else "<span class='status-pill muted'>停用</span>"
    row_keys = set(row.keys())
    signature_enabled = bool(row["auto_signature_enabled"]) if "auto_signature_enabled" in row_keys else False
    signature_text = row["signature_text"] if "signature_text" in row_keys and row["signature_text"] else ""
    signature_image_path = row["signature_image_path"] if "signature_image_path" in row_keys and row["signature_image_path"] else ""
    signature_status = (
        f"<span class='status-pill ok'>已启用</span> {'PNG' if signature_image_path else '文字'} {h(signature_text)}"
        if signature_enabled
        else "<span class='status-pill muted'>未启用</span>"
    )
    note = row["notes"] or ""
    if str(note).startswith("示例"):
        note = f"{note}；示例数据"
    toggle_form = f"""
    <form method="post" action="/settings/common-people/toggle" class="inline-form">
      <input type="hidden" name="id" value="{h(row['id'])}">
      <input type="hidden" name="active" value="{toggle_value}">
      <button class="secondary-action" type="submit">{toggle_label}</button>
    </form>
    """
    delete_form = f"""
    <form method="post" action="/settings/common-people/delete" class="inline-form" onsubmit="return confirm('确认删除这个常用人员？已上传任务不会受影响。')">
      <input type="hidden" name="id" value="{h(row['id'])}">
      <button class="danger small-action" type="submit">删除</button>
    </form>
    """
    return f"""
    <tr>
      <td>{h(row['display_name'])}</td>
      <td>{h(row['aliases'] if 'aliases' in row.keys() and row['aliases'] else '')}</td>
      <td>{role_badges(row['default_role'])}</td>
      <td>{h(row['id_type'])}</td>
      <td>{h(row['id_number'])}</td>
      <td>{h(row['nationality'])}</td>
      <td>{h(row['phone'])}</td>
      <td>{status}</td>
      <td>{signature_status}</td>
      <td>{h(note)}</td>
      <td><div class="row-actions"><a href="/settings?edit_person={h(row['id'])}#people">编辑</a>{toggle_form}{delete_form}</div></td>
    </tr>
    """


def load_template_registry() -> dict[str, dict[str, object]]:
    if not TEMPLATE_REGISTRY_PATH.exists():
        return {}
    try:
        data = json.loads(TEMPLATE_REGISTRY_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_template_registry(data: dict[str, dict[str, object]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_REGISTRY_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def package_code(package_name: str, task_type: str = "") -> str:
    if task_type == "incorporation":
        return "P1"
    return {
        "普通变更 DR 包": "M01",
        "转入包": "M02",
        "股东决议包": "M02/MR",
        "董事/秘书任免包": "M01",
        "股份转让包": "M03",
        "增资配股包": "M04",
        "年审包": "M05",
        "内部核对": "CHECK",
    }.get(package_name, "")


def generate_button_label(code: str, user: dict | None) -> str:
    admin = bool(user and user.get("role") == "admin")
    labels = {
        "P1": "生成注册文件 PDF 包",
        "M01": "生成普通变更董事决议 PDF 包",
        "M02": "生成转入文件 PDF 包",
        "M03": "生成股份转让 PDF 包",
        "M04": "生成增资配股 PDF 包",
        "M05": "生成年审 PDF 包",
    }
    if admin and code != "P1":
        return f"{labels.get(code, '生成 PDF 包')}（{code}）"
    return labels.get(code, "生成 PDF 包")


def download_label(code: str, user: dict | None) -> str:
    admin = bool(user and user.get("role") == "admin")
    labels = {
        "P1": "下载注册文件 PDF 包",
        "M01": "下载普通变更董事决议 PDF 包",
        "M02": "下载转入文件 PDF 包",
        "M03": "下载股份转让 PDF 包",
        "M04": "下载增资配股 PDF 包",
        "M05": "下载年审 PDF 包",
    }
    if admin and code != "P1":
        return f"{labels.get(code, '下载 PDF 包')}（{code}）"
    return labels.get(code, "下载 PDF 包")


def summary_cards(summary: dict[str, object]) -> str:
    task_type = str(summary.get("task_type", ""))
    if task_type == "maintenance":
        cards = [
            ("任务类型", task_label(task_type)),
            ("文件包", summary.get("package_count", "-")),
            ("识别事项", summary.get("detected_change_count", "-")),
            ("DR 组", summary.get("ordinary_dr_groups", "-")),
            ("转入", summary.get("transfer_in", "-")),
            ("转股", summary.get("share_transfers", "-")),
            ("增资", summary.get("share_allotments", "-")),
            ("年审", summary.get("annual_review", "-")),
            ("复核项", summary.get("manual_review_items", "-")),
        ]
    elif task_type == "incorporation":
        cards = [
            ("任务类型", task_label(task_type)),
            ("P1 版本", summary.get("p1_version", P1_VERSION)),
            ("董事", summary.get("directors", "-")),
            ("秘书", summary.get("secretaries", "-")),
            ("股东", summary.get("shareholders", "-")),
            ("RORC 控制人", summary.get("registrable_controllers", "-")),
        ]
    else:
        cards = [
            ("任务类型", task_label(task_type)),
            ("识别事项", summary.get("detected_change_count", "-")),
            ("严重错误", len(summary.get("blocking_errors", []) or [])),
            ("提醒", len(summary.get("warnings", []) or [])),
        ]
    html_cards = "".join(f"<div class='metric'><span>{h(label)}</span><strong>{h(value)}</strong></div>" for label, value in cards)
    return f"<div class='metrics'>{html_cards}</div>"


def package_status_cards(files: list[dict[str, object]], summary: dict[str, object], user: dict | None) -> str:
    task_type = str(summary.get("task_type", ""))
    if task_type == "incorporation":
        code = dev_code("P1", user)
        return f"""
        <div class="package-grid">
          <div class="package-card available">
            <span>可生成</span>
            <strong>注册文件包</strong>
            <p>董事决议、Form 45/45B、股权证书、Form 24、RORC、秘书服务协议和挂名董事协议。</p>
            {code}
          </div>
        </div>
        """

    grouped: dict[str, list[dict[str, object]]] = {}
    for item in files:
        package = str(item.get("package") or "其他文件")
        grouped.setdefault(package, []).append(item)
    if not grouped:
        return ""

    cards = []
    for package, items in grouped.items():
        manual = any(bool(item.get("manual_review")) for item in items)
        available = (
            (package == "普通变更 DR 包" and summary.get("m01_available") == "Yes")
            or (package == "转入包" and summary.get("m02_available") == "Yes")
            or (package == "股份转让包" and summary.get("m03_available") == "Yes")
            or (package == "增资配股包" and summary.get("m04_available") == "Yes")
            or (package == "年审包" and summary.get("m05_available") == "Yes")
        )
        status = "可生成" if available else "需复核" if manual else "预览"
        class_name = "available" if available else "review" if manual else "pending"
        names = "、".join(str(item.get("name") or "") for item in items[:3] if item.get("name"))
        if len(items) > 3:
            names += f" 等 {len(items)} 项"
        code = dev_code(package_code(package), user)
        cards.append(
            f"""
            <div class="package-card {class_name}">
              <span>{h(status)}</span>
              <strong>{h(package)}</strong>
              <p>{h(names or '暂无明细')}</p>
              {code}
            </div>
            """
        )
    return f"<div class='package-grid'>{''.join(cards)}</div>"


def auto_generation_notice(is_generating: bool) -> str:
    if not is_generating:
        return ""
    return """
    <div class="info-box generating-box">
      <span class="spinner" aria-hidden="true"></span>
      <div>
        <strong>正在生成 PDF 文件包</strong>
        <p>上传已完成，系统正在后台生成。页面会自动刷新，完成后下载按钮会显示在这里。</p>
      </div>
    </div>
    <script>setTimeout(() => window.location.reload(), 5000);</script>
    """


def auto_generation_control(row, has_download: bool, can_generate_any: bool, has_errors: bool) -> str:
    if is_generation_status(row["status"]):
        return "<p class='muted'>系统正在生成文件包，请稍等。生成完成后页面会自动出现下载按钮。</p>"
    if has_download:
        return "<p class='muted'>文件已经生成。请直接下载上方文件包；如果判断不对，修改 Excel 后重新上传一次即可。</p>"
    if row["status"] == "generation_failed":
        return "<div class='error-box'><strong>自动生成失败</strong><p>请检查表格字段或稍后重新上传。管理员可以在后台日志查看失败原因。</p></div>"
    if has_errors:
        return "<p class='error'>系统检查发现严重错误，请修改 Excel 后重新上传。</p>"
    if can_generate_any:
        return "<p class='muted'>这个旧任务尚未自动生成。请重新上传一次表格，系统会自动生成文件包。</p>"
    return "<p class='muted'>当前表格没有识别到已接入的生成器项目。请检查表格选项或换用当前版本模板。</p>"


def generated_files_panel(packages: list[tuple[str, Path]]) -> str:
    sections = []
    for label, path in packages:
        if not path.exists():
            continue
        try:
            with zipfile.ZipFile(path) as zf:
                names = [name for name in zf.namelist() if name.lower().endswith(".pdf")]
        except zipfile.BadZipFile:
            names = []
        items = "".join(f"<li>{h(name)}</li>" for name in names[:30])
        more = f"<li>... 另有 {len(names) - 30} 个文件</li>" if len(names) > 30 else ""
        sections.append(
            f"""
            <details class="generated-package" open>
              <summary>{h(label)} · {len(names)} 个 PDF</summary>
              <ul>{items}{more}</ul>
            </details>
            """
        )
    if not sections:
        return ""
    return f"<div class='generated-files'><strong>已生成文件</strong>{''.join(sections)}</div>"


def review_workflow(summary: dict[str, object], warnings: list[str], blocking_errors: list[str], user: dict | None) -> str:
    if blocking_errors:
        items = "".join(f"<li>{h(item)}</li>" for item in blocking_errors)
        return f"""
        <div class="error-box">
          <strong>必须先修正资料表</strong>
          <p>下面问题会影响文件生成。请回到 Excel 修改后重新上传一份新的资料表。</p>
          <ul>{items}</ul>
        </div>
        """
    manual_count = int(summary.get("manual_review_items") or 0)
    if not warnings and not manual_count:
        return """
        <div class="info-box">
          <strong>检查通过</strong>
          <p>没有严重错误。请核对文件预览和签字人，确认后生成 PDF。</p>
        </div>
        """
    items = "".join(f"<li>{h(item)}</li>" for item in warnings)
    admin_note = "<p class='muted'>管理员可在下方文件规则明细中查看哪些文件标记为“需要复核”。</p>" if user and user.get("role") == "admin" else ""
    return f"""
    <div class="warning">
      <strong>需要人工复核</strong>
      <p>这些不是系统阻断错误，但生成或发给客户前需要人工确认。复核方式：检查 Excel 原始数据、BizFile 信息、签字人和文件清单；确认无误后可以继续生成。</p>
      <ul>{items or '<li>文件清单中存在需要复核的项目。</li>'}</ul>
      {admin_note}
    </div>
    """


def is_generation_status(status: str) -> bool:
    return status == "generating_pdf" or status.endswith("_generating_pdf")


def review_workflow(summary: dict[str, object], warnings: list[str], blocking_errors: list[str], user: dict | None) -> str:
    if blocking_errors:
        items = "".join(f"<li>{h(item)}</li>" for item in blocking_errors)
        return f"""
        <div class="error-box">
          <strong>需要先修改表格</strong>
          <p>下面问题会影响文件生成。请回到 Excel 修改后重新上传。</p>
          <ul>{items}</ul>
        </div>
        """
    manual_count = int(summary.get("manual_review_items") or 0)
    if not warnings and not manual_count:
        return """
        <div class="info-box">
          <strong>检查通过</strong>
          <p>没有严重错误。系统会自动生成 PDF 文件包；下方仍保留判断信息，方便你快速核对。</p>
        </div>
        """
    items = "".join(f"<li>{h(item)}</li>" for item in warnings)
    admin_note = "<p class='muted'>管理员可在下方文件规则明细中查看哪些文件标记为需要复核。</p>" if user and user.get("role") == "admin" else ""
    return f"""
    <div class="warning">
      <strong>需要人工复核</strong>
      <p>这些不是系统阻断错误，文件仍会自动生成；发送客户前请核对 Excel、BizFile、签字人和文件清单。</p>
      <ul>{items or '<li>文件清单中存在需要复核的项目。</li>'}</ul>
      {admin_note}
    </div>
    """


def review_progress(row, summary: dict[str, object], can_generate: bool, can_generate_p2: bool, has_download: bool) -> str:
    has_errors = bool(summary.get("blocking_errors"))
    is_generating = is_generation_status(row["status"])
    can_generate_any = bool(can_generate or can_generate_p2)
    steps = [
        ("1", "上传资料表", "完成", "done"),
        ("2", "系统检查", "需修正" if has_errors else "已通过", "error" if has_errors else "done"),
        ("3", "人工复核", "检查文件清单和签字人", "active" if not has_errors and not has_download and not is_generating else "done" if has_download or is_generating else "pending"),
        ("4", "生成 PDF", "生成中" if is_generating else "可生成" if can_generate_any else "等待接入" if not has_errors else "不可生成", "active" if is_generating or (can_generate_any and not has_download) else "done" if has_download else "pending"),
        ("5", "下载文件包", "已生成" if has_download else "生成后下载", "done" if has_download else "pending"),
    ]
    cards = "".join(
        f"""
        <div class="process-step {h(class_name)}">
          <span>{h(num)}</span>
          <strong>{h(title)}</strong>
          <small>{h(detail)}</small>
        </div>
        """
        for num, title, detail, class_name in steps
    )
    return f"<div class='process-flow'>{cards}</div>"


def generation_steps(row, summary: dict[str, object], can_generate: bool, can_generate_p2: bool, user: dict | None) -> str:
    task_type = row["task_type"]
    has_errors = bool(summary.get("blocking_errors"))
    if is_generation_status(row["status"]):
        steps = [
            ("1", "后台生成中", "系统正在生成 Word 和 PDF 文件包。"),
            ("2", "页面自动刷新", "生成完成后会出现下载按钮。"),
            ("3", "下载文件包", "如果文件较多，可能需要几十秒。"),
        ]
    elif has_errors:
        steps = [
            ("1", "下载或打开原 Excel", "修正严重错误对应字段。"),
            ("2", "重新上传资料表", "系统会重新检查并生成新的任务。"),
            ("3", "再生成文件", "通过检查后才开放生成按钮。"),
        ]
    elif task_type == "incorporation":
        steps = [
            ("1", "核对注册资料", "确认公司名、董事、秘书、股东、FYE 和签字人。"),
            ("2", "处理复核提醒", "如有提醒，人工确认后继续。"),
            ("3", "生成注册 PDF 包", "下载后发客户签署或内部存档。"),
        ]
    elif task_type == "maintenance":
        action = "生成已接入的 PDF 文件包" if can_generate_p2 else "等待对应文件包接入"
        detail = "当前已接入普通董事决议、转入、股份转让、增资配股和年审包；生成前请先核对复核提醒和签字人。" if can_generate_p2 else "当前事项还没有对应生成器，先按文件预览复核。"
        steps = [
            ("1", "核对变更事项", "确认每个事项是否需要生成、是否同组 DR、签字人是否正确。"),
            ("2", "处理复核提醒", "高风险事项先人工确认，再进入生成。"),
            ("3", action, detail),
        ]
    else:
        steps = [
            ("1", "核对文件清单", "旧变更表先做判断预览。"),
            ("2", "换用 v3 表", "需要生成时建议使用维护/变更/年审 v3 表。"),
            ("3", "重新上传", "由系统重新判断。"),
        ]
    cards = "".join(
        f"<div class='step-card'><span>{h(num)}</span><strong>{h(title)}</strong><p>{h(detail)}</p></div>"
        for num, title, detail in steps
    )
    return f"<div class='step-grid'>{cards}</div>"


def review_progress(row, summary: dict[str, object], can_generate: bool, can_generate_p2: bool, has_download: bool) -> str:
    has_errors = bool(summary.get("blocking_errors"))
    is_generating = is_generation_status(row["status"])
    can_generate_any = bool(can_generate or can_generate_p2)
    steps = [
        ("1", "上传资料表", "完成", "done"),
        ("2", "系统检查", "需修改" if has_errors else "已通过", "error" if has_errors else "done"),
        (
            "3",
            "自动生成",
            "生成中" if is_generating else "已生成" if has_download else "等待重新上传" if can_generate_any else "无可生成项目",
            "active" if is_generating else "done" if has_download else "pending",
        ),
        ("4", "下载文件包", "已生成" if has_download else "生成后下载", "done" if has_download else "pending"),
    ]
    cards = "".join(
        f"""
        <div class="process-step {h(class_name)}">
          <span>{h(num)}</span>
          <strong>{h(title)}</strong>
          <small>{h(detail)}</small>
        </div>
        """
        for num, title, detail, class_name in steps
    )
    return f"<div class='process-flow'>{cards}</div>"


def generation_steps(row, summary: dict[str, object], can_generate: bool, can_generate_p2: bool, user: dict | None) -> str:
    has_errors = bool(summary.get("blocking_errors"))
    if is_generation_status(row["status"]):
        steps = [
            ("1", "后台生成中", "系统正在生成 Word 和 PDF 文件包。"),
            ("2", "页面自动刷新", "生成完成后会出现下载按钮和实际文件清单。"),
            ("3", "下载文件包", "文件较多时可能需要几十秒。"),
        ]
    elif row["status"] == "generation_failed":
        steps = [
            ("1", "检查表格", "生成失败通常是字段缺失、模板占位符或转换 PDF 出错。"),
            ("2", "修改后重传", "当前流程建议直接改 Excel 后重新上传。"),
            ("3", "后台日志", "管理员可在后台查看失败记录。"),
        ]
    elif has_errors:
        steps = [
            ("1", "打开原 Excel", "修正严重错误对应字段。"),
            ("2", "重新上传", "系统会重新检查并自动生成。"),
            ("3", "下载新文件", "新任务生成完成后直接下载。"),
        ]
    else:
        steps = [
            ("1", "查看判断信息", "系统保留文件清单和签字逻辑，方便快速核对。"),
            ("2", "直接下载", "上传后自动生成，无需再点确认生成。"),
            ("3", "有误就重传", "如果判断不对，修改 Excel 后重新上传一次。"),
        ]
    cards = "".join(
        f"<div class='step-card'><span>{h(num)}</span><strong>{h(title)}</strong><p>{h(detail)}</p></div>"
        for num, title, detail in steps
    )
    return f"<div class='step-grid'>{cards}</div>"


def preview_table(rows: list[dict[str, object]]) -> str:
    trs = "".join(
        f"<tr><td>{h(row.get('file'))}</td><td>{h(row.get('count'))}</td><td>{h(row.get('signing_logic'))}</td><td>{h(row.get('signer_detail'))}</td></tr>"
        for row in rows
    )
    return f"<table><thead><tr><th>文件</th><th>份数</th><th>签字/生成逻辑</th><th>人员/字段</th></tr></thead><tbody>{trs or '<tr><td colspan=\"4\">暂无预览</td></tr>'}</tbody></table>"


def files_table(files: list[dict[str, object]], user: dict | None = None) -> str:
    is_admin = bool(user and user.get("role") == "admin")
    code_head = "<th>内部编号</th>" if is_admin else ""
    trs = "".join(
        f"<tr><td>{h(f.get('package'))}</td>{'<td>' + h(package_code(str(f.get('package') or ''))) + '</td>' if is_admin else ''}<td>{h(f.get('doc_type'))}</td><td>{h(f.get('name'))}</td><td>{h(f.get('suggested'))}</td><td>{h(f.get('reason'))}</td><td>{h(f.get('signing'))}</td><td>{'需要复核' if f.get('manual_review') else ''}</td></tr>"
        for f in files
    )
    colspan = 8 if is_admin else 7
    empty_row = f'<tr><td colspan="{colspan}">暂无建议</td></tr>'
    return f"<table><thead><tr><th>文件包</th>{code_head}<th>类型</th><th>文件</th><th>建议</th><th>触发原因</th><th>签字方式</th><th>备注</th></tr></thead><tbody>{trs or empty_row}</tbody></table>"


def delete_job_form(row, user: dict) -> str:
    if user.get("role") != "admin":
        return ""
    return f"""
    <section class="panel danger-zone">
      <h3>任务清理</h3>
      <p class="muted">删除任务会移除上传文件、生成文件包和任务记录。</p>
      <form method="post" action="/job/delete">
        <input type="hidden" name="job_id" value="{h(row['id'])}">
        <button class="danger" type="submit">删除此任务</button>
      </form>
    </section>
    """


def parse_multipart_file(content_type: str, body: bytes, field_name: str):
    marker = "boundary="
    if marker not in content_type:
        return "", b""
    boundary = content_type.split(marker, 1)[1].strip().strip('"')
    boundary_bytes = ("--" + boundary).encode()
    parts = body.split(boundary_bytes)
    for part in parts:
        if b"Content-Disposition" not in part:
            continue
        header, _, payload = part.partition(b"\r\n\r\n")
        if f'name="{field_name}"'.encode() not in header:
            continue
        filename = ""
        disposition = header.decode("utf-8", "ignore").strip().splitlines()[0]
        for bit in disposition.split(";"):
            bit = bit.strip()
            if bit.startswith("filename="):
                filename = bit.split("=", 1)[1].strip().strip('"')
        payload = payload.rsplit(b"\r\n", 1)[0]
        return filename, payload
    return "", b""


def normalize_form_fields(fields: dict[str, str] | dict[str, list[str]]) -> dict[str, list[str]]:
    normalized: dict[str, list[str]] = {}
    for key, value in fields.items():
        if isinstance(value, list):
            normalized[key] = [str(item) for item in value]
        else:
            normalized[key] = [str(value)]
    return normalized


def save_uploaded_signature_png(display_name: str, filename: str, payload: bytes) -> str:
    if len(payload) > SIGNATURE_UPLOAD_MAX_BYTES:
        raise ValueError("签名图片不能超过 2MB。")
    if not payload:
        raise ValueError("请上传有效的 PNG 签名图片。")

    source_name = Path(filename or "").stem or display_name
    target_name = safe_slug(display_name or source_name)
    target_path = SIGNATURE_DIR / f"{target_name}.png"

    try:
        image = Image.open(BytesIO(payload))
        if image.format != "PNG":
            raise ValueError("签名图片目前只支持 PNG 格式。")
        image = image.convert("RGBA")
    except (UnidentifiedImageError, OSError) as exc:
        raise ValueError("请上传有效的 PNG 签名图片。") from exc

    SIGNATURE_DIR.mkdir(parents=True, exist_ok=True)
    image.save(target_path)
    return signature_relative_path(target_path)


def parse_multipart_form(content_type: str, body: bytes):
    marker = "boundary="
    if marker not in content_type:
        return {}, {}
    boundary = content_type.split(marker, 1)[1].strip().strip('"')
    boundary_bytes = ("--" + boundary).encode()
    fields: dict[str, str] = {}
    files: dict[str, tuple[str, bytes]] = {}
    for part in body.split(boundary_bytes):
        if b"Content-Disposition" not in part:
            continue
        header, _, payload = part.partition(b"\r\n\r\n")
        if not payload:
            continue
        header_text = header.decode("utf-8", "ignore")
        disposition = ""
        for line in header_text.splitlines():
            if line.lower().startswith("content-disposition:"):
                disposition = line
                break
        name = ""
        filename = ""
        for bit in disposition.split(";"):
            bit = bit.strip()
            if bit.startswith("name="):
                name = bit.split("=", 1)[1].strip().strip('"')
            elif bit.startswith("filename="):
                filename = bit.split("=", 1)[1].strip().strip('"')
        payload = payload.rsplit(b"\r\n", 1)[0]
        if not name:
            continue
        if filename:
            files[name] = (filename, payload)
        else:
            fields[name] = payload.decode("utf-8", "ignore").strip()
    return fields, files


def cleanup_old_uploads() -> None:
    if not UPLOAD_DIR.exists():
        return
    cutoff = datetime.now(UTC).timestamp() - UPLOAD_RETENTION_DAYS * 24 * 60 * 60
    for path in UPLOAD_DIR.iterdir():
        if path.is_file() and path.stat().st_mtime < cutoff:
            path.unlink(missing_ok=True)


def cleanup_job_files(row) -> None:
    upload_path = Path(row["upload_path"]) if row["upload_path"] else None
    if upload_path:
        safe_remove_file(upload_path, UPLOAD_DIR)
    code = safe_filename(row["job_code"])
    safe_remove_file(GENERATED_DIR / f"{code}_P1_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P1_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P1_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P1_pdf", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M01_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M01_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M01_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M01_pdf", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M02_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M02_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M02_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M02_pdf", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M03_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M03_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M03_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M03_pdf", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M04_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M04_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M04_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M04_pdf", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M05_docx_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M05_docs", GENERATED_DIR)
    safe_remove_file(GENERATED_DIR / f"{code}_P2_M05_pdf_package.zip", GENERATED_DIR)
    safe_remove_dir(GENERATED_DIR / f"{code}_P2_M05_pdf", GENERATED_DIR)


def safe_remove_file(path: Path, root: Path) -> None:
    try:
        resolved = path.resolve()
        if resolved.is_file() and resolved.is_relative_to(root.resolve()):
            resolved.unlink(missing_ok=True)
    except OSError:
        return


def safe_remove_dir(path: Path, root: Path) -> None:
    try:
        resolved = path.resolve()
        if resolved.is_dir() and resolved.is_relative_to(root.resolve()):
            shutil.rmtree(resolved)
    except OSError:
        return


def content_type_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if suffix == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if suffix == ".pdf":
        return "application/pdf"
    if suffix == ".md":
        return "text/markdown; charset=utf-8"
    if suffix == ".zip":
        return "application/zip"
    if suffix == ".css":
        return "text/css; charset=utf-8"
    return "application/octet-stream"


def task_label(task_type: str) -> str:
    return {"incorporation": "新公司注册", "change": "现有公司变更", "maintenance": "公司维护/变更年审", "unknown": "未知"}.get(task_type, task_type)


def role_label(role: str) -> str:
    return {"admin": "管理员", "staff": "普通用户"}.get(role, role)


def status_label(status: str) -> str:
    return {
        "blocked": "需修正",
        "needs_review": "待确认",
        "parsed": "已解析",
        "docx_generated": "已生成旧文件包",
        "p2_m01_generated": "已生成普通董事决议",
        "generating_pdf": "正在生成 PDF",
        "p2_m01_generating_pdf": "正在生成普通董事决议 PDF",
        "p2_m02_generating_pdf": "正在生成转入文件 PDF",
        "p2_m03_generating_pdf": "正在生成股份转让 PDF",
        "p2_m04_generating_pdf": "正在生成增资配股 PDF",
        "p2_m05_generating_pdf": "正在生成年审 PDF",
        "generation_failed": "生成失败，可重试",
        "pdf_generated": "已生成 PDF 包",
        "p2_m01_pdf_generated": "已生成普通董事决议 PDF",
        "p2_m02_pdf_generated": "已生成转入文件 PDF",
        "p2_m03_pdf_generated": "已生成股份转让 PDF",
        "p2_m04_pdf_generated": "已生成增资配股 PDF",
        "p2_m05_pdf_generated": "已生成年审 PDF",
    }.get(status, status)


def main():
    init_db()
    cleanup_old_uploads()
    host = os.environ.get("APP_HOST", "127.0.0.1")
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("APP_PORT", "8088"))
    httpd = ThreadingHTTPServer((host, port), App)
    print(f"{APP_NAME} running at http://{host}:{port}")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
