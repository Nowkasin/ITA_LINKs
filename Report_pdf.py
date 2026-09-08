from __future__ import annotations

import hmac
import os
from contextlib import contextmanager
from datetime import datetime
from functools import wraps
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import quote, urlparse

import pymssql
from flask import Flask, Response, request, jsonify, render_template, session, redirect
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from weasyprint import HTML
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from dotenv import load_dotenv

load_dotenv()
app = Flask(__name__)

# SESSION_SECRET_KEY ใน .env จำเป็นสำหรับเซ็น cookie — สร้างด้วย:
# python -c "import secrets; print(secrets.token_hex(32))"
app.secret_key = os.getenv("SESSION_SECRET_KEY", "")
if not app.secret_key:
    raise RuntimeError("ไม่พบ SESSION_SECRET_KEY ใน .env กรุณาตั้งค่าก่อนรันระบบ")

# ความปลอดภัยของ session cookie — SESSION_COOKIE_SECURE ควรเป็น true เสมอเมื่อขึ้น production จริง
# (รันหลัง HTTPS แล้ว) แต่ปล่อย false ไว้ตอน dev/ทดสอบในเครื่องผ่าน http://127.0.0.1 เพราะ Secure
# cookie จะไม่ถูกส่งเลยถ้าไม่ใช่ HTTPS ทำให้ session ใช้งานไม่ได้ตอนทดสอบในเครื่อง
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true"

BASE_DIR = Path(__file__).resolve().parent

# ===================== K2 SSO TOKEN (signed, หมดอายุอัตโนมัติ) =====================
# K2_TOKEN_SECRET: แนะนำให้ตั้งแยกจาก SESSION_SECRET_KEY (คนละ key คนละหน้าที่) — สร้างด้วย:
#   python -c "import secrets; print(secrets.token_hex(32))"
K2_TOKEN_SECRET = os.getenv("K2_TOKEN_SECRET", app.secret_key)
# K2_API_KEY: secret key เฉพาะสำหรับให้ "เซิร์ฟเวอร์ K2" (ไม่ใช่ browser user) เรียกขอ token — ต้องตั้งใน .env ก่อนใช้จริง
K2_API_KEY = os.getenv("K2_API_KEY", "")
K2_TOKEN_MAX_AGE_SECONDS = int(os.getenv("K2_TOKEN_MAX_AGE_SECONDS", "300"))  # token ใช้ได้ 5 นาทีหลังออก

_k2_token_serializer = URLSafeTimedSerializer(K2_TOKEN_SECRET, salt="k2-sso-token")


def issue_k2_token(buasri_id: str) -> str:
    """ออก token ที่เซ็นด้วย K2_TOKEN_SECRET ผูกกับ buasri_id — ปลอม/แก้ไขไม่ได้ถ้าไม่รู้ secret"""
    return _k2_token_serializer.dumps({"buasri_id": buasri_id})


def verify_k2_token(token: str) -> Optional[str]:
    """ตรวจลายเซ็นและอายุ token — คืน buasri_id ถ้าผ่าน, None ถ้าไม่ผ่าน (ปลอม/หมดอายุ)"""
    try:
        data = _k2_token_serializer.loads(token, max_age=K2_TOKEN_MAX_AGE_SECONDS)
        return (data or {}).get("buasri_id") or None
    except (BadSignature, SignatureExpired):
        return None


DB_CONFIG = {
    "server": os.getenv("DB_SERVER"),
    "port": int(os.getenv("DB_PORT", "1433")),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": os.getenv("DB_NAME"),
    "charset": "UTF-8",
}

SP_NAME = "dbo.SP_4001_Ita_Procurement"

THAI_MONTHS = {
    1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน",
    5: "พฤษภาคม", 6: "มิถุนายน", 7: "กรกฎาคม", 8: "สิงหาคม",
    9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม",
}
THAI_MONTH_NAME_TO_NUM = {name: num for num, name in THAI_MONTHS.items()}  # กัน K2 ส่งชื่อเดือนไทยแทนตัวเลข เช่น "สิงหาคม"

THAI_MONTH_ABBR = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}


def format_thai_short_date(value: Any) -> str:
    """
    แปลง date/datetime (หรือ string 'YYYY-MM-DD') -> 'D เดือนย่อ. YY' (พ.ศ. 2 หลัก)
    เช่น 4 ส.ค. 69 — ให้ตรงรูปแบบเดียวกับที่ SP_4001 คำนวณเองใน RefDocNoDisplay
    """
    if not value:
        return "-"
    try:
        if hasattr(value, "day") and hasattr(value, "month") and hasattr(value, "year"):
            day, month, year_be = value.day, value.month, value.year + 543
        else:
            parts = str(value)[:10].split("-")
            if len(parts) != 3:
                return safe_text(value)
            year_ce, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            year_be = year_ce + 543
        return f"{day} {THAI_MONTH_ABBR.get(month, '')} {str(year_be)[-2:]}"
    except Exception:
        return safe_text(value)


try:
    from pythainlp.tokenize import word_tokenize
except ImportError:
    word_tokenize = None


def insert_thai_break_hints(text: str) -> str:
    """
    แทรก zero-width space (U+200B) ระหว่างคำไทยที่ตัดคำได้ (ผ่าน pythainlp)
    เพื่อให้ WeasyPrint/Pango มีจุดตัดบรรทัดที่ถูกต้อง — ปกติ Pango บน Windows ไม่มี
    ไลบรารีตัดคำไทย (libthai) จึงตัดกลางคำแบบสุ่มถ้าไม่มีช่องว่าง/จุดตัดให้เลย
    ถ้าไม่ได้ติดตั้ง pythainlp จะคืนข้อความเดิมเฉยๆ (fallback ปลอดภัย ไม่ error)
    """
    if not text or text == "-" or word_tokenize is None:
        return text
    try:
        return "\u200b".join(word_tokenize(text, engine="newmm"))
    except Exception:
        return text

# ===================== HELPERS =====================
def get_arg(*names: str, default: str = "") -> str:
    """อ่านค่าได้ทั้งจาก query string (GET) และฟอร์ม POST (request.values = args รวมกับ form)"""
    for name in names:
        value = request.values.get(name)
        if value is not None and str(value).strip() != "":
            return str(value).strip()
    return default

def to_buddhist_year(value: Optional[int]) -> Optional[int]:
    """
    แปลงปี ค.ศ. -> พ.ศ. อัตโนมัติ ถ้าค่าที่ได้มาต่ำกว่า 2400 (แปลว่าเป็น ค.ศ. แน่ๆ เพราะ พ.ศ. ยุคนี้เกิน 2400 เสมอ)
    เผื่อ K2 หรือใครก็ตามส่ง FiscalYear มาเป็น ค.ศ. (เช่น 2026) แทนที่จะเป็น พ.ศ. (2569) ตามที่ SP ต้องการ
    """
    if value is None:
        return None
    return value + 543 if value < 2400 else value


def to_int_or_none(value: str) -> Optional[int]:
    value = (value or "").strip()
    try:
        return int(value) if value else None
    except ValueError:
        return None


def parse_month(value: str) -> Optional[int]:
    """รับได้ทั้ง '8' (ตัวเลข) และ 'สิงหาคม' (ชื่อเดือนไทยเต็ม) — เผื่อ K2 ส่งมาเป็นชื่อเดือนแทนตัวเลข"""
    value = (value or "").strip()
    if not value:
        return None
    as_int = to_int_or_none(value)
    if as_int is not None:
        return as_int
    return THAI_MONTH_NAME_TO_NUM.get(value)


def thai_month_name(month_no: Optional[int]) -> str:
    return THAI_MONTHS.get(month_no, "") if month_no else ""


def format_currency(value: Any) -> str:
    if value is None or value == "":
        return "-"
    try:
        num = float(value)
        return f"{int(num):,}" if num.is_integer() else f"{num:,.2f}"
    except Exception:
        return str(value)


def safe_text(value: Any) -> str:
    text = str(value).strip() if value is not None else ""
    return text if text else "-"


def content_disposition_header(filename: str) -> str:
    """HTTP header เป็น Latin-1 เท่านั้น ชื่อไฟล์ไทยต้อง encode แบบ RFC 5987 (filename*=UTF-8''...)"""
    return f'attachment; filename="download"; filename*=UTF-8\'\'{quote(filename, safe="")}'


def safe_fetch(fn, default):
    try:
        return fn()
    except Exception:
        return default


def is_safe_next_path(path: str) -> bool:
    """
    กัน open redirect: next ต้องเป็น path ภายในระบบเดียวกันเท่านั้น (ขึ้นต้นด้วย '/' ตัวเดียว,
    ไม่ใช่ '//evil.com', ไม่มี scheme/host แปลกปน) ไม่งั้นใครก็ปลอมลิงก์ /sso?...&next=https://phishing.example ได้
    """
    if not path or not path.startswith("/") or path.startswith("//"):
        return False
    parsed = urlparse(path)
    return not parsed.netloc and not parsed.scheme


# ===================== DB =====================
@contextmanager
def db_cursor(commit: bool = False):
    conn = pymssql.connect(**DB_CONFIG)
    try:
        cur = conn.cursor(as_dict=True)
        yield cur
        if commit:
            conn.commit()
    finally:
        conn.close()


def fetch_rows(fiscal_year, month, dept_id, plant: Optional[int] = None) -> List[Dict[str, Any]]:
    """SP_4001: @FiscalYear, @Month, @Search_Dept_ID, @Search_Plant — กรองปีงบ/เดือนจากคอลัมน์ RefDocNoDate (ยืนยันจาก SP จริง), ไม่กรอง UseOrNot"""
    with db_cursor() as cur:
        cur.execute(f"EXEC {SP_NAME} %s, %s, %s, %s", (fiscal_year, month, dept_id, plant))
        return cur.fetchall()


def fetch_departments(only_active: bool = True) -> List[Dict[str, Any]]:
    """ใช้เฉพาะตอน resolve ชื่อหน่วยงานที่ K2 ส่งมา (CUR_INDEPT_LNAME_TH) -> Dept_id"""
    with db_cursor() as cur:
        cur.execute("EXEC dbo.SP_4005_Plant_List %s", (1 if only_active else 0,))
        return [{"id": safe_text(r.get("Dept_id")), "name": safe_text(r.get("Dept_name"))} for r in cur.fetchall()]


def fetch_purchasing_units() -> List[Dict[str, Any]]:
    """หน่วยจัดซื้อ (Plant) — ใช้ resolve ชื่อที่ K2 ส่งมา และใช้แสดงชื่อหน่วยงานในหัวรายงาน"""
    with db_cursor() as cur:
        cur.execute(
            "SELECT [Plant], [PlanName], [Group_Report] FROM [dbo].[Plant_ms] "
            "WHERE [active] = 1 ORDER BY [PlanName]"
        )
        return [
            {"id": safe_text(r.get("Plant")), "name": safe_text(r.get("PlanName")),
             "group": safe_text(r.get("Group_Report"))}
            for r in cur.fetchall()
        ]


def fetch_group_report_list() -> List[Dict[str, Any]]:
    with db_cursor() as cur:
        cur.execute("SELECT [Group_Report], [Group_Report_Desc] FROM [dbo].[Group_Report_ms] ORDER BY [Group_Report]")
        return [{"value": safe_text(r.get("Group_Report")), "label": safe_text(r.get("Group_Report_Desc"))} for r in cur.fetchall()]


def fetch_plant_list_by_group(group_report: str) -> List[str]:
    with db_cursor() as cur:
        cur.execute("SELECT [Plant] FROM [dbo].[Plant_ms] WHERE [active] = 1 AND [Group_Report] = %s", (group_report,))
        return [safe_text(r.get("Plant")) for r in cur.fetchall()]


# ===================== AUTH / USER ACCOUNT =====================
# หมายเหตุ: ไม่มีหน้า login เดิม/หน้าเว็บจัดการข้อมูลแล้ว — มีแค่ 2 อย่าง: /sso (login จาก K2)
# และ /procurement-report.pdf|.xlsx (ดาวน์โหลดรายงาน) เท่านั้น

def check_buasri_id(buasri_id: str) -> Optional[Dict[str, Any]]:
    with db_cursor() as cur:
        cur.execute("EXEC dbo.SP_4006_User_CheckBuasriID %s", (buasri_id,))
        return cur.fetchone()


def get_plant_by_user(buasri_id: str) -> List[Dict[str, Any]]:
    with db_cursor() as cur:
        cur.execute("EXEC dbo.SP_4001_GetPlantByUser %s", (buasri_id,))
        return cur.fetchall() or []


def is_resource_admin_dept(buasri_id: str) -> bool:
    with db_cursor() as cur:
        cur.execute(
            "SELECT cur_dept_cd FROM [**.*.**.***\\MS_ASSET_ERP].[SAP_FI].[dbo].[v_ess_Person] WHERE BUASRI_ID = %s",
            (buasri_id,),
        )
        row = cur.fetchone()
    return bool(row) and str(row.get("cur_dept_cd")) == "5029"


def get_or_create_user(buasri_id: str, full_name: str, is_admin: bool) -> Optional[Dict[str, Any]]:
    with db_cursor(commit=True) as cur:
        cur.execute("EXEC dbo.SP_4007_User_GetOrCreate %s, %s, %s", (buasri_id, full_name, 1 if is_admin else 0))
        return cur.fetchone()


def save_dept_access(user_id: int, buasri_id: str, plant_rows: List[Dict[str, Any]]) -> None:
    with db_cursor(commit=True) as cur:
        if not plant_rows:
            cur.execute("DELETE FROM dbo.User_Dept_Access WHERE BuasriID = %s", (buasri_id,))
            return
        for idx, row in enumerate(plant_rows):
            cur.execute(
                """
                EXEC dbo.SP_4008_User_SaveDeptAccess
                    @UserID=%s, @BuasriID=%s, @Plant=%s, @Dept_id=%s, @Dept_name=%s, @Group_Report=%s, @IsFirstRow=%s;
                """,
                (user_id, buasri_id, row.get("Plant"), row.get("Indept"), row.get("PlanName"),
                 row.get("Group_Report"), 1 if idx == 0 else 0),
            )


def establish_session_for_buasri_id(buasri_id: str):
    """
    สร้าง session ให้ผู้ใช้จาก buasri_id เพียงอย่างเดียว (เชื่อว่ายืนยันตัวตนมาจาก K2 แล้ว)
    คืนค่า None ถ้าสำเร็จ, หรือ (json_response, status_code) ถ้ามี error ให้ตอบกลับ
    """
    person = check_buasri_id(buasri_id)
    if not person:
        return jsonify({"error": "ไม่พบข้อมูลผู้ใช้นี้ในระบบมหาวิทยาลัย หรือสถานะไม่ active"}), 403

    full_name = safe_text(person.get("FullName"))
    is_admin = safe_fetch(lambda: is_resource_admin_dept(buasri_id), False)

    user = get_or_create_user(buasri_id, full_name if full_name != "-" else None, is_admin)
    if not user:
        return jsonify({"error": "เข้าสู่ระบบไม่สำเร็จ กรุณาลองใหม่อีกครั้ง"}), 500

    plant_rows = safe_fetch(lambda: get_plant_by_user(buasri_id), [])
    try:
        save_dept_access(user["ID"], buasri_id, plant_rows)
    except Exception:
        pass

    dept_ids = [safe_text(r.get("Indept")) for r in plant_rows if r.get("Indept") is not None]
    plant_codes = [safe_text(r.get("Plant")) for r in plant_rows if r.get("Plant") is not None]
    dept_name = safe_text(plant_rows[0].get("PlanName")) if plant_rows else "-"

    session["user_id"] = user["ID"]
    session["buasri_id"] = user["BuasriID"]
    session["full_name"] = safe_text(user.get("FullName"))
    session["role"] = user.get("Role") or "staff"
    session["dept_ids"] = dept_ids
    session["plants"] = plant_codes
    session["dept_name"] = dept_name if dept_name != "-" else ""
    return None


def login_required(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if not session.get("user_id"):
            return jsonify({"error": "กรุณาเข้าใช้งานผ่านลิงก์ที่ส่งมาจากระบบ K2"}), 401
        return view_func(*args, **kwargs)
    return wrapped

# ===================== K2 INTEGRATION HELPERS =====================
def resolve_dept_id_by_name(name: str) -> Optional[str]:
    """แปลงชื่อหน่วยงานจาก K2 (CUR_INDEPT_LNAME_TH) -> Dept_id ด้วย exact match"""
    return resolve_dept_by_name(name)[0]

def resolve_dept_by_name(name: str) -> Tuple[Optional[str], Optional[str]]:
    """
    คืน (Dept_id, Dept_name) จาก DB ถ้า match ชื่อที่ K2 ส่งมาแบบ exact เป๊ะ, (None, None) ถ้าไม่เจอ
    สำคัญ: ใช้ Dept_name ที่คืนมา (ไม่ใช่ข้อความดิบจาก K2) ไปโชว์หัวรายงาน เพื่อการันตีว่าชื่อที่โชว์
    ตรงกับข้อมูลที่ถูกกรองจริงเสมอ — ถ้า resolve ไม่เจอ ห้ามเอาข้อความดิบไปโชว์เด็ดขาด (จะโชว์ชื่อที่ไม่มีจริง)
    """
    if not name:
        return None, None
    name_norm = name.strip()
    for d in safe_fetch(lambda: fetch_departments(True), []):
        if d["name"].strip() == name_norm:
            return d["id"], d["name"]
    app.logger.warning("resolve_dept_by_name: ไม่พบหน่วยงาน '%s'", name_norm)
    return None, None

def resolve_plant_by_name(name: str) -> Optional[str]:
    """แปลงชื่อ Plant จาก K2 (PlantName) -> รหัส Plant"""
    if not name:
        return None
    name_norm = name.strip()
    for u in safe_fetch(fetch_purchasing_units, []):
        if u["name"].strip() == name_norm:
            return u["id"]
    app.logger.warning("resolve_plant_by_name: ไม่พบหน่วยจัดซื้อ '%s'", name_norm)
    return None

def resolve_group_report_by_label(text: str) -> Optional[str]:
    """
    แปลงข้อความ Group_Report_Desc (เช่น 'สำนักงานอธิการบดี', 'คณะ/สถาบัน/สำนัก') -> Group_Report code
    ถ้า K2 ส่งข้อความมาแทน code ตรง ๆ, หรือส่งข้อความที่ไม่ตรงกับ Group_Report_ms เลย (เช่น
    'มหาวิทยาลัยศรีนครินทรวิโรฒ' ซึ่งหมายถึง "ทั้งหมด" ไม่ใช่กลุ่มใดกลุ่มหนึ่ง) จะคืน None -> ไม่กรอง (=ทั้งมหาวิทยาลัย)
    """
    if not text:
        return None
    text_norm = text.strip()
    for g in safe_fetch(fetch_group_report_list, []):
        if g.get("label", "").strip() == text_norm:
            return g.get("value")
    app.logger.info("resolve_group_report_by_label: ไม่พบกลุ่ม '%s' ใน Group_Report_ms ถือว่าไม่กรอง (ทั้งมหาวิทยาลัย)", text_norm)
    return None

def extract_plant_code(row: Dict[str, Any]) -> Optional[str]:
    """ลองหา Plant จากแถว SP_4001 — เช็คด้วย /debug/sp4001-columns ว่าคอลัมน์จริงชื่ออะไร"""
    for key in ("Plant", "PlantCode", "plant", "Plant_id"):
        if row.get(key):
            return safe_text(row.get(key))
    return None

def filter_rows_by_plant(rows: List[Dict[str, Any]], plant_filter: Optional[str], allowed_plants: Optional[Set[str]]):
    """กรอง rows ตาม Plant เดียวหรือกลุ่ม Plant ฝั่ง Python (SP ไม่รองรับกรองแบบกลุ่ม)"""
    if not plant_filter and not allowed_plants:
        return rows
    filtered = []
    for r in rows:
        plant_code = extract_plant_code(r)
        if plant_code is None:
            continue  # หา Plant ไม่ได้ -> ตัดออก ปลอดภัยกว่าเดา
        if (plant_filter and plant_code == plant_filter) or (allowed_plants and plant_code in allowed_plants):
            filtered.append(r)
    return filtered


def is_university_wide_group(group_report: Optional[str]) -> bool:
    """0 = มหาวิทยาลัย (ทั้งหมด), 1 = สนอ., 2 = คณะ — เช็ค '0' หลัก, fallback เช็ค label"""
    if group_report is None:
        return False
    if str(group_report) == "0":
        return True
    found = next((g for g in safe_fetch(fetch_group_report_list, []) if g.get("value") == group_report), None)
    return bool(found and "มหาวิทยาลัย" in (found.get("label") or ""))

def resolve_unit_display_name(dept_id, sp_plant_param, group_report, rows) -> str:
    """หัวรายงาน: ลำดับ Plant เจาะจง > Dept เจาะจง > กลุ่มที่เลือก > ว่าง (=ทั้งมหาวิทยาลัย)"""
    if sp_plant_param:
        found = next((u for u in safe_fetch(fetch_purchasing_units, []) if str(u.get("id")) == str(sp_plant_param)), None)
        if found and found.get("name"):
            return found["name"]

    if dept_id and rows:
        name = safe_text(rows[0].get("Dept_name"))
        if name != "-":
            return name

    if group_report and not is_university_wide_group(group_report):
        found = next((g for g in safe_fetch(fetch_group_report_list, []) if g.get("value") == group_report), None)
        if found and found.get("label"):
            return found["label"]

    return ""

def resolve_query_filters():
    """parse + resolve query params ที่ใช้ร่วมกันในทั้ง pdf และ xlsx route"""
    fiscal_year = to_buddhist_year(to_int_or_none(get_arg("FiscalYear", "fiscal_year", "Year")))
    # ใช้เดือนที่ผู้ใช้เลือกตรงๆ ทั้งกรองข้อมูลและตั้งชื่อเดือนในหัวรายงาน (ไม่ถอยหลัง 1 เดือนอีกต่อไป —
    # เดือน 8 ที่เลือก ต้องได้ทั้งข้อมูลของเดือน 8 และหัวรายงานเขียนว่า "เดือนสิงหาคม" ตรงกัน)
    month = parse_month(get_arg("Month", "month"))
    dept_id = get_arg("Search_Dept_ID", "DeptID", "dept_id", default="").strip() or None
    dept_name = get_arg("CUR_INDEPT_LNAME_TH", "DeptName", default="").strip() or None
    plant_id_direct = get_arg("PlantName_id", "PlantID", "Search_Plant", "Plant", default="").strip() or None
    plant_name = get_arg("PlantName", default="").strip() or None
    group_report_raw = get_arg("Group_Report", "ReportGroup", default="").strip() or None

    if dept_id and "," in dept_id:
        app.logger.warning("Search_Dept_ID เป็น comma list ('%s') ไม่รองรับ ตัดทิ้ง", dept_id)
        dept_id = None
    resolved_dept_name = None  # ชื่อหน่วยงานจริงจาก DB ที่ยืนยันแล้วว่าตรงกับ dept_id — ใช้โชว์หัวรายงานได้อย่างปลอดภัยเท่านั้น
    if not dept_id and dept_name:
        dept_id, resolved_dept_name = resolve_dept_by_name(dept_name)

    plant_filter = plant_id_direct or (resolve_plant_by_name(plant_name) if plant_name else None)
    try:
        sp_plant_param = int(plant_filter) if plant_filter else None
    except (TypeError, ValueError):
        sp_plant_param = None

    # Group_Report: ยอมรับได้ทั้ง code ตรงๆ (เช่น "1") หรือข้อความ label (เช่น "สำนักงานอธิการบดี")
    # ถ้าเป็น code ที่มีอยู่จริงใน Group_Report_ms ใช้เลย, ถ้าไม่ใช่ลองแปลงจาก label, ถ้าแปลงไม่ได้ = ไม่กรอง (ทั้งมหาวิทยาลัย)
    group_report = None
    if group_report_raw:
        known_codes = {g.get("value") for g in safe_fetch(fetch_group_report_list, [])}
        group_report = group_report_raw if group_report_raw in known_codes else resolve_group_report_by_label(group_report_raw)

    return fiscal_year, month, dept_id, sp_plant_param, group_report, resolved_dept_name

def fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report):
    """ดึงจาก SP_4001 แล้วกรองตามกลุ่ม (Group_Report) เพิ่มถ้าจำเป็น — SP กรองได้แค่ Plant เดียว"""
    rows = fetch_rows(fiscal_year, month, dept_id, sp_plant_param)
    if group_report and not sp_plant_param:
        allowed = safe_fetch(lambda: set(fetch_plant_list_by_group(group_report)), None)
        rows = filter_rows_by_plant(rows, None, allowed)
    return rows


def scoped_fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report):
    """
    ตัวห่อ fetch_filtered_rows ที่บังคับสิทธิ์ตาม session (สำคัญ: กัน user ที่ไม่ใช่ admin
    แก้ query string เอง เช่น Search_Dept_ID/Search_Plant เป็นของหน่วยงานอื่น แล้วดึงข้อมูลนอกสิทธิ์ตนเอง)
    คืนค่า (rows, error_response_or_None) — ถ้ามี error_response ให้ return ค่านั้นตรงๆ จาก route
    """
    if session.get("role") == "admin":
        return fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report), None

    allowed_depts = set(str(d) for d in (session.get("dept_ids") or []))
    allowed_plants = set(str(p) for p in (session.get("plants") or []))

    if dept_id is not None and str(dept_id) not in allowed_depts:
        return None, (jsonify({"error": "ไม่มีสิทธิ์เข้าถึงข้อมูลของหน่วยงานนี้"}), 403)
    if sp_plant_param is not None and str(sp_plant_param) not in allowed_plants:
        return None, (jsonify({"error": "ไม่มีสิทธิ์เข้าถึงข้อมูลของหน่วยจัดซื้อนี้"}), 403)

    rows = fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report)
    if dept_id is None and sp_plant_param is None:
        # ไม่ได้เจาะจงหน่วยงาน/หน่วยจัดซื้อมา -> จำกัดผลลัพธ์ให้เหลือเฉพาะที่ตนเองมีสิทธิ์เท่านั้น
        # (ไม่งั้น user ทั่วไปจะเห็นข้อมูลทั้งมหาวิทยาลัยได้ถ้าไม่ระบุ filter)
        rows = filter_rows_by_plant(rows, None, allowed_plants)
    return rows, None


def current_report_date() -> Tuple[str, str, str]:
    """
    วันที่ออกรายงาน (แถว 'วันที่ ... เดือน ... พ.ศ. ...' บนหัวเอกสาร) ใช้วันที่ปัจจุบันของเซิร์ฟเวอร์
    เสมอ ไม่รับ override จาก query string/K2 อีกต่อไป — เพราะวันที่นี้หมายถึง "วันที่พิมพ์รายงาน"
    ไม่ใช่ช่วงเวลาของข้อมูล (ช่วงเวลาของข้อมูลควบคุมด้วย FiscalYear/Month แยกต่างหาก)
    """
    today = datetime.now()
    return str(today.day), thai_month_name(today.month), str(today.year + 543)

# ===================== DATA MAPPING =====================
def map_row(row: Dict[str, Any]) -> Dict[str, str]:
    contract_date_str = format_thai_short_date(row.get("RefDocNoDate"))

    return {
        "seq": safe_text(row.get("SeqNo")),
        "work": insert_thai_break_hints(safe_text(row.get("WorkDetail"))),
        "budget": format_currency(row.get("Budget")),
        "median": format_currency(row.get("MedianPrice")),
        "method": insert_thai_break_hints(safe_text(row.get("Method_name"))),
        "bidder_list": insert_thai_break_hints(safe_text(row.get("BidderList"))),
        "winner": insert_thai_break_hints(safe_text(row.get("Winner"))),
        "reason": insert_thai_break_hints(safe_text(row.get("Reason_name"))),
        "refdoc": safe_text(row.get("RefDocNo")),
        "contract_date": contract_date_str,
        "dept_name": safe_text(row.get("Dept_name")) if row.get("Dept_name") else safe_text(row.get("Dept_id")),
    }


def render_pdf_bytes(rows, month_name, report_year, unit_name, report_day, report_month, report_be_year) -> bytes:
    """PDF ด้วย WeasyPrint (เสถียรกว่า Playwright/Chromium กับหน้าข้อมูลเยอะ)"""
    html_content = render_template(
        "pdf_report.html",
        rows=[map_row(r) for r in rows],
        month_name=month_name, report_year=report_year, unit_name=unit_name,
        report_day=report_day, report_month=report_month, report_be_year=report_be_year,
    )
    return HTML(string=html_content, base_url=str(BASE_DIR)).write_pdf()


# หัวคอลัมน์ตรงกับ pdf_report.html เป๊ะๆ — คอลัมน์สุดท้าย (เลขที่/วันที่สัญญา) merge เป็นหัวเดียว colspan=2 เหมือน PDF
# หมายเหตุ: PDF ไม่มีคอลัมน์ "หน่วยงาน" (หัวรายงานระบุหน่วยงานไว้ที่ subtitle ด้านบนแทน) จึงตัดออกจาก xlsx ด้วยเพื่อให้ตรงกัน
REPORT_COLUMN_HEADERS = [
    "ลำดับ",
    "งานที่จัดซื้อหรือจัดจ้าง",
    "วงเงินที่จะซื้อหรือจ้าง (บาท)",
    "ราคากลาง (บาท)",
    "วิธีซื้อหรือจ้าง",
    "รายชื่อผู้เสนอราคาและราคาที่เสนอ",
    "ผู้ได้รับการคัดเลือกและราคาที่ตกลงซื้อหรือจ้าง",
    "เหตุผลที่คัดเลือกโดยสรุป",
]
CONTRACT_HEADER = "เลขที่และวันที่ของสัญญาหรือข้อตกลงในการซื้อหรือจ้าง"  # merge คอลัมน์ 9-10 เหมือน <th colspan="2"> ใน PDF


def render_xlsx_bytes(
    rows: List[Dict[str, Any]],
    month_name: str,
    fiscal_year: Optional[int],
    unit_name: str,
    report_day: str,
    report_month: str,
    report_be_year: str,
) -> bytes:
    """
    สร้างไฟล์ .xlsx ด้วย openpyxl — ข้อความหัวเรื่อง/หน่วยงาน/วันที่ และหัวคอลัมน์ ใช้ประโยคเดียวกับ pdf_report.html
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "รายงาน"

    col_count = len(REPORT_COLUMN_HEADERS) + 2  # 8 หัวเดี่ยว + 2 คอลัมน์ที่ merge กันเป็นหัวเดียว = 10 คอลัมน์ข้อมูล

    title_font = Font(name="TH Sarabun New", size=16, bold=True)
    subtitle_font = Font(name="TH Sarabun New", size=14)
    header_font = Font(name="TH Sarabun New", size=14, bold=True)
    body_font = Font(name="TH Sarabun New", size=14)
    header_fill = PatternFill(start_color="EBEBEB", end_color="EBEBEB", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right_top = Alignment(horizontal="right", vertical="top")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row_idx = 1

    # เลขแบบฟอร์มมุมขวาบน — เทียบเท่า <div class="form-code">สขร.1</div> ใน pdf_report.html
    # ไม่ merge เซลล์นี้ (ต่างจากหัวเรื่อง/หน่วยงาน/วันที่) เพราะต้องการแค่ชิดขวาสุดคอลัมน์สุดท้าย
    code_cell = ws.cell(row=row_idx, column=col_count, value="สขร.1")
    code_cell.font = subtitle_font
    code_cell.alignment = Alignment(horizontal="right", vertical="center")
    row_idx += 1

    # หัวเรื่อง — ตรงกับ <div class="title"> ใน pdf_report.html
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=col_count)
    title_cell = ws.cell(row=row_idx, column=1, value=f"สรุปผลการดำเนินการจัดซื้อจัดจ้างในรอบเดือน{month_name} {fiscal_year or ''}")
    title_cell.font = title_font
    title_cell.alignment = center
    row_idx += 1

    # หน่วยงาน — ตรงกับ <div class="subtitle"> (unit_name + มหาวิทยาลัยศรีนครินทรวิโรฒ)
    subtitle = f"{unit_name} มหาวิทยาลัยศรีนครินทรวิโรฒ" if unit_name else "มหาวิทยาลัยศรีนครินทรวิโรฒ"
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=col_count)
    subtitle_cell = ws.cell(row=row_idx, column=1, value=subtitle)
    subtitle_cell.font = subtitle_font
    subtitle_cell.alignment = center
    row_idx += 1

    # วันที่ออกรายงาน — ตรงกับ <div class="date-line">
    if report_day and report_month and report_be_year:
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=col_count)
        date_cell = ws.cell(row=row_idx, column=1, value=f"วันที่ {report_day} เดือน{report_month} พ.ศ. {report_be_year}")
        date_cell.font = subtitle_font
        date_cell.alignment = center
        row_idx += 1

    row_idx += 1  # แถวว่างคั่นก่อนตาราง

    # หัวตาราง — 8 คอลัมน์เดี่ยว + 1 หัว merge คู่ (เลขที่และวันที่ของสัญญา) เหมือน PDF
    header_row = row_idx
    for col, text in enumerate(REPORT_COLUMN_HEADERS, start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    last_col = len(REPORT_COLUMN_HEADERS)
    ws.merge_cells(start_row=header_row, start_column=last_col + 1, end_row=header_row, end_column=last_col + 2)
    for col in (last_col + 1, last_col + 2):
        cell = ws.cell(row=header_row, column=col, value=CONTRACT_HEADER if col == last_col + 1 else None)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    row_idx += 1

    # แถวข้อมูล — ลำดับคอลัมน์: ลำดับ, งาน, วงเงิน, ราคากลาง, วิธี, ผู้เสนอราคา, ผู้ชนะ, เหตุผล, เลขที่สัญญา, วันที่สัญญา
    numeric_cols = {3, 4}       # วงเงิน, ราคากลาง -> ชิดขวา
    center_cols = {1, 5, 9, 10}  # ลำดับ, วิธีซื้อหรือจ้าง, เลขที่สัญญา, วันที่สัญญา -> กึ่งกลาง
    data_start_row = row_idx
    for row in rows:
        # หมายเหตุ: คู่กับ RefDocNo คือ RefDocNoDate (ไม่ใช่ ContractDate) — เหมือนที่แก้ใน map_row สำหรับ PDF
        contract_date_str = format_thai_short_date(row.get("RefDocNoDate"))
        values = [
            safe_text(row.get("SeqNo")), safe_text(row.get("WorkDetail")),
            safe_text(row.get("Budget")), safe_text(row.get("MedianPrice")), safe_text(row.get("Method_name")),
            safe_text(row.get("BidderList")), safe_text(row.get("Winner")), safe_text(row.get("Reason_name")),
            safe_text(row.get("RefDocNo")), contract_date_str,
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = body_font
            cell.border = border
            if col in numeric_cols:
                cell.alignment = right_top
            elif col in center_cols:
                cell.alignment = center
            else:
                cell.alignment = left_top
        row_idx += 1

    last_data_row = row_idx - 1

    # ความกว้างคอลัมน์โดยประมาณ (หน่วย: จำนวนตัวอักษร) — 10 คอลัมน์ตรงกับข้อมูล
    column_widths = [6, 30, 14, 14, 16, 26, 26, 20, 16, 12]
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = ws.cell(row=data_start_row, column=1).coordinate

    # ===== ตั้งค่าหน้าพิมพ์ — ให้เปิดใน Excel แล้วกด Print/Save as PDF ได้หน้าสวย ไม่ตัดคอลัมน์กลางตาราง =====
    # A4 แนวนอน (เหมือนหน้า PDF ต้นฉบับ), บีบให้พอดี 1 หน้าตามความกว้าง แต่ปล่อยความสูงอัตโนมัติ
    # (fitToHeight=0) เพื่อให้ยาวได้หลายหน้าตามจำนวนแถวข้อมูลโดยไม่บีบตัวอักษรจนอ่านไม่ออก
    last_col_letter = get_column_letter(col_count)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True
    # พื้นที่พิมพ์ครอบทั้งตาราง (รวมแถว สขร.1/หัวเรื่อง/หน่วยงาน/วันที่ ด้านบนด้วย)
    ws.print_area = f"A1:{last_col_letter}{last_data_row}"
    # แถวหัวตาราง (ลำดับ/งานที่จัดซื้อ/ฯลฯ) พิมพ์ซ้ำทุกหน้า ให้รู้ว่าแต่ละคอลัมน์คือคอลัมน์อะไรแม้ขึ้นหน้าใหม่
    ws.print_title_rows = f"{header_row}:{header_row}"

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def report_filename(month_name: str, fiscal_year, ext: str) -> str:
    return f"สรุปผลการดำเนินการจัดซื้อจัดจ้างในรอบเดือน{month_name} {fiscal_year or ''}.{ext}"


# ===================== ROUTES: K2 SSO =====================
@app.post("/api/k2/issue-token")
def api_k2_issue_token():
    """
    Server-to-server เท่านั้น — ฝั่ง K2 (workflow/backend ของ K2 เอง ไม่ใช่ browser ของ user) เรียก
    endpoint นี้ก่อนสร้างปุ่มดาวน์โหลด เพื่อขอ token ที่ผูกกับ buasri_id ของ user คนที่กำลัง login
    อยู่ใน K2 ตอนนั้น แล้วเอา token ไปฝังในลิงก์/ฟอร์มปุ่ม PDF และ Excel แทนการส่ง buasri_id เปล่าๆ

    ต้องแนบ header:  X-API-Key: <K2_API_KEY ที่ตั้งไว้ใน .env ทั้งสองฝั่ง>
    Body (JSON):     {"buasri_id": "xxxxxxx"}
    Response (JSON): {"token": "...", "expires_in": 300}
    """
    if not K2_API_KEY:
        return jsonify({"error": "เซิร์ฟเวอร์ยังไม่ได้ตั้งค่า K2_API_KEY กรุณาตั้งค่าใน .env ก่อนใช้งานจริง"}), 500
    if not hmac.compare_digest(request.headers.get("X-API-Key", ""), K2_API_KEY):
        return jsonify({"error": "unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    buasri_id = (payload.get("buasri_id") or "").strip()
    if not buasri_id:
        return jsonify({"error": "buasri_id is required"}), 400

    return jsonify({"token": issue_k2_token(buasri_id), "expires_in": K2_TOKEN_MAX_AGE_SECONDS})


@app.route("/sso", methods=["GET", "POST"])
def sso_login():
    """
    Entry point ที่ K2 ยิงมาแทนหน้า login เดิม รองรับทั้ง 2 แบบ:
      - GET  /sso?buasri_id=xxxxxxx&next=...        (ลิงก์/ปุ่มธรรมดา)
      - POST /sso  (form body: buasri_id=xxxxxxx&next=...)   (ฟอร์ม submit จากหน้า K2)
    K2 ถือว่ายืนยันตัวตนผู้ใช้มาให้แล้ว จึงไม่เช็ค LDAP ซ้ำในฝั่งนี้ — ใช้ buasri_id
    ไป lookup ผู้ใช้ในฐานข้อมูลของระบบเองเท่านั้น แล้วสร้าง session ให้ทันที
    ระบบนี้ไม่มีหน้าเว็บให้แสดงผล (ไม่มี index.html) — จุดประสงค์เดียวคือ login แล้วพาไปดาวน์โหลด
    รายงานตาม next= ทันที next= จึงเป็นค่าที่ "ต้องมี" เสมอ

    รองรับ 2 วิธีระบุตัวตน:
      1. token=  (แนะนำ/ปลอดภัย) — ได้มาจาก POST /api/k2/issue-token ฝั่ง K2 backend ก่อนสร้างลิงก์
      2. buasri_id=  (ไม่ปลอดภัย ใครก็ปลอมได้ — ใช้ได้เฉพาะตอน dev/ทดสอบในเครื่องเท่านั้น
         ต้องปิดออกก่อนขึ้น production จริงเมื่อ K2 เปลี่ยนมาใช้ token= แล้ว)
    """
    token = get_arg("token", "Token", default="")
    buasri_id = get_arg("buasri_id", "BuasriID", "BUASRI_ID", default="").strip()

    if token:
        verified_id = verify_k2_token(token)
        if not verified_id:
            return jsonify({"error": "ลิงก์หมดอายุหรือไม่ถูกต้อง กรุณากดปุ่มจากหน้า K2 ใหม่อีกครั้ง"}), 401
        buasri_id = verified_id
    elif buasri_id:
        app.logger.warning(
            "/sso ถูกเรียกด้วย buasri_id เปล่า (ไม่มี token ยืนยัน) — ใช้ได้เฉพาะตอนทดสอบ ต้องปิดก่อนขึ้น production"
        )
    else:
        return jsonify({"error": "ไม่พบ token หรือ buasri_id จากลิงก์ที่ส่งมา"}), 400

    next_path = get_arg("next", "Next", default="")
    if not next_path or not is_safe_next_path(next_path):
        return jsonify({"error": "ไม่พบ next (ปลายทางที่จะพาไปหลัง login) หรือรูปแบบไม่ถูกต้อง — ระบบนี้ไม่มีหน้าเว็บ ต้องระบุปลายทางเสมอ"}), 400

    try:
        error_response = establish_session_for_buasri_id(buasri_id)
    except Exception:
        app.logger.exception("establish_session_for_buasri_id ล้มเหลว")
        return jsonify({"error": "เข้าสู่ระบบไม่สำเร็จ กรุณาลองใหม่อีกครั้ง หรือติดต่อผู้ดูแลระบบ"}), 500
    if error_response is not None:
        return error_response

    return redirect(next_path)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/procurement-report.pdf")
@login_required
def procurement_report_pdf():
    fiscal_year, month, dept_id, sp_plant_param, group_report, resolved_dept_name = resolve_query_filters()
    # unit_name ที่โชว์บนหัวรายงาน — ลำดับความน่าเชื่อถือ:
    #   1) resolved_dept_name: ชื่อจาก DB ที่ยืนยันแล้วว่าตรงกับ dept_id จริง (ปลอดภัยสุด)
    #   2) UnitName/unit: ค่า override ที่ตั้งใจส่งมาตรง ๆ (ไม่ใช่ CUR_INDEPT_LNAME_TH/DeptName ที่อาจ resolve ไม่ตรง)
    #   3) resolve_unit_display_name(...): คำนวณจากข้อมูลที่กรองได้จริงหลัง fetch
    #   4) session dept_name: ชื่อหน่วยงานจริงของ user ที่ login อยู่ (fallback สุดท้าย)
    # ห้ามเอาข้อความดิบจาก CUR_INDEPT_LNAME_TH/DeptName มาโชว์ตรง ๆ ถ้า resolve ไม่เจอ เพราะจะโชว์
    # ชื่อหน่วยงานที่ไม่มีจริง ทั้งที่ข้อมูลในตารางถูกกรองด้วยหน่วยงานอื่น (ของ user ที่ login อยู่)
    unit_name = resolved_dept_name or get_arg("UnitName", "unit", default="").strip()
    # เดือนในหัวรายงาน = เดือนเดียวกับที่ใช้กรองข้อมูลเป๊ะๆ (เลือกเดือน 8 -> ได้ทั้งข้อมูลและหัวรายงานของเดือน 8)
    month_name = get_arg("MonthName", "month_name", default="") or thai_month_name(month)
    # วันที่ออกรายงานใช้วันที่ปัจจุบันเสมอ ไม่รับ override จาก query string
    report_day, report_month, report_be_year = current_report_date()

    rows, scope_error = scoped_fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report)
    if scope_error is not None:
        return scope_error
    if not unit_name:
        unit_name = resolve_unit_display_name(dept_id, sp_plant_param, group_report, rows)
    if not unit_name:
        unit_name = session.get("dept_name", "").strip()

    pdf_bytes = render_pdf_bytes(
        rows, month_name, str(fiscal_year or ""), unit_name or "",
        report_day, report_month, report_be_year,
    )
    return Response(
        pdf_bytes, mimetype="application/pdf",
        headers={"Content-Disposition": content_disposition_header(report_filename(month_name, fiscal_year, "pdf"))},
    )

@app.get("/procurement-report.xlsx")
@login_required
def procurement_report_xlsx():
    fiscal_year, month, dept_id, sp_plant_param, group_report, resolved_dept_name = resolve_query_filters()
    unit_name = resolved_dept_name or get_arg("UnitName", "unit", default="").strip()
    # เดือนในหัวรายงาน = เดือนเดียวกับที่ใช้กรองข้อมูลเป๊ะๆ — เหมือน PDF
    month_name = get_arg("MonthName", "month_name", default="") or thai_month_name(month)
    # วันที่ออกรายงานใช้วันที่ปัจจุบันเสมอ ไม่รับ override จาก query string — เหมือน PDF
    report_day, report_month, report_be_year = current_report_date()

    rows, scope_error = scoped_fetch_filtered_rows(fiscal_year, month, dept_id, sp_plant_param, group_report)
    if scope_error is not None:
        return scope_error
    if not unit_name:
        unit_name = resolve_unit_display_name(dept_id, sp_plant_param, group_report, rows)
    if not unit_name:
        unit_name = session.get("dept_name", "").strip()

    xlsx_bytes = render_xlsx_bytes(
        rows, month_name, fiscal_year, unit_name or "",
        report_day, report_month, report_be_year,
    )
    return Response(
        xlsx_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": content_disposition_header(report_filename(month_name, fiscal_year, "xlsx"))},
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False)