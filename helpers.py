from __future__ import annotations

from typing import Any, Optional
from urllib.parse import quote, urlparse

from flask import request

try:
    from pythainlp.tokenize import word_tokenize
except ImportError:
    word_tokenize = None

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


def safe_text(value: Any) -> str:
    text = str(value).strip() if value is not None else ""
    return text if text else "-"


def safe_fetch(fn, default):
    try:
        return fn()
    except Exception:
        return default


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


def content_disposition_header(filename: str) -> str:
    """HTTP header เป็น Latin-1 เท่านั้น ชื่อไฟล์ไทยต้อง encode แบบ RFC 5987 (filename*=UTF-8''...)"""
    return f'attachment; filename="download"; filename*=UTF-8\'\'{quote(filename, safe="")}'


def is_safe_next_path(path: str) -> bool:
    """
    กัน open redirect: next ต้องเป็น path ภายในระบบเดียวกันเท่านั้น (ขึ้นต้นด้วย '/' ตัวเดียว,
    ไม่ใช่ '//evil.com', ไม่มี scheme/host แปลกปน) ไม่งั้นใครก็ปลอมลิงก์ /sso?...&next=https://phishing.example ได้
    """
    if not path or not path.startswith("/") or path.startswith("//"):
        return False
    parsed = urlparse(path)
    return not parsed.netloc and not parsed.scheme