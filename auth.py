from __future__ import annotations

from functools import wraps
from typing import Optional

from flask import jsonify, session
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired

from config import K2_TOKEN_SECRET, K2_TOKEN_MAX_AGE_SECONDS
from db import (
    check_buasri_id,
    get_or_create_user,
    get_plant_by_user,
    is_resource_admin_dept,
    save_dept_access,
)
from helpers import safe_fetch, safe_text

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