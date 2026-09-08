from __future__ import annotations

import hmac

from flask import Blueprint, Response, current_app, jsonify, redirect, request, session

from auth import establish_session_for_buasri_id, issue_k2_token, login_required, verify_k2_token
from config import K2_API_KEY, K2_TOKEN_MAX_AGE_SECONDS
from helpers import content_disposition_header, get_arg, is_safe_next_path, thai_month_name
from reports import (
    current_report_date,
    render_pdf_bytes,
    render_xlsx_bytes,
    report_filename,
    resolve_query_filters,
    resolve_unit_display_name,
    scoped_fetch_filtered_rows,
)

bp = Blueprint("report", __name__)


# ===================== K2 SSO =====================
@bp.post("/api/k2/issue-token")
def api_k2_issue_token():
    
    if not K2_API_KEY:
        return jsonify({"error": "เซิร์ฟเวอร์ยังไม่ได้ตั้งค่า K2_API_KEY กรุณาตั้งค่าใน .env ก่อนใช้งานจริง"}), 500
    if not hmac.compare_digest(request.headers.get("X-API-Key", ""), K2_API_KEY):
        return jsonify({"error": "unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    buasri_id = (payload.get("buasri_id") or "").strip()
    if not buasri_id:
        return jsonify({"error": "buasri_id is required"}), 400

    return jsonify({"token": issue_k2_token(buasri_id), "expires_in": K2_TOKEN_MAX_AGE_SECONDS})


@bp.route("/sso", methods=["GET", "POST"])
def sso_login():
    
    token = get_arg("token", "Token", default="")
    buasri_id = get_arg("buasri_id", "BuasriID", "BUASRI_ID", default="").strip()

    if token:
        verified_id = verify_k2_token(token)
        if not verified_id:
            return jsonify({"error": "ลิงก์หมดอายุหรือไม่ถูกต้อง กรุณากดปุ่มจากหน้า K2 ใหม่อีกครั้ง"}), 401
        buasri_id = verified_id
    elif buasri_id:
        current_app.logger.warning(
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
        current_app.logger.exception("establish_session_for_buasri_id ล้มเหลว")
        return jsonify({"error": "เข้าสู่ระบบไม่สำเร็จ กรุณาลองใหม่อีกครั้ง หรือติดต่อผู้ดูแลระบบ"}), 500
    if error_response is not None:
        return error_response

    return redirect(next_path)


@bp.get("/health")
def health():
    return {"status": "ok"}


# ===================== REPORT DOWNLOADS =====================
@bp.get("/procurement-report.pdf")
@login_required
def procurement_report_pdf():
    fiscal_year, month, dept_id, sp_plant_param, group_report, resolved_dept_name = resolve_query_filters()

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


@bp.get("/procurement-report.xlsx")
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