from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Dict, List, Optional

import pymssql

from config import DB_CONFIG, SP_NAME
from helpers import safe_text


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