from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent

# ===================== SESSION =====================
# SESSION_SECRET_KEY ใน .env จำเป็นสำหรับเซ็น cookie — สร้างด้วย:
# python -c "import secrets; print(secrets.token_hex(32))"
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY", "")

SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "false").lower() == "true"

# ===================== K2 SSO TOKEN =====================
# K2_TOKEN_SECRET: แนะนำให้ตั้งแยกจาก SESSION_SECRET_KEY (คนละ key คนละหน้าที่) — สร้างด้วย:

K2_TOKEN_SECRET = os.getenv("K2_TOKEN_SECRET", SESSION_SECRET_KEY)
# K2_API_KEY: secret key เฉพาะสำหรับให้ "เซิร์ฟเวอร์ K2" (ไม่ใช่ browser user) เรียกขอ token — ต้องตั้งใน .env ก่อนใช้จริง
K2_API_KEY = os.getenv("K2_API_KEY", "")
K2_TOKEN_MAX_AGE_SECONDS = int(os.getenv("K2_TOKEN_MAX_AGE_SECONDS", "300"))  # token ใช้ได้ 5 นาทีหลังออก

# ===================== DATABASE =====================
DB_CONFIG = {
    "server": os.getenv("DB_SERVER"),
    "port": int(os.getenv("DB_PORT", "1433")),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": os.getenv("DB_NAME"),
    "charset": "UTF-8",
}

SP_NAME = "dbo.SP_4001_Ita_Procurement"