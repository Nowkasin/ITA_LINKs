from __future__ import annotations

from flask import Flask

import config
from routes import bp

app = Flask(__name__)

# SESSION_SECRET_KEY ใน .env จำเป็นสำหรับเซ็น cookie — สร้างด้วย:
# python -c "import secrets; print(secrets.token_hex(32))"
app.secret_key = config.SESSION_SECRET_KEY
if not app.secret_key:
    raise RuntimeError("ไม่พบ SESSION_SECRET_KEY ใน .env กรุณาตั้งค่าก่อนรันระบบ")

# ความปลอดภัยของ session cookie — SESSION_COOKIE_SECURE ควรเป็น true เสมอเมื่อขึ้น production จริง
# (รันหลัง HTTPS แล้ว) แต่ปล่อย false ไว้ตอน dev/ทดสอบในเครื่องผ่าน http://127.0.0.1 เพราะ Secure
# cookie จะไม่ถูกส่งเลยถ้าไม่ใช่ HTTPS ทำให้ session ใช้งานไม่ได้ตอนทดสอบในเครื่อง
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = config.SESSION_COOKIE_SECURE

app.register_blueprint(bp)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False)