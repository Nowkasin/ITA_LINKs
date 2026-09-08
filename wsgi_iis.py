from __future__ import annotations

import os

from waitress import serve

from app import app

if __name__ == "__main__":
    port = int(os.environ.get("HTTP_PLATFORM_PORT", "5001"))
    # host ต้องเป็น 127.0.0.1 เท่านั้น — IIS เป็นตัวรับ request จากภายนอกแล้ว reverse-proxy
    # มาที่ process นี้ในเครื่องเดียวกัน ไม่ต้องเปิดรับ request ตรงจากภายนอกอีกชั้น
    serve(app, host="127.0.0.1", port=port)