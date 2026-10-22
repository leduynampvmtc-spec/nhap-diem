"""Đăng nhập cục bộ (khoá mềm). Mật khẩu không lưu ở dạng chữ — chỉ lưu băm PBKDF2-SHA256."""
from __future__ import annotations

import hashlib
import hmac

_SALT = b"nhapdiem-pvc-2026"
_ITER = 200_000
_USER = "leduynam"
_HASH = "ab4a4103f4d389a223ace605e3942e9e64751f574694b6624dc4d5acdf35a6e1"


def _hash(pw: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), _SALT, _ITER).hex()


def check_login(user: str, password: str) -> bool:
    ok_user = hmac.compare_digest((user or "").strip().lower().encode(), _USER.encode())
    ok_pw = hmac.compare_digest(_hash(password or "").encode(), _HASH.encode())
    return ok_user and ok_pw
