"""
Giao diện web chạy cục bộ (127.0.0.1) — mở bằng trình duyệt, giống cách app "Valve Test Calculator" hoạt động.

* Chỉ nghe trên 127.0.0.1, kiểm tra Host/Origin, dùng cookie phiên ngẫu nhiên sau khi đăng nhập.
* Tệp tải lên được lưu vào thư mục tạm; file xuất ghi vào thư mục người dùng chọn rồi cho mở/tải về.
* Chỉ dùng thư viện chuẩn (http.server) nên đóng gói PyInstaller rất gọn, chạy giống nhau trên Windows và macOS.
"""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import socket
import sys
import tempfile
import threading
import time
import urllib.parse
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Dict, Optional

from . import auth, core
from . import controller as ctl
from . import forms_pdf as F
from .osutil import open_file, open_in_file_manager
from .footer import FooterOptions
from .pdfgrid import ensure_fonts, resource_path

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".nhapdiem_web_config.json")
DEFAULT_OUT = os.path.join(os.path.expanduser("~"), "Downloads", "NhapDiem_xuat")
MAX_UPLOAD = 40 * 1024 * 1024
SAVED_KEYS = {"out_dir", "tbm", "ma_lhp", "dia_diem", "so_ca", "tg_ca", "tg_tt", "quy_che", "ht_lt", "dd_th",
              "ca_thi", "ngay", "fmt", "ft_dk", "ft_tl", "ft_lt", "ft_tk", "ft_lbh", "ft_nhl", "lan_tk", "ft_lbh_dk", "ft_lbh_tl", "ft_lbh_lt", "ft_lbh_tk", "ft_nhl_dk", "ft_nhl_tl", "ft_nhl_lt", "ft_nhl_tk", "tn_may", "tn_giay", "van_dap", "tu_luan", "ht_khac", "drop", "last_user"}


def _log(msg: str) -> None:
    try:
        if sys.stdout:
            print(msg, flush=True)
    except Exception:
        pass


class AppState:
    def __init__(self) -> None:
        self.tmp = tempfile.mkdtemp(prefix="nhapdiem_")
        self.token = ""                       # cookie phiên (rỗng = chưa đăng nhập)
        self.user = ""
        self.paths: Dict[str, str] = {}       # input / att -> đường dẫn tệp tải lên
        self.names: Dict[str, str] = {}       # tên gốc để hiển thị
        self.files: Dict[str, str] = {}       # id tải về -> đường dẫn file đã xuất
        self.font = ""
        self.lock = threading.Lock()
        self.server: Optional[ThreadingHTTPServer] = None
        self.last_ping = 0.0
        self.pinged = False

    # --- cấu hình ---
    def load_config(self) -> dict:
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            return {k: v for k, v in d.items() if k in SAVED_KEYS}
        except Exception:
            return {}

    def save_config(self, upd: dict) -> None:
        d = self.load_config()
        d.update({k: v for k, v in upd.items() if k in SAVED_KEYS})
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False, indent=1)
        except Exception:
            pass


STATE = AppState()


def _s(opts: dict, key: str, default: str = "") -> str:
    v = opts.get(key, default)
    return default if v is None else str(v).strip()


def _b(opts: dict, key: str, default: bool = False) -> bool:
    v = opts.get(key, default)
    return bool(v)


def student_rows(sess: ctl.Session) -> list:
    rows = []
    for st in sess.students:
        rows.append({
            "stt": st.stt, "name": st.name, "mssv": st.mssv or "",
            "o": core.fmt_score(st.o), "s": core.fmt_score(st.s), "x": st.x or "",
            "vang": core.fmt_pct(st.vang_tong),
        })
    return rows


def meta_dict(sess: ctl.Session) -> dict:
    m = sess.meta
    return {"loai": m.loai, "ten": m.ten, "lop": m.lop, "khoa": m.khoa, "lop_khoa": m.lop_khoa,
            "tong_gio": m.tong_gio, "gio_lt": m.gio_lt, "gio_th": m.gio_th, "he": m.he,
            "la_mo_dun": m.la_mo_dun, "n": len(sess.students)}


def do_load(opts: dict, use_att: bool, drop: bool, lan: int = 1) -> ctl.Session:
    inp = STATE.paths.get("input", "")
    att = STATE.paths.get("att", "") if use_att else ""
    return ctl.load_session(inp, att, drop_unmatched=drop, lan=lan)


def do_export(kind: str, opts: dict) -> dict:
    out_dir = _s(opts, "out_dir") or DEFAULT_OUT
    ngay = ctl.parse_date(_s(opts, "ngay"))
    tbm = _s(opts, "tbm", "CƠ KHÍ").upper()
    ma = _s(opts, "ma_lhp")
    extra = ""
    fmt = "xlsx" if _s(opts, "fmt", "pdf").lower() == "xlsx" else "pdf"
    ft = FooterOptions.from_opts(opts)
    if kind == "dk":
        sess = do_load(opts, True, _b(opts, "drop", True))
        o = F.DKOptions(
            ma_lhp=ma, tn_may=_b(opts, "tn_may", True), tn_giay=_b(opts, "tn_giay"), van_dap=_b(opts, "van_dap"),
            tu_luan=_b(opts, "tu_luan"), ht_khac=_b(opts, "ht_khac"), dia_diem=_s(opts, "dia_diem"),
            so_hssv_ca=_s(opts, "so_ca", "-"), tg_ca=_s(opts, "tg_ca", "-"), tg_tap_trung=_s(opts, "tg_tt", "-"),
            quy_che=_s(opts, "quy_che", "2022"), to_bo_mon=tbm, ngay_lap=ngay)
        path = ctl.export_dieu_kien(sess, out_dir, o, fmt, ft)
    elif kind == "tl":
        sess = do_load(opts, False, False)
        o = F.TLOptions(ma_lhp=ma, hinh_thuc_lt=_s(opts, "ht_lt"), dia_diem_th=_s(opts, "dd_th"),
                        to_bo_mon=tbm, ngay_lap=ngay)
        path = ctl.export_thi_lai(sess, out_dir, o, fmt, ft)
    elif kind == "lt":
        sess = do_load(opts, False, False)
        lan = 2 if str(opts.get("lan", 1)) == "2" else 1
        o = F.LTOptions(lan_thi=lan, ca_thi=_s(opts, "ca_thi"), to_bo_mon=tbm, ngay_lap=ngay)
        path = ctl.export_diem_lt(sess, out_dir, o, fmt, ft)
    elif kind == "xtt":
        sess = do_load(opts, False, False)
        path = ctl.export_xep_thu_tu(sess, out_dir, ngay, fmt, ft)
    elif kind in ("tong1", "tong2"):
        lan = 2 if kind == "tong2" else 1
        sess = do_load(opts, False, False, lan)
        path, n = ctl.export_tong_ket(STATE.paths.get("input", ""), sess, out_dir, ngay, lan)
        extra = f"{n} HSSV"
    elif kind == "qt":
        sess = do_load(opts, True, _b(opts, "drop", True))
        path, ws_ = ctl.export_nhap_qua_trinh(sess, STATE.paths.get("tpl_qt", ""), out_dir, ngay, _s(opts, "quy_che", "2022"))
        sess.warnings = list(sess.warnings) + ws_
    elif kind == "kt":
        sess = do_load(opts, False, False)
        lan = 2 if str(opts.get("lan_kt", 1)) == "2" else 1
        path, ws_ = ctl.export_nhap_ket_thuc(sess, STATE.paths.get("tpl_kt", ""), out_dir, ngay, lan)
        sess.warnings = list(sess.warnings) + ws_
    elif kind == "phieu":
        sess = do_load(opts, False, False)
        bai = int(opts.get("bai", 1))
        path, n = ctl.export_phieu(sess, out_dir, bai, ngay)
        extra = f"{n} HSSV có điểm"
    else:
        raise ctl.AppError("Chức năng không hợp lệ.")
    fid = secrets.token_hex(8)
    with STATE.lock:
        STATE.files[fid] = path
    return {"ok": True, "id": fid, "filename": os.path.basename(path), "path": path, "note": extra,
            "warnings": sess.warnings, "font": STATE.font}


# ---------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    server_version = "NhapDiem/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args) -> None:  # im lặng
        pass

    # --- tiện ích ---
    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").split(":")[0].lower()
        if host not in ("127.0.0.1", "localhost"):
            return False
        origin = self.headers.get("Origin")
        if origin:
            o = urllib.parse.urlparse(origin)
            if (o.hostname or "").lower() not in ("127.0.0.1", "localhost") or o.port != self.server.server_address[1]:
                return False
        return True

    def _cookie(self) -> str:
        for part in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = part.strip().partition("=")
            if k == "nd_session":
                return v
        return ""

    def _authed(self) -> bool:
        return bool(STATE.token) and secrets.compare_digest(self._cookie(), STATE.token)

    def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8",
              headers: Optional[dict] = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj: dict, code: int = 200, headers: Optional[dict] = None) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"), headers=headers)

    def _body(self) -> bytes:
        return getattr(self, "_raw", b"")

    def _read_json(self) -> dict:
        raw = self._body()
        try:
            return json.loads(raw.decode("utf-8")) if raw else {}
        except Exception:
            raise ctl.AppError("Dữ liệu gửi lên không hợp lệ.")

    # --- GET ---
    def do_GET(self) -> None:  # noqa: N802
        if not self._host_ok():
            return self._send(HTTPStatus.FORBIDDEN, b"Forbidden", "text/plain")
        u = urllib.parse.urlparse(self.path)
        path = u.path
        try:
            if path in ("/", "/index.html"):
                with open(resource_path("assets", "index.html"), "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            if path in ("/logo.png", "/favicon.png"):
                with open(resource_path("assets", "logo_240.png"), "rb") as f:
                    return self._send(200, f.read(), "image/png", {"Cache-Control": "max-age=86400"})
            if path == "/logo72.png":
                with open(resource_path("assets", "logo_72.png"), "rb") as f:
                    return self._send(200, f.read(), "image/png", {"Cache-Control": "max-age=86400"})
            if path == "/api/ping":
                STATE.last_ping = time.time()
                STATE.pinged = True
                return self._json({"ok": True})
            if path == "/api/me":
                cfg = STATE.load_config()
                return self._json({"authed": self._authed(), "user": STATE.user if self._authed() else "",
                                   "last_user": cfg.get("last_user", "")})
            if not self._authed():
                return self._json({"ok": False, "error": "Chưa đăng nhập."}, 401)
            if path == "/api/config":
                cfg = STATE.load_config()
                cfg.setdefault("out_dir", DEFAULT_OUT)
                return self._json({"config": cfg, "font": STATE.font,
                                   "files": {k: STATE.names.get(k, "") for k in ("input", "att")}})
            m = re.match(r"^/download/([0-9a-f]{16})$", path)
            if m:
                p = STATE.files.get(m.group(1))
                if not p or not os.path.isfile(p):
                    return self._send(404, b"Not found", "text/plain")
                q = urllib.parse.parse_qs(u.query)
                inline = q.get("inline", ["0"])[0] == "1"
                ctype = "application/pdf" if p.lower().endswith(".pdf") else \
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                disp = "inline" if inline and p.lower().endswith(".pdf") else "attachment"
                fn = urllib.parse.quote(os.path.basename(p))
                with open(p, "rb") as f:
                    data = f.read()
                return self._send(200, data, ctype, {"Content-Disposition": f"{disp}; filename*=UTF-8''{fn}"})
            return self._send(404, b"Not found", "text/plain")
        except ctl.AppError as e:
            return self._json({"ok": False, "error": str(e)}, 400)
        except Exception as e:  # lỗi không lường trước
            return self._json({"ok": False, "error": f"Lỗi: {e}"}, 500)

    # --- POST ---
    def do_POST(self) -> None:  # noqa: N802
        # đọc hết thân yêu cầu ngay từ đầu để kết nối keep-alive không bị lệch khi trả lỗi sớm
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            n = 0
        if n > MAX_UPLOAD:
            self.close_connection = True
            return self._json({"ok": False, "error": "Tệp quá lớn (tối đa 40 MB)."}, 413, {"Connection": "close"})
        self._raw = self.rfile.read(n) if n else b""
        if not self._host_ok():
            return self._send(HTTPStatus.FORBIDDEN, b"Forbidden", "text/plain")
        u = urllib.parse.urlparse(self.path)
        path = u.path
        try:
            if path == "/api/login":
                d = self._read_json()
                if auth.check_login(str(d.get("user", "")), str(d.get("password", ""))):
                    with STATE.lock:
                        STATE.token = secrets.token_urlsafe(24)
                        STATE.user = str(d.get("user", "")).strip()
                    STATE.save_config({"last_user": STATE.user})
                    return self._json({"ok": True, "user": STATE.user},
                                      headers={"Set-Cookie": f"nd_session={STATE.token}; Path=/; HttpOnly; SameSite=Strict"})
                return self._json({"ok": False, "error": "Sai tài khoản hoặc mật khẩu."}, 401)
            if not self._authed():
                return self._json({"ok": False, "error": "Chưa đăng nhập."}, 401)
            if path == "/api/logout":
                with STATE.lock:
                    STATE.token = ""
                    STATE.user = ""
                return self._json({"ok": True}, headers={"Set-Cookie": "nd_session=; Path=/; Max-Age=0"})
            if path == "/api/config":
                STATE.save_config(self._read_json())
                return self._json({"ok": True})
            if path == "/api/upload":
                q = urllib.parse.parse_qs(u.query)
                kind = q.get("kind", [""])[0]
                if kind not in ("input", "att", "tpl_qt", "tpl_kt"):
                    raise ctl.AppError("Loại tệp không hợp lệ.")
                name = os.path.basename(urllib.parse.unquote(q.get("name", ["file.xlsx"])[0])) or "file.xlsx"
                ok_ext = (".xlsx", ".xlsm", ".xls") if kind.startswith("tpl_") else (".xlsx", ".xlsm")
                if not name.lower().endswith(ok_ext):
                    raise ctl.AppError("Chỉ nhận tệp Excel (" + ", ".join(ok_ext) + ").")
                data = self._body()
                if not data:
                    raise ctl.AppError("Tệp rỗng.")
                dest = os.path.join(STATE.tmp, f"{kind}{os.path.splitext(name)[1].lower()}")
                with open(dest, "wb") as f:
                    f.write(data)
                with STATE.lock:
                    STATE.paths[kind] = dest
                    STATE.names[kind] = name
                return self._json({"ok": True, "name": name})
            if path == "/api/clear":
                kind = self._read_json().get("kind", "")
                with STATE.lock:
                    STATE.paths.pop(kind, None)
                    STATE.names.pop(kind, None)
                return self._json({"ok": True})
            if path == "/api/load":
                d = self._read_json()
                sess = do_load(d, bool(STATE.paths.get("att")), bool(d.get("drop", False)))
                return self._json({"ok": True, "meta": meta_dict(sess), "warnings": sess.warnings,
                                   "students": student_rows(sess)})
            m = re.match(r"^/api/export/(tong1|tong2|dk|tl|lt|xtt|phieu|qt|kt)$", path)
            if m:
                d = self._read_json()
                STATE.save_config({k: v for k, v in d.items() if k in SAVED_KEYS})
                return self._json(do_export(m.group(1), d))
            if path == "/api/tinh":
                d = self._read_json()
                lan = 2 if str(d.get("lan_tk", 1)) == "2" else 1
                sess = do_load(d, False, False, lan)
                out = ctl.bang_tong_ket(sess)
                out.update({"ok": True, "meta": meta_dict(sess), "lan": lan,
                            "chua_l2": sum(1 for r in out["rows"] if lan == 2 and r["t"] == "")})
                return self._json(out)
            if path == "/api/open":
                d = self._read_json()
                p = STATE.files.get(str(d.get("id", "")))
                if not p or not os.path.exists(p):
                    raise ctl.AppError("Không tìm thấy file.")
                (open_in_file_manager if d.get("what") == "folder" else open_file)(p)
                return self._json({"ok": True})
            if path == "/api/quit":
                self._json({"ok": True})
                threading.Thread(target=STATE.server.shutdown, daemon=True).start()
                return
            return self._send(404, b"Not found", "text/plain")
        except ctl.AppError as e:
            return self._json({"ok": False, "error": str(e)}, 400)
        except ctl.Cancelled:
            return self._json({"ok": False, "error": "Đã hủy."}, 400)
        except PermissionError:
            return self._json({"ok": False, "error": "Không ghi được file (đang mở trong chương trình khác? "
                                                     "hoặc không có quyền ghi vào thư mục lưu). Hãy đóng file rồi thử lại."}, 400)
        except OSError as e:
            return self._json({"ok": False, "error": f"Lỗi ghi file: {e}"}, 400)
        except Exception as e:
            return self._json({"ok": False, "error": f"Lỗi: {e}"}, 500)


class QuietServer(ThreadingHTTPServer):
    def handle_error(self, request, client_address) -> None:   # bỏ qua lỗi ngắt kết nối của trình duyệt
        pass


def _free_port(start: int = 47610, tries: int = 40) -> int:
    for p in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    raise RuntimeError("Không tìm được cổng trống để chạy ứng dụng.")


def make_server(port: int = 0) -> ThreadingHTTPServer:
    port = port or _free_port()
    srv = QuietServer(("127.0.0.1", port), Handler)
    srv.daemon_threads = True
    STATE.server = srv
    try:
        STATE.font = ensure_fonts()
    except Exception as e:
        STATE.font = f"Lỗi font: {e}"
    return srv


def main(open_browser: bool = True) -> None:
    srv = make_server()
    url = f"http://127.0.0.1:{srv.server_address[1]}/"
    _log(f"Xuất biểu mẫu từ Form nhập điểm — mở {url} (đóng cửa sổ này hoặc bấm Thoát để dừng)")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    def watchdog() -> None:   # tự thoát khi đã đóng hết tab trình duyệt
        while True:
            time.sleep(10)
            if STATE.pinged and time.time() - STATE.last_ping > 150:
                srv.shutdown()
                return
    threading.Thread(target=watchdog, daemon=True).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()
        shutil.rmtree(STATE.tmp, ignore_errors=True)
