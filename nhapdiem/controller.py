"""
Lớp điều phối (không phụ thuộc giao diện): đọc file nhập điểm, tính điểm, xuất file.
GUI chỉ gọi các hàm ở đây nên có thể kiểm thử tự động mà không cần mở cửa sổ.
"""
from __future__ import annotations

import datetime as _dt
import os
import random
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from . import core, forms_import, forms_pdf as F, phieu_xlsx, xlsx_forms as X
from .footer import FooterOptions


class AppError(Exception):
    """Lỗi có thông điệp thân thiện để hiển thị cho người dùng."""


class Cancelled(Exception):
    """Người dùng chủ động hủy (vd. không ghi đè file cũ) — không phải lỗi."""


@dataclass
class Session:
    path: str
    meta: core.Meta
    students: List[core.Student]
    warnings: List[str] = field(default_factory=list)


def parse_date(text: str) -> _dt.date:
    t = (text or "").strip()
    if not t:
        return _dt.date.today()
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$", t)
    if not m:
        raise AppError("Ngày lập không đúng định dạng — nhập dạng dd/mm/yyyy (ví dụ 21/10/2026).")
    d, mo, y = map(int, m.groups())
    try:
        return _dt.date(y, mo, d)
    except ValueError:
        raise AppError("Ngày lập không hợp lệ.")


def load_session(path: str, att_path: str = "", drop_unmatched: bool = False) -> Session:
    if not path or not os.path.isfile(path):
        raise AppError("Chưa chọn file nhập điểm (Form nhập điểm .xlsx) hoặc file không tồn tại.")
    try:
        meta, students, warns = core.read_input(path)
    except Exception as e:  # file hỏng / không phải xlsx
        raise AppError(f"Không đọc được file nhập điểm:\n{e}")
    if not students:
        raise AppError("Không tìm thấy HSSV nào trong sheet DIEM TK (cột B từ hàng 10).")
    core.compute_all(students, meta)
    if att_path:
        if not os.path.isfile(att_path):
            raise AppError("File điểm danh không tồn tại.")
        try:
            att = core.load_attendance(att_path, names=[s.name for s in students])
        except Exception as e:
            raise AppError(f"Không đọc được file điểm danh:\n{e}")
        missing = core.apply_attendance(students, att)
        if missing and drop_unmatched:
            miss = set(missing)
            students = [s for s in students if s.name not in miss]
            for i, s in enumerate(students, 1):
                s.stt = i
            if not students:
                raise AppError("Không còn HSSV nào khớp với file điểm danh.")
            warns.append(f"Đã loại {len(miss)} HSSV không có trong file điểm danh: " + ", ".join(missing))
        elif missing:
            warns.append("Không khớp tên trong file điểm danh: " + ", ".join(missing[:6]) + (" ..." if len(missing) > 6 else ""))
    return Session(path, meta, students, warns)


def _target(out_dir: str, filename: str) -> str:
    if not out_dir:
        raise AppError("Chưa chọn thư mục lưu file.")
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as e:
        raise AppError(f"Không tạo được thư mục lưu:\n{e}")
    return os.path.join(out_dir, filename)


def target_path(out_dir: str, sess: Session, prefix: str, ngay: _dt.date, ext: str, extra: str = "") -> str:
    return _target(out_dir, core.out_filename(prefix, sess.meta, ngay, ext, extra))


# ---------------------------------------------------------------------------
def _fmt(fmt: str) -> str:
    return "xlsx" if str(fmt).lower() == "xlsx" else "pdf"


def export_dieu_kien(sess: Session, out_dir: str, opt: F.DKOptions, fmt: str = "pdf", ft=None) -> str:
    ft = ft or FooterOptions()
    fmt = _fmt(fmt)
    path = target_path(out_dir, sess, "DSDT", opt.ngay_lap, fmt)
    if fmt == "xlsx":
        X.xlsx_dieu_kien(path, sess.meta, sess.students, opt, ft)
    else:
        F.pdf_dieu_kien_du_thi(path, sess.meta, sess.students, opt, ft)
    return path


def export_thi_lai(sess: Session, out_dir: str, opt: F.TLOptions, fmt: str = "pdf", ft=None) -> str:
    ft = ft or FooterOptions()
    fmt = _fmt(fmt)
    ds = core.danh_sach_thi_lai(sess.students, sess.meta, opt.hinh_thuc_lt, opt.text_th)
    if not ds:
        raise AppError("Không có HSSV nào thuộc diện thi lại (cột Đánh giá không có \"Thi lại\").")
    path = target_path(out_dir, sess, "DSTL", opt.ngay_lap, fmt)
    (X.xlsx_thi_lai if fmt == "xlsx" else F.pdf_thi_lai)(path, sess.meta, ds, opt, ft)
    return path


def export_diem_lt(sess: Session, out_dir: str, opt: F.LTOptions, fmt: str = "pdf", ft=None) -> str:
    ft = ft or FooterOptions()
    fmt = _fmt(fmt)
    if not sess.meta.la_mo_dun:
        raise AppError("File điểm thi LT chỉ xuất cho MÔ ĐUN.\n"
                       f"File hiện tại là \"{sess.meta.loai or 'không rõ'}\" ({sess.meta.ten}). "
                       "Hãy chọn file Form nhập điểm của mô đun (ô A5 ghi \"TÊN MÔ ĐUN: …\").")
    ds = core.bang_diem_lt(sess.students, opt.lan_thi)
    if not ds:
        raise AppError("Không có HSSV nào để xuất (lần thi 2 chỉ gồm HSSV có điểm thi LT lần 1 < 5).")
    path = target_path(out_dir, sess, "LT", opt.ngay_lap, fmt, "L2" if opt.lan_thi == 2 else "")
    (X.xlsx_diem_lt if fmt == "xlsx" else F.pdf_diem_lt)(path, sess.meta, ds, opt, ft)
    return path


def export_xep_thu_tu(sess: Session, out_dir: str, ngay: _dt.date, fmt: str = "pdf", ft=None) -> str:
    ft = ft or FooterOptions()
    fmt = _fmt(fmt)
    data = core.xep_thu_tu(sess.students)
    if not any(data.values()):
        raise AppError("Chưa có điểm tổng kết để xếp thứ tự.")
    path = target_path(out_dir, sess, "XTT", ngay, fmt)
    if fmt == "xlsx":
        X.xlsx_xep_thu_tu(path, sess.meta, data, ft)
    else:
        F.pdf_xep_thu_tu(path, sess.meta, data, ngay, ft)
    return path


def export_phieu(sess: Session, out_dir: str, bai: int, ngay: _dt.date,
                 rng: Optional[random.Random] = None) -> Tuple[str, int]:
    has = sum(1 for s in sess.students if core.is_num(s.dk[bai - 1]))
    if has == 0:
        raise AppError(f"Chưa có điểm kiểm tra định kỳ {bai} (cột {'EFGHI'[bai - 1]}) trong file nhập điểm.")
    path = target_path(out_dir, sess, "PDGTH", ngay, "xlsx", f"Bai{bai}")
    n = phieu_xlsx.xuat_phieu_xlsx(path, sess.meta, sess.students, bai, rng)
    return path, n


def _tpl(path: str, what: str) -> str:
    if not path or not os.path.exists(path):
        raise AppError(f"Chưa chọn file mẫu {what} (.xls hoặc .xlsx).")
    return path


def export_nhap_qua_trinh(sess: Session, tpl: str, out_dir: str, ngay: _dt.date, quy_che: str = "2022") -> Tuple[str, List[str]]:
    tpl = _tpl(tpl, "Form nhập quá trình")
    path = target_path(out_dir, sess, "NHAPQT", ngay, "xls" if tpl.lower().endswith(".xls") else "xlsx")
    try:
        w = forms_import.fill_qua_trinh(tpl, path, sess.meta, sess.students, quy_che)
    except forms_import.TemplateError as e:
        raise AppError(str(e))
    return path, w


def export_nhap_ket_thuc(sess: Session, tpl: str, out_dir: str, ngay: _dt.date, lan: int = 1) -> Tuple[str, List[str]]:
    tpl = _tpl(tpl, "Form nhập thi kết thúc mô đun")
    path = target_path(out_dir, sess, "NHAPKT", ngay, "xls" if tpl.lower().endswith(".xls") else "xlsx", "L2" if lan == 2 else "")
    try:
        w = forms_import.fill_ket_thuc(tpl, path, sess.meta, sess.students, lan)
    except forms_import.TemplateError as e:
        raise AppError(str(e))
    return path, w
