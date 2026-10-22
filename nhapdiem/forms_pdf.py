"""
Bốn biểu mẫu xuất PDF — dựng lại đúng bố cục các sheet "DK DT", "TL", "LT", "TK"
trong Form Code VBA.xlsm (cùng chữ, cột, hàng, khung kẻ; không tô màu — vì macro VBA
xóa màu nền trước khi xuất PDF).
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Dict, List, Sequence

from reportlab.pdfbase import pdfmetrics

from .core import Meta, Student, DASH, fmt_pct, fmt_score, ngay_thang_nam, xet_dieu_kien
from .pdfgrid import BI, BOLD, ITAL, REG, Grid, TableDoc, a4, box, draw_page_footer, draw_text, make_grid


def _cell(c, g: Grid, c0: int, c1: int, y_top: float, h_excel: float, text: str = "",
          font: str = REG, size: float = 13, align: str = "center", valign: str = "middle",
          border: bool = True, wrap: bool = True, h_pt: float | None = None, underline: bool = False):
    x, w = g.x(c0), g.w(c0, c1)
    h = h_pt if h_pt is not None else h_excel * g.s
    if border:
        box(c, x, y_top, w, h)
    if text:
        draw_text(c, text, x, y_top, w, h, font=font, size=size * g.s, align=align, valign=valign,
                  pad=2.0 * g.s, wrap=wrap)
        if underline:
            from reportlab.pdfbase import pdfmetrics as pm
            sz = size * g.s
            tw = pm.stringWidth(text.split("\n")[-1], font, sz)
            ty = y_top - (h - sz * 1.15) / 2 - sz * 0.85 - 1.5
            c.setLineWidth(0.5)
            c.line(x + (w - tw) / 2, ty, x + (w + tw) / 2, ty)


def _pf(ft, key):
    """Hàm vẽ footer mỗi trang (Mã biểu mẫu | LBH/NHL | Trang i/N) hoặc None."""
    if ft is None:
        return None
    return lambda c, i, n, w, h, ml, mr: draw_page_footer(c, i, n, w, ml, mr, ft.left(key), ft.center(key))


def dk_rows(students: List[Student], opt: "DKOptions") -> List[list]:
    rows = []
    for st in students:
        du, note = xet_dieu_kien(st, opt.quy_che)
        ex = [st.tx1, st.tx2] + st.dk[:6]
        rows.append([str(st.stt), st.mssv, st.name] + [fmt_score(v) for v in ex] + [
            fmt_score(st.o), fmt_pct(st.vang_tong), "X" if du else "", "" if du else "X", note])
    return rows


def tl_rows(ds: List[dict]) -> List[list]:
    return [[str(d["stt"]), d["mssv"], d["name"], "X" if d["lt"] else "", "X" if d["th"] else "", d["ghichu"]]
            for d in ds]


def lt_rows(ds: List[dict]) -> List[list]:
    return [[str(d["stt"]), d["mssv"], d["name"], d["dob"], d["diem"], d["ghichu"]] for d in ds]


# ===========================================================================
# 1) DANH SÁCH XÉT ĐIỀU KIỆN DỰ THI  (sheet DK DT)
# ===========================================================================
@dataclass
class DKOptions:
    ma_lhp: str = ""
    # hàng 12 của bảng "Hình thức thi"
    tn_may: bool = True          # B12  Trắc nghiệm - Trên máy
    tn_giay: bool = False        # C12  Trắc nghiệm - Trên giấy
    van_dap: bool = False        # D12
    tu_luan: bool = False        # E12
    ht_khac: bool = False        # G12
    dia_diem: str = "Phòng máy tính"   # I12  Thi thực hành - địa điểm
    so_hssv_ca: str = "-"        # L12  Số HSSV / ca thi
    tg_ca: str = "-"             # N12  Thời gian (theo ca)
    tg_tap_trung: str = "-"      # O12  Thời gian thi tập trung cả lớp (phút)
    quy_che: str = "2022"        # "2021" | "2022"
    to_bo_mon: str = "CƠ KHÍ"
    ngay_lap: _dt.date = field(default_factory=_dt.date.today)


DK_UNITS = [5.66, 18, 47.33, 8, 8, 7.33, 7.33, 7.33, 7.33, 7.33, 6, 7.66, 11.66, 10, 10, 22.33]


def pdf_dieu_kien_du_thi(path: str, meta: Meta, students: List[Student], opt: DKOptions, ft=None) -> int:
    doc = TableDoc(path, a4(True), margins=(28, 28, 28, 28))
    g = make_grid(DK_UNITS, doc.page_w, 28, 28, hcenter=True)
    s = g.s
    last = len(DK_UNITS) - 1
    intro_rows = [33, 17, 20, 17, 25.5, 20.25, 20.25, 20.25, 20.25, 20.25, 45, 20.25, 4.5]
    head_rows = [57, 43.5, 43.5]
    foot_rows = [6, 20.25, 32.25, 71.25]

    X = lambda b: "X" if b else "-"

    def intro(c, y):
        def top(i):
            return y - sum(intro_rows[:i]) * s

        _cell(c, g, 0, 6, top(0), intro_rows[0], f"PHÒNG ĐÀO TẠO\nTỔ BỘ MÔN {opt.to_bo_mon}", size=13, border=False)
        _cell(c, g, 7, last, top(0), intro_rows[0],
              "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc", size=13, border=False)
        _cell(c, g, 0, last, top(2), intro_rows[2], "DANH SÁCH XÉT ĐIỀU KIỆN HSSV DỰ THI MH/MĐ", font=BOLD, size=16,
              border=False)
        _cell(c, g, 0, 2, top(4), intro_rows[4], meta.a5_raw, size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 3, 4, top(4), intro_rows[4], f"SỐ TC: {meta.tin_chi}", size=13, align="left", valign="bottom",
              border=False, wrap=False)
        _cell(c, g, 5, 6, top(4), intro_rows[4], "Tổng số giờ:", size=13, align="right", valign="bottom", border=False, wrap=False)
        _cell(c, g, 7, 7, top(4), intro_rows[4], f"{meta.tong_gio}", size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 0, 2, top(5), intro_rows[5], f"LỚP: {meta.lop_khoa}", size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 6, 6, top(5), intro_rows[5], "LT:", size=13, align="right", valign="bottom", border=False, wrap=False)
        _cell(c, g, 7, 7, top(5), intro_rows[5], f"{meta.gio_lt:g}", size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 0, 2, top(6), intro_rows[6], f"MÃ LHP: {opt.ma_lhp}", size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 6, 6, top(6), intro_rows[6], "TH:", size=13, align="right", valign="bottom", border=False, wrap=False)
        _cell(c, g, 7, 7, top(6), intro_rows[6], f"{meta.gio_th:g}", size=13, align="left", valign="bottom", border=False, wrap=False)
        _cell(c, g, 0, 2, top(7), intro_rows[7], "HÌNH THỨC THI (*):", size=13, align="left", valign="bottom", border=False, wrap=False)
        t9, t10, t11, t12 = top(8), top(9), top(10), top(11)
        h9, h10, h11, h12 = intro_rows[8], intro_rows[9], intro_rows[10], intro_rows[11]
        _cell(c, g, 1, 7, t9, h9, "Lý Thuyết", font=BOLD, size=11)
        _cell(c, g, 8, last, t9, h9, "Thi Thực hành", font=BOLD, size=11)
        _cell(c, g, 1, 2, t10, h10, "Trắc nghiệm", font=ITAL, size=11)
        _cell(c, g, 3, 3, t10, h10 + h11, "Vấn đáp", font=ITAL, size=11)
        _cell(c, g, 4, 5, t10, h10 + h11, "Tự luận", font=ITAL, size=11)
        _cell(c, g, 6, 7, t10, h10 + h11, "HT khác", font=ITAL, size=11)
        _cell(c, g, 8, 10, t10, h10 + h11, "Địa điểm thi", font=ITAL, size=11)
        _cell(c, g, 11, 13, t10, h10, "Thi theo ca", font=ITAL, size=11)
        _cell(c, g, 14, last, t10, h10, "Thi tập trung cả lớp", font=ITAL, size=11)
        _cell(c, g, 1, 1, t11, h11, "Trên máy", font=ITAL, size=11)
        _cell(c, g, 2, 2, t11, h11, "Trên giấy", font=ITAL, size=11)
        _cell(c, g, 11, 12, t11, h11, "Số HSSV/ ca thi", font=ITAL, size=11)
        _cell(c, g, 13, 13, t11, h11, "Thời gian", font=ITAL, size=11)
        _cell(c, g, 14, last, t11, h11, "Thời gian thi (Phút)", font=ITAL, size=11)
        _cell(c, g, 1, 1, t12, h12, X(opt.tn_may), size=11)
        _cell(c, g, 2, 2, t12, h12, X(opt.tn_giay), size=11)
        _cell(c, g, 3, 3, t12, h12, X(opt.van_dap), size=11)
        _cell(c, g, 4, 5, t12, h12, X(opt.tu_luan), size=11)
        _cell(c, g, 6, 7, t12, h12, X(opt.ht_khac), size=11)
        _cell(c, g, 8, 10, t12, h12, opt.dia_diem or "-", size=11)
        _cell(c, g, 11, 12, t12, h12, opt.so_hssv_ca or "-", size=11)
        _cell(c, g, 13, 13, t12, h12, opt.tg_ca or "-", size=11)
        _cell(c, g, 14, last, t12, h12, opt.tg_tap_trung or "-", size=11)
        return sum(intro_rows) * s

    def header(c, y, pi):
        t = [y, y - head_rows[0] * s, y - (head_rows[0] + head_rows[1]) * s]
        h = head_rows
        _cell(c, g, 0, 0, t[0], sum(h), "STT", font=BOLD)
        _cell(c, g, 1, 1, t[0], sum(h), "Mã số HSSV", font=BOLD)
        _cell(c, g, 2, 2, t[0], sum(h), "HỌ VÀ TÊN HSSV", font=BOLD)
        _cell(c, g, 3, 4, t[0], h[0], "ĐIỂM KT THƯỜNG XUYÊN", font=BOLD)
        _cell(c, g, 3, 3, t[1], h[1] + h[2], "1", font=BOLD)
        _cell(c, g, 4, 4, t[1], h[1] + h[2], "2", font=BOLD)
        _cell(c, g, 5, 10, t[0], h[0], "ĐIỂM KT ĐỊNH KỲ", font=BOLD)
        for i in range(6):
            _cell(c, g, 5 + i, 5 + i, t[1], h[1] + h[2], str(i + 1), font=BOLD)
        _cell(c, g, 11, 11, t[0], sum(h), "TBC KT", font=BOLD)
        _cell(c, g, 12, 12, t[0], h[0] + h[1], "% Tiết vắng học (có phép và không phép)", font=BOLD)
        _cell(c, g, 12, 12, t[2], h[2], "LT + TH", font=BOLD)
        _cell(c, g, 13, 14, t[0], h[0], "XÉT ĐK DỰ THI\n(đánh dấu x)", font=BOLD)
        _cell(c, g, 13, 13, t[1], h[1] + h[2], "Đủ ĐK thi", font=BOLD)
        _cell(c, g, 14, 14, t[1], h[1] + h[2], "Không đủ ĐK thi", font=BOLD)
        _cell(c, g, 15, 15, t[0], sum(h), "GHI CHÚ", font=BOLD)
        return sum(h) * s

    rows = dk_rows(students, opt)

    aligns = ["center", "center", "left"] + ["center"] * 8 + ["center"] * 5 + ["left"]
    row_h = 20 * s
    spans = [(0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 9), (10, 10),
             (11, 11), (12, 12), (13, 13), (14, 14), (15, 15)]

    def row(c, y, r):
        for (c0, c1), txt, al in zip(spans, r, aligns):
            _cell(c, g, c0, c1, y, 20, txt, size=13, align=al, h_pt=row_h)

    def footer(c, y):
        y -= foot_rows[0] * s
        _cell(c, g, 12, last, y, foot_rows[1], ngay_thang_nam(opt.ngay_lap), font=ITAL, size=13, border=False)
        y -= foot_rows[1] * s
        _cell(c, g, 0, 4, y, foot_rows[2], f"TỔ BỘ MÔN {opt.to_bo_mon}\nTỔ PHÓ", font=BOLD, size=13,
              valign="top", border=False)
        _cell(c, g, 12, last, y, foot_rows[2], "GIÁO VIÊN GIẢNG DẠY", font=BOLD, size=13, valign="top", border=False)

    return doc.render(rows, row_h, sum(intro_rows) * s, sum(head_rows) * s, sum(foot_rows) * s,
                      intro, header, row, footer, title="Danh sách xét điều kiện dự thi", page_footer=_pf(ft, "dk"))


# ===========================================================================
# 2) DANH SÁCH THI LẠI  (sheet TL)
# ===========================================================================
@dataclass
class TLOptions:
    ma_lhp: str = ""
    hinh_thuc_lt: str = "Thi trắc nghiệm trên máy tính"     # TL!H1
    dia_diem_th: str = "xưởng TBĐ"                            # TL!H2 = "Thi thực hành tại " + địa điểm
    to_bo_mon: str = "CƠ KHÍ"
    ngay_lap: _dt.date = field(default_factory=_dt.date.today)

    @property
    def text_th(self) -> str:
        return f"Thi thực hành tại {self.dia_diem_th}".strip()


TL_UNITS = [5.3, 15.7, 32, 14.3, 14.3, 19.7]


def pdf_thi_lai(path: str, meta: Meta, ds: List[dict], opt: TLOptions, ft=None) -> int:
    doc = TableDoc(path, a4(True), margins=(57, 28, 42, 42))
    g = make_grid(TL_UNITS, doc.page_w, 57, 28, max_scale=1.1, hcenter=True)
    s = g.s
    intro_rows = [15.75, 15, 26.25, 18, 26.25, 17, 17, 17, 15]
    head_rows = [31.5, 30]
    foot_rows = [6, 20, 14, 60]

    def intro(c, y):
        def top(i):
            return y - sum(intro_rows[:i]) * s

        _cell(c, g, 0, 2, top(0), intro_rows[0], "PHÒNG ĐÀO TẠO", size=11, border=False)
        _cell(c, g, 3, 5, top(0), intro_rows[0], "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=10, border=False)
        _cell(c, g, 0, 2, top(1), intro_rows[1], f"TỔ BỘ MÔN {opt.to_bo_mon}", font=BOLD, size=11, border=False)
        _cell(c, g, 3, 5, top(1), intro_rows[1], "Độc lập - Tự do - Hạnh phúc", font=BOLD, size=11, border=False,
              underline=True)
        _cell(c, g, 0, 5, top(3), intro_rows[3], "DANH SÁCH HSSV THI LẠI MH/MĐ (LẦN 2)", font=BOLD, size=14,
              valign="bottom", border=False)
        _cell(c, g, 0, 2, top(5), intro_rows[5], meta.a5_raw, font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 3, 3, top(5), intro_rows[5], f"SỐ TC: {meta.tin_chi}", font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 4, 4, top(5), intro_rows[5], "Tổng số giờ:", font=BOLD, size=11, align="right", border=False, wrap=False)
        _cell(c, g, 5, 5, top(5), intro_rows[5], f"{meta.tong_gio}", font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 0, 2, top(6), intro_rows[6], f"LỚP: {meta.lop_khoa}", font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 4, 4, top(6), intro_rows[6], "LT:", font=BOLD, size=11, align="right", border=False, wrap=False)
        _cell(c, g, 5, 5, top(6), intro_rows[6], f"{meta.gio_lt:g}", font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 0, 2, top(7), intro_rows[7], f"MÃ LHP: {opt.ma_lhp}", font=BOLD, size=11, align="left", border=False, wrap=False)
        _cell(c, g, 4, 4, top(7), intro_rows[7], "TH:", font=BOLD, size=11, align="right", border=False, wrap=False)
        _cell(c, g, 5, 5, top(7), intro_rows[7], f"{meta.gio_th:g}", font=BOLD, size=11, align="left", border=False, wrap=False)
        return sum(intro_rows) * s

    def header(c, y, pi):
        h = head_rows
        _cell(c, g, 0, 0, y, sum(h), "STT", font=BOLD, size=10)
        _cell(c, g, 1, 1, y, sum(h), "Mã số HSSV", font=BOLD, size=10)
        _cell(c, g, 2, 2, y, sum(h), "HỌ VÀ TÊN HSSV", font=BOLD, size=10)
        _cell(c, g, 3, 4, y, h[0], "NỘI DUNG THI LẠI", font=BOLD, size=10)
        _cell(c, g, 3, 3, y - h[0] * s, h[1], "LT", font=BOLD, size=10)
        _cell(c, g, 4, 4, y - h[0] * s, h[1], "TH", font=BOLD, size=10)
        _cell(c, g, 5, 5, y, sum(h), "GHI CHÚ", font=BOLD, size=10)
        return sum(h) * s

    rows = tl_rows(ds)
    aligns = ["center", "left", "left", "center", "center", "left"]
    row_h = 33 * s

    def row(c, y, r):
        for i, (txt, al) in enumerate(zip(r, aligns)):
            _cell(c, g, i, i, y, 33, txt, size=13, align=al, h_pt=row_h)

    def footer(c, y):
        y -= foot_rows[0] * s
        _cell(c, g, 3, 5, y, foot_rows[1], ngay_thang_nam(opt.ngay_lap), font=ITAL, size=11, border=False)
        y -= foot_rows[1] * s
        _cell(c, g, 0, 2, y, foot_rows[2], "TRƯỞNG BỘ MÔN", font=BOLD, size=11, border=False)
        _cell(c, g, 3, 5, y, foot_rows[2], "GIÁO VIÊN BỘ MÔN", font=BOLD, size=11, border=False)

    return doc.render(rows, row_h, sum(intro_rows) * s, sum(head_rows) * s, sum(foot_rows) * s,
                      intro, header, row, footer, title="Danh sách HSSV thi lại", page_footer=_pf(ft, "tl"))


# ===========================================================================
# 3) BẢNG ĐIỂM TỔNG HỢP — ĐIỂM THI LT  (sheet LT)
# ===========================================================================
@dataclass
class LTOptions:
    lan_thi: int = 1
    ca_thi: str = ""
    to_bo_mon: str = "CƠ KHÍ"
    ngay_lap: _dt.date = field(default_factory=_dt.date.today)


LT_UNITS = [4.7, 19.2, 30, 16, 13, 31]


def pdf_diem_lt(path: str, meta: Meta, ds: List[dict], opt: LTOptions, ft=None) -> int:
    doc = TableDoc(path, a4(False), margins=(42, 28, 42, 42))
    g = make_grid(LT_UNITS, doc.page_w, 42, 28, hcenter=True)
    s = g.s
    intro_rows = [15.75, 15, 15, 20.25, 18, 18.75, 18, 18, 18, 18, 19.5]
    head_rows = [33]
    foot_rows = [6, 20.25, 18, 18, 50]

    def intro(c, y):
        def top(i):
            return y - sum(intro_rows[:i]) * s

        _cell(c, g, 0, 2, top(0), intro_rows[0], "TRƯỜNG CAO ĐẲNG DẦU KHÍ", font=BOLD, size=12, border=False)
        _cell(c, g, 4, 5, top(0), intro_rows[0], "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", size=12, border=False)
        _cell(c, g, 0, 2, top(1), intro_rows[1], "PHÒNG ĐÀO TẠO", font=BOLD, size=12, border=False)
        _cell(c, g, 4, 5, top(1), intro_rows[1], "Độc lập - Tự do - Hạnh phúc", font=BOLD, size=12, border=False,
              underline=True)
        _cell(c, g, 0, 2, top(2), intro_rows[2], f"TỔ BỘ MÔN {opt.to_bo_mon}", font=BOLD, size=12, border=False)
        _cell(c, g, 0, 5, top(4), intro_rows[4], "BẢNG ĐIỂM TỔNG HỢP THEO LỚP", font=BOLD, size=14, border=False)
        _cell(c, g, 0, 5, top(6), intro_rows[6], f"Tên MH/MĐ: {meta.ten}", size=13, align="left", border=False, wrap=False)
        _cell(c, g, 0, 5, top(7), intro_rows[7], f"Lần thi: {opt.lan_thi:02d}", size=13, align="left", border=False, wrap=False)
        _cell(c, g, 0, 5, top(8), intro_rows[8], f"Lớp: {meta.lop_khoa}", size=13, align="left", border=False, wrap=False)
        _cell(c, g, 0, 5, top(9), intro_rows[9], f"Ca thi: {opt.ca_thi}", size=13, align="left", border=False, wrap=False)
        return sum(intro_rows) * s

    def header(c, y, pi):
        for i, t in enumerate(["Stt", "Mã số HSSV", "Họ và tên", "Ngày sinh", "Điểm", "Ghi chú"]):
            _cell(c, g, i, i, y, head_rows[0], t, font=BOLD, size=13)
        return head_rows[0] * s

    rows = lt_rows(ds)
    aligns = ["center", "center", "left", "center", "center", "center"]
    row_h = 20 * s

    def row(c, y, r):
        for i, (txt, al) in enumerate(zip(r, aligns)):
            _cell(c, g, i, i, y, 20, txt, size=13, align=al, h_pt=row_h)

    def footer(c, y):
        y -= foot_rows[0] * s
        _cell(c, g, 3, 5, y, foot_rows[1], ngay_thang_nam(opt.ngay_lap), font=ITAL, size=13, border=False)
        y -= foot_rows[1] * s
        _cell(c, g, 3, 5, y, foot_rows[2], "CÁN BỘ GIÁM SÁT / CÁN BỘ COI THI", font=BOLD, size=13, border=False)
        y -= foot_rows[2] * s
        _cell(c, g, 3, 5, y, foot_rows[3], "GIÁM KHẢO 1 / GIÁM KHẢO 2", font=BOLD, size=13, border=False)

    return doc.render(rows, row_h, sum(intro_rows) * s, sum(head_rows) * s, sum(foot_rows) * s,
                      intro, header, row, footer, title="Bảng điểm tổng hợp theo lớp", page_footer=_pf(ft, "lt"))


# ===========================================================================
# 4) XẾP THỨ TỰ  (sheet TK)
# ===========================================================================
TK_UNITS = [5.7, 24, 10.7, 1.7, 5, 20, 10.7, 2, 5.7, 23, 10.7, 2, 6.7, 25, 11]
TK_GROUPS = [(0, "SINH VIÊN CÓ TỔNG KẾT >=7", "lon7"),
             (4, "SINH VIÊN CÓ TỔNG KẾT < 5", "nho5"),
             (8, "SINH VIÊN TOP 3", "top3"),
             (12, "SINH VIÊN BOTTOM 3", "bottom3")]


def pdf_xep_thu_tu(path: str, meta: Meta, data: Dict[str, list], ngay_lap: _dt.date | None = None, ft=None) -> int:
    doc = TableDoc(path, a4(True), margins=(28, 28, 32, 32))
    g = make_grid(TK_UNITS, doc.page_w, 28, 28, hcenter=True)
    s = g.s
    cap_h = 34 * s
    head_rows = [18, 18]
    n = max([len(data[k]) for _, _, k in TK_GROUPS] + [0])

    def intro(c, y):
        sz = 13 * s
        c.setFont(BOLD, sz)
        line1 = f"XẾP THỨ TỰ ĐIỂM TỔNG KẾT – {meta.ten} – LỚP {meta.lop_khoa}"
        draw_text(c, line1, g.x(0), y, g.total, cap_h, font=BOLD, size=14 * s, align="center", wrap=False)
        return cap_h

    def header(c, y, pi):
        for c0, title, _ in TK_GROUPS:
            _cell(c, g, c0, c0 + 2, y, head_rows[0], title, font=BOLD, size=12, border=False)
            yy = y - head_rows[0] * s
            _cell(c, g, c0, c0, yy, head_rows[1], "Stt", font=BOLD, size=13)
            _cell(c, g, c0 + 1, c0 + 1, yy, head_rows[1], "Họ và tên", font=BOLD, size=13)
            _cell(c, g, c0 + 2, c0 + 2, yy, head_rows[1], "Điểm tổng", font=BOLD, size=13)
        return sum(head_rows) * s

    rows = list(range(n))
    row_h = 20.25 * s

    def row(c, y, i):
        for c0, _, key in TK_GROUPS:
            lst = data[key]
            if i < len(lst):
                nm, sc = lst[i]
                _cell(c, g, c0, c0, y, 20.25, str(i + 1), size=13, h_pt=row_h)
                _cell(c, g, c0 + 1, c0 + 1, y, 20.25, nm, size=13, align="left", h_pt=row_h)
                _cell(c, g, c0 + 2, c0 + 2, y, 20.25, sc, size=13, h_pt=row_h)

    def footer(c, y):
        return None

    return doc.render(rows, row_h, cap_h, sum(head_rows) * s, 0, intro, header, row, footer,
                      title="Xếp thứ tự điểm tổng kết", page_footer=_pf(ft, "tk"))
