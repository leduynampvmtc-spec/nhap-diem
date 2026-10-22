"""
Bộ vẽ PDF dạng lưới (reportlab) — thay cho ExportAsFixedFormat / AppleScriptTask của Excel.
Ghi file PDF trực tiếp bằng Python nên chạy y hệt trên Windows và macOS.

Font: ưu tiên Times New Roman của hệ thống (giống mẫu Word/Excel), nếu không có thì dùng
Liberation Serif đi kèm (cùng kích thước ký tự với Times New Roman, có đủ tiếng Việt).
"""
from __future__ import annotations

import os
import sys
from typing import Callable, List, Sequence

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

REG, BOLD, ITAL, BI = "NDR", "NDB", "NDI", "NDBI"
_fonts_ready = False
FONT_SOURCE = ""


def resource_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    if hasattr(sys, "_MEIPASS"):
        base = os.path.join(base, "nhapdiem")
    return os.path.join(base, *parts)


def _system_tnr():
    cands = [
        # Windows
        [r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\timesbd.ttf",
         r"C:\Windows\Fonts\timesi.ttf", r"C:\Windows\Fonts\timesbi.ttf"],
        # macOS
        ["/System/Library/Fonts/Supplemental/Times New Roman.ttf",
         "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
         "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
         "/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf"],
        ["/Library/Fonts/Times New Roman.ttf", "/Library/Fonts/Times New Roman Bold.ttf",
         "/Library/Fonts/Times New Roman Italic.ttf", "/Library/Fonts/Times New Roman Bold Italic.ttf"],
        # Microsoft Office cho Mac (nếu font hệ thống bị thiếu)
        *[[f"/Applications/Microsoft {app}.app/Contents/Resources/DFonts/{fn}" for fn in
           ("Times New Roman.ttf", "Times New Roman Bold.ttf", "Times New Roman Italic.ttf", "Times New Roman Bold Italic.ttf")]
          for app in ("Word", "Excel", "PowerPoint")],
        # Linux (msttcorefonts)
        ["/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
         "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
         "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Italic.ttf",
         "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold_Italic.ttf"],
    ]
    for c in cands:
        if all(os.path.exists(p) for p in c):
            return c
    return None


def ensure_fonts() -> str:
    global _fonts_ready, FONT_SOURCE
    if _fonts_ready:
        return FONT_SOURCE
    chosen, label = None, ""
    sysf = _system_tnr()
    if sysf:
        try:
            for name, path in zip((REG, BOLD, ITAL, BI), sysf):
                pdfmetrics.registerFont(TTFont(name, path))
            # kiểm tra có đủ ký tự tiếng Việt
            f = pdfmetrics.getFont(REG)
            if all(ord(ch) in f.face.charToGlyph for ch in "ăâêôơưđẠặỗữỳ"):
                chosen, label = sysf, "Times New Roman (hệ thống)"
            else:
                chosen = None
        except Exception:
            chosen = None
    if not chosen:
        d = resource_path("fonts")
        files = ["LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf",
                 "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"]
        for name, fn in zip((REG, BOLD, ITAL, BI), files):
            pdfmetrics.registerFont(TTFont(name, os.path.join(d, fn)))
        label = "Liberation Serif (đi kèm)"
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    registerFontFamily(REG, normal=REG, bold=BOLD, italic=ITAL, boldItalic=BI)
    _fonts_ready = True
    FONT_SOURCE = label
    return label


# ---------------------------------------------------------------------------
# Vẽ chữ trong ô
# ---------------------------------------------------------------------------
def _wrap(text: str, font: str, size: float, maxw: float) -> List[str]:
    lines: List[str] = []
    for para in str(text).split("\n"):
        words = para.split(" ")
        cur = ""
        for w in words:
            trial = w if not cur else cur + " " + w
            if pdfmetrics.stringWidth(trial, font, size) <= maxw or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w
        # từ đơn quá dài -> cắt theo ký tự
        while pdfmetrics.stringWidth(cur, font, size) > maxw and len(cur) > 1:
            k = len(cur)
            while k > 1 and pdfmetrics.stringWidth(cur[:k], font, size) > maxw:
                k -= 1
            lines.append(cur[:k])
            cur = cur[k:]
        lines.append(cur)
    return lines


def draw_text(c: canvas.Canvas, text: str, x: float, y_top: float, w: float, h: float,
              font: str = REG, size: float = 10, align: str = "center", valign: str = "middle",
              pad: float = 2.0, wrap: bool = True, min_size: float = 5.0, lead: float = 1.15) -> None:
    """Vẽ text vào hộp (x, y_top → y_top-h). Tự thu nhỏ chữ nếu tràn."""
    if text is None or text == "":
        return
    text = str(text)
    inner_w = max(w - 2 * pad, 1)
    sz = size
    while True:
        lines = _wrap(text, font, sz, inner_w) if wrap else text.split("\n")
        tot_h = len(lines) * sz * lead
        widest = max(pdfmetrics.stringWidth(l, font, sz) for l in lines)
        if (tot_h <= h - 1 and widest <= inner_w + 0.5) or sz <= min_size:
            break
        sz -= 0.25
    c.setFont(font, sz)
    if valign == "top":
        y = y_top - pad - sz * 0.85
    elif valign == "bottom":
        y = y_top - h + pad + (len(lines) - 1) * sz * lead + sz * 0.2
    else:
        y = y_top - (h - tot_h) / 2 - sz * 0.85 + 0.0
    for ln in lines:
        lw = pdfmetrics.stringWidth(ln, font, sz)
        if align == "left":
            tx = x + pad
        elif align == "right":
            tx = x + w - pad - lw
        else:
            tx = x + (w - lw) / 2
        c.drawString(tx, y, ln)
        y -= sz * lead


def box(c: canvas.Canvas, x: float, y_top: float, w: float, h: float, lw: float = 0.5) -> None:
    c.setLineWidth(lw)
    c.rect(x, y_top - h, w, h, stroke=1, fill=0)


# ---------------------------------------------------------------------------
# Lưới cột
# ---------------------------------------------------------------------------
class Grid:
    """Cột theo đơn vị 'ký tự Excel'; đổi sang pt bằng hệ số u. Font/hàng nhân hệ số s."""

    def __init__(self, units: Sequence[float], x0: float, u: float, s: float):
        self.units, self.u, self.s = list(units), u, s
        self.xs = [x0]
        for w in self.units:
            self.xs.append(self.xs[-1] + w * u)

    def x(self, c0: int) -> float:
        return self.xs[c0]

    def w(self, c0: int, c1: int) -> float:
        return self.xs[c1 + 1] - self.xs[c0]

    @property
    def total(self) -> float:
        return self.xs[-1] - self.xs[0]


CHAR_PT = 5.5  # 1 ký tự Excel ~ 5.5pt ở 100%


def make_grid(units: Sequence[float], page_w: float, left: float, right: float,
              max_scale: float | None = None, hcenter: bool = True) -> Grid:
    usable = page_w - left - right
    s_fit = usable / (sum(units) * CHAR_PT)
    s = s_fit if max_scale is None else min(s_fit, max_scale)
    u = s * CHAR_PT
    total = sum(units) * u
    x0 = left + ((usable - total) / 2 if hcenter else 0)
    return Grid(units, x0, u, s)


# ---------------------------------------------------------------------------
# Tài liệu dạng bảng nhiều trang
# ---------------------------------------------------------------------------
class TableDoc:
    """
    intro  : vẽ ở trang đầu (trên cùng), trả về chiều cao đã dùng
    header : vẽ ở đầu MỖI trang (trước các hàng dữ liệu), trả về chiều cao
    row    : vẽ 1 hàng dữ liệu
    footer : vẽ sau hàng cuối (chữ ký…)
    """

    def __init__(self, path: str, pagesize, margins=(28, 28, 28, 28)):
        ensure_fonts()
        self.path = path
        self.page_w, self.page_h = pagesize
        self.ml, self.mr, self.mt, self.mb = margins

    def render(self, rows: Sequence, row_h: float, intro_h: float, header_h: float, footer_h: float,
               draw_intro: Callable, draw_header: Callable, draw_row: Callable, draw_footer: Callable,
               title: str = "", page_footer: Callable | None = None) -> int:
        c = canvas.Canvas(self.path, pagesize=(self.page_w, self.page_h))
        c.setTitle(title or os.path.basename(self.path))
        c.setAuthor("Nhập điểm – PV College")
        avail = self.page_h - self.mt - self.mb

        # chia trang
        pages: List[List] = []
        cur: List = []
        used = intro_h + header_h
        for r in rows:
            if used + row_h > avail and cur:
                pages.append(cur)
                cur, used = [], header_h
            cur.append(r)
            used += row_h
        pages.append(cur)
        # chữ ký không vừa trang cuối -> kéo vài hàng cuối sang trang mới (tránh trang chỉ có chữ ký)
        last = pages[-1]
        used_last = (intro_h if len(pages) == 1 else 0) + header_h + len(last) * row_h
        if used_last + footer_h > avail:
            if len(last) >= 3:
                pages[-1], moved = last[:-2], last[-2:]
                pages.append(moved)
            elif len(pages) > 1 and len(pages[-2]) >= 3:
                moved = pages[-2][-2:]
                pages[-2] = pages[-2][:-2]
                pages[-1] = moved + last
            else:
                pages.append([])

        for pi, chunk in enumerate(pages):
            y = self.page_h - self.mt
            if pi == 0:
                y -= draw_intro(c, y) or 0
            y -= draw_header(c, y, pi) or 0
            for r in chunk:
                draw_row(c, y, r)
                y -= row_h
            if pi == len(pages) - 1:
                draw_footer(c, y)
            if page_footer:
                page_footer(c, pi + 1, len(pages), self.page_w, self.page_h, self.ml, self.mr)
            c.showPage()
        c.save()
        return len(pages)


def draw_page_footer(c, i: int, n: int, page_w: float, ml: float, mr: float, left: str, center: str,
                     size: float = 12, y: float = 18) -> None:
    """Footer mỗi trang: trái = mã biểu mẫu, giữa = LBH/NHL, phải = Trang i/N (Times New Roman 12)."""
    c.setFont(REG, size)
    if left:
        c.drawString(ml, y, left)
    if center:
        c.drawCentredString(page_w / 2, y, center)
    c.drawRightString(page_w - mr, y, f"Trang {i}/{n}")


def a4(landscape_mode: bool):
    return landscape(A4) if landscape_mode else A4
