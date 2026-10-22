"""
Giao diện tkinter (có sẵn trong Python trên Windows và macOS — không cần cài thêm gì).
5 tab: Điều kiện dự thi · Thi lại · Điểm thi LT · Xếp thứ tự · Phiếu đánh giá TH.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from . import auth
from . import controller as ctl
from . import forms_pdf as F
from .osutil import open_file, open_in_file_manager
from .pdfgrid import ensure_fonts, resource_path

APP_TITLE = "Xuất biểu mẫu từ Form nhập điểm — PV College"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".nhapdiem_config.json")

GREEN = "#1B8A47"      # xanh lá của logo PV College
BLUE = "#23327F"       # xanh dương của logo
BG = "#F3F6F4"
AUTHOR = "Người viết ứng dụng: Giảng viên Lê Duy Nam"
ORG = "TRƯỜNG CAO ĐẲNG DẦU KHÍ  ·  TỔ BỘ MÔN CƠ KHÍ"


def load_logo(size: int):
    """Nạp logo trường (PNG) — trả None nếu thiếu file để app vẫn chạy."""
    try:
        return tk.PhotoImage(file=resource_path("assets", f"logo_{size}.png"))
    except Exception:
        return None


def apply_theme(root) -> None:
    try:
        st = ttk.Style(root)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        root.configure(bg=BG)
        st.configure(".", background=BG)
        st.configure("TFrame", background=BG)
        st.configure("TLabel", background=BG)
        st.configure("TLabelframe", background=BG)
        st.configure("TLabelframe.Label", background=BG, foreground=BLUE, font=("", 10, "bold"))
        st.configure("TCheckbutton", background=BG)
        st.configure("TRadiobutton", background=BG)
        st.configure("TNotebook", background=BG)
        st.configure("TNotebook.Tab", padding=(12, 6), font=("", 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", GREEN)], foreground=[("selected", "white")])
        st.configure("TButton", padding=(10, 5))
        st.configure("Accent.TButton", background=GREEN, foreground="white", font=("", 10, "bold"), padding=(14, 7))
        st.map("Accent.TButton", background=[("active", BLUE)])
        st.configure("Header.TFrame", background="white")
        st.configure("Bar.TFrame", background=GREEN)
        st.configure("Title.TLabel", background="white", foreground=BLUE, font=("", 15, "bold"))
        st.configure("Org.TLabel", background="white", foreground=GREEN, font=("", 10, "bold"))
        st.configure("Sub.TLabel", background="white", foreground="#555555")
        st.configure("Foot.TLabel", background=BG, foreground="#666666")
    except Exception:
        pass


class LoginWindow(tk.Tk):
    """Màn hình đăng nhập có logo Trường. `self.user` được đặt khi đăng nhập thành công."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Đăng nhập — " + APP_TITLE)
        self.resizable(False, False)
        apply_theme(self)
        self.user = ""
        self._tries = 0
        self.v_user = tk.StringVar()
        self.v_pw = tk.StringVar()
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                self.v_user.set(json.load(f).get("last_user", ""))
        except Exception:
            pass

        card = ttk.Frame(self, padding=28)
        card.pack()
        self._logo = load_logo(240)
        if self._logo:
            ttk.Label(card, image=self._logo).pack(pady=(0, 6))
        ttk.Label(card, text="XUẤT BIỂU MẪU TỪ FORM NHẬP ĐIỂM", foreground=BLUE,
                  font=("", 14, "bold")).pack()
        ttk.Label(card, text=ORG, foreground=GREEN, font=("", 9, "bold")).pack(pady=(2, 14))

        form = ttk.Frame(card)
        form.pack()
        ttk.Label(form, text="Tài khoản:").grid(row=0, column=0, sticky="e", padx=6, pady=5)
        self.e_user = ttk.Entry(form, textvariable=self.v_user, width=26)
        self.e_user.grid(row=0, column=1, pady=5)
        ttk.Label(form, text="Mật khẩu:").grid(row=1, column=0, sticky="e", padx=6, pady=5)
        self.e_pw = ttk.Entry(form, textvariable=self.v_pw, width=26, show="•")
        self.e_pw.grid(row=1, column=1, pady=5)
        self.msg = ttk.Label(card, text="", foreground="#B00020")
        self.msg.pack(pady=(8, 0))
        ttk.Button(card, text="Đăng nhập", style="Accent.TButton", command=self._submit).pack(pady=(8, 10))
        ttk.Label(card, text=AUTHOR, style="Foot.TLabel").pack()

        self.bind("<Return>", lambda _e: self._submit())
        self.e_pw.focus_set() if self.v_user.get() else self.e_user.focus_set()
        self.update_idletasks()
        try:  # căn giữa màn hình
            w, h = self.winfo_reqwidth(), self.winfo_reqheight()
            self.geometry(f"+{(self.winfo_screenwidth() - w) // 2}+{(self.winfo_screenheight() - h) // 3}")
        except Exception:
            pass

    def _submit(self) -> None:
        if auth.check_login(self.v_user.get(), self.v_pw.get()):
            self.user = self.v_user.get().strip()
            try:  # chỉ nhớ tên tài khoản, không lưu mật khẩu
                try:
                    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = {}
                data["last_user"] = self.user
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=1)
            except Exception:
                pass
            self.destroy()
            return
        self._tries += 1
        self.v_pw.set("")
        self.msg.configure(text="Sai tài khoản hoặc mật khẩu." + (
            f" (lần {self._tries})" if self._tries > 1 else ""))
        self.e_pw.focus_set()


class App(tk.Tk):
    def __init__(self, user: str = "") -> None:
        super().__init__()
        self.user = user
        apply_theme(self)
        self.title(APP_TITLE)
        self.geometry("920x860")
        self.minsize(840, 720)
        try:
            self.font_label = ensure_fonts()
        except Exception as e:  # thiếu font -> báo sớm
            messagebox.showerror("Lỗi font", f"Không nạp được font PDF:\n{e}")
            raise

        import datetime as dt
        # ---- biến dùng chung ----
        self.v_in = tk.StringVar()
        self.v_out = tk.StringVar()
        self.v_ngay = tk.StringVar(value=dt.date.today().strftime("%d/%m/%Y"))
        self.v_tbm = tk.StringVar(value="CƠ KHÍ")
        self.v_malhp = tk.StringVar()
        # ---- tab 1 ----
        self.v_tn_may = tk.BooleanVar(value=True)
        self.v_tn_giay = tk.BooleanVar(value=False)
        self.v_van_dap = tk.BooleanVar(value=False)
        self.v_tu_luan = tk.BooleanVar(value=False)
        self.v_ht_khac = tk.BooleanVar(value=False)
        self.v_dia_diem = tk.StringVar(value="Phòng máy tính")
        self.v_so_ca = tk.StringVar(value="-")
        self.v_tg_ca = tk.StringVar(value="-")
        self.v_tg_tt = tk.StringVar(value="-")
        self.v_quyche = tk.StringVar(value="2022")
        self.v_att = tk.StringVar()
        self.v_drop = tk.BooleanVar(value=True)
        # ---- tab 2 ----
        self.v_ht_lt = tk.StringVar(value="Thi trắc nghiệm trên máy tính")
        self.v_dd_th = tk.StringVar(value="xưởng TBĐ")
        # ---- tab 3 ----
        self.v_lan = tk.IntVar(value=1)
        self.v_ca = tk.StringVar()
        # ---- tab 5 ----
        self.v_bai = tk.IntVar(value=1)

        self._load_config()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log(f"Font file xuất: {self.font_label}")
        if "Times New Roman (hệ thống)" not in self.font_label:
            self.log("⚠ Máy chưa có Times New Roman nên PDF dùng Liberation Serif (cùng kích thước chữ, hình dáng gần như giống hệt). "
                     "Cài Times New Roman rồi mở lại app để PDF dùng đúng font này.")
        self.log("Chọn file Form nhập điểm và thư mục lưu, rồi bấm Xuất ở từng tab.")

    # ------------------------------------------------------------------
    # cấu hình nhớ giữa các lần chạy
    # ------------------------------------------------------------------
    _SAVED = ("v_in", "v_out", "v_tbm", "v_malhp", "v_dia_diem", "v_so_ca", "v_tg_ca", "v_tg_tt",
              "v_ht_lt", "v_dd_th", "v_ca", "v_quyche")

    def _load_config(self) -> None:
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k in self._SAVED:
                if k in data:
                    getattr(self, k).set(data[k])
        except Exception:
            pass

    def _save_config(self) -> None:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump({k: getattr(self, k).get() for k in self._SAVED}, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _on_close(self) -> None:
        self._save_config()
        self.destroy()

    # ------------------------------------------------------------------
    # dựng giao diện
    # ------------------------------------------------------------------
    def _row_entry(self, parent, r: int, label: str, var: tk.StringVar, browse: Callable | None = None,
                   width: int = 60, clear: Callable | None = None) -> None:
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", padx=(8, 6), pady=3)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=r, column=1, sticky="we", pady=3)
        if browse:
            ttk.Button(parent, text="Chọn…", command=browse).grid(row=r, column=2, padx=6, pady=3)
        if clear:
            ttk.Button(parent, text="Xóa", command=clear, width=4).grid(row=r, column=3, padx=(0, 8), pady=3)

    def _build_header(self) -> None:
        hd = ttk.Frame(self, style="Header.TFrame", padding=(14, 8))
        hd.pack(fill="x")
        self._logo72 = load_logo(72)
        if self._logo72:
            ttk.Label(hd, image=self._logo72, background="white").pack(side="left", padx=(0, 12))
        mid = ttk.Frame(hd, style="Header.TFrame")
        mid.pack(side="left", fill="x", expand=True)
        ttk.Label(mid, text="XUẤT BIỂU MẪU TỪ FORM NHẬP ĐIỂM", style="Title.TLabel").pack(anchor="w")
        ttk.Label(mid, text=ORG, style="Org.TLabel").pack(anchor="w")
        ttk.Label(mid, text=AUTHOR, style="Sub.TLabel").pack(anchor="w")
        right = ttk.Frame(hd, style="Header.TFrame")
        right.pack(side="right")
        ttk.Label(right, text=f"Tài khoản: {self.user}" if self.user else "", style="Sub.TLabel").pack(anchor="e")
        ttk.Button(right, text="Đăng xuất", command=self._logout).pack(anchor="e", pady=(4, 0))
        ttk.Frame(self, style="Bar.TFrame", height=4).pack(fill="x")

    def _logout(self) -> None:
        self._save_config()
        self.relogin = True
        self.destroy()

    def _build(self) -> None:
        self._build_header()
        top = ttk.LabelFrame(self, text="Dùng chung")
        top.pack(fill="x", padx=10, pady=(10, 4))
        top.columnconfigure(1, weight=1)
        self._row_entry(top, 0, "File Form nhập điểm (.xlsx):", self.v_in, self._pick_in)
        self._row_entry(top, 1, "Thư mục lưu file xuất:", self.v_out, self._pick_out)
        r2 = ttk.Frame(top)
        r2.grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=3)
        ttk.Label(r2, text="Ngày lập (dd/mm/yyyy):").pack(side="left")
        ttk.Entry(r2, textvariable=self.v_ngay, width=12).pack(side="left", padx=(6, 24))
        ttk.Label(r2, text="Tổ bộ môn:").pack(side="left")
        ttk.Entry(r2, textvariable=self.v_tbm, width=18).pack(side="left", padx=6)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=False, padx=10, pady=6)
        self._tab_dk(nb)
        self._tab_tl(nb)
        self._tab_lt(nb)
        self._tab_xtt(nb)
        self._tab_phieu(nb)

        logf = ttk.LabelFrame(self, text="Nhật ký")
        logf.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.txt = tk.Text(logf, height=9, wrap="word", state="disabled")
        sb = ttk.Scrollbar(logf, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        ttk.Label(self, text=AUTHOR + "  ·  Trường Cao đẳng Dầu khí (PV College)", style="Foot.TLabel").pack(pady=(0, 6))

    def _tab_dk(self, nb) -> None:
        f = ttk.Frame(nb)
        nb.add(f, text="1. Điều kiện dự thi")
        f.columnconfigure(1, weight=1)
        self._row_entry(f, 0, "Mã LHP:", self.v_malhp, width=30)
        lt = ttk.LabelFrame(f, text="Hình thức thi — Lý thuyết (hàng 11–12)")
        lt.grid(row=1, column=0, columnspan=4, sticky="we", padx=8, pady=4)
        for i, (txt, var) in enumerate([("Trắc nghiệm – Trên máy", self.v_tn_may),
                                        ("Trắc nghiệm – Trên giấy", self.v_tn_giay),
                                        ("Vấn đáp", self.v_van_dap), ("Tự luận", self.v_tu_luan),
                                        ("HT khác", self.v_ht_khac)]):
            ttk.Checkbutton(lt, text=txt, variable=var).grid(row=0, column=i, padx=8, pady=4, sticky="w")
        th = ttk.LabelFrame(f, text="Hình thức thi — Thực hành (hàng 11–12)")
        th.grid(row=2, column=0, columnspan=4, sticky="we", padx=8, pady=4)
        th.columnconfigure(1, weight=1)
        self._row_entry(th, 0, "Địa điểm thi:", self.v_dia_diem, width=40)
        self._row_entry(th, 1, "Thi theo ca – Số HSSV/ca:", self.v_so_ca, width=12)
        self._row_entry(th, 2, "Thi theo ca – Thời gian:", self.v_tg_ca, width=12)
        self._row_entry(th, 3, "Thi tập trung cả lớp – Thời gian (phút):", self.v_tg_tt, width=12)
        qc = ttk.Frame(f)
        qc.grid(row=3, column=0, columnspan=4, sticky="w", padx=8, pady=2)
        ttk.Label(qc, text="Quy chế xét ĐK dự thi:").pack(side="left")
        ttk.Radiobutton(qc, text="2022 (vắng > 20% tổng số tiết hoặc TBC KT < 5)", variable=self.v_quyche,
                        value="2022").pack(side="left", padx=8)
        ttk.Radiobutton(qc, text="2021 (vắng LT > 30%, vắng TH > 0 hoặc TBC KT < 5)", variable=self.v_quyche,
                        value="2021").pack(side="left", padx=8)
        self._row_entry(f, 4, "File điểm danh (tuỳ chọn):", self.v_att, self._pick_att, clear=lambda: self.v_att.set(""))
        ttk.Button(f, text="Xuất PDF – Danh sách xét điều kiện dự thi", command=self.do_dk).grid(
            row=5, column=0, columnspan=4, pady=8)
        ttk.Checkbutton(f, text="Loại HSSV không có trong file điểm danh khỏi danh sách xuất",
                        variable=self.v_drop).grid(row=6, column=0, columnspan=4, sticky="w", padx=8)

    def _tab_tl(self, nb) -> None:
        f = ttk.Frame(nb)
        nb.add(f, text="2. Thi lại")
        f.columnconfigure(1, weight=1)
        self._row_entry(f, 0, "Mã LHP:", self.v_malhp, width=30)
        self._row_entry(f, 1, "Hình thức thi LT:", self.v_ht_lt, width=50)
        self._row_entry(f, 2, "Địa điểm thi TH:", self.v_dd_th, width=50)
        ttk.Label(f, text="Ghi chú thi TH sẽ là:  “Thi thực hành tại <địa điểm thi TH>”", foreground="#555").grid(
            row=3, column=0, columnspan=3, sticky="w", padx=8)
        ttk.Button(f, text="Xuất PDF – Danh sách HSSV thi lại", command=self.do_tl).grid(
            row=4, column=0, columnspan=3, pady=10)

    def _tab_lt(self, nb) -> None:
        f = ttk.Frame(nb)
        nb.add(f, text="3. Điểm thi LT")
        f.columnconfigure(1, weight=1)
        r = ttk.Frame(f)
        r.grid(row=0, column=0, columnspan=3, sticky="w", padx=8, pady=4)
        ttk.Label(r, text="Lần thi:").pack(side="left")
        ttk.Radiobutton(r, text="Lần 1 (tất cả HSSV)", variable=self.v_lan, value=1).pack(side="left", padx=8)
        ttk.Radiobutton(r, text="Lần 2 (HSSV có điểm LT lần 1 < 5)", variable=self.v_lan, value=2).pack(side="left", padx=8)
        self._row_entry(f, 1, "Ca thi (nhập tay):", self.v_ca, width=50)
        ttk.Label(f, text="Ví dụ: Ca 4 - Ca 7 (23/1/2026). Nội dung này in vào dòng \"Ca thi:\" của bảng điểm.",
                  foreground="#555").grid(row=2, column=1, sticky="w")
        ttk.Label(f, text="⚠ LƯU Ý: Chỉ xuất File điểm thi LT khi đây là MÔ ĐUN (ô A5 ghi \"TÊN MÔ ĐUN: …\"). "
                          "Với MÔN HỌC, ứng dụng sẽ không xuất.", foreground="#B00020",
                  font=("", 10, "bold"), wraplength=780, justify="left").grid(
            row=3, column=0, columnspan=3, sticky="w", padx=8, pady=(6, 0))
        ttk.Button(f, text="Xuất PDF – Bảng điểm tổng hợp theo lớp (LT)", style="Accent.TButton",
                   command=self.do_lt).grid(row=4, column=0, columnspan=3, pady=10)

    def _tab_xtt(self, nb) -> None:
        f = ttk.Frame(nb)
        nb.add(f, text="4. Xếp thứ tự")
        ttk.Label(f, justify="left", text=(
            "Xuất 4 bảng theo điểm tổng kết (lần 1): HSSV có tổng kết ≥ 7, HSSV < 5 (kể cả HSSV không có điểm),\n"
            "Top 3 và Bottom 3 — đúng nội dung sheet TK trong Form Code VBA.")).pack(anchor="w", padx=8, pady=8)
        ttk.Button(f, text="Xuất PDF – Xếp thứ tự", command=self.do_xtt).pack(pady=10)

    def _tab_phieu(self, nb) -> None:
        f = ttk.Frame(nb)
        nb.add(f, text="5. Phiếu đánh giá TH")
        r = ttk.Frame(f)
        r.pack(anchor="w", padx=8, pady=8)
        ttk.Label(r, text="Lấy điểm kiểm tra định kỳ:").pack(side="left")
        for i in range(1, 6):
            ttk.Radiobutton(r, text=f"Bài {i}", variable=self.v_bai, value=i).pack(side="left", padx=8)
        ttk.Label(f, justify="left", foreground="#555", text=(
            "Điểm kiểm tra định kỳ được tách thành 10 tiêu chí (TC1–TC10), mỗi tiêu chí 0 / 0,5 / 1 sao cho tổng đúng bằng điểm.\n"
            "Mỗi lần xuất, cách phân bổ tiêu chí được chọn ngẫu nhiên lại. File xuất định dạng .xlsx.")).pack(
            anchor="w", padx=8)
        ttk.Button(f, text="Xuất Excel – Phiếu đánh giá kết quả thực hành", command=self.do_phieu).pack(pady=10)

    # ------------------------------------------------------------------
    def log(self, msg: str) -> None:
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _pick_in(self) -> None:
        p = filedialog.askopenfilename(title="Chọn file Form nhập điểm",
                                       filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self.v_in.set(p)
            if not self.v_out.get():
                self.v_out.set(os.path.dirname(p))

    def _pick_out(self) -> None:
        p = filedialog.askdirectory(title="Chọn thư mục lưu file xuất")
        if p:
            self.v_out.set(p)

    def _pick_att(self) -> None:
        p = filedialog.askopenfilename(title="Chọn file điểm danh",
                                       filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")])
        if p:
            self.v_att.set(p)

    # ------------------------------------------------------------------
    def _run(self, fn: Callable[[ctl.Session], str], att: bool = False, need_malhp: bool = False) -> None:
        try:
            ngay = ctl.parse_date(self.v_ngay.get())
            sess = ctl.load_session(self.v_in.get().strip(), self.v_att.get().strip() if att else "",
                                    drop_unmatched=bool(self.v_drop.get()))
            for w in sess.warnings:
                self.log("⚠ " + w)
            if need_malhp and not self.v_malhp.get().strip():
                if not messagebox.askyesno("Thiếu Mã LHP", "Mã LHP đang để trống. Vẫn xuất file?"):
                    return
            path = fn(sess)
        except ctl.Cancelled:
            self.log("Đã hủy (không ghi đè file cũ).")
            return
        except ctl.AppError as e:
            self.log("✖ " + str(e))
            messagebox.showwarning("Không xuất được", str(e))
            return
        except PermissionError:
            msg = "Không ghi được file (đang mở trong chương trình khác?). Hãy đóng file rồi thử lại."
            self.log("✖ " + msg)
            messagebox.showerror("Lỗi ghi file", msg)
            return
        except Exception as e:  # lỗi không lường trước
            self.log(f"✖ Lỗi: {e!r}")
            messagebox.showerror("Lỗi", f"Có lỗi khi xuất file:\n{e}")
            return
        self._save_config()
        self.log("✔ Đã xuất: " + path)
        if messagebox.askyesno("Hoàn tất", f"Đã xuất file:\n{path}\n\nMở file ngay?"):
            open_file(path)

    def _exists_ok(self, sess: ctl.Session, prefix: str, ext: str, extra: str = "") -> bool:
        p = ctl.target_path(self.v_out.get().strip(), sess, prefix, ctl.parse_date(self.v_ngay.get()), ext, extra)
        if os.path.exists(p):
            return messagebox.askyesno("File đã tồn tại", f"{os.path.basename(p)}\nđã có trong thư mục. Ghi đè?")
        return True

    def do_dk(self) -> None:
        def go(sess):
            opt = F.DKOptions(
                ma_lhp=self.v_malhp.get().strip(), tn_may=self.v_tn_may.get(), tn_giay=self.v_tn_giay.get(),
                van_dap=self.v_van_dap.get(), tu_luan=self.v_tu_luan.get(), ht_khac=self.v_ht_khac.get(),
                dia_diem=self.v_dia_diem.get().strip(), so_hssv_ca=self.v_so_ca.get().strip(),
                tg_ca=self.v_tg_ca.get().strip(), tg_tap_trung=self.v_tg_tt.get().strip(),
                quy_che=self.v_quyche.get(), to_bo_mon=self.v_tbm.get().strip().upper(),
                ngay_lap=ctl.parse_date(self.v_ngay.get()))
            if not self._exists_ok(sess, "DSDT", "pdf"):
                raise ctl.Cancelled()
            return ctl.export_dieu_kien(sess, self.v_out.get().strip(), opt)
        self._run(go, att=True, need_malhp=True)

    def do_tl(self) -> None:
        def go(sess):
            opt = F.TLOptions(ma_lhp=self.v_malhp.get().strip(), hinh_thuc_lt=self.v_ht_lt.get().strip(),
                              dia_diem_th=self.v_dd_th.get().strip(), to_bo_mon=self.v_tbm.get().strip().upper(),
                              ngay_lap=ctl.parse_date(self.v_ngay.get()))
            if not self._exists_ok(sess, "DSTL", "pdf"):
                raise ctl.Cancelled()
            return ctl.export_thi_lai(sess, self.v_out.get().strip(), opt)
        self._run(go, need_malhp=True)

    def do_lt(self) -> None:
        def go(sess):
            lan = self.v_lan.get()
            if sess.meta.la_mo_dun and not self.v_ca.get().strip():
                if not messagebox.askyesno("Thiếu Ca thi", "Mục Ca thi đang để trống. Vẫn xuất file?"):
                    raise ctl.Cancelled()
            opt = F.LTOptions(lan_thi=lan, ca_thi=self.v_ca.get().strip(), to_bo_mon=self.v_tbm.get().strip().upper(),
                              ngay_lap=ctl.parse_date(self.v_ngay.get()))
            if not self._exists_ok(sess, "LT", "pdf", "L2" if lan == 2 else ""):
                raise ctl.Cancelled()
            return ctl.export_diem_lt(sess, self.v_out.get().strip(), opt)
        self._run(go)

    def do_xtt(self) -> None:
        def go(sess):
            if not self._exists_ok(sess, "XTT", "pdf"):
                raise ctl.Cancelled()
            return ctl.export_xep_thu_tu(sess, self.v_out.get().strip(), ctl.parse_date(self.v_ngay.get()))
        self._run(go)

    def do_phieu(self) -> None:
        def go(sess):
            bai = self.v_bai.get()
            if not self._exists_ok(sess, "PDGTH", "xlsx", f"Bai{bai}"):
                raise ctl.Cancelled()
            path, n = ctl.export_phieu(sess, self.v_out.get().strip(), bai, ctl.parse_date(self.v_ngay.get()))
            self.log(f"  Bài {bai}: {n} HSSV có điểm.")
            return path
        self._run(go)


def main() -> None:
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    while True:
        login = LoginWindow()
        login.mainloop()
        if not login.user:        # đóng cửa sổ đăng nhập -> thoát
            return
        app = App(login.user)
        app.relogin = False
        try:
            icon = resource_path("assets", "logo_72.png")
            app._icon = tk.PhotoImage(file=icon)
            app.iconphoto(True, app._icon)
        except Exception:
            pass
        app.mainloop()
        if not getattr(app, "relogin", False):
            return
