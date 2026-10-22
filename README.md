# Xuất biểu mẫu từ Form nhập điểm (Python – Windows & macOS)

Ứng dụng đọc file **Form nhập điểm.xlsx** (sheet `DIEM TK`), tính điểm theo đúng các macro VBA trong
**Form Code VBA.xlsm** rồi xuất:

| Tab | Xuất ra | Theo sheet / macro VBA |
|---|---|---|
| 1. Điều kiện dự thi | PDF hoặc xlsx | `DK DT` – `Copy_Select_Data`, `Xet_DK_2021/2022`, `PDF_DK` |
| 2. Thi lại | PDF hoặc xlsx | `TL` – `THIL`, `PDF_TL` |
| 3. Điểm thi LT | PDF hoặc xlsx | `LT` – `LYT` (lần 1), `LYT2` (lần 2), `INLT` |
| 4. Xếp thứ tự | PDF hoặc xlsx | `TK` – `SQL_lon7`, `SQL_nho5`, `SQL_3lonnhat`, `SQL_3nhonhat` |
| 5. Phiếu đánh giá TH | **xlsx** | `DIEM TC` – `Bai1…Bai5`, `Diem_TC` |

Không dùng Excel/VBA/COM nên chạy giống nhau trên Windows và macOS. File PDF được ghi trực tiếp bằng
Python (reportlab) — không còn lỗi `ExportAsFixedFormat`/AppleScript trên Mac.

## Chạy thử từ mã nguồn
```
pip install -r requirements.txt
python main.py
```
App mở **giao diện web cục bộ** (địa chỉ `http://127.0.0.1:<cổng>/`, tự mở trình duyệt, kiểu "Valve Test Calculator"):
đăng nhập → chọn file Form nhập điểm (bấm hoặc kéo thả) → chọn tab → **Xuất**. File xuất ghi vào "Thư mục lưu"
(mặc định `~/Downloads/NhapDiem_xuat`), có nút Xem PDF / Mở file / Mở thư mục / Tải về. Đóng hết tab trình duyệt thì
app tự thoát sau ~2,5 phút, hoặc bấm nút **Thoát**. Dữ liệu chỉ xử lý trên máy, không gửi ra Internet.
Muốn dùng giao diện cửa sổ cũ (tkinter): `python main.py --tk`.
Python 3.10+ là đủ (giao diện web không cần Tk). Kiểm thử: `python tests/run_tests.py`.

## Đóng gói thành .exe / .app
* **Windows:** chạy `build_win.bat` → `dist\NhapDiem.exe`
* **macOS:** chạy `./build_mac.sh` → `dist/NhapDiem.app` (lần đầu mở: chuột phải → Open vì app chưa ký số)
* Hoặc đẩy thư mục lên GitHub, vào **Actions → Build app → Run workflow** để build cả 2 hệ điều hành,
  tải file ở mục *Artifacts*.

(Một file .exe chỉ chạy trên Windows, một file .app chỉ chạy trên Mac — phải build riêng trên từng hệ điều hành.)

## Cách dùng
0. Mở app sẽ hiện màn hình **Đăng nhập** có logo Trường (tài khoản `leduynam`). Sau đăng nhập, thanh tiêu đề có logo, tên tổ bộ môn, tài khoản và nút **Đăng xuất**. Mật khẩu chỉ lưu dạng băm PBKDF2 trong `nhapdiem/auth.py` (khoá mềm — người rành code vẫn có thể gỡ).
1. Chọn **File Form nhập điểm** và **Thư mục lưu** (phần "Dùng chung"), kiểm tra Ngày lập / Tổ bộ môn.
2. Vào tab cần xuất, nhập các ô, bấm **Xuất**. Tên file giống VBA: `ddmmyy_DSDT_<MÔN KHÔNG DẤU>_<LỚP-KHÓA>.pdf`.

**Tab 1** – Mã LHP; tick hình thức thi LT (Trên máy / Trên giấy / Vấn đáp / Tự luận / HT khác → ô hàng 12 hiện `X`
hoặc `-`); Địa điểm thi, Thi theo ca (số HSSV/ca, thời gian), Thi tập trung (phút). Quy chế xét ĐK 2021/2022.
Cột "% tiết vắng" chỉ có số khi chọn *File điểm danh* (`Form_Diem_danh.xlsx`: hàng 10 = LT/TH, hàng 11 = giờ của buổi (mỗi buổi tính 1), hàng 12+ = X / P / KP; P và KP đều tính là vắng). Tự chọn sheet (CKT/LT…) có nhiều tên HSSV khớp nhất. % vắng theo macro `Tinhnghihoc`: số buổi vắng / tổng số buổi (ô có số ở hàng 11), ghép tên không dấu. HSSV không có trong file điểm danh để trống và được báo cảnh báo.

**Tab 2** – Mã LHP, Hình thức thi LT, **Địa điểm thi TH** (ghi chú thành "Thi thực hành tại …").

**Tab 3** – Lần thi 1 (tất cả) hoặc 2 (HSSV có điểm LT lần 1 < 5), Ca thi.

**Tab 5** – chọn Bài 1–5 (= cột KT định kỳ 1–5). Mỗi điểm được tách thành 10 tiêu chí 0/0,5/1.

## Dữ liệu đầu vào cần có
Sheet `DIEM TK` giữ nguyên bố cục: A5 (tên môn), O5/Q5 (số giờ, LT/TH), B6/E6 (lớp, khóa), HSSV từ hàng 10,
cột C–N (điểm KT), P/Q (thi lần 1 LT/TH), T/U (thi lần 2), **AA = Mã số HSSV, AE = Ngày sinh** (nếu trống thì
cột tương ứng trong file xuất để trống).

## Khác biệt có chủ ý so với VBA
* Điểm KT định kỳ 6 (cột J) được chép sang DK DT (VBA cũ cắt mất do `Resize(…, 8)`).
* Phiếu TH: VBA lặp ngẫu nhiên tới khi tổng = điểm nên **treo** với điểm như 5,7; bản này cộng phần lẻ vào 1 tiêu chí.
* Bảng "< 5" và "Bottom 3": HSSV điểm tổng kết `-` (học lại) đứng đầu bảng với điểm hiển thị `-`, đúng ý đồ của macro (code VBA gán "-" cho nhóm này).
* Sửa lỗi chính tả "NÔI DUNG THI LẠI" → "NỘI DUNG THI LẠI" và dấu ngoặc ở "% Tiết vắng học (…)".
* Danh sách thi lại sắp theo họ tên giảm dần (chuỗi nhị phân) đúng như `Order by F2 Desc` của VBA — chữ "Đ" đứng trước "L".

Font: dùng Times New Roman của máy nếu có, nếu không dùng Liberation Serif đi kèm (SIL OFL, thư mục `nhapdiem/fonts`).

## Định dạng xuất & Footer (tab 1–4)
- Khung "Định dạng xuất & Footer" cho chọn **PDF** hoặc **Excel (.xlsx)** cho 4 file Điều kiện dự thi / Thi lại / Điểm thi LT / Xếp thứ tự.
- Nút **Cập nhật Footer** lưu: Mã biểu mẫu từng file (mặc định BM/CL/05.01, 05.02, 05.04a; Xếp thứ tự để trống), LBH, NHL.
- Footer in ở chân mỗi trang: `Mã biểu mẫu | LBH … NHL … | Trang x/y` (Times New Roman 12) — cả PDF lẫn xlsx (xlsx dùng footer in của Excel, khổ A4, tiêu đề bảng lặp lại mỗi trang).

## Tab 6 · Nhập điểm hệ thống
Chọn file mẫu của hệ thống (Form nhập quá trình / Form nhập thi kết thúc mô đun, **đã lưu thành .xlsx**) → app điền điểm
từ Form nhập điểm theo Mã sinh viên rồi xuất file mới (giữ nguyên định dạng mẫu). Cột được nhận bằng mã ở hàng 6
(DiemHeSo11/12/21–25, DiemTBThuongKy, DuocDuThiKetThuc, DiemTongKet, DiemThiKN1 = LT, DiemThiKN2 = TH).

## Đóng gói 1 file chạy
- Windows: chạy `build_win.bat` → `dist\Phan_mem_Nhap_Diem_LDN_PV_College.exe` (bấm đúp, tự mở trình duyệt).
- macOS: chạy `build_mac.sh` → `dist/Phan_mem_Nhap_Diem_LDN_PV_College.app`.
- Hoặc đẩy lên GitHub → Actions → "Build app" để tự build cả hai.
