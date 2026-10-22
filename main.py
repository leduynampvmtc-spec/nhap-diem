"""Điểm vào của ứng dụng (cũng là file PyInstaller đóng gói thành .exe / .app).

Mặc định chạy giao diện web cục bộ (mở trình duyệt). Dùng  python main.py --tk  để mở giao diện cửa sổ cũ (tkinter).
"""
import sys

if __name__ == "__main__":
    if "--tk" in sys.argv:
        from nhapdiem.gui import main
        main()
    else:
        from nhapdiem.webapp import main
        main(open_browser="--no-browser" not in sys.argv)
