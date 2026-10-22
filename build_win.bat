@echo off
REM Dong goi thanh Phan_mem_Nhap_Diem_LDN_PV_College.exe (1 file, chay tren Windows, Python 3.10+)
python -m pip install -r requirements.txt || goto :err
python -m PyInstaller --noconfirm --clean --onefile --name Phan_mem_Nhap_Diem_LDN_PV_College ^
  --add-data "nhapdiem\fonts;nhapdiem\fonts" --add-data "nhapdiem\assets;nhapdiem\assets" --icon nhapdiem\assets\icon.ico main.py || goto :err
echo.
echo Xong: dist\Phan_mem_Nhap_Diem_LDN_PV_College.exe  (bam doup de chay - tu mo trinh duyet)
exit /b 0
:err
echo Co loi khi dong goi.
exit /b 1
