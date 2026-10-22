#!/bin/bash
# Dong goi thanh Phan_mem_Nhap_Diem_LDN_PV_College.app (chay tren macOS)
set -e
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
python3 -m PyInstaller --noconfirm --clean --windowed --name Phan_mem_Nhap_Diem_LDN_PV_College --icon nhapdiem/assets/icon.icns \
  --add-data "nhapdiem/fonts:nhapdiem/fonts" --add-data "nhapdiem/assets:nhapdiem/assets" main.py
echo
echo "Xong: dist/Phan_mem_Nhap_Diem_LDN_PV_College.app"
