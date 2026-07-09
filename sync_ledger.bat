@echo off
chcp 65001 > nul
echo ================================
echo   台帳同期：Excel → GitHub
echo ================================
cd /d "C:\Users\Y.Kinomura\OneDrive - 株式会社ハシマ\アプリ\claude_code"
python sync_ledger.py
echo.
pause
