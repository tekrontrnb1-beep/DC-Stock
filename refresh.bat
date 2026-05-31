@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   DC-Stock — Дата шинэчлэх
echo ============================================
echo.
echo import\ доторх файлуудаас data\stock-data.enc уусгэж байна...
echo.
python build_data.py
echo.
if errorlevel 1 (
  echo [!] Алдаа гарлаа. Дээрх мессежийг шалгана уу.
) else (
  echo ============================================
  echo   Бэлэн! Дараа нь нийтлэхийн тулд:
  echo     git add -A
  echo     git commit -m "data update"
  echo     git push
  echo ============================================
)
echo.
pause
