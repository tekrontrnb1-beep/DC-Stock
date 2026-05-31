@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==================================================
echo   DC-Stock - GitHub-d push hiih
echo ==================================================
echo.
echo  Эхлээд GitHub дээр "DC-Stock" нэртэй PUBLIC repo
echo  үүсгэсэн байх ёстой:  https://github.com/new
echo.
echo  Доор GitHub token-оо буулгана уу (Right-click = Paste).
echo.
set /p TOKEN=GitHub Token:
if "%TOKEN%"=="" ( echo Token хоосон байна. Цуцлав. & pause & exit /b 1 )
echo.
echo --- commit ---
git add -A
git commit -m "DC-Stock dashboard update"
echo.
echo --- push (хэдэн секунд хүлээнэ) ---
git -c credential.helper= push -u "https://tekrontrnb1-beep:%TOKEN%@github.com/tekrontrnb1-beep/DC-Stock.git" main
set "RC=%ERRORLEVEL%"
set "TOKEN="
echo.
if "%RC%"=="0" (
  echo ==================================================
  echo   AMJILTTAI! Одоо GitHub repo дээр:
  echo   Settings - Pages - Branch: main / (root) - Save
  echo   Дараа нь: https://tekrontrnb1-beep.github.io/DC-Stock/
  echo ==================================================
) else (
  echo [!] Push амжилтгүй. Шалгах зүйлс:
  echo     - "DC-Stock" repo үүсгэсэн үү?
  echo     - Token зөв, эрх ^(Contents: write^) хүрэлцэх үү?
)
echo.
pause
