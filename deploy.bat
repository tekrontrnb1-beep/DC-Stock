@echo off
cd /d "%~dp0"
echo ============================================
echo    DC-Stock  -  rebuild + publish to GitHub
echo ============================================
echo.
set "TOKEN="
if exist token.txt set /p TOKEN=<token.txt
if not "%TOKEN%"=="" goto BUILD
echo Paste your GitHub token below (one time only - it will be saved
echo to token.txt so next time you just double-click, no typing).
echo.
set /p TOKEN=GitHub token:
if "%TOKEN%"=="" goto NOTOKEN
> token.txt echo %TOKEN%
echo Saved. Next time no typing needed.

:BUILD
echo.
echo [1/2] Rebuilding data from import\ folder ...
python build_data.py
if errorlevel 1 goto BUILDFAIL
echo.
echo [2/2] Pushing to GitHub ...
git add -A
git commit -m "DC-Stock update"
git -c credential.helper= push "https://tekrontrnb1-beep:%TOKEN%@github.com/tekrontrnb1-beep/DC-Stock.git" main
set RC=%ERRORLEVEL%
set "TOKEN="
git remote remove origin 1>nul 2>nul
git remote add origin "https://github.com/tekrontrnb1-beep/DC-Stock.git"
echo.
if "%RC%"=="0" goto OK
goto PUSHFAIL

:OK
echo ============================================
echo    DONE!  Live in ~1-2 min:
echo    https://tekrontrnb1-beep.github.io/DC-Stock/
echo ============================================
goto END

:BUILDFAIL
echo [X] build_data.py failed - check the files in import\ . Nothing was pushed.
goto END

:PUSHFAIL
echo [X] Push failed - check token (needs Contents:write) and that the repo exists.
echo     If the token changed, delete token.txt and run again.
goto END

:NOTOKEN
echo Token was empty. Cancelled.

:END
echo.
pause
