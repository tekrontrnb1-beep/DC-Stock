@echo off
cd /d "%~dp0"
echo ============================================
echo    DC-Stock  -  push to GitHub
echo ============================================
echo.
echo 1) First create an EMPTY public repo "DC-Stock" at https://github.com/new
echo 2) Paste your GitHub token below (mouse right-click = paste), then Enter
echo.
set /p TOKEN=GitHub token:
if "%TOKEN%"=="" goto NOTOKEN
echo.
echo Committing...
git add -A
git commit -m "DC-Stock dashboard update"
echo.
echo Pushing... please wait
git -c credential.helper= push "https://tekrontrnb1-beep:%TOKEN%@github.com/tekrontrnb1-beep/DC-Stock.git" main
set RC=%ERRORLEVEL%
set "TOKEN="
git remote remove origin 1>nul 2>nul
git remote add origin "https://github.com/tekrontrnb1-beep/DC-Stock.git"
echo.
if "%RC%"=="0" goto OK
goto FAIL

:OK
echo ============================================
echo    SUCCESS!  Now turn on GitHub Pages:
echo    repo  Settings  ^>  Pages  ^>  Branch: main  ^>  folder: root  ^>  Save
echo    Then open:  https://tekrontrnb1-beep.github.io/DC-Stock/
echo ============================================
goto END

:FAIL
echo [X] Push failed. Check:
echo     - Did you create the "DC-Stock" repo on GitHub?
echo     - Is the token valid and has Contents:write permission?
goto END

:NOTOKEN
echo Token was empty. Cancelled.

:END
echo.
pause
