@echo off
chcp 65001 >nul
where py >nul 2>nul
if errorlevel 1 goto powershell_fallback
py -3 "%~dp0publish-skills.py" %*
goto finished

:powershell_fallback
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0publish-skills.ps1" %*

:finished
if errorlevel 1 (
  echo.
  echo 发布失败，请检查上方提示。
)
pause
