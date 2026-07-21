@echo off
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-skill-sync.ps1"
if errorlevel 1 (
  echo.
  echo 安装失败，请保留此窗口并联系插件维护人员。
) else (
  echo.
  echo 安装完成，请完全退出并重新打开 Codex 或 WorkBuddy。
)
pause
