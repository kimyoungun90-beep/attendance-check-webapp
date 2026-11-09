@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js가 필요합니다. 또는 index.html을 브라우저로 여세요.
  pause
  exit /b 1
)
start "" "http://127.0.0.1:8766/"
node server.cjs
pause
