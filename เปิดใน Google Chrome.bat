@echo off
setlocal

set "CHROME="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if defined CHROME (
  start "" "%CHROME%" "%~dp0index.html"
  exit /b 0
)

where chrome.exe >nul 2>nul
if not errorlevel 1 (
  start "" chrome.exe "%~dp0index.html"
  exit /b 0
)

echo ไม่พบ Google Chrome ในเครื่อง กรุณาติดตั้ง Google Chrome แล้วลองอีกครั้ง
pause
