@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Please install Node.js 22.12 or newer from https://nodejs.org/
  pause
  exit /b 1
)
if not exist "node_modules\vite\bin\vite.js" (
  echo Installing dependencies for the first run...
  call npm.cmd ci --cache .npm-cache --no-fund
  if errorlevel 1 (
    pause
    exit /b 1
  )
)
echo Starting Word Merge Studio. Busy ports are skipped automatically.
echo The actual local URL will appear below and open in your browser.
echo Keep this window open. Press Ctrl+C to stop.
call npm.cmd run dev -- --open
if errorlevel 1 pause
