@echo off
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed. Please install Node.js 20+ first.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo npm is not available. Please reinstall Node.js.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing dependencies...
  call npm install || goto :fail
)

echo Starting app...
call npm run start || goto :fail
exit /b 0

:fail
echo Operation failed.
pause
exit /b 1
