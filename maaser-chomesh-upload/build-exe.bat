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

echo Installing dependencies...
call npm install || goto :fail

echo Building portable EXE...
call npm run build:exe || goto :fail

echo Build completed.
echo Check dist folder for the EXE file.
pause
exit /b 0

:fail
echo Build failed.
pause
exit /b 1
