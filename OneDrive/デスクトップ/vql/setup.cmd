@echo off
REM Bypasses PowerShell blocking npm.ps1 (ExecutionPolicy). Same as: npm run setup
setlocal
cd /d "%~dp0"
where npm.cmd >nul 2>nul
if errorlevel 1 (
  echo npm.cmd not found. Install Node.js from https://nodejs.org/ and reopen the terminal.
  exit /b 1
)
call npm.cmd run setup
exit /b %ERRORLEVEL%
