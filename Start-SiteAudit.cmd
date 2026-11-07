@echo off
REM Double-click this file to run the SharePoint site audit.
where pwsh >nul 2>nul
if errorlevel 1 (
  echo PowerShell 7 is not installed. Get it from https://aka.ms/powershell and try again.
  pause
  exit /b 1
)
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-SiteAudit.ps1"
