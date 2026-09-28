@echo off
REM Double-click: converts everything in the "input" folder.
REM Drag-and-drop: drop PowerPoint files or folders onto this file to convert them.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-PptToPdf.ps1" %*
echo.
pause
