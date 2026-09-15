@echo off
chcp 65001 >nul
title Tự động đồng bộ dữ liệu GitHub
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_push.ps1"
pause
