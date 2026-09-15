@echo off
chcp 65001 >nul
cd /d "D:\data"
start /min "Auto Push GitHub" powershell -NoProfile -ExecutionPolicy Bypass -File "D:\data\auto_push.ps1"
