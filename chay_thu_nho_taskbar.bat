@echo off
chcp 65001 >nul
cd /d "D:\Antigravity Tai\Github\data"
start /min "Auto Push GitHub" powershell -NoProfile -ExecutionPolicy Bypass -File "D:\Antigravity Tai\Github\data\auto_push.ps1"
