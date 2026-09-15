@echo off
chcp 65001 >nul
title Kiểm tra tiến trình tự động Push
echo ========================================================
echo KIỂM TRA TIẾN TRÌNH AUTO PUSH
echo ========================================================
tasklist /fi "imagename eq powershell.exe" | findstr /i "powershell" >nul
if %errorlevel% equ 0 (
    echo [OK] Script PowerShell DANG CHAY NGAM!
) else (
    echo [CANH BAO] Script hien KHONG CHAY!
)
echo.
echo ========================================================
echo NHẬT KÝ HOẠT ĐỘNG (10 DÒNG GẦN NHẤT):
echo ========================================================
if exist "D:\Antigravity Tai\Github\data\auto_push.log" (
    powershell -NoProfile -Command "Get-Content 'D:\Antigravity Tai\Github\data\auto_push.log' -Tail 10"
) else (
    echo Chưa có file nhật ký auto_push.log
)
echo ========================================================
pause
