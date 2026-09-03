@echo off
chcp 65001 >nul
title PEA Scrap Metal Return Calculator - Provincial Electricity Authority
cls
echo ======================================================================
echo    ระบบคำนวณการคืนพัสดุประเภทเศษเหล็ก การไฟฟ้าส่วนภูมิภาค (PEA)
echo ======================================================================
echo.

:: Detect Python executable
set PYTHON_CMD=
where py >nul 2>&1
if %ERRORLEVEL% equ 0 (
    set PYTHON_CMD=py
) else (
    where python >nul 2>&1
    if %ERRORLEVEL% equ 0 (
        set PYTHON_CMD=python
    )
)

if "%PYTHON_CMD%"=="" (
    echo [ERROR] ไม่พบโปรแกรม Python หรือ Python Launcher (py) บนเครื่องนี้
    echo กรุณาติดตั้ง Python จาก https://www.python.org ก่อนใช้งาน
    echo.
    pause
    exit /b 1
)

echo [OK] ตรวจพบตัวประมวลผล: %PYTHON_CMD%
echo.
echo กำลังเริ่มต้น Web Application บน http://localhost:8080 ...
echo * หมายเหตุ: กรุณาอย่าปิดหน้าต่างนี้ขณะใช้งานระบบ *
echo.

:: Launch browser after 1.5 seconds delay in background so server is ready
start "" powershell -WindowStyle Hidden -Command "Start-Sleep -Milliseconds 1500; Start-Process 'http://localhost:8080'"

:: Start the Python server in foreground
%PYTHON_CMD% local_server.py

echo.
echo เซิร์ฟเวอร์หยุดทำงานแล้ว
pause
