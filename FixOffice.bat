@echo off
:: Mengecek hak akses Administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    goto :run_fix
) else (
    echo =====================================================
    echo ERROR: Harap jalankan file ini sebagai Administrator!
    echo Klik kanan file ini lalu pilih 'Run as administrator'.
    echo =====================================================
    pause
    exit
)

:run_fix
title Perbaikan Error Office 0xc0000142
echo Menyiapkan proses perbaikan...
echo.

:: 1. Menghentikan layanan ClickToRunSvc
echo [1/2] Menghentikan Layanan Microsoft Office (ClickToRunSvc)...
net stop ClickToRunSvc /y
timeout /t 2 >nul

:: 2. Menjalankan kembali layanan ClickToRunSvc
echo [2/2] Menjalankan kembali Layanan Microsoft Office...
net start ClickToRunSvc

if %errorLevel% == 0 (
    echo.
    echo =====================================================
    echo BERHASIL: Layanan Office telah direstart.
    echo Silakan coba buka kembali Word atau Excel Anda.
    echo =====================================================
) else (
    echo.
    echo Terjadi kesalahan saat mencoba menjalankan layanan.
    echo Pastikan Microsoft Office terinstal dengan benar.
)

pause
exit
