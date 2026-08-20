@echo off
title Server Report Sistem PG2
echo ====================================================
echo   Menjalankan Server Report Sistem PG2...
echo   Mohon JANGAN TUTUP jendela hitam ini selama 
echo   Anda masih menggunakan websitenya.
echo ====================================================
echo.

:: Memastikan semua library pendukung sistem sudah terinstal (termasuk server Waitress & engine Calamine)
echo Menyiapkan mesin server dan mengecek kelengkapan sistem...
python -m pip install -r requirements.txt -q

:: Menjalankan server flask di belakang layar pada terminal ini
start /b python app.py

:: Menunggu 3 detik agar server siap
timeout /t 3 /nobreak >nul

:: Membuka browser otomatis
echo Membuka browser...
start http://127.0.0.1:5000

:: Menjaga jendela tetap terbuka agar user bisa melihat log/error (opsional)
:: Tapi karena start /b, proses nempel di cmd ini
pause
