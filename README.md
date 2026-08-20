\\

# Report Sistem PG2

Aplikasi web internal untuk memproses laporan absensi secara otomatis.

## Fitur
- Pemrosesan Excel otomatis (mencocokkan nama mandor, jabatan, bagian).
- Generate Pivot Summary untuk kalkulasi kegagalan absen.
- Antarmuka web profesional berbasis Flask.
- Unduhan file otomatis.

## Cara Menjalankan
1. Pastikan Python sudah terinstal.
2. Instal pustaka yang dibutuhkan:
   ```bash
   pip install Flask Werkzeug openpyxl pandas
   ```
3. Jalankan server Flask:
   ```bash
   python app.py
   ```
4. Buka browser di `http://127.0.0.1:5000`
