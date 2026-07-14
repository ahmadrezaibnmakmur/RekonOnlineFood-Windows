# Build Windows Installer

Jalankan langkah ini di Windows.

## Untuk user akhir

User akhir tidak perlu buka terminal, tidak perlu install Python, dan tidak perlu menjalankan command.

Share file ini saja:

```text
dist\installer\RekonOnlineFoodSetup-1.01.exe
```

User cukup:

1. Download `RekonOnlineFoodSetup-1.01.exe`.
2. Double-click untuk install.
3. Klik shortcut `Rekon Online Food`.

Browser akan terbuka otomatis ke aplikasi.

## Untuk pembuat installer

Jalankan langkah ini di Windows, bukan di laptop user akhir.

Cara paling mudah: double-click file ini dari folder project:

```text
Build Installer Windows.bat
```

Untuk membuat installer `.exe`, Inno Setup 6 wajib terinstall:
https://jrsoftware.org/isinfo.php

Alternatif via terminal:

```bat
py -m pip install -r requirements.txt pyinstaller
py build_windows.py
```

Output:

- App standalone: `dist\RekonOnlineFood-1.01\RekonOnlineFood-1.01.exe`
- Installer: `dist\installer\RekonOnlineFoodSetup-1.01.exe`

Jika Inno Setup belum terinstall, build akan berhenti supaya tidak ada installer lama yang tidak sengaja diupload.

## Test user journey

1. Jalankan `dist\installer\RekonOnlineFoodSetup-1.01.exe`.
2. Buka aplikasi dari Start Menu atau Desktop shortcut.
3. Browser akan terbuka ke `http://localhost:8080`.
4. Pilih folder project yang berisi `Raw Data Transaksi`.
   Struktur data standar:
   `Raw Data Transaksi\Download ERP\YYYY-MM-DD`,
   `Raw Data Transaksi\Grabfood\YYYY-MM-DD`,
   `Raw Data Transaksi\GoFood\YYYY-MM-DD`,
   `Raw Data Transaksi\ShopeeFood\YYYY-MM-DD`.
5. Proses rekonsiliasi dan export Excel.

## Catatan

- Build Windows harus dijalankan di Windows, bukan macOS.
- Aplikasi tidak membutuhkan Python di komputer user setelah diinstall.
- Jika aplikasi sudah berjalan dan shortcut diklik lagi, aplikasi hanya membuka browser ke instance yang sudah aktif.
