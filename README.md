# Rekon Online Food Windows Build Repo

Repo ini khusus untuk build Windows `.exe` dan installer lewat GitHub Actions.

## Isi repo

- source app: `rekon.py`, `webapp/`, `store_mapping.json`
- build script Windows: `build_windows.py`
- installer script: `installer/RekonOnlineFood.iss`
- workflow GitHub Actions: `.github/workflows/build-windows.yml`

## Journey

1. Push repo ini ke GitHub.
2. Buka tab `Actions`.
3. Jalankan workflow `Build Windows Installer`, atau biarkan jalan otomatis saat push ke `main`.
4. Tunggu job selesai.
5. Download artifact:
   - `RekonOnlineFoodSetup-1.02`
   - `RekonOnlineFood-1.02-portable`

## Output

- installer: `dist/installer/RekonOnlineFoodSetup-1.02.exe`
- app portable: `dist/RekonOnlineFood-1.02/`

## Fallback manual di Windows

Kalau mau build manual di laptop Windows:

1. Install Python.
2. Install Inno Setup 6.
3. Jalankan `BUAT INSTALLER - DOUBLE CLICK.bat`.
