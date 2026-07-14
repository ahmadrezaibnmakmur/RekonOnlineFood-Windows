#!/usr/bin/env python3
"""
Build Rekon Online Food for Windows.

Run this file on Windows:
    py -m pip install -r requirements.txt pyinstaller
    py build_windows.py

Optional installer build:
    Install Inno Setup, then run this script again. It will create:
    dist/installer/RekonOnlineFoodSetup-1.03.exe
"""

import os
import platform
import shutil
import subprocess
import sys

from app_version import APP_VERSION


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(SCRIPT_DIR, "webapp")
APP_PY = os.path.join(APP_DIR, "app.py")
REKON_PY = os.path.join(SCRIPT_DIR, "rekon.py")
TEMPLATES_DIR = os.path.join(APP_DIR, "templates")
STATIC_DIR = os.path.join(APP_DIR, "static")
STORE_MAPPING = os.path.join(SCRIPT_DIR, "store_mapping.json")
APP_VERSION_FILE = os.path.join(SCRIPT_DIR, "app_version.py")
DIST_DIR = os.path.join(SCRIPT_DIR, "dist")
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")
INSTALLER_SCRIPT = os.path.join(SCRIPT_DIR, "installer", "RekonOnlineFood.iss")
APP_BUILD_NAME = f"RekonOnlineFood-{APP_VERSION}"
SETUP_NAME = f"RekonOnlineFoodSetup-{APP_VERSION}.exe"


def require_windows():
    if platform.system() != "Windows":
        print("ERROR: build_windows.py harus dijalankan di Windows.")
        print("PyInstaller tidak bisa membuat Windows .exe dari macOS.")
        sys.exit(1)


def check_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("ERROR: PyInstaller belum terinstall.")
        print("Jalankan: py -m pip install pyinstaller")
        sys.exit(1)


def clean():
    for path in [
        BUILD_DIR,
        os.path.join(DIST_DIR, APP_BUILD_NAME),
        os.path.join(DIST_DIR, "installer", SETUP_NAME),
    ]:
        if os.path.exists(path):
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)


def pyinstaller_args():
    args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name",
        APP_BUILD_NAME,
        "--noconfirm",
        "--clean",
        "--windowed",
        "--onedir",
    ]

    data_files = [
        (TEMPLATES_DIR, "webapp/templates"),
        (REKON_PY, "."),
        (STORE_MAPPING, "."),
        (APP_VERSION_FILE, "."),
    ]
    if os.path.exists(STATIC_DIR):
        data_files.append((STATIC_DIR, "webapp/static"))

    for src, dest in data_files:
        if os.path.exists(src):
            args.extend(["--add-data", f"{src}{os.pathsep}{dest}"])

    args.extend(
        [
            "--hidden-import",
            "pandas",
            "--hidden-import",
            "openpyxl",
            "--hidden-import",
            "flask",
            "--hidden-import",
            "jinja2",
            "--hidden-import",
            "werkzeug",
            "--hidden-import",
            "tkinter",
        ]
    )
    args.append(APP_PY)
    return args


def build_exe():
    print(f"[1/3] Building {APP_BUILD_NAME}.exe...")
    subprocess.run(pyinstaller_args(), cwd=SCRIPT_DIR, check=True)
    exe_path = os.path.join(DIST_DIR, APP_BUILD_NAME, f"{APP_BUILD_NAME}.exe")
    if not os.path.exists(exe_path):
        print(f"ERROR: Output tidak ditemukan: {exe_path}")
        sys.exit(1)
    print(f"OK: {exe_path}")


def find_inno_compiler():
    env_path = os.environ.get("INNO_SETUP_COMPILER")
    candidates = [
        env_path,
        shutil.which("iscc"),
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def build_installer():
    print("[2/3] Looking for Inno Setup...")
    compiler = find_inno_compiler()
    if not compiler:
        print("ERROR: Inno Setup tidak ditemukan.")
        print(f"Install Inno Setup 6 untuk membuat dist\\installer\\{SETUP_NAME}")
        sys.exit(1)

    print("[3/3] Building installer...")
    subprocess.run([compiler, INSTALLER_SCRIPT], cwd=SCRIPT_DIR, check=True)
    setup_path = os.path.join(DIST_DIR, "installer", SETUP_NAME)
    if not os.path.exists(setup_path):
        print(f"ERROR: Installer tidak ditemukan: {setup_path}")
        sys.exit(1)
    print(f"OK: {setup_path}")


def main():
    require_windows()
    check_pyinstaller()
    clean()
    build_exe()
    build_installer()


if __name__ == "__main__":
    main()
