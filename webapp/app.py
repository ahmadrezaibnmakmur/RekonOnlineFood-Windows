#!/usr/bin/env python3
"""
Rekon Online Food - Web Interface
==================================
Flask-based web app for reconciling ERP transactions with online food platforms.

Usage:
    python3 app.py
    Then open http://localhost:5000 in browser
"""

import os
import sys
import json
import warnings
import urllib.request
from datetime import datetime, date
from pathlib import Path
from collections import defaultdict

warnings.filterwarnings("ignore")

IS_FROZEN = getattr(sys, "frozen", False)
BUNDLE_DIR = getattr(
    sys,
    "_MEIPASS",
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
)

# Add bundled/source app root to path for rekon module.
sys.path.insert(0, BUNDLE_DIR)

from app_version import APP_VERSION

try:
    from flask import Flask, render_template, request, jsonify, send_file
except ImportError:
    print("ERROR: Flask belum terinstall. Jalankan: pip3 install flask")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas belum terinstall. Jalankan: pip3 install pandas openpyxl")
    sys.exit(1)

import rekon

# Import reconciliation logic from rekon.py
from rekon import (
    STORE_MAP, ERP1_TO_FOLDER, ERP2_TO_FOLDER,
    find_erp_files, load_erp_penerimaan, load_erp_transaksi,
    find_platform_reports, load_grabfood_reports, load_gofood_reports,
    load_shopeefood_reports,
    reconcile, generate_summary_per_store_day, generate_summary_per_platform, generate_detail,
    export_to_excel,
)

TEMPLATE_DIR = os.path.join(BUNDLE_DIR, "webapp", "templates")
app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max

# Default project path
if IS_FROZEN:
    DEFAULT_PROJECT_PATH = os.path.join(
        os.path.expanduser("~"), "Documents", "Rekon Online Food"
    )
else:
    DEFAULT_PROJECT_PATH = BUNDLE_DIR

BUNDLED_MAPPING_PATH = os.path.join(BUNDLE_DIR, "store_mapping.json")
FALLBACK_MAPPING_PATH = os.path.join(
    BUNDLE_DIR, "RekonOnlineFood-Windows", "store_mapping.json"
)


def read_mapping_payload(path):
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        return None


def default_mapping_payload():
    """Return the bundled mapping payload or an empty default structure."""
    for mapping_path in (BUNDLED_MAPPING_PATH, FALLBACK_MAPPING_PATH):
        payload = read_mapping_payload(mapping_path)
        if payload is not None:
            return payload

    return {
        "_comment": "Mapping nama toko antara ERP dan Platform. Tambah toko baru di sini.",
        "_usage": {
            "folder": "Nama folder di Grabfood/ atau GoFood/ (UPPERCASE)",
            "erp1": "Nama di ERP File 1 (Penerimaan) - null jika tidak ada",
            "erp2": "Nama di ERP File 2 (Laporan Transaksi) - null jika tidak ada",
            "gof_merchant_id": "Merchant ID GoFood untuk identifikasi toko",
            "display": "Nama tampilan yang mudah dibaca",
            "grabfood_store": "Nama kolom 'Store Name' di file GrabFood - null jika tidak ada",
            "shopeefood_store": "Nama kolom 'Store Name' di file ShopeeFood - null jika tidak ada",
        },
        "stores": {},
    }


def mapping_payload(stores):
    """Build the persisted store mapping JSON payload."""
    payload = default_mapping_payload()
    payload["stores"] = stores
    return payload


def mapping_path_for(project_path):
    return os.path.join(project_path, "store_mapping.json")


def load_mapping_for_project(project_path):
    """Load project mapping, falling back to the bundled default mapping."""
    mapping_path = mapping_path_for(project_path)
    payload = read_mapping_payload(mapping_path)
    if payload is not None:
        return payload
    return default_mapping_payload()


def activate_project_mapping(project_path):
    """Refresh rekon.py globals so reconcile uses the selected project mapping."""
    stores = load_mapping_for_project(project_path).get("stores", {})

    rekon.STORE_MAP.clear()
    rekon.STORE_MAP.update(stores)

    rekon.ERP1_TO_FOLDER.clear()
    rekon.ERP2_TO_FOLDER.clear()
    rekon.GRABFOOD_STORE_TO_FOLDER.clear()
    rekon.GOFOOD_MERCHANT_TO_FOLDER.clear()
    rekon.SHOPEEFOOD_STORE_TO_FOLDER.clear()

    for folder, info in rekon.STORE_MAP.items():
        if info.get("erp1"):
            rekon.ERP1_TO_FOLDER[info["erp1"]] = folder
        if info.get("erp2"):
            rekon.ERP2_TO_FOLDER[info["erp2"]] = folder
        if info.get("grabfood_store"):
            rekon.GRABFOOD_STORE_TO_FOLDER[info["grabfood_store"]] = folder
        if info.get("gof_merchant_id"):
            rekon.GOFOOD_MERCHANT_TO_FOLDER[info["gof_merchant_id"]] = folder
        if info.get("shopeefood_store"):
            rekon.SHOPEEFOOD_STORE_TO_FOLDER[info["shopeefood_store"]] = folder


def calculate_reconciliation(project_path, start_date, end_date):
    """Load, reconcile, and generate report rows for a project/date range."""
    activate_project_mapping(project_path)

    penerimaan_files, transaksi_files = find_erp_files(project_path, start_date, end_date)
    erp_data = []
    for f in penerimaan_files:
        erp_data.extend(load_erp_penerimaan(f, start_date, end_date))
    for f in transaksi_files:
        erp_data.extend(load_erp_transaksi(f, start_date, end_date))

    grabfood_data = load_grabfood_reports(project_path, start_date, end_date)
    gofood_data = load_gofood_reports(project_path, start_date, end_date)
    shopeefood_data = load_shopeefood_reports(project_path, start_date, end_date)

    erp_gofood = [e for e in erp_data if e["platform"] == "GoFood"]
    erp_grabfood = [e for e in erp_data if e["platform"] == "GrabFood"]
    erp_shopee = [e for e in erp_data if e["platform"] == "ShopeeFood"]

    matched_gofood, un_erp_gofood, un_plat_gofood = reconcile(erp_gofood, gofood_data)
    matched_grabfood, un_erp_grabfood, un_plat_grabfood = reconcile(
        erp_grabfood, grabfood_data
    )
    matched_shopee, un_erp_shopee, un_plat_shopee = reconcile(
        erp_shopee, shopeefood_data, allow_fuzzy=False
    )

    matched = matched_gofood + matched_grabfood + matched_shopee
    unmatched_erp = un_erp_gofood + un_erp_grabfood + un_erp_shopee
    unmatched_platform = un_plat_gofood + un_plat_grabfood + un_plat_shopee

    summary_rows = generate_summary_per_store_day(
        matched, unmatched_erp, unmatched_platform, start_date, end_date
    )
    platform_summary_rows = generate_summary_per_platform(
        matched, unmatched_erp, unmatched_platform, start_date, end_date
    )
    detail_rows = generate_detail(matched, unmatched_erp, unmatched_platform)

    return {
        "matched": matched,
        "unmatched_erp": unmatched_erp,
        "unmatched_platform": unmatched_platform,
        "summary_rows": summary_rows,
        "platform_summary_rows": platform_summary_rows,
        "detail_rows": detail_rows,
    }


@app.route("/")
def index():
    """Main page."""
    return render_template(
        "index.html", default_path=DEFAULT_PROJECT_PATH, app_version=APP_VERSION
    )


@app.route("/api/health")
def health_check():
    """Health endpoint used by packaged launchers to detect an existing app."""
    return jsonify({"app": "RekonOnlineFood", "status": "ok"})


@app.route("/config")
def config_page():
    """Config page for store mapping."""
    return render_template("config.html", default_path=DEFAULT_PROJECT_PATH)


@app.route("/api/config/mapping", methods=["GET"])
def get_store_mapping():
    """Get store mapping from JSON file."""
    data = request.args
    project_path = data.get("project_path", DEFAULT_PROJECT_PATH)

    mapping = load_mapping_for_project(project_path)

    return jsonify({"stores": mapping.get("stores", {})})


@app.route("/api/config/mapping", methods=["POST"])
def save_store_mapping():
    """Save store mapping to JSON file."""
    data = request.get_json()
    project_path = data.get("project_path", DEFAULT_PROJECT_PATH)
    stores = data.get("stores", {})
    mapping_path = mapping_path_for(project_path)

    os.makedirs(project_path, exist_ok=True)
    with open(mapping_path, "w", encoding="utf-8") as f:
        mapping = mapping_payload(stores)
        json.dump(mapping, f, indent=4, ensure_ascii=False)

    activate_project_mapping(project_path)

    return jsonify({"success": True, "message": "Mapping tersimpan"})


@app.route("/api/browse", methods=["POST"])
def browse_folders():
    """Browse directories for folder picker."""
    data = request.get_json()
    current_path = data.get("path", "/")
    
    # Handle special cases
    if current_path == "~":
        current_path = os.path.expanduser("~")
    
    # Ensure path exists
    if not os.path.isdir(current_path):
        current_path = "/"
    
    # Get parent directory
    parent_path = os.path.dirname(current_path)
    if parent_path == current_path:
        parent_path = None
    
    # List directories
    directories = []
    try:
        for item in sorted(os.listdir(current_path)):
            item_path = os.path.join(current_path, item)
            if os.path.isdir(item_path) and not item.startswith("."):
                directories.append({
                    "name": item,
                    "path": item_path,
                })
    except PermissionError:
        return jsonify({"error": "Tidak ada akses ke folder ini"}), 403
    
    return jsonify({
        "current_path": current_path,
        "parent_path": parent_path,
        "directories": directories,
    })


@app.route("/api/select-folder", methods=["POST"])
def select_folder():
    """Open the native OS folder picker on the local machine."""
    data = request.get_json() or {}
    initial_dir = data.get("path") or DEFAULT_PROJECT_PATH
    if not os.path.isdir(initial_dir):
        initial_dir = DEFAULT_PROJECT_PATH

    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askdirectory(
            initialdir=initial_dir,
            title="Pilih Folder Project Rekon Online Food",
        )
        root.destroy()
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not selected:
        return jsonify({"cancelled": True})

    return jsonify({"path": selected})


@app.route("/api/scan", methods=["POST"])
def scan_folders():
    """Scan project folder and return available data info."""
    data = request.get_json(silent=True) or {}
    project_path = data.get("project_path", DEFAULT_PROJECT_PATH)
    start_str = data.get("start_date")
    end_str = data.get("end_date")

    if not start_str or not end_str:
        return jsonify({"error": "Pilih tanggal awal dan akhir dulu"}), 400

    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        return jsonify({"error": "Format tanggal salah. Gunakan YYYY-MM-DD"}), 400

    if start_date > end_date:
        return jsonify({"error": "Tanggal awal lebih besar dari tanggal akhir"}), 400

    if not os.path.isdir(project_path):
        return jsonify({"error": f"Folder tidak ditemukan: {project_path}"}), 400

    try:
        activate_project_mapping(project_path)

        penerimaan_files, transaksi_files = find_erp_files(project_path, start_date, end_date)
        penerimaan_rows = []
        transaksi_rows = []
        for f in penerimaan_files:
            penerimaan_rows.extend(load_erp_penerimaan(f, start_date, end_date))
        for f in transaksi_files:
            transaksi_rows.extend(load_erp_transaksi(f, start_date, end_date))

        grabfood_rows = load_grabfood_reports(project_path, start_date, end_date)
        gofood_rows = load_gofood_reports(project_path, start_date, end_date)
        shopeefood_rows = load_shopeefood_reports(project_path, start_date, end_date)
    except Exception as e:
        return jsonify({"error": f"Gagal scan folder: {e}"}), 500

    def files_from(rows):
        return sorted({r.get("source_file", "") for r in rows if r.get("source_file")})

    def platform_check(name, platform, rows, extensions=(".xlsx",), scan_all_folders=False):
        report_files = [report["filename"] for report in find_platform_reports(
            project_path, platform, start_date, end_date, extensions=extensions,
            scan_all_folders=scan_all_folders,
        )]
        return {
            "name": name,
            "required": False,
            "found": bool(rows),
            "files": files_from(rows) or sorted(set(report_files)),
            "rows": len(rows),
            "note": (
                "File ditemukan, tetapi tidak ada transaksi yang cocok dengan periode atau formatnya tidak terbaca."
                if report_files and not rows else ""
            ),
        }

    checks = [
        {
            "name": "ERP Penerimaan",
            "required": False,
            "required_label": "salah satu ERP",
            "found": bool(penerimaan_rows),
            "files": files_from(penerimaan_rows),
            "rows": len(penerimaan_rows),
        },
        {
            "name": "ERP Transaksi",
            "required": False,
            "required_label": "salah satu ERP",
            "found": bool(transaksi_rows),
            "files": files_from(transaksi_rows),
            "rows": len(transaksi_rows),
        },
        platform_check("GrabFood", "Grabfood", grabfood_rows, extensions=(".xlsx", ".csv")),
        platform_check("GoFood", "GoFood", gofood_rows, scan_all_folders=True),
        platform_check("ShopeeFood", "ShopeeFood", shopeefood_rows),
    ]

    result = {
        "project_path": project_path,
        "periode": f"{start_date} s/d {end_date}",
        "ready": bool(penerimaan_rows or transaksi_rows),
        "checks": checks,
        "erp": {
            "penerimaan": [
                {"name": os.path.basename(f), "path": f} for f in penerimaan_files
            ],
            "transaksi": [
                {"name": os.path.basename(f), "path": f} for f in transaksi_files
            ],
        },
        "stores": [
            {"folder": k, "display": v.get("display", k)}
            for k, v in sorted(STORE_MAP.items())
        ],
    }

    return jsonify(result)


@app.route("/api/rekon", methods=["POST"])
def run_reconciliation():
    """Run reconciliation and return results."""
    data = request.get_json()
    project_path = data.get("project_path", DEFAULT_PROJECT_PATH)
    start_str = data.get("start_date")
    end_str = data.get("end_date")

    if not start_str or not end_str:
        return jsonify({"error": "Tanggal awal dan akhir harus diisi"}), 400

    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        return jsonify({"error": "Format tanggal salah. Gunakan YYYY-MM-DD"}), 400

    if start_date > end_date:
        return jsonify({"error": "Tanggal awal lebih besar dari tanggal akhir"}), 400

    if not os.path.isdir(project_path):
        return jsonify({"error": f"Folder tidak ditemukan: {project_path}"}), 400

    results = calculate_reconciliation(project_path, start_date, end_date)
    summary_rows = results["summary_rows"]
    platform_summary_rows = results["platform_summary_rows"]
    detail_rows = results["detail_rows"]

    # Convert dates to strings for JSON
    for row in summary_rows:
        row["Tanggal"] = str(row["Tanggal"])
    for row in detail_rows:
        row["Tanggal"] = str(row["Tanggal"])

    # Stats
    total_erp = sum(r["Total ERP"] for r in summary_rows)
    total_platform = sum(r["Total Platform"] for r in summary_rows)
    n_matched = sum(1 for r in detail_rows if r["Status"] == "COCOK")
    n_erp_only = sum(1 for r in detail_rows if r["Status"] == "DI ERP SAJA")
    n_plat_only = sum(1 for r in detail_rows if r["Status"] == "DI PLATFORM SAJA")

    result = {
        "periode": f"{start_date} s/d {end_date}",
        "stats": {
            "total_erp": total_erp,
            "total_platform": total_platform,
            "selisih": total_erp - total_platform,
            "matched": n_matched,
            "erp_only": n_erp_only,
            "platform_only": n_plat_only,
            "total_transactions": len(detail_rows),
        },
        "summary": summary_rows,
        "platform_summary": platform_summary_rows,
        "detail": detail_rows,
    }

    return jsonify(result)


@app.route("/api/export", methods=["POST"])
def export_excel():
    """Export reconciliation results to Excel."""
    data = request.get_json()
    project_path = data.get("project_path", DEFAULT_PROJECT_PATH)
    start_str = data.get("start_date")
    end_str = data.get("end_date")

    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return jsonify({"error": "Invalid date"}), 400

    if start_date > end_date:
        return jsonify({"error": "Tanggal awal lebih besar dari tanggal akhir"}), 400

    if not os.path.isdir(project_path):
        return jsonify({"error": f"Folder tidak ditemukan: {project_path}"}), 400

    # Generate output path
    date_str = start_date.strftime("%Y%m%d")
    if start_date != end_date:
        date_str += f"-{end_date.strftime('%Y%m%d')}"
    output_dir = os.path.join(project_path, "Hasil Rekon")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"Rekon_OnlineFood_{date_str}.xlsx")

    results = calculate_reconciliation(project_path, start_date, end_date)

    export_to_excel(
        results["summary_rows"],
        results["detail_rows"],
        results["matched"],
        results["unmatched_erp"],
        results["unmatched_platform"],
        start_date,
        end_date,
        output_path,
    )

    return jsonify({
        "success": True,
        "path": output_path,
        "filename": os.path.basename(output_path),
    })


@app.route("/api/download/<path:filename>")
def download_file(filename):
    """Download a file."""
    return send_file(filename, as_attachment=True)


if __name__ == "__main__":
    import webbrowser
    import threading

    port = 8080
    app_url = f"http://localhost:{port}"
    health_url = f"http://127.0.0.1:{port}/api/health"

    def is_existing_app_running():
        try:
            with urllib.request.urlopen(health_url, timeout=1) as response:
                payload = json.loads(response.read().decode("utf-8"))
            return payload.get("app") == "RekonOnlineFood"
        except Exception:
            return False

    if is_existing_app_running():
        webbrowser.open(app_url)
        sys.exit(0)

    def open_browser():
        import time
        for _ in range(40):
            if is_existing_app_running():
                try:
                    webbrowser.open(app_url)
                except Exception:
                    pass
                return
            time.sleep(0.5)

    if os.environ.get("REKON_NO_BROWSER") != "1":
        threading.Thread(target=open_browser, daemon=True).start()

    print()
    print("=" * 50)
    print("  REKON ONLINE FOOD - Web Interface")
    print("=" * 50)
    print(f"  Buka browser: http://localhost:{port}")
    print(f"  Tekan Ctrl+C untuk berhenti")
    print("=" * 50)
    print()

    app.run(host="127.0.0.1", port=port, debug=False)
