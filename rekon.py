#!/usr/bin/env python3
"""
Rekon Online Food - Offline Reconciliation Tool
================================================
Reconciles ERP transactions with online food platform settlements
(GoFood, GrabFood, ShopeeFood).

Usage:
    python3 rekon.py --start-date 2026-05-19 --end-date 2026-05-19
    python3 rekon.py --start-date 2026-06-11 --end-date 2026-06-11
    python3 rekon.py  # interactive mode - will ask for dates
    
Merchant ID Mapping:
    python3 rekon.py --map-merchant           # Show merchant ID mapping status
    python3 rekon.py --auto-map-merchant      # Auto-map all merchant IDs
"""

import argparse
import os
import sys
import json
import re
from datetime import datetime, date
from io import BytesIO
from pathlib import Path
from collections import defaultdict
import warnings
from zipfile import ZipFile

warnings.filterwarnings("ignore")

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas belum terinstall. Jalankan: pip3 install pandas openpyxl")
    sys.exit(1)

# ============================================================
# CONFIGURATION - Load from store_mapping.json
# ============================================================

CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))
STORE_MAPPING_FILE = os.path.join(CONFIG_DIR, "store_mapping.json")
RAW_DATA_DIR_NAME = "Raw Data Transaksi"


def load_store_mapping():
    """Load store mapping from JSON config file."""
    if os.path.exists(STORE_MAPPING_FILE):
        with open(STORE_MAPPING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("stores", {})
    return {}


def resolve_data_root(project_path):
    """Return the folder that contains Download ERP, Grabfood, GoFood, ShopeeFood."""
    raw_data_dir = os.path.join(project_path, RAW_DATA_DIR_NAME)
    if os.path.isdir(raw_data_dir):
        return raw_data_dir
    return project_path


def date_scoped_roots(base_dir, start_date=None, end_date=None):
    """Return dated subfolders in range, falling back to base_dir for legacy layouts."""
    if not start_date or not end_date:
        return [base_dir]

    roots = []
    for name in sorted(os.listdir(base_dir)):
        full = os.path.join(base_dir, name)
        if not os.path.isdir(full):
            continue
        folder_date = parse_report_folder_date(name)
        if folder_date and start_date <= folder_date <= end_date:
            roots.append(full)

    return roots or [base_dir]


# Load on import
STORE_MAP = load_store_mapping()

# Build reverse maps for quick lookup
ERP1_TO_FOLDER = {}
ERP2_TO_FOLDER = {}
GRABFOOD_STORE_TO_FOLDER = {}
GOFOOD_MERCHANT_TO_FOLDER = {}
SHOPEEFOOD_STORE_TO_FOLDER = {}
for folder, info in STORE_MAP.items():
    if info.get("erp1"):
        ERP1_TO_FOLDER[info["erp1"]] = folder
    if info.get("erp2"):
        ERP2_TO_FOLDER[info["erp2"]] = folder
    if info.get("grabfood_store"):
        GRABFOOD_STORE_TO_FOLDER[info["grabfood_store"]] = folder
    if info.get("gof_merchant_id"):
        GOFOOD_MERCHANT_TO_FOLDER[info["gof_merchant_id"]] = folder
    if info.get("shopeefood_store"):
        SHOPEEFOOD_STORE_TO_FOLDER[info["shopeefood_store"]] = folder


# Track unmapped stores for reporting
UNMAPPED_STORES = []


def map_platform_store(name, source):
    """Map a platform identifier only through its platform-specific exact mapping."""
    lookup = {
        "grabfood": GRABFOOD_STORE_TO_FOLDER,
        "gofood": GOFOOD_MERCHANT_TO_FOLDER,
        "shopeefood": SHOPEEFOOD_STORE_TO_FOLDER,
    }.get(source.lower(), {})
    value = str(name).strip() if name is not None else ""
    if not value:
        return "UNMAPPED"
    for mapped_value, folder in lookup.items():
        if str(mapped_value).casefold() == value.casefold():
            return folder
    return "UNMAPPED"


def auto_detect_store(name, source="unknown"):
    """
    Auto-detect store from name. If not in mapping, try fuzzy match.
    Returns (folder_name, is_new).
    """
    if source.lower() in {"grabfood", "gofood", "shopeefood"}:
        return map_platform_store(name, source), False

    if not name or name == "nan":
        return "UNKNOWN", False

    name = str(name).strip()

    # Exact match in existing mappings
    for folder, info in STORE_MAP.items():
        # Check erp1
        if info.get("erp1") and info["erp1"].lower() == name.lower():
            return folder, False
        # Check erp2
        if info.get("erp2") and info["erp2"].lower() == name.lower():
            return folder, False
        # Check grabfood_store
        if info.get("grabfood_store") and info["grabfood_store"].lower() == name.lower():
            return folder, False
        # Check shopeefood_store
        if info.get("shopeefood_store") and info["shopeefood_store"].lower() == name.lower():
            return folder, False
        # Check gof_merchant_id
        if info.get("gof_merchant_id") and info["gof_merchant_id"].lower() == name.lower():
            return folder, False
        # Check folder name itself
        if folder.lower() == name.lower():
            return folder, False
        # Check display name
        if info.get("display") and info["display"].lower() == name.lower():
            return folder, False

    # Fuzzy match: check if name contains any known folder name
    name_upper = name.upper()
    for folder in STORE_MAP:
        if folder in name_upper or name_upper in folder:
            return folder, False

    # Try to extract store name from "Procil Kios XXX" or "Procil Bubur Tim Organik - Kios XXX"
    import re
    match = re.search(r'Kios\s+(.+)', name, re.IGNORECASE)
    if match:
        extracted = match.group(1).strip().upper()
        # Check if extracted name matches any folder
        for folder in STORE_MAP:
            if folder.upper() == extracted or folder.upper() in extracted:
                return folder, False

    # NEW STORE - create mapping automatically
    import re

    # Try to extract clean store name from various formats
    new_folder = None

    # Pattern 1: "Procil Kios XXX" or "Procil Bubur Tim Organik - Kios XXX"
    match = re.search(r'Kios\s+(.+)', name, re.IGNORECASE)
    if match:
        extracted = match.group(1).strip()
        new_folder = extracted.upper().replace(" ", "_").replace("-", "_")

    # Pattern 2: Just a simple name like "Depok Baru"
    if not new_folder:
        # Remove common prefixes
        clean = name
        for prefix in ["Procil Kios ", "Procil Bubur Tim Organik - Kios ", "Procil "]:
            if clean.lower().startswith(prefix.lower()):
                clean = clean[len(prefix):]
                break
        new_folder = clean.upper().replace(" ", "_").replace("-", "_")

    # Clean up folder name
    new_folder = re.sub(r'[^A-Z0-9_]', '', new_folder)  # Remove special chars
    new_folder = re.sub(r'_+', '_', new_folder)  # Remove double underscores
    new_folder = new_folder.strip('_')  # Remove leading/trailing underscores

    if new_folder and new_folder not in STORE_MAP:
        # Auto-add to mapping
        new_mapping = {
            "erp1": None,
            "erp2": None,
            "display": name.title(),
            "grabfood_store": None,
            "shopeefood_store": None,
        }

        # Set the right field based on source
        if "erp" in source.lower() and "penerimaan" in source.lower():
            new_mapping["erp1"] = name
        elif "erp" in source.lower():
            new_mapping["erp2"] = name
        elif "grabfood" in source.lower():
            new_mapping["grabfood_store"] = name
        elif "shopeefood" in source.lower():
            new_mapping["shopeefood_store"] = name

        STORE_MAP[new_folder] = new_mapping

        # Update reverse maps
        if new_mapping["erp1"]:
            ERP1_TO_FOLDER[new_mapping["erp1"]] = new_folder
        if new_mapping["erp2"]:
            ERP2_TO_FOLDER[new_mapping["erp2"]] = new_folder
        if new_mapping["grabfood_store"]:
            GRABFOOD_STORE_TO_FOLDER[new_mapping["grabfood_store"]] = new_folder
        if new_mapping["shopeefood_store"]:
            SHOPEEFOOD_STORE_TO_FOLDER[new_mapping["shopeefood_store"]] = new_folder

        # Track for reporting
        UNMAPPED_STORES.append({
            "name": name,
            "folder": new_folder,
            "source": source,
        })

        return new_folder, True

    return new_folder, False


def save_store_mapping():
    """Save current STORE_MAP back to JSON file."""
    data = {
        "_comment": "Mapping nama toko antara ERP dan Platform. Tambah toko baru di sini.",
        "_usage": {
            "folder": "Nama folder di Grabfood/ atau GoFood/ (UPPERCASE)",
            "erp1": "Nama di ERP File 1 (Penerimaan) - null jika tidak ada",
            "erp2": "Nama di ERP File 2 (Laporan Transaksi) - null jika tidak ada",
            "gof_merchant_id": "Merchant ID GoFood untuk identifikasi toko",
            "display": "Nama tampilan yang mudah dibaca",
            "grabfood_store": "Nama kolom 'Store Name' di file GrabFood - null jika tidak ada",
            "shopeefood_store": "Nama kolom 'Store Name' di file ShopeeFood - null jika tidak ada"
        },
        "stores": STORE_MAP
    }
    with open(STORE_MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def scan_gofood_merchant_ids(project_path):
    """
    Scan all GoFood files to extract Merchant IDs.
    Returns dict: {merchant_id: [list of folder names]}
    """
    gofood_base = os.path.join(resolve_data_root(project_path), "GoFood")
    if not os.path.isdir(gofood_base):
        return {}
    
    merchant_folders = defaultdict(list)
    
    for root, _, files in os.walk(gofood_base):
        for excel_file in files:
            if not excel_file.endswith('.xlsx') or excel_file.startswith('~$'):
                continue
            
            filepath = os.path.join(root, excel_file)
            try:
                df = pd.read_excel(filepath, header=0)
                if 'Merchant ID' in df.columns and len(df) > 0:
                    merchant_id = str(df['Merchant ID'].iloc[0]).strip()
                    # Extract store name from filename
                    store_name = excel_file.replace('Laporan Gofood ', '').replace('.xlsx', '').replace('.xls', '')
                    # Remove date pattern if present
                    import re
                    store_name = re.sub(r'\d{2}-\d{2}-\d{4}', '', store_name).strip()
                    
                    if merchant_id and store_name:
                        merchant_folders[merchant_id].append(store_name)
            except Exception:
                continue
    
    return dict(merchant_folders)


def show_merchant_id_mapping(project_path):
    """
    Show current Merchant ID mapping status.
    Returns (mapped, unmapped) tuples.
    """
    # Scan GoFood files for Merchant IDs
    scanned = scan_gofood_merchant_ids(project_path)
    
    mapped = []
    unmapped = []
    
    for merchant_id, folders in sorted(scanned.items()):
        # Get the most common folder name (usually the store name)
        from collections import Counter
        folder_counts = Counter(folders)
        most_common_folder = folder_counts.most_common(1)[0][0]
        
        # Check if mapped
        is_mapped = merchant_id in GOFOOD_MERCHANT_TO_FOLDER
        mapped_store = GOFOOD_MERCHANT_TO_FOLDER.get(merchant_id, "")
        
        entry = {
            "merchant_id": merchant_id,
            "folder_name": most_common_folder,
            "is_mapped": is_mapped,
            "mapped_store": mapped_store
        }
        
        if is_mapped:
            mapped.append(entry)
        else:
            unmapped.append(entry)
    
    return mapped, unmapped


def map_merchant_id_to_store(merchant_id, folder_name):
    """
    Map a Merchant ID to a store folder.
    Updates both STORE_MAP and GOFOOD_MERCHANT_TO_FOLDER.
    """
    # Ensure the store exists in STORE_MAP
    if folder_name not in STORE_MAP:
        STORE_MAP[folder_name] = {}
    
    # Add merchant ID mapping
    STORE_MAP[folder_name]["gof_merchant_id"] = merchant_id
    
    # Update reverse map
    GOFOOD_MERCHANT_TO_FOLDER[merchant_id] = folder_name
    
    # Save to file
    save_store_mapping()
    
    return True

# Indonesian month mapping
ID_MONTHS = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


# ============================================================
# DATA LOADERS
# ============================================================

def find_erp_files(project_path, start_date=None, end_date=None):
    """Find all ERP files in Download ERP folder."""
    erp_dir = os.path.join(resolve_data_root(project_path), "Download ERP")
    if not os.path.isdir(erp_dir):
        print(f"  [WARN] Folder tidak ditemukan: {erp_dir}")
        return [], []

    penerimaan_files = []
    transaksi_files = []

    for scan_root in date_scoped_roots(erp_dir, start_date, end_date):
        for root, _, files in os.walk(scan_root):
            for f in files:
                if not f.endswith(".xlsx") or f.startswith("~$"):
                    continue
                full = os.path.join(root, f)
                if "penerimaan" in f.lower():
                    penerimaan_files.append(full)
                elif "laporan" in f.lower() or "transaksi" in f.lower():
                    transaksi_files.append(full)

    return penerimaan_files, transaksi_files


def load_erp_penerimaan(filepath, start_date, end_date):
    """
    Load ERP File 1 (Penerimaan Penjualan Per Tipe Pembayaran).
    Returns list of dicts with standardized fields.
    """
    try:
        df = pd.read_excel(filepath, header=4)
    except Exception as e:
        print(f"  [ERROR] Gagal baca {filepath}: {e}")
        return []

    # Rename columns (some are "Unnamed")
    cols = list(df.columns)
    clean_cols = []
    for i, c in enumerate(cols):
        if str(c).startswith("Unnamed"):
            clean_cols.append(f"_col{i}")
        else:
            clean_cols.append(str(c))
    df.columns = clean_cols

    # Filter only online food
    online_platforms = ["GoFood", "GrabFood", "Shopee Food"]
    df = df[df["Tipe Pembayaran"].isin(online_platforms)].copy()

    if df.empty:
        return []

    # Parse dates
    df["Waktu Transaksi (POS)"] = pd.to_datetime(df["Waktu Transaksi (POS)"], errors="coerce")
    df = df.dropna(subset=["Waktu Transaksi (POS)"])

    # Filter by date range
    df["date"] = df["Waktu Transaksi (POS)"].dt.date
    df = df[(df["date"] >= start_date) & (df["date"] <= end_date)]

    results = []
    for _, row in df.iterrows():
        cabang = str(row.get("Nama Cabang Faktur Penjualan", "")).strip()
        folder, is_new = auto_detect_store(cabang, "erp_penerimaan")

        # Normalize platform name
        platform = row["Tipe Pembayaran"]
        if platform == "Shopee Food":
            platform = "ShopeeFood"

        results.append({
            "source": "ERP",
            "source_file": os.path.basename(filepath),
            "platform": platform,
            "store_folder": folder,
            "store_erp": cabang,
            "no_faktur": str(row.get("Nomor # Faktur Penjualan", "")),
            "amount": float(row.get("Total Penerimaan", 0) or 0),
            "datetime": row["Waktu Transaksi (POS)"],
            "date": row["date"],
        })

    return results


def load_erp_transaksi(filepath, start_date, end_date):
    """
    Load ERP File 2 (Laporan Transaksi Penjualan).
    Returns list of dicts with standardized fields.
    """
    try:
        df = pd.read_excel(filepath, header=1)
    except Exception as e:
        print(f"  [ERROR] Gagal baca {filepath}: {e}")
        return []

    # Filter only online food
    online_channels = ["GoFood", "GrabFood", "ShopeeFood"]
    df = df[df["Sales Channel"].isin(online_channels)].copy()

    if df.empty:
        return []

    # Parse dates
    df["tanggal_dt"] = pd.to_datetime(df["tanggal dan waktu pembayaran"], errors="coerce")
    df = df.dropna(subset=["tanggal_dt"])

    # Filter by date range
    df["date"] = df["tanggal_dt"].dt.date
    df = df[(df["date"] >= start_date) & (df["date"] <= end_date)]

    results = []
    for _, row in df.iterrows():
        cabang = str(row.get("Cabang", "")).strip()
        folder, is_new = auto_detect_store(cabang, "erp_transaksi")

        results.append({
            "source": "ERP",
            "source_file": os.path.basename(filepath),
            "platform": str(row["Sales Channel"]),
            "store_folder": folder,
            "store_erp": cabang,
            "no_faktur": str(row.get("No Penggunaan", "")),
            "amount": float(row.get("Total Harga Jual (Net)", 0) or 0),
            "datetime": row["tanggal_dt"],
            "date": row["date"],
        })

    return results


def find_platform_reports(
    project_path, platform, start_date, end_date, extensions=(".xlsx",), scan_all_folders=False
):
    """
    Find platform report files. Transaction dates are filtered by each loader.
    Reports are in: Raw Data Transaksi/<platform>/<YYYY-MM-DD>/<filename>.xlsx.
    Legacy project-root folders still work.
    """
    base_dir = os.path.join(resolve_data_root(project_path), platform)
    if not os.path.isdir(base_dir):
        return []

    extensions = tuple(extension.lower() for extension in extensions)

    def is_report_file(filename):
        return filename.lower().endswith(extensions) and not filename.startswith("~$")

    found_files = []
    scan_roots = [base_dir] if scan_all_folders else date_scoped_roots(
        base_dir, start_date, end_date
    )

    # A consolidated multi-outlet file may be saved directly in the platform folder.
    if not scan_all_folders and scan_roots != [base_dir]:
        for f in os.listdir(base_dir):
            if is_report_file(f):
                found_files.append({
                    "path": os.path.join(base_dir, f),
                    "filename": f,
                    "folder_date": None,
                    "folder_name": os.path.basename(base_dir),
                })

    for scan_root in scan_roots:
        for root, _, files in os.walk(scan_root):
            folder_name = os.path.basename(root)
            folder_date = parse_report_folder_date(folder_name)

            for f in files:
                if is_report_file(f):
                    found_files.append({
                        "path": os.path.join(root, f),
                        "filename": f,
                        "folder_date": folder_date,
                        "folder_name": folder_name,
                    })

    return found_files


def parse_report_folder_date(date_str):
    """Parse supported report folder dates."""
    try:
        return datetime.strptime(date_str.strip(), "%Y-%m-%d").date()
    except ValueError:
        return parse_indonesian_date(date_str)


def parse_indonesian_date(date_str):
    """Parse '19 Mei 2026' -> date(2026, 5, 19)."""
    try:
        parts = date_str.strip().split()
        day = int(parts[0])
        month_name = parts[1].lower()
        year = int(parts[2])

        month_map = {v.lower(): k for k, v in ID_MONTHS.items()}
        month = month_map.get(month_name)
        if month is None:
            return None

        return date(year, month, day)
    except (ValueError, IndexError):
        return None


def row_value(row, column, default=""):
    """Return a stripped cell value without leaking pandas NaN strings."""
    value = row.get(column, default)
    if pd.isna(value):
        return default
    return str(value).strip()


def normalize_gofood_columns(df):
    """Normalize known GoFood headers exported with different casing or spacing."""
    aliases = {
        "merchant id": "Merchant ID",
        "merchant name": "Nama Merchant",
        "nama merchant": "Nama Merchant",
        "outlet name": "Nama Outlet",
        "nama outlet": "Nama Outlet",
        "nama toko": "Nama Outlet",
        "waktu transaksi": "Waktu transaksi",
        "nomor pesanan": "Nomor pesanan",
        "penjualan": "Penjualan",
        "biaya gofood": "Biaya GoFood",
        "total biaya": "Total Biaya",
        "pendapatan bersih": "Pendapatan Bersih",
    }
    rename = {}
    for column in df.columns:
        normalized = " ".join(str(column).replace("\ufeff", "").split()).casefold()
        if normalized in aliases:
            rename[column] = aliases[normalized]
    return df.rename(columns=rename)


REVERSED_XLSX_CELL_REFERENCE = re.compile(
    rb'(<c\b[^>]*\br=")(\d+)([A-Z]+)(")'
)


def repair_reversed_xlsx_references(filepath):
    """Return a repaired in-memory workbook for malformed cell refs such as 1A."""
    repaired = BytesIO()
    repaired_count = 0
    with ZipFile(filepath) as source, ZipFile(repaired, "w") as target:
        for entry in source.infolist():
            content = source.read(entry.filename)
            if entry.filename.startswith("xl/worksheets/") and entry.filename.endswith(".xml"):
                content, count = REVERSED_XLSX_CELL_REFERENCE.subn(
                    lambda match: match.group(1) + match.group(3) + match.group(2) + match.group(4),
                    content,
                )
                repaired_count += count
            target.writestr(entry, content)
    if not repaired_count:
        raise ValueError("Tidak menemukan referensi sel XLSX yang dapat diperbaiki")
    repaired.seek(0)
    return repaired


def read_gofood_report_source(source):
    """Read the first matching GoFood header from an Excel source."""
    last_df = None
    for header in range(5):
        if hasattr(source, "seek"):
            source.seek(0)
        df = normalize_gofood_columns(pd.read_excel(source, header=header))
        if "Waktu transaksi" in df.columns:
            return df
        last_df = df
    return last_df


def read_gofood_report(filepath):
    """Read GoFood files whose headers may start after a title row."""
    try:
        return read_gofood_report_source(filepath)
    except ValueError as error:
        if "invalid literal for int()" not in str(error):
            raise
        return read_gofood_report_source(repair_reversed_xlsx_references(filepath))


def parse_gofood_datetime(value):
    """Parse ISO and Indonesian-style GoFood transaction timestamps safely."""
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    iso_format = len(text) >= 5 and text[:4].isdigit() and text[4] in "-/"
    return pd.to_datetime(value, errors="coerce", dayfirst=not iso_format)


def read_grabfood_report(filepath):
    """Read a standard GrabFood Excel or CSV transaction export."""
    if filepath.lower().endswith(".csv"):
        try:
            return pd.read_csv(filepath, encoding="utf-8-sig", sep=None, engine="python")
        except UnicodeDecodeError:
            return pd.read_csv(filepath, encoding="latin-1", sep=None, engine="python")
    return pd.read_excel(filepath, header=0)


def load_grabfood_reports(project_path, start_date, end_date):
    """
    Load all Grabfood reports within date range.
    Returns list of dicts with standardized fields.
    """
    reports = find_platform_reports(
        project_path, "Grabfood", start_date, end_date, extensions=(".xlsx", ".csv")
    )
    if not reports:
        return []

    results = []
    for rpt in reports:
        try:
            df = read_grabfood_report(rpt["path"])
        except Exception as e:
            print(f"  [WARN] Gagal baca {rpt['filename']}: {e}")
            continue

        if df.empty:
            continue

        default_store_name = (
            row_value(df.iloc[0], "Store Name") if "Store Name" in df.columns else ""
        )

        # Parse dates
        if "Transfer Date" in df.columns:
            df["transfer_date"] = pd.to_datetime(df["Transfer Date"], errors="coerce")
        if "Created On" in df.columns:
            df["created_on"] = pd.to_datetime(df["Created On"], errors="coerce")

        # Use Created On as transaction date for matching
        df["txn_date"] = df.get("created_on", df.get("transfer_date"))

        for _, row in df.iterrows():
            txn_date = row.get("txn_date")
            if pd.isna(txn_date):
                continue

            txn_date_val = txn_date.date() if hasattr(txn_date, "date") else None
            if txn_date_val is None:
                continue

            # Filter by date range
            if not (start_date <= txn_date_val <= end_date):
                continue

            store_name = row_value(row, "Store Name") or default_store_name
            folder = map_platform_store(store_name, "grabfood")
            data_issue = (
                f"Outlet GrabFood belum dimapping exact: {store_name or '-'}"
                if folder == "UNMAPPED" else None
            )

            results.append({
                "source": "Grabfood",
                "source_file": rpt["filename"],
                "platform": "GrabFood",
                "store_folder": folder,
                "store_platform": store_name,
                "data_issue": data_issue,
                "transaction_id": str(row.get("Transaction ID", "")),
                "short_order_id": str(row.get("Short Order ID", "")),
                "amount": float(row.get("Net Sales", 0) or 0),
                "gross_amount": float(row.get("Amount", 0) or 0),
                "net_sales": float(row.get("Net Sales", 0) or 0),
                "grab_fee": float(row.get("Grab Fee", 0) or 0),
                "total_received": float(row.get("Total", 0) or 0),
                "status": str(row.get("Status", "")),
                "payment_method": str(row.get("Payment Method", "")),
                "datetime": txn_date,
                "date": txn_date_val,
            })

    return results


def load_gofood_reports(project_path, start_date, end_date):
    """
    Load all GoFood reports within date range.
    Returns list of dicts with standardized fields.
    """
    reports = find_platform_reports(
        project_path, "GoFood", start_date, end_date, scan_all_folders=True
    )
    if not reports:
        return []

    results = []
    for rpt in reports:
        try:
            df = read_gofood_report(rpt["path"])
        except Exception as e:
            print(f"  [WARN] Gagal baca {rpt['filename']}: {e}")
            continue

        if df.empty:
            continue

        if "Waktu transaksi" not in df.columns:
            print(f"  [WARN] {rpt['filename']}: kolom 'Waktu transaksi' tidak ditemukan")
            continue

        # Consolidated exports can leave Merchant ID blank after the first row
        # of each outlet. Carry it only within its own preceding outlet block.
        if "Merchant ID" in df.columns:
            df["Merchant ID"] = df["Merchant ID"].replace(r"^\s*$", pd.NA, regex=True).ffill()

        df["txn_dt"] = df["Waktu transaksi"].apply(parse_gofood_datetime)

        for _, row in df.iterrows():
            txn_dt = row.get("txn_dt")
            if pd.isna(txn_dt):
                continue

            txn_date_val = txn_dt.date() if hasattr(txn_dt, "date") else None
            if txn_date_val is None:
                continue

            if not (start_date <= txn_date_val <= end_date):
                continue

            merchant_id = row_value(row, "Merchant ID")
            folder_store = map_platform_store(merchant_id, "gofood")
            outlet_name = next(
                (row_value(row, column) for column in ("Nama Outlet", "Nama Merchant")
                 if column in df.columns and row_value(row, column)),
                "",
            )
            data_issue = (
                f"Merchant ID GoFood belum dimapping exact: {merchant_id or outlet_name or '-'}"
                if folder_store == "UNMAPPED" else None
            )

            results.append({
                "source": "GoFood",
                "source_file": rpt["filename"],
                "platform": "GoFood",
                "store_folder": folder_store,
                "store_platform": outlet_name or folder_store,
                "data_issue": data_issue,
                "order_id": str(row.get("Nomor pesanan", "")),
                "merchant_id": merchant_id,
                "penjualan": float(row.get("Penjualan", 0) or 0),
                "biaya_gofeed": float(row.get("Biaya GoFood", 0) or 0),
                "total_biaya": float(row.get("Total Biaya", 0) or 0),
                "pendapatan_bersih": float(row.get("Pendapatan Bersih", 0) or 0),
                "amount": float(row.get("Penjualan", 0) or 0),  # unified field
                "datetime": txn_dt,
                "date": txn_date_val,
            })

    return results


def parse_shopee_amount(value):
    """Parse ShopeeFood amount strings like '10.000' into 10000."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace(".", "").replace(",", ".")
    return float(text)


def load_shopeefood_reports(project_path, start_date, end_date):
    """
    Load all ShopeeFood transaction reports within date range.
    Returns list of dicts with standardized fields.
    """
    reports = find_platform_reports(project_path, "ShopeeFood", start_date, end_date)
    if not reports:
        return []

    results = []
    for rpt in reports:
        if "transactions" not in rpt["filename"].lower():
            continue

        try:
            df = pd.read_excel(
                rpt["path"],
                sheet_name="Order_Payment_Details",
                dtype=str,
            )
            overall = pd.read_excel(
                rpt["path"],
                sheet_name="Overall",
                dtype=str,
            )
        except Exception as e:
            print(f"  [WARN] Gagal baca {rpt['filename']}: {e}")
            continue

        if df.empty or "Order Complete Time" not in df.columns:
            continue

        default_store_name = (
            row_value(df.iloc[0], "Store Name") if "Store Name" in df.columns else ""
        )
        df["txn_dt"] = pd.to_datetime(
            df["Order Complete Time"],
            format="%m/%d/%Y %H:%M:%S",
            errors="coerce",
        )
        overall_amounts = {}
        if {"ID", "Amount"}.issubset(overall.columns):
            overall_amounts = {
                str(row.get("ID", "")).strip(): parse_shopee_amount(row.get("Amount", 0))
                for _, row in overall.iterrows()
            }

        for _, row in df.iterrows():
            txn_dt = row.get("txn_dt")
            if pd.isna(txn_dt):
                continue

            txn_date_val = txn_dt.date() if hasattr(txn_dt, "date") else None
            if txn_date_val is None or not (start_date <= txn_date_val <= end_date):
                continue

            amount = parse_shopee_amount(row.get("Order Amount", 0))
            order_id = str(row.get("Order ID", ""))
            net_income = parse_shopee_amount(row.get("Net Income", 0))
            overall_amount = overall_amounts.get(order_id)
            data_issue = None
            if overall_amount is not None and overall_amount != net_income:
                data_issue = (
                    "ShopeeFood Transactions tidak konsisten: "
                    f"Overall Amount {overall_amount:,.0f} != Net Income {net_income:,.0f}"
                )
            report_amount = overall_amount if data_issue and overall_amount is not None else amount
            store_name = row_value(row, "Store Name") or default_store_name
            folder = map_platform_store(store_name, "shopeefood")
            if folder == "UNMAPPED":
                data_issue = f"Outlet ShopeeFood belum dimapping exact: {store_name or '-'}"

            results.append({
                "source": "ShopeeFood",
                "source_file": rpt["filename"],
                "platform": "ShopeeFood",
                "store_folder": folder,
                "store_platform": store_name,
                "order_id": order_id,
                "pickup_id": str(row.get("Order Pickup ID", "")),
                "settlement_id": str(row.get("Settlement ID", "")),
                "amount": amount,
                "report_amount": report_amount,
                "gross_amount": amount,
                "commission": parse_shopee_amount(row.get("Commission", 0)),
                "net_income": net_income,
                "overall_amount": overall_amount,
                "data_issue": data_issue,
                "status": str(row.get("Order Status", "")),
                "datetime": txn_dt,
                "date": txn_date_val,
            })

    return results


def extract_store_from_gofilename(filename):
    """Extract store folder name from GoFood filename."""
    # Patterns:
    #   "Laporan Gofood BINTARA.xlsx"
    #   "Laporan Gofood 19-05-2026 MUCHTAR TABRANI.xlsx"
    #   "Laporan Gofood Seroja.xlsx" (note: mixed case)

    name = filename.replace(".xlsx", "").replace(".xls", "")

    # Remove "Laporan Gofood" prefix
    name = name.replace("Laporan Gofood", "").replace("Laporan GoFood", "").strip()

    # Remove date pattern if present (e.g., "19-05-2026")
    import re
    name = re.sub(r"\d{2}-\d{2}-\d{4}", "", name).strip()

    # Convert to uppercase for matching
    name_upper = name.upper()

    # Map to canonical folder name
    for folder in STORE_MAP:
        if folder.upper() == name_upper:
            return folder

    # Try partial match
    for folder in STORE_MAP:
        if folder.upper() in name_upper or name_upper in folder.upper():
            return folder

    return name_upper


# ============================================================
# RECONCILIATION ENGINE
# ============================================================

def reconcile(erp_data, platform_data, allow_fuzzy=False):
    """
    Match ERP transactions with platform transactions.
    Matching criteria: same store + same date + same amount.

    Returns:
        matched: list of (erp_row, platform_row) tuples
        unmatched_erp: list of erp_rows not matched
        unmatched_platform: list of platform_rows not matched
    """
    # Group by (store, date, amount) for matching
    # Build index from platform data
    platform_index = defaultdict(list)
    for p in platform_data:
        if p.get("data_issue"):
            continue
        key = (p["store_folder"], p["date"], p["amount"])
        platform_index[key].append(p)

    # Also build a fuzzy index (store + date, for amount tolerance matching)
    platform_by_store_date = defaultdict(list)
    for p in platform_data:
        if p.get("data_issue"):
            continue
        key = (p["store_folder"], p["date"])
        platform_by_store_date[key].append(p)

    matched = []
    unmatched_erp = []
    used_platform = set()

    for erp in erp_data:
        # Try exact match first
        key = (erp["store_folder"], erp["date"], erp["amount"])
        candidates = platform_index.get(key, [])

        found = False
        for cand in candidates:
            cand_id = id(cand)
            if cand_id not in used_platform:
                matched.append((erp, cand))
                used_platform.add(cand_id)
                found = True
                break

        if found:
            continue

        if not allow_fuzzy:
            unmatched_erp.append(erp)
            continue

        # Try fuzzy match: same store + same date, amount within tolerance
        key2 = (erp["store_folder"], erp["date"])
        candidates2 = platform_by_store_date.get(key2, [])

        for cand in candidates2:
            cand_id = id(cand)
            if cand_id in used_platform:
                continue
            # Allow 5% tolerance or Rp 5,000 tolerance
            diff = abs(erp["amount"] - cand["amount"])
            tolerance = max(erp["amount"] * 0.05, 5000)
            if diff <= tolerance:
                matched.append((erp, cand))
                used_platform.add(cand_id)
                found = True
                break

        if not found:
            unmatched_erp.append(erp)

    unmatched_platform = [p for p in platform_data if id(p) not in used_platform]

    return matched, unmatched_erp, unmatched_platform


# ============================================================
# REPORT GENERATORS
# ============================================================

def platform_report_amount(row):
    """Amount to show in reports; invalid Shopee rows show the changed source value."""
    return row.get("report_amount", row["amount"])


def generate_summary_per_store_day(matched, unmatched_erp, unmatched_platform, start_date, end_date):
    """Generate summary reconciliation per store per day."""
    rows = []

    # Collect all stores and dates
    all_stores = set()
    all_dates = set()

    # From matched
    for erp, plat in matched:
        all_stores.add(erp["store_folder"])
        all_dates.add(erp["date"])

    # From unmatched ERP
    for erp in unmatched_erp:
        all_stores.add(erp["store_folder"])
        all_dates.add(erp["date"])

    # From unmatched platform
    for plat in unmatched_platform:
        all_stores.add(plat["store_folder"])
        all_dates.add(plat["date"])

    all_dates = sorted(all_dates)
    all_stores = sorted(all_stores)

    for store in all_stores:
        for d in all_dates:
            # ERP totals
            erp_gofood = sum(e["amount"] for e, p in matched
                           if e["store_folder"] == store and e["date"] == d and e["platform"] == "GoFood")
            erp_grabfood = sum(e["amount"] for e, p in matched
                             if e["store_folder"] == store and e["date"] == d and e["platform"] == "GrabFood")
            erp_shopee = sum(e["amount"] for e, p in matched
                           if e["store_folder"] == store and e["date"] == d and e["platform"] == "ShopeeFood")
            erp_gofood_un = sum(e["amount"] for e in unmatched_erp
                              if e["store_folder"] == store and e["date"] == d and e["platform"] == "GoFood")
            erp_grabfood_un = sum(e["amount"] for e in unmatched_erp
                                if e["store_folder"] == store and e["date"] == d and e["platform"] == "GrabFood")
            erp_shopee_un = sum(e["amount"] for e in unmatched_erp
                              if e["store_folder"] == store and e["date"] == d and e["platform"] == "ShopeeFood")

            # Platform totals
            plat_gofood = sum(p["amount"] for e, p in matched
                            if e["store_folder"] == store and e["date"] == d and p["platform"] == "GoFood")
            plat_grabfood = sum(p["amount"] for e, p in matched
                              if e["store_folder"] == store and e["date"] == d and p["platform"] == "GrabFood")
            plat_shopee = sum(p["amount"] for e, p in matched
                            if e["store_folder"] == store and e["date"] == d and p["platform"] == "ShopeeFood")
            plat_gofood_un = sum(p["amount"] for p in unmatched_platform
                               if p["store_folder"] == store and p["date"] == d and p["platform"] == "GoFood")
            plat_grabfood_un = sum(p["amount"] for p in unmatched_platform
                                 if p["store_folder"] == store and p["date"] == d and p["platform"] == "GrabFood")
            plat_shopee_un = sum(platform_report_amount(p) for p in unmatched_platform
                               if p["store_folder"] == store and p["date"] == d and p["platform"] == "ShopeeFood")

            # Skip empty rows
            total_erp = erp_gofood + erp_grabfood + erp_shopee + erp_gofood_un + erp_grabfood_un + erp_shopee_un
            total_plat = plat_gofood + plat_grabfood + plat_shopee + plat_gofood_un + plat_grabfood_un + plat_shopee_un
            if total_erp == 0 and total_plat == 0:
                continue

            display = STORE_MAP.get(store, {}).get("display", store)

            rows.append({
                "Toko": display,
                "Tanggal": d,
                "ERP GoFood": erp_gofood + erp_gofood_un,
                "ERP GrabFood": erp_grabfood + erp_grabfood_un,
                "ERP ShopeeFood": erp_shopee + erp_shopee_un,
                "Platform GoFood": plat_gofood + plat_gofood_un,
                "Platform GrabFood": plat_grabfood + plat_grabfood_un,
                "Platform ShopeeFood": plat_shopee + plat_shopee_un,
                "Selisih GoFood": (erp_gofood + erp_gofood_un) - (plat_gofood + plat_gofood_un),
                "Selisih GrabFood": (erp_grabfood + erp_grabfood_un) - (plat_grabfood + plat_grabfood_un),
                "Selisih ShopeeFood": (erp_shopee + erp_shopee_un) - (plat_shopee + plat_shopee_un),
                "Total ERP": total_erp,
                "Total Platform": total_plat,
                "Selisih Total": total_erp - total_plat,
            })

    return rows


def generate_summary_per_platform(matched, unmatched_erp, unmatched_platform, start_date, end_date):
    """
    Generate summary: per Platform per Toko.
    Shows: Cocok, Hanya ERP, Hanya Platform, Total.
    """
    platforms = ["GoFood", "GrabFood", "ShopeeFood"]
    all_stores = set()
    for e, p in matched:
        all_stores.add(e["store_folder"])
    for e in unmatched_erp:
        all_stores.add(e["store_folder"])
    for p in unmatched_platform:
        all_stores.add(p["store_folder"])
    all_stores = sorted(all_stores)

    rows = []
    for platform in platforms:
        platform_erp_total = 0
        platform_plat_total = 0
        platform_matched_count = 0
        platform_erp_only_count = 0
        platform_plat_only_count = 0

        for store in all_stores:
            display = STORE_MAP.get(store, {}).get("display", store)

            # Matched: ERP amount + Platform amount
            erp_matched = sum(e["amount"] for e, p in matched
                             if e["store_folder"] == store and e["platform"] == platform)
            plat_matched = sum(p["amount"] for e, p in matched
                              if e["store_folder"] == store and p["platform"] == platform)
            cnt_matched = len([1 for e, p in matched
                              if e["store_folder"] == store and e["platform"] == platform])

            # Unmatched ERP only
            erp_only = sum(e["amount"] for e in unmatched_erp
                          if e["store_folder"] == store and e["platform"] == platform)
            cnt_erp_only = len([1 for e in unmatched_erp
                               if e["store_folder"] == store and e["platform"] == platform])

            # Unmatched Platform only
            plat_only = sum(platform_report_amount(p) for p in unmatched_platform
                           if p["store_folder"] == store and p["platform"] == platform)
            cnt_plat_only = len([1 for p in unmatched_platform
                                if p["store_folder"] == store and p["platform"] == platform])

            if cnt_matched == 0 and cnt_erp_only == 0 and cnt_plat_only == 0:
                continue

            rows.append({
                "Platform": platform,
                "Toko": display,
                "Cocok (ERP)": erp_matched,
                "Cocok (Platform)": plat_matched,
                "Jumlah Cocok": cnt_matched,
                "Hanya ERP (Rp)": erp_only,
                "Jumlah Hanya ERP": cnt_erp_only,
                "Hanya Platform (Rp)": plat_only,
                "Jumlah Hanya Platform": cnt_plat_only,
                "Total ERP": erp_matched + erp_only,
                "Total Platform": plat_matched + plat_only,
                "Selisih": (erp_matched + erp_only) - (plat_matched + plat_only),
            })

            platform_erp_total += erp_matched + erp_only
            platform_plat_total += plat_matched + plat_only
            platform_matched_count += cnt_matched
            platform_erp_only_count += cnt_erp_only
            platform_plat_only_count += cnt_plat_only

        # Platform subtotal
        if platform_erp_total > 0 or platform_plat_total > 0:
            rows.append({
                "Platform": f"  SUBTOTAL {platform}",
                "Toko": "",
                "Cocok (ERP)": "",
                "Cocok (Platform)": "",
                "Jumlah Cocok": platform_matched_count,
                "Hanya ERP (Rp)": "",
                "Jumlah Hanya ERP": platform_erp_only_count,
                "Hanya Platform (Rp)": "",
                "Jumlah Hanya Platform": platform_plat_only_count,
                "Total ERP": platform_erp_total,
                "Total Platform": platform_plat_total,
                "Selisih": platform_erp_total - platform_plat_total,
            })

    return rows


def generate_detail(matched, unmatched_erp, unmatched_platform):
    """Generate detail per transaction."""
    rows = []

    # Matched transactions
    for erp, plat in matched:
        selisih = erp["amount"] - plat["amount"]
        rows.append({
            "Status": "COCOK",
            "Toko": STORE_MAP.get(erp["store_folder"], {}).get("display", erp["store_folder"]),
            "Tanggal": erp["date"],
            "Platform": erp["platform"],
            "No Faktur ERP": erp.get("no_faktur", ""),
            "No Order Platform": plat.get("transaction_id", "") or plat.get("order_id", ""),
            "ERP Amount": erp["amount"],
            "Platform Amount": plat["amount"],
            "Selisih": selisih,
            "Keterangan": f"Source ERP: {erp.get('source_file', '')} | Source Platform: {plat.get('source_file', '')}",
        })

    # Unmatched ERP
    for erp in unmatched_erp:
        rows.append({
            "Status": "DI ERP SAJA",
            "Toko": STORE_MAP.get(erp["store_folder"], {}).get("display", erp["store_folder"]),
            "Tanggal": erp["date"],
            "Platform": erp["platform"],
            "No Faktur ERP": erp.get("no_faktur", ""),
            "No Order Platform": "",
            "ERP Amount": erp["amount"],
            "Platform Amount": 0,
            "Selisih": erp["amount"],
            "Keterangan": f"Source: {erp.get('source_file', '')}",
        })

    # Unmatched platform
    for plat in unmatched_platform:
        rows.append({
            "Status": "DI PLATFORM SAJA",
            "Toko": STORE_MAP.get(plat["store_folder"], {}).get("display", plat["store_folder"]),
            "Tanggal": plat["date"],
            "Platform": plat["platform"],
            "No Faktur ERP": "",
            "No Order Platform": plat.get("transaction_id", "") or plat.get("order_id", ""),
            "ERP Amount": 0,
            "Platform Amount": platform_report_amount(plat),
            "Selisih": -platform_report_amount(plat),
            "Keterangan": plat.get("data_issue") or f"Source: {plat.get('source_file', '')}",
        })

    # Sort by date, store, platform
    rows.sort(key=lambda r: (str(r["Tanggal"]), r["Toko"], r["Platform"], r["Status"]))

    return rows


# ============================================================
# OUTPUT
# ============================================================

def print_terminal_summary(summary_rows, detail_rows, matched, unmatched_erp, unmatched_platform, start_date, end_date):
    """Print reconciliation results to terminal."""
    print()
    print("=" * 90)
    print(f"  HASIL REKONSILIASI ONLINE FOOD")
    print(f"  Periode: {start_date} s/d {end_date}")
    print("=" * 90)

    if not summary_rows:
        print()
        print("  Tidak ada data transaksi online food dalam periode ini.")
        return

    # Summary per store per day
    print()
    print("-" * 90)
    print("  REKAP PER TOKO PER HARI")
    print("-" * 90)

    for row in summary_rows:
        print()
        print(f"  {row['Toko']}  |  {row['Tanggal']}")
        print(f"  {'Platform':<15} {'ERP':>12} {'Platform':>12} {'Selisih':>12}")
        print(f"  {'-'*15} {'-'*12} {'-'*12} {'-'*12}")

        if row["ERP GoFood"] > 0 or row["Platform GoFood"] > 0:
            print(f"  {'GoFood':<15} {row['ERP GoFood']:>12,.0f} {row['Platform GoFood']:>12,.0f} {row['Selisih GoFood']:>12,.0f}")
        if row["ERP GrabFood"] > 0 or row["Platform GrabFood"] > 0:
            print(f"  {'GrabFood':<15} {row['ERP GrabFood']:>12,.0f} {row['Platform GrabFood']:>12,.0f} {row['Selisih GrabFood']:>12,.0f}")
        if row["ERP ShopeeFood"] > 0 or row["Platform ShopeeFood"] > 0:
            print(f"  {'ShopeeFood':<15} {row['ERP ShopeeFood']:>12,.0f} {row['Platform ShopeeFood']:>12,.0f} {row['Selisih ShopeeFood']:>12,.0f}")

        print(f"  {'TOTAL':<15} {row['Total ERP']:>12,.0f} {row['Total Platform']:>12,.0f} {row['Selisih Total']:>12,.0f}")

    # Summary per platform per store
    platform_rows = generate_summary_per_platform(matched, unmatched_erp, unmatched_platform, start_date, end_date)
    if platform_rows:
        print()
        print("-" * 90)
        print("  REKAP PER PLATFORM PER TOKO")
        print("-" * 90)
        print()
        print(f"  {'Platform':<12} {'Toko':<22} {'Cocok':>8} {'Hanya ERP':>10} {'Hanya Plat':>10} {'Total ERP':>14} {'Total Plat':>14} {'Selisih':>14}")
        print(f"  {'-'*12} {'-'*22} {'-'*8} {'-'*10} {'-'*10} {'-'*14} {'-'*14} {'-'*14}")

        for row in platform_rows:
            platform = row["Platform"]
            tokoo = row["Toko"]
            if "SUBTOTAL" in str(platform):
                print(f"  {'':<12} {'':<22} {'':>8} {'':>10} {'':>10} {'':>14} {'':>14} {'':>14}")
                print(f"  {platform:<12} {'':<22} {row['Jumlah Cocok']:>8} {row['Jumlah Hanya ERP']:>10} {row['Jumlah Hanya Platform']:>10} {row['Total ERP']:>14,.0f} {row['Total Platform']:>14,.0f} {row['Selisih']:>14,.0f}")
                print(f"  {'-'*12} {'-'*22} {'-'*8} {'-'*10} {'-'*10} {'-'*14} {'-'*14} {'-'*14}")
            elif tokoo:
                print(f"  {platform:<12} {tokoo:<22} {row['Jumlah Cocok']:>8} {row['Jumlah Hanya ERP']:>10} {row['Jumlah Hanya Platform']:>10} {row['Total ERP']:>14,.0f} {row['Total Platform']:>14,.0f} {row['Selisih']:>14,.0f}")

    # Grand total
    print()
    print("-" * 90)
    print("  GRAND TOTAL")
    print("-" * 90)
    grand_erp = sum(r["Total ERP"] for r in summary_rows)
    grand_plat = sum(r["Total Platform"] for r in summary_rows)
    print(f"  ERP Total:      Rp {grand_erp:>15,.0f}")
    print(f"  Platform Total: Rp {grand_plat:>15,.0f}")
    print(f"  Selisih:        Rp {grand_erp - grand_plat:>15,.0f}")

    # Match stats
    n_matched = sum(1 for r in detail_rows if r["Status"] == "COCOK")
    n_erp_only = sum(1 for r in detail_rows if r["Status"] == "DI ERP SAJA")
    n_plat_only = sum(1 for r in detail_rows if r["Status"] == "DI PLATFORM SAJA")
    print()
    print(f"  Transaksi Cocok:       {n_matched}")
    print(f"  Hanya di ERP:          {n_erp_only}")
    print(f"  Hanya di Platform:     {n_plat_only}")

    # Detail
    print()
    print("-" * 90)
    print("  DETAIL PER TRANSAKSI")
    print("-" * 90)

    for row in detail_rows:
        status_icon = {"COCOK": "[OK]", "DI ERP SAJA": "[!!]", "DI PLATFORM SAJA": "[??]"}.get(row["Status"], "[--]")
        print(f"  {status_icon} {row['Toko']} | {row['Tanggal']} | {row['Platform']}")
        print(f"       ERP: {row['No Faktur ERP']:<25} Rp {row['ERP Amount']:>12,.0f}")
        print(f"       Plat: {row['No Order Platform']:<24} Rp {row['Platform Amount']:>12,.0f}")
        if row["Selisih"] != 0:
            print(f"       *** SELISIH: Rp {row['Selisih']:>12,.0f}")
        print()


def export_to_excel(summary_rows, detail_rows, matched, unmatched_erp, unmatched_platform, start_date, end_date, output_path):
    """Export results to Excel file."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: Summary
        if summary_rows:
            df_summary = pd.DataFrame(summary_rows)
            df_summary.to_excel(writer, sheet_name="Rekap per Toko", index=False)

            # Format date column
            ws = writer.sheets["Rekap per Toko"]
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    if cell.column == 2:  # Tanggal column
                        cell.number_format = "YYYY-MM-DD"
                    elif cell.column >= 3:  # Amount columns
                        cell.number_format = "#,##0"

        # Sheet 2: Platform Summary
        platform_summary = generate_summary_per_platform(matched, unmatched_erp, unmatched_platform, start_date, end_date)
        if platform_summary:
            df_platform = pd.DataFrame(platform_summary)
            df_platform.to_excel(writer, sheet_name="Rekap per Platform", index=False)

            ws = writer.sheets["Rekap per Platform"]
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    if cell.column >= 3:  # Amount columns
                        cell.number_format = "#,##0"

        # Sheet 3: Detail
        if detail_rows:
            df_detail = pd.DataFrame(detail_rows)
            df_detail.to_excel(writer, sheet_name="Detail Transaksi", index=False)

            ws = writer.sheets["Detail Transaksi"]
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    if cell.column == 3:  # Tanggal column
                        cell.number_format = "YYYY-MM-DD"
                    elif cell.column in [8, 9, 10]:  # Amount columns
                        cell.number_format = "#,##0"

        # Sheet 4: Unmatched ERP only
        unmatched_erp = [r for r in detail_rows if r["Status"] == "DI ERP SAJA"]
        if unmatched_erp:
            pd.DataFrame(unmatched_erp).to_excel(
                writer, sheet_name="Hanya di ERP", index=False
            )

        # Sheet 5: Unmatched Platform only
        unmatched_plat = [r for r in detail_rows if r["Status"] == "DI PLATFORM SAJA"]
        if unmatched_plat:
            pd.DataFrame(unmatched_plat).to_excel(
                writer, sheet_name="Hanya di Platform", index=False
            )

    print(f"\n  [OK] Hasil disimpan ke: {output_path}")


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Rekon Online Food - Reconcile ERP with platform settlements"
    )
    parser.add_argument("--start-date", type=str, help="Tanggal awal (YYYY-MM-DD)")
    parser.add_argument("--end-date", type=str, help="Tanggal akhir (YYYY-MM-DD)")
    parser.add_argument(
        "--project-path",
        type=str,
        default=os.path.dirname(os.path.abspath(__file__)),
        help="Path ke project folder (default: folder script ini)",
    )
    parser.add_argument(
        "--output", type=str, default=None, help="Path output Excel (default: auto)"
    )
    parser.add_argument(
        "--no-excel", action="store_true", help="Tidak export ke Excel"
    )
    parser.add_argument(
        "--map-merchant", action="store_true", 
        help="Mapping Merchant ID GoFood ke nama toko ERP"
    )
    parser.add_argument(
        "--auto-map-merchant", action="store_true",
        help="Auto-map semua Merchant ID berdasarkan nama folder"
    )

    args = parser.parse_args()
    project_path = args.project_path

    # Handle merchant ID mapping commands
    if args.map_merchant or args.auto_map_merchant:
        print()
        print("=" * 60)
        print("  MERCHANT ID MAPPING")
        print("=" * 60)
        print()
        
        mapped, unmapped = show_merchant_id_mapping(project_path)
        
        print(f"✓ Ter-mapping: {len(mapped)}")
        for entry in mapped:
            print(f"  {entry['merchant_id']} -> {entry['mapped_store']}")
        
        print()
        print(f"✗ Belum ter-mapping: {len(unmapped)}")
        for entry in unmapped:
            print(f"  {entry['merchant_id']} -> {entry['folder_name']}")
        
        if args.auto_map_merchant and unmapped:
            print()
            print("=== AUTO-MAPPING ===")
            for entry in unmapped:
                success = map_merchant_id_to_store(entry['merchant_id'], entry['folder_name'])
                if success:
                    print(f"  ✓ {entry['merchant_id']} -> {entry['folder_name']}")
            print()
            print("✓ Semua merchant ID sudah ter-mapping!")
        elif unmapped:
            print()
            print("Gunakan --auto-map-merchant untuk auto-mapping semua.")
            print("Atau edit store_mapping.json manual.")
        
        return

    # Get dates
    start_str = args.start_date
    end_str = args.end_date

    if not start_str:
        start_str = input("  Tanggal awal (YYYY-MM-DD): ").strip()
    if not end_str:
        end_str = input("  Tanggal akhir (YYYY-MM-DD): ").strip()

    try:
        start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        print("  [ERROR] Format tanggal salah. Gunakan YYYY-MM-DD")
        sys.exit(1)

    if start_date > end_date:
        print("  [ERROR] Tanggal awal lebih besar dari tanggal akhir")
        sys.exit(1)

    project_path = args.project_path

    print()
    print("=" * 60)
    print("  REKON ONLINE FOOD - Offline Reconciliation Tool")
    print("=" * 60)
    print(f"  Project path : {project_path}")
    print(f"  Periode      : {start_date} s/d {end_date}")
    print()

    # ---- Load ERP Data ----
    print("  [1/4] Loading data ERP...")
    penerimaan_files, transaksi_files = find_erp_files(project_path, start_date, end_date)
    print(f"        File penerimaan: {len(penerimaan_files)}")
    print(f"        File transaksi : {len(transaksi_files)}")

    erp_data = []
    for f in penerimaan_files:
        data = load_erp_penerimaan(f, start_date, end_date)
        erp_data.extend(data)
        print(f"        -> {os.path.basename(f)}: {len(data)} transaksi online food")

    for f in transaksi_files:
        data = load_erp_transaksi(f, start_date, end_date)
        erp_data.extend(data)
        print(f"        -> {os.path.basename(f)}: {len(data)} transaksi online food")

    print(f"        Total ERP: {len(erp_data)} transaksi")

    # ---- Load Platform Data ----
    print()
    print("  [2/4] Loading data platform...")

    print("        Grabfood:")
    grabfood_data = load_grabfood_reports(project_path, start_date, end_date)
    print(f"          {len(grabfood_data)} transaksi")

    print("        GoFood:")
    gofood_data = load_gofood_reports(project_path, start_date, end_date)
    print(f"          {len(gofood_data)} transaksi")

    print("        ShopeeFood:")
    shopeefood_data = load_shopeefood_reports(project_path, start_date, end_date)
    print(f"          {len(shopeefood_data)} transaksi")

    platform_data = grabfood_data + gofood_data + shopeefood_data
    print(f"        Total Platform: {len(platform_data)} transaksi")

    # ---- Reconcile ----
    print()
    print("  [3/4] Rekonsiliasi...")

    # Separate ERP by platform for targeted matching
    erp_gofood = [e for e in erp_data if e["platform"] == "GoFood"]
    erp_grabfood = [e for e in erp_data if e["platform"] == "GrabFood"]
    erp_shopee = [e for e in erp_data if e["platform"] == "ShopeeFood"]

    # Match per platform
    matched_gofood, un_erp_gofood, un_plat_gofood = reconcile(erp_gofood, gofood_data)
    matched_grabfood, un_erp_grabfood, un_plat_grabfood = reconcile(erp_grabfood, grabfood_data)
    matched_shopee, un_erp_shopee, un_plat_shopee = reconcile(
        erp_shopee, shopeefood_data, allow_fuzzy=False
    )

    all_matched = matched_gofood + matched_grabfood + matched_shopee
    all_unmatched_erp = un_erp_gofood + un_erp_grabfood + un_erp_shopee
    all_unmatched_platform = un_plat_gofood + un_plat_grabfood + un_plat_shopee

    print(f"        Cocok            : {len(all_matched)}")
    print(f"        Hanya di ERP     : {len(all_unmatched_erp)}")
    print(f"        Hanya di Platform: {len(all_unmatched_platform)}")

    # ---- Generate Reports ----
    print()
    print("  [4/4] Generate laporan...")

    summary_rows = generate_summary_per_store_day(
        all_matched, all_unmatched_erp, all_unmatched_platform, start_date, end_date
    )
    detail_rows = generate_detail(all_matched, all_unmatched_erp, all_unmatched_platform)

    # Print to terminal
    print_terminal_summary(summary_rows, detail_rows, all_matched, all_unmatched_erp, all_unmatched_platform, start_date, end_date)

    # Export to Excel
    if not args.no_excel:
        if args.output:
            output_path = args.output
        else:
            date_str = start_date.strftime("%Y%m%d")
            if start_date != end_date:
                date_str += f"-{end_date.strftime('%Y%m%d')}"
            output_dir = os.path.join(project_path, "Hasil Rekon")
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, f"Rekon_OnlineFood_{date_str}.xlsx")

        export_to_excel(summary_rows, detail_rows, all_matched, all_unmatched_erp, all_unmatched_platform, start_date, end_date, output_path)

    # Auto-save mapping if there are new stores
    if UNMAPPED_STORES:
        save_store_mapping()
        print()
        print("  [INFO] Toko baru terdeteksi dan sudah ditambahkan ke store_mapping.json:")
        for s in UNMAPPED_STORES:
            print(f"    - {s['name']} (folder: {s['folder']}, source: {s['source']})")
        print()


if __name__ == "__main__":
    main()
