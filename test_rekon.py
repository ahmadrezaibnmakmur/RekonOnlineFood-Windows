from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import re
from zipfile import ZipFile

import rekon


def test_gofood_multi_outlet_with_title_row_and_merged_merchant_id():
    title_row = rekon.pd.DataFrame({"Laporan GoFood": ["rekap outlet"]})
    report = rekon.pd.DataFrame([
        {" Merchant ID ": "G847824388", "Waktu Transaksi": "11/06/2026 10:00", "Penjualan": 23000},
        {" Merchant ID ": "", "Waktu Transaksi": "11/06/2026 10:30", "Penjualan": 24000},
        {" Merchant ID ": "G803824393", "Waktu Transaksi": "11/06/2026 11:00", "Penjualan": 25000},
        {" Merchant ID ": "", "Waktu Transaksi": "11/06/2026 11:30", "Penjualan": 26000},
    ])

    def read_excel(_, header=0, **__):
        return title_row if header == 0 else report

    with patch.object(rekon, "find_platform_reports", return_value=[{
        "path": "gofood_multi.xlsx",
        "filename": "Laporan Gofood MULTI.xlsx",
    }]), patch.object(rekon.pd, "read_excel", side_effect=read_excel):
        rows = rekon.load_gofood_reports(".", date(2026, 6, 11), date(2026, 6, 11))

    assert [row["store_folder"] for row in rows] == [
        "BINTARA", "BINTARA", "SEROJA", "SEROJA"
    ]
    assert all(row["date"] == date(2026, 6, 11) for row in rows)


def test_platform_scan_includes_top_level_bulk_file():
    with TemporaryDirectory() as temporary_dir:
        base = Path(temporary_dir) / "Raw Data Transaksi" / "GoFood"
        dated = base / "2026-06-11"
        dated.mkdir(parents=True)
        (dated / "per-outlet.xlsx").touch()
        (base / "LAPORAN-BULK.XLSX").touch()

        reports = rekon.find_platform_reports(
            temporary_dir, "GoFood", date(2026, 6, 12), date(2026, 6, 12),
            scan_all_folders=True,
        )

    assert {report["filename"] for report in reports} == {
        "per-outlet.xlsx", "LAPORAN-BULK.XLSX"
    }


def test_gofood_loader_repairs_reversed_xlsx_cell_references():
    report = rekon.pd.DataFrame([
        {"Merchant ID": "G847824388", "Waktu transaksi": "2026-06-11 10:00", "Penjualan": 23000},
        {"Merchant ID": "G803824393", "Waktu transaksi": "2026-06-11 11:00", "Penjualan": 24000},
    ])

    with TemporaryDirectory() as temporary_dir:
        valid = Path(temporary_dir) / "valid.xlsx"
        broken = Path(temporary_dir) / "gofood-multi-outlet.xlsx"
        report.to_excel(valid, index=False)
        with ZipFile(valid) as source, ZipFile(broken, "w") as target:
            for entry in source.infolist():
                content = source.read(entry.filename)
                if entry.filename.startswith("xl/worksheets/") and entry.filename.endswith(".xml"):
                    content = re.sub(
                        rb'(<c\b[^>]*\br=")([A-Z]+)(\d+)(")',
                        lambda match: match.group(1) + match.group(3) + match.group(2) + match.group(4),
                        content,
                    )
                target.writestr(entry, content)

        with patch.object(rekon, "find_platform_reports", return_value=[{
            "path": str(broken),
            "filename": broken.name,
        }]):
            rows = rekon.load_gofood_reports(".", date(2026, 6, 11), date(2026, 6, 11))

    assert [row["store_folder"] for row in rows] == ["BINTARA", "SEROJA"]


def test_grabfood_csv_is_scanned_and_loaded():
    report = rekon.pd.DataFrame([{
        "Store Name": "Procil Bubur Tim Organik - Kios Bintara",
        "Created On": "2026-06-11 10:00:00",
        "Transaction ID": "GRAB-1",
        "Net Sales": 23000,
    }])

    with TemporaryDirectory() as temporary_dir:
        base = Path(temporary_dir) / "Raw Data Transaksi" / "Grabfood" / "2026-06-11"
        base.mkdir(parents=True)
        (base / "grab-report.CSV").touch()
        reports = rekon.find_platform_reports(
            temporary_dir, "Grabfood", date(2026, 6, 11), date(2026, 6, 11),
            extensions=(".xlsx", ".csv"),
        )

    with patch.object(rekon, "find_platform_reports", return_value=reports), \
            patch.object(rekon.pd, "read_csv", return_value=report):
        rows = rekon.load_grabfood_reports(".", date(2026, 6, 11), date(2026, 6, 11))

    assert [row["store_folder"] for row in rows] == ["BINTARA"]
    assert rows[0]["amount"] == 23000


if __name__ == "__main__":
    test_gofood_multi_outlet_with_title_row_and_merged_merchant_id()
    test_platform_scan_includes_top_level_bulk_file()
    test_gofood_loader_repairs_reversed_xlsx_cell_references()
    test_grabfood_csv_is_scanned_and_loaded()
    print("ok")
