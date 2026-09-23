import math
import os
import json
from datetime import date
from tempfile import TemporaryDirectory
from unittest import TestCase, main
from unittest.mock import patch

import openpyxl

import rekon
from webapp import app as webapp_module


class DataQualityTests(TestCase):
    def test_mapping_in_raw_data_folder_overrides_legacy_root_mapping(self):
        with TemporaryDirectory() as temporary_dir:
            raw_data_dir = os.path.join(temporary_dir, "Raw Data Transaksi")
            os.mkdir(raw_data_dir)
            with open(os.path.join(temporary_dir, "store_mapping.json"), "w") as f:
                json.dump({"stores": {"ROOT": {}}}, f)
            with open(os.path.join(raw_data_dir, "store_mapping.json"), "w") as f:
                json.dump({"stores": {"PONDOK KELAPA": {}}}, f)

            mapping = webapp_module.load_mapping_for_project(temporary_dir)

        self.assertIn("PONDOK KELAPA", mapping["stores"])

    def test_erp_penerimaan_accepts_gofood_case_variant(self):
        report = rekon.pd.DataFrame([{
            "Tipe Pembayaran": "Gofood",
            "Waktu Transaksi (POS)": "2026-09-10 09:41:07",
            "Nama Cabang Faktur Penjualan": "Procil Kios Pondok Kelapa",
            "Nomor # Faktur Penjualan": "PKP-100926.00038",
            "Total Penerimaan": 42500,
        }])
        with patch.object(rekon.pd, "read_excel", return_value=report), patch.object(
            rekon, "auto_detect_store", return_value=("PONDOK KELAPA", False)
        ):
            rows = rekon.load_erp_penerimaan(
                "erp.xlsx", date(2026, 9, 10), date(2026, 9, 10)
            )

        self.assertEqual(rows[0]["platform"], "GoFood")

    def test_grabfood_kaliabang_alias_is_mapped(self):
        self.assertEqual(
            rekon.map_platform_store(
                "Procil Bubur Tim Organik - Kios Kaliabang Tengah", "grabfood"
            ),
            "KALIABANG",
        )

    def test_platform_mapping_ignores_spreadsheet_apostrophe_and_extra_space(self):
        self.assertEqual(
            rekon.map_platform_store(
                "'Procil Bubur Tim Organik -  Kios Pondok Kelapa", "shopeefood"
            ),
            "PONDOK KELAPA",
        )

    def grab_rows(self, rows, diagnostics=None):
        report = rekon.pd.DataFrame(rows)
        with patch.object(rekon, "find_platform_reports", return_value=[{
            "path": "grab.csv",
            "filename": "grab.csv",
        }]), patch.object(rekon, "read_grabfood_report", return_value=report):
            return rekon.load_grabfood_reports(
                ".", date(2026, 7, 22), date(2026, 7, 22), diagnostics
            )

    def test_grabfood_cancelled_or_unpaid_is_ignored_with_diagnostic(self):
        diagnostics = []
        rows = self.grab_rows([{
            "Store Name": "Procil Bubur Organik - Perum Permata Mension",
            "Created On": "2026-07-22 06:42:00",
            "Status": "Cancelled",
            "Order Type": "Not Paid",
            "Short Order ID": "GF-884",
            "Net Sales": rekon.pd.NA,
        }], diagnostics)

        self.assertEqual(rows, [])
        payload = rekon.build_diagnostics(diagnostics)
        self.assertEqual(payload["summary"]["ignored_rows"], 1)
        self.assertEqual(payload["events"][0]["order_id"], "GF-884")

    def test_valid_grabfood_transfer_statuses_are_kept(self):
        rows = self.grab_rows([
            {
                "Store Name": "Procil Bubur Tim Organik - Kios Bintara",
                "Created On": f"2026-07-22 {index + 8:02d}:00:00",
                "Status": status,
                "Order Type": "Auto-Paid",
                "Transaction ID": f"G-{index}",
                "Net Sales": 10000 + index,
            }
            for index, status in enumerate(
                ["Transferred", "Completed", "Transfer failed"]
            )
        ])

        self.assertEqual(len(rows), 3)
        self.assertTrue(all(math.isfinite(row["amount"]) for row in rows))

    def test_active_invalid_amount_becomes_zero_and_cannot_match(self):
        diagnostics = []
        rows = self.grab_rows([{
            "Store Name": "Procil Bubur Tim Organik - Kios Bintara",
            "Created On": "2026-07-22 10:00:00",
            "Status": "Transferred",
            "Order Type": "Auto-Paid",
            "Transaction ID": "G-BAD",
            "Net Sales": float("nan"),
        }], diagnostics)

        self.assertEqual(rows[0]["amount"], 0)
        self.assertIn("tidak dicocokkan", rows[0]["data_issue"])
        payload = rekon.build_diagnostics(diagnostics)
        self.assertEqual(payload["summary"]["amount_defaulted"], 1)

        erp = [{
            "store_folder": "BINTARA",
            "date": date(2026, 7, 22),
            "amount": 0,
        }]
        matched, unmatched_erp, unmatched_platform = rekon.reconcile(erp, rows)
        self.assertEqual(matched, [])
        self.assertEqual(unmatched_erp, erp)
        self.assertEqual(unmatched_platform, rows)

    def test_finite_number_rejects_all_non_finite_values(self):
        for value in (None, "", float("nan"), float("inf"), float("-inf")):
            number, valid = rekon.finite_number(value)
            self.assertEqual(number, 0)
            self.assertFalse(valid)
            self.assertTrue(math.isfinite(number))

    def test_api_rejects_remaining_nan_with_valid_json_error(self):
        bad_results = {
            "summary_rows": [{
                "Toko": "Bintara",
                "Tanggal": date(2026, 7, 22),
                "Total ERP": 0,
                "Total Platform": float("nan"),
            }],
            "platform_summary_rows": [],
            "detail_rows": [],
            "diagnostics": rekon.build_diagnostics(),
        }
        client = webapp_module.app.test_client()
        with patch.object(
            webapp_module, "calculate_reconciliation", return_value=bad_results
        ):
            response = client.post("/api/rekon", json={
                "project_path": os.getcwd(),
                "start_date": "2026-07-22",
                "end_date": "2026-07-22",
            })

        self.assertEqual(response.status_code, 422)
        self.assertIn("tidak valid", response.get_json()["error"])
        self.assertNotIn("NaN", response.get_data(as_text=True))

    def test_export_always_contains_catatan_data_sheet(self):
        diagnostics = rekon.build_diagnostics([{
            "kind": "ignored",
            "code": "grabfood_cancelled_or_unpaid",
            "platform": "GrabFood",
            "date": "2026-07-22",
            "source_file": "grab.csv",
            "store": "Perum Permata Mension",
            "order_id": "GF-884",
            "field": "Status / Order Type",
            "original_value": "Cancelled / Not Paid",
            "action": "Dilewati dari rekonsiliasi",
            "message": "Transaksi dibatalkan atau tidak dibayar",
        }])
        with TemporaryDirectory() as temporary_dir:
            output = os.path.join(temporary_dir, "rekon.xlsx")
            rekon.export_to_excel(
                [], [], [], [], [],
                date(2026, 7, 22), date(2026, 7, 22), output,
                diagnostics=diagnostics,
            )
            workbook = openpyxl.load_workbook(output, read_only=True)
            sheet = workbook["Catatan Data"]
            values = list(sheet.values)

        self.assertEqual(values[0][0], "Jenis")
        self.assertEqual(values[1][0], "DILEWATI")
        self.assertEqual(values[1][5], "GF-884")


if __name__ == "__main__":
    main()
