import asyncio
import io
import unittest

import pandas as pd
from fastapi import UploadFile
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

import main
from insights import analytics_dataframe, anomaly_report


class DataLensFeatureTests(unittest.TestCase):
    def setUp(self):
        main.datasets.clear()
        main.reconciliation_reports.clear()
        self.content = self.make_workbook()

    def tearDown(self):
        main.datasets.clear()
        main.reconciliation_reports.clear()

    @staticmethod
    def make_workbook():
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Hesabat"
        sheet.merge_cells("A1:B1")
        sheet["A1"] = "Maliyyə"
        sheet.append(["ID", "Məbləğ", "Ad"])
        sheet.append(["A", 100, " A "])
        sheet["B3"].fill = PatternFill("solid", fgColor="FFFF00")
        sheet["C3"].comment = Comment("Müştəri adı mənbədə boşluqludur", "Tester")
        sheet.append(["A", 100, " A "])
        sheet.append(["B", -10, " B"])
        other = workbook.create_sheet("İkinci vərəq")
        other.append(["ID", "Məbləğ"])
        other.append(["A", 100])
        buffer = io.BytesIO()
        workbook.save(buffer)
        workbook.close()
        return buffer.getvalue()

    def upload(self):
        result = asyncio.run(
            main.upload_files(
                [
                    UploadFile(
                        filename="hesabat.xlsx",
                        file=io.BytesIO(self.content),
                    )
                ]
            )
        )
        return result["sheet_names"][0]["file_id"], result

    def test_upload_ingests_every_sheet_and_profile_reports_quality(self):
        file_id, result = self.upload()
        self.assertEqual(len(result["sheet_names"]), 2)
        self.assertEqual(result["sheet_names"][0]["sheet_name"], "Hesabat")
        profile = asyncio.run(
            main.profile_sheet(
                main.ValidationTarget(file_id=file_id, sheet_name="Hesabat")
            )
        )
        self.assertEqual(profile["rows"], 3)
        self.assertEqual(profile["exact_duplicate_rows"], 1)
        self.assertEqual(profile["trimmed_values"], 3)
        self.assertEqual(profile["quality_score"], 83)

    def test_merged_headers_fill_colors_and_comments_are_extracted(self):
        file_id, _ = self.upload()
        overview = asyncio.run(main.get_overview(file_id, "Hesabat"))
        columns = overview["columns"]
        self.assertEqual(columns[:2], ["Maliyyə / ID", "Maliyyə / Məbləğ"])
        formatting = asyncio.run(main.get_formatting(file_id, "Hesabat"))
        first_row = formatting["cells"]["0"]
        self.assertEqual(first_row["Maliyyə / Məbləğ"]["fill_rgb"], "FFFF00")
        self.assertEqual(
            first_row["Ad"]["comment"],
            "Müştəri adı mənbədə boşluqludur",
        )

    def test_color_analysis_groups_fills_and_calculates_numeric_totals(self):
        file_id, _ = self.upload()
        result = asyncio.run(
            main.analyze_fill_colors(main.ColorAnalysisRequest(file_id=file_id))
        )
        self.assertEqual(len(result["groups"]), 1)
        group = result["groups"][0]
        self.assertEqual(group["color"], "#FFFF00")
        self.assertEqual(group["row_count"], 1)
        self.assertEqual(group["numeric_metrics"][0]["sum"], 100)

    def test_reconciliation_classifies_matches_differences_and_one_sided_rows(self):
        left_id, right_id = "left", "right"
        main.datasets[left_id] = {
            "filename": "left.xlsx",
            "sheets": {
                "L": pd.DataFrame(
                    {"ID": ["A", "B", "C"], "Məbləğ": [10, 20, 30]}
                )
            },
        }
        main.datasets[right_id] = {
            "filename": "right.xlsx",
            "sheets": {
                "R": pd.DataFrame({"Ref": ["A", "B", "D"], "Cəm": [10, 25, 1]})
            },
        }
        result = asyncio.run(
            main.reconcile_sheets(
                main.ReconciliationRequest(
                    left_file_id=left_id,
                    left_sheet_name="L",
                    left_key_column="ID",
                    left_amount_column="Məbləğ",
                    right_file_id=right_id,
                    right_sheet_name="R",
                    right_key_column="Ref",
                    right_amount_column="Cəm",
                )
            )
        )
        self.assertEqual(result["counts"]["MATCH"], 1)
        self.assertEqual(result["counts"]["AMOUNT_DIFF"], 1)
        self.assertEqual(result["counts"]["LEFT_ONLY"], 1)
        self.assertEqual(result["counts"]["RIGHT_ONLY"], 1)
        self.assertEqual(result["totals"]["net_difference"], 24)

    def test_duckdb_analytics_windows_and_fuzzy_anomaly_detection(self):
        frame = pd.DataFrame(
            {
                "Region": ["North", "North", "South", "South"],
                "Amount": [10, 20, 20, -5],
                "Date": pd.to_datetime(
                    ["2025-01-01", "2025-01-02", "2025-01-03", "2025-01-04"]
                ),
                "Merchant": ["Acme", "Acmee", "Other", "Other"],
            }
        )
        result = analytics_dataframe(
            frame,
            group_by="Region",
            metric="Amount",
            order_by="Date",
        )
        self.assertEqual(result["engine"], "DuckDB")
        self.assertEqual(result["grouped_summary"][0]["row_count"], 2)
        self.assertEqual(len(result["window_analytics"]), 4)
        self.assertIn("dense_rank", result["window_analytics"][0])
        report = anomaly_report(frame)
        self.assertEqual(len(report["numeric_anomalies"]), 1)
        self.assertEqual(len(report["fuzzy_text_matches"]), 1)

    def test_full_export_contains_required_report_and_cleaned_sheets(self):
        file_id, _ = self.upload()
        main.reconciliation_reports[file_id] = {
            "rows": [
                {
                    "key": "A",
                    "left_amount": 10,
                    "right_amount": 9,
                    "difference": 1,
                    "status": "AMOUNT_DIFF",
                }
            ],
            "totals": {"left": 10, "right": 9, "net_difference": 1},
        }
        response = asyncio.run(
            main.export_full_workbook(
                main.FullExportRequest(file_id=file_id, sheet_name="Hesabat")
            )
        )
        exported = load_workbook(io.BytesIO(response.body), data_only=True)
        self.assertTrue(
            {
                "Xülasə",
                "AI auditi",
                "Rəng analizi",
                "Qruplaşdırma",
                "Uzlaşma",
                "Anomaliyalar",
                "Təmizlənmiş · Hesabat",
                "Təmizlənmiş · İkinci vərəq",
            }.issubset(exported.sheetnames)
        )
        reconciliation = exported["Uzlaşma"]
        self.assertEqual(reconciliation["E2"].value, "AMOUNT_DIFF")
        self.assertEqual(reconciliation["E2"].fill.fgColor.rgb, "00FFEB9C")
        exported.close()


if __name__ == "__main__":
    unittest.main()
