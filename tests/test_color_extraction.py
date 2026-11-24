import asyncio
import io
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, Font
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
import pandas as pd
from fastapi import HTTPException

import main


class ColorExtractionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.workbook_bytes = self.make_workbook()
        main.datasets.clear()
        main.datasets["color-test"] = {
            "filename": "color-test.xlsx",
            "file_extension": ".xlsx",
            "workbook_bytes": self.workbook_bytes,
            "sheets": {
                "Movements": pd.DataFrame(
                    columns=["ID", "Description", "Type", "Amount"]
                ),
                "Different schema": pd.DataFrame(columns=["Name", "Total"]),
                "Conditional": pd.DataFrame(columns=["Name"]),
                "Theme": pd.DataFrame(columns=["Theme color"]),
            },
            "previews": {},
            "history": {},
            "cleaning_log": {},
        }

    def tearDown(self) -> None:
        main.datasets.clear()

    @staticmethod
    def make_workbook() -> bytes:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Movements"
        sheet.append(["ID", "Description", "Type", "Amount"])
        sheet.append([1, "red and purple", "shared row", 50])
        sheet["B2"].font = Font(color="FFFF0000")
        sheet["C2"].font = Font(color="FFA568D2")
        sheet["D2"].font = Font(bold=True)
        sheet["D2"].number_format = '"+ "##0.00" AZN saxla'
        sheet.append([2, "ordinary default black", "not selected", 20])
        sheet.append([3, "near red", "not selected", 30])
        sheet["B4"].font = Font(color="FFFF0001")
        sheet.append([4, "explicit black", "selected black", 40])
        sheet["B5"].font = Font(color="FF000000")
        sheet.merge_cells("E1:F1")
        sheet["E1"] = "Merged header"
        sheet["F2"] = "value under merged header"

        other = workbook.create_sheet("Different schema")
        other.append(["Name", "Total"])
        other.append(["blue ARGB", 100])
        other["A2"].font = Font(color="FF00B0F0")
        other.append(["indexed red", 200])
        other["A3"].font = Font(color=Color(indexed=2))

        conditional = workbook.create_sheet("Conditional")
        conditional.append(["Name"])
        conditional.append(["duplicate"])
        conditional.append(["duplicate"])
        conditional.conditional_formatting.add(
            "A2:A3",
            Rule(
                type="duplicateValues",
                dxf=DifferentialStyle(font=Font(color="FF9C0006")),
            ),
        )

        theme_sheet = workbook.create_sheet("Theme")
        theme_sheet.append(["Theme color"])
        theme_sheet.append(["theme blue"])
        theme_sheet["A2"].font = Font(color=Color(theme=4))
        buffer = io.BytesIO()
        workbook.save(buffer)
        workbook.close()
        themed_workbook = load_workbook(io.BytesIO(buffer.getvalue()))
        theme = themed_workbook.loaded_theme.decode("utf-8")
        themed_workbook.loaded_theme = theme.replace(
            "4F81BD",
            "0070C0",
        ).encode("utf-8")
        themed_buffer = io.BytesIO()
        themed_workbook.save(themed_buffer)
        themed_workbook.close()
        return themed_buffer.getvalue()

    def test_extracts_selected_font_colors_across_sheets(self) -> None:
        grouped, headers = main.extract_color_rows(
            self.workbook_bytes,
            ["red", "purple", "black", "blue"],
        )

        self.assertEqual(len(grouped["red"]), 4)
        self.assertEqual(len(grouped["purple"]), 1)
        self.assertEqual(len(grouped["black"]), 1)
        self.assertEqual(len(grouped["blue"]), 2)
        self.assertEqual(grouped["red"][0]["source_sheet"], "Movements")
        self.assertEqual(grouped["red"][0]["source_row"], 2)
        self.assertEqual(grouped["red"][0]["matched_cells"], ["B2"])
        self.assertEqual(grouped["purple"][0]["matched_columns"], ["Type"])
        self.assertEqual(grouped["blue"][0]["source_sheet"], "Different schema")
        self.assertEqual(grouped["blue"][1]["source_sheet"], "Theme")
        self.assertEqual(grouped["red"][2]["source_sheet"], "Conditional")
        self.assertEqual(grouped["black"][0]["matched_cells"], ["D2"])
        self.assertIn("Name", headers)
        self.assertIn("ID", headers)
        self.assertIn("Sütun F", headers)

    def test_default_black_text_and_similar_colors_are_not_selected(self) -> None:
        grouped, _ = main.extract_color_rows(self.workbook_bytes, ["black", "red"])
        self.assertNotIn(
            "ordinary default black",
            [row["values"].get("Description") for row in grouped["black"]],
        )
        self.assertNotIn(
            "near red",
            [row["values"].get("Description") for row in grouped["red"]],
        )

    def test_header_detection_handles_read_only_worksheets(self) -> None:
        workbook = load_workbook(io.BytesIO(self.workbook_bytes), read_only=True)
        worksheet = workbook["Movements"]
        self.assertEqual(
            main.worksheet_header_row(
                worksheet,
                worksheet.max_row,
                worksheet.max_column,
            ),
            1,
        )
        workbook.close()

    def test_preview_includes_counts_sheets_and_overlapping_groups(self) -> None:
        response = asyncio.run(
            main.preview_color_extraction(
                main.ColorExtractionRequest(
                    file_id="color-test",
                    colors=["red", "purple"],
                )
            )
        )
        self.assertEqual(response["scanned_sheets"], 4)
        self.assertEqual(response["colors"], ["red", "purple"])
        self.assertEqual(response["groups"][0]["count"], 4)
        self.assertEqual(response["groups"][1]["count"], 1)
        self.assertEqual(response["groups"][0]["preview_rows"][0]["excel_row"], 2)
        self.assertEqual(response["groups"][0]["preview_rows"][0]["values"]["ID"], 1)

    def test_export_contains_one_sheet_per_group_and_source_metadata(self) -> None:
        response = asyncio.run(
            main.export_color_extraction(
                main.ColorExtractionRequest(
                    file_id="color-test",
                    colors=["red", "blue"],
                )
            )
        )
        exported = load_workbook(io.BytesIO(response.body), data_only=False)
        self.assertEqual(exported.sheetnames, ["Qırmızı", "Mavi"])
        red_sheet = exported["Qırmızı"]
        self.assertEqual(
            [red_sheet.cell(1, column).value for column in range(1, 5)],
            ["Mənbə vərəqi", "Excel sətir №", "Uyğun sütunlar", "Uyğun xanalar"],
        )
        self.assertEqual(red_sheet["A2"].value, "Movements")
        self.assertEqual(red_sheet["B2"].value, 2)
        self.assertEqual(red_sheet["C2"].value, "Description")
        self.assertEqual(red_sheet["D2"].value, "B2")
        self.assertEqual(exported["Mavi"].max_row, 3)
        exported.close()

    def test_csv_and_legacy_xls_are_rejected_with_clear_message(self) -> None:
        for extension in (".csv", ".xls"):
            with self.subTest(extension=extension):
                main.datasets["unsupported"] = {
                    "filename": f"legacy{extension}",
                    "file_extension": extension,
                    "workbook_bytes": None,
                    "sheets": {},
                }
                with self.assertRaises(HTTPException) as error:
                    asyncio.run(
                        main.preview_color_extraction(
                            main.ColorExtractionRequest(
                                file_id="unsupported",
                                colors=["red"],
                            )
                        )
                    )
                self.assertEqual(error.exception.status_code, 400)
                self.assertIn(".xlsx", error.exception.detail)

    def test_unknown_file_is_not_found(self) -> None:
        with self.assertRaises(HTTPException) as error:
            asyncio.run(
                main.preview_color_extraction(
                    main.ColorExtractionRequest(
                        file_id="missing",
                        colors=["red"],
                    )
                )
            )
        self.assertEqual(error.exception.status_code, 404)


if __name__ == "__main__":
    unittest.main()
