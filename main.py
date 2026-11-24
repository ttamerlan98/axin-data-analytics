from __future__ import annotations

import csv
import json
import logging
import math
import re
import colorsys
from collections import Counter
from io import BytesIO, StringIO
from datetime import date, datetime
from pathlib import Path
from typing import Any, Literal
from urllib.parse import quote
from uuid import uuid4
from xml.etree import ElementTree

import pandas as pd
from fastapi import FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field


PROJECT_DIR = Path(__file__).resolve().parent
logger = logging.getLogger(__name__)
app = FastAPI(title="Local Data Analytics API")
app.mount(
    "/static",
    StaticFiles(directory=PROJECT_DIR / "static"),
    name="static",
)

# Data is kept in process memory and is cleared when the server restarts.
datasets: dict[str, dict[str, Any]] = {}


@app.exception_handler(RequestValidationError)
async def validation_error_handler(
    request: Request,
    error: RequestValidationError,
) -> JSONResponse:
    logger.info(
        "Invalid request for %s (%d validation errors)",
        request.url.path,
        len(error.errors()),
    )
    return JSONResponse(
        status_code=422,
        content={"detail": "Sorğu məlumatları etibarsızdır. Daxil etdiyiniz dəyərləri yoxlayın."},
    )


@app.exception_handler(Exception)
async def unexpected_error_handler(request: Request, error: Exception) -> JSONResponse:
    logger.error(
        "Unhandled API error for %s",
        request.url.path,
        exc_info=(type(error), error, error.__traceback__),
    )
    return JSONResponse(
        status_code=500,
        content={"detail": "Sorğunu emal etmək mümkün olmadı. Məlumatları yoxlayıb yenidən cəhd edin."},
    )


class CleaningRequest(BaseModel):
    file_id: str
    sheet_name: str
    remove_duplicates: bool = False
    missing_values: str = Field(default="none", pattern="^(none|zero|mean|drop_rows)$")
    standardize_text: bool = False
    standardize_dates: bool = False


class CleaningTarget(BaseModel):
    file_id: str
    sheet_name: str


class ValidationTarget(BaseModel):
    file_id: str
    sheet_name: str


class MergeSource(BaseModel):
    file_id: str
    sheet_name: str
    column_mapping: dict[str, str] = Field(default_factory=dict)


class MergeRequest(BaseModel):
    sources: list[MergeSource] = Field(min_length=1)


class ColorExtractionRequest(BaseModel):
    file_id: str
    colors: list[Literal["red", "purple", "black", "blue"]] = Field(min_length=1)


COLOR_LABELS = {
    "red": "Qırmızı",
    "purple": "Bənövşəyi",
    "black": "Qara",
    "blue": "Mavi",
}
COLOR_RGB_VALUES = {
    "red": {"FF0000", "9C0006"},
    "purple": {"A568D2", "7030A0"},
    "blue": {"0070C0", "00B0F0"},
}
COLOR_METADATA_COLUMNS = (
    "Mənbə vərəqi",
    "Excel sətir №",
    "Uyğun sütunlar",
    "Uyğun xanalar",
)
WORKBOOK_THEME_NAMESPACE = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
}


def normalized_rgb(rgb: str | None, tint: float = 0.0) -> str | None:
    if not rgb:
        return None
    value = rgb.strip().lstrip("#").upper()
    if len(value) == 8:
        value = value[-6:]
    if len(value) != 6:
        return None
    try:
        red, green, blue = (int(value[offset : offset + 2], 16) for offset in (0, 2, 4))
    except ValueError:
        return None
    if tint:
        hue, lightness, saturation = colorsys.rgb_to_hls(
            red / 255,
            green / 255,
            blue / 255,
        )
        if tint < 0:
            lightness *= 1 + tint
        else:
            lightness = lightness * (1 - tint) + tint
        red, green, blue = (
            round(channel * 255)
            for channel in colorsys.hls_to_rgb(hue, lightness, saturation)
        )
    return f"{red:02X}{green:02X}{blue:02X}"


def workbook_theme_colors(workbook: Any) -> dict[int, str]:
    theme = workbook.loaded_theme
    if not theme:
        return {}
    try:
        root = ElementTree.fromstring(theme)
    except ElementTree.ParseError:
        return {}
    color_scheme = root.find(".//a:clrScheme", WORKBOOK_THEME_NAMESPACE)
    if color_scheme is None:
        return {}

    colors: dict[int, str] = {}
    for index, theme_entry in enumerate(color_scheme):
        if not len(theme_entry):
            continue
        color_attributes = theme_entry[0].attrib
        color_value = color_attributes.get("lastClr") or color_attributes.get("val")
        normalized = normalized_rgb(color_value)
        if normalized:
            colors[index] = normalized
    return colors


def cell_font_rgb(cell: Any, theme_colors: dict[int, str]) -> str | None:
    return excel_color_rgb(cell.font.color, theme_colors)


def excel_color_rgb(color: Any, theme_colors: dict[int, str]) -> str | None:
    if color is None:
        return None
    tint = float(color.tint or 0)
    if color.type == "rgb":
        return normalized_rgb(color.rgb, tint)
    if color.type == "theme":
        base = theme_colors.get(color.theme)
    elif (
        color.type == "indexed"
        and isinstance(color.indexed, int)
        and 0 <= color.indexed < len(COLOR_INDEX)
    ):
        base = normalized_rgb(COLOR_INDEX[color.indexed])
    else:
        return None
    return normalized_rgb(base, tint)


def is_color_match(cell: Any, color_name: str, theme_colors: dict[int, str]) -> bool:
    if color_name == "black":
        return (
            cell.font.bold
            and "azn saxla" in cell.number_format.casefold()
            and (
                cell.font.color is None
                or cell_font_rgb(cell, theme_colors) == "000000"
            )
        )
    return cell_font_rgb(cell, theme_colors) in COLOR_RGB_VALUES[color_name]


def conditional_font_color_cells(
    worksheet: Any,
    selected_colors: list[str],
    theme_colors: dict[int, str],
) -> dict[str, set[str]]:
    matched_coordinates = {color: set() for color in selected_colors}
    for conditional_format in worksheet.conditional_formatting:
        for rule in worksheet.conditional_formatting[conditional_format]:
            if rule.type != "duplicateValues" or not rule.dxf or not rule.dxf.font:
                continue
            rgb = excel_color_rgb(rule.dxf.font.color, theme_colors)
            matching_colors = [
                color
                for color in selected_colors
                if color != "black" and rgb in COLOR_RGB_VALUES[color]
            ]
            if not matching_colors:
                continue

            for cell_range in conditional_format.sqref.ranges:
                cells = [
                    cell
                    for row in worksheet.iter_rows(
                        min_row=cell_range.min_row,
                        max_row=cell_range.max_row,
                        min_col=cell_range.min_col,
                        max_col=cell_range.max_col,
                    )
                    for cell in row
                    if cell.value is not None
                ]
                counts = Counter(
                    (
                        "text",
                        cell.value.casefold(),
                    )
                    if isinstance(cell.value, str)
                    else ("number", float(cell.value))
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                    else (type(cell.value).__name__, cell.value)
                    for cell in cells
                )
                for cell in cells:
                    value_key = (
                        ("text", cell.value.casefold())
                        if isinstance(cell.value, str)
                        else ("number", float(cell.value))
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                        else (type(cell.value).__name__, cell.value)
                    )
                    if counts[value_key] > 1:
                        for color in matching_colors:
                            matched_coordinates[color].add(cell.coordinate)
    return matched_coordinates


def unique_column_headers(cells: list[Any]) -> list[str]:
    used = set(COLOR_METADATA_COLUMNS)
    headers: list[str] = []
    for cell in cells:
        base = (
            str(cell.value)
            if cell.value is not None and str(cell.value).strip()
            else f"Sütun {get_column_letter(cell.column)}"
        )
        candidate = base
        suffix = 2
        while candidate in used:
            candidate = f"{base} ({suffix})"
            suffix += 1
        used.add(candidate)
        headers.append(candidate)
    return headers


def json_safe_cell(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def worksheet_header_row(worksheet: Any, max_data_row: int, max_data_column: int) -> int:
    minimum_values = min(3, max_data_column)
    for row_number, row in enumerate(
        worksheet.iter_rows(
            min_row=1,
            max_row=max_data_row,
            max_col=max_data_column,
        ),
        start=1,
    ):
        values = [cell.value for cell in row if cell.value is not None]
        text_values = sum(isinstance(value, str) for value in values)
        if (
            len(values) >= minimum_values
            and text_values >= minimum_values
            and text_values / len(values) >= 0.65
        ):
            return row_number
    return 0


def extract_color_rows(
    workbook_bytes: bytes,
    selected_colors: list[str],
) -> tuple[dict[str, list[dict[str, Any]]], list[str]]:
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True)
    formula_workbook = load_workbook(BytesIO(workbook_bytes), data_only=False)
    theme_colors = workbook_theme_colors(workbook)
    grouped_rows: dict[str, list[dict[str, Any]]] = {
        color: [] for color in selected_colors
    }
    all_headers: list[str] = []

    for worksheet, formula_worksheet in zip(workbook.worksheets, formula_workbook.worksheets):
        nonempty_cells = [
            cell
            for row in formula_worksheet.iter_rows()
            for cell in row
            if cell.value is not None
        ]
        max_data_column = max((cell.column for cell in nonempty_cells), default=0)
        max_data_row = max((cell.row for cell in nonempty_cells), default=0)
        if max_data_column == 0:
            continue

        header_row = worksheet_header_row(
            worksheet,
            max_data_row,
            max_data_column,
        )
        if header_row:
            headers = unique_column_headers(
                list(
                    worksheet.iter_rows(
                        min_row=header_row,
                        max_row=header_row,
                        max_col=max_data_column,
                    )
                )[0]
            )
        else:
            headers = [
                f"Sütun {get_column_letter(column)}"
                for column in range(1, max_data_column + 1)
            ]
        for header in headers:
            if header not in all_headers:
                all_headers.append(header)

        conditional_matches = conditional_font_color_cells(
            worksheet,
            selected_colors,
            theme_colors,
        )
        for row in worksheet.iter_rows(
            min_row=header_row + 1 if header_row else 1,
            max_row=max_data_row,
            max_col=max_data_column,
        ):
            matched_cells: dict[str, list[str]] = {
                color: [] for color in selected_colors
            }
            matched_columns: dict[str, list[str]] = {
                color: [] for color in selected_colors
            }
            for cell, header in zip(row, headers):
                formula_cell = formula_worksheet[cell.coordinate]
                value = (
                    formula_cell.value
                    if cell.value is None and formula_cell.data_type == "f"
                    else cell.value
                )
                if value is None:
                    continue
                for color in selected_colors:
                    if (
                        is_color_match(cell, color, theme_colors)
                        or cell.coordinate in conditional_matches[color]
                    ):
                        matched_cells[color].append(cell.coordinate)
                        matched_columns[color].append(header)

            values = {
                header: (
                    formula_worksheet[cell.coordinate].value
                    if cell.value is None
                    and formula_worksheet[cell.coordinate].data_type == "f"
                    else cell.value
                )
                for cell, header in zip(row, headers)
            }
            for color in selected_colors:
                if matched_cells[color]:
                    grouped_rows[color].append(
                        {
                            "source_sheet": worksheet.title,
                            "source_row": row[0].row,
                            "matched_cells": matched_cells[color],
                            "matched_columns": matched_columns[color],
                            "values": values,
                        }
                    )

    workbook.close()
    formula_workbook.close()
    return grouped_rows, all_headers


def color_extraction_summary(
    grouped_rows: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    for color, rows in grouped_rows.items():
        sheet_counts: dict[str, int] = {}
        for row in rows:
            sheet = row["source_sheet"]
            sheet_counts[sheet] = sheet_counts.get(sheet, 0) + 1
        groups.append(
            {
                "color": color,
                "label": COLOR_LABELS[color],
                "count": len(rows),
                "sheets": [
                    {"sheet_name": sheet, "count": count}
                    for sheet, count in sheet_counts.items()
                ],
                "preview_rows": [
                    {
                        "sheet_name": row["source_sheet"],
                        "excel_row": row["source_row"],
                        "matched_cells": row["matched_cells"],
                        "matched_columns": row["matched_columns"],
                        "values": {
                            column: json_safe_cell(value)
                            for column, value in row["values"].items()
                        },
                    }
                    for row in rows[:5]
                ],
            }
        )
    return groups


def color_extraction_workbook(
    grouped_rows: dict[str, list[dict[str, Any]]],
    all_headers: list[str],
) -> bytes:
    workbook = Workbook()
    workbook.remove(workbook.active)
    output_headers = [*COLOR_METADATA_COLUMNS, *all_headers]

    for color, rows in grouped_rows.items():
        worksheet = workbook.create_sheet(COLOR_LABELS[color])
        worksheet.append(output_headers)
        for row in rows:
            values = [
                row["source_sheet"],
                row["source_row"],
                ", ".join(row["matched_columns"]),
                ", ".join(row["matched_cells"]),
                *(row["values"].get(header) for header in all_headers),
            ]
            worksheet.append(values)
            for cell in worksheet[worksheet.max_row]:
                if isinstance(cell.value, str) and cell.value.startswith(("=", "+", "-", "@")):
                    cell.data_type = "s"
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            column_letter = column_cells[0].column_letter
            max_length = max(
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=10,
            )
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 50)

    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def get_dataset_sheet(file_id: str, sheet_name: str) -> tuple[dict[str, Any], pd.DataFrame]:
    dataset = datasets.get(file_id)
    if dataset is None:
        raise HTTPException(status_code=404, detail="File not found")
    dataframe = dataset["sheets"].get(sheet_name)
    if dataframe is None:
        raise HTTPException(status_code=404, detail="Sheet not found")
    return dataset, dataframe


def json_rows(dataframe: pd.DataFrame) -> list[dict[str, Any]]:
    serialized = dataframe.head(100).to_json(
        orient="records",
        date_format="iso",
    )
    rows = json.loads(serialized)
    json.dumps(rows, allow_nan=False)
    return rows


def read_csv_with_fallback(content: bytes) -> pd.DataFrame:
    def parse_csv_text(text: str) -> pd.DataFrame:
        try:
            dialect = csv.Sniffer().sniff(text[:65536], delimiters=",;\t|")
            delimiter = dialect.delimiter
        except csv.Error:
            delimiter = ","

        try:
            records = csv.reader(
                StringIO(text),
                delimiter=delimiter,
                strict=True,
            )
            field_counts = {len(record) for record in records if record}
        except csv.Error as error:
            raise ValueError("CSV faylının quruluşu düzgün deyil.") from error

        if len(field_counts) > 1:
            raise ValueError("CSV faylında sütun sayı sətirlər arasında uyğun deyil.")

        return pd.read_csv(StringIO(text), sep=delimiter, engine="python")

    try:
        decoded = content.decode("utf-8-sig")
    except UnicodeDecodeError:
        decoded = ""
    else:
        return parse_csv_text(decoded)

    candidates: list[tuple[int, str]] = []
    errors: list[str] = []
    for encoding in ("cp1254", "cp1251"):
        try:
            text = content.decode(encoding)
        except UnicodeDecodeError as error:
            errors.append(f"{encoding}: {error}")
            continue
        cyrillic = len(re.findall(r"[\u0400-\u04ff]", text))
        turkish = len(re.findall(r"[çğıöşüÇĞİÖŞÜ]", text))
        score = 2 * cyrillic + 3 * turkish
        candidates.append((score, text))

    if candidates:
        _, selected_text = max(candidates, key=lambda candidate: candidate[0])
        return parse_csv_text(selected_text)
    raise ValueError(
        "CSV faylının kodlaşdırılması UTF-8, Windows-1254 və ya Windows-1251 "
        "olmalıdır. " + "; ".join(errors)
    )


def export_filename(file_id: str, extension: str) -> str:
    filename = str(datasets[file_id]["filename"])
    stem = Path(filename).stem
    safe_stem = re.sub(r"[^\w.-]+", "_", stem, flags=re.UNICODE).strip("._")
    return f"{safe_stem or 'melumat'}.{extension}"


def attachment_header(filename: str) -> str:
    fallback = filename.encode("ascii", "ignore").decode("ascii") or "export"
    encoded = quote(filename, safe="")
    return f'attachment; filename="{fallback}"; filename*=UTF-8\'\'{encoded}'


def sql_identifier(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def sql_value(value: Any, dialect: str) -> str:
    if value is None or pd.isna(value):
        return "NULL"
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, (datetime, date)):
        text = value.strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, datetime) else value.isoformat()
        if dialect == "oracle":
            return f"TO_DATE('{text}', 'YYYY-MM-DD HH24:MI:SS')" if isinstance(value, datetime) else f"TO_DATE('{text}', 'YYYY-MM-DD')"
        return f"'{text}'"
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(value) if math.isfinite(value) else "NULL"
    return "'" + str(value).replace("'", "''") + "'"


def sql_column_type(dataframe: pd.DataFrame, column: Any, dialect: str) -> str:
    series = dataframe[column]
    if pd.api.types.is_bool_dtype(series):
        return "NUMBER(1)" if dialect == "oracle" else "INTEGER"
    if pd.api.types.is_integer_dtype(series):
        return "NUMBER(38, 0)" if dialect == "oracle" else "INTEGER"
    if pd.api.types.is_float_dtype(series):
        return "NUMBER" if dialect == "oracle" else "REAL"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "DATE" if dialect == "oracle" else "TEXT"
    return "VARCHAR2(4000)" if dialect == "oracle" else "TEXT"


def sql_script(dataframe: pd.DataFrame, table_name: str, dialect: str) -> str:
    columns = list(dataframe.columns)
    table_identifier = sql_identifier(table_name)
    column_definitions = ",\n".join(
        f"  {sql_identifier(str(column))} {sql_column_type(dataframe, column, dialect)}"
        for column in columns
    )
    lines = [
        f"-- SQL dialect: {dialect.upper()}",
        f"CREATE TABLE {table_identifier} (\n{column_definitions}\n);",
        "",
    ]
    if dialect == "sqlite":
        lines.append("BEGIN TRANSACTION;")

    column_list = ", ".join(sql_identifier(str(column)) for column in columns)
    for row in dataframe.itertuples(index=False, name=None):
        values = ", ".join(sql_value(value, dialect) for value in row)
        lines.append(
            f"INSERT INTO {table_identifier} ({column_list}) VALUES ({values});"
        )

    if dialect == "sqlite":
        lines.append("COMMIT;")
    return "\n".join(lines) + "\n"


def standardize_az_text(value: str) -> str:
    normalized = re.sub(r"\s+", " ", value.strip())
    if normalized.casefold() == "baki":
        return "Bakı"

    normalized = normalized.replace("İ", "i").replace("I", "ı").lower()

    def capitalize_word(word: str) -> str:
        for index, character in enumerate(word):
            if character.isalpha():
                upper = "İ" if character == "i" else "I" if character == "ı" else character.upper()
                return word[:index] + upper + word[index + 1 :]
        return word

    return " ".join(capitalize_word(word) for word in normalized.split(" "))


def parse_mixed_dates(values: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(values, format="ISO8601", errors="coerce")
    fallback_mask = values.notna() & parsed.isna()
    if fallback_mask.any():
        fallback = pd.to_datetime(
            values[fallback_mask],
            format="mixed",
            errors="coerce",
            dayfirst=True,
        )
        parsed.loc[fallback_mask] = fallback
    return parsed


def date_column(dataframe: pd.DataFrame, column: Any) -> bool:
    series = dataframe[column]
    if pd.api.types.is_datetime64_any_dtype(series):
        return True

    label_suggests_date = bool(
        re.search(r"(^|[_\s/-])(date|time|tarix|gün|ay)([_\s/-]|$)", str(column), re.I)
        or re.search(r"date|time|tarix", str(column), re.I)
    )
    values = series.dropna()
    if values.empty:
        return False
    if not pd.api.types.is_object_dtype(series) and not pd.api.types.is_string_dtype(series):
        return False

    string_values = values.astype(str)
    has_date_pattern = string_values.str.contains(
        r"\d{1,4}[-/.]\d{1,2}|\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)",
        case=False,
        regex=True,
    ).any()
    if not label_suggests_date and not has_date_pattern:
        return False

    parsed = parse_mixed_dates(values)
    return label_suggests_date or parsed.notna().mean() >= 0.6


def clean_dataframe(
    original: pd.DataFrame,
    options: CleaningRequest,
) -> tuple[pd.DataFrame, list[str], list[dict[str, Any]], list[int]]:
    cleaned = original.copy(deep=True)
    messages: list[str] = []

    if options.remove_duplicates:
        removed = int(cleaned.duplicated().sum())
        cleaned = cleaned.drop_duplicates()
        messages.append(f"{removed} təkrar sətir silindi")

    if options.missing_values == "zero":
        filled = int(cleaned.isna().sum().sum())
        cleaned = cleaned.fillna(0)
        messages.append(f"{filled} boş xana 0 ilə dolduruldu")
    elif options.missing_values == "mean":
        missing_before = int(cleaned.isna().sum().sum())
        numeric_columns = cleaned.select_dtypes(include="number").columns
        cleaned[numeric_columns] = cleaned[numeric_columns].fillna(
            cleaned[numeric_columns].mean()
        )
        filled = missing_before - int(cleaned.isna().sum().sum())
        remaining = int(cleaned.isna().sum().sum())
        messages.append(
            f"{filled} boş xana sütun ortalaması ilə dolduruldu; "
            f"{remaining} qeyri-rəqəm və ya ortalaması olmayan boş xana saxlanıldı"
        )
    elif options.missing_values == "drop_rows":
        removed = len(cleaned) - len(cleaned.dropna())
        cleaned = cleaned.dropna()
        messages.append(f"{removed} boş xana olan sətir silindi")

    if options.standardize_text:
        changed = 0
        for column in cleaned.select_dtypes(include=["object", "string"]).columns:
            if options.standardize_dates and date_column(cleaned, column):
                continue
            values = cleaned[column]
            standardized = values.map(
                lambda value: standardize_az_text(value)
                if isinstance(value, str)
                else value
            )
            changed += int(
                sum(
                    isinstance(before, str)
                    and before != after
                    for before, after in zip(values, standardized)
                )
            )
            cleaned[column] = standardized
        messages.append(f"{changed} mətn xanası standartlaşdırıldı")

    if options.standardize_dates:
        changed = 0
        for column in cleaned.columns:
            if not date_column(cleaned, column):
                continue
            values = cleaned[column]
            parsed = parse_mixed_dates(values)
            standardized = parsed.dt.strftime("%Y-%m-%d")
            converted = values.notna() & parsed.notna()
            changed += int(
                sum(
                    not isinstance(before, str) or before != after
                    for before, after in zip(values[converted], standardized[converted])
                )
            )
            cleaned[column] = standardized.where(parsed.notna(), values)
        messages.append(f"{changed} tarix standartlaşdırıldı (YYYY-MM-DD)")

    source_positions = {index: position + 1 for position, index in enumerate(original.index)}
    preview_indexes = cleaned.head(100).index
    row_indices = [source_positions[index] for index in preview_indexes]
    changed_cells: list[dict[str, Any]] = []
    for row_index, dataframe_index in zip(row_indices, preview_indexes):
        before = original.loc[dataframe_index]
        after = cleaned.loc[dataframe_index]
        for column in cleaned.columns:
            before_value = before[column]
            after_value = after[column]
            before_missing = pd.isna(before_value)
            after_missing = pd.isna(after_value)
            if before_missing and after_missing:
                continue
            if before_missing != after_missing or before_value != after_value:
                changed_cells.append({"row": row_index, "column": str(column)})

    return cleaned, messages, changed_cells, row_indices


@app.get("/")
async def home() -> FileResponse:
    return FileResponse(PROJECT_DIR / "static" / "index.html")


def build_overview(
    selected_file_id: str | None = None,
    selected_sheet_name: str | None = None,
) -> dict[str, Any]:
    total_rows = 0
    total_columns = 0
    total_missing_cells = 0
    total_duplicate_rows = 0
    sheet_names: list[dict[str, Any]] = []
    selected: tuple[str, str, pd.DataFrame] | None = None

    for file_id, dataset in datasets.items():
        for sheet_name, dataframe in dataset["sheets"].items():
            sheet_names.append(
                {
                    "file_id": file_id,
                    "filename": dataset["filename"],
                    "sheet_name": sheet_name,
                    "columns": [str(column) for column in dataframe.columns],
                }
            )
            total_rows += len(dataframe.index)
            total_columns += len(dataframe.columns)
            total_missing_cells += int(dataframe.isna().sum().sum())
            total_duplicate_rows += int(dataframe.duplicated().sum())

            if (
                selected is None
                and (selected_file_id is None or file_id == selected_file_id)
                and (
                    selected_sheet_name is None
                    or sheet_name == selected_sheet_name
                )
            ):
                selected = (file_id, sheet_name, dataframe)

    if selected_file_id is not None and selected_file_id not in datasets:
        raise HTTPException(status_code=404, detail="File not found")

    if selected_file_id is not None and selected_sheet_name is not None:
        dataset = datasets[selected_file_id]
        if selected_sheet_name not in dataset["sheets"]:
            raise HTTPException(status_code=404, detail="Sheet not found")
    elif selected_sheet_name is not None and selected is None:
        raise HTTPException(status_code=404, detail="Sheet not found")

    selected_file = None
    selected_sheet = None
    columns: list[str] = []
    rows: list[dict[str, Any]] = []
    if selected is not None:
        file_id, selected_sheet, dataframe = selected
        selected_file = {
            "file_id": file_id,
            "filename": datasets[file_id]["filename"],
        }
        columns = [str(column) for column in dataframe.columns]
        rows = json_rows(dataframe)

    return {
        "total_files": len(datasets),
        "sheet_names": sheet_names,
        "total_rows": total_rows,
        "total_columns": total_columns,
        "total_missing_cells": total_missing_cells,
        "total_duplicate_rows": total_duplicate_rows,
        "selected_file": selected_file,
        "selected_sheet": selected_sheet,
        "columns": columns,
        "rows": rows,
    }


@app.post("/api/upload")
async def upload_files(files: list[UploadFile] = File(...)) -> dict[str, Any]:
    if not files:
        raise HTTPException(status_code=400, detail="Upload at least one file")

    parsed_files: list[tuple[str, str, bytes, dict[str, pd.DataFrame]]] = []
    for upload in files:
        filename = upload.filename or ""
        extension = Path(filename).suffix.lower()
        if extension not in {".xlsx", ".xls", ".csv"}:
            raise HTTPException(
                status_code=400,
                detail=f"Dəstəklənməyən fayl formatı: {filename or 'adsız fayl'}",
            )

        try:
            content = await upload.read()
            if extension == ".csv":
                frame = read_csv_with_fallback(content)
                sheets = {Path(filename).stem or "Sheet1": frame}
            else:
                excel_sheets = pd.read_excel(
                    BytesIO(content),
                    sheet_name=None,
                    engine="openpyxl" if extension == ".xlsx" else "xlrd",
                )
                sheets = {
                    str(sheet_name): dataframe
                    for sheet_name, dataframe in excel_sheets.items()
                }
        except Exception as error:
            logger.warning(
                "Could not read uploaded file %r",
                filename,
                exc_info=(type(error), error, error.__traceback__),
            )
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{filename or 'Fayl'} oxuna bilmədi. "
                    "Faylın formatını, kodlaşdırmasını və məzmununu yoxlayın."
                ),
            ) from error

        parsed_files.append((filename, extension, content, sheets))

    for filename, extension, content, sheets in parsed_files:
        datasets[str(uuid4())] = {
            "filename": filename,
            "file_extension": extension,
            "workbook_bytes": content if extension == ".xlsx" else None,
            "sheets": sheets,
            "previews": {},
            "history": {},
            "cleaning_log": {},
        }

    return build_overview()


@app.get("/api/overview")
async def get_overview(
    file_id: str | None = None,
    sheet_name: str | None = None,
) -> dict[str, Any]:
    overview = build_overview(file_id, sheet_name)
    if file_id is not None and sheet_name is not None:
        dataset, _ = get_dataset_sheet(file_id, sheet_name)
        overview["cleaning_log"] = dataset["cleaning_log"].get(sheet_name, [])
        overview["has_undo"] = bool(dataset["history"].get(sheet_name, []))
    return overview


def color_extraction_source(file_id: str) -> tuple[dict[str, Any], bytes]:
    dataset = datasets.get(file_id)
    if dataset is None:
        raise HTTPException(status_code=404, detail="Fayl tapılmadı.")
    if dataset.get("file_extension") != ".xlsx" or not dataset.get("workbook_bytes"):
        raise HTTPException(
            status_code=400,
            detail=(
                "Rənglərə görə ayırma yalnız formatlı .xlsx fayllarında mümkündür. "
                ".csv, .xls və birləşdirilmiş məlumat dəstləri dəstəklənmir."
            ),
        )
    return dataset, dataset["workbook_bytes"]


@app.post("/api/colors/preview")
async def preview_color_extraction(
    request: ColorExtractionRequest,
) -> dict[str, Any]:
    _, workbook_bytes = color_extraction_source(request.file_id)
    selected_colors = list(dict.fromkeys(request.colors))
    try:
        grouped_rows, _ = extract_color_rows(workbook_bytes, selected_colors)
    except Exception as error:
        logger.warning(
            "Could not scan workbook colors for file %s",
            request.file_id,
            exc_info=(type(error), error, error.__traceback__),
        )
        raise HTTPException(
            status_code=400,
            detail="Excel faylının rəng formatlarını oxumaq mümkün olmadı.",
        ) from error
    return {
        "file_id": request.file_id,
        "colors": selected_colors,
        "scanned_sheets": len(datasets[request.file_id]["sheets"]),
        "groups": color_extraction_summary(grouped_rows),
    }


@app.post("/api/colors/export")
async def export_color_extraction(
    request: ColorExtractionRequest,
) -> Response:
    _, workbook_bytes = color_extraction_source(request.file_id)
    selected_colors = list(dict.fromkeys(request.colors))
    try:
        grouped_rows, all_headers = extract_color_rows(
            workbook_bytes,
            selected_colors,
        )
        export_bytes = color_extraction_workbook(grouped_rows, all_headers)
    except Exception as error:
        logger.warning(
            "Could not export workbook color groups for file %s",
            request.file_id,
            exc_info=(type(error), error, error.__traceback__),
        )
        raise HTTPException(
            status_code=400,
            detail="Rənglərə görə ayrılmış Excel faylını yaratmaq mümkün olmadı.",
        ) from error

    return Response(
        content=export_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": attachment_header(
                "rənglər_üzrə_ayrılmış.xlsx"
            )
        },
    )


@app.post("/api/clean/preview")
async def preview_cleaning(request: CleaningRequest) -> dict[str, Any]:
    dataset, original = get_dataset_sheet(request.file_id, request.sheet_name)
    cleaned, messages, changed_cells, row_indices = clean_dataframe(original, request)
    key = request.sheet_name
    dataset["previews"][key] = {
        "dataframe": cleaned,
        "messages": messages,
        "has_changes": not original.equals(cleaned),
    }

    return {
        "columns": [str(column) for column in original.columns],
        "original_rows": json_rows(original),
        "rows": json_rows(cleaned),
        "row_indices": row_indices,
        "changed_cells": changed_cells,
        "preview_log": messages,
        "cleaning_log": dataset["cleaning_log"].get(key, []),
        "has_changes": not original.equals(cleaned),
    }


@app.post("/api/clean/apply")
async def apply_cleaning(request: CleaningTarget) -> dict[str, Any]:
    dataset, original = get_dataset_sheet(request.file_id, request.sheet_name)
    preview = dataset["previews"].get(request.sheet_name)
    if preview is None:
        raise HTTPException(status_code=400, detail="Cleaning preview not found")
    if not preview["has_changes"]:
        raise HTTPException(status_code=400, detail="Preview has no changes to apply")

    dataset["history"].setdefault(request.sheet_name, []).append(original.copy(deep=True))
    dataset["sheets"][request.sheet_name] = preview["dataframe"]
    cleaning_log = dataset["cleaning_log"].setdefault(request.sheet_name, [])
    cleaning_log.extend(
        {"action": "Tətbiq edildi", "message": message}
        for message in preview["messages"]
    )
    dataset["previews"].pop(request.sheet_name, None)

    return {
        "overview": build_overview(request.file_id, request.sheet_name),
        "cleaning_log": cleaning_log,
        "has_undo": True,
    }


@app.post("/api/clean/undo")
async def undo_cleaning(request: CleaningTarget) -> dict[str, Any]:
    dataset, _ = get_dataset_sheet(request.file_id, request.sheet_name)
    history = dataset["history"].get(request.sheet_name, [])
    if not history:
        raise HTTPException(status_code=400, detail="No applied cleaning changes to undo")

    dataset["sheets"][request.sheet_name] = history.pop()
    dataset["previews"].pop(request.sheet_name, None)
    cleaning_log = dataset["cleaning_log"].setdefault(request.sheet_name, [])
    cleaning_log.append(
        {"action": "Geri qaytarıldı", "message": "Son tətbiq edilmiş dəyişiklik ləğv edildi"}
    )
    return {
        "overview": build_overview(request.file_id, request.sheet_name),
        "cleaning_log": cleaning_log,
        "has_undo": bool(history),
    }


@app.post("/api/merge")
async def merge_sheets(request: MergeRequest) -> dict[str, Any]:
    frames: list[pd.DataFrame] = []
    selected_sources: set[tuple[str, str]] = set()

    for source in request.sources:
        source_key = (source.file_id, source.sheet_name)
        if source_key in selected_sources:
            raise HTTPException(
                status_code=400,
                detail=f"Sheet selected more than once: {source.sheet_name}",
            )
        selected_sources.add(source_key)

        _, dataframe = get_dataset_sheet(source.file_id, source.sheet_name)
        column_names = {str(column) for column in dataframe.columns}
        unknown_columns = set(source.column_mapping) - column_names
        if unknown_columns:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"Unknown columns for {source.sheet_name}: "
                    f"{', '.join(map(str, sorted(unknown_columns)))}"
                ),
            )

        mapping = {
            column: source.column_mapping.get(str(column), str(column))
            for column in dataframe.columns
        }
        empty_targets = [column for column, target in mapping.items() if not target.strip()]
        if empty_targets:
            raise HTTPException(
                status_code=400,
                detail="Mapped column names cannot be empty",
            )
        if len(set(mapping.values())) != len(mapping):
            raise HTTPException(
                status_code=400,
                detail=f"Each source column must map to a unique column in {source.sheet_name}",
            )
        frames.append(dataframe.rename(columns=mapping).copy(deep=True))

    merged = pd.concat(frames, ignore_index=True, sort=False)
    merged_file_id = str(uuid4())
    merged_sheet_name = "Birləşdirilmiş"
    files_count = len({source.file_id for source in request.sources})
    log_entry = {
        "action": "Birləşdirmə",
        "message": (
            f"{files_count} fayl və {len(request.sources)} vərəq birləşdirildi "
            f"({len(merged)} sətir)"
        ),
    }
    datasets[merged_file_id] = {
        "filename": f"Birləşdirilmiş məlumat ({files_count} fayl)",
        "sheets": {merged_sheet_name: merged},
        "previews": {},
        "history": {},
        "cleaning_log": {merged_sheet_name: [log_entry]},
    }

    return {
        "overview": build_overview(merged_file_id, merged_sheet_name),
        "cleaning_log": [log_entry],
        "merged_files": files_count,
        "merged_sheets": len(request.sources),
    }


@app.post("/api/validate")
async def validate_sheet(request: ValidationTarget) -> dict[str, Any]:
    _, dataframe = get_dataset_sheet(request.file_id, request.sheet_name)
    row_reasons: dict[int, set[str]] = {}
    flagged_cells: list[dict[str, Any]] = []
    negative_cells = 0
    outlier_cells = 0
    numeric_columns = dataframe.select_dtypes(include="number").columns

    for column in numeric_columns:
        values = pd.to_numeric(dataframe[column], errors="coerce").astype("float64")
        finite = values.notna() & values.ne(float("inf")) & values.ne(float("-inf"))
        finite_values = values[finite]
        negative_rows = finite_values[finite_values < 0].index
        outlier_rows: pd.Index = pd.Index([])

        if not finite_values.empty:
            first_quartile = finite_values.quantile(0.25)
            third_quartile = finite_values.quantile(0.75)
            interquartile_range = third_quartile - first_quartile
            lower_fence = first_quartile - 1.5 * interquartile_range
            upper_fence = third_quartile + 1.5 * interquartile_range
            outlier_rows = finite_values[
                (finite_values < lower_fence) | (finite_values > upper_fence)
            ].index

        negative_indexes = set(negative_rows)
        outlier_indexes = set(outlier_rows)
        negative_cells += len(negative_indexes)
        outlier_cells += len(outlier_indexes)

        for index in negative_indexes | outlier_indexes:
            reasons: list[str] = []
            if index in negative_indexes:
                reasons.append("Mənfi dəyər")
            if index in outlier_indexes:
                reasons.append("IQR kənar dəyəri")
            row_number = dataframe.index.get_loc(index)
            row_reasons.setdefault(row_number, set()).update(reasons)
            flagged_cells.append(
                {
                    "row": row_number + 1,
                    "column": str(column),
                    "value": float(values.loc[index]),
                    "reasons": reasons,
                }
            )

    flagged_rows = [
        {
            "row": row_number + 1,
            "reasons": sorted(reasons),
        }
        for row_number, reasons in sorted(row_reasons.items())
    ]
    return {
        "numeric_columns": [str(column) for column in numeric_columns],
        "flagged_rows": flagged_rows,
        "flagged_cells": flagged_cells,
        "summary": {
            "flagged_rows": len(flagged_rows),
            "negative_values": negative_cells,
            "outliers": outlier_cells,
            "scanned_rows": len(dataframe),
        },
    }


@app.get("/api/export/{export_format}")
async def export_data(
    export_format: str,
    file_id: str,
    sheet_name: str,
    dialect: str = Query(default="sqlite", pattern="^(sqlite|oracle)$"),
) -> Response:
    dataset, dataframe = get_dataset_sheet(file_id, sheet_name)
    if dataframe.columns.empty:
        raise HTTPException(status_code=400, detail="Export üçün sütun tapılmadı")

    if export_format == "xlsx":
        buffer = BytesIO()
        try:
            worksheet_name = re.sub(r"[\[\]:*?/\\]", "_", sheet_name)[:31] or "Data"
            excel_dataframe = dataframe.copy()
            for column in excel_dataframe.select_dtypes(
                include=["datetimetz"]
            ).columns:
                excel_dataframe[column] = excel_dataframe[column].dt.tz_localize(None)
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                excel_dataframe.to_excel(writer, sheet_name=worksheet_name, index=False)
                worksheet = writer.sheets[worksheet_name]
                for column_index, column in enumerate(excel_dataframe.columns, start=1):
                    values = excel_dataframe[column].tolist()
                    max_length = max(
                        [len(str(column))]
                        + [len(str(value)) for value in values if not pd.isna(value)]
                    )
                    worksheet.column_dimensions[
                        worksheet.cell(row=1, column=column_index).column_letter
                    ].width = min(max(max_length + 2, 10), 60)
                for row_index, row in enumerate(
                    excel_dataframe.itertuples(index=False, name=None),
                    start=2,
                ):
                    for column_index, value in enumerate(row, start=1):
                        if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
                            worksheet.cell(
                                row=row_index,
                                column=column_index,
                            ).data_type = "s"
            content = buffer.getvalue()
        except Exception as error:
            logger.error(
                "Could not create Excel export",
                exc_info=(type(error), error, error.__traceback__),
            )
            raise HTTPException(
                status_code=500,
                detail="Excel faylını yaratmaq mümkün olmadı",
            ) from error
        filename = export_filename(file_id, "xlsx")
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif export_format == "sql":
        content = sql_script(dataframe, sheet_name, dialect).encode("utf-8")
        filename = export_filename(file_id, "sql")
        media_type = "application/sql; charset=utf-8"
    elif export_format == "txt":
        log_entries = dataset["cleaning_log"].get(sheet_name, [])
        report_lines = [
            "Təmizləmə Hesabatı",
            f"Fayl: {dataset['filename']}",
            f"Vərəq: {sheet_name}",
            "",
        ]
        if log_entries:
            report_lines.extend(
                f"{entry['action']}: {entry['message']}" for entry in log_entries
            )
        else:
            report_lines.append("Bu vərəq üçün təmizləmə əməliyyatı qeydə alınmayıb.")
        content = ("\n".join(report_lines) + "\n").encode("utf-8")
        filename = export_filename(file_id, "txt")
        media_type = "text/plain; charset=utf-8"
    else:
        raise HTTPException(
            status_code=400,
            detail="İxrac formatı xlsx, sql və ya txt olmalıdır",
        )

    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": attachment_header(filename)},
    )
