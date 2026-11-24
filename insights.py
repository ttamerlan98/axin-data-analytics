"""Local data-quality, statistical, reconciliation, and workbook insight helpers."""

from __future__ import annotations

import difflib
import math
import re
from datetime import date, datetime
from io import BytesIO
from typing import Any

import pandas as pd

try:
    import duckdb
except ImportError:  # Analytics still works if DuckDB's optional native wheel is unavailable.
    duckdb = None


def json_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        value = value.item()
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if pd.isna(value):
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _quoted(identifier: Any) -> str:
    return '"' + str(identifier).replace('"', '""') + '"'


def _records(dataframe: pd.DataFrame) -> list[dict[str, Any]]:
    return [
        {str(key): json_value(value) for key, value in row.items()}
        for row in dataframe.to_dict(orient="records")
    ]


def profile_dataframe(dataframe: pd.DataFrame) -> dict[str, Any]:
    row_count, column_count = dataframe.shape
    duplicate_count = int(dataframe.duplicated().sum()) if row_count else 0
    null_count = int(dataframe.isna().sum().sum())
    trimmed_count = 0
    columns: list[dict[str, Any]] = []

    for column in dataframe.columns:
        values = dataframe[column]
        missing = int(values.isna().sum())
        text_values = values.dropna().map(lambda value: value if isinstance(value, str) else None).dropna()
        trim_count = int(text_values.map(lambda value: value != value.strip()).sum())
        trimmed_count += trim_count
        lengths = text_values.map(len)
        columns.append(
            {
                "name": str(column),
                "dtype": str(values.dtype),
                "non_null": int(values.notna().sum()),
                "null_count": missing,
                "null_percent": round(missing / row_count * 100, 2) if row_count else 0.0,
                "distinct_count": int(values.nunique(dropna=True)),
                "trim_count": trim_count,
                "text_length_min": int(lengths.min()) if not lengths.empty else None,
                "text_length_max": int(lengths.max()) if not lengths.empty else None,
                "text_length_average": round(float(lengths.mean()), 2) if not lengths.empty else None,
                "samples": [json_value(value) for value in values.dropna().head(3).tolist()],
            }
        )

    total_cells = row_count * column_count
    missing_ratio = null_count / total_cells if total_cells else 0.0
    duplicate_ratio = duplicate_count / row_count if row_count else 0.0
    trim_ratio = trimmed_count / total_cells if total_cells else 0.0
    quality_score = max(
        0,
        round(100 - missing_ratio * 50 - duplicate_ratio * 30 - trim_ratio * 20),
    )
    return {
        "rows": row_count,
        "columns_count": column_count,
        "columns": columns,
        "null_count": null_count,
        "null_percent": round(null_count / total_cells * 100, 2) if total_cells else 0.0,
        "exact_duplicate_rows": duplicate_count,
        "trimmed_values": trimmed_count,
        "quality_score": quality_score,
        "quality_grade": "Yaxşı" if quality_score >= 85 else "Orta" if quality_score >= 65 else "Diqqət tələb edir",
    }


def _date_columns(dataframe: pd.DataFrame) -> list[str]:
    dates: list[str] = []
    for column in dataframe.columns:
        values = dataframe[column].dropna()
        name_suggests_date = bool(re.search(r"tarix|date|time|gün|ay", str(column), re.I))
        if pd.api.types.is_datetime64_any_dtype(dataframe[column]):
            dates.append(str(column))
        elif name_suggests_date and not values.empty:
            dates.append(str(column))
    return dates


def analytics_dataframe(
    dataframe: pd.DataFrame,
    group_by: str | None = None,
    metric: str | None = None,
    order_by: str | None = None,
) -> dict[str, Any]:
    if group_by is not None and group_by not in dataframe.columns:
        raise ValueError(f"Naməlum qruplaşdırma sütunu: {group_by}")
    numeric_columns = [
        str(column)
        for column in dataframe.select_dtypes(include="number").columns
    ]
    if metric is not None and metric not in dataframe.columns:
        raise ValueError(f"Naməlum ölçü sütunu: {metric}")
    if order_by is not None and order_by not in dataframe.columns:
        raise ValueError(f"Naməlum sıralama sütunu: {order_by}")
    if metric is None and numeric_columns:
        metric = numeric_columns[0]

    numeric_stats: list[dict[str, Any]] = []
    for column in numeric_columns:
        values = pd.to_numeric(dataframe[column], errors="coerce").dropna()
        if values.empty:
            continue
        numeric_stats.append(
            {
                "column": column,
                "count": int(values.count()),
                "sum": json_value(values.sum()),
                "average": json_value(values.mean()),
                "minimum": json_value(values.min()),
                "maximum": json_value(values.max()),
                "median": json_value(values.median()),
                "standard_deviation": json_value(values.std()) if len(values) > 1 else 0.0,
            }
        )

    date_stats: list[dict[str, Any]] = []
    for column in _date_columns(dataframe):
        parsed = pd.to_datetime(dataframe[column], errors="coerce").dropna()
        if parsed.empty:
            continue
        date_stats.append(
            {
                "column": column,
                "minimum": json_value(parsed.min()),
                "maximum": json_value(parsed.max()),
                "span_days": int((parsed.max() - parsed.min()).days),
                "valid_dates": int(parsed.count()),
            }
        )

    categorical: list[dict[str, Any]] = []
    for column in dataframe.select_dtypes(include=["object", "string", "category"]).columns:
        counts = dataframe[column].dropna().astype(str).value_counts().head(10)
        if not counts.empty:
            categorical.append(
                {
                    "column": str(column),
                    "distinct": int(dataframe[column].nunique(dropna=True)),
                    "top10": [
                        {"value": str(value), "count": int(count)}
                        for value, count in counts.items()
                    ],
                }
            )

    grouped: list[dict[str, Any]] = []
    if group_by is not None:
        grouped = _grouped_summary(dataframe, group_by, metric)

    top10: list[dict[str, Any]] = []
    windows: list[dict[str, Any]] = []
    if metric is not None and pd.api.types.is_numeric_dtype(dataframe[metric]):
        enriched = dataframe.copy()
        enriched["__datalens_row__"] = range(1, len(enriched) + 1)
        top10 = _records(
            enriched.sort_values(metric, ascending=False, na_position="last")
            .head(10)
            .drop(columns="__datalens_row__")
        )
        windows = _window_summary(enriched, metric, group_by, order_by)

    return {
        "rows": len(dataframe),
        "numeric_statistics": numeric_stats,
        "date_statistics": date_stats,
        "categorical_groups": categorical,
        "grouped_summary": grouped,
        "top10": top10,
        "window_analytics": windows,
        "window_metric": metric,
        "window_order": order_by or "__source_row__",
        "engine": "DuckDB" if duckdb is not None else "pandas",
    }


def _grouped_summary(
    dataframe: pd.DataFrame,
    group_by: str,
    metric: str | None,
) -> list[dict[str, Any]]:
    if duckdb is not None:
        try:
            connection = duckdb.connect(":memory:")
            connection.register("datalens_frame", dataframe)
            metric_sql = (
                f", SUM(TRY_CAST({_quoted(metric)} AS DOUBLE)) AS total, "
                f"AVG(TRY_CAST({_quoted(metric)} AS DOUBLE)) AS average"
                if metric is not None
                else ""
            )
            result = connection.execute(
                f"SELECT {_quoted(group_by)} AS group_value, COUNT(*) AS row_count"
                f"{metric_sql} FROM datalens_frame "
                f"GROUP BY {_quoted(group_by)} ORDER BY row_count DESC LIMIT 100"
            ).df()
            connection.close()
            return _records(result)
        except Exception:
            pass

    work = dataframe.copy()
    work["__group__"] = work[group_by].map(lambda value: "(boş)" if pd.isna(value) else str(value))
    if metric is None:
        result = work.groupby("__group__", dropna=False).size().rename("row_count").reset_index()
    else:
        result = work.groupby("__group__", dropna=False)[metric].agg(
            row_count="count", total="sum", average="mean"
        ).reset_index()
    result = result.rename(columns={"__group__": "group_value"}).sort_values("row_count", ascending=False).head(100)
    return _records(result)


def _window_summary(
    dataframe: pd.DataFrame,
    metric: str,
    group_by: str | None,
    order_by: str | None,
) -> list[dict[str, Any]]:
    if dataframe.empty:
        return []
    order = order_by or "__datalens_row__"
    if duckdb is not None:
        try:
            connection = duckdb.connect(":memory:")
            connection.register("datalens_frame", dataframe)
            partition = f"PARTITION BY {_quoted(group_by)} " if group_by else ""
            group_selection = f"{_quoted(group_by)} AS group_value, " if group_by else ""
            rows = connection.execute(
                f"SELECT {_quoted('__datalens_row__')} AS source_row, "
                f"{group_selection}{_quoted(metric)} AS value, "
                f"DENSE_RANK() OVER ({partition}ORDER BY {_quoted(metric)} DESC) AS dense_rank, "
                f"{_quoted(metric)} - LAG({_quoted(metric)}) OVER "
                f"({partition}ORDER BY {_quoted(order)}) AS difference, "
                f"SUM({_quoted(metric)}) OVER ({partition}ORDER BY {_quoted(order)} "
                "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS running_total, "
                f"LAG({_quoted(metric)}) OVER ({partition}ORDER BY {_quoted(order)}) AS lag "
                f"FROM datalens_frame WHERE {_quoted(metric)} IS NOT NULL "
                f"ORDER BY {_quoted('__datalens_row__')} LIMIT 1000"
            ).df()
            connection.close()
            return _records(rows)
        except Exception:
            pass

    ordered = dataframe.copy()
    ordered["__datalens_row__"] = range(1, len(ordered) + 1)
    if order_by:
        ordered = ordered.sort_values(order_by, kind="stable", na_position="last")
    groups = ordered.groupby(group_by, dropna=False, sort=False) if group_by else [(None, ordered)]
    output: list[dict[str, Any]] = []
    for group_value, part in groups:
        values = pd.to_numeric(part[metric], errors="coerce")
        previous = values.shift()
        dense_ranks = values.rank(method="dense", ascending=False)
        running = values.fillna(0).cumsum()
        for index, value in values.items():
            if pd.isna(value):
                continue
            row = {
                "source_row": int(part.loc[index, "__datalens_row__"]),
                "value": json_value(value),
                "dense_rank": json_value(dense_ranks.loc[index]),
                "difference": json_value(value - previous.loc[index]) if pd.notna(previous.loc[index]) else None,
                "running_total": json_value(running.loc[index]),
                "lag": json_value(previous.loc[index]),
            }
            if group_by:
                row["group_value"] = json_value(group_value)
            output.append(row)
    return output[:1000]


def reconcile_dataframes(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_key: str,
    right_key: str,
    left_amount: str,
    right_amount: str,
    tolerance: float = 0.01,
) -> dict[str, Any]:
    for dataframe, column in (
        (left, left_key), (right, right_key),
        (left, left_amount), (right, right_amount),
    ):
        if column not in dataframe.columns:
            raise ValueError(f"Naməlum sütun: {column}")

    def prepare(frame: pd.DataFrame, key: str, amount: str, prefix: str) -> pd.DataFrame:
        result = frame[[key, amount]].copy().reset_index(drop=True)
        result.columns = [f"{prefix}_key_value", f"{prefix}_amount_value"]
        result[f"{prefix}_row"] = range(1, len(result) + 1)
        result[f"{prefix}_key"] = result[f"{prefix}_key_value"].map(
            lambda value: None if pd.isna(value) or not str(value).strip() else str(value).strip().casefold()
        )
        result[f"{prefix}_join_key"] = result[f"{prefix}_key"]
        result[f"{prefix}_amount"] = pd.to_numeric(result[f"{prefix}_amount_value"], errors="coerce")
        result[f"{prefix}_occurrence"] = result.groupby(
            f"{prefix}_key", dropna=False
        ).cumcount()
        missing_mask = result[f"{prefix}_key"].isna()
        result.loc[missing_mask, f"{prefix}_join_key"] = [
            f"__missing_{prefix}_{row_number}"
            for row_number in result.loc[missing_mask, f"{prefix}_row"]
        ]
        return result

    left_prepared = prepare(left, left_key, left_amount, "left")
    right_prepared = prepare(right, right_key, right_amount, "right")
    merged = left_prepared.merge(
        right_prepared,
        left_on=["left_join_key", "left_occurrence"],
        right_on=["right_join_key", "right_occurrence"],
        how="outer",
        suffixes=("_left", "_right"),
        indicator=True,
    )
    results: list[dict[str, Any]] = []
    counts = {"MATCH": 0, "AMOUNT_DIFF": 0, "LEFT_ONLY": 0, "RIGHT_ONLY": 0}
    for item in merged.to_dict(orient="records"):
        left_present = item["_merge"] in {"both", "left_only"}
        right_present = item["_merge"] in {"both", "right_only"}
        left_value = item.get("left_amount")
        right_value = item.get("right_amount")
        if item["_merge"] == "left_only":
            status = "LEFT_ONLY"
        elif item["_merge"] == "right_only":
            status = "RIGHT_ONLY"
        elif left_value is None or pd.isna(left_value) or right_value is None or pd.isna(right_value):
            status = "AMOUNT_DIFF"
        else:
            status = "MATCH" if abs(float(left_value) - float(right_value)) <= tolerance else "AMOUNT_DIFF"
        counts[status] += 1
        results.append(
            {
                "key": json_value(
                    item.get("left_key_value") if left_present
                    else item.get("right_key_value")
                ),
                "left_row": json_value(item.get("left_row")),
                "right_row": json_value(item.get("right_row")),
                "left_amount": json_value(left_value),
                "right_amount": json_value(right_value),
                "difference": json_value(
                    float(left_value) - float(right_value)
                    if left_value is not None and right_value is not None
                    and not pd.isna(left_value) and not pd.isna(right_value)
                    else None
                ),
                "status": status,
            }
        )
    total_left = float(left_prepared["left_amount"].sum())
    total_right = float(right_prepared["right_amount"].sum())
    return {
        "rows": results,
        "counts": counts,
        "totals": {
            "left": total_left,
            "right": total_right,
            "net_difference": total_left - total_right,
        },
        "summary": {
            "left_rows": len(left),
            "right_rows": len(right),
            "matched": counts["MATCH"],
            "amount_differences": counts["AMOUNT_DIFF"],
            "left_only": counts["LEFT_ONLY"],
            "right_only": counts["RIGHT_ONLY"],
            "net_difference": total_left - total_right,
        },
        "tolerance": tolerance,
    }


def anomaly_report(dataframe: pd.DataFrame) -> dict[str, Any]:
    outliers: list[dict[str, Any]] = []
    for column in dataframe.select_dtypes(include="number").columns:
        values = pd.to_numeric(dataframe[column], errors="coerce")
        finite = values.dropna()
        if finite.empty:
            continue
        q1, q3 = finite.quantile(0.25), finite.quantile(0.75)
        spread = q3 - q1
        mask = (values < q1 - 1.5 * spread) | (values > q3 + 1.5 * spread)
        negative = values < 0
        selected = mask | negative
        for row_number, number in enumerate(values.tolist()):
            if not bool(selected.iloc[row_number]):
                continue
            reasons = []
            if number < 0:
                reasons.append("Mənfi dəyər")
            if bool(mask.iloc[row_number]):
                reasons.append("IQR kənar dəyəri")
            outliers.append(
                {"row": row_number + 1, "column": str(column), "value": json_value(number), "reasons": reasons}
            )

    fuzzy: list[dict[str, Any]] = []
    for column in dataframe.select_dtypes(include=["object", "string", "category"]).columns:
        counts = dataframe[column].dropna().astype(str).value_counts()
        values = counts.index[:250].tolist()
        pairs = 0
        for index, first in enumerate(values):
            for second in values[index + 1:]:
                score = difflib.SequenceMatcher(None, first.casefold(), second.casefold()).ratio()
                if score >= 0.88:
                    fuzzy.append(
                        {
                            "column": str(column),
                            "first": first,
                            "first_count": int(counts[first]),
                            "second": second,
                            "second_count": int(counts[second]),
                            "similarity": round(score, 3),
                        }
                    )
                    pairs += 1
                    if len(fuzzy) >= 100:
                        break
            if pairs and len(fuzzy) >= 100:
                break

    return {
        "exact_duplicate_rows": int(dataframe.duplicated().sum()),
        "numeric_anomalies": outliers,
        "fuzzy_text_matches": fuzzy,
        "summary": {
            "scanned_rows": len(dataframe),
            "numeric_anomalies": len(outliers),
            "fuzzy_text_matches": len(fuzzy),
            "exact_duplicate_rows": int(dataframe.duplicated().sum()),
        },
    }


def compile_audit_summary(
    filename: str,
    sheet_name: str,
    dataframe: pd.DataFrame,
) -> dict[str, Any]:
    profile = profile_dataframe(dataframe)
    anomalies = anomaly_report(dataframe)
    return {
        "filename": filename,
        "sheet_name": sheet_name,
        "rows": profile["rows"],
        "columns": profile["columns_count"],
        "quality_score": profile["quality_score"],
        "quality_grade": profile["quality_grade"],
        "exact_duplicate_rows": profile["exact_duplicate_rows"],
        "missing_values": profile["null_count"],
        "trimmed_values": profile["trimmed_values"],
        "numeric_anomalies": anomalies["summary"]["numeric_anomalies"],
        "fuzzy_text_matches": anomalies["summary"]["fuzzy_text_matches"],
        "column_health": [
            {
                "column": column["name"],
                "null_percent": column["null_percent"],
                "distinct": column["distinct_count"],
                "dtype": column["dtype"],
            }
            for column in profile["columns"]
        ],
    }


def audit_prompt(summary: dict[str, Any], anomalies: dict[str, Any]) -> str:
    return (
        "Audit etibarlı şəkildə və yalnız aşağıdakı məxfi olmayan, "
        "ümumiləşdirilmiş Excel göstəricilərini təhlil et. Müşahidələri, "
        "riskləri və yoxlanmalı addımları qısa maddələrlə yaz; uydurma nəticə çıxarma.\n\n"
        f"Xülasə: {summary}\n"
        f"Anomaliya göstəriciləri: {anomalies.get('summary', {})}\n"
        f"Nümunə anomaliyalar: {anomalies.get('numeric_anomalies', [])[:15]}\n"
        f"Oxşar mətnlər: {anomalies.get('fuzzy_text_matches', [])[:15]}"
    )
