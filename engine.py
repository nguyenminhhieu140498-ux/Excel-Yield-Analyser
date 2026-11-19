from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.drawing.image import Image as XLImage
from rapidfuzz import fuzz, process

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


class AnalyzerError(ValueError):
    """Input/configuration error for UI reporting."""


@dataclass
class MergeResult:
    data: pd.DataFrame
    total_yield_rows: int
    matched_yield_rows: int
    unmatched_yield_rows: int
    mapping_duplicate_keys: int
    match_rate: float


def list_sheets(path: str | Path) -> list[str]:
    path = Path(path)
    _validate_excel(path)
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception as exc:
        raise AnalyzerError(f"Could not read workbook: {exc}") from exc


def read_sheets(path: str | Path, sheets: Iterable[str]) -> pd.DataFrame:
    path = Path(path)
    _validate_excel(path)
    selected = list(sheets)
    if not selected:
        raise AnalyzerError("Select at least one worksheet.")
    try:
        available = list_sheets(path)
        missing = [sheet for sheet in selected if sheet not in available]
        if missing:
            raise AnalyzerError(f"Sheet not found: {', '.join(missing)}")
        frames = []
        for sheet in selected:
            frame = pd.read_excel(path, sheet_name=sheet)
            if frame.empty and len(frame.columns) == 0:
                continue
            frame.columns = [str(col).strip() for col in frame.columns]
            frame["__Source_Sheet__"] = sheet
            frames.append(frame)
        if not frames:
            raise AnalyzerError("Selected worksheets have no data.")
        return pd.concat(frames, ignore_index=True, sort=False)
    except AnalyzerError:
        raise
    except Exception as exc:
        raise AnalyzerError(f"Could not read selected sheets: {exc}") from exc


def _normalize_match_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def _nonblank_rows(data: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    result = data.copy()
    for column in columns:
        if column not in result.columns:
            continue
        result = result[
            result[column].notna()
            & result[column].astype(str).str.strip().ne("")
        ]
    return result


def merge_yield_data(yield_data, mapping_data, yield_key, mapping_key,
                     append_columns, how="left", fuzzy=False, fuzzy_threshold=90):
    if how not in {"left", "inner", "right"}:
        raise AnalyzerError("Join type must be left, inner, or right.")
    if yield_key not in yield_data.columns:
        raise AnalyzerError(f"Yield column not found: {yield_key}")
    if mapping_key not in mapping_data.columns:
        raise AnalyzerError(f"Mapping column not found: {mapping_key}")
    if not 0 <= fuzzy_threshold <= 100:
        raise AnalyzerError("Fuzzy threshold must be between 0 and 100.")

    selected = list(dict.fromkeys(append_columns))
    invalid = [c for c in selected if c not in mapping_data.columns]
    if invalid:
        raise AnalyzerError(f"Mapping column(s) not found: {', '.join(invalid)}")
    selected = [c for c in selected if c != mapping_key]

    left = yield_data.copy()
    right = mapping_data[[mapping_key, *selected]].copy()
    total = len(left)

    right["__eya_key_norm__"] = right[mapping_key].map(_normalize_match_key)
    right = right[right["__eya_key_norm__"] != ""].copy()
    duplicate_count = int(right["__eya_key_norm__"].duplicated(keep="first").sum())
    right = right.drop_duplicates(subset="__eya_key_norm__", keep="first")

    exact_lookup = set(right["__eya_key_norm__"].dropna().tolist())
    choices = right["__eya_key_norm__"].dropna().tolist()
    cache = {}

    def match_key(value):
        norm = _normalize_match_key(value)
        if not norm:
            return None, None
        if norm in cache:
            return cache[norm]
        if norm in exact_lookup:
            result = (norm, 100.0)
        elif fuzzy and choices:
            found = process.extractOne(norm, choices, scorer=fuzz.ratio,
                                       score_cutoff=fuzzy_threshold)
            result = (found[0], float(found[1])) if found else (None, None)
        else:
            result = (None, None)
        cache[norm] = result
        return result

    match_values = [match_key(value) for value in left[yield_key]]
    left["__eya_order__"] = np.arange(len(left), dtype=np.int64)
    left["__eya_match_key__"] = [item[0] for item in match_values]
    left["Fuzzy_Match_Score"] = [item[1] for item in match_values]
    right = right.rename(columns={"__eya_key_norm__": "__eya_match_key__"})

    try:
        merged = left.merge(
            right[["__eya_match_key__", *selected]],
            on="__eya_match_key__", how=how, sort=False,
            indicator="__eya_merge_status__", suffixes=("", "_mapping"),
            validate="many_to_one",
        )
    except Exception as exc:
        raise AnalyzerError(f"Merge failed: {exc}") from exc

    matched = int((merged["__eya_merge_status__"] == "both").sum())
    unmatched = max(0, total - matched)
    match_rate = matched / total * 100 if total else 0.0
    merged["__eya_sort_order__"] = merged["__eya_order__"].fillna(float("inf"))
    merged = merged.sort_values("__eya_sort_order__", kind="stable").reset_index(drop=True)
    merged = merged.drop(columns=["__eya_order__", "__eya_sort_order__",
                                  "__eya_match_key__", "__eya_merge_status__"],
                         errors="ignore")
    return MergeResult(merged, total, matched, unmatched, duplicate_count, match_rate)


def commonality(data: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    columns = list(dict.fromkeys(columns))
    if not columns:
        raise AnalyzerError("Select at least one commonality field.")
    missing = [c for c in columns if c not in data.columns]
    if missing:
        raise AnalyzerError(f"Field(s) not found: {', '.join(missing)}")
    filtered = _nonblank_rows(data, columns)
    if filtered.empty:
        return pd.DataFrame(columns=["Rank", *columns, "Count"])
    result = (
        filtered.groupby(columns, dropna=False, observed=True)
        .size().reset_index(name="Count")
        .sort_values("Count", ascending=False, kind="stable")
        .reset_index(drop=True)
    )
    result.insert(0, "Rank", np.arange(1, len(result) + 1))
    return result


def make_pivot(data: pd.DataFrame, rows: Iterable[str], columns: Iterable[str],
               value: str | None, aggregation: str = "count") -> pd.DataFrame:
    rows = list(rows)
    columns = list(columns)
    valid_aggs = {"count", "sum", "mean"}
    if aggregation not in valid_aggs:
        raise AnalyzerError("Aggregation must be count, sum, or mean.")

    for name in rows + columns:
        if name not in data.columns:
            raise AnalyzerError(f"Column not found: {name}")

    filtered = _nonblank_rows(data, rows + columns)
    if value is not None and value not in filtered.columns:
        raise AnalyzerError(f"Value column not found: {value}")

    try:
        if aggregation == "count":
            if rows:
                pivot = filtered.groupby(rows, dropna=False, observed=True).size()
                pivot = pivot.rename("Count")
                if len(rows) == 1:
                    frame = pivot.reset_index()
                else:
                    frame = pivot.reset_index()
                if columns:
                    grouped = filtered.groupby(rows + columns, dropna=False, observed=True).size()
                    grouped = grouped.rename("Count").reset_index()
                    if len(columns) == 1:
                        frame = grouped.pivot(index=rows,
                                              columns=columns[0],
                                              values="Count").fillna(0).reset_index()
                    else:
                        frame = grouped
                return _clean_pivot_columns(frame)

            if columns:
                pivot = filtered.groupby(columns, dropna=False, observed=True).size()
                frame = pivot.rename("Count").reset_index()
                return _clean_pivot_columns(frame)

            frame = pd.DataFrame({"Count": [len(filtered)]})
            return frame

        if value is None:
            raise AnalyzerError("Select a value column for sum or mean.")
        work = filtered.copy()
        work[value] = pd.to_numeric(work[value], errors="coerce")
        index = rows if rows else None
        pivot = pd.pivot_table(
            work, index=index, columns=columns if columns else None,
            values=value, aggfunc=aggregation, fill_value=0,
            dropna=False, observed=True,
        )
        if isinstance(pivot, pd.Series):
            pivot = pivot.to_frame()
        return _clean_pivot_columns(pivot.reset_index())
    except AnalyzerError:
        raise
    except Exception as exc:
        raise AnalyzerError(f"Could not generate pivot summary: {exc}") from exc


def _clean_pivot_columns(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.reset_index(drop=True)
    cleaned = []
    for col in frame.columns:
        if isinstance(col, tuple):
            cleaned.append(" / ".join(str(part) for part in col if str(part) != ""))
        else:
            cleaned.append(str(col))
    frame.columns = cleaned
    return frame


def commonality_matrix(data: pd.DataFrame, row_column: str, column_column: str) -> pd.DataFrame:
    if row_column not in data.columns:
        raise AnalyzerError(f"Row column not found: {row_column}")
    if column_column not in data.columns:
        raise AnalyzerError(f"Column column not found: {column_column}")
    filtered = _nonblank_rows(data, [row_column, column_column])
    if filtered.empty:
        return pd.DataFrame()
    return pd.crosstab(filtered[row_column].fillna("(blank)"),
                       filtered[column_column].fillna("(blank)"),
                       dropna=False)


def export_report(output_path: str | Path, merged: pd.DataFrame,
                  pivot: pd.DataFrame | None = None,
                  common: pd.DataFrame | None = None,
                  top_issues: pd.DataFrame | None = None,
                  matrix: pd.DataFrame | None = None) -> Path:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    pivot = pivot if pivot is not None else pd.DataFrame()
    common = common if common is not None else pd.DataFrame()
    top_issues = top_issues if top_issues is not None else pd.DataFrame()
    matrix = matrix if matrix is not None else pd.DataFrame()
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged.to_excel(writer, sheet_name="1_MergedData", index=False)
            pivot.to_excel(writer, sheet_name="2_PivotTable", index=False)
            common.to_excel(writer, sheet_name="3_Commonality", index=False)
            top_issues.to_excel(writer, sheet_name="4_TopIssues", index=False)
            matrix.to_excel(writer, sheet_name="5_Charts", index=True)
        chart_files = _create_charts(merged, top_issues, matrix, output_path.parent)
        workbook = load_workbook(output_path)
        chart_sheet = workbook["5_Charts"]
        if chart_files:
            for idx, image_path in enumerate(chart_files):
                img = XLImage(str(image_path)); img.width = 720; img.height = 400
                img.anchor = f"A{10 + idx * 22}"; chart_sheet.add_image(img)
        if not matrix.empty:
            col_count = matrix.shape[1]; row_count = matrix.shape[0]
            chart_sheet.conditional_formatting.add(
                f"B2:{_excel_col(col_count + 1)}{row_count + 2}",
                ColorScaleRule(start_type="min", start_color="63BE7B",
                               mid_type="percentile", mid_value=50,
                               mid_color="FFEB84", end_type="max",
                               end_color="F8696B"))
        workbook.save(output_path)
        return output_path
    except Exception as exc:
        raise AnalyzerError(f"Could not export report: {exc}") from exc


def _create_charts(merged, top_issues, matrix, directory):
    paths = []
    if not top_issues.empty and "Count" in top_issues.columns:
        label_col = next((c for c in top_issues.columns if c not in {"Rank", "Count"}), None)
        if label_col:
            top10 = top_issues.head(10).copy()
            if not top10.empty:
                fig, ax = plt.subplots(figsize=(10, 5))
                ax.barh(top10[label_col].astype(str), top10["Count"], color="#4472C4")
                ax.set_title("Top 10 Commonality"); ax.set_xlabel("Count")
                fig.tight_layout(); path = directory / "eya_top_issues.png"
                fig.savefig(path, dpi=130); plt.close(fig); paths.append(path)
    if not matrix.empty:
        fig, ax = plt.subplots(figsize=(10, 6))
        image = ax.imshow(matrix.to_numpy(), cmap="RdYlGn_r", aspect="auto")
        ax.set_title("Commonality Heatmap"); fig.colorbar(image, ax=ax, label="Count")
        fig.tight_layout(); path = directory / "eya_heatmap.png"
        fig.savefig(path, dpi=130); plt.close(fig); paths.append(path)
    return paths


def _validate_excel(path: Path) -> None:
    if not path.exists():
        raise AnalyzerError(f"File does not exist: {path}")
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise AnalyzerError("Choose a .xlsx or .xlsm workbook.")
    if path.stat().st_size == 0:
        raise AnalyzerError("The selected workbook is empty.")


def _excel_col(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result
