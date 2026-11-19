"""Engineering measurement analytics, charts, and Excel export helpers."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import Iterable

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot import as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

CHART_TYPES = [
    "Multi-Line Trend Chart", "Vendor Comparison Chart", "Average Trend Chart",
    "Box Plot by Vendor", "Histogram", "Scatter Plot", "Control/SPC Chart",
]


def numeric_columns(data: pd.DataFrame) -> list[str]:
    result = []
    for column in data.columns:
        values = pd.to_numeric(data[column], errors="coerce")
        if values.notna().any():
            result.append(str(column))
    return result


def _measurement_values(data: pd.DataFrame, columns: Iterable[str]) -> np.ndarray:
    columns = list(columns)
    if not columns:
        raise ValueError("Select at least one measurement column.")
    missing = [c for c in columns if c not in data.columns]
    if missing:
        raise ValueError("Measurement column(s) not found: " + ", ".join(missing))
    numeric = data[columns].apply(pd.to_numeric, errors="coerce")
    return numeric.to_numpy(dtype=float)


def _capability(values: np.ndarray, usl=None, lsl=None) -> dict:
    finite = np.asarray(values, dtype=float)
    finite = finite[np.isfinite(finite)]
    count = int(finite.size)
    mean = float(np.mean(finite)) if count else np.nan
    std = float(np.std(finite, ddof=1)) if count > 1 else np.nan

    # Moving-range estimate of within-subgroup variation for Cp/Cpk.
    within = np.nan
    if count > 1:
        differences = np.diff(finite)
        if differences.size:
            within = float(np.std(differences, ddof=1) / np.sqrt(2)) if differences.size > 1 else float(abs(differences[0]) / np.sqrt(2))

    above = int(np.sum(finite > usl)) if usl is not None else 0
    below = int(np.sum(finite < lsl)) if lsl is not None else 0

    if usl is not None and lsl is not None:
        fail = int(np.sum((finite > usl) | (finite < lsl)))
    elif usl is not None:
        fail = int(np.sum(finite > usl))
    elif lsl is not None:
        fail = int(np.sum(finite < lsl))
    else:
        fail = 0

    has_specs = usl is not None or lsl is not None
    yield_pct = (count - fail) / count * 100 if count and has_specs else np.nan

    cp = pp = cpk = ppk = np.nan
    if usl is not None and lsl is not None:
        spec_width = float(usl - lsl)
        if np.isfinite(within) and within > 0:
            cp = spec_width / (6 * within)
            cpk = min((usl - mean) / (3 * within), (mean - lsl) / (3 * within))
        if np.isfinite(std) and std > 0:
            pp = spec_width / (6 * std)
            ppk = min((usl - mean) / (3 * std), (mean - lsl) / (3 * std))

    if count:
        q1, q3 = np.percentile(finite, [25, 75])
        spread = q3 - q1
        outliers = int(np.sum((finite < q1 - 1.5 * spread) |
                              (finite > q3 + 1.5 * spread)))
    else:
        outliers = 0

    return {
        "Count": count,
        "Mean": mean,
        "Median": float(np.median(finite)) if count else np.nan,
        "Std Dev": std,
        "Min": float(np.min(finite)) if count else np.nan,
        "Max": float(np.max(finite)) if count else np.nan,
        "Range": float(np.ptp(finite)) if count else np.nan,
        "Cp": cp,
        "Cpk": cpk,
        "Pp": pp,
        "Ppk": ppk,
        "Above USL": above,
        "Below LSL": below,
        "Pass": count - fail,
        "Fail": fail,
        "Yield %": yield_pct,
        "Outlier Count": outliers,
    }


def analyze_measurements(data: pd.DataFrame, columns: Iterable[str],
                         usl=None, lsl=None) -> tuple[pd.DataFrame, dict]:
    columns = list(columns)
    values = _measurement_values(data, columns)
    overall = _capability(values.ravel(), usl, lsl)
    per_point = pd.DataFrame([
        {"Measurement Point": str(column),
         **_capability(values[:, index], usl, lsl)}
        for index, column in enumerate(columns)
    ])
    return per_point, overall


def vendor_summary(data: pd.DataFrame, columns: Iterable[str],
                   group_column: str, usl=None, lsl=None) -> pd.DataFrame:
    if not group_column or group_column not in data.columns:
        return pd.DataFrame()
    columns = list(columns)
    matrix = data[columns].apply(pd.to_numeric, errors="coerce")
    records = []
    for group, indexes in data.groupby(group_column, dropna=False, sort=False).groups.items():
        values = matrix.loc[indexes].to_numpy(dtype=float)
        stats = _capability(values.ravel(), usl, lsl)
        records.append({
            "Vendor": "(blank)" if pd.isna(group) else str(group),
            "Sample Count": stats["Count"],
            "Average": stats["Mean"],
            "Minimum": stats["Min"],
            "Maximum": stats["Max"],
            "Std Dev": stats["Std Dev"],
            "Cpk": stats["Cpk"],
            "Yield %": stats["Yield %"],
            "Out-of-spec Count": stats["Fail"],
        })

    result = pd.DataFrame(records)
    if not result.empty:
        result = result.sort_values(
            ["Yield %", "Cpk", "Std Dev"],
            ascending=[False, False, True], na_position="last",
        ).reset_index(drop=True)
        result.insert(0, "Rank", np.arange(1, len(result) + 1))
    return result


def dashboard_findings(data: pd.DataFrame, columns: Iterable[str],
                       group_column: str, usl=None, lsl=None) -> pd.DataFrame:
    columns = list(columns)
    values = data[columns].apply(pd.to_numeric, errors="coerce")
    failed = pd.DataFrame(False, index=data.index, columns=columns)
    if usl is not None:
        failed |= values.gt(usl)
    if lsl is not None:
        failed |= values.lt(lsl)

    rows = [{"Finding": "Top Fail Measurement Point", "Category": column,
             "Count": int(failed[column].sum())}
            for column in columns if failed[column].any()]

    if failed.to_numpy().any():
        unit_fail = failed.any(axis=1)
        if group_column and group_column in data.columns:
            common = data.loc[unit_fail, group_column].fillna("(blank)").astype(str).value_counts()
            if not common.empty:
                rows.append({"Finding": "Top Fail Vendor", "Category": common.index[0],
                             "Count": int(common.iloc[0])})
        for candidate, label in (("Build Matrix", "Top Fail BM"), ("BM", "Top Fail BM"),
                                 ("Station", "Top Fail Station")):
            if candidate in data.columns:
                common = data.loc[unit_fail, candidate].fillna("(blank)").astype(str).value_counts()
                if not common.empty:
                    rows.append({"Finding": label, "Category": common.index[0],
                                 "Count": int(common.iloc[0])})
    return pd.DataFrame(rows, columns=["Finding", "Category", "Count"])


def render_chart(data: pd.DataFrame, columns: Iterable[str], chart_type: str,
                 *, group_column="", id_column="", vendor_column="", bm_column="",
                 usl=None, lsl=None, target=None, highlight_oos=True,
                 selected_vendor="", sample_limit=500, figure=None):
    columns = list(columns)
    values = data[columns].apply(pd.to_numeric, errors="coerce")
    fig = figure or plt.figure(figsize=(11, 6), constrained_layout=True)
    fig.clear()
    ax = fig.add_subplot(111)

    group = (data[group_column].fillna("(blank)").astype(str)
             if group_column in data.columns else pd.Series("All", index=data.index))
    failed = pd.Series(False, index=data.index)
    if usl is not None:
        failed |= values.gt(usl).any(axis=1)
    if lsl is not None:
        failed |= values.lt(lsl).any(axis=1)

    if chart_type == "Multi-Line Trend Chart":
        n = len(data)
        if n > sample_limit:
            indexes = np.linspace(0, n - 1, sample_limit, dtype=int)
            plot_values = values.iloc[indexes]
        else:
            indexes = np.arange(n)
            plot_values = values
        for row_idx, (_, row) in zip(indexes, plot_values.iterrows()):
            is_fail = bool(failed.iloc[row_idx])
            is_vendor = bool(selected_vendor and group.iloc[row_idx] == selected_vendor)
            color = "#d62728" if highlight_oos and is_fail else ("#2878b5" if is_vendor else "#9aa0a6")
            alpha = 0.9 if is_fail or is_vendor else 0.16
            ax.plot(np.arange(len(columns)), row.to_numpy(dtype=float), color=color,
                    alpha=alpha, linewidth=1.6 if is_fail else 0.65)
        ax.set_xticks(np.arange(len(columns)), columns, rotation=45, ha="right")
        ax.set_xlabel("Measurement point"); ax.set_ylabel("Measurement value")

    elif chart_type in ("Vendor Comparison Chart", "Average Trend Chart"):
        if chart_type == "Vendor Comparison Chart" and group_column in data.columns:
            for name, ix in data.groupby(group_column, dropna=False, sort=False).groups.items():
                ax.plot(columns, values.loc[ix].mean(axis=0), marker="o", label=str(name))
            ax.legend(title=group_column, bbox_to_anchor=(1.02, 1), loc="upper left")
        else:
            ax.plot(columns, values.mean(axis=0), marker="o", color="#2878b5", label="Mean")
            ax.legend()
        ax.set_ylabel("Measurement value"); ax.tick_params(axis="x", rotation=45)

    elif chart_type == "Box Plot by Vendor":
        if group_column not in data.columns:
            raise ValueError("Select a grouping column for vendor box plots.")
        long = values.assign(__group=group).melt(id_vars="__group", var_name="Measurement", value_name="Value").dropna()
        vendors = long["__group"].astype(str).unique().tolist()[:40]
        groups = [long.loc[long["__group"].eq(vendor), "Value"].to_numpy() for vendor in vendors]
        ax.boxplot(groups, tick_labels=vendors, showfliers=False)
        ax.set_ylabel("Measurement value"); ax.tick_params(axis="x", rotation=45)

    elif chart_type == "Histogram":
        ax.hist(values.to_numpy().ravel(), bins="fd", color="#4472c4", alpha=.8)
        ax.set_xlabel("Measurement value"); ax.set_ylabel("Frequency")

    elif chart_type == "Scatter Plot":
        if len(columns) < 2:
            raise ValueError("Select at least two measurement columns for scatter plot.")
        ax.scatter(values.iloc[:, 0], values.iloc[:, 1], s=10, alpha=.5,
                   c=np.where(failed, "#d62728", "#4472c4"))
        ax.set_xlabel(columns[0]); ax.set_ylabel(columns[1])

    elif chart_type == "Control/SPC Chart":
        series = values.iloc[:, 0].to_numpy(dtype=float)
        if len(series) > sample_limit:
            indexes = np.linspace(0, len(series) - 1, sample_limit, dtype=int)
            series = series[indexes]
        else:
            indexes = np.arange(len(series))
        mean = float(np.nanmean(series)) if np.isfinite(series).any() else np.nan
        sigma = float(np.nanstd(series, ddof=1)) if np.isfinite(series).sum() > 1 else np.nan
        ax.plot(indexes, series, color="#4472c4", linewidth=.8, marker=".", markersize=3)
        ax.axhline(mean, color="green", label="Center line")
        if np.isfinite(sigma):
            ax.axhline(mean + 3 * sigma, color="#d62728", linestyle="--", label="UCL (3σ)")
            ax.axhline(mean - 3 * sigma, color="#d62728", linestyle="--", label="LCL (3σ)")
        ax.set_xlabel("Unit / row"); ax.set_ylabel(columns[0]); ax.legend()

    else:
        raise ValueError(f"Unsupported chart type: {chart_type}")

    if chart_type != "Control/SPC Chart" and (usl is not None or lsl is not None or target is not None):
        if usl is not None:
            ax.axhline(usl, color="blue", linestyle="--", label="USL")
        if target is not None:
            ax.axhline(target, color="green", linestyle="-", label="Target")
        if lsl is not None:
            ax.axhline(lsl, color="red", linestyle="--", label="LSL")
        ax.legend(loc="best")

    ax.set_title(chart_type)
    ax.grid(alpha=.2)
    return fig


def export_engineering_charts(output_path: str | Path, figure, stats: pd.DataFrame,
                              overall: dict, vendors: pd.DataFrame,
                              spec_analysis: pd.DataFrame,
                              spc_metrics: pd.DataFrame | None = None) -> Path:
    output_path = Path(output_path)
    if output_path.suffix.lower() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        workbook = load_workbook(output_path)
    else:
        from openpyxl import Workbook
        workbook = Workbook()
        workbook.remove(workbook.active)

    name = "6_EngineeringCharts"
    if name in workbook.sheetnames:
        del workbook[name]
    sheet = workbook.create_sheet(name)
    sheet["A1"] = "Engineering Chart Analyzer"
    sheet["A3"] = "Overall statistics"

    row = 4
    for key, value in overall.items():
        sheet.cell(row, 1, key)
        sheet.cell(row, 2, None if pd.isna(value) else value)
        row += 1

    image_buffer = BytesIO()
    figure.savefig(image_buffer, format="png", dpi=140, bbox_inches="tight")
    image_buffer.seek(0)
    image = XLImage(image_buffer)
    image.width = 960
    image.height = 520
    image.anchor = "D3"
    sheet.add_image(image)

    def add_table(title, frame, start_row, start_col=1):
        sheet.cell(start_row, start_col, title)
        if frame is None or frame.empty:
            sheet.cell(start_row + 1, start_col, "No data")
            return
        for column_index, column in enumerate(frame.columns, start_col):
            sheet.cell(start_row + 1, column_index, str(column))
        for row_index, values in enumerate(frame.itertuples(index=False, name=None), start_row + 2):
            for column_index, value in enumerate(values, start_col):
                if pd.isna(value):
                    value = None
                elif isinstance(value, np.generic):
                    value = value.item()
                sheet.cell(row_index, column_index, value)

    add_table("Measurement statistics", stats, 20)
    add_table("Vendor summary", vendors, 20, max(8, len(stats.columns) + 2))
    add_table("Spec analysis", spec_analysis, 40)
    add_table("SPC metrics", spc_metrics if spc_metrics is not None else pd.DataFrame(), 40, 8)

    for column in sheet.columns:
        letter = column[0].column_letter
        width = min(32, max(14, max(len(str(cell.value or "")) for cell in column) + 2))
        sheet.column_dimensions[letter].width = width

    workbook.save(output_path)
    return output_path
