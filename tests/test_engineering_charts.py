import pytest
import pandas as pd

from engineering_charts import (
    analyze_measurements,
    dashboard_findings,
    numeric_columns,
    vendor_summary,
)


def test_numeric_column_detection_and_stats_with_specs():
    data = pd.DataFrame({
        "SN": ["a", "b", "c"],
        "A": [0.0, 0.1, -0.4],
        "B": [0.02, 0.03, 0.04],
    })

    assert numeric_columns(data) == ["A", "B"]
    table, overall = analyze_measurements(
        data,
        ["A", "B"],
        usl=0.07,
        lsl=-0.33,
    )

    assert overall["Count"] == 6
    assert overall["Above USL"] == 1
    assert overall["Below LSL"] == 1
    assert overall["Fail"] == 2
    assert overall["Pass"] == 4
    assert overall["Yield %"] == pytest.approx(100 * 4 / 6)
    assert len(table) == 2


def test_vendor_ranking_and_dashboard_findings():
    data = pd.DataFrame({
        "Vendor": ["A", "A", "B"],
        "Build Matrix": ["x", "y", "x"],
        "Station": ["S1", "S1", "S2"],
        "M1": [0.0, 0.2, -0.4],
    })

    ranked = vendor_summary(
        data,
        ["M1"],
        "Vendor",
        usl=0.1,
        lsl=-0.2,
    )
    assert ranked.iloc[0]["Vendor"] == "A"

    findings = dashboard_findings(
        data,
        ["M1"],
        "Vendor",
        usl=0.1,
        lsl=-0.2,
    )
    assert "Top Fail Measurement Point" in findings["Finding"].tolist()
    assert "Top Fail Vendor" in findings["Finding"].tolist()
