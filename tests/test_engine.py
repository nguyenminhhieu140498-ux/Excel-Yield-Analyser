import pandas as pd

from engine import commonality, make_pivot, merge_yield_data


def test_left_merge_and_duplicate_mapping_keys():
    yield_data = pd.DataFrame({"SN": ["a", "b", "c"], "Issue": ["x", "y", "x"]})
    mapping = pd.DataFrame({"Key": ["a", "a", "b"], "BM": ["one", "two", "three"]})

    result = merge_yield_data(
        yield_data, mapping, "SN", "Key", ["BM"], "left"
    )

    assert result.total_yield_rows == 3
    assert result.matched_yield_rows == 2
    assert result.unmatched_yield_rows == 1
    assert result.mapping_duplicate_keys == 1
    assert result.data["SN"].tolist() == ["a", "b", "c"]
    assert result.data.loc[0, "BM"] == "one"


def test_commonality_and_count_pivot():
    frame = pd.DataFrame({
        "BM": ["A", "A", "B"],
        "Issue": ["x", "x", "y"],
    })
    result = commonality(frame, ["Issue", "BM"])
    assert result.iloc[0]["Count"] == 2

    pivot = make_pivot(frame, ["BM"], [], None, "count")
    assert not pivot.empty