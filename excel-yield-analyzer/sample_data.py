from pathlib import Path

import pandas as pd


def main():
    output = Path("sample_data")
    output.mkdir(exist_ok=True)

    yield_rows = pd.DataFrame({
        "Serial_Number": [f"SN{i:05d}" for i in range(1, 101)],
        "Issue_Code": ["Error_101", "Error_102", "Error_103", "PASS"] * 25,
        "Date": pd.date_range("2026-01-01", periods=100, freq="h"),
        "Factory": ["Factory_A", "Factory_B"] * 50,
        "Station": ["FATP_01", "FATP_02", "CR_01", "ORT_01"] * 25,
    })
    mapping = pd.DataFrame({
        "SN": [f"SN{i:05d}" for i in range(1, 96)] + ["SN00001"],
        "BM": [f"BM_{i % 5 + 1:03d}" for i in range(96)],
        "Product": ["Product_A", "Product_B", "Product_C"] * 32,
        "Line": ["SMT_Line_01", "SMT_Line_02"] * 48,
        "Issue_Category": ["Electrical", "Mechanical"] * 48,
    })

    with pd.ExcelWriter(output / "Daily_Yield.xlsx", engine="openpyxl") as writer:
        yield_rows.iloc[:50].to_excel(writer, sheet_name="FATP", index=False)
        yield_rows.iloc[50:].to_excel(writer, sheet_name="CR", index=False)
        pd.DataFrame({"Summary": ["Sample report"]}).to_excel(
            writer, sheet_name="Summary", index=False
        )
    mapping.to_excel(output / "Mapping_List.xlsx", index=False)
    print(f"Created sample workbooks in {output.resolve()}")


if __name__ == "__main__":
    main()