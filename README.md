# Excel Yield Analyzer

A macOS desktop starter application for merging manufacturing yield workbooks,
adding mapping data, exploring commonality, creating pivot-style summaries,
and exporting analysis to Excel.

## Requirements

- macOS Ventura (13) or later
- Python 3.12
- Excel workbooks in `.xlsx` or `.xlsm` format

## Run from source

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python app.py
```

## Generate sample input files

```bash
python sample_data.py
```

Import `sample_data/Daily_Yield.xlsx`, select the FATP and CR sheets, then
import `sample_data/Mapping_List.xlsx`. Use `Serial_Number` and `SN` as the
merge keys, select mapping fields to append, and run the merge.

## Tests

```bash
pytest
```

## Build the macOS app

```bash
chmod +x build_mac.sh
./build_mac.sh
```

The app bundle is created under `dist/Excel Yield Analyzer.app`. For
distribution outside your development Mac, code signing and notarization are
recommended. Build separately on each target architecture (Apple Silicon or
Intel) unless you have verified a universal build.

## Merge behavior

- Supports left, inner, and right joins.
- Mapping records with duplicate keys are deduplicated, keeping the first
  record; the duplicate count is shown in merge statistics.
- Yield-row order is retained. Unmatched mapping rows in a right join appear
  after yield rows.
- Blank keys do not match other blank keys.
- Match rate is matched yield rows divided by input yield rows.

## Pivot and Excel limitations

The `2_PivotTable` worksheet is a generated pivot-style aggregate table, not a
native interactive Excel PivotTable. `openpyxl` does not provide a dependable
way to author Excel PivotTable objects. The export includes merged data,
pivot summary, commonality results, top commonality, and a charts sheet with
embedded charts and a heatmap-colored matrix.

## Performance and privacy

Workbooks are processed locally and are not uploaded. Data is held in memory;
large files need sufficient available RAM. Runtime depends on workbook size,
formats, and Mac hardware. The 100,000-row/15-second target should be
benchmarked against representative production workbooks before rollout.