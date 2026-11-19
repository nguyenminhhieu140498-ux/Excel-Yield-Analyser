

## Engineering Chart Analyzer

A standalone desktop window for engineering trend/SPC visualization, spec analysis, vendor ranking, and an out-of-spec dashboard. Start it from the repository root:

```bash
python engineering_chart_analyzer.py
```

Open an `.xlsx` or `.xlsm` workbook, select a worksheet and one or more auto-detected numeric measurement columns, then choose one of the seven chart types. The interactive Matplotlib toolbar provides zoom and pan; chart controls include USL/LSL/target, vendor grouping, out-of-spec highlighting, and a bounded sample count for large trend plots. Save charts as PNG or PDF, or export the visualization and summary tables to the `6_EngineeringCharts` worksheet in an Excel workbook.

Capability indices (Cp/Cpk/Pp/Ppk) are only meaningful when both spec limits are entered and there is sufficient non-constant data. Cp/Cpk use an estimate of within variation from successive observations; Pp/Ppk use overall standard deviation. SPC chart limits are descriptive ±3 standard deviations and do not implement optional Western Electric rules. Plotting is sampled for large trend charts; calculations remain based on all selected data. Performance depends on workbook size and machine, so the 100k-row / five-second target should be benchmarked with representative production files before relying on it.
