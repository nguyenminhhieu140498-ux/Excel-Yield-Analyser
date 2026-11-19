"""Desktop Engineering Chart Analyzer window for Excel Yield Analyzer."""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from engineering_charts import (
    CHART_TYPES,
    analyze_measurements,
    dashboard_findings,
    export_engineering_charts,
    numeric_columns,
    render_chart,
    vendor_summary,
)


class EngineeringChartWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Engineering Chart Analyzer")
        self.resize(1400, 900)
        self.data = None
        self.workbook_path = None

        root = QVBoxLayout(self)

        file_row = QHBoxLayout()
        self.file_label = QLabel("Choose an Excel workbook to begin")
        choose = QPushButton("Open workbook…")
        choose.clicked.connect(self.open_workbook)
        file_row.addWidget(choose)
        file_row.addWidget(self.file_label, 1)
        root.addLayout(file_row)

        split = QSplitter(Qt.Orientation.Horizontal)
        controls = QWidget()
        form = QFormLayout(controls)

        self.sheet = QComboBox()
        self.sheet.currentTextChanged.connect(self.load_sheet)

        self.chart_type = QComboBox()
        self.chart_type.addItems(CHART_TYPES)
        self.chart_type.currentTextChanged.connect(self.update_chart)

        self.columns = QListWidget()
        self.columns.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.columns.itemSelectionChanged.connect(self.update_chart)

        self.id_column = QComboBox()
        self.vendor_column = QComboBox()
        self.bm_column = QComboBox()
        self.group_column = QComboBox()
        self.vendor_filter = QComboBox()

        for combo in (
            self.id_column,
            self.vendor_column,
            self.bm_column,
            self.group_column,
            self.vendor_filter,
        ):
            combo.addItem("")

        self.usl = QLineEdit()
        self.lsl = QLineEdit()
        self.target = QLineEdit()
        self.highlight = QCheckBox("Highlight out-of-spec units")
        self.highlight.setChecked(True)
        self.sample = QLineEdit("500")

        form.addRow("Worksheet", self.sheet)
        form.addRow("Chart type", self.chart_type)
        form.addRow("Measurement columns", self.columns)
        form.addRow("Serial / unit", self.id_column)
        form.addRow("Vendor", self.vendor_column)
        form.addRow("Build Matrix", self.bm_column)
        form.addRow("Group by", self.group_column)
        form.addRow("Selected vendor", self.vendor_filter)
        form.addRow("USL", self.usl)
        form.addRow("Target", self.target)
        form.addRow("LSL", self.lsl)
        form.addRow("Trend sample limit", self.sample)
        form.addRow(self.highlight)

        buttons = QHBoxLayout()
        analyze = QPushButton("Analyze / refresh chart")
        analyze.clicked.connect(self.update_chart)
        save_png = QPushButton("Save PNG…")
        save_png.clicked.connect(self.save_png)
        save_pdf = QPushButton("Save PDF…")
        save_pdf.clicked.connect(self.save_pdf)
        export = QPushButton("Export to Excel…")
        export.clicked.connect(self.export_excel)
        buttons.addWidget(analyze)
        buttons.addWidget(save_png)
        buttons.addWidget(save_pdf)
        buttons.addWidget(export)
        form.addRow(buttons)

        split.addWidget(controls)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvasQTAgg(self.figure)
        right_layout.addWidget(NavigationToolbar2QT(self.canvas, self))
        right_layout.addWidget(self.canvas, 4)
        right_layout.addWidget(QLabel("Engineering statistics and vendor ranking"))

        self.table = QTableWidget()
        self.table.setMaximumHeight(190)
        right_layout.addWidget(self.table)

        self.findings = QLabel("Yield dashboard findings will appear here.")
        self.findings.setWordWrap(True)
        right_layout.addWidget(self.findings)

        split.addWidget(right)
        split.setStretchFactor(1, 1)
        root.addWidget(split)

        self.status = QLabel(
            "Zoom and pan with the chart toolbar. Hover charts for navigation coordinates."
        )
        root.addWidget(self.status)

        for combo in (
            self.id_column,
            self.vendor_column,
            self.bm_column,
            self.group_column,
            self.vendor_filter,
        ):
            combo.currentTextChanged.connect(self.update_chart)

        for line in (self.usl, self.lsl, self.target, self.sample):
            line.editingFinished.connect(self.update_chart)

        self.highlight.stateChanged.connect(self.update_chart)

    def open_workbook(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Excel workbook",
            "",
            "Excel workbooks (*.xlsx *.xlsm)",
        )
        if not path:
            return

        try:
            self.workbook_path = Path(path)
            sheet_names = pd.ExcelFile(path).sheet_names
            self.sheet.blockSignals(True)
            self.sheet.clear()
            self.sheet.addItems(sheet_names)
            self.sheet.blockSignals(False)
            self.file_label.setText(self.workbook_path.name)
            self.load_sheet(self.sheet.currentText())
        except Exception as exc:
            QMessageBox.warning(self, "Engineering Chart Analyzer", str(exc))

    @staticmethod
    def _guess_column(names: list[str], keywords: tuple[str, ...]) -> str:
        for name in names:
            folded = name.casefold()
            if any(keyword in folded for keyword in keywords):
                return name
        return ""

    def load_sheet(self, name):
        if not self.workbook_path or not name:
            return

        try:
            self.data = pd.read_excel(self.workbook_path, sheet_name=name)

            self.columns.clear()
            for column in numeric_columns(self.data):
                self.columns.addItem(column)

            names = [str(column) for column in self.data.columns]

            combo_settings = (
                (self.id_column, ("serial", "serial number", "sn")),
                (self.vendor_column, ("vendor", "supplier", "manufacturer")),
                (self.bm_column, ("build matrix", "bm")),
                (self.group_column, ("vendor", "supplier", "manufacturer", "build matrix", "product")),
            )

            for combo, keywords in combo_settings:
                previous = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItem("")
                combo.addItems(names)

                if previous in names:
                    combo.setCurrentText(previous)
                else:
                    guess = self._guess_column(names, keywords)
                    if guess:
                        combo.setCurrentText(guess)

                combo.blockSignals(False)

            if self.vendor_column.currentText():
                self.group_column.setCurrentText(self.vendor_column.currentText())

            self._refresh_vendor_filter()
            self.update_chart()

        except Exception as exc:
            QMessageBox.warning(self, "Could not read worksheet", str(exc))

    def _refresh_vendor_filter(self):
        if self.data is None:
            return

        column = self.group_column.currentText()
        previous = self.vendor_filter.currentText()
        values = [""]

        if column and column in self.data.columns:
            values += sorted(
                str(value)
                for value in self.data[column].dropna().astype(str).unique()
            )

        self.vendor_filter.blockSignals(True)
        self.vendor_filter.clear()
        self.vendor_filter.addItems(values)
        if previous in values:
            self.vendor_filter.setCurrentText(previous)
        self.vendor_filter.blockSignals(False)

    def selected_columns(self) -> list[str]:
        return [item.text() for item in self.columns.selectedItems()]

    @staticmethod
    def _number(line: QLineEdit):
        text = line.text().strip()
        if not text:
            return None
        value = float(text)
        if not np.isfinite(value):
            raise ValueError("Spec values must be finite numbers.")
        return value

    def update_chart(self, *_):
        if self.data is None:
            return

        columns = self.selected_columns()
        if not columns:
            return

        try:
            usl = self._number(self.usl)
            lsl = self._number(self.lsl)
            target = self._number(self.target)

            render_chart(
                self.data,
                columns,
                self.chart_type.currentText(),
                group_column=self.group_column.currentText(),
                id_column=self.id_column.currentText(),
                vendor_column=self.vendor_column.currentText(),
                bm_column=self.bm_column.currentText(),
                usl=usl,
                lsl=lsl,
                target=target,
                highlight_oos=self.highlight.isChecked(),
                selected_vendor=self.vendor_filter.currentText(),
                sample_limit=max(50, int(self.sample.text() or 500)),
                figure=self.figure,
            )
            self.canvas.draw_idle()

            stats, overall = analyze_measurements(
                self.data, columns, usl, lsl
            )
            vendors = vendor_summary(
                self.data,
                columns,
                self.group_column.currentText(),
                usl,
                lsl,
            )
            spec = pd.DataFrame([
                {"Metric": key, "Value": value}
                for key, value in overall.items()
                if key in ("Above USL", "Below LSL", "Pass", "Fail", "Yield %")
            ])
            findings = dashboard_findings(
                self.data,
                columns,
                self.group_column.currentText(),
                usl,
                lsl,
            )

            self._show_table(vendors if not vendors.empty else stats)

            messages = []
            if not findings.empty:
                for row in findings.itertuples():
                    messages.append(
                        f"{row.Finding}: {row.Category} ({row.Count:,})"
                    )

            if not vendors.empty:
                best = vendors.iloc[0]["Vendor"]
                messages.insert(
                    0,
                    f"Best ranked vendor: {best} "
                    "(ranking prioritizes yield, Cpk, then lower variation).",
                )

                out_of_spec_total = vendors["Out-of-spec Count"].sum()
                if out_of_spec_total:
                    top = vendors.sort_values(
                        "Out-of-spec Count", ascending=False
                    ).iloc[0]
                    share = top["Out-of-spec Count"] / out_of_spec_total * 100
                    messages.append(
                        f"{top['Vendor']} contributes {share:.1f}% "
                        "of observed out-of-spec results."
                    )

            self.findings.setText(
                "\n".join(messages)
                if messages
                else "No out-of-spec units found (or no spec limits entered)."
            )

            sample_limit = max(50, int(self.sample.text() or 500))
            self.status.setText(
                f"{len(self.data):,} rows · {len(columns)} measurement points "
                f"· chart sampling is capped at {sample_limit} units"
            )

            self._last_stats = stats
            self._last_overall = overall
            self._last_vendors = vendors
            self._last_spec = spec

        except Exception as exc:
            self.status.setText(str(exc))

    def _show_table(self, frame):
        self.table.clear()
        self.table.setRowCount(min(100, len(frame)))
        self.table.setColumnCount(len(frame.columns))
        self.table.setHorizontalHeaderLabels([str(column) for column in frame.columns])

        for row_index, row in enumerate(frame.head(100).itertuples(index=False, name=None)):
            for column_index, value in enumerate(row):
                if pd.isna(value):
                    text = ""
                elif isinstance(value, (float, np.floating)):
                    text = f"{value:,.3f}"
                else:
                    text = str(value)
                self.table.setItem(row_index, column_index, QTableWidgetItem(text))

        self.table.resizeColumnsToContents()

    def save_png(self):
        self._save_figure("PNG image (*.png)", "png")

    def save_pdf(self):
        self._save_figure("PDF document (*.pdf)", "pdf")

    def _save_figure(self, file_filter, suffix):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save chart",
            f"EngineeringChart.{suffix}",
            file_filter,
        )
        if not path:
            return
        try:
            self.figure.savefig(path, format=suffix, dpi=160, bbox_inches="tight")
        except Exception as exc:
            QMessageBox.warning(self, "Save failed", str(exc))

    def export_excel(self):
        if self.data is None or not hasattr(self, "_last_stats"):
            QMessageBox.information(
                self,
                "Export",
                "Load a worksheet and select measurement columns first.",
            )
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export engineering analysis",
            "Engineering_Charts.xlsx",
            "Excel workbook (*.xlsx)",
        )
        if not path:
            return

        try:
            spc = pd.DataFrame([
                {
                    "Metric": "Control chart",
                    "Value": "X-bar chart; center ± 3σ control limits",
                }
            ])
            output = export_engineering_charts(
                path,
                self.figure,
                self._last_stats,
                self._last_overall,
                self._last_vendors,
                self._last_spec,
                spc,
            )
            QMessageBox.information(
                self,
                "Export complete",
                f"Engineering charts exported to:\n{output}",
            )
        except Exception as exc:
            QMessageBox.warning(self, "Export failed", str(exc))


def main():
    app = QApplication(sys.argv)
    window = EngineeringChartWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
