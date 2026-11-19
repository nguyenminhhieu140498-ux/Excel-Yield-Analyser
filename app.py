from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtGui import QDrag
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from engine import (
    AnalyzerError,
    commonality,
    commonality_matrix,
    export_report,
    list_sheets,
    merge_yield_data,
    read_sheets,
)


FIELD_MIME = "application/x-excel-yield-analyzer-field"
STAT_MIME = "application/x-excel-yield-analyzer-stat"
STATISTICS = ["N", "Row %"]
STATUS_ORDER = ["Fail", "Pass"]


class SheetDialog(QDialog):
    def __init__(self, path: str, parent=None, multi: bool = True):
        super().__init__(parent)
        self.setWindowTitle("Select worksheets")
        self.resize(360, 420)

        layout = QVBoxLayout(self)
        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection
            if multi
            else QAbstractItemView.SelectionMode.SingleSelection
        )

        for sheet in list_sheets(path):
            self.sheet_list.addItem(sheet)

        if self.sheet_list.count():
            self.sheet_list.item(0).setSelected(True)

        layout.addWidget(QLabel("Select worksheet(s):"))
        layout.addWidget(self.sheet_list)

        buttons = QHBoxLayout()
        use_button = QPushButton("Use selected")
        cancel_button = QPushButton("Cancel")
        use_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(use_button)
        buttons.addWidget(cancel_button)
        layout.addLayout(buttons)

    def selected_sheets(self) -> list[str]:
        return [item.text() for item in self.sheet_list.selectedItems()]


class DragSourceList(QListWidget):
    def __init__(self, mime_type: str, parent=None):
        super().__init__(parent)
        self.mime_type = mime_type
        self.setDragEnabled(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)

    def startDrag(self, supported_actions):
        item = self.currentItem()
        if item is None:
            return

        mime = QMimeData()
        mime.setData(self.mime_type, item.text().encode("utf-8"))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.CopyAction)


class FieldList(DragSourceList):
    """Available-column catalog; dropping a field back here removes it."""

    def __init__(self, parent=None):
        super().__init__(FIELD_MIME, parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(FIELD_MIME):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(FIELD_MIME):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        source = event.source()
        if isinstance(source, DropZone):
            row = source.currentRow()
            if row >= 0:
                source.takeItem(row)
            event.acceptProposedAction()
        else:
            event.ignore()


class StatisticList(DragSourceList):
    def __init__(self, parent=None):
        super().__init__(STAT_MIME, parent)
        for statistic in STATISTICS:
            self.addItem(statistic)


class DropZone(QListWidget):
    def __init__(self, accepted_mime: str, parent=None):
        super().__init__(parent)
        self.accepted_mime = accepted_mime
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDropIndicatorShown(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

    def startDrag(self, supported_actions):
        item = self.currentItem()
        if item is None:
            return

        mime = QMimeData()
        mime.setData(self.accepted_mime, item.text().encode("utf-8"))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(self.accepted_mime):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(self.accepted_mime):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasFormat(self.accepted_mime):
            event.ignore()
            return

        text = bytes(
            event.mimeData().data(self.accepted_mime)
        ).decode("utf-8").strip()
        if not text:
            event.ignore()
            return

        source = event.source()
        exists = any(self.item(i).text() == text for i in range(self.count()))
        if not exists:
            self.addItem(text)

        if isinstance(source, DropZone) and source is not self:
            row = source.currentRow()
            if row >= 0:
                source.takeItem(row)

        event.acceptProposedAction()

    def names(self) -> list[str]:
        return [self.item(i).text() for i in range(self.count())]

    def remove_selected(self):
        row = self.currentRow()
        if row >= 0:
            self.takeItem(row)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Yield Analyzer")
        self.resize(1450, 900)

        self.yield_data: pd.DataFrame | None = None
        self.mapping_data: pd.DataFrame | None = None
        self.merged: pd.DataFrame | None = None
        self.pivot_data: pd.DataFrame | None = None
        self.common_data: pd.DataFrame | None = None
        self.matrix: pd.DataFrame | None = None

        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout(root)
        self.tabs = QTabWidget()
        root_layout.addWidget(self.tabs)

        self.import_tab = QWidget()
        self.mapping_tab = QWidget()
        self.analysis_tab = QWidget()
        self.export_tab = QWidget()

        self.tabs.addTab(self.import_tab, "Import Files")
        self.tabs.addTab(self.mapping_tab, "Mapping")
        self.tabs.addTab(self.analysis_tab, "Analysis")
        self.tabs.addTab(self.export_tab, "Export")

        self._build_import_tab()
        self._build_mapping_tab()
        self._build_analysis_tab()
        self._build_export_tab()

    def _build_import_tab(self):
        layout = QVBoxLayout(self.import_tab)

        row = QHBoxLayout()
        choose_yield = QPushButton("Choose Yield File…")
        choose_yield.clicked.connect(self.choose_yield_file)
        self.yield_label = QLabel("No yield workbook selected")
        row.addWidget(choose_yield)
        row.addWidget(self.yield_label, 1)
        layout.addLayout(row)

        row = QHBoxLayout()
        choose_mapping = QPushButton("Choose Mapping File…")
        choose_mapping.clicked.connect(self.choose_mapping_file)
        self.mapping_label = QLabel("No mapping workbook selected")
        row.addWidget(choose_mapping)
        row.addWidget(self.mapping_label, 1)
        layout.addLayout(row)

        layout.addWidget(QLabel("Preview — first 100 rows"))
        self.import_preview = QTableWidget()
        self._style_table(self.import_preview)
        layout.addWidget(self.import_preview, 1)

    def _build_mapping_tab(self):
        layout = QVBoxLayout(self.mapping_tab)

        self.yield_key = QComboBox()
        self.mapping_key = QComboBox()
        self.join_type = QComboBox()
        self.join_type.addItems(["left", "inner", "right"])
        self.join_type.setCurrentText("left")

        layout.addWidget(QLabel("Yield merge key"))
        layout.addWidget(self.yield_key)
        layout.addWidget(QLabel("Mapping merge key"))
        layout.addWidget(self.mapping_key)
        layout.addWidget(QLabel("Join type"))
        layout.addWidget(self.join_type)

        fuzzy_row = QHBoxLayout()
        self.fuzzy_enabled = QCheckBox(
            "Allow similar (fuzzy) merge-key matches"
        )
        self.fuzzy_threshold = QSpinBox()
        self.fuzzy_threshold.setRange(50, 100)
        self.fuzzy_threshold.setValue(90)
        self.fuzzy_threshold.setSuffix("%")
        fuzzy_row.addWidget(self.fuzzy_enabled)
        fuzzy_row.addWidget(QLabel("Minimum similarity"))
        fuzzy_row.addWidget(self.fuzzy_threshold)
        layout.addLayout(fuzzy_row)

        layout.addWidget(QLabel("Mapping columns to append"))
        self.append_list = QListWidget()
        self.append_list.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection
        )
        layout.addWidget(self.append_list, 1)

        merge_button = QPushButton("Merge data")
        merge_button.clicked.connect(self.run_merge)
        layout.addWidget(merge_button)

        self.merge_stats = QLabel("Merge statistics will appear here.")
        self.merge_stats.setWordWrap(True)
        layout.addWidget(self.merge_stats)

    def _build_analysis_tab(self):
        layout = QVBoxLayout(self.analysis_tab)

        instruction = QLabel(
            "Drag grouping fields into Rows. Drag test-result fields into "
            "Values. Each test field is displayed as Fail and Pass columns, "
            "with N and Row % under each status."
        )
        instruction.setWordWrap(True)
        layout.addWidget(instruction)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        source_panel = QWidget()
        source_layout = QVBoxLayout(source_panel)
        source_layout.addWidget(QLabel("Available columns"))

        self.field_search = QLineEdit()
        self.field_search.setPlaceholderText("Search columns…")
        self.field_search.textChanged.connect(self.filter_fields)
        source_layout.addWidget(self.field_search)

        self.field_list = FieldList()
        source_layout.addWidget(self.field_list, 1)

        source_layout.addWidget(QLabel("Statistics"))
        self.statistic_list = StatisticList()
        source_layout.addWidget(self.statistic_list, 1)

        work_panel = QWidget()
        work_layout = QVBoxLayout(work_panel)
        zones = QHBoxLayout()

        self.rows_zone = DropZone(FIELD_MIME)
        self.columns_zone = DropZone(FIELD_MIME)
        self.values_zone = DropZone(FIELD_MIME)
        self.statistics_zone = DropZone(STAT_MIME)

        zones.addWidget(self._zone_panel("Rows (e.g. phase, vendors)", self.rows_zone), 1)
        zones.addWidget(self._zone_panel("Columns (optional grouping)", self.columns_zone), 1)
        zones.addWidget(self._zone_panel("Values (Pass/Fail test fields)", self.values_zone), 1)
        zones.addWidget(self._zone_panel("Statistics (N, Row %)", self.statistics_zone), 1)
        work_layout.addLayout(zones, 1)

        controls = QHBoxLayout()
        remove_button = QPushButton("Remove selected")
        remove_button.clicked.connect(self.remove_selected_item)
        reset_button = QPushButton("Start over")
        reset_button.clicked.connect(self.reset_analysis)
        update_button = QPushButton("Update table")
        update_button.clicked.connect(self.run_tabulate)

        controls.addWidget(remove_button)
        controls.addWidget(reset_button)
        controls.addStretch(1)
        controls.addWidget(update_button)
        work_layout.addLayout(controls)

        work_layout.addWidget(QLabel("Resulting table"))
        self.analysis_preview = QTableWidget()
        self._style_table(self.analysis_preview)
        self.analysis_preview.setWordWrap(True)
        self.analysis_preview.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        work_layout.addWidget(self.analysis_preview, 4)

        self.analysis_status = QLabel(
            "Import data, then drag fields and statistics into the drop zones."
        )
        self.analysis_status.setWordWrap(True)
        work_layout.addWidget(self.analysis_status)

        splitter.addWidget(source_panel)
        splitter.addWidget(work_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 4)
        layout.addWidget(splitter, 1)

    @staticmethod
    def _zone_panel(title: str, zone: DropZone) -> QWidget:
        panel = QWidget()
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(2, 2, 2, 2)
        panel_layout.addWidget(QLabel(title))
        panel_layout.addWidget(zone)
        return panel

    def _build_export_tab(self):
        layout = QVBoxLayout(self.export_tab)
        layout.addWidget(QLabel("Optional vendor box plot"))

        self.boxplot_enabled = QCheckBox(
            "Include vendor box plot in Excel report"
        )
        self.boxplot_enabled.setChecked(True)
        layout.addWidget(self.boxplot_enabled)

        self.vendor_column = QComboBox()
        self.measurement_column = QComboBox()
        self.status_column = QComboBox()

        for label, widget in (
            ("Vendor column", self.vendor_column),
            ("Numeric measurement column", self.measurement_column),
            ("Pass/fail status column", self.status_column),
        ):
            layout.addWidget(QLabel(label))
            layout.addWidget(widget)

        labels = QHBoxLayout()
        labels.addWidget(QLabel("Pass label"))
        self.pass_label = QLineEdit("pass")
        labels.addWidget(self.pass_label)
        labels.addWidget(QLabel("Fail label"))
        self.fail_label = QLineEdit("fail")
        labels.addWidget(self.fail_label)
        layout.addLayout(labels)

        export_button = QPushButton("Export Excel Report…")
        export_button.clicked.connect(self.export_excel)
        layout.addWidget(export_button)

        self.export_status = QLabel("Import data to enable report export.")
        self.export_status.setWordWrap(True)
        layout.addWidget(self.export_status)
        layout.addStretch(1)

    @staticmethod
    def _style_table(table: QTableWidget):
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        table.setWordWrap(True)
        table.setHorizontalScrollMode(
            QAbstractItemView.ScrollMode.ScrollPerPixel
        )
        table.setVerticalScrollMode(
            QAbstractItemView.ScrollMode.ScrollPerPixel
        )
        table.horizontalHeader().setVisible(True)
        table.horizontalHeader().setDefaultAlignment(
            Qt.AlignmentFlag.AlignCenter
        )
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        table.verticalHeader().setVisible(True)
        table.setStyleSheet(
            "QHeaderView::section {"
            " background-color: #e8edf5;"
            " color: #172033;"
            " font-weight: bold;"
            " border: 1px solid #b8c2d1;"
            " padding: 6px;"
            "}"
            "QTableWidget { gridline-color: #c8ced8; }"
        )

    def choose_yield_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose yield workbook",
            "",
            "Excel workbooks (*.xlsx *.xlsm)",
        )
        if not path:
            return

        try:
            dialog = SheetDialog(path, self, multi=True)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return

            self.yield_data = read_sheets(path, dialog.selected_sheets())
            self.merged = None
            self.yield_label.setText(Path(path).name)
            self.show_frame(self.import_preview, self.yield_data.head(100))
            self.refresh_columns()
            self.analysis_status.setText(
                f"Loaded {len(self.yield_data):,} yield rows."
            )
        except Exception as exc:
            self.show_error(exc)

    def choose_mapping_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose mapping workbook",
            "",
            "Excel workbooks (*.xlsx *.xlsm)",
        )
        if not path:
            return

        try:
            dialog = SheetDialog(path, self, multi=False)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return

            self.mapping_data = read_sheets(path, dialog.selected_sheets())
            self.mapping_label.setText(Path(path).name)
            self.refresh_columns()
        except Exception as exc:
            self.show_error(exc)

    def current_data(self) -> pd.DataFrame | None:
        return self.merged if self.merged is not None else self.yield_data

    def refresh_columns(self):
        if self.yield_data is not None:
            self.yield_key.clear()
            self.yield_key.addItems(
                [str(column) for column in self.yield_data.columns]
            )

        if self.mapping_data is not None:
            self.mapping_key.clear()
            self.mapping_key.addItems(
                [str(column) for column in self.mapping_data.columns]
            )
            self.append_list.clear()
            for column in self.mapping_data.columns:
                self.append_list.addItem(str(column))

        data = self.current_data()
        if data is None:
            return

        names = [str(column) for column in data.columns]
        self.field_list.clear()
        self.field_list.addItems(names)

        for combo in (
            self.vendor_column,
            self.measurement_column,
            self.status_column,
        ):
            previous = combo.currentText()
            combo.clear()
            combo.addItems([""] + names)
            if previous in names:
                combo.setCurrentText(previous)

        self._guess_boxplot_columns(data)

    def _guess_boxplot_columns(self, data: pd.DataFrame):
        columns = [str(column) for column in data.columns]

        if not self.vendor_column.currentText():
            candidate = next(
                (
                    c for c in columns
                    if "vendor" in c.casefold() or "supplier" in c.casefold()
                ),
                None,
            )
            if candidate:
                self.vendor_column.setCurrentText(candidate)

        if not self.status_column.currentText():
            candidate = next(
                (
                    c for c in columns
                    if any(word in c.casefold() for word in (
                        "status", "pass", "fail", "result", "judge"
                    ))
                ),
                None,
            )
            if candidate:
                self.status_column.setCurrentText(candidate)

        if not self.measurement_column.currentText():
            numeric = [
                str(column) for column in data.columns
                if pd.api.types.is_numeric_dtype(data[column])
            ]
            if numeric:
                self.measurement_column.setCurrentText(numeric[0])

    def run_merge(self):
        try:
            if self.yield_data is None or self.mapping_data is None:
                raise AnalyzerError(
                    "Choose both yield and mapping workbooks first."
                )

            append_columns = [
                item.text() for item in self.append_list.selectedItems()
            ]
            if not append_columns:
                raise AnalyzerError("Select mapping columns to append.")

            result = merge_yield_data(
                self.yield_data,
                self.mapping_data,
                self.yield_key.currentText(),
                self.mapping_key.currentText(),
                append_columns,
                self.join_type.currentText(),
                fuzzy=self.fuzzy_enabled.isChecked(),
                fuzzy_threshold=self.fuzzy_threshold.value(),
            )

            self.merged = result.data
            self.pivot_data = None
            self.common_data = None
            self.matrix = None

            self.show_frame(self.import_preview, self.merged.head(100))
            self.merge_stats.setText(
                f"Yield rows: {result.total_yield_rows:,}    "
                f"Matched: {result.matched_yield_rows:,}    "
                f"Unmatched: {result.unmatched_yield_rows:,}    "
                f"Match rate: {result.match_rate:.2f}%    "
                f"Duplicate mapping keys: {result.mapping_duplicate_keys:,}"
            )
            self.refresh_columns()
        except Exception as exc:
            self.show_error(exc)

    def filter_fields(self, text: str):
        query = text.strip().casefold()
        for index in range(self.field_list.count()):
            item = self.field_list.item(index)
            item.setHidden(query not in item.text().casefold())

    def remove_selected_item(self):
        for zone in (
            self.rows_zone,
            self.columns_zone,
            self.values_zone,
            self.statistics_zone,
        ):
            if zone.hasFocus() and zone.currentItem() is not None:
                zone.remove_selected()
                return

        QMessageBox.information(
            self,
            "Remove item",
            "Select an item in a drop zone, then click Remove selected.",
        )

    def reset_analysis(self):
        for zone in (
            self.rows_zone,
            self.columns_zone,
            self.values_zone,
            self.statistics_zone,
        ):
            zone.clear()

        self.analysis_preview.clear()
        self.analysis_preview.setRowCount(0)
        self.analysis_preview.setColumnCount(0)
        self.pivot_data = None
        self.common_data = None
        self.matrix = None
        self.analysis_status.setText("Analysis layout cleared.")

    def run_tabulate(self):
        try:
            data = self.current_data()
            if data is None:
                raise AnalyzerError("Import a yield workbook first.")

            row_fields = self.rows_zone.names()
            column_fields = self.columns_zone.names()
            value_fields = self.values_zone.names()
            statistics = self.statistics_zone.names()

            if not row_fields and not column_fields:
                raise AnalyzerError(
                    "Drag at least one grouping field into Rows or Columns."
                )
            if not value_fields:
                raise AnalyzerError(
                    "Drag one or more Pass/Fail test fields into Values."
                )
            if not statistics:
                raise AnalyzerError(
                    "Drag N and/or Row % into Statistics."
                )

            result = build_pass_fail_table(
                data=data,
                row_fields=row_fields,
                column_fields=column_fields,
                value_fields=value_fields,
                statistics=statistics,
            )

            self.pivot_data = result
            all_group_fields = row_fields + column_fields
            self.common_data = (
                commonality(data, all_group_fields)
                if all_group_fields
                else pd.DataFrame()
            )

            if len(all_group_fields) >= 2:
                self.matrix = commonality_matrix(
                    data,
                    all_group_fields[0],
                    all_group_fields[1],
                )
            else:
                self.matrix = pd.DataFrame()

            self.show_frame(self.analysis_preview, result.head(1000))
            self.analysis_status.setText(
                f"Generated {len(result):,} row(s). "
                "Each test field has Fail and Pass groups, each with N and Row %."
            )
        except Exception as exc:
            self.show_error(exc)

    def export_excel(self):
        try:
            data = self.current_data()
            if data is None:
                raise AnalyzerError("Import a yield workbook first.")

            path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Excel report",
                "Merged_Result.xlsx",
                "Excel workbook (*.xlsx)",
            )
            if not path:
                return

            output = export_report(
                path,
                data,
                self.pivot_data if self.pivot_data is not None else pd.DataFrame(),
                self.common_data if self.common_data is not None else pd.DataFrame(),
                self.common_data if self.common_data is not None else pd.DataFrame(),
                self.matrix if self.matrix is not None else pd.DataFrame(),
            )

            if self.boxplot_enabled.isChecked():
                add_vendor_boxplot_to_excel(
                    output_path=output,
                    data=data,
                    vendor_column=self.vendor_column.currentText(),
                    measurement_column=self.measurement_column.currentText(),
                    status_column=self.status_column.currentText(),
                    pass_label=self.pass_label.text().strip() or "pass",
                    fail_label=self.fail_label.text().strip() or "fail",
                )

            self.export_status.setText(f"Report saved: {output}")
            QMessageBox.information(
                self, "Export complete", f"Excel report saved to:\n{output}"
            )
        except Exception as exc:
            self.show_error(exc)

    @staticmethod
    def show_frame(table: QTableWidget, frame: pd.DataFrame):
        table.clear()
        table.setRowCount(len(frame))
        table.setColumnCount(len(frame.columns))
        table.setHorizontalHeaderLabels([str(c) for c in frame.columns])
        table.horizontalHeader().setVisible(True)
        table.verticalHeader().setVisible(True)

        for row_index, values in enumerate(
            frame.itertuples(index=False, name=None)
        ):
            for col_index, value in enumerate(values):
                if pd.isna(value):
                    text = ""
                elif isinstance(value, (int, np.integer)):
                    text = f"{int(value):,}"
                elif isinstance(value, (float, np.floating)):
                    text = f"{float(value):,.2f}"
                    if str(frame.columns[col_index]).endswith("Row %"):
                        text += "%"
                else:
                    text = str(value)

                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(row_index, col_index, item)

        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(False)

    def show_error(self, error):
        QMessageBox.warning(self, "Excel Yield Analyzer", str(error))


def _normalize_status(value) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip().casefold()
    if text in {"fail", "failed", "ng", "nok"}:
        return "Fail"
    if text in {"pass", "passed", "ok"}:
        return "Pass"
    return None


def build_pass_fail_table(
    data: pd.DataFrame,
    row_fields: list[str],
    column_fields: list[str],
    value_fields: list[str],
    statistics: list[str],
) -> pd.DataFrame:
    """Create one row per grouping combination and N/Row % per test/status."""
    row_fields = list(dict.fromkeys(row_fields))
    column_fields = list(dict.fromkeys(column_fields))
    value_fields = list(dict.fromkeys(value_fields))
    grouping_fields = list(dict.fromkeys(row_fields + column_fields))

    if not value_fields:
        raise AnalyzerError("Select at least one test field in Values.")

    missing = [
        field for field in grouping_fields + value_fields
        if field not in data.columns
    ]
    if missing:
        raise AnalyzerError("Column(s) not found: " + ", ".join(missing))

    statistics = list(dict.fromkeys(statistics))
    if not statistics or any(s not in STATISTICS for s in statistics):
        raise AnalyzerError("Select N and/or Row % from Statistics.")

    frame = data.copy()

    # Remove records with blank grouping values so no blank group is shown.
    for field in grouping_fields:
        frame = frame.loc[
            frame[field].notna()
            & frame[field].astype(str).str.strip().ne("")
        ]

    output_columns = list(grouping_fields)
    for value_field in value_fields:
        for status in STATUS_ORDER:
            for statistic in statistics:
                output_columns.append(
                    f"{value_field}\n{status}\n{statistic}"
                )

    if frame.empty:
        return pd.DataFrame(columns=output_columns)

    if grouping_fields:
        groups = frame.groupby(
            grouping_fields,
            dropna=False,
            observed=True,
            sort=False,
        )
        grouped_items = list(groups)
    else:
        grouped_items = [((), frame)]

    records = []
    for group_key, group in grouped_items:
        if not isinstance(group_key, tuple):
            group_key = (group_key,)

        record = dict(zip(grouping_fields, group_key))

        for value_field in value_fields:
            statuses = group[value_field].map(_normalize_status)
            counts = {
                status: int((statuses == status).sum())
                for status in STATUS_ORDER
            }
            denominator = sum(counts.values())

            for status in STATUS_ORDER:
                count = counts[status]
                row_percent = (
                    count / denominator * 100.0 if denominator else 0.0
                )

                if "N" in statistics:
                    record[f"{value_field}\n{status}\nN"] = count
                if "Row %" in statistics:
                    record[f"{value_field}\n{status}\nRow %"] = row_percent

        records.append(record)

    result = pd.DataFrame(records, columns=output_columns)

    # Keep integer count columns as integers rather than decimals.
    for column in result.columns:
        if column.endswith("\nN"):
            result[column] = result[column].fillna(0).astype("int64")

    return result


def add_vendor_boxplot_to_excel(
    output_path: str | Path,
    data: pd.DataFrame,
    vendor_column: str,
    measurement_column: str,
    status_column: str,
    pass_label: str = "pass",
    fail_label: str = "fail",
):
    required = {
        "Vendor": vendor_column,
        "measurement": measurement_column,
        "pass/fail status": status_column,
    }
    missing = [
        label for label, column in required.items()
        if not column or column not in data.columns
    ]
    if missing:
        raise AnalyzerError(
            "Select valid " + ", ".join(missing) + " columns in Export."
        )

    frame = data[[vendor_column, measurement_column, status_column]].copy()
    frame.columns = ["Vendor", "Measurement", "Status"]
    frame["Measurement"] = pd.to_numeric(
        frame["Measurement"], errors="coerce"
    )
    frame["Vendor"] = frame["Vendor"].astype("string").str.strip()
    frame = frame.dropna(subset=["Vendor", "Measurement"])
    frame = frame[frame["Vendor"].ne("")]

    if frame.empty:
        raise AnalyzerError(
            "No valid vendor and numeric measurement values are available."
        )

    status = frame["Status"].astype("string").str.strip().str.casefold()
    pass_key = pass_label.strip().casefold()
    fail_key = fail_label.strip().casefold()
    frame["Plot status"] = np.select(
        [status.eq(pass_key), status.eq(fail_key)],
        ["Pass", "Fail"],
        default="Other",
    )

    vendors = frame["Vendor"].astype(str).unique().tolist()
    values = [
        frame.loc[
            frame["Vendor"].astype(str).eq(vendor), "Measurement"
        ].to_numpy()
        for vendor in vendors
    ]

    summary = (
        frame.groupby("Vendor", observed=True)
        .agg(
            N=("Measurement", "count"),
            Mean=("Measurement", "mean"),
            Median=("Measurement", "median"),
            Min=("Measurement", "min"),
            Max=("Measurement", "max"),
        )
        .reset_index()
    )

    fail_counts = (
        frame.assign(Fail=frame["Plot status"].eq("Fail"))
        .groupby("Vendor", observed=True)["Fail"]
        .sum()
        .rename("Fail Count")
        .reset_index()
    )
    pass_counts = (
        frame.assign(Pass=frame["Plot status"].eq("Pass"))
        .groupby("Vendor", observed=True)["Pass"]
        .sum()
        .rename("Pass Count")
        .reset_index()
    )
    summary = summary.merge(fail_counts, on="Vendor", how="left")
    summary = summary.merge(pass_counts, on="Vendor", how="left")
    summary["Fail Rate %"] = (
        summary["Fail Count"] / summary["N"] * 100
    ).round(2)

    with tempfile.TemporaryDirectory(prefix="eya_boxplot_") as temp_dir:
        image_path = Path(temp_dir) / "vendor_boxplot.png"

        fig, ax = plt.subplots(figsize=(max(10, len(vendors)), 7))
        positions = np.arange(1, len(vendors) + 1)
        box = ax.boxplot(
            values,
            positions=positions,
            patch_artist=True,
            showmeans=True,
            showfliers=True,
        )

        for patch in box["boxes"]:
            patch.set_facecolor("#b7c9e2")
            patch.set_edgecolor("#3d5a80")
        for mean in box["means"]:
            mean.set_marker("D")
            mean.set_markerfacecolor("#111111")
            mean.set_markeredgecolor("#111111")

        rng = np.random.default_rng(42)
        colors = {"Pass": "#16803c", "Fail": "#d62728"}

        for position, vendor in zip(positions, vendors):
            vendor_rows = frame[frame["Vendor"].astype(str).eq(vendor)]
            for label in ("Pass", "Fail"):
                subset = vendor_rows[vendor_rows["Plot status"].eq(label)]
                if subset.empty:
                    continue
                if len(subset) > 250:
                    subset = subset.sample(250, random_state=42)
                x = position + rng.uniform(-0.12, 0.12, len(subset))
                ax.scatter(
                    x,
                    subset["Measurement"],
                    s=16,
                    alpha=0.55,
                    color=colors[label],
                    label=label,
                    edgecolors="none",
                )

        handles, labels = ax.get_legend_handles_labels()
        unique = dict(zip(labels, handles))
        if unique:
            ax.legend(unique.values(), unique.keys(), title="Status")

        ax.set_xticks(positions)
        ax.set_xticklabels(vendors, rotation=45, ha="right")
        ax.set_xlabel("Vendor")
        ax.set_ylabel(measurement_column)
        ax.set_title(f"{measurement_column} distribution by vendor (mean shown)")
        ax.grid(axis="y", alpha=0.25)
        fig.tight_layout()
        fig.savefig(image_path, dpi=150)
        plt.close(fig)

        workbook = load_workbook(output_path)
        sheet_name = "6_VendorBoxPlot"
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
        sheet = workbook.create_sheet(sheet_name)

        sheet["A1"] = "Vendor box plot"
        sheet["A2"] = f"Vendor field: {vendor_column}"
        sheet["A3"] = f"Measurement: {measurement_column}"
        sheet["A4"] = f"Status field: {status_column}"
        sheet["A5"] = f"Pass: {pass_label} | Fail: {fail_label}"
        sheet["A7"] = "Summary by vendor"

        for col_index, column in enumerate(summary.columns, start=1):
            sheet.cell(row=8, column=col_index, value=str(column))

        for row_index, values_row in enumerate(
            summary.itertuples(index=False, name=None), start=9
        ):
            for col_index, value in enumerate(values_row, start=1):
                if pd.isna(value):
                    value = None
                elif isinstance(value, np.generic):
                    value = value.item()
                sheet.cell(row=row_index, column=col_index, value=value)

        image = XLImage(str(image_path))
        image.width = 1000
        image.height = 600
        image.anchor = f"A{max(11, len(summary) + 11)}"
        sheet.add_image(image)

        for cells in sheet.columns:
            letter = cells[0].column_letter
            width = max(
                14,
                max(len(str(cell.value or "")) for cell in cells) + 2,
            )
            sheet.column_dimensions[letter].width = min(32, width)

        workbook.save(output_path)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Excel Yield Analyzer")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())