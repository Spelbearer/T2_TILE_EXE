import sys

from PyQt6 import QtWidgets, QtCore, QtGui
# import pandas as pd
from pandas import isna, read_excel ,read_csv, concat, DataFrame, merge
import os
import re
from openpyxl import load_workbook
from s2sphere import CellId, LatLng
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QGuiApplication

WKT_POINT_RE = re.compile(r"POINT\s*\(\s*([\d.\-]+)\s+([\d.\-]+)\s*\)")


LIGHT_THEME = {
    "window_bg": "#F5F5F7",
    "surface": "#FFFFFF",
    "text": "#1D1D1F",
    "secondary": "#6E6E73",
    "border": "#D2D2D7",
    "accent": "#0A84FF",
    "hover": "#F2F2F7",
    "pressed": "#E5E5EA",
}

DARK_THEME = {
    "window_bg": "#1C1C1E",
    "surface": "#2C2C2E",
    "text": "#FFFFFF",
    "secondary": "#98989D",
    "border": "#3A3A3C",
    "accent": "#0A84FF",
    "hover": "#3A3A3C",
    "pressed": "#2C2C2E",
}


def build_stylesheet(theme: dict) -> str:
    return f"""
    QWidget {{
        background: {theme['window_bg']};
        color: {theme['text']};
    }}
    QGroupBox#formatBox, QGroupBox#uploadBox, QGroupBox#outputBox {{
        background: {theme['surface']};
        border: 1px solid {theme['border']};
        border-radius: 8px;
        margin-top: 8px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 4px;
        color: {theme['text']};
        font-size: 17px;
        font-weight: 600;
    }}
    QLineEdit {{
        background: {theme['surface']};
        border: 1px solid {theme['border']};
        border-radius: 6px;
        padding: 6px 10px;
        color: {theme['text']};
    }}
    QLineEdit:focus {{
        border: 1px solid {theme['accent']};
        outline: none;
    }}
    QPushButton {{
        background: {theme['surface']};
        border: 1px solid {theme['border']};
        border-radius: 6px;
        padding: 6px 12px;
        color: {theme['text']};
    }}
    QPushButton:hover {{ background: {theme['hover']}; }}
    QPushButton:pressed {{ background: {theme['pressed']}; }}
    QPushButton:focus {{ border: 1px solid {theme['accent']}; }}
    QComboBox {{
        background: {theme['surface']};
        border: 1px solid {theme['border']};
        border-radius: 6px;
        padding: 6px 10px;
        color: {theme['text']};
    }}
    QComboBox:focus {{ border: 1px solid {theme['accent']}; }}
    """


def apply_theme(app: QtWidgets.QApplication, dark: bool = False) -> None:
    theme = DARK_THEME if dark else LIGHT_THEME
    app.setStyleSheet(build_stylesheet(theme))


def set_app_font(app: QtWidgets.QApplication) -> None:
    for family in ["SF Pro Text", "Inter", "Segoe UI", "Helvetica Neue", "Arial"]:
        font = QtGui.QFont(family, 13)
        if QtGui.QFontInfo(font).family() == family:
            app.setFont(font)
            break

class ProcessingWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int)
    finished = QtCore.pyqtSignal(object)
    error = QtCore.pyqtSignal(str)

    def __init__(self, file_path, match_file_path, input_format, output_dir, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.match_file_path = match_file_path
        self.input_format = input_format
        self.output_dir = output_dir
        self.columns_needed = [
            "s2_cell_id_13",
            "geounit_name",
            "ADM_name",
            "town_name",
            "tele2_scoring_qual",
            "mts_scoring_qual",
            "megafon_scoring_qual",
            "beeline_scoring_qual",
            "gap_scorinq_qual_mts",
            "gap_scorinq_qual_megafon",
            "gap_scorinq_qual_beeline",
            "Sale_Potential",
            "SAVE_potential",
        ]

    def parse_position(self, position):
        try:
            if isinstance(position, float) and isna(position):
                return None, None
            wkt_match = WKT_POINT_RE.match(str(position))
            if wkt_match:
                lon = float(wkt_match.group(1))
                lat = float(wkt_match.group(2))
                return lat, lon
        except Exception:
            pass
        return None, None

    def get_tile_id(self, lat, lon):
        if lat is None or lon is None:
            return None
        try:
            return str(CellId.from_lat_lng(LatLng.from_degrees(lat, lon)).parent(13).id())
        except Exception:
            return None

    @QtCore.pyqtSlot()
    def run(self):
        try:
            result = self.process_file()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))

    def process_file(self):
        fmt = self.input_format
        # Determine how to read the input file. Excel files are read from the
        # first sheet, while text-based formats use a semicolon-separated CSV
        # reader as before.
        if self.file_path.lower().endswith(('.xls', '.xlsx')):
            df = read_excel(self.file_path, sheet_name=0)
        else:
            df = read_csv(self.file_path, sep=';')

        if fmt == 'WKT':
            if 'BS_POSITION' not in df.columns:
                raise ValueError("Нет колонки 'BS_POSITION'")
        elif fmt == 'LAT / LON':
            if 'LATITUDE' not in df.columns or 'LONGITUDE' not in df.columns:
                raise ValueError("Требуются колонки 'LATITUDE' и 'LONGITUDE'")

            def to_wkt(lat, lon):
                try:
                    lat_f = float(str(lat).replace(',', '.'))
                    lon_f = float(str(lon).replace(',', '.'))
                    return f"POINT ({lon_f} {lat_f})"
                except Exception:
                    return None

            df['BS_POSITION'] = df.apply(lambda r: to_wkt(r['LATITUDE'], r['LONGITUDE']), axis=1)
        else:
            raise ValueError(f"Неподдерживаемый формат: {fmt}")

        total_rows = len(df)
        tile_ids = []

        for i, row in enumerate(df.itertuples(), 1):
            position = getattr(row, 'BS_POSITION', None)
            lat, lon = self.parse_position(position)
            tile_id = self.get_tile_id(lat, lon)
            tile_ids.append(tile_id)
            if i % 100 == 0 or i == total_rows:
                self.progress.emit(i, total_rows)

        df['tile_id'] = tile_ids
        tile_ids_set = set(df['tile_id'].astype(str).str.strip())

        matches = []
        chunk_size = 100_000
        found_rows = 0
        total_match_rows = 0
        for chunk in read_csv(self.match_file_path, sep=';', dtype=str, chunksize=chunk_size, usecols=self.columns_needed):
            chunk['s2_cell_id_13'] = chunk['s2_cell_id_13'].astype(str).str.strip()
            filtered = chunk[(chunk['s2_cell_id_13'].isin(tile_ids_set))]
            found_rows += len(filtered)
            total_match_rows += len(chunk)
            if not filtered.empty:
                matches.append(filtered)

        if matches:
            df2_filtered = concat(matches, ignore_index=True)
        else:
            df2_filtered = DataFrame(columns=self.columns_needed)

        df['tile_id'] = df['tile_id'].astype(str).str.strip()
        df2_filtered['s2_cell_id_13'] = df2_filtered['s2_cell_id_13'].astype(str).str.strip()

        merged = merge(
            df,
            df2_filtered,
            how='left',
            left_on='tile_id',
            right_on='s2_cell_id_13',
            suffixes=('', '_spr')
        )

        if 'tile_id' in merged.columns:
            merged.drop(columns=['tile_id'], inplace=True)

        if fmt == 'LAT / LON':
            for col in ['LATITUDE', 'LONGITUDE']:
                if col in merged.columns:
                    merged.drop(columns=[col], inplace=True)

        fn1 = os.path.basename(self.file_path)
        fn2 = os.path.basename(self.match_file_path)
        result_name = f"Потенциал"
        out_path = os.path.join(self.output_dir, result_name)
        if not out_path.lower().endswith('.xlsx'):
            out_path += '.xlsx'

        merged.to_excel(out_path, index=False)
        wb = load_workbook(out_path)
        ws = wb.active
        ws.auto_filter.ref = ws.dimensions
        wb.save(out_path)

        return out_path, len(merged), found_rows, total_match_rows

class TileIntersectionApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Определение тайлов")
        self.match_file_path = (
            r"\\corp.tele2.ru\operations_MR\Operations_All\Потенциал_рынка\яархив_исходники\T_Potential\T_Potential_filtered_last.txt"
        )
        self.file_path = None
        self.input_format = "WKT"
        self.output_dir = os.getcwd()
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        self.format_box = QtWidgets.QGroupBox("1. Выберите формат входных данных")
        self.format_box.setObjectName("formatBox")
        fb_layout = QtWidgets.QVBoxLayout(self.format_box)
        self.format_combo = QtWidgets.QComboBox()
        self.format_combo.addItems(["WKT", "LAT / LON"])
        self.format_combo.currentIndexChanged.connect(self.on_format_change)
        fb_layout.addWidget(self.format_combo)
        layout.addWidget(self.format_box)

        self.upload_box = QtWidgets.QGroupBox("2. Загрузка исходного файла")
        self.upload_box.setObjectName("uploadBox")
        ub_layout = QtWidgets.QHBoxLayout(self.upload_box)
        self.file_line = QtWidgets.QLineEdit()
        self.file_line.setReadOnly(True)
        self.btn_browse = QtWidgets.QPushButton("Выбрать файл…")
        self.btn_browse.clicked.connect(self.select_file)
        ub_layout.addWidget(self.file_line)
        ub_layout.addWidget(self.btn_browse)
        layout.addWidget(self.upload_box)

        self.output_box = QtWidgets.QGroupBox("3. Папка для сохранения")
        self.output_box.setObjectName("outputBox")
        ob_layout = QtWidgets.QHBoxLayout(self.output_box)
        self.output_line = QtWidgets.QLineEdit()
        self.output_line.setReadOnly(True)
        self.output_line.setText(self.output_dir)
        self.btn_browse_output = QtWidgets.QPushButton("Выбрать папку…")
        self.btn_browse_output.clicked.connect(self.select_output_dir)
        ob_layout.addWidget(self.output_line)
        ob_layout.addWidget(self.btn_browse_output)
        layout.addWidget(self.output_box)

        self.btn_process = QtWidgets.QPushButton("Обработать")
        self.btn_process.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_process)

        self.progress = QtWidgets.QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        self.result_label = QtWidgets.QLabel()
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

        for box in (self.format_box, self.upload_box, self.output_box):
            effect = QtWidgets.QGraphicsDropShadowEffect(
                blurRadius=12,
                xOffset=0,
                yOffset=3,
                color=QtGui.QColor(0, 0, 0, 25),
            )
            box.setGraphicsEffect(effect)

    def on_format_change(self, index):
        self.input_format = self.format_combo.currentText()

    def select_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",
            "Text/Excel Files (*.txt *.csv *.xls *.xlsx);;All Files (*)",
        )
        if path:
            self.file_path = path
            self.file_line.setText(path)

    def select_output_dir(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self,
            "Выберите папку",
            self.output_dir,
        )
        if path:
            self.output_dir = path
            self.output_line.setText(path)

    def start_processing(self):
        if not self.file_path:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Выберите файл")
            return
        self.btn_process.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.result_label.clear()

        self.thread = QtCore.QThread(self)
        self.worker = ProcessingWorker(self.file_path, self.match_file_path, self.input_format, self.output_dir)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_progress(self, current, total):
        self.progress.setMaximum(total)
        self.progress.setValue(current)

    def on_finished(self, result):
        self.btn_process.setEnabled(True)
        self.progress.setVisible(False)
        if result:
            out_path, final_count, found_rows, total_rows = result
            self.result_label.setText(
                f"Результат сохранён в:\n{out_path}"
            )
            QtWidgets.QMessageBox.information(self, "Готово", f"Файл сохранён: {out_path}")
        else:
            self.result_label.setText("Ошибка во время обработки.")

    def on_error(self, message):
        self.btn_process.setEnabled(True)
        self.progress.setVisible(False)
        QtWidgets.QMessageBox.critical(self, "Ошибка", message)

if __name__ == "__main__":
    dark = "--dark" in sys.argv
    # Enable High DPI pixmaps if supported by the current Qt build.
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
    Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
)
    app = QtWidgets.QApplication(sys.argv)

    set_app_font(app)
    apply_theme(app, dark)
    window = TileIntersectionApp()
    window.show()
    app.exec()
