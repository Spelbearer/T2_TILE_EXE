import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import math
import os
import re
import threading
import time
from s2sphere import RegionCoverer, Cap, LatLng, Angle
from openpyxl import load_workbook


class TileIntersectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Определение тайлов")
        self.root.geometry("830x560")
        self.root.resizable(False, False)

        self.file_path = None
        self.match_file_path = None
        self.input_format = 'wkt'

        self.bg_color = "#f6f6f6"
        frame = tk.Frame(self.root, bg=self.bg_color)
        frame.pack(fill="both", expand=True, padx=12, pady=10)

        # 1. Выбор формата входных данных
        tk.Label(
            frame,
            text="1. Выберите формат входных данных:",
            anchor='w', bg=self.bg_color, font=("Arial", 11, 'bold'),
        ).pack(fill="x", pady=(0, 5))
        self.input_format_var = tk.StringVar(value='wkt')
        format_options = ['wkt', 'coords']
        self.format_combo = ttk.Combobox(frame, textvariable=self.input_format_var, values=format_options,
                                        state="readonly", width=15)
        self.format_combo.pack(pady=(0, 15), anchor='w')
        self.format_combo.current(0)

        # 2. Загрузка исходного файла
        self.file_label_text = tk.StringVar()
        self.update_file_label_text()

        tk.Label(
            frame,
            textvariable=self.file_label_text,
            anchor='w', bg=self.bg_color, font=("Arial", 11, 'bold'),
        ).pack(fill="x", pady=(0, 2))
        tk.Button(frame, text="Выбрать исходный файл", command=self.load_file).pack(pady=(0, 8))

        self.filename_label = tk.Label(frame, text="", bg=self.bg_color, fg="grey")
        self.filename_label.pack(fill="x", pady=(0, 10))

        # 3. Загрузка справочника
        tk.Label(
            frame,
            text="3. Загрузите файл-справочник (s2_cell_id_13):",
            anchor='w', bg=self.bg_color, font=("Arial", 11, 'bold'),
        ).pack(fill="x", pady=(0, 2))
        tk.Button(frame, text="Выбрать справочник", command=self.load_match_file).pack(pady=(0, 8))

        self.match_filename_label = tk.Label(frame, text="", bg=self.bg_color, fg="grey")
        self.match_filename_label.pack(fill="x", pady=(0, 15))

        # Прогрессбар и счетчик
        self.progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=650, mode='determinate')
        self.progress.pack(pady=12)
        self.progress.pack_forget()

        self.counter_var = tk.StringVar(value="")
        tk.Label(frame, textvariable=self.counter_var, fg="#34495e", bg=self.bg_color).pack(pady=(0, 15))

        # Кнопка запуска обработки
        self.btn_process = tk.Button(
            frame, text="Начать обработку", state=tk.DISABLED, command=self.start_processing, bg='#3498db', fg='white'
        )
        self.btn_process.pack(pady=16)

        self.result = tk.Label(frame, text="", fg="#2c3e50", bg=self.bg_color, justify='left', wraplength=780)
        self.result.pack(pady=10)

        self.output_dir = os.path.expanduser(r"~/Downloads/Tile_Results")
        os.makedirs(self.output_dir, exist_ok=True)

        # Привязка обновления подписи к смене формата
        self.format_combo.bind("<<ComboboxSelected>>", lambda e: self.on_format_change())

        # Для обновления прогресса из потока
        self.total_rows = 0
        self.current_row = 0
        self.last_update_time = 0
        self.lock = threading.Lock()

    def on_format_change(self):
        self.input_format = self.format_combo.get()
        self.update_file_label_text()
        self.filename_label.config(text="")
        self.file_path = None
        self.check_all_files()

    def update_file_label_text(self):
        fmt = self.input_format_var.get()
        if fmt == 'wkt':
            text = "2. Загрузите исходный файл (обязательна колонка BS_POSITION в формате WKT):"
        else:
            text = "2. Загрузите исходный файл (обязательны колонки LATITUDE и LONGITUDE):"
        self.file_label_text.set(text)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[('Text files', '*.txt'), ('CSV files', '*.csv'), ('All files', '*.*')],
            title="Выберите исходный файл"
        )
        if file_path:
            self.file_path = file_path
            self.input_format = self.format_combo.get()
            self.filename_label.config(
                text=f"{os.path.basename(self.file_path)} (формат: {self.input_format})",
                fg="green"
            )
        self.check_all_files()

    def load_match_file(self):
        match_file_path = filedialog.askopenfilename(
            filetypes=[('Text files', '*.txt'), ('CSV files', '*.csv'), ('All files', '*.*')],
            title="Выберите справочник (s2_cell_id_13)"
        )
        if match_file_path:
            self.match_file_path = match_file_path
            self.match_filename_label.config(
                text=os.path.basename(self.match_file_path),
                fg="green"
            )
        self.check_all_files()

    def check_all_files(self):
        if self.file_path and self.match_file_path:
            self.btn_process.config(state=tk.NORMAL)
        else:
            self.btn_process.config(state=tk.DISABLED)

    def parse_position(self, position):
        try:
            if isinstance(position, float) and pd.isna(position):
                return None, None
            wkt_match = re.match(r'POINT\s*\(\s*([\d\.\-]+)\s+([\d\.\-]+)\s*\)', str(position))
            if wkt_match:
                lon = float(wkt_match.group(1))
                lat = float(wkt_match.group(2))
                return lat, lon
            else:
                raise ValueError(f"Неизвестный формат WKT: {position}")
        except Exception as e:
            print(f"Ошибка парсинга WKT: {position} --- {e}")
            return None, None

    def get_tile_id(self, lat, lon):
        try:
            if lat is None or lon is None:
                return None
            region = Cap.from_axis_angle(
                LatLng.from_degrees(lat, lon).to_point(),
                Angle.from_degrees(360 * 1 / (2 * math.pi * 6371000))
            )
            coverer = RegionCoverer()
            coverer.min_level = 13
            coverer.max_level = 13
            cells = coverer.get_covering(region)
            return str(sorted([x.id() for x in cells])[0]) if cells else None
        except Exception as e:
            print(f"Ошибка определения тайла ({lat}, {lon}) --- {e}")
            return None

    def update_progress_ui(self):
        """Обновление прогресса и счетчика в главном потоке через after()"""
        with self.lock:
            current = self.current_row
            total = self.total_rows

        if total > 0:
            self.progress['maximum'] = total
            self.progress['value'] = current
            percent = (current / total)
            self.counter_var.set(f"tile_id: {current}/{total} ({percent:.0%})")
            self.root.update_idletasks()

        if current < total:
            self.root.after(500, self.update_progress_ui)

    def process_file(self):
        """Обработка файла в отдельном потоке"""
        try:
            fmt = getattr(self, "input_format", "wkt")
            print(f"Используем формат: {fmt}")

            # Читаем файлы напрямую без кэширования
            df = pd.read_csv(self.file_path, sep=';')

            if fmt == 'wkt':
                if 'BS_POSITION' not in df.columns:
                    self.show_error_threadsafe("Нет колонки 'BS_POSITION'")
                    return None
            elif fmt == 'coords':
                if 'LATITUDE' not in df.columns or 'LONGITUDE' not in df.columns:
                    self.show_error_threadsafe("Требуются колонки 'LATITUDE' и 'LONGITUDE' для формата coords")
                    return None

                def to_wkt(lat, lon):
                    try:
                        lat_f = float(str(lat).replace(',', '.'))
                        lon_f = float(str(lon).replace(',', '.'))
                        return f"POINT ({lon_f} {lat_f})"
                    except Exception:
                        return None

                df['BS_POSITION'] = df.apply(
                    lambda row: to_wkt(row['LATITUDE'], row['LONGITUDE']), axis=1
                )
            else:
                self.show_error_threadsafe(f"Неподдерживаемый формат: {fmt}")
                return None

            with self.lock:
                self.total_rows = len(df)
                self.current_row = 0

            self.progress.pack()
            tile_ids = []

            for i, row in enumerate(df.itertuples(), 1):
                position = getattr(row, 'BS_POSITION', None)
                lat, lon = self.parse_position(position)
                tile_id = self.get_tile_id(lat, lon)
                tile_ids.append(tile_id)

                with self.lock:
                    self.current_row = i

                now = time.time()
                if now - self.last_update_time > 0.5 or i == self.total_rows:
                    self.last_update_time = now

            df['tile_id'] = tile_ids
            tile_ids_set = set(df['tile_id'].astype(str).str.strip())

            # Обработка справочника
            matches = []
            chunk_size = 100_000
            found_rows = 0
            total_rows = 0

            for chunk in pd.read_csv(self.match_file_path, sep=';', dtype=str, chunksize=chunk_size):
                chunk['s2_cell_id_13'] = chunk['s2_cell_id_13'].astype(str).str.strip()
                filtered = chunk[chunk['s2_cell_id_13'].isin(tile_ids_set)]
                found_rows += len(filtered)
                total_rows += len(chunk)
                if not filtered.empty:
                    matches.append(filtered)

            if matches:
                df2_filtered = pd.concat(matches, ignore_index=True)
            else:
                df2_filtered = pd.DataFrame(columns=['s2_cell_id_13'])

            df['tile_id'] = df['tile_id'].astype(str).str.strip()
            df2_filtered['s2_cell_id_13'] = df2_filtered['s2_cell_id_13'].astype(str).str.strip()

            merged = pd.merge(
                df,
                df2_filtered,
                how='left',
                left_on='tile_id',
                right_on='s2_cell_id_13',
                suffixes=('', '_spr')
            )

            if 'tile_id' in merged.columns:
                merged.drop(columns=['tile_id'], inplace=True)

            if fmt == 'coords':
                for col in ['LATITUDE', 'LONGITUDE']:
                    if col in merged.columns:
                        merged.drop(columns=[col], inplace=True)

            fn1 = os.path.basename(self.file_path)
            fn2 = os.path.basename(self.match_file_path)
            result_name = f"MERGED_{fn1}_BY_{fn2}"
            out_path = os.path.join(self.output_dir, result_name)
            if not out_path.lower().endswith('.xlsx'):
                out_path += '.xlsx'

            # Сохраняем в Excel
            merged.to_excel(out_path, index=False)

            # Загружаем в openpyxl для установки фильтра (автофильтр на всю таблицу)
            wb = load_workbook(out_path)
            ws = wb.active

            # Устанавливаем автофильтр на весь диапазон с данными
            ws.auto_filter.ref = ws.dimensions

            wb.save(out_path)

            return out_path, len(merged), found_rows, total_rows

        except Exception as e:
            self.show_error_threadsafe(str(e))
            return None

    def show_error_threadsafe(self, message):
        """Вывод ошибки из другого потока с применением after"""
        self.root.after(0, lambda: messagebox.showerror("Ошибка", message))

    def start_processing(self):
        self.btn_process.config(state=tk.DISABLED)
        self.counter_var.set("Обработка...")
        self.progress.pack()
        self.progress['value'] = 0
        self.result.config(text="", fg="#2c3e50")

        self.last_update_time = 0
        self.root.after(100, self.update_progress_ui)

        threading.Thread(target=self.run_processing, daemon=True).start()

    def run_processing(self):
        result = self.process_file()
        self.root.after(0, self.on_processing_finished, result)

    def on_processing_finished(self, result):
        self.btn_process.config(state=tk.NORMAL)
        self.progress.pack_forget()
        self.counter_var.set("")
        if result:
            output_path, final_count, found_rows, total_rows = result
            self.result.config(
                text=f"Готово!\n"
                     f"В объединённой выгрузке {final_count} строк.\n"
                     f"Найдено соответствий во втором файле: {found_rows} из {total_rows}.\n"
                     f"Результат сохранён в:\n{output_path}",
                fg="#27ae60"
            )
            try:
                os.startfile(self.output_dir)
            except Exception:
                pass
        else:
            self.result.config(text="Ошибка во время обработки.", fg="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = TileIntersectionApp(root)
    root.mainloop()
