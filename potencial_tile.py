import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
import threading
import time
from openpyxl import load_workbook
from s2sphere import CellId, LatLng

WKT_POINT_RE = re.compile(r"POINT\s*\(\s*([\d.\-]+)\s+([\d.\-]+)\s*\)")


class TileIntersectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Определение тайлов")
        self.root.geometry("830x560")
        self.root.resizable(False, False)
        self.root.configure(bg="#f0f2f5")

        self.file_path = None
        self.match_file_path = None
        self.input_format = 'wkt'

        # Цвета и шрифты
        self.bg_color = "#f0f2f5"
        self.frame_bg = "white"
        self.font_bold = ("Segoe UI", 11, "bold")
        self.font_normal = ("Segoe UI", 10)
        self.accent_color = "#0078D7"  # Синий Microsoft-style
        self.success_color = "#28a745"
        self.error_color = "#c0392b"
        self.grey_text = "#454649"

        frame = tk.Frame(self.root, bg=self.bg_color)
        frame.pack(fill="both", expand=True, padx=15, pady=15)

        # --- Блок выбора формата ---
        format_frame = tk.LabelFrame(frame, text="1. Выберите формат входных данных:",
                                    bg=self.frame_bg, fg=self.accent_color, font=self.font_bold)
        format_frame.pack(fill="x", pady=(0, 15))

        self.input_format_var = tk.StringVar(value='wkt')
        format_options = ['wkt', 'coords']
        self.format_combo = ttk.Combobox(format_frame, textvariable=self.input_format_var, values=format_options,
                                        state="readonly", width=17, font=self.font_normal)
        self.format_combo.pack(padx=10, pady=8, anchor='w')
        self.format_combo.current(0)
        self.format_combo.bind("<<ComboboxSelected>>", lambda e: self.on_format_change())

        # --- Загрузка исходного файла ---
        input_file_frame = tk.LabelFrame(frame, text="2. Загрузка исходного файла",
                                        bg=self.frame_bg, fg=self.accent_color, font=self.font_bold)
        input_file_frame.pack(fill="x", pady=(0, 15))

        self.file_label_text = tk.StringVar()
        self.update_file_label_text()
        label_file_desc = tk.Label(input_file_frame, textvariable=self.file_label_text, bg=self.frame_bg,
                                fg="black", font=self.font_normal, anchor="w")
        label_file_desc.pack(fill="x", padx=10, pady=(10, 3))

        btn_file = tk.Button(input_file_frame, text="Выбрать исходный файл", command=self.load_file,
                            bg=self.accent_color, fg="white", font=self.font_bold,
                            activebackground="#005a9e", cursor="hand2", relief="flat", padx=15, pady=5)
        btn_file.pack(padx=10, pady=(0, 10), anchor='w')
        btn_file.bind("<Enter>", lambda e: btn_file.config(bg="#005a9e"))
        btn_file.bind("<Leave>", lambda e: btn_file.config(bg=self.accent_color))

        self.filename_label = tk.Label(input_file_frame, text="", bg=self.frame_bg, fg=self.grey_text,
                                    font=("Segoe UI", 9, "italic"), anchor="w")
        self.filename_label.pack(fill="x", padx=10, pady=(0, 10))

        # --- Загрузка справочника ---
        ref_file_frame = tk.LabelFrame(frame, text="3. Загрузка файла-справочника (s2_cell_id_13):",
                                    bg=self.frame_bg, fg=self.accent_color, font=self.font_bold)
        ref_file_frame.pack(fill="x", pady=(0, 15))

        btn_ref = tk.Button(ref_file_frame, text="Выбрать справочник", command=self.load_match_file,
                            bg=self.accent_color, fg="white", font=self.font_bold,
                            activebackground="#005a9e", cursor="hand2", relief="flat", padx=15, pady=5)
        btn_ref.pack(padx=10, pady=(10, 8), anchor='w')
        btn_ref.bind("<Enter>", lambda e: btn_ref.config(bg="#005a9e"))
        btn_ref.bind("<Leave>", lambda e: btn_ref.config(bg=self.accent_color))

        self.match_filename_label = tk.Label(ref_file_frame, text="", bg=self.frame_bg, fg=self.grey_text,
                                            font=("Segoe UI", 9, "italic"), anchor="w")
        self.match_filename_label.pack(fill="x", padx=10, pady=(0, 10))

        # --- Прогресс ---
        progress_frame = tk.Frame(frame, bg=self.bg_color)
        progress_frame.pack(fill="x", pady=(10, 10))

        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=680, mode='determinate')
        self.progress.pack(side="left", padx=(0, 10), pady=5)

        self.counter_var = tk.StringVar(value="")
        self.counter_label = tk.Label(progress_frame, textvariable=self.counter_var, fg="#34495e",
                                    bg=self.bg_color, font=self.font_normal)
        self.counter_label.pack(side="left", pady=5)

        # --- Кнопка запуска обработки ---
        self.btn_process = tk.Button(
            frame, text="Начать обработку", state=tk.DISABLED, command=self.start_processing,
            bg=self.accent_color, fg='white', font=self.font_bold, activebackground="#005a9e",
            cursor="hand2", relief="flat", padx=25, pady=8
        )
        self.btn_process.pack(pady=20)
        # Обеспечиваем белый цвет текста при наведении и уходе мыши
        self.btn_process.bind("<Enter>", lambda e: self.btn_process.config(bg="#005a9e", fg='white'))
        self.btn_process.bind("<Leave>", lambda e: self.btn_process.config(bg=self.accent_color, fg='white'))

        self.result = tk.Label(frame, text="", fg="#2c3e50", bg=self.bg_color,
                            justify='left', wraplength=780, font=self.font_normal)
        self.result.pack(pady=10)

        self.output_dir = os.path.expanduser(r"~/Downloads/Tile_Results")
        os.makedirs(self.output_dir, exist_ok=True)

        # Для обновления прогресса из потока
        self.total_rows = 0
        self.current_row = 0
        self.last_update_time = 0
        self.lock = threading.Lock()

        # Столбцы, которые нужно тянуть из справочника (добавлен operator_name)
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
            "operator_name",
        ]

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
            wkt_match = WKT_POINT_RE.match(str(position))
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
            return str(CellId.from_lat_lng(LatLng.from_degrees(lat, lon)).parent(13).id())
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

            # Читаем исходный файл
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

            # Обработка справочника с фильтрацией по оператору Tele2
            matches = []
            chunk_size = 100_000
            found_rows = 0
            total_rows = 0

            for chunk in pd.read_csv(self.match_file_path, sep=';', dtype=str, chunksize=chunk_size, usecols=self.columns_needed):
                chunk['s2_cell_id_13'] = chunk['s2_cell_id_13'].astype(str).str.strip()
                # Фильтрация по tile_id и operator_name == "Tele2"
                filtered = chunk[(chunk['s2_cell_id_13'].isin(tile_ids_set)) & (chunk['operator_name'] == "Tele2")]
                found_rows += len(filtered)
                total_rows += len(chunk)
                if not filtered.empty:
                    matches.append(filtered)

            if matches:
                df2_filtered = pd.concat(matches, ignore_index=True)
            else:
                df2_filtered = pd.DataFrame(columns=self.columns_needed)

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
                fg=self.success_color
            )
            try:
                os.startfile(self.output_dir)
            except Exception:
                pass
        else:
            self.result.config(text="Ошибка во время обработки.", fg=self.error_color)


if __name__ == "__main__":
    root = tk.Tk()
    app = TileIntersectionApp(root)
    root.mainloop()
