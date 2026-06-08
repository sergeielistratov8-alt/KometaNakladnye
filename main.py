import sys
import json
import os
from pathlib import Path
from threading import Lock
from PyQt5.QtCore import QObject

# ГЛОБАЛЬНЫЕ ИМПОРТЫ ДЛЯ СТАТИЧЕСКИХ МЕТОДОВ
try:
    import pandas as pd
except Exception:
    pass

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except Exception:
    pass

# Определяем, запущено ли приложение через PyInstaller
is_pyinstaller = getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

# Проверка наличия критичных зависимостей
missing = []

if not is_pyinstaller:
    try:
        import pandas as pd
    except Exception as e:
        missing.append(f'pandas ({e})')
    try:
        import openpyxl
    except Exception as e:
        missing.append(f'openpyxl ({e})')

qt_backend = None
try:
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QListWidget,
        QListWidgetItem, QPushButton, QFileDialog, QLabel, QMessageBox, QProgressBar, QSplitter
    )
    from PyQt5.QtCore import Qt, pyqtSignal, QObject, QMimeData, QThread
    from PyQt5.QtGui import QIcon, QColor, QFont, QDrag
    from PyQt5.QtWidgets import QAbstractItemView
    qt_backend = 'PyQt5'
except Exception as e:
    try:
        from PySide6.QtWidgets import (
            QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QListWidget,
            QListWidgetItem, QPushButton, QFileDialog, QLabel, QMessageBox, QProgressBar, QSplitter
        )
        from PySide6.QtCore import Qt, Signal as pyqtSignal, QObject, QMimeData, QThread
        from PySide6.QtGui import QIcon, QColor, QFont, QDrag
        from PySide6.QtWidgets import QAbstractItemView
        qt_backend = 'PySide6'
    except Exception as e2:
        if not is_pyinstaller:
            missing.append(f'PyQt5/PySide6 ({e}; {e2})')
        else:
            error_details = f"Qt библиотека не найдена. Ошибки: {e} {e2}"
            print(error_details, file=sys.stderr)
            # Не прерываем выполнение при импорте модуля, чтобы можно было
            # использовать утилитарные функции (например, ExcelCleaner) в тестах.
            qt_backend = None

if missing and not is_pyinstaller:
    msg = f"Отсутствуют зависимости: {chr(10).join(missing)} Установите через: pip install -r requirements.txt"
    try:
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, msg, "Ошибка зависимостей", 0)
    except Exception:
        pass
    print(msg, file=sys.stderr)
    sys.exit(1)


class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    file_processed = pyqtSignal(str, str)
    error = pyqtSignal(str)


class FileProcessor(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    file_processed = pyqtSignal(str, str)

    def __init__(self, files: list, staging_dir: str):
        super().__init__()
        self.files = list(files)
        self.staging_dir = staging_dir
        self.stop_requested = False

    def process_files(self):
        done = 0
        for file_path in list(self.files):
            if self.stop_requested:
                break
            try:
                file_p = Path(file_path)
                file_name = file_p.name

                # Читаем и чистим исходник один раз, сохраняем в .xls.
                combined, error_msg = ExcelCleaner.build_cleaned_dataframe(str(file_p))
                if combined is None:
                    self.file_processed.emit(file_name, f'❌ {error_msg}')
                else:
                    for fmt in ExcelCleaner.EXPORT_FORMATS:
                        dest_name = file_p.stem + fmt
                        dest_path = Path(self.staging_dir) / dest_name
                        try:
                            ExcelCleaner.save_dataframe(combined, str(dest_path))
                            self.file_processed.emit(dest_name, '✅ Обработан')
                        except Exception as e:
                            self.file_processed.emit(dest_name, f'❌ {str(e)}')
            except Exception as e:
                self.file_processed.emit(Path(file_path).name, f'❌ {str(e)}')
            done += 1
            self.progress.emit(done)
        self.finished.emit()


class ExcelCleaner:
    HEADER_KEYWORDS = {
        'Артикул': ['артикул', 'код', 'sku', 'артик', 'артик.'],
        'Товар': ['товар', 'наименование', 'описание', 'наим', 'предмет'],
        'Кол-во': ['кол-во', 'количество', 'кол ', 'qty', 'шт', 'штук', 'кол.'],
        'Цена': ['цена', 'price', 'стоимость', 'цена за ед', 'цена за шт'],
        'Сумма': ['сумма', 'итого', 'total', 'стоимость', 'суммарно']
    }
    TARGET_COLUMNS = ['Артикул', 'Товар', 'Кол-во', 'Цена', 'Сумма']
    FOOTER_KEYWORDS = ['итого', 'итого:', 'в том числе', 'ндс', 'налог']
    # Колонки, которые при экспорте принудительно приводятся к числу в Excel.
    NUMERIC_EXPORT_COLUMNS = ['Артикул', 'Кол-во', 'Цена', 'Сумма']
    INTEGER_EXPORT_COLUMNS = ['Артикул', 'Кол-во']
    EXPORT_FORMATS = ['.xls']

    @staticmethod
    def is_processable_filename(name: str) -> bool:
        return not name.startswith('._') and not name.startswith('~$')

    @staticmethod
    def sanitize_sheet_name(name: str) -> str:
        import re
        cleaned = re.sub(r'[\\/*?:\[\]]', '_', str(name).strip())
        return (cleaned[:31] if cleaned else 'Sheet')

    @staticmethod
    def read_csv_file(file_path: str):
        import pandas as pd
        encodings = ('utf-8-sig', 'utf-8', 'cp1251')
        last_error = None
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    sample = f.read(4096)
                if not sample.strip():
                    return pd.DataFrame()
                sep = ';' if sample.count(';') >= sample.count(',') else ','
                return pd.read_csv(
                    file_path, header=None, dtype=str, keep_default_na=False,
                    encoding=encoding, sep=sep,
                )
            except Exception as e:
                last_error = e
        raise last_error or ValueError('Не удалось прочитать CSV')

    @staticmethod
    def normalize_header_value(value: str) -> str:
        if value is None:
            return ''
        text = str(value).strip().lower()
        for canonical, patterns in ExcelCleaner.HEADER_KEYWORDS.items():
            for pattern in patterns:
                if pattern in text:
                    return canonical
        return text.title() if text else ''

    @staticmethod
    def is_numeric_text(value: str) -> bool:
        if value is None:
            return False
        text = str(value).strip().replace(' ', '').replace(',', '.').replace('+', '').replace('-', '')
        return bool(text) and text.replace('.', '').isdigit()

    @staticmethod
    def is_numeric_row(row_values) -> bool:
        text = ' '.join(str(value).strip() for value in row_values if value is not None and str(value).strip())
        if not text:
            return False
        normalized = text.replace(' ', '').replace(',', '.').replace('+', '').replace('-', '')
        return normalized.replace('.', '').isdigit()

    @staticmethod
    def remove_footer_rows(df):
        import pandas as pd
        cleaned = df.copy()
        while len(cleaned) > 0 and cleaned.iloc[-1].replace('', pd.NA).isna().all():
            cleaned = cleaned.iloc[:-1]
        if len(cleaned) == 0:
            return cleaned

        footer_index = None
        for idx in range(len(cleaned) - 1, -1, -1):
            row_text = ' '.join(str(value).strip().lower() for value in cleaned.iloc[idx].values if value is not None and str(value).strip())
            if any(keyword in row_text for keyword in ExcelCleaner.FOOTER_KEYWORDS):
                footer_index = idx
                break

        if footer_index is not None:
            return cleaned.iloc[:footer_index].copy()

        if len(cleaned) >= 2:
            prev_text = ' '.join(str(value).strip().lower() for value in cleaned.iloc[-2].values if value is not None and str(value).strip())
            if any(keyword in prev_text for keyword in ExcelCleaner.FOOTER_KEYWORDS) and ExcelCleaner.is_numeric_row(cleaned.iloc[-1].values):
                return cleaned.iloc[:-2].copy()
        return cleaned

    @staticmethod
    def drop_footer_rows(cleaned):
        def is_footer_row(row):
            text = ' '.join(str(value).strip().lower() for value in row.values if value is not None and str(value).strip())
            return any(keyword in text for keyword in ExcelCleaner.FOOTER_KEYWORDS)
        return cleaned.loc[~cleaned.apply(is_footer_row, axis=1)].copy()

    @staticmethod
    def find_header_row(df) -> int:
        import pandas as pd
        for idx, row in df.iterrows():
            found = set()
            for value in row.values:
                if pd.isna(value):
                    continue
                text = str(value).strip().lower()
                for canonical, patterns in ExcelCleaner.HEADER_KEYWORDS.items():
                    if any(pattern in text for pattern in patterns):
                        found.add(canonical)
            if len(found) >= 2 and any(col in found for col in ['Товар', 'Цена', 'Сумма']):
                return idx
        return -1

    @staticmethod
    def looks_like_item_table(df) -> bool:
        if df.shape[1] < 5 or len(df) == 0:
            return False
        sample = df.iloc[:min(3, len(df))]
        numeric_rows = 0
        for _, row in sample.iterrows():
            if ExcelCleaner.is_numeric_text(row.iloc[2]) and ExcelCleaner.is_numeric_text(row.iloc[3]):
                numeric_rows += 1
        return numeric_rows >= 1

    @staticmethod
    def extract_positional_table(df):
        import pandas as pd
        if not ExcelCleaner.looks_like_item_table(df):
            return pd.DataFrame(columns=ExcelCleaner.TARGET_COLUMNS)
        cleaned = ExcelCleaner.remove_footer_rows(df.copy())
        cleaned = cleaned.iloc[:, :5]
        cleaned.columns = ExcelCleaner.TARGET_COLUMNS
        cleaned = cleaned.replace(r'^\s*$', pd.NA, regex=True)
        cleaned = cleaned.dropna(axis=0, how='all', subset=ExcelCleaner.TARGET_COLUMNS).fillna('')
        return cleaned

    @staticmethod
    def clean_dataframe(df):
        import pandas as pd
        df = df.fillna('').astype(str)
        if df.empty:
            return pd.DataFrame(columns=ExcelCleaner.TARGET_COLUMNS)

        header_index = ExcelCleaner.find_header_row(df)
        if header_index >= 0:
            header_values = df.loc[header_index].astype(str).tolist()
            cleaned = df.loc[header_index + 1:].copy().reset_index(drop=True)
            cleaned.columns = [ExcelCleaner.normalize_header_value(col) for col in header_values]
            cleaned = cleaned.loc[:, cleaned.columns.notna()]

            for canonical in ExcelCleaner.TARGET_COLUMNS:
                if canonical not in cleaned.columns:
                    cleaned[canonical] = ''
            cleaned = cleaned[ExcelCleaner.TARGET_COLUMNS]
            cleaned = cleaned.replace(r'^\s*$', pd.NA, regex=True)
            cleaned = cleaned.dropna(axis=0, how='all', subset=ExcelCleaner.TARGET_COLUMNS)
            cleaned = ExcelCleaner.remove_footer_rows(cleaned)
            cleaned = ExcelCleaner.drop_footer_rows(cleaned)
            cleaned = cleaned.fillna('')
            return cleaned

        positional = ExcelCleaner.extract_positional_table(df)
        if not positional.empty:
            return positional

        cleaned = df.replace(r'^\s*$', pd.NA, regex=True).dropna(axis=0, how='all').dropna(axis=1, how='all')
        cleaned.columns = [ExcelCleaner.normalize_header_value(col) for col in cleaned.columns]
        for canonical in ExcelCleaner.TARGET_COLUMNS:
            if canonical not in cleaned.columns:
                cleaned[canonical] = ''
        cleaned = cleaned[ExcelCleaner.TARGET_COLUMNS]
        cleaned = cleaned.replace(r'^\s*$', pd.NA, regex=True).dropna(axis=0, how='all', subset=ExcelCleaner.TARGET_COLUMNS)
        cleaned = ExcelCleaner.remove_footer_rows(cleaned)
        cleaned = ExcelCleaner.drop_footer_rows(cleaned)
        cleaned = cleaned.fillna('')
        return cleaned

    @staticmethod
    def convert_numeric_columns(df, columns=None):
        import pandas as pd
        if columns is None:
            columns = ['Кол-во', 'Цена', 'Сумма', 'Артикул']
        cols = [c for c in columns if c in df.columns]
        for col in cols:
            try:
                s = df[col].astype(str).str.strip().replace('', pd.NA)
                s = s.str.replace(' ', '')
                # Обрабатывать разные форматы: если в строке есть и ',' и '.',
                # то скорее всего ',' — разделитель тысяч (удаляем), '.' — десятичный.
                def _normalize(x):
                    if x is None:
                        return x
                    xs = str(x)
                    if ',' in xs and '.' in xs:
                        return xs.replace(',', '')
                    if ',' in xs:
                        return xs.replace(',', '.')
                    return xs

                s = s.apply(_normalize)
                conv = pd.to_numeric(s, errors='coerce')
                if conv.notna().any():
                    mask = conv.notna()
                    # Переводим столбец в object, чтобы можно было записать числа
                    df[col] = df[col].astype(object)
                    df.loc[mask, col] = conv[mask]
            except Exception:
                pass
        return df

    # Невидимые символы, которые остаются после копирования из 1С/Excel и
    # ломают импорт (распознавание чисел): неразрывные пробелы, узкий
    # неразрывный пробел, символы нулевой ширины, BOM.
    INVISIBLE_CHARS = ['\u00a0', '\u202f', '\u200b', '\u200c', '\u200d', '\ufeff']

    @staticmethod
    def prepare_for_export(df):
        """Готовит DataFrame к сохранению так, чтобы импорт точно увидел числа.

        1) Во ВСЕХ текстовых ячейках удаляет невидимые символы (\\xa0 и пр.)
           и крайние пробелы (внутренние пробелы в тексте сохраняются).
        2) Для числовых колонок ('Артикул', 'Кол-во', 'Цена', 'Сумма') приводит
           значения к числу, чтобы Excel видел тип Number, а не Text.
        """
        import pandas as pd

        export_df = df.copy()

        def _strip_invisible(text: str) -> str:
            for ch in ExcelCleaner.INVISIBLE_CHARS:
                text = text.replace(ch, '')
            return text

        # 1) Общая чистка невидимых символов во всех нечисловых колонках.
        #    Проверяем по «не числовая», а не по dtype == object, чтобы
        #    работало и в новых версиях pandas (строковый dtype 'str').
        for col in export_df.columns:
            if not pd.api.types.is_numeric_dtype(export_df[col]):
                s = export_df[col].astype(str).map(_strip_invisible)
                export_df[col] = s.str.strip()

        # 2) Числовые колонки -> настоящий тип float.
        def _normalize(value):
            text = _strip_invisible(str(value)).strip()
            if text in ('', 'nan', 'none', 'None', '<NA>'):
                return ''
            # Удаляем обычные пробелы (разделители разрядов).
            text = text.replace(' ', '')
            if ',' in text and '.' in text:
                # И ',' и '.' -> ',' считаем разделителем тысяч.
                return text.replace(',', '')
            if ',' in text:
                # Только ',' -> десятичный разделитель.
                return text.replace(',', '.')
            return text

        for col in ExcelCleaner.NUMERIC_EXPORT_COLUMNS:
            if col not in export_df.columns:
                continue
            normalized = export_df[col].map(_normalize)
            if col == 'Артикул':
                def _artikul(value, raw_text):
                    if raw_text == '':
                        return ''
                    num = pd.to_numeric(raw_text, errors='coerce')
                    if pd.isna(num):
                        return _strip_invisible(str(value)).strip()
                    return int(num) if num == int(num) else float(num)
                export_df[col] = [
                    _artikul(v, n) for v, n in zip(export_df[col], normalized)
                ]
            else:
                export_df[col] = pd.to_numeric(normalized, errors='coerce')

        return export_df

    @staticmethod
    def coerce_to_excel_value(value, col_name: str):
        """Возвращает значение для xlwt: число (int/float) или текст."""
        import pandas as pd
        import math

        if col_name == 'Товар':
            return str(value).strip() if value not in (None, '') else '', 'text'

        if value is None or value == '':
            return '', 'empty'
        if isinstance(value, float) and pd.isna(value):
            return '', 'empty'

        if isinstance(value, int):
            return value, 'int'
        if isinstance(value, float):
            if col_name in ExcelCleaner.INTEGER_EXPORT_COLUMNS and value == int(value):
                return int(value), 'int'
            return value, 'float'

        text = str(value).strip()
        for ch in ExcelCleaner.INVISIBLE_CHARS:
            text = text.replace(ch, '')
        text = text.strip().replace(' ', '')
        if not text or text.lower() in ('nan', 'none', '<na>'):
            return '', 'empty'

        if ',' in text and '.' in text:
            text = text.replace(',', '')
        elif ',' in text:
            text = text.replace(',', '.')

        try:
            num = float(text)
            if math.isfinite(num):
                if col_name in ExcelCleaner.INTEGER_EXPORT_COLUMNS and num == int(num):
                    return int(num), 'int'
                return num, 'float'
        except ValueError:
            pass

        return str(value).strip(), 'text'

    @staticmethod
    def save_dataframe(df, dest_path: str):
        """Сохраняет DataFrame в .xls с жирными рамками и сеткой ячеек."""
        ext = Path(dest_path).suffix.lower()
        export_df = ExcelCleaner.prepare_for_export(df)

        if ext == '.xls':
            ExcelCleaner.save_xls(export_df, dest_path)
        else:
            raise ValueError(f"Неподдерживаемый формат сохранения: {ext}")

    @staticmethod
    def save_xls(export_df, dest_path: str):
        import xlwt

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Sheet1')

        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THICK
        borders.right = xlwt.Borders.THICK
        borders.top = xlwt.Borders.THICK
        borders.bottom = xlwt.Borders.THICK

        def _make_style(wrap_text=False, num_format=None):
            style = xlwt.XFStyle()
            style.borders = borders
            alignment = xlwt.Alignment()
            alignment.horz = xlwt.Alignment.HORZ_CENTER
            alignment.vert = xlwt.Alignment.VERT_CENTER
            alignment.wrap = 1 if wrap_text else 0
            style.alignment = alignment
            if num_format:
                style.num_format_str = num_format
            return style

        int_style = _make_style(num_format='0')
        decimal_style = _make_style(num_format='0.00')
        text_style = _make_style()
        wrap_style = _make_style(wrap_text=True)

        for col_idx, width in enumerate((12, 45, 8, 12, 12)):
            ws.col(col_idx).width = 256 * width

        for row_idx in range(len(export_df)):
            for col_idx, col_name in enumerate(export_df.columns):
                value = export_df.iloc[row_idx, col_idx]
                cell_value, kind = ExcelCleaner.coerce_to_excel_value(value, col_name)

                if kind == 'empty':
                    style = wrap_style if col_name == 'Товар' else int_style
                    ws.write(row_idx, col_idx, '', style)
                elif kind == 'text':
                    style = wrap_style if col_name == 'Товар' else text_style
                    ws.write(row_idx, col_idx, cell_value, style)
                elif kind == 'int':
                    ws.write(row_idx, col_idx, cell_value, int_style)
                else:
                    ws.write(row_idx, col_idx, cell_value, decimal_style)

        wb.save(dest_path)

    @staticmethod
    def build_cleaned_dataframe(file_path: str):
        """Читает исходник и возвращает единый очищенный DataFrame.

        Возвращает (df, ''). При неподдерживаемом формате -> (None, текст).
        Парсинг вынесен отдельно, чтобы прочитать файл один раз и затем
        сохранить результат сразу в несколько форматов.
        """
        import pandas as pd
        p = Path(file_path)
        ext = p.suffix.lower()

        frames = []
        if ext == '.csv':
            raw = ExcelCleaner.read_csv_file(file_path)
            raw = raw.fillna('').astype(str)
            cleaned = ExcelCleaner.clean_dataframe(raw)
            if not cleaned.empty:
                frames.append(cleaned)
        elif ext in ('.xls', '.xlsx'):
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            sheets = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str, engine=engine)
            for raw in (sheets or {}).values():
                if raw is None or raw.empty:
                    continue
                raw = raw.fillna('').astype(str)
                cleaned = ExcelCleaner.clean_dataframe(raw)
                if not cleaned.empty:
                    frames.append(cleaned)
        else:
            return None, f"Неподдерживаемый формат: {ext}"

        if frames:
            combined = pd.concat(frames, ignore_index=True)
        else:
            combined = pd.DataFrame(columns=ExcelCleaner.TARGET_COLUMNS)
        return combined, ''

    @staticmethod
    def remove_empty_rows_and_columns(file_path: str, dest_path: str) -> tuple[bool, str]:
        try:
            combined, error = ExcelCleaner.build_cleaned_dataframe(file_path)
            if combined is None:
                return False, error
            ExcelCleaner.save_dataframe(combined, dest_path)
            return True, ''
        except Exception as e:
            return False, f"Ошибка структуры: {str(e)}"


class ExcelCleanerUI(QMainWindow):
    def __init__(self):
        super().__init__()
        user_home = Path(os.environ.get('USERPROFILE', os.path.expanduser('~')))

        self.config_file = user_home / ".kometa_config.json"
        self.ready_dir = user_home / "KometaReady"
        self.ready_dir.mkdir(exist_ok=True)
        self.staging_dir = user_home / "KometaStaging"
        self.staging_dir.mkdir(exist_ok=True)

        self.config = self.load_config()
        self.processing = False
        self.files_to_process = []
        self.processed_files = []
        self.worker_thread = None
        self.worker = None
        self.worker_lock = Lock()
        self.signals = WorkerSignals()
        self.signals.file_processed.connect(self.on_file_processed)
        self.signals.progress.connect(self.on_progress)
        self.signals.finished.connect(self.on_finished)

        self.init_ui()
        self.setWindowTitle("Комета - Накладные")
        self.setGeometry(100, 100, 1000, 600)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        folders_layout = QHBoxLayout()

        self.source_btn = QPushButton("📂 Выбрать папку источника")
        self.source_btn.clicked.connect(self.select_source_folder)
        folders_layout.addWidget(self.source_btn)

        self.dest_btn = QPushButton("📂 Выбрать папку назначения")
        self.dest_btn.clicked.connect(self.select_dest_folder)
        folders_layout.addWidget(self.dest_btn)
        main_layout.addLayout(folders_layout)

        info_layout = QHBoxLayout()
        self.source_label = QLabel("Источник: не выбрана")
        self.dest_label = QLabel("Назначение: не выбрана")
        info_layout.addWidget(self.source_label)
        info_layout.addWidget(self.dest_label)
        main_layout.addLayout(info_layout)

        columns_layout = QHBoxLayout()
        left_layout = QVBoxLayout()
        left_label = QLabel("📥 Входные файлы")
        left_label.setFont(QFont("Arial", 11, QFont.Bold))
        left_layout.addWidget(left_label)

        self.input_list = QListWidget()
        self.input_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.input_list.setStyleSheet("QListWidget { border: 1px solid #ddd; border-radius: 3px; background-color: #2b2b2b; color: #ffffff; outline: none; font-weight: 600; } QListWidget::item { padding: 4px; border-bottom: 1px solid #444444; } QListWidget::item:hover { background-color: #3a3a3a; } QListWidget::item:selected { background-color: #5a9bd6; color: #ffffff; }")
        self.setup_drag_drop(self.input_list)
        left_layout.addWidget(self.input_list)

        input_btn_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("➕ Добавить файлы")
        self.add_files_btn.clicked.connect(self.add_input_files)
        self.remove_selected_btn = QPushButton("❌ Удалить выбранные")
        self.remove_selected_btn.clicked.connect(self.remove_selected_files)
        input_btn_layout.addWidget(self.add_files_btn)
        input_btn_layout.addWidget(self.remove_selected_btn)
        left_layout.addLayout(input_btn_layout)

        right_layout = QVBoxLayout()
        right_label = QLabel("✅ Готовые файлы")
        right_label.setFont(QFont("Arial", 11, QFont.Bold))
        right_layout.addWidget(right_label)

        self.output_list = QListWidget()
        self.output_list.setStyleSheet("QListWidget { border: 1px solid #ddd; border-radius: 3px; background-color: #2b2b2b; color: #ffffff; outline: none; font-weight: 600; } QListWidget::item { padding: 4px; border-bottom: 1px solid #444444; }")
        right_layout.addWidget(self.output_list)

        self.clear_output_btn = QPushButton("🗑️ Очистить список")
        self.clear_output_btn.clicked.connect(self.clear_output_list)
        right_layout.addWidget(self.clear_output_btn)

        columns_layout.addLayout(left_layout, 1)
        columns_layout.addLayout(right_layout, 1)
        main_layout.addLayout(columns_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.process_btn = QPushButton("🚀 Подготовить файлы")
        self.process_btn.setFont(QFont("Arial", 12, QFont.Bold))
        self.process_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; padding: 10px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: #45a049; }")
        self.process_btn.clicked.connect(self.start_processing)
        main_layout.addWidget(self.process_btn)

        self.unload_btn = QPushButton("📤 Выгрузить готовые")
        self.unload_btn.setFont(QFont("Arial", 11, QFont.Bold))
        self.unload_btn.setStyleSheet("background-color: #1976D2; color: white; padding: 8px; border-radius: 4px;")
        self.unload_btn.clicked.connect(self.unload_ready)
        self.unload_btn.setEnabled(False)
        main_layout.addWidget(self.unload_btn)

        central_widget.setLayout(main_layout)
        self.update_folder_labels()

    def setup_drag_drop(self, widget):
        widget.setAcceptDrops(True)
        widget.dragEnterEvent = self.drag_enter_event
        widget.dropEvent = self.drop_event

    def drag_enter_event(self, event):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def drop_event(self, event):
        files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            p = Path(path)
            if p.suffix.lower() in ['.xlsx', '.xls', '.csv'] and ExcelCleaner.is_processable_filename(p.name):
                files.append(path)
        if files:
            self.add_files_to_input(files)
            event.accept()

    def load_config(self) -> dict:
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f: return json.load(f)
            except: pass
        return {"source_folder": "", "dest_folder": ""}

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("Ошибка сохранения конфига:", e)

    def select_source_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с входными файлами", self.config.get("source_folder", str(Path.home())))
        if folder:
            self.config["source_folder"] = folder
            self.save_config()
            self.update_folder_labels()
            self.scan_source_folder()

    def select_dest_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для готовых файлов", self.config.get("dest_folder", str(self.ready_dir)))
        if folder:
            self.config["dest_folder"] = folder
            self.save_config()
            self.update_folder_labels()

    def update_folder_labels(self):
        source = self.config.get("source_folder", "")
        dest = self.config.get("dest_folder", "")
        self.source_label.setText(f"📁 Источник: {Path(source).name if source else 'не выбрана'}")
        self.dest_label.setText(f"📁 Назначение: {Path(dest).name if dest else 'папка KometaReady'}")
        if not dest: self.config["dest_folder"] = str(self.ready_dir)

    def scan_source_folder(self):
        source = self.config.get("source_folder", "")
        if not source or not Path(source).exists(): return
        self.input_list.clear()
        self.files_to_process = []
        for file_path in Path(source).iterdir():
            if file_path.is_file() and file_path.suffix.lower() in ['.xlsx', '.xls', '.csv']:
                if ExcelCleaner.is_processable_filename(file_path.name):
                    self.files_to_process.append(str(file_path))
        self.update_input_list()

    def update_input_list(self):
        self.input_list.setUpdatesEnabled(False)
        self.input_list.clear()
        for idx, file_path in enumerate(self.files_to_process, 1):
            item = QListWidgetItem(f"{idx}. {Path(file_path).name}")
            item.setData(Qt.UserRole, file_path)
            self.input_list.addItem(item)
        self.input_list.setUpdatesEnabled(True)

    def add_input_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Выберите Excel-файлы", self.config.get("source_folder", str(Path.home())), "Excel файлы (*.xlsx *.XLSX *.xls *.XLS *.csv *.CSV)")
        if files: self.add_files_to_input(files)

    def add_files_to_input(self, files: list):
        for file_path in files:
            if not ExcelCleaner.is_processable_filename(Path(file_path).name):
                continue
            if file_path not in self.files_to_process:
                self.files_to_process.append(file_path)
        self.update_input_list()

    def remove_selected_files(self):
        for item in reversed(self.input_list.selectedItems()):
            file_path = item.data(Qt.UserRole)
            if file_path in self.files_to_process: self.files_to_process.remove(file_path)
        self.update_input_list()

    def clear_output_list(self):
        self.output_list.clear()
        self.processed_files = []

    def unload_ready(self):
        dest_folder = Path(self.config.get("dest_folder", str(self.ready_dir)))
        if not dest_folder.exists():
            return
        import shutil
        moved = 0
        errors = []
        for fname in list(self.processed_files):
            src = self.staging_dir / fname
            dest = dest_folder / fname
            if not src.exists():
                errors.append(f"{fname}: файл не найден")
                continue
            try:
                if dest.exists():
                    dest.unlink()
                shutil.move(str(src), str(dest))
                moved += 1
            except Exception as e:
                errors.append(f"{fname}: {e}")
        self.files_to_process, self.processed_files = [], []
        self.input_list.clear()
        self.output_list.clear()
        msg = f"Перемещено {moved} файлов."
        if errors:
            msg += "\n\nОшибки:\n" + "\n".join(errors)
        QMessageBox.information(self, "Выгружено", msg)

    def start_processing(self):
        if not self.files_to_process: return
        self.processing = True
        self.process_btn.setEnabled(False)
        self.unload_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(self.files_to_process))

        self.worker = FileProcessor(list(self.files_to_process), str(self.staging_dir))
        self.worker_thread = QThread()
        self.worker.moveToThread(self.worker_thread)
        self.worker.progress.connect(self.on_progress)
        self.worker.file_processed.connect(self.on_file_processed)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker_thread.started.connect(self.worker.process_files)
        self.worker_thread.finished.connect(self.cleanup_worker)
        self.worker_thread.start()
        self.files_to_process = []
        self.update_input_list()

    def cleanup_worker(self):
        if self.worker_thread is not None: self.worker_thread.wait(1000)
        self.worker, self.worker_thread = None, None

    def on_file_processed(self, file_name: str, status: str):
        self.processed_files.append(file_name)
        self.output_list.addItem(f"{len(self.processed_files)}. {file_name} {status}")

    def on_progress(self, value: int): self.progress_bar.setValue(value)
    def on_finished(self):
        self.processing = False
        self.process_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, "Готово", "Все файлы обработаны!")

def main():
    app = QApplication(sys.argv)
    window = ExcelCleanerUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
