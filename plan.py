Лицензия MIT

Copyright (c) 2026 Культяпов Андрей (Camellan)

Предоставляется разрешение бесплатно и без ограничений любому лицу, получившему копию данного программного обеспечения и сопутствующей документации (далее «Программное обеспечение»), на действия с Программным обеспечением без ограничений, включая, но не ограничиваясь, права на использование, копирование, модификацию, объединение, публикацию, распространение, сублицензирование и/или продажу копий Программного обеспечения, а также на предоставление права другим лицам на выполнение этих действий, при соблюдении следующих условий:

Вышеуказанное уведомление об авторских правах и данное уведомление о разрешении должны быть включены во все копии или существенные части Программного обеспечения.

ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ ПРЕДОСТАВЛЯЕТСЯ «КАК ЕСТЬ», БЕЗ КАКИХ-ЛИБО ГАРАНТИЙ, ЯВНЫХ ИЛИ ПОДРАЗУМЕВАЕМЫХ, ВКЛЮЧАЯ, НО НЕ ОГРАНИЧИВАЯСЬ ГАРАНТИЯМИ КОМЕРЧЕСКОЙ ЦЕННОСТИ И ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЁННЫХ ЦЕЛЕЙ. АВТОРЫ И ПРАВООБЛАДАТЕЛИ НЕ НЕСУТ ОТВЕТСТВЕННОСТИ ЗА КАКИЕ-ЛИБО ИСКИ, УЩЕРБ ИЛИ ДРУГИЕ ТРЕБОВАНИЯ, ВОЗНИКШИЕ ИЗ ИЛИ В СВЯЗИ С ИСПОЛЬЗОВАНИЕМ ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ ИЛИ ЕГО ИСПОЛЬЗОВАНИЕМ ИЛИ ИНЫМИ ДЕЙСТВИЯМИ.

import sys
import os
import re
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# Подключаем ctypes для принудительного включения иконки в Windows
if sys.platform == "win32":
    import ctypes
    myappid = 'psk.iskra.vinanalyzer.1.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QLabel,
    QLineEdit,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox
)
from PyQt6.QtCore import Qt, QLocale, QTranslator
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIntValidator, QIcon

EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
REQUIRED_COLUMNS = {"модель", "vin", "доп версия"}


def get_resource_path(relative_path: str) -> str:
    """Возвращает путь к ресурсу для запуска из .py и собранного .exe."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def extract_date_from_filename(file_path: Path) -> Optional[date]:
    """Извлекает дату ДД.ММ.ГГГГ или ДД,ММ,ГГГГ из имени файла."""
    match = re.search(r"(?<!\d)(\d{1,2})[.,](\d{1,2})[.,](\d{4})(?!\d)", file_path.stem)
    if not match:
        return None
    day, month, year = map(int, match.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def file_sort_key(file_path: Path):
    """Сначала файлы с датой по возрастанию, затем файлы без распознанной даты."""
    file_date = extract_date_from_filename(file_path)
    return (file_date is None, file_date or date.max, file_path.name.casefold())


class VinAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(get_resource_path("lada.ico")))
        self.setWindowTitle('ПСК "ИСКРА" Расчёт плана')
        self.setGeometry(100, 100, 760, 500)
        self.setAcceptDrops(True)

        # Список файлов и объединённые отфильтрованные данные.
        self.loaded_files: list[Path] = []
        self.file_data: dict[str, pd.DataFrame] = {}
        self.df_filtered: Optional[pd.DataFrame] = None
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        top_layout = QHBoxLayout()

        self.drop_zone = QLabel(
            "Перетащите сюда один или несколько Excel-файлов\n"
            "или нажмите «Добавить файлы…»"
        )
        self.drop_zone.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_zone.setWordWrap(True)
        self.drop_zone.setStyleSheet(
            "border: 2px dashed gray; padding: 15px; font-size: 13px; "
            "color: gray; background-color: #f9f9f9;"
        )
        top_layout.addWidget(self.drop_zone, stretch=2)

        inputs_group = QWidget()
        inputs_layout = QVBoxLayout(inputs_group)
        buttons_layout = QHBoxLayout()

        self.open_button = QPushButton("Добавить файлы…")
        self.open_button.clicked.connect(self.select_files_via_dialog)
        buttons_layout.addWidget(self.open_button)

        self.clear_button = QPushButton("Очистить список")
        self.clear_button.clicked.connect(self.clear_files)
        buttons_layout.addWidget(self.clear_button)

        self.about_button = QPushButton("О программе")
        self.about_button.clicked.connect(self.show_about_dialog)
        buttons_layout.addWidget(self.about_button)
        inputs_layout.addLayout(buttons_layout)

        only_digits = QIntValidator()
        vin_layout = QHBoxLayout()
        vin_layout.addWidget(QLabel("VIN (5 цифр):"))
        self.vin_input = QLineEdit()
        self.vin_input.setValidator(only_digits)
        self.vin_input.setMaxLength(5)
        self.vin_input.setPlaceholderText("Например: 12345")
        vin_layout.addWidget(self.vin_input)
        inputs_layout.addLayout(vin_layout)

        plan_layout = QHBoxLayout()
        plan_layout.addWidget(QLabel("План:"))
        self.plan_input = QLineEdit()
        self.plan_input.setValidator(only_digits)
        self.plan_input.setPlaceholderText("Количество")
        plan_layout.addWidget(self.plan_input)
        inputs_layout.addLayout(plan_layout)

        self.analyze_button = QPushButton("Анализ")
        self.analyze_button.setStyleSheet(
            "font-weight: bold; background-color: #e8f5e9; color: #2e7d32; "
            "border: 1px solid #a5d6a7; padding: 5px;"
        )
        self.analyze_button.clicked.connect(self.analyze_data)
        inputs_layout.addWidget(self.analyze_button)
        top_layout.addWidget(inputs_group, stretch=2)

        stats_group = QWidget()
        stats_group.setStyleSheet(
            "background-color: #f5f5f5; border-radius: 5px; border: 1px solid #ddd;"
        )
        stats_layout = QVBoxLayout(stats_group)
        results_title = QLabel("<b>План на смену:</b>")
        results_title.setStyleSheet("color: #000000; font-size: 13px;")
        stats_layout.addWidget(results_title)

        self.ljo_counter_label = QLabel("Количество LJO: 0")
        self.ljo_counter_label.setStyleSheet("font-size: 14px; color: #1565c0; font-weight: bold;")
        stats_layout.addWidget(self.ljo_counter_label)
        self.bjo_counter_label = QLabel("Количество BJO: 0")
        self.bjo_counter_label.setStyleSheet("font-size: 14px; color: #43a047; font-weight: bold;")
        stats_layout.addWidget(self.bjo_counter_label)
        self.bjoc_counter_label = QLabel("Количество BJOC: 0")
        self.bjoc_counter_label.setStyleSheet("font-size: 14px; color: #ef6c00; font-weight: bold;")
        stats_layout.addWidget(self.bjoc_counter_label)
        top_layout.addWidget(stats_group, stretch=1)
        main_layout.addLayout(top_layout)

        self.table_widget = QTableWidget()
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.table_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ожидание загрузки файлов…")
        self.setCentralWidget(central_widget)

    # --- Drag and drop: принимаем сразу несколько файлов ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
            if paths and all(path.suffix.lower() in EXCEL_EXTENSIONS for path in paths):
                event.acceptProposedAction()
                self.drop_zone.setStyleSheet(
                    "border: 2px dashed green; padding: 15px; color: green; background-color: #e8f5e9;"
                )

    def dragLeaveEvent(self, event):
        self.reset_drop_zone_style()

    def dropEvent(self, event: QDropEvent):
        self.reset_drop_zone_style()
        paths = [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile() and Path(url.toLocalFile()).suffix.lower() in EXCEL_EXTENSIONS
        ]
        if paths:
            self.add_excel_files(paths)
            event.acceptProposedAction()

    def reset_drop_zone_style(self):
        self.drop_zone.setStyleSheet(
            "border: 2px dashed gray; padding: 15px; color: gray; background-color: #f9f9f9;"
        )

    def select_files_via_dialog(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите один или несколько файлов Excel",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm)",
        )
        if file_paths:
            self.add_excel_files([Path(path) for path in file_paths])

    def clear_files(self):
        self.loaded_files = []
        self.file_data = {}
        self.df_filtered = None
        self.drop_zone.setText("Перетащите сюда один или несколько Excel-файлов\nили нажмите «Добавить файлы…»")
        self.reset_drop_zone_style()
        self.reset_results()
        self.status_bar.showMessage("Список файлов очищен.")

    def reset_results(self):
        self.ljo_counter_label.setText("Количество LJO: 0")
        self.bjo_counter_label.setText("Количество BJO: 0")
        self.bjoc_counter_label.setText("Количество BJOC: 0")
        self.table_widget.setRowCount(0)
        self.table_widget.setColumnCount(0)

    def _find_header_row(self, file_path: Path) -> int:
        df_preview = pd.read_excel(file_path, header=None, nrows=20)
        for idx, row in df_preview.iterrows():
            row_values = {str(value).strip().casefold() for value in row.dropna()}
            if REQUIRED_COLUMNS.issubset(row_values):
                return idx
        raise ValueError(
            "Не найдена строка заголовков с обязательными столбцами: "
            "Модель, VIN, ДОП ВЕРСИЯ (поиск в первых 20 строках)."
        )

    def _read_excel_file(self, file_path: Path) -> pd.DataFrame:
        detected_header = self._find_header_row(file_path)
        df = pd.read_excel(file_path, header=detected_header)
        df.columns = [str(col).strip().casefold() for col in df.columns]

        rename_dict = {
            "модель": "Модель",
            "vin": "VIN",
            "доп версия": "ДОП ВЕРСИЯ",
        }
        df = df.rename(columns=rename_dict)
        missing = {"Модель", "VIN", "ДОП ВЕРСИЯ"} - set(df.columns)
        if missing:
            raise ValueError("В файле отсутствуют обязательные столбцы: " + ", ".join(sorted(missing)))

        df = df[["Модель", "VIN", "ДОП ВЕРСИЯ"]].copy()
        model_series = df["Модель"].astype(str).str.strip().str.upper()
        return df.loc[model_series.isin(["LJO", "BJO"])].reset_index(drop=True)

    def _check_date_sequence(self, candidate_files: list[Path], new_files: set[str]) -> bool:
        dated = [(path, extract_date_from_filename(path)) for path in candidate_files]
        undated = [(path, file_date) for path, file_date in dated if file_date is None]
        dated = [(path, file_date) for path, file_date in dated if file_date is not None]
        dated.sort(key=lambda item: (item[1], item[0].name.casefold()))

        warnings = []
        for (previous_path, previous_date), (next_path, next_date) in zip(dated, dated[1:]):
            # Проверяем пары, затронутые добавлением новых файлов.
            if str(previous_path.resolve()) not in new_files and str(next_path.resolve()) not in new_files:
                continue
            difference = (next_date - previous_date).days
            if difference > 1:
                warnings.append(
                    f"• {previous_path.name} — {previous_date.strftime('%d.%m.%Y')}\n"
                    f"  {next_path.name} — {next_date.strftime('%d.%m.%Y')}\n"
                    f"  Разрыв: {difference} дн."
                )

        if undated:
            warnings.extend(
                f"• Не удалось определить дату в имени файла: {path.name}"
                for path, _ in undated
                if str(path.resolve()) in new_files
            )

        if warnings:
            message = (
                "Обнаружена непоследовательность дат или дата не распознана:\n\n"
                + "\n\n".join(warnings)
                + "\n\nПродолжить загрузку файлов?"
            )
            answer = QMessageBox.warning(
                self,
                "Проверка последовательности дат",
                message,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            return answer == QMessageBox.StandardButton.Yes
        return True

    def add_excel_files(self, file_paths: list[Path]):
        # Убираем дубликаты и файлы, которые уже загружены.
        new_files = []
        known_paths = {str(path.resolve()) for path in self.loaded_files}
        for path in file_paths:
            path = path.resolve()
            if path.suffix.lower() not in EXCEL_EXTENSIONS:
                continue
            if str(path) not in known_paths and path.is_file():
                new_files.append(path)
                known_paths.add(str(path))

        if not new_files:
            QMessageBox.information(self, "Файлы", "Новых Excel-файлов для добавления нет.")
            return

        candidate_files = sorted(self.loaded_files + new_files, key=file_sort_key)
        new_paths_resolved = {str(path.resolve()) for path in new_files}
        if not self._check_date_sequence(candidate_files, new_paths_resolved):
            self.status_bar.showMessage("Загрузка новых файлов отменена.")
            return

        # Читаем все новые файлы до изменения текущего состояния: при ошибке список останется прежним.
        try:
            newly_read = {str(path.resolve()): self._read_excel_file(path) for path in new_files}
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка загрузки", f"Не удалось обработать файл:\n{exc}")
            return

        self.file_data.update(newly_read)
        self.loaded_files = candidate_files
        frames = [self.file_data[str(path.resolve())] for path in self.loaded_files]
        self.df_filtered = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
            columns=["Модель", "VIN", "ДОП ВЕРСИЯ"]
        )

        file_lines = [
            f"{path.name} — {extract_date_from_filename(path).strftime('%d.%m.%Y')}"
            if extract_date_from_filename(path)
            else f"{path.name} — дата не распознана"
            for path in self.loaded_files
        ]
        self.drop_zone.setText("Загружено файлов: " + str(len(self.loaded_files)) + "\n\n" + "\n".join(file_lines))
        self.reset_results()
        self.status_bar.showMessage(
            f"Загружено файлов: {len(self.loaded_files)}. Строк с LJO/BJO: {len(self.df_filtered)}"
        )

    def show_about_dialog(self):
        about_text = (
            '<h3>ПСК "ИСКРА" Расчёт плана</h3>'
            "<p>Программа для автоматизации обработки планов производства.</p>"
            "<p><b>Разработчик:</b> Культяпов А.В. / 6152 / 021А</p>"
            "<p>Исходный код проекта:<br>"
            "<a href='https://github.com/camellan/plan'>https://github.com/camellan/plan</a></p>"
            "<p><b>Версия:</b> 1.1.0</p>"
            '<p>© ПСК "ИСКРА", 2026</p>'
        )
        QMessageBox.information(self, "О программе", about_text)

    def analyze_data(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Предупреждение", "Сначала загрузите корректный Excel-файл или файлы.")
            return

        vin_tail = self.vin_input.text().strip()
        plan_str = self.plan_input.text().strip()
        if not vin_tail or not plan_str:
            QMessageBox.warning(self, "Предупреждение", "Заполните оба поля: 'VIN' и 'План'.")
            return

        plan_count = int(plan_str)
        if plan_count <= 0:
            QMessageBox.warning(self, "Предупреждение", "План должен быть больше нуля.")
            return

        # Ищем в объединённом списке, упорядоченном по датам файлов.
        data = self.df_filtered.copy()
        data["VIN_str"] = data["VIN"].fillna("").astype(str).str.strip()
        found_index = None
        for idx, row in data.iterrows():
            if row["VIN_str"].endswith(vin_tail):
                found_index = idx
                break

        if found_index is None:
            QMessageBox.critical(
                self,
                "Ошибка",
                f"VIN с окончанием '{vin_tail}' не найден среди моделей LJO/BJO во всех загруженных файлах.",
            )
            return

        df_plan_slice = data.iloc[found_index:found_index + plan_count].copy().fillna("")
        models_series = df_plan_slice["Модель"].astype(str).str.strip().str.upper()
        extra_series = df_plan_slice["ДОП ВЕРСИЯ"].astype(str).str.strip().str.upper()
        ljo_count = (models_series == "LJO").sum()
        bjoc_count = ((models_series == "BJO") & (extra_series == "VSUVE")).sum()
        total_bjo = (models_series == "BJO").sum()
        bjo_count = total_bjo - bjoc_count

        self.ljo_counter_label.setText(f"Количество LJO: {ljo_count}")
        self.bjo_counter_label.setText(f"Количество BJO: {bjo_count}")
        self.bjoc_counter_label.setText(f"Количество BJOC: {bjoc_count}")
        display_df = df_plan_slice.drop(columns=["VIN_str"], errors="ignore")

        self.table_widget.setRowCount(display_df.shape[0])
        self.table_widget.setColumnCount(display_df.shape[1])
        self.table_widget.setHorizontalHeaderLabels(list(display_df.columns))
        for row_idx, row in enumerate(display_df.itertuples(index=False, name=None)):
            for col_idx, value in enumerate(row):
                self.table_widget.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        self.status_bar.showMessage(f"Анализ завершен. Найдено строк в плане: {len(display_df)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Устанавливаем русскую локаль
    locale = QLocale(QLocale.Language.Russian, QLocale.Country.Russia)
    QLocale.setDefault(locale)

    # Подключаем переводчик (если есть файлы переводов Qt)
    translator = QTranslator()
    if translator.load("qtbase_ru", ":/translations"):
        app.installTranslator(translator)
        
    window = VinAnalyzerApp()
    
    window.show()
    sys.exit(app.exec())
