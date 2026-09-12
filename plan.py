import sys
import os
import pandas as pd
from pathlib import Path
from typing import Optional

# Подключаем ctypes для принудительного включения иконки в Windows
if sys.platform == "win32":
    import ctypes
    myappid = 'psk.iskra.vinanalyzer.1.0'  # Уникальный ID приложения
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
    QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIntValidator, QIcon

def get_resource_path(relative_path: str) -> str:
    """Возвращает абсолютный путь к ресурсу, адаптированный для работы и в .py, и в .exe"""
    try:
        # PyInstaller создает временную папку _MEIPASS при работе .exe
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class VinAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Сначала настраиваем иконку, чтобы она подгрузилась до отрисовки рамки
        icon_path = get_resource_path("lada.ico")
        self.setWindowIcon(QIcon(icon_path))
        
        self.setWindowTitle('ПСК "ИСКРА" Расчёт плана')
        self.setGeometry(100, 100, 600, 400)
        
        # Разрешаем Drop файлов на окно
        self.setAcceptDrops(True)
        
        # Переменные состояния
        self.current_file_path: Optional[Path] = None
        self.df_filtered: Optional[pd.DataFrame] = None  # Только LJO и BJO

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)

        # --- ВЕРХНЯЯ ПАНЕЛЬ: Управление ---
        top_layout = QHBoxLayout()
        
        # Зона перетаскивания файла
        self.drop_zone = QLabel("Перетащите Excel файл сюда\nили нажмите кнопку\nвыбора файла")
        self.drop_zone.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_zone.setStyleSheet(
            "border: 2px dashed gray; padding: 15px; font-size: 13px; color: gray; background-color: #f9f9f9;"
        )
        top_layout.addWidget(self.drop_zone, stretch=2)
        
        # Блок фильтров и ввода параметров
        inputs_group = QWidget()
        inputs_layout = QVBoxLayout(inputs_group)
        
        # Горизонтальный слой для кнопок выбора файла и информации
        buttons_layout = QHBoxLayout()
        self.open_button = QPushButton("Выбрать файл...")
        self.open_button.clicked.connect(self.select_file_via_dialog)
        buttons_layout.addWidget(self.open_button)
        
        # Кнопка "О программе"
        self.about_button = QPushButton("О программе")
        self.about_button.clicked.connect(self.show_about_dialog)
        buttons_layout.addWidget(self.about_button)
        
        inputs_layout.addLayout(buttons_layout)
        
        # Валидатор для ввода только цифр
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
        
        # Кнопка "Анализировать" зеленого цвета
        self.analyze_button = QPushButton("Анализ")
        self.analyze_button.setStyleSheet(
            "font-weight: bold; background-color: #e8f5e9; color: #2e7d32; border: 1px solid #a5d6a7; padding: 5px;"
        )
        self.analyze_button.clicked.connect(self.analyze_data)
        inputs_layout.addWidget(self.analyze_button)
        
        top_layout.addWidget(inputs_group, stretch=1)
        
        # Блок Счётчиков (LJO / BJO / BJOC)
        stats_group = QWidget()
        stats_group.setStyleSheet("background-color: #f5f5f5; border-radius: 5px; border: 1px solid #ddd;")
        stats_layout = QVBoxLayout(stats_group)
        
        # Заголовок результатов черного цвета
        results_title = QLabel("<b>План на смену:</b>")
        results_title.setStyleSheet("color: #000000; font-size: 13px;")
        stats_layout.addWidget(results_title)
        
        self.ljo_counter_label = QLabel("Количество LJO: 0")
        self.ljo_counter_label.setStyleSheet("font-size: 14px; color: #1565c0; font-weight: bold;")
        stats_layout.addWidget(self.ljo_counter_label)
        
        self.bjo_counter_label = QLabel("Количество BJO: 0")
        self.bjo_counter_label.setStyleSheet("font-size: 14px; color: #43a047; font-weight: bold;")
        stats_layout.addWidget(self.bjo_counter_label)

        # Метка для BJOC
        self.bjoc_counter_label = QLabel("Количество BJOC: 0")
        self.bjoc_counter_label.setStyleSheet("font-size: 14px; color: #ef6c00; font-weight: bold;")
        stats_layout.addWidget(self.bjoc_counter_label)
        
        top_layout.addWidget(stats_group, stretch=1)
        main_layout.addLayout(top_layout)
        
        # --- ТАБЛИЦА (Показывает срез по плану) ---
        self.table_widget = QTableWidget()
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.table_widget)
        
        # --- СТАТУС БАР ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ожидание загрузки файла...")
        
        self.setCentralWidget(central_widget)

    # --- Логика Drag and Drop ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                url = urls[0].toLocalFile()
                if url.endswith(('.xlsx', '.xls', '.xlsm')):
                    event.acceptProposedAction()
                    self.drop_zone.setStyleSheet(
                        "border: 2px dashed green; padding: 15px; color: green; background-color: #e8f5e9;"
                    )

    def dragLeaveEvent(self, event):
        self.reset_drop_zone_style()

    def dropEvent(self, event: QDropEvent):
        self.reset_drop_zone_style()
        urls = event.mimeData().urls()
        if urls:
            file_path = Path(urls[0].toLocalFile())
            self.load_excel_file(file_path)

    def reset_drop_zone_style(self):
        self.drop_zone.setStyleSheet(
            "border: 2px dashed gray; padding: 15px; color: gray; background-color: #f9f9f9;"
        )

    def select_file_via_dialog(self):
        file_filter = "Excel Files (*.xlsx *.xls *.xlsm)"
        file_path_str, _ = QFileDialog.getOpenFileName(self, "Выберите файл Excel", "", file_filter)
        if file_path_str:
            self.load_excel_file(Path(file_path_str))
    # Метод вывода окна информации о разработчике
    def show_about_dialog(self):
        about_text = (
            "<h3>ПСК \"ИСКРА\" Расчёт плана</h3>\n"
            "<p>Программа для автоматизации обработки планов производства.</p>\n"
            "<p><b>Разработчик:</b> Культяпов А.В. / 6152 / 021А</p>\n"
            "<p>Исходный код проекта:<br>"
            "<a href='https://github.com/camellan/plan'>https://github.com/camellan/plan</a></p>"
            "<p><b>Версия:</b> 1.0.0</p>\n"
            "<p>© ПСК \"ИСКРА\", 2026</p>"
        )
        QMessageBox.information(self, "О программе", about_text)            

    # --- Автоматический поиск строки заголовков ---
    def _find_header_row(self, file_path: Path) -> int:
        df_preview = pd.read_excel(file_path, header=None, nrows=20)
        required_headers = {"модель", "vin", "доп версия"}
        
        for idx, row in df_preview.iterrows():
            row_values = [str(val).strip().lower() for val in row.dropna()]
            if all(req in row_values for req in required_headers):
                return idx
        return 0

    # --- Загрузка и фильтрация LJO / BJO ---
    def load_excel_file(self, file_path: Path):
        try:
            self.current_file_path = file_path
            self.drop_zone.setText(f"Файл выбран:\n{file_path.name}")
            
            detected_header = self._find_header_row(file_path)
            df = pd.read_excel(file_path, header=detected_header)
            
            df.columns = [str(col).strip().lower() for col in df.columns]
            
            rename_dict = {}
            for col in df.columns:
                if col == "модель": rename_dict[col] = "Модель"
                elif col == "vin": rename_dict[col] = "VIN"
                elif col == "доп версия": rename_dict[col] = "ДОП ВЕРСИЯ"
            df = df.rename(columns=rename_dict)
            
            available_cols = [c for c in ["Модель", "VIN", "ДОП ВЕРСИЯ"] if c in df.columns]
            df = df[available_cols]
            
            if "Модель" in df.columns:
                df["Модель_str"] = df["Модель"].astype(str).str.strip().str.upper()
                self.df_filtered = df[df["Модель_str"].isin(["LJO", "BJO"])].copy()
                self.df_filtered = self.df_filtered.drop(columns=["Модель_str"], errors="ignore")
                self.df_filtered = self.df_filtered.reset_index(drop=True)
            else:
                self.df_filtered = df
                
            self.status_bar.showMessage(f"Файл загружен. Строк с LJO/BJO: {len(self.df_filtered)}")
            
            self.ljo_counter_label.setText("Количество LJO: 0")
            self.bjo_counter_label.setText("Количество BJO: 0")
            self.bjoc_counter_label.setText("Количество BJOC: 0")
            self.table_widget.setRowCount(0)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка загрузки", f"Не удалось обработать файл:\n{str(e)}")

    # --- Анализ по VIN и Плану ---
    def analyze_data(self):
        if self.df_filtered is None or self.df_filtered.empty:
            QMessageBox.warning(self, "Предупреждение", "Сначала загрузите корректный Excel файл.")
            return
            
        vin_tail = self.vin_input.text().strip()
        plan_str = self.plan_input.text().strip()
        
        if not vin_tail or not plan_str:
            QMessageBox.warning(self, "Предупреждение", "Заполните оба поля: 'VIN' и 'План'.")
            return
            
        plan_count = int(plan_str)
        
        found_index = None
        self.df_filtered["VIN_str"] = self.df_filtered["VIN"].astype(str).str.strip()
        
        for idx, row in self.df_filtered.iterrows():
            if row["VIN_str"].endswith(vin_tail):
                found_index = idx
                break
                
        if found_index is None:
            QMessageBox.critical(self, "Ошибка", f"VIN с окончанием '{vin_tail}' не найден среди моделей LJO/BJO.")
            return

        df_plan_slice = self.df_filtered.iloc[found_index : found_index + plan_count].copy()
        df_plan_slice = df_plan_slice.fillna("")

        # Вспомогательные серии для точного подсчета
        models_series = df_plan_slice["Модель"].astype(str).str.strip().str.upper()
        extra_series = df_plan_slice["ДОП ВЕРСИЯ"].astype(str).str.strip().str.upper()

        # Расчет количества моделей
        ljo_count = (models_series == "LJO").sum()
        
        # BJOC — это BJO со значением VSUVE
        bjoc_count = ((models_series == "BJO") & (extra_series == "VSUVE")).sum()
        
        # BJO теперь исключает из себя BJOC
        total_bjo = (models_series == "BJO").sum()
        bjo_count = total_bjo - bjoc_count

        # Обновляем все текстовые метки на экране
        self.ljo_counter_label.setText(f"Количество LJO: {ljo_count}")
        self.bjo_counter_label.setText(f"Количество BJO: {bjo_count}")
        self.bjoc_counter_label.setText(f"Количество BJOC: {bjoc_count}")

        display_df = df_plan_slice.drop(columns=["VIN_str"], errors="ignore")
        
        self.table_widget.setRowCount(display_df.shape[0])
        self.table_widget.setColumnCount(display_df.shape[1])
        
        self.table_widget.setHorizontalHeaderLabels(list(display_df.columns))
        
        for row_idx, row in display_df.reset_index(drop=True).iterrows():
            for col_idx, value in enumerate(row):
                self.table_widget.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))

        self.status_bar.showMessage(f"Анализ завершен. Найдено строк в плане: {len(display_df)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VinAnalyzerApp()
    window.show()
    sys.exit(app.exec())
