# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Культяпов Андрей (Camellan)

import sys
import os
import re
import pickle
from datetime import date
from pathlib import Path
from typing import Optional
from collections import Counter
from PyQt6.QtNetwork import QLocalServer, QLocalSocket

SINGLE_INSTANCE_KEY = "PSK_ISKRA_PlanAnalyzer_SingleInstance"


import pandas as pd

if sys.platform == "win32":
    import ctypes
    import winreg
    myappid = 'psk.iskra.vinanalyzer.1.3'
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
    QSystemTrayIcon,
    QMenu,
    QCheckBox,
)
from PyQt6.QtCore import Qt, QLocale, QTranslator, QSettings
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIntValidator, QIcon, QAction, QCloseEvent

EXCEL_EXTENSIONS = (".xlsx", ".xls", ".xlsm")
REQUIRED_COLUMNS = {"модель", "vin", "доп версия"}
RUN_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_REG_NAME = "PSK_ISKRA_PlanAnalyzer"


def get_resource_path(relative_path: str) -> str:
    """Возвращает путь к ресурсу для запуска из .py и собранного .exe."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def get_app_data_dir() -> Path:
    """Возвращает директорию для хранения кэша сессии."""
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path.home() / ".local" / "share"
    data_dir = base / "PSK_ISKRA" / "PlanAnalyzer"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir


def get_autostart_command() -> str:
    """Формирует команду запуска с флагом --autostart для реестра Windows."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --autostart'
    script_path = os.path.abspath(__file__)
    pythonw = sys.executable.replace("python.exe", "pythonw.exe")
    exe_to_use = pythonw if os.path.exists(pythonw) else sys.executable
    return f'"{exe_to_use}" "{script_path}" --autostart'


def is_autostart_enabled() -> bool:
    """Проверяет наличие записи автозапуска в реестре текущего пользователя."""
    if sys.platform != "win32":
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, APP_REG_NAME)
            return bool(value)
    except OSError:
        return False


def set_autostart(enabled: bool) -> bool:
    """Включает или отключает автозапуск программы вместе с Windows."""
    if sys.platform != "win32":
        return False
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, get_autostart_command())
            else:
                try:
                    winreg.DeleteValue(key, APP_REG_NAME)
                except FileNotFoundError:
                    pass
        return True
    except OSError:
        return False
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
        self.app_icon = QIcon(get_resource_path("lada.ico"))
        self.setWindowIcon(self.app_icon)
        self.setWindowTitle('ПСК "ИСКРА" Расчёт плана')
        self.setGeometry(100, 100, 760, 500)
        self.setAcceptDrops(True)

        self.settings = QSettings("PSK_ISKRA", "VinAnalyzer")
        self.cache_file = get_app_data_dir() / "session_cache.pkl"
        self._force_quit = False

        self.loaded_files: list[Path] = list()
        self.file_data: dict[str, pd.DataFrame] = dict()
        self.df_filtered: Optional[pd.DataFrame] = None
        self.sequence_ratio_label = None

        self.init_ui()
        self.init_tray()
        self.restore_state_and_session()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        top_layout = QHBoxLayout()

        self.drop_zone = QLabel(
            "Перетащите сюда один или несколько Excel-файлов\n"
            "или нажмите «Добавить файлы…»"
        )
        self.drop_zone.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.drop_zone.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.drop_zone.setWordWrap(True)
        self.reset_drop_zone_style()

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
        self.vin_input.textChanged.connect(self.save_settings)
        vin_layout.addWidget(self.vin_input)
        inputs_layout.addLayout(vin_layout)

        plan_layout = QHBoxLayout()
        plan_layout.addWidget(QLabel("План:"))
        self.plan_input = QLineEdit()
        self.plan_input.setValidator(only_digits)
        self.plan_input.setPlaceholderText("Количество")
        self.plan_input.textChanged.connect(self.save_settings)
        plan_layout.addWidget(self.plan_input)
        inputs_layout.addLayout(plan_layout)

        self.analyze_button = QPushButton("Анализ")
        self.analyze_button.setFixedWidth(120)
        self.analyze_button.setStyleSheet(
            "QPushButton {"
            "  font-weight: bold;"
            "  background-color: #e8f5e9;"
            "  color: #2e7d32;"
            "  border: 1px solid #a5d6a7;"
            "  border-radius: 4px;"
            "  padding: 5px;"
            "}"
            "QPushButton:hover {"
            "  background-color: #dcedc8;"
            "}"
            "QPushButton:pressed {"
            "  background-color: #c5e1a5;"
            "}"
        )
        self.analyze_button.clicked.connect(self.analyze_data)

        analyze_layout = QHBoxLayout()
        analyze_layout.addStretch()
        analyze_layout.addWidget(self.analyze_button)
        inputs_layout.addLayout(analyze_layout)



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

        self.sequence_ratio_label = QLabel("Очерёдность LJO/BJO(C): —")
        self.sequence_ratio_label.setStyleSheet("font-size: 14px; color: #424242; font-weight: bold;")
        stats_layout.addWidget(self.sequence_ratio_label)

        top_layout.addWidget(stats_group, stretch=1)
        main_layout.addLayout(top_layout)

        self.table_widget = QTableWidget()
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.table_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ожидание загрузки файлов…")
        self.setCentralWidget(central_widget)
    # --- Системный трей и автозапуск ---
    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(self.app_icon, self)
        self.tray_icon.setToolTip('ПСК "ИСКРА" Расчёт плана')

        tray_menu = QMenu(self)

        restore_action = QAction("Развернуть", self)
        restore_action.triggered.connect(self.restore_from_tray)
        tray_menu.addAction(restore_action)

        self.autostart_action = QAction("Запускать вместе с Windows (в трей)", self)
        self.autostart_action.setCheckable(True)
        self.autostart_action.setChecked(is_autostart_enabled())
        self.autostart_action.triggered.connect(self.toggle_autostart)
        tray_menu.addAction(self.autostart_action)

        reset_close_pref_action = QAction("Спрашивать при закрытии", self)
        reset_close_pref_action.triggered.connect(self.reset_close_behavior)
        tray_menu.addAction(reset_close_pref_action)

        tray_menu.addSeparator()

        quit_action = QAction("Выход", self)
        quit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def toggle_autostart(self, checked: bool):
        if set_autostart(checked):
            state_text = "включён" if checked else "отключён"
            self.status_bar.showMessage(f"Автозапуск с Windows {state_text}.")
        else:
            self.autostart_action.setChecked(not checked)
            QMessageBox.warning(
                self,
                "Автозапуск",
                "Не удалось изменить параметр автозапуска в реестре Windows.",
            )

    def on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason):
        if reason in (
            QSystemTrayIcon.ActivationReason.Trigger,
            QSystemTrayIcon.ActivationReason.DoubleClick,
        ):
            self.restore_from_tray()

    def restore_from_tray(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def attach_single_instance_server(self, server: QLocalServer):
        """Принимает сигнал от повторно запущенной копии и разворачивает окно."""
        self.instance_server = server
        self.instance_server.newConnection.connect(self._on_new_instance_connection)

    def _on_new_instance_connection(self):
        socket = self.instance_server.nextPendingConnection()
        if socket:
            socket.waitForReadyRead(500)
            command = bytes(socket.readAll()).decode("utf-8", errors="ignore").strip()
            socket.disconnectFromServer()
            if command == "SHOW":
                self.restore_from_tray()

    def reset_close_behavior(self):
        self.settings.remove("close_action")
        QMessageBox.information(
            self,
            "Настройки закрытия",
            "Настройка сброшена. При следующем закрытии программа снова спросит, свернуть её в трей или выйти.",
        )

    def quit_application(self):
        self._force_quit = True
        self.save_settings()
        self.save_session_cache()

        # 1. Закрываем локальный сервер одиночного запуска
        if hasattr(self, "instance_server") and self.instance_server is not None:
            self.instance_server.close()
            QLocalServer.removeServer(SINGLE_INSTANCE_KEY)

        # 2. Скрываем и освобождаем иконку трея
        if hasattr(self, "tray_icon") and self.tray_icon is not None:
            self.tray_icon.hide()
            self.tray_icon.deleteLater()

        # 3. Возвращаем стандартное поведение выхода и закрываем окно
        QApplication.setQuitOnLastWindowClosed(True)
        self.close()
        QApplication.quit()


    def closeEvent(self, event: QCloseEvent):
        self.save_settings()
        self.save_session_cache()

        if self._force_quit or not QSystemTrayIcon.isSystemTrayAvailable():
            event.accept()
            return

        saved_action = self.settings.value("close_action", "")
        if saved_action == "tray":
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                'ПСК "ИСКРА" Расчёт плана',
                "Программа свёрнута в системный трей.",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
            return
        elif saved_action == "exit":
            event.accept()
            self.quit_application()
            return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Закрытие программы")
        msg_box.setText("Свернуть программу в системный трей\nили завершить работу?")
        msg_box.setIcon(QMessageBox.Icon.Question)

        tray_btn = msg_box.addButton("В трей", QMessageBox.ButtonRole.AcceptRole)
        exit_btn = msg_box.addButton("Выход", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = msg_box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)

        remember_cb = QCheckBox("Запомнить мой выбор")
        msg_box.setCheckBox(remember_cb)

        msg_box.exec()
        clicked = msg_box.clickedButton()

        if clicked == cancel_btn:
            event.ignore()
        elif clicked == tray_btn:
            if remember_cb.isChecked():
                self.settings.setValue("close_action", "tray")
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                'ПСК "ИСКРА" Расчёт плана',
                "Программа свёрнута в системный трей.",
                QSystemTrayIcon.MessageIcon.Information,
                2000,
            )
        elif clicked == exit_btn:
            if remember_cb.isChecked():
                self.settings.setValue("close_action", "exit")
            event.accept()
            self.quit_application()

    # --- Сохранение и восстановление настроек и сессии ---
    def save_settings(self):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("vin_tail", self.vin_input.text().strip())
        self.settings.setValue("plan_count", self.plan_input.text().strip())

    def save_session_cache(self):
        """Сохраняет список путей и уже распарсенные DataFrame на диск."""
        try:
            payload = {
                "loaded_files": [str(p) for p in self.loaded_files],
                "file_data": self.file_data,
            }
            with open(self.cache_file, "wb") as f:
                pickle.dump(payload, f)
        except Exception:
            pass

    def restore_state_and_session(self):
        """Восстанавливает геометрию окна, поля ввода и кэшированные таблицы."""
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

        saved_vin = self.settings.value("vin_tail", "")
        saved_plan = self.settings.value("plan_count", "")

        self.vin_input.blockSignals(True)
        self.plan_input.blockSignals(True)
        self.vin_input.setText(str(saved_vin))
        self.plan_input.setText(str(saved_plan))
        self.vin_input.blockSignals(False)
        self.plan_input.blockSignals(False)

        if not self.cache_file.exists():
            return

        try:
            with open(self.cache_file, "rb") as f:
                payload = pickle.load(f)

            raw_paths = payload.get("loaded_files", list())
            cached_data = payload.get("file_data", dict())

            if not raw_paths or not cached_data:
                return

            self.loaded_files = [Path(p) for p in raw_paths]
            self.file_data = cached_data
            self._rebuild_combined_dataframe(restored_from_cache=True)

            # Если ранее были введены VIN и План — автоматически восстанавливаем таблицу расчёта
            if saved_vin and saved_plan and self.df_filtered is not None and not self.df_filtered.empty:
                self.analyze_data(silent=True)

        except Exception:
            self.loaded_files = list()
            self.file_data = dict()
            self.df_filtered = None

    def _rebuild_combined_dataframe(self, restored_from_cache: bool = False):
        frames = [self.file_data[str(path)] for path in self.loaded_files if str(path) in self.file_data]
        self.df_filtered = (
            pd.concat(frames, ignore_index=True)
            if frames
            else pd.DataFrame(columns=["Модель", "VIN", "ДОП ВЕРСИЯ"])
        )

        file_lines = list()
        missing_count = 0
        for path in self.loaded_files:
            file_date = extract_date_from_filename(path)
            date_str = file_date.strftime("%d.%m.%Y") if file_date else "дата не распознана"
            exists_mark = "" if path.exists() else " [из кэша, исходный файл удалён]"
            if not path.exists():
                missing_count += 1
            file_lines.append(f"{path.name} — {date_str}{exists_mark}")

        self.drop_zone.setText(
            f"Загружено файлов: {len(self.loaded_files)}\n\n" + "\n".join(file_lines)
        )
        self.reset_drop_zone_style()

        status_msg = f"Загружено файлов: {len(self.loaded_files)}. Строк с LJO/BJO: {len(self.df_filtered)}"
        if restored_from_cache:
            status_msg = f"Сессия восстановлена из кэша. {status_msg}"
        if missing_count:
            status_msg += f" (исходных файлов недоступно: {missing_count})"
        self.status_bar.showMessage(status_msg)

    # --- Drag and drop: принимаем сразу несколько файлов ---
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
            if paths and all(path.suffix.lower() in EXCEL_EXTENSIONS for path in paths):
                event.acceptProposedAction()
                self.drop_zone.setStyleSheet(
                    "border: 2px dashed #2e7d32; border-radius: 6px; padding: 15px; "
                    "font-size: 13px; color: #1b5e20; background-color: #c8e6c9;"
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
        if self.loaded_files:
            # Стиль, когда файлы уже загружены
            self.drop_zone.setStyleSheet(
                "border: 1px solid #90caf9; border-radius: 6px; padding: 15px; "
                "font-size: 13px; color: #0d47a1; background-color: #e3f2fd;"
            )
        else:
            # Стиль пустого поля (ожидание загрузки)
            self.drop_zone.setStyleSheet(
                "border: 2px dashed gray; border-radius: 6px; padding: 15px; "
                "font-size: 13px; color: gray; background-color: #f9f9f9;"
            )


    def select_files_via_dialog(self):
        last_dir = str(self.settings.value("last_open_dir", ""))
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите один или несколько файлов Excel",
            last_dir,
            "Excel Files (*.xlsx *.xls *.xlsm)",
        )
        if file_paths:
            self.settings.setValue("last_open_dir", str(Path(file_paths[0]).parent))
            self.add_excel_files([Path(path) for path in file_paths])

    def clear_files(self):
        self.loaded_files = list()
        self.file_data = dict()
        self.df_filtered = None
        if self.cache_file.exists():
            try:
                self.cache_file.unlink()
            except Exception:
                pass
        self.drop_zone.setText("Перетащите сюда один или несколько Excel-файлов\nили нажмите «Добавить файлы…»")
        self.reset_drop_zone_style()
        self.reset_results()
        self.status_bar.showMessage("Список файлов и кэш сессии очищены.")

    def reset_results(self):
        self.ljo_counter_label.setText("Количество LJO: 0")
        self.bjo_counter_label.setText("Количество BJO: 0")
        self.bjoc_counter_label.setText("Количество BJOC: 0")
        self.sequence_ratio_label.setText("Очерёдность LJO/BJO(C): —")
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

        warnings = list()
        for (previous_path, previous_date), (next_path, next_date) in zip(dated, dated[1:]):
            if str(previous_path) not in new_files and str(next_path) not in new_files:
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
                if str(path) in new_files
            )

        if warnings:
            message = (
                "Обнаружена непоследовательность дат или дата не распознана:\n\n"
                + "\n\n".join(warnings)
                + "\n\nПродолжить загрузку файлов?"
            )
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Проверка последовательности дат")
            msg_box.setText(message)
            msg_box.setIcon(QMessageBox.Icon.Warning)

            yes_btn = msg_box.addButton("Да", QMessageBox.ButtonRole.YesRole)
            no_btn = msg_box.addButton("Нет", QMessageBox.ButtonRole.NoRole)
            msg_box.setDefaultButton(no_btn)

            msg_box.exec()
            return msg_box.clickedButton() == yes_btn
        return True


    def add_excel_files(self, file_paths: list[Path]):
        new_files = list()
        known_paths = {str(path) for path in self.loaded_files}
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
        new_paths_resolved = {str(path) for path in new_files}
        if not self._check_date_sequence(candidate_files, new_paths_resolved):
            self.status_bar.showMessage("Загрузка новых файлов отменена.")
            return

        try:
            newly_read = {str(path): self._read_excel_file(path) for path in new_files}
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка загрузки", f"Не удалось обработать файл:\n{exc}")
            return

        self.file_data.update(newly_read)
        self.loaded_files = candidate_files
        self.reset_results()
        self._rebuild_combined_dataframe(restored_from_cache=False)
        self.save_session_cache()
    def show_about_dialog(self):
        about_text = (
            '<h3>ПСК "ИСКРА" Расчёт плана</h3>'
            "<p>Программа для автоматизации обработки планов производства.</p>"
            "<p><b>Разработчик:</b> Культяпов А.В. / 6152 / 021А</p>"
            "<p>Исходный код проекта:<br>"
            "<a href='https://github.com/camellan/plan'>https://github.com/camellan/plan</a></p>"
            "<p><b>Версия:</b> 1.3.0</p>"
            '<p>© ПСК "ИСКРА", 2026</p>'
        )
        QMessageBox.information(self, "О программе", about_text)

    def analyze_data(self, silent: bool = False):
        if self.df_filtered is None or self.df_filtered.empty:
            if not silent:
                QMessageBox.warning(self, "Предупреждение", "Сначала загрузите корректный Excel-файл или файлы.")
            return

        vin_tail = self.vin_input.text().strip()
        plan_str = self.plan_input.text().strip()
        if not vin_tail or not plan_str:
            if not silent:
                QMessageBox.warning(self, "Предупреждение", "Заполните оба поля: 'VIN' и 'План'.")
            return

        plan_count = int(plan_str)
        if plan_count <= 0:
            if not silent:
                QMessageBox.warning(self, "Предупреждение", "План должен быть больше нуля.")
            return

        data = self.df_filtered.copy()
        data["VIN_str"] = data["VIN"].fillna("").astype(str).str.strip()
        found_index = None
        for idx, row in data.iterrows():
            if row["VIN_str"].endswith(vin_tail):
                found_index = idx
                break

        if found_index is None:
            if not silent:
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

        mode_ljo, mode_bjo_c = self._get_most_common_run_sizes(models_series)
        self.sequence_ratio_label.setText(f"Очерёдность LJO/BJO(C): {mode_ljo}/{mode_bjo_c}")

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

    @staticmethod
    def _most_common_size(sizes: list[int]) -> int:
        """Самый частый размер пачки. При равенстве частот — минимальный."""
        if not sizes:
            return 0
        counter = Counter(sizes)
        max_count = max(counter.values())
        candidates = [size for size, count in counter.items() if count == max_count]
        return min(candidates)

    def _get_most_common_run_sizes(self, models_series: pd.Series):
        """Находит самый частый размер непрерывной пачки для LJO и BJO(C)."""
        if models_series.empty:
            return 0, 0

        simplified = models_series.apply(lambda x: "BJO(C)" if x == "BJO" else x)

        runs = list()
        current_model = simplified.iloc[0]
        run_len = 1

        for val in simplified.iloc[1:]:
            if val == current_model:
                run_len += 1
            else:
                runs.append((current_model, run_len))
                current_model = val
                run_len = 1
        runs.append((current_model, run_len))

        ljo_lengths = [length for model, length in runs if model == "LJO"]
        bjo_lengths = [length for model, length in runs if model == "BJO(C)"]

        mode_ljo = self._most_common_size(ljo_lengths)
        mode_bjo = self._most_common_size(bjo_lengths)

        return mode_ljo, mode_bjo


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Проверяем, не запущена ли уже первая копия программы
    probe_socket = QLocalSocket()
    probe_socket.connectToServer(SINGLE_INSTANCE_KEY)
    if probe_socket.waitForConnected(300):
        if "--autostart" not in sys.argv:
            probe_socket.write(b"SHOW")
            probe_socket.waitForBytesWritten(300)
        probe_socket.disconnectFromServer()
        probe_socket.close()
        sys.exit(0)

    QLocalServer.removeServer(SINGLE_INSTANCE_KEY)
    instance_server = QLocalServer()
    instance_server.listen(SINGLE_INSTANCE_KEY)

    app.setQuitOnLastWindowClosed(False)

    locale = QLocale(QLocale.Language.Russian, QLocale.Country.Russia)
    QLocale.setDefault(locale)

    translator = QTranslator()
    if translator.load(get_resource_path("qtbase_ru"), get_resource_path("translations")):
        app.installTranslator(translator)

    window = VinAnalyzerApp()
    window.attach_single_instance_server(instance_server)

    if "--autostart" in sys.argv:
        window.hide()
    else:
        window.show()

    exit_code = app.exec()

    # Гарантированно освобождаем ресурсы до завершения python.exe
    instance_server.close()
    QLocalServer.removeServer(SINGLE_INSTANCE_KEY)
    del window
    del instance_server

    sys.exit(exit_code)
