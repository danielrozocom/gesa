import json
import os
import sys
import threading
import re
import ctypes
import copy
import subprocess
import tempfile

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller _MEIPASS."""
    if hasattr(sys, '_MEIPASS'):
        p = os.path.join(sys._MEIPASS, relative_path)
        if os.path.exists(p):
            return p
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, relative_path)


_instance_file_lock = None

def is_another_instance_running():
    global _instance_file_lock
    lock_file = os.path.join(tempfile.gettempdir(), "gesa_app_v520d.lock")
    try:
        _instance_file_lock = open(lock_file, "a+")
        if sys.platform == "win32":
            import msvcrt
            _instance_file_lock.seek(0)
            msvcrt.locking(_instance_file_lock.fileno(), msvcrt.LK_NBLCK, 1)
        return False
    except (IOError, OSError):
        return True
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QListWidget,
    QListWidgetItem, QTextEdit, QProgressBar, QFrame, QSplitter,
    QFileDialog, QMessageBox, QAbstractItemView,
    QListView, QDialog, QDialogButtonBox, QCalendarWidget, QSizePolicy, QToolButton,
    QSplashScreen,
)
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, QDate, QLocale, QSettings, QRectF
from PyQt6.QtGui import QTextCharFormat, QColor, QPalette, QShortcut, QKeySequence, QIcon, QPainter, QPixmap
from PyQt6.QtSvg import QSvgRenderer


def _qta():
    import qtawesome as _qta_mod
    return _qta_mod


try:
    myappid = "danielrozocom.gesa.academic.v1"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

MONTH_MAP = {"ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4, "MAY": 5, "JUN": 6,
             "JUL": 7, "AGO": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DIC": 12}
MONTH_REV = {v: k for k, v in MONTH_MAP.items()}


def set_window_dark_mode(hwnd, dark: bool):
    try:
        value = ctypes.c_int(1 if dark else 0)
        res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            int(hwnd), 20, ctypes.byref(value), ctypes.sizeof(value)
        )
        if res != 0:
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                int(hwnd), 19, ctypes.byref(value), ctypes.sizeof(value)
            )
    except Exception:
        pass


class DateSelector(QWidget):
    dateChanged = pyqtSignal(QDate)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self._date = QDate.currentDate()

        self.display = QLineEdit()
        self.display.setReadOnly(True)
        self.display.setObjectName("date-display")
        layout.addWidget(self.display)

        self.btn = QPushButton()
        self.btn.setObjectName("btn-icon")
        self.btn.clicked.connect(self._open_calendar)
        layout.addWidget(self.btn)

        self._update_display()

    def _update_display(self):
        d = self._date
        locale = QLocale(QLocale.Language.Spanish)
        day_full = locale.dayName(d.dayOfWeek(), QLocale.FormatType.LongFormat).capitalize()
        self.display.setText(f"{day_full} {str(d.day()).zfill(2)}/{MONTH_REV[d.month()]}/{d.year()}")

    def _open_calendar(self):
        main_win = self.window()
        dlg = QDialog(main_win or self)
        dlg.setWindowTitle("Seleccionar fecha")
        dlg.setMinimumWidth(360)

        is_dark = True
        if hasattr(main_win, "_effective_theme"):
            is_dark = (main_win._effective_theme() == "dark")
        set_window_dark_mode(dlg.winId(), is_dark)

        bg_color = "#18181b" if is_dark else "#ffffff"
        text_color = "#fafafa" if is_dark else "#0f172a"
        border_color = "#27272a" if is_dark else "#e2e8f0"
        hover_bg = "#27272a" if is_dark else "#f1f5f9"
        sel_bg = "#3b82f6" if is_dark else "#2563eb"
        weekend_color = "#f87171" if is_dark else "#e11d48"

        dlg.setStyleSheet(f"""
            QDialog {{
                background-color: {bg_color};
                color: {text_color};
                border-radius: 12px;
            }}
            QCalendarWidget {{
                background-color: {bg_color};
                color: {text_color};
            }}
            QCalendarWidget QWidget#qt_calendar_navigationbar {{
                background-color: {bg_color};
                border-bottom: 1px solid {border_color};
                padding: 4px;
            }}
            QCalendarWidget QToolButton {{
                color: {text_color};
                background-color: transparent;
                border-radius: 6px;
                font-family: "Segoe UI";
                font-size: 13px;
                font-weight: 600;
                padding: 4px 8px;
            }}
            QCalendarWidget QToolButton:hover {{
                background-color: {hover_bg};
            }}
            QCalendarWidget QMenu {{
                background-color: {bg_color};
                color: {text_color};
                border: 1px solid {border_color};
            }}
            QCalendarWidget QSpinBox {{
                color: {text_color};
                background-color: {hover_bg};
                border-radius: 4px;
                font-size: 13px;
            }}
            QCalendarWidget QTableView {{
                background-color: {bg_color};
                color: {text_color};
                selection-background-color: {sel_bg};
                selection-color: #ffffff;
                font-family: "Segoe UI";
                font-size: 13px;
                outline: 0;
                gridline-color: transparent;
            }}
        """)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        cal = QCalendarWidget()
        cal.setSelectedDate(self._date)
        cal.setGridVisible(False)
        cal.setLocale(QLocale(QLocale.Language.Spanish))
        cal.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        cal.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)

        prev_btn = cal.findChild(QToolButton, "qt_calendar_prevmonth")
        next_btn = cal.findChild(QToolButton, "qt_calendar_nextmonth")
        if prev_btn:
            prev_btn.setIcon(_qta().icon("fa5s.chevron-left", color=text_color))
        if next_btn:
            next_btn.setIcon(_qta().icon("fa5s.chevron-right", color=text_color))

        fmt_weekend = QTextCharFormat()
        fmt_weekend.setForeground(QColor(weekend_color))
        cal.setWeekdayTextFormat(Qt.DayOfWeek.Saturday, fmt_weekend)
        cal.setWeekdayTextFormat(Qt.DayOfWeek.Sunday, fmt_weekend)

        layout.addWidget(cal)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setObjectName("btn-outline")
        cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        cancel_btn.clicked.connect(dlg.reject)

        ok_btn = QPushButton("Aceptar")
        ok_btn.setObjectName("btn-primary")
        ok_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        ok_btn.clicked.connect(dlg.accept)

        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(ok_btn)

        layout.addLayout(btn_layout)

        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._date = cal.selectedDate()
            self._update_display()
            self.dateChanged.emit(self._date)

    def setDate(self, date):
        self._date = QDate(date)
        self._update_display()

    def date(self):
        return QDate(self._date)


SHADCN_DARK = {
    "bg": "#0a0e17", "card": "#131927", "border": "#1e293b",
    "text": "#f8fafc", "muted": "#94a3b8",
    "blue": "#93abd9", "green": "#10b981", "red": "#ef4444",
    "blue_hover": "#7d92c3", "green_hover": "#059669", "red_hover": "#dc2626",
    "input_bg": "#131927",
    "hover_bg": "#1e293b",
    "disabled_bg": "#1e293b",
    "disabled_text": "#64748b",
    "red_disabled": "#7f1d1d",
    "red_disabled_text": "#fca5a5",
}

SHADCN_LIGHT = {
    "bg": "#f4f6fa", "card": "#ffffff", "border": "#cbd5e1",
    "text": "#0f172a", "muted": "#64748b",
    "blue": "#364B87", "green": "#059669", "red": "#dc2626",
    "blue_hover": "#2b3c6d", "green_hover": "#047857", "red_hover": "#b91c1c",
    "input_bg": "#ffffff",
    "hover_bg": "#e2e8f0",
    "disabled_bg": "#e2e8f0",
    "disabled_text": "#94a3b8",
    "red_disabled": "#fee2e2",
    "red_disabled_text": "#991b1b",
}


def make_palette(c):
    p = QPalette()
    bg = QColor(c["bg"])
    card = QColor(c["card"])
    text = QColor(c["text"])
    muted = QColor(c["muted"])
    blue = QColor(c["blue"])

    for group in [QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled]:
        p.setColor(group, QPalette.ColorRole.Window, card)
        p.setColor(group, QPalette.ColorRole.WindowText, text)
        p.setColor(group, QPalette.ColorRole.Base, card)
        p.setColor(group, QPalette.ColorRole.AlternateBase, bg)
        p.setColor(group, QPalette.ColorRole.ToolTipBase, card)
        p.setColor(group, QPalette.ColorRole.ToolTipText, text)
        p.setColor(group, QPalette.ColorRole.Text, text)
        p.setColor(group, QPalette.ColorRole.Button, card)
        p.setColor(group, QPalette.ColorRole.ButtonText, text)
        p.setColor(group, QPalette.ColorRole.Highlight, blue)
        p.setColor(group, QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        p.setColor(group, QPalette.ColorRole.PlaceholderText, muted)
    return p


def make_stylesheet(c):
    return f"""
QWidget {{
    background-color: transparent;
    color: {c["text"]};
}}
QMainWindow, QWidget#central {{
    background-color: {c["bg"]};
}}
QLabel {{
    color: {c["text"]};
    font-family: "Segoe UI";
    background-color: transparent;
}}
QLabel.muted {{
    color: {c["muted"]};
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.3px;
}}
QLabel.section-title {{
    font-size: 13px;
    font-weight: 700;
    color: {c["text"]};
    letter-spacing: 0.2px;
}}
QLabel#app-title {{
    font-size: 26px;
    font-weight: 800;
    color: {"#ffffff" if c["bg"] == "#0a0e17" else c["blue"]};
    letter-spacing: -0.5px;
}}
QLabel#app-subtitle {{
    font-size: 14px;
    color: {c["muted"]};
    font-weight: 400;
}}
QLabel#status-label {{
    font-size: 12px;
    font-weight: 700;
}}
QLabel#preview-box {{
    background-color: {c["input_bg"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 8px;
    padding: 8px 10px;
    font-family: "Segoe UI";
    font-size: 11px;
}}
QLineEdit {{
    background-color: {c["input_bg"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 5px 10px;
    font-family: "Segoe UI";
    font-size: 12px;
    height: 20px;
    selection-background-color: {c["blue"]};
    selection-color: #ffffff;
}}
QLineEdit:focus {{
    border-color: {c["blue"]};
    border-width: 2px;
    padding: 4px 9px;
}}
QLineEdit#date-display {{
    font-size: 14px;
    font-weight: 500;
    letter-spacing: 0.5px;
}}

QCalendarWidget {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 10px;
}}
QCalendarWidget QWidget#qt_calendar_navigationbar {{
    background-color: {c["input_bg"]};
    border-bottom: 1px solid {c["border"]};
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    padding: 4px;
}}
QCalendarWidget QToolButton {{
    color: {c["text"]};
    background-color: transparent;
    border: none;
    border-radius: 6px;
    padding: 4px 8px;
    font-weight: 500;
    font-size: 13px;
}}
QCalendarWidget QToolButton:hover {{
    background-color: {c["hover_bg"]};
}}
QCalendarWidget QToolButton:pressed {{
    background-color: {c["border"]};
}}
QCalendarWidget QToolButton::menu-indicator {{
    image: none;
    width: 0px;
}}
QCalendarWidget QSpinBox {{
    color: {c["text"]};
    background-color: {c["card"]};
    border: 1px solid {c["border"]};
    border-radius: 4px;
    padding: 2px 6px;
    font-weight: 500;
}}
QCalendarWidget QMenu {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 4px;
}}
QCalendarWidget QMenu::item:selected {{
    background-color: {c["blue"]};
    color: #ffffff;
    border-radius: 4px;
}}
QCalendarWidget QTableView {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: none;
    selection-background-color: {c["blue"]};
    selection-color: #ffffff;
    font-size: 12px;
    outline: 0;
}}
QCalendarWidget QTableView::item {{
    border-radius: 4px;
    font-weight: normal;
}}
QCalendarWidget QTableView::item:hover {{
    background-color: {c["hover_bg"]};
}}
QCalendarWidget QTableView::item:selected {{
    background-color: {c["blue"]};
    color: #ffffff;
    font-weight: normal;
}}
QCalendarWidget QHeaderView {{
    background-color: {c["card"]};
}}
QCalendarWidget QHeaderView::section {{
    background-color: {c["card"]};
    color: {c["muted"]};
    font-weight: 500;
    font-size: 11px;
    padding: 6px 0px;
    border: none;
    border-bottom: 1px solid {c["border"]};
}}
QCalendarWidget QAbstractItemView:disabled {{
    color: {c["muted"]};
}}
QComboBox {{
    background-color: {c["input_bg"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 5px 10px;
    font-family: "Segoe UI";
    font-size: 12px;
    height: 20px;
}}
QComboBox:focus {{
    border-color: {c["blue"]};
    border-width: 2px;
    padding: 4px 9px;
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox::down-arrow {{
    image: none;
    border: none;
}}
QComboBox QFrame {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
}}
QComboBox QAbstractItemView {{
    background-color: {c["card"]};
    color: {c["text"]};
    selection-background-color: {c["blue"]};
    selection-color: #ffffff;
    border: none;
    outline: none;
    padding: 4px;
}}
QComboBox QAbstractItemView::item {{
    padding: 6px 10px;
    border-radius: 4px;
    min-height: 28px;
    color: {c["text"]};
    background-color: transparent;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {c["hover_bg"]};
    color: {c["text"]};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {c["blue"]};
    color: #ffffff;
}}
QComboBox QAbstractScrollArea {{
    background-color: {c["card"]};
    color: {c["text"]};
}}
QComboBox QScrollBar:vertical {{
    background-color: {c["card"]};
    width: 8px;
    margin: 2px;
    border: none;
}}
QComboBox QScrollBar::handle:vertical {{
    background-color: {c["border"]};
    border-radius: 4px;
    min-height: 20px;
}}
QComboBox QScrollBar::handle:vertical:hover {{
    background-color: {c["muted"]};
}}
QComboBox QScrollBar::add-line:vertical,
QComboBox QScrollBar::sub-line:vertical {{
    height: 0px;
    background: none;
    border: none;
}}
QPushButton {{
    font-family: "Segoe UI";
    font-size: 12px;
    border-radius: 6px;
    padding: 6px 16px;
    border: none;
    font-weight: 600;
}}
QPushButton:disabled {{
    background-color: {c["disabled_bg"]};
    color: {c["disabled_text"]};
    border: none;
}}
QPushButton::icon {{
    margin-right: 4px;
}}
QPushButton#btn-primary {{
    background-color: {c["blue"]};
    color: #ffffff;
    padding: 8px 20px;
    font-size: 13px;
}}
QPushButton#btn-primary:hover {{
    background-color: {c["blue_hover"]};
}}
QPushButton#btn-primary:pressed {{
    background-color: {c["blue_hover"]};
}}
QPushButton#btn-primary:disabled {{
    background-color: {c["disabled_bg"]};
    color: {c["disabled_text"]};
}}
QPushButton#btn-green {{
    background-color: {c["green"]};
    color: #ffffff;
    padding: 8px 20px;
    font-size: 13px;
}}
QPushButton#btn-green:hover {{
    background-color: {c["green_hover"]};
}}
QPushButton#btn-green:pressed {{
    background-color: {c["green_hover"]};
}}
QPushButton#btn-green:disabled {{
    background-color: {c["disabled_bg"]};
    color: {c["disabled_text"]};
}}
QPushButton#btn-red {{
    background-color: {c["red"]};
    color: #ffffff;
    padding: 8px 20px;
    font-size: 13px;
}}
QPushButton#btn-red:hover {{
    background-color: {c["red_hover"]};
}}
QPushButton#btn-red:pressed {{
    background-color: {c["red_hover"]};
}}
QPushButton#btn-red:disabled {{
    background-color: {c["red_disabled"]};
    color: {c["red_disabled_text"]};
    border: none;
}}
QToolTip {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 5px 9px;
    font-family: "Segoe UI";
    font-size: 11px;
}}
QPushButton#btn-outline {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    padding: 6px 12px;
    font-weight: 500;
    border-radius: 6px;
}}
QPushButton#btn-outline:hover {{
    background-color: {c["hover_bg"]};
    border-color: {c["blue"]};
    color: {c["blue"]};
}}
QPushButton#btn-outline:pressed {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-outline:disabled {{
    background-color: {c["card"]};
    color: {c["muted"]};
    border: 1px solid {c["border"]};
}}
QPushButton#btn-ghost {{
    background-color: transparent;
    color: {c["text"]};
    border: none;
    padding: 6px 12px;
    font-weight: 500;
    border-radius: 6px;
}}
QPushButton#btn-ghost:hover {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-ghost:pressed {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-ghost:disabled {{
    color: {c["muted"]};
}}
QPushButton#btn-icon {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 4px;
    min-width: 28px;
    min-height: 28px;
}}
QPushButton#btn-icon:hover {{
    background-color: {c["hover_bg"]};
    color: {c["blue"]};
    border-color: {c["blue"]};
}}
QPushButton#btn-icon:pressed {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-icon:disabled {{
    background-color: {c["card"]};
    color: {c["muted"]};
    border: 1px solid {c["border"]};
}}
QPushButton#btn-add {{
    background-color: {c["card"]};
    color: {c["green"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 4px;
    min-width: 28px;
    min-height: 28px;
}}
QPushButton#btn-add:hover {{
    background-color: {c["hover_bg"]};
    color: {c["green"]};
    border-color: {c["green"]};
}}
QPushButton#btn-add:pressed {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-add:disabled {{
    background-color: {c["card"]};
    color: {c["muted"]};
    border: 1px solid {c["border"]};
}}
QPushButton#btn-del {{
    background-color: {c["card"]};
    color: {c["red"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    padding: 4px;
    min-width: 28px;
    min-height: 28px;
}}
QPushButton#btn-del:hover {{
    background-color: {c["hover_bg"]};
    color: {c["red"]};
    border-color: {c["red"]};
}}
QPushButton#btn-del:pressed {{
    background-color: {c["hover_bg"]};
}}
QPushButton#btn-del:disabled {{
    background-color: {c["card"]};
    color: {c["muted"]};
    border: 1px solid {c["border"]};
}}
QPushButton#theme-btn {{
    background-color: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 8px;
    padding: 4px 10px;
    min-width: 36px;
    min-height: 32px;
    font-size: 15px;
}}
QPushButton#theme-btn:hover {{
    background-color: {c["hover_bg"]};
    border-color: {c["blue"]};
    color: {c["blue"]};
}}
QPushButton#theme-btn:pressed {{
    background-color: {c["hover_bg"]};
}}
QFrame#card {{
    background-color: {c["card"]};
    border: 1px solid {c["border"]};
    border-radius: 10px;
}}
QFrame#divider {{
    background-color: {c["border"]};
    max-height: 1px;
    margin: 4px 0;
}}
QListWidget {{
    background-color: {c["input_bg"]};
    color: {c["text"]};
    border: 1px solid {c["border"]};
    border-radius: 8px;
    font-family: "Segoe UI";
    font-size: 12px;
    outline: none;
    padding: 6px;
}}
QListWidget::item {{
    padding: 7px 10px;
    border-radius: 5px;
    margin: 1px 0;
}}
QListWidget::item:selected {{
    background-color: {c["blue"]};
    color: #ffffff;
}}
QListWidget::item:hover:!selected {{
    background-color: {c["hover_bg"]};
}}
QTextEdit {{
    background-color: {c["input_bg"]};
    color: {c["green"]};
    border: 1px solid {c["border"]};
    border-radius: 8px;
    font-family: "Consolas";
    font-size: 10px;
    padding: 8px;
    selection-background-color: {c["blue"]};
    selection-color: #ffffff;
}}
QProgressBar {{
    background-color: {c["input_bg"]};
    border: 1px solid {c["border"]};
    border-radius: 6px;
    height: 6px;
    text-align: center;
    font-size: 10px;
    color: transparent;
}}
QProgressBar::chunk {{
    background-color: {c["green"]};
    border-radius: 5px;
}}
QSplitter::handle {{
    background-color: {c["border"]};
    width: 1px;
    margin: 0 4px;
}}
QScrollBar:vertical {{
    background-color: transparent;
    width: 6px;
    border: none;
    border-radius: 3px;
}}
QScrollBar::handle:vertical {{
    background-color: {c["border"]};
    border-radius: 3px;
    min-height: 30px;
}}
QScrollBar::handle:vertical:hover {{
    background-color: {c["muted"]};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QMessageBox {{
    background-color: {c["card"]};
}}
QMessageBox QLabel {{
    color: {c["text"]};
    font-size: 12px;
    min-width: 300px;
}}
QMessageBox QPushButton {{
    background-color: {c["blue"]};
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 7px 22px;
    font-size: 12px;
    font-weight: 600;
    min-width: 80px;
}}
QMessageBox QPushButton:hover {{
    background-color: {c["blue_hover"]};
}}
QDialog {{
    background-color: {c["card"]};
    color: {c["text"]};
}}
QDialog QLabel {{
    color: {c["text"]};
}}
QDialog QPushButton {{
    background-color: {c["blue"]};
    color: #ffffff;
    border-radius: 6px;
    padding: 7px 22px;
}}
QDialog QPushButton:hover {{
    background-color: {c["blue_hover"]};
}}
"""


class ReorderableList(QListWidget):
    reordered = pyqtSignal()
    filesDropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setMovement(QListView.Movement.Snap)
        self.setUniformItemSizes(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            file_paths = []
            for url in event.mimeData().urls():
                p = url.toLocalFile()
                if p and os.path.isfile(p) and p.lower().endswith(".docx"):
                    file_paths.append(p)
            if file_paths:
                self.filesDropped.emit(file_paths)
        else:
            super().dropEvent(event)
            self.reordered.emit()


class GenerateWorker(QObject):
    progress_updated = pyqtSignal(float, str)
    log_message = pyqtSignal(str)
    generation_finished = pyqtSignal(list)

    def __init__(self, template_path, sessions, grade, period, name_template, title_template, output_dir):
        super().__init__()
        self.template_path = template_path
        self.sessions = sessions
        self.grade = grade
        self.period = period
        self.name_template = name_template
        self.title_template = title_template
        self.output_dir = output_dir
        self.stop_requested = False

    def run(self):
        from Code import GRADES_INFO, merge_docx_with_guaranteed_header, expand_template
        grade_info = GRADES_INFO[self.grade]
        clean_grade = self.grade.replace("\u00b0", "")
        results = []

        target_subs = [
            (s, sub) for s in self.sessions for sub in s["subsessions"] if sub["files"]
        ]
        total_tasks = len(target_subs)
        completed_tasks = 0
        cur_offset = 0
        current_session_name = None

        for s, sub in target_subs:
            if self.stop_requested:
                results.append(("stopped", "Proceso detenido por el usuario."))
                break

            if current_session_name != s["name"]:
                current_session_name = s["name"]
                cur_offset = 0

            dia = s["day"].zfill(2)
            fecha = f"{dia}/{s['month']}/{s['year']}"
            sub_code = sub["name"].replace("Subsesi\u00f3n ", "S")

            name_context = {
                "grade": self.grade,
                "period": self.period,
                "session": sub_code,
                "year": s["year"],
                "level": grade_info["level"],
            }
            filename = expand_template(self.name_template, name_context) + ".docx"

            clean_num = self.grade.replace("\u00b0", "").strip()
            if grade_info["multicourse"]:
                if clean_num in ("10", "11"):
                    p_c_val = f"{clean_num}-___"
                else:
                    p_c_val = f"{clean_num}0___"
            else:
                p_c_val = self.grade

            config = {
                "level": grade_info["level"],
                "grade": self.grade,
                "grade_clean": self.grade,
                "grade_display": self.grade,
                "period": self.period,
                "session_code": sub_code,
                "date": fecha,
                "year": s["year"],
                "p_c_value": p_c_val,
                "title_template": self.title_template,
            }
            out = filename
            out_dir = self.output_dir.strip()
            if out_dir:
                out_dir = os.path.abspath(out_dir)
                os.makedirs(out_dir, exist_ok=True)
                out = os.path.join(out_dir, out)
            else:
                out = os.path.join(os.getcwd(), out)

            try:
                last_num = merge_docx_with_guaranteed_header(
                    self.template_path, sub["files"], out, config, cur_offset
                )
                cur_offset = last_num
                results.append(("ok", out))
            except Exception as e:
                results.append(("error", f"{out}: {e}"))
                break

            completed_tasks += 1
            progress_ratio = completed_tasks / total_tasks
            pct = int(progress_ratio * 100)
            self.progress_updated.emit(progress_ratio, f"{pct}%")

        self.generation_finished.emit(results)


class UpdateWorker(QThread):
    finished_signal = pyqtSignal(bool, str)

    def run(self):
        project_dir = os.path.dirname(os.path.abspath(__file__))
        git_dir = os.path.join(project_dir, ".git")
        if not os.path.exists(git_dir):
            self.finished_signal.emit(
                False,
                "Modo independiente.\nPara actualizar autom\u00e1ticamente desde GitHub, clona el repositorio con Git o ejecuta start.bat."
            )
            return

        try:
            res = subprocess.run(
                ["git", "pull", "origin", "main"],
                cwd=project_dir,
                capture_output=True,
                text=True,
                timeout=12
            )
            out = (res.stdout or res.stderr or "").strip()
            if "Already up to date" in out or "Ya est\u00e1 actualizado" in out or "Already up-to-date" in out:
                self.finished_signal.emit(True, "La aplicaci\u00f3n ya est\u00e1 en la \u00faltima versi\u00f3n de GitHub.")
            else:
                self.finished_signal.emit(True, f"Se han descargado las \u00faltimas novedades de GitHub:\n\n{out[:300]}")
        except Exception as e:
            self.finished_signal.emit(False, f"No se pudo consultar GitHub: {e}")


class DesktopApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GESA \u2014 Gestor de Evaluaciones de Suficiencia Acad\u00e9mica")
        self.setMinimumSize(1100, 720)

        icon_file = get_resource_path("app_icon.ico")
        if not os.path.exists(icon_file):
            icon_file = get_resource_path("app_icon.png")
        if os.path.exists(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        self.settings = QSettings("GESA", "AcademicManager")
        saved_theme = str(self.settings.value("theme", "system"))
        self._theme = saved_theme if saved_theme in ["dark", "light", "system"] else "system"
        self._icon_refs = []

        self._undo_stack = []
        self._redo_stack = []
        self._setup_shortcuts()

        self.template_path = ""
        self.output_dir = ""
        self.grade = ""
        self.period = "P3"
        self.name_template = "E.S.A._{grade}_{period}_{session}_{year}"
        self.title_template = "Evaluaciones de Suficiencia Académica - {grade} - {period} - {session} - {year}"

        self.sessions = []
        self._selected_session = -1
        self._selected_sub = -1

        self.processing = False
        self._stop_requested = False

        self._worker = None
        self._worker_thread = None

        self._build()
        self._add_session()
        self._apply_theme()

        icon_file = get_resource_path("app_icon.png")
        if not os.path.exists(icon_file):
            icon_file = get_resource_path("app_icon.ico")
        if os.path.exists(icon_file):
            app_icon = QIcon(icon_file)
            self.setWindowIcon(app_icon)
            app = QApplication.instance()
            if app:
                app.setWindowIcon(app_icon)

        self.showMaximized()

    # ─── undo / redo ───────────────────────────────────────────

    def _setup_shortcuts(self):
        self.shortcut_undo = QShortcut(QKeySequence("Ctrl+Z"), self)
        self.shortcut_undo.activated.connect(self._undo)

        self.shortcut_redo_y = QShortcut(QKeySequence("Ctrl+Y"), self)
        self.shortcut_redo_y.activated.connect(self._redo)

        self.shortcut_redo_shift = QShortcut(QKeySequence("Ctrl+Shift+Z"), self)
        self.shortcut_redo_shift.activated.connect(self._redo)

    def _save_state_for_undo(self):
        state = {
            "sessions": copy.deepcopy(self.sessions),
            "selected_session": self._selected_session,
            "selected_sub": self._selected_sub,
            "grade": self.grade_combo.currentText() if hasattr(self, "grade_combo") else "",
            "period": self.period_combo.currentText() if hasattr(self, "period_combo") else "P3",
            "name_template": self.name_template,
            "title_template": self.title_template,
            "template_path": self.template_path,
            "output_dir": self.output_dir,
        }
        self._undo_stack.append(state)
        if len(self._undo_stack) > 50:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def _apply_state_dict(self, state):
        self.sessions = copy.deepcopy(state["sessions"])
        self._selected_session = state["selected_session"]
        self._selected_sub = state["selected_sub"]
        self.template_path = state["template_path"]
        self.output_dir = state["output_dir"]
        self.name_template = state["name_template"]
        self.title_template = state["title_template"]

        if hasattr(self, "template_entry"):
            self.template_entry.setText(self.template_path)
            self.template_entry.setCursorPosition(0)
        if hasattr(self, "output_entry"):
            self.output_entry.setText(self.output_dir)
            self.output_entry.setCursorPosition(0)
        if hasattr(self, "name_template_entry"):
            self.name_template_entry.setText(self.name_template)
            self.name_template_entry.setCursorPosition(0)
        if hasattr(self, "title_template_entry"):
            self.title_template_entry.setText(self.title_template)
            self.title_template_entry.setCursorPosition(0)
        if hasattr(self, "grade_combo") and state["grade"]:
            self.grade_combo.setCurrentText(state["grade"])
        if hasattr(self, "period_combo") and state["period"]:
            self.period_combo.setCurrentText(state["period"])

        self._refresh_sessions()
        self._update_preview()

    def _undo(self):
        if not self._undo_stack:
            self._log("Nada que deshacer (Ctrl+Z)")
            return
        current_state = {
            "sessions": copy.deepcopy(self.sessions),
            "selected_session": self._selected_session,
            "selected_sub": self._selected_sub,
            "grade": self.grade_combo.currentText() if hasattr(self, "grade_combo") else "",
            "period": self.period_combo.currentText() if hasattr(self, "period_combo") else "P3",
            "name_template": self.name_template,
            "title_template": self.title_template,
            "template_path": self.template_path,
            "output_dir": self.output_dir,
        }
        self._redo_stack.append(current_state)
        prev_state = self._undo_stack.pop()
        self._apply_state_dict(prev_state)
        self._log("\u21a9 Deshecho (Ctrl+Z)")

    def _redo(self):
        if not self._redo_stack:
            self._log("Nada que rehacer (Ctrl+Y)")
            return
        current_state = {
            "sessions": copy.deepcopy(self.sessions),
            "selected_session": self._selected_session,
            "selected_sub": self._selected_sub,
            "grade": self.grade_combo.currentText() if hasattr(self, "grade_combo") else "",
            "period": self.period_combo.currentText() if hasattr(self, "period_combo") else "P3",
            "name_template": self.name_template,
            "title_template": self.title_template,
            "template_path": self.template_path,
            "output_dir": self.output_dir,
        }
        self._undo_stack.append(current_state)
        next_state = self._redo_stack.pop()
        self._apply_state_dict(next_state)
        self._log("\u21aa Rehacer (Ctrl+Y)")

    # ─── data model ────────────────────────────────────────────

    def _session(self, idx=None):
        i = idx if idx is not None else self._selected_session
        if 0 <= i < len(self.sessions):
            return self.sessions[i]
        return None

    def _sub(self, s_idx=None, sub_idx=None):
        s = self._session(s_idx)
        if s is None:
            return None
        i = sub_idx if sub_idx is not None else self._selected_sub
        if 0 <= i < len(s["subsessions"]):
            return s["subsessions"][i]
        return None

    def _refresh_sessions(self):
        self.sessions_lb.clear()
        for s in self.sessions:
            item = QListWidgetItem(f"  {s['name']}")
            self.sessions_lb.addItem(item)
        if self.sessions:
            if self._selected_session >= len(self.sessions):
                self._selected_session = len(self.sessions) - 1
            self.sessions_lb.setCurrentRow(self._selected_session)
        self._refresh_subs()

    def _refresh_subs(self):
        self.subs_lb.clear()
        s = self._session()
        if s:
            for sub in s["subsessions"]:
                self.subs_lb.addItem(QListWidgetItem(f"  {sub['name']}"))
            if self._selected_sub >= len(s["subsessions"]):
                self._selected_sub = len(s["subsessions"]) - 1
            if s["subsessions"]:
                self.subs_lb.setCurrentRow(self._selected_sub)
            d = int(s["day"])
            m = MONTH_MAP.get(s["month"], 1)
            y = int(s["year"])
            self.date_edit.dateChanged.disconnect(self._on_date_changed)
            self.date_edit.setDate(QDate(y, m, d))
            self.date_edit.dateChanged.connect(self._on_date_changed)
        self._refresh_files()
        self._update_preview()

    def _refresh_files(self):
        self.files_lb.clear()
        sub = self._sub()
        if sub:
            for f in sub["files"]:
                item = QListWidgetItem(f"  {os.path.basename(f)}")
                item.setToolTip(f)
                self.files_lb.addItem(item)

    def _open_file_location(self, item=None):
        sub = self._sub()
        if not sub or not sub["files"]:
            self._show_info("Aviso", "No hay archivos en la subsesi\u00f3n seleccionada.")
            return
        row = self.files_lb.currentRow()
        if row < 0 or row >= len(sub["files"]):
            row = 0
        filepath = sub["files"][row]
        if os.path.exists(filepath):
            subprocess.run(f'explorer /select,"{os.path.abspath(filepath)}"', shell=True)
        elif os.path.exists(os.path.dirname(filepath)):
            os.startfile(os.path.dirname(filepath))
        else:
            self._show_warning("Archivo no encontrado", f"No se encontr\u00f3 el archivo:\n{filepath}")

    # ─── add / remove session / sub ────────────────────────────

    def _add_session(self):
        self._save_state_for_undo()
        n = len(self.sessions) + 1
        if self.sessions:
            last_s = self.sessions[-1]
            try:
                y = int(last_s["year"])
                m = MONTH_MAP.get(last_s["month"], 1)
                d = int(last_s["day"])
                qd = QDate(y, m, d).addDays(1)
            except Exception:
                qd = self.date_edit.date()
        else:
            qd = self.date_edit.date()

        self.sessions.append({
            "name": f"Sesi\u00f3n {n}",
            "day": str(qd.day()),
            "month": MONTH_REV[qd.month()],
            "year": str(qd.year()),
            "subsessions": [],
        })
        self._selected_session = len(self.sessions) - 1
        self._add_sub(save_undo=False)
        self._refresh_sessions()

    def _remove_sessions(self):
        rows = sorted(set(i.row() for i in self.sessions_lb.selectedIndexes()), reverse=True)
        if not rows:
            return
        if len(rows) == len(self.sessions):
            self._show_info("Aviso", "Debe haber al menos una sesi\u00f3n.")
            return

        count = len(rows)
        msg = f"\u00bfDeseas eliminar la sesi\u00f3n seleccionada?" if count == 1 else f"\u00bfDeseas eliminar las {count} sesiones seleccionadas?"
        if not self._ask_yes_no("Confirmar eliminaci\u00f3n", msg):
            return

        self._save_state_for_undo()
        for r in rows:
            self.sessions.pop(r)
        self._selected_session = min(self._selected_session, len(self.sessions) - 1)
        self._selected_sub = -1
        self._refresh_sessions()

    def _add_sub(self, save_undo=True):
        if save_undo:
            self._save_state_for_undo()
        s = self._session()
        if s is None:
            return
        n = len(s["subsessions"]) + 1
        s["subsessions"].append({
            "name": f"Subsesi\u00f3n {s['name'].split()[-1]}.{n}",
            "files": [],
        })
        self._selected_sub = len(s["subsessions"]) - 1
        self._refresh_subs()

    def _remove_subs(self):
        s = self._session()
        if s is None:
            return
        rows = sorted(set(i.row() for i in self.subs_lb.selectedIndexes()), reverse=True)
        if not rows:
            return
        if len(rows) == len(s["subsessions"]):
            self._show_info("Aviso", "Debe haber al menos una subsesi\u00f3n.")
            return

        count = len(rows)
        msg = f"\u00bfDeseas eliminar la subsesi\u00f3n seleccionada?" if count == 1 else f"\u00bfDeseas eliminar las {count} subsesiones seleccionadas?"
        if not self._ask_yes_no("Confirmar eliminaci\u00f3n", msg):
            return

        self._save_state_for_undo()
        for r in rows:
            s["subsessions"].pop(r)
        self._selected_sub = min(self._selected_sub, len(s["subsessions"]) - 1) if s["subsessions"] else -1
        self._refresh_subs()

    # ─── select session / sub ──────────────────────────────────

    def _update_preview(self, sub_name=None):
        if not hasattr(self, "preview_box"):
            return
        from Code import GRADES_INFO, expand_template
        s = self._session()
        sub = self._sub()

        grade_text = self.grade_combo.currentText() if hasattr(self, "grade_combo") else ""
        period_text = self.period_combo.currentText() if hasattr(self, "period_combo") else ""
        grade = "{grade}" if (not grade_text or grade_text.startswith("Seleccionar")) else grade_text
        period = "{period}" if (not period_text or period_text.startswith("Seleccionar")) else period_text
        session_name = s["name"] if s else "Sesi\u00f3n 1"
        target_sub_name = sub_name if sub_name else (sub["name"] if sub else "Subsesi\u00f3n 1.1")
        sub_code = target_sub_name.replace("Subsesi\u00f3n ", "S").replace("Subsesion ", "S")
        year = s["year"] if s else "2026"
        day = s["day"] if s else "15"
        month = s["month"] if s else "SEP"
        level = GRADES_INFO.get(grade_text, {}).get("level", "{level}")
        ctx = {
            "grade": grade,
            "period": period,
            "session": sub_code,
            "year": year,
            "level": level,
            "day": day,
            "month": month,
        }

        fn = expand_template(self.name_template, ctx)
        if not fn.lower().endswith(".docx"):
            fn += ".docx"

        total_files = 0
        if target_sub_name and hasattr(self, 'sessions') and self.sessions:
            for sess_obj in self.sessions:
                subs_list = sess_obj.get("subsessions", []) if isinstance(sess_obj, dict) else getattr(sess_obj, "subsessions", [])
                for sub_obj in subs_list:
                    sub_name_val = sub_obj.get("name", "") if isinstance(sub_obj, dict) else getattr(sub_obj, "name", "")
                    sub_files_val = sub_obj.get("files", []) if isinstance(sub_obj, dict) else getattr(sub_obj, "files", [])
                    if sub_name_val == target_sub_name or sub_name_val.replace("Subsesión ", "S").replace("Subsesion ", "S") == sub_code:
                        total_files = len(sub_files_val)
                        break

        eval_prefix = "Evaluación" if total_files <= 1 else "Evaluaciones"
        tpl_title = self.title_template
        if tpl_title.startswith("Evaluación de") or tpl_title.startswith("Evaluaciones de"):
            tpl_title = re.sub(r'^Evaluaci\u00f3n(es)?\s+de', f'{eval_prefix} de', tpl_title)

        dt = expand_template(tpl_title, ctx)

        self.preview_box.setText(
            f"<b>\ud83d\udcc4 Archivo:</b> {fn}<br>"
            f"<b>\ud83d\udcdd T\u00edtulo:</b> {dt}"
        )

    def _on_sub_hover(self, item):
        if item and item.text().strip():
            self._update_preview(sub_name=item.text().strip())

    def _on_name_template_changed(self, v):
        self.name_template = v
        self._update_preview()

    def _on_title_template_changed(self, v):
        self.title_template = v
        self._update_preview()

    def _on_select_session(self, row):
        if row >= 0:
            self._selected_session = row
            self._selected_sub = -1
            self._refresh_subs()

    def _on_select_sub(self, row):
        if row >= 0:
            self._selected_sub = row
            self._refresh_files()
            self._update_preview()

    # ─── drag-drop reorder sync ────────────────────────────────

    def _on_sessions_reordered(self):
        old_map = {s["name"]: s for s in self.sessions}
        new_list = []
        for i in range(self.sessions_lb.count()):
            name = self.sessions_lb.item(i).text().strip()
            if name in old_map:
                new_list.append(old_map[name])
        if new_list:
            self.sessions = new_list
        self._selected_session = self.sessions_lb.currentRow()
        self._refresh_subs()

    def _on_subs_reordered(self):
        s = self._session()
        if s is None:
            return
        old_map = {sub["name"]: sub for sub in s["subsessions"]}
        new_list = []
        for i in range(self.subs_lb.count()):
            name = self.subs_lb.item(i).text().strip()
            if name in old_map:
                new_list.append(old_map[name])
        if new_list:
            s["subsessions"] = new_list
        self._selected_sub = self.subs_lb.currentRow()
        self._refresh_files()

    def _on_files_reordered(self):
        sub = self._sub()
        if sub is None:
            return
        name_to_path = {os.path.basename(f): f for f in sub["files"]}
        new_list = []
        for i in range(self.files_lb.count()):
            name = self.files_lb.item(i).text().strip()
            if name in name_to_path:
                new_list.append(name_to_path[name])
        if new_list:
            sub["files"] = new_list

    # ─── file management ───────────────────────────────────────

    def _add_files_from_paths(self, paths):
        sub = self._sub()
        if sub is None:
            self._show_info("Aviso", "Selecciona una subsesi\u00f3n primero.")
            return
        if not paths:
            return

        if len(paths) > 1:
            paths = list(reversed(paths))

        self._save_state_for_undo()
        chosen_dir = os.path.dirname(paths[0])
        resolved_any = False

        for p in paths:
            if p not in sub["files"]:
                sub["files"].append(p)

        for s in self.sessions:
            for sb in s["subsessions"]:
                for i, f in enumerate(sb["files"]):
                    if os.path.exists(f):
                        continue
                    candidate = os.path.join(chosen_dir, os.path.basename(f))
                    if candidate != f and os.path.exists(candidate) and candidate not in sb["files"]:
                        sb["files"][i] = candidate
                        resolved_any = True

        if resolved_any:
            self._log(f"Archivos adicionales reubicados desde: {chosen_dir}")

        self._refresh_files()
        if resolved_any:
            self._refresh_sessions()

    def _add_files(self):
        sub = self._sub()
        if sub is None:
            self._show_info("Aviso", "Selecciona una subsesi\u00f3n primero.")
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar ex\u00e1menes",
            "", "Word (*.docx);;Todos (*.*)")
        if paths:
            self._add_files_from_paths(paths)

    def _move_file(self, direction):
        sub = self._sub()
        if sub is None:
            return
        row = self.files_lb.currentRow()
        if row < 0:
            return
        new = row + direction
        if not 0 <= new < len(sub["files"]):
            return
        sub["files"][row], sub["files"][new] = sub["files"][new], sub["files"][row]
        self._refresh_files()
        self.files_lb.setCurrentRow(new)

    def _remove_files(self):
        sub = self._sub()
        if sub is None:
            return
        rows = sorted(set(i.row() for i in self.files_lb.selectedIndexes()))
        if not rows:
            return

        count = len(rows)
        msg = f"\u00bfDeseas eliminar el archivo seleccionado?" if count == 1 else f"\u00bfDeseas eliminar los {count} archivos seleccionados?"
        if not self._ask_yes_no("Confirmar eliminaci\u00f3n", msg):
            return

        self._save_state_for_undo()
        for row in reversed(rows):
            sub["files"].pop(row)
        self._refresh_files()

    def _on_date_changed(self, qdate=None):
        s = self._session()
        if s is not None and qdate is not None:
            s["day"] = str(qdate.day())
            s["month"] = MONTH_REV[qdate.month()]
            s["year"] = str(qdate.year())

    # ─── theme ─────────────────────────────────────────────────

    def _detect_system_theme(self):
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            val = winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
            winreg.CloseKey(key)
            return "light" if val == 1 else "dark"
        except Exception:
            return "dark"

    def _check_updates(self):
        if hasattr(self, "_update_thread") and self._update_thread and self._update_thread.isRunning():
            return

        if hasattr(self, "update_btn"):
            self.update_btn.setEnabled(False)

        self._log("\ud83d\udd04 Consultando actualizaciones en GitHub...")

        self._update_thread = UpdateWorker()
        self._update_thread.finished_signal.connect(self._on_update_finished)
        self._update_thread.start()

    def _on_update_finished(self, success, message):
        if hasattr(self, "update_btn"):
            self.update_btn.setEnabled(True)

        if success:
            first_line = message.splitlines()[0] if message else "Sincronizaci\u00f3n completada."
            self._log(f"\u2713 {first_line}")
            self._show_info("Actualizaciones", message)
        else:
            self._show_warning("Actualizaciones", message)

    def _effective_theme(self):
        if self._theme == "system":
            return self._detect_system_theme()
        return self._theme

    def _theme_colors(self):
        return SHADCN_DARK if self._effective_theme() == "dark" else SHADCN_LIGHT

    def _apply_theme(self):
        c = self._theme_colors()
        app = QApplication.instance()
        if app:
            app.setPalette(make_palette(c))
        self.setStyleSheet(make_stylesheet(c))
        is_dark = (self._effective_theme() == "dark")
        set_window_dark_mode(self.winId(), is_dark)
        if self._theme == "dark":
            icon_name = "fa5s.moon"
            tip = "Tema: Oscuro"
        elif self._theme == "light":
            icon_name = "fa5s.sun"
            tip = "Tema: Claro"
        else:
            eff = self._effective_theme().capitalize()
            icon_name = "fa5s.desktop"
            tip = f"Tema: Sistema ({eff})"
        self.theme_btn.setIcon(self._make_icon(icon_name, "text"))
        self.theme_btn.setToolTip(f"{tip} | Clic: {self._next_theme_name()}")
        logo_color = "#ffffff" if is_dark else c["blue"]
        if hasattr(self, "logo_icon_lbl"):
            self.logo_icon_lbl.setPixmap(self._get_svg_logo_pixmap(logo_color, 26))
        self._refresh_icons()
        self.date_edit.btn.setIcon(self._make_icon("fa5s.calendar-alt", "muted"))
        for w in [self.date_edit.display, self.date_edit.btn]:
            w.style().unpolish(w)
            w.style().polish(w)
        lb_bg = c["input_bg"]
        lb_fg = c["text"]
        lb_sel = c["blue"]
        lb_border = c["border"]
        lb_card = c["hover_bg"]
        for lb in [self.sessions_lb, self.subs_lb, self.files_lb]:
            lb.setStyleSheet(f"""
                QListWidget {{
                    background-color: {lb_bg};
                    color: {lb_fg};
                    border: 1px solid {lb_border};
                    border-radius: 8px;
                    font-family: "Segoe UI";
                    font-size: 12px;
                    outline: none;
                    padding: 6px;
                }}
                QListWidget::item {{
                    padding: 7px 10px;
                    border-radius: 5px;
                    margin: 1px 0;
                }}
                QListWidget::item:selected {{
                    background-color: {lb_sel};
                    color: #ffffff;
                }}
                QListWidget::item:hover:!selected {{
                    background-color: {lb_card};
                }}
            """)

    def _get_svg_logo_pixmap(self, color_hex, size=26):
        svg_data = f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200">
          <rect x="24" y="14" width="152" height="172" rx="28" fill="none" stroke="{color_hex}" stroke-width="16"/>
          <path d="M 40 142 L 160 142" fill="none" stroke="{color_hex}" stroke-width="16" stroke-linecap="round"/>
          <path d="M 68 84 L 94 108 L 138 60" fill="none" stroke="{color_hex}" stroke-width="18" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>"""
        renderer = QSvgRenderer(svg_data.encode('utf-8'))
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return pixmap

    def _make_icon(self, name, color_key="muted"):
        c = self._theme_colors()
        color = c.get(color_key, color_key)
        return _qta().icon(name, color=color)

    def _reg_icon(self, btn, icon_name, color_key="muted"):
        self._icon_refs.append((btn, icon_name, color_key))
        return btn

    def _refresh_icons(self):
        for btn, name, color_key in self._icon_refs:
            btn.setIcon(self._make_icon(name, color_key))

    def _make_icon_btn(self, icon_name, obj_name, tooltip, color_key="muted"):
        btn = QPushButton()
        btn.setObjectName(obj_name)
        btn.setToolTip(tooltip)
        self._reg_icon(btn, icon_name, color_key)
        return btn



    def _next_theme_name(self):
        order = ["dark", "light", "system"]
        names = {"dark": "Oscuro", "light": "Claro", "system": "Sistema"}
        try:
            idx = order.index(self._theme)
            return names[order[(idx + 1) % len(order)]]
        except ValueError:
            return "Oscuro"

    def _cycle_theme(self):
        order = ["dark", "light", "system"]
        try:
            idx = order.index(self._theme)
            self._theme = order[(idx + 1) % len(order)]
        except ValueError:
            self._theme = "dark"
        if hasattr(self, "settings"):
            self.settings.setValue("theme", self._theme)
        self._apply_theme()

    # ─── styled dialogs ────────────────────────────────────────

    def _show_info(self, title, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle(title)
        dlg.setText(message)
        c = self._theme_colors()
        try:
            ico = _qta().icon("fa5s.check-circle", color=c["green"])
            dlg.setIconPixmap(ico.pixmap(48, 48))
        except Exception:
            dlg.setIcon(QMessageBox.Icon.Information)
        btn = dlg.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        set_window_dark_mode(dlg.winId(), self._effective_theme() == "dark")
        dlg.setStyleSheet(f"QMessageBox {{ background-color: {c['card']}; color: {c['text']}; }} QMessageBox QLabel {{ color: {c['text']}; font-family: 'Segoe UI'; font-size: 13px; }}")
        btn.setStyleSheet(f"QPushButton {{ background-color: {c['blue']}; color: #ffffff; border: none; border-radius: 6px; padding: 7px 22px; font-weight: 600; min-width: 75px; }} QPushButton:hover {{ background-color: {c['blue_hover']}; }}")
        dlg.setDefaultButton(btn)
        dlg.exec()

    def _show_warning(self, title, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle(title)
        dlg.setText(message)
        c = self._theme_colors()
        try:
            ico = _qta().icon("fa5s.exclamation-triangle", color="#f59e0b" if self._effective_theme() == "dark" else "#d97706")
            dlg.setIconPixmap(ico.pixmap(48, 48))
        except Exception:
            dlg.setIcon(QMessageBox.Icon.Warning)
        btn = dlg.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        set_window_dark_mode(dlg.winId(), self._effective_theme() == "dark")
        dlg.setStyleSheet(f"QMessageBox {{ background-color: {c['card']}; color: {c['text']}; }} QMessageBox QLabel {{ color: {c['text']}; font-family: 'Segoe UI'; font-size: 13px; }}")
        btn.setStyleSheet(f"QPushButton {{ background-color: {c['blue']}; color: #ffffff; border: none; border-radius: 6px; padding: 7px 22px; font-weight: 600; min-width: 75px; }} QPushButton:hover {{ background-color: {c['blue_hover']}; }}")
        dlg.setDefaultButton(btn)
        dlg.exec()

    def _show_error(self, title, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle(title)
        dlg.setText(message)
        dlg.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        c = self._theme_colors()
        try:
            ico = _qta().icon("fa5s.times-circle", color=c["red"])
            dlg.setIconPixmap(ico.pixmap(48, 48))
        except Exception:
            dlg.setIcon(QMessageBox.Icon.Critical)
        btn_copy = dlg.addButton("Copiar Error", QMessageBox.ButtonRole.ActionRole)
        btn = dlg.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        set_window_dark_mode(dlg.winId(), self._effective_theme() == "dark")
        dlg.setStyleSheet(f"QMessageBox {{ background-color: {c['card']}; color: {c['text']}; }} QMessageBox QLabel {{ color: {c['text']}; font-family: 'Segoe UI'; font-size: 13px; }}")
        btn.setStyleSheet(f"QPushButton {{ background-color: {c['red']}; color: #ffffff; border: none; border-radius: 6px; padding: 7px 22px; font-weight: 600; min-width: 75px; }} QPushButton:hover {{ background-color: {c['red_hover']}; }}")
        btn_copy.setStyleSheet(f"QPushButton {{ background-color: transparent; color: {c['text']}; border: 1px solid {c['border']}; border-radius: 6px; padding: 7px 18px; font-weight: 500; }} QPushButton:hover {{ background-color: {c['bg']}; }}")
        dlg.setDefaultButton(btn)
        dlg.exec()
        if dlg.clickedButton() == btn_copy:
            QApplication.clipboard().setText(message)

    def _ask_yes_no(self, title, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle(title)
        dlg.setText(message)
        c = self._theme_colors()
        try:
            ico = _qta().icon("fa5s.question-circle", color=c["blue"])
            dlg.setIconPixmap(ico.pixmap(48, 48))
        except Exception:
            dlg.setIcon(QMessageBox.Icon.Question)
        yes_btn = dlg.addButton("  S\u00ed  ", QMessageBox.ButtonRole.YesRole)
        no_btn = dlg.addButton("  No  ", QMessageBox.ButtonRole.NoRole)
        
        c = self._theme_colors()
        set_window_dark_mode(dlg.winId(), self._effective_theme() == "dark")
        
        dlg.setStyleSheet(f"QMessageBox {{ background-color: {c['card']}; color: {c['text']}; }} QMessageBox QLabel {{ color: {c['text']}; font-family: 'Segoe UI'; font-size: 13px; }}")
        
        yes_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {c["blue"]};
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 7px 22px;
                font-weight: 600;
                min-width: 75px;
            }}
            QPushButton:hover {{
                background-color: {c["blue_hover"]};
            }}
        """)
        
        no_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {c["text"]};
                border: 1px solid {c["border"]};
                border-radius: 6px;
                padding: 7px 22px;
                font-weight: 500;
                min-width: 75px;
            }}
            QPushButton:hover {{
                background-color: {c["hover_bg"]};
                border-color: {c["text"]};
            }}
        """)

        dlg.setDefaultButton(yes_btn)
        dlg.exec()
    def _show_shortcodes_help(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Etiquetas y Shortcodes — GESA")
        dlg.setMinimumWidth(560)
        
        is_dark = (self._effective_theme() == "dark")
        set_window_dark_mode(dlg.winId(), is_dark)
        c = self._theme_colors()

        dlg.setStyleSheet(f"""
            QDialog {{
                background-color: {c["card"]};
                color: {c["text"]};
            }}
            QLabel {{
                color: {c["text"]};
                font-family: "Segoe UI";
            }}
        """)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        title = QLabel("Etiquetas y Shortcodes Disponibles")
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {c['blue']};")
        layout.addWidget(title)

        desc = QLabel("Usa estas etiquetas dentro de tu plantilla Word maestra (.docx) o en las plantillas de salida para automatizar la personalizaci\u00f3n:")
        desc.setWordWrap(True)
        desc.setStyleSheet(f"font-size: 12px; color: {c['muted']};")
        layout.addWidget(desc)

        content = QTextEdit()
        content.setReadOnly(True)
        content.setHtml(f"""
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; font-size: 12px; color: {c['text']}; line-height: 1.4; }}
            h4 {{ margin-top: 10px; margin-bottom: 6px; color: {c['blue']}; font-size: 13px; }}
            p {{ margin-top: 2px; margin-bottom: 8px; font-size: 11px; }}
            table {{ border-collapse: collapse; width: 100%; margin-bottom: 12px; }}
            th {{ background-color: {c['hover_bg']}; color: {c['text']}; padding: 6px 8px; text-align: left; font-size: 11px; border: 1px solid {c['border']}; }}
            td {{ padding: 6px 8px; border: 1px solid {c['border']}; font-size: 11px; }}
            code {{ font-family: 'Consolas', monospace; font-size: 11px; background-color: {c['hover_bg']}; padding: 2px 6px; border-radius: 4px; color: {c['blue']}; font-weight: bold; }}
        </style>

        <h4>📄 1. Etiquetas para la Plantilla Maestra en Word (.docx)</h4>
        <p>Escribe estas etiquetas en el cuerpo o encabezado de tu documento Word:</p>
        <table>
            <tr><th>Etiqueta</th><th>Descripción</th><th>Ejemplo de Salida</th></tr>
            <tr><td><code>(GRADE)</code></td><td>Grado del examen</td><td><b>6°</b></td></tr>
            <tr><td><code>(P_C)</code></td><td>Casilla de llenado de curso</td><td><b>60___</b> (6°-9°), <b>10-___</b> (10°-11°)</td></tr>
            <tr><td><code>(EDU_LEVEL)</code></td><td>Nivel educativo</td><td><b>BÁSICA SECUNDARIA</b></td></tr>
            <tr><td><code>(TERM)</code></td><td>Periodo académico</td><td><b>P3</b></td></tr>
            <tr><td><code>(SESSION)</code></td><td>Código de subsesión</td><td><b>S1.1</b></td></tr>
            <tr><td><code>(DATE)</code></td><td>Fecha de aplicación</td><td><b>15/SEP/2026</b></td></tr>
        </table>

        <h4>🏷️ 2. Etiquetas para Nombre de Archivo y Título del Documento</h4>
        <p>Usa estas etiquetas en los campos de entrada de la barra lateral:</p>
        <table>
            <tr><th>Etiqueta</th><th>Descripción</th><th>Ejemplo</th></tr>
            <tr><td><code>{{grade}}</code></td><td>Grado asignado</td><td>6° / Prejardín</td></tr>
            <tr><td><code>{{period}}</code></td><td>Periodo académico</td><td>P3</td></tr>
            <tr><td><code>{{session}}</code></td><td>Código de subsesión</td><td>S1.1</td></tr>
            <tr><td><code>{{year}}</code></td><td>Año configurado</td><td>2026</td></tr>
            <tr><td><code>{{level}}</code></td><td>Nivel educativo</td><td>Básica Secundaria</td></tr>
            <tr><td><code>{{day}}</code></td><td>Día de aplicación</td><td>15</td></tr>
            <tr><td><code>{{month}}</code></td><td>Mes abreviado</td><td>SEP</td></tr>
        </table>
        """)
        layout.addWidget(content)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        close_btn = QPushButton("Entendido")
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {c["blue"]};
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 7px 24px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {c["blue_hover"]};
            }}
        """)
        close_btn.clicked.connect(dlg.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        dlg.exec()

    # ─── build UI ──────────────────────────────────────────────

    def _build(self):
        from Code import GRADES_INFO
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(14)

        # ─── Header ─────────────────────────────────────────────
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(10)

        self.logo_icon_lbl = QLabel()
        self.logo_icon_lbl.setFixedSize(26, 26)
        header_layout.addWidget(self.logo_icon_lbl)

        title = QLabel("GESA")
        title.setObjectName("app-title")
        header_layout.addWidget(title)

        subtitle = QLabel("\u2014  Gestor de Evaluaciones de Suficiencia Acad\u00e9mica")
        subtitle.setObjectName("app-subtitle")
        header_layout.addWidget(subtitle)

        header_layout.addStretch()

        self.update_btn = self._make_icon_btn("fa5s.sync-alt", "theme-btn", "Buscar actualizaciones desde GitHub", "text")
        self.update_btn.clicked.connect(self._check_updates)
        header_layout.addWidget(self.update_btn)

        self.theme_btn = QPushButton()
        self.theme_btn.setObjectName("theme-btn")
        self.theme_btn.clicked.connect(self._cycle_theme)
        header_layout.addWidget(self.theme_btn)

        main_layout.addWidget(header)

        # ─── Panels splitter ────────────────────────────────────
        panels = QSplitter(Qt.Orientation.Horizontal)
        panels.setHandleWidth(1)
        panels.setChildrenCollapsible(False)

        # ─── Sidebar ────────────────────────────────────────────
        sidebar = QFrame()
        sidebar.setObjectName("card")
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(20, 18, 20, 18)
        sidebar_layout.setSpacing(10)
        sidebar.setMinimumWidth(340)
        sidebar.setMaximumWidth(400)

        sidebar_title = QLabel("Configuraci\u00f3n")
        sidebar_title.setProperty("class", "section-title")
        sidebar_layout.addWidget(sidebar_title)

        # Template
        tmpl_lbl_row = QHBoxLayout()
        tmpl_lbl_row.setSpacing(4)
        tmpl_label = QLabel("Plantilla Maestra")
        tmpl_label.setProperty("class", "muted")
        tmpl_lbl_row.addWidget(tmpl_label)
        tmpl_lbl_row.addStretch()
        tmpl_help_btn = self._make_icon_btn("fa5s.question-circle", "btn-ghost", "Ver guía de etiquetas / shortcodes", "blue")
        tmpl_help_btn.clicked.connect(self._show_shortcodes_help)
        tmpl_lbl_row.addWidget(tmpl_help_btn)
        sidebar_layout.addLayout(tmpl_lbl_row)

        tmpl_row = QHBoxLayout()
        tmpl_row.setSpacing(6)
        self.template_entry = QLineEdit()
        self.template_entry.setPlaceholderText("Seleccione archivo .docx...")
        tmpl_row.addWidget(self.template_entry)
        tmpl_btn = self._make_icon_btn("fa5s.search", "btn-icon", "Buscar plantilla")
        tmpl_btn.clicked.connect(self._choose_template)
        tmpl_row.addWidget(tmpl_btn)
        sidebar_layout.addLayout(tmpl_row)

        # Output dir
        out_label = QLabel("Carpeta de Destino")
        out_label.setProperty("class", "muted")
        sidebar_layout.addWidget(out_label)
        out_row = QHBoxLayout()
        out_row.setSpacing(6)
        self.output_entry = QLineEdit()
        self.output_entry.setPlaceholderText("Seleccione carpeta de salida...")
        out_row.addWidget(self.output_entry)
        out_btn = self._make_icon_btn("fa5s.folder-open", "btn-icon", "Seleccionar carpeta")
        out_btn.clicked.connect(self._choose_output_dir)
        out_row.addWidget(out_btn)
        sidebar_layout.addLayout(out_row)

        sidebar_layout.addSpacing(4)

        # Grade + Period
        gp_row = QHBoxLayout()
        gp_row.setSpacing(10)
        g_col = QVBoxLayout()
        g_col.setSpacing(4)
        g_label = QLabel("Grado")
        g_label.setProperty("class", "muted")
        g_col.addWidget(g_label)
        self.grade_combo = QComboBox()
        self.grade_combo.addItems(["Seleccionar..."] + list(GRADES_INFO.keys()))
        self.grade_combo.setCurrentIndex(0)
        self.grade_combo.currentIndexChanged.connect(lambda _: self._update_preview())
        g_col.addWidget(self.grade_combo)
        gp_row.addLayout(g_col)
        p_col = QVBoxLayout()
        p_col.setSpacing(4)
        p_label = QLabel("Periodo")
        p_label.setProperty("class", "muted")
        p_col.addWidget(p_label)
        self.period_combo = QComboBox()
        self.period_combo.addItems(["Seleccionar...", "P1", "P2", "P3", "P4"])
        self.period_combo.setCurrentIndex(0)
        self.period_combo.currentIndexChanged.connect(lambda _: self._update_preview())
        p_col.addWidget(self.period_combo)
        gp_row.addLayout(p_col)
        sidebar_layout.addLayout(gp_row)

        # Divider
        div1 = QFrame()
        div1.setObjectName("divider")
        sidebar_layout.addWidget(div1)

        # Date
        date_label = QLabel("Fecha de Aplicaci\u00f3n")
        date_label.setProperty("class", "section-title")
        sidebar_layout.addWidget(date_label)

        self.date_edit = DateSelector()
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(self._on_date_changed)
        sidebar_layout.addWidget(self.date_edit)

        # Divider
        div2 = QFrame()
        div2.setObjectName("divider")
        sidebar_layout.addWidget(div2)

        # Naming templates
        tpl_lbl_row = QHBoxLayout()
        tpl_lbl_row.setSpacing(4)
        tpl_label = QLabel("Plantillas de Salida")
        tpl_label.setProperty("class", "section-title")
        tpl_lbl_row.addWidget(tpl_label)
        tpl_lbl_row.addStretch()
        tpl_help_btn = self._make_icon_btn("fa5s.question-circle", "btn-ghost", "Ver guía de etiquetas / shortcodes", "blue")
        tpl_help_btn.clicked.connect(self._show_shortcodes_help)
        tpl_lbl_row.addWidget(tpl_help_btn)
        sidebar_layout.addLayout(tpl_lbl_row)

        shortcodes_hint = QLabel("{grade} {period} {session} {year} {level}")
        shortcodes_hint.setProperty("class", "muted")
        sidebar_layout.addWidget(shortcodes_hint)

        fn_label = QLabel("Nombre del archivo")
        fn_label.setProperty("class", "muted")
        sidebar_layout.addWidget(fn_label)
        self.name_template_entry = QLineEdit()
        self.name_template_entry.setText(self.name_template)
        self.name_template_entry.setCursorPosition(0)
        self.name_template_entry.textChanged.connect(self._on_name_template_changed)
        sidebar_layout.addWidget(self.name_template_entry)

        dt_label = QLabel("T\u00edtulo del documento")
        dt_label.setProperty("class", "muted")
        sidebar_layout.addWidget(dt_label)
        self.title_template_entry = QLineEdit()
        self.title_template_entry.setText(self.title_template)
        self.title_template_entry.setCursorPosition(0)
        self.title_template_entry.textChanged.connect(self._on_title_template_changed)
        sidebar_layout.addWidget(self.title_template_entry)

        # Output preview box
        sidebar_layout.addSpacing(4)
        prev_header = QLabel("Vista previa de salida")
        prev_header.setProperty("class", "section-title")
        sidebar_layout.addWidget(prev_header)

        self.preview_box = QLabel()
        self.preview_box.setObjectName("preview-box")
        self.preview_box.setWordWrap(True)
        self.preview_box.setTextFormat(Qt.TextFormat.RichText)
        sidebar_layout.addWidget(self.preview_box)

        sidebar_layout.addStretch()

        # Import/Export/Limpiar
        ie_row = QHBoxLayout()
        ie_row.setSpacing(6)
        imp_btn = QPushButton(" Importar")
        imp_btn.setObjectName("btn-outline")
        imp_btn.clicked.connect(self._import_config)
        self._reg_icon(imp_btn, "fa5s.file-import", "text")
        ie_row.addWidget(imp_btn)
        exp_btn = QPushButton(" Exportar")
        exp_btn.setObjectName("btn-outline")
        exp_btn.clicked.connect(self._export_config)
        self._reg_icon(exp_btn, "fa5s.file-export", "text")
        ie_row.addWidget(exp_btn)
        clr_btn = QPushButton(" Limpiar")
        clr_btn.setObjectName("btn-outline")
        clr_btn.clicked.connect(self._clear_all)
        self._reg_icon(clr_btn, "fa5s.broom", "red")
        ie_row.addWidget(clr_btn)
        sidebar_layout.addLayout(ie_row)

        panels.addWidget(sidebar)

        # ─── Workspace ──────────────────────────────────────────
        workspace = QSplitter(Qt.Orientation.Horizontal)
        workspace.setHandleWidth(1)
        workspace.setChildrenCollapsible(False)

        # Column 1: Sessions
        col1 = QFrame()
        col1.setObjectName("card")
        col1_layout = QVBoxLayout(col1)
        col1_layout.setContentsMargins(16, 14, 16, 14)
        col1_layout.setSpacing(8)

        c1_header = QWidget()
        c1_header_layout = QHBoxLayout(c1_header)
        c1_header_layout.setContentsMargins(0, 0, 0, 0)
        c1_title = QLabel("Sesiones")
        c1_title.setProperty("class", "section-title")
        c1_header_layout.addWidget(c1_title)
        c1_header_layout.addStretch()
        c1_del = self._make_icon_btn("fa5s.trash-alt", "btn-del", "Eliminar sesiones", "red")
        c1_del.clicked.connect(self._remove_sessions)
        c1_header_layout.addWidget(c1_del)
        c1_add = self._make_icon_btn("fa5s.plus", "btn-add", "A\u00f1adir sesi\u00f3n", "green")
        c1_add.clicked.connect(self._add_session)
        c1_header_layout.addWidget(c1_add)
        col1_layout.addWidget(c1_header)

        self.sessions_lb = ReorderableList()
        self.sessions_lb.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.sessions_lb.currentRowChanged.connect(self._on_select_session)
        self.sessions_lb.reordered.connect(self._on_sessions_reordered)
        col1_layout.addWidget(self.sessions_lb)

        workspace.addWidget(col1)

        # Column 2: Subsessions
        col2 = QFrame()
        col2.setObjectName("card")
        col2_layout = QVBoxLayout(col2)
        col2_layout.setContentsMargins(16, 14, 16, 14)
        col2_layout.setSpacing(8)

        c2_header = QWidget()
        c2_header_layout = QHBoxLayout(c2_header)
        c2_header_layout.setContentsMargins(0, 0, 0, 0)
        c2_title = QLabel("Subsesiones")
        c2_title.setProperty("class", "section-title")
        c2_header_layout.addWidget(c2_title)
        c2_header_layout.addStretch()
        c2_del = self._make_icon_btn("fa5s.trash-alt", "btn-del", "Eliminar subsesiones", "red")
        c2_del.clicked.connect(self._remove_subs)
        c2_header_layout.addWidget(c2_del)
        c2_add = self._make_icon_btn("fa5s.plus", "btn-add", "A\u00f1adir subsesi\u00f3n", "green")
        c2_add.clicked.connect(self._add_sub)
        c2_header_layout.addWidget(c2_add)
        col2_layout.addWidget(c2_header)

        self.subs_lb = ReorderableList()
        self.subs_lb.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.subs_lb.setMouseTracking(True)
        self.subs_lb.currentRowChanged.connect(self._on_select_sub)
        self.subs_lb.itemEntered.connect(self._on_sub_hover)
        self.subs_lb.reordered.connect(self._on_subs_reordered)
        col2_layout.addWidget(self.subs_lb)

        workspace.addWidget(col2)

        # Column 3: Files
        col3 = QFrame()
        col3.setObjectName("card")
        col3_layout = QVBoxLayout(col3)
        col3_layout.setContentsMargins(16, 14, 16, 14)
        col3_layout.setSpacing(8)

        c3_header = QWidget()
        c3_header_layout = QHBoxLayout(c3_header)
        c3_header_layout.setContentsMargins(0, 0, 0, 0)
        c3_title = QLabel("Archivos")
        c3_title.setProperty("class", "section-title")
        c3_header_layout.addWidget(c3_title)
        c3_header_layout.addStretch()

        c3_open = self._make_icon_btn("fa5s.folder-open", "btn-icon", "Abrir ubicaci\u00f3n del archivo en el explorador", "text")
        c3_open.clicked.connect(self._open_file_location)
        c3_header_layout.addWidget(c3_open)

        c3_up = self._make_icon_btn("fa5s.chevron-up", "btn-icon", "Subir", "muted")
        c3_up.clicked.connect(lambda: self._move_file(-1))
        c3_header_layout.addWidget(c3_up)

        c3_down = self._make_icon_btn("fa5s.chevron-down", "btn-icon", "Bajar", "muted")
        c3_down.clicked.connect(lambda: self._move_file(1))
        c3_header_layout.addWidget(c3_down)

        c3_del = self._make_icon_btn("fa5s.trash-alt", "btn-del", "Eliminar archivos", "red")
        c3_del.clicked.connect(self._remove_files)
        c3_header_layout.addWidget(c3_del)

        c3_add = self._make_icon_btn("fa5s.plus", "btn-add", "A\u00f1adir archivos", "green")
        c3_add.clicked.connect(self._add_files)
        c3_header_layout.addWidget(c3_add)

        col3_layout.addWidget(c3_header)

        self.files_lb = ReorderableList()
        self.files_lb.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.files_lb.reordered.connect(self._on_files_reordered)
        self.files_lb.filesDropped.connect(self._add_files_from_paths)
        self.files_lb.itemDoubleClicked.connect(self._open_file_location)
        col3_layout.addWidget(self.files_lb)

        workspace.addWidget(col3)
        workspace.setSizes([300, 300, 300])

        panels.addWidget(workspace)
        panels.setSizes([360, 700])

        main_layout.addWidget(panels, 1)

        # ─── Bottom panel ───────────────────────────────────────
        bottom = QFrame()
        bottom.setObjectName("card")
        bottom.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        bottom_layout = QVBoxLayout(bottom)
        bottom_layout.setContentsMargins(18, 14, 18, 14)

        bottom_inner = QWidget()
        bottom_inner_layout = QHBoxLayout(bottom_inner)
        bottom_inner_layout.setContentsMargins(0, 0, 0, 0)
        bottom_inner_layout.setSpacing(20)

        # Log
        log_side = QWidget()
        log_layout = QVBoxLayout(log_side)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.setSpacing(4)
        log_label = QLabel("Registro (Logs)")
        log_label.setProperty("class", "muted")
        log_layout.addWidget(log_label)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(100)
        self.log.setMaximumHeight(140)
        log_layout.addWidget(self.log)
        bottom_inner_layout.addWidget(log_side, stretch=1)

        # Action panel
        action_side = QWidget()
        action_side.setFixedWidth(250)
        action_layout = QVBoxLayout(action_side)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(4)

        prog_header = QWidget()
        prog_header_layout = QHBoxLayout(prog_header)
        prog_header_layout.setContentsMargins(0, 0, 0, 0)
        self.status = QLabel("Listo")
        self.status.setObjectName("status-label")
        prog_header_layout.addWidget(self.status)
        prog_header_layout.addStretch()
        self.progress_lbl = QLabel("0%")
        self.progress_lbl.setProperty("class", "muted")
        prog_header_layout.addWidget(self.progress_lbl)
        action_layout.addWidget(prog_header)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setValue(0)
        action_layout.addWidget(self.progress_bar)

        action_layout.addSpacing(8)

        btn_frame = QWidget()
        btn_layout = QVBoxLayout(btn_frame)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(6)

        self.generate_btn = QPushButton("  GENERAR TODO")
        self.generate_btn.setObjectName("btn-green")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.clicked.connect(self._generate_all)
        self._reg_icon(self.generate_btn, "fa5s.play", "#ffffff")
        btn_layout.addWidget(self.generate_btn)

        self.stop_btn = QPushButton("  DETENER")
        self.stop_btn.setObjectName("btn-red")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.clicked.connect(self._stop_generation)
        self.stop_btn.setIcon(_qta().icon("fa5s.stop", color="#ffffff", color_disabled="#ffffff"))
        self.stop_btn.hide()
        btn_layout.addWidget(self.stop_btn)

        action_layout.addWidget(btn_frame)
        bottom_inner_layout.addWidget(action_side)

        bottom_layout.addWidget(bottom_inner)
        main_layout.addWidget(bottom, 0)

    # ─── helpers ───────────────────────────────────────────────

    def _set_status_color(self, color):
        self.status.setStyleSheet(f"#status-label {{ font-size: 12px; font-weight: 700; color: {color}; }}")

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar plantilla maestra",
            "", "Word (*.docx);;Todos (*.*)")
        if path:
            self.template_path = path
            self.template_entry.setText(path)
            self._log(f"Plantilla cargada: {os.path.basename(path)}")

    def _choose_output_dir(self):
        path = QFileDialog.getExistingDirectory(
            self, "Carpeta donde guardar los documentos generados")
        if path:
            self.output_dir = path
            self.output_entry.setText(path)
            self._log(f"Carpeta de salida: {path}")

    def _log(self, text):
        self.log.append(f"[{threading.current_thread().name}] {text}")

    def _config_dict(self):
        return {
            "template": self.template_path,
            "output_dir": self.output_dir,
            "grade": self.grade_combo.currentText(),
            "period": self.period_combo.currentText(),
            "year": self.sessions[0]["year"] if self.sessions else "2026",
            "sessions": [
                {
                    "name": s["name"],
                    "day": s["day"],
                    "month": s["month"],
                    "year": s["year"],
                    "subsessions": [
                        {"name": sub["name"], "files": sub["files"]}
                        for sub in s["subsessions"]
                    ],
                }
                for s in self.sessions
            ],
        }

    def _apply_config(self, cfg):
        tmpl = cfg.get("template") or cfg.get("template_filename")
        if tmpl and os.path.exists(tmpl):
            self.template_path = tmpl
            self.template_entry.setText(tmpl)
        elif tmpl:
            fname = os.path.basename(tmpl)
            local = os.path.join(os.getcwd(), fname)
            if os.path.exists(local):
                self.template_path = os.path.abspath(local)
                self.template_entry.setText(self.template_path)
            else:
                if self._ask_yes_no(
                    "Plantilla no encontrada",
                    f"No se encontr\u00f3 la plantilla:\n{tmpl}\n\n\u00bfBuscarla manualmente?"
                ):
                    path, _ = QFileDialog.getOpenFileName(
                        self, "Seleccionar plantilla",
                        "", "Word (*.docx);;Todos (*.*)")
                    if path:
                        self.template_path = path
                        self.template_entry.setText(path)

        out_dir = cfg.get("output_dir")
        if out_dir and os.path.exists(out_dir):
            self.output_dir = out_dir
            self.output_entry.setText(out_dir)
        elif out_dir:
            local_out = os.path.join(os.getcwd(), os.path.basename(out_dir))
            if os.path.exists(local_out):
                self.output_dir = local_out
                self.output_entry.setText(local_out)
            else:
                if self._ask_yes_no(
                    "Carpeta de salida no encontrada",
                    f"No se encontr\u00f3 la carpeta:\n{out_dir}\n\n\u00bfSeleccionar otra?"
                ):
                    path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida")
                    if path:
                        self.output_dir = path
                        self.output_entry.setText(path)

        for eng, esp in [("grade", "grado"), ("period", "periodo")]:
            v = cfg.get(eng) or cfg.get(esp)
            if v:
                combo = self.grade_combo if eng == "grade" else self.period_combo
                idx = combo.findText(v)
                if idx >= 0:
                    combo.setCurrentIndex(idx)

        self.name_template = cfg.get("name_template", "E.S.A._{grade}_{period}_{session}_{year}")
        self.name_template_entry.setText(self.name_template)
        self.title_template = cfg.get("title_template", "Evaluaciones de Suficiencia Académica - {grade} - {period} - {session} - {year}")
        self.title_template_entry.setText(self.title_template)

        raw = cfg.get("sessions")
        if not raw:
            raw = [
                {
                    "name": "Sesi\u00f3n 1",
                    "day": cfg.get("day") or cfg.get("dia", "15"),
                    "month": cfg.get("month") or cfg.get("mes", "SEP"),
                    "year": cfg.get("year") or cfg.get("anio", "2026"),
                    "subsessions": [
                        {
                            "name": "Subsesi\u00f3n 1.1",
                            "files": cfg.get("files", []),
                        }
                    ],
                }
            ]
        for s in raw:
            sn = s.get("name", "Sesi\u00f3n 1")
            subs = s.get("subsessions", [])
            if not subs:
                subs = [{"name": sn + ".1", "files": s.get("files", [])}]
            for sub in subs:
                fn = sub.get("ordered_filenames") or sub.get("files", [])
                sub["files"] = fn if isinstance(fn, list) else []
                sub["name"] = sub.get("name", sn + ".1")
            s["subsessions"] = subs

        self.sessions = [
            {
                "name": s.get("name", f"Sesi\u00f3n {i+1}"),
                "day": str(s.get("day", "15")),
                "month": s.get("month", "SEP"),
                "year": str(s.get("year", "2026")),
                "subsessions": s["subsessions"],
            }
            for i, s in enumerate(raw)
        ]

        missing = []
        existing_dirs = set()
        for s in self.sessions:
            for sub in s["subsessions"]:
                for f in sub["files"]:
                    if not os.path.exists(f):
                        missing.append(f)
                    else:
                        existing_dirs.add(os.path.dirname(f))
        if missing:
            old_root = os.path.commonpath(missing) if len(missing) > 1 else os.path.dirname(missing[0])
            if existing_dirs:
                common_roots = set()
                for d in existing_dirs:
                    p = os.path.commonpath([d, old_root])
                    common_roots.add(p)
                if common_roots:
                    old_root = max(common_roots, key=len) if len(common_roots) > 1 else old_root
            if self._ask_yes_no(
                "Archivos no encontrados",
                f"Se encontraron {len(missing)} archivo(s) sin localizar.\n"
                f"Ejemplo: {os.path.basename(missing[0])}\n\n"
                "\u00bfSeleccionar la carpeta ra\u00edz actual para reubicarlos?"
            ):
                new_root = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta ra\u00edz")
                if new_root:
                    for s in self.sessions:
                        for sub in s["subsessions"]:
                            sub["files"] = [
                                f.replace(old_root, new_root) if not os.path.exists(f) and old_root in f else f
                                for f in sub["files"]
                            ]
            total_missing = sum(
                1 for s in self.sessions for sub in s["subsessions"] for f in sub["files"]
                if not os.path.exists(f)
            )
            if total_missing:
                detail = "\n".join(
                    f"  \u2022 {os.path.basename(f)}"
                    for s in self.sessions for sub in s["subsessions"] for f in sub["files"]
                    if not os.path.exists(f)
                )[:500]
                self._show_warning(
                    "Atenci\u00f3n",
                    f"A\u00fan quedan {total_missing} archivo(s) sin ubicar.\n\n{detail}\n\n"
                    "Se seleccionar\u00e1 la primera subsesi\u00f3n con archivos faltantes.\n"
                    "Puedes agregarlos manualmente."
                )
                for si, s in enumerate(self.sessions):
                    for subi, sub in enumerate(s["subsessions"]):
                        if any(not os.path.exists(f) for f in sub["files"]):
                            self._selected_session = si
                            self._selected_sub = subi
                            self._refresh_sessions()
                            break
                    else:
                        continue
                    break
            elif missing:
                self._show_info(
                    "Archivos reubicados",
                    f"Todos los {len(missing)} archivo(s) se reubicaron correctamente."
                )

        self._selected_session = 0
        self._selected_sub = 0
        self._refresh_sessions()

    # ─── export / import ───────────────────────────────────────

    def _export_config(self):
        g_name = self.grade_combo.currentText()
        p_name = self.period_combo.currentText()
        g_str = "ESQUEMA" if g_name.startswith("Seleccionar") else g_name
        p_str = "BASE" if p_name.startswith("Seleccionar") else p_name
        default_fn = f"CONFIG_{g_str}_{p_str}_{self.sessions[0]['year'] if self.sessions else '2026'}.json"
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar configuraci\u00f3n",
            default_fn,
            "JSON (*.json)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self._config_dict(), f, indent=2, ensure_ascii=False)
        self._log(f"Configuraci\u00f3n exportada: {path}")

    def _clear_all(self):
        if not self._ask_yes_no("Limpiar configuraci\u00f3n", "\u00bfDeseas limpiar toda la configuraci\u00f3n actual y dejar la aplicaci\u00f3n por defecto limpia?"):
            return
        self._save_state_for_undo()
        self.template_path = ""
        self.output_dir = ""
        self.grade = ""
        self.period = ""
        self.name_template = "E.S.A._{grade}_{period}_{session}_{year}"
        self.title_template = "Evaluaciones de Suficiencia Acad\u00e9mica - {grade} - {period} - {session} - {year}"

        if hasattr(self, "template_entry"):
            self.template_entry.setText("")
        if hasattr(self, "output_entry"):
            self.output_entry.setText("")
        if hasattr(self, "name_template_entry"):
            self.name_template_entry.setText(self.name_template)
        if hasattr(self, "title_template_entry"):
            self.title_template_entry.setText(self.title_template)
        if hasattr(self, "grade_combo"):
            self.grade_combo.setCurrentIndex(0)
        if hasattr(self, "period_combo"):
            self.period_combo.setCurrentIndex(0)

        qd = QDate.currentDate()
        self.sessions = [{
            "name": "Sesión 1",
            "day": str(qd.day()),
            "month": MONTH_REV[qd.month()],
            "year": str(qd.year()),
            "subsessions": [{
                "name": "Subsesión 1.1",
                "files": []
            }],
        }]
        self._selected_session = 0
        self._selected_sub = 0
        self._refresh_sessions()
        self._log("🧹 Configuración limpiada por defecto.")

    def _import_config(self):
        if self.sessions and any(sub.get("files") for s in self.sessions for sub in s.get("subsessions", [])):
            if self._ask_yes_no("Importar configuración", "¿Deseas limpiar la configuración actual antes de importar la nueva?"):
                self._save_state_for_undo()

        path, _ = QFileDialog.getOpenFileName(
            self, "Importar configuración",
            "", "JSON (*.json)")
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        self._apply_config(cfg)
        total = sum(len(sub["files"]) for s in self.sessions for sub in s["subsessions"])
        self._log(f"Configuración importada: {path} ({total} archivos)")

    # ─── generate all & stop ───────────────────────────────────

    def _stop_generation(self):
        if self.processing and not self._stop_requested:
            self._stop_requested = True
            if hasattr(self, "_worker") and self._worker:
                self._worker.stop_requested = True
            self.stop_btn.setIcon(_qta().icon("fa5s.stop", color="#fca5a5", color_disabled="#fca5a5"))
            self.stop_btn.setText("  Cancelando...")
            self.stop_btn.setStyleSheet("""
                QPushButton, QPushButton:disabled {
                    background-color: #7f1d1d;
                    color: #fca5a5;
                    border: 1px solid #991b1b;
                    border-radius: 6px;
                    font-weight: 700;
                    font-size: 13px;
                    min-height: 40px;
                }
            """)
            self.stop_btn.setEnabled(False)
            self.status.setText("Cancelando...")
            self._set_status_color("orange")
            self._log("⚠️ Cancelación solicitada por el usuario...")

    def _generate_all(self):
        if self.processing:
            return
        g_val = self.grade_combo.currentText()
        p_val = self.period_combo.currentText()
        if not g_val or g_val.startswith("Seleccionar"):
            self._show_warning("Grado no seleccionado", "Por favor selecciona un Grado antes de generar.")
            return
        if not p_val or p_val.startswith("Seleccionar"):
            self._show_warning("Periodo no seleccionado", "Por favor selecciona un Periodo antes de generar.")
            return
        if not self.sessions:
            self._show_warning("Sin datos", "No hay sesiones configuradas.")
            return
        if not self.template_path or not os.path.isfile(self.template_path):
            self._show_warning("Sin plantilla", "Selecciona la plantilla maestra.")
            return
        total_files = sum(
            len(sub["files"]) for s in self.sessions for sub in s["subsessions"]
        )
        if total_files == 0:
            self._show_warning("Sin archivos", "Agrega archivos a las subsesiones.")
            return

        empty_subs = [
            f"{s['name']} ➔ {sub['name']}"
            for s in self.sessions for sub in s["subsessions"]
            if not sub["files"]
        ]
        if empty_subs:
            count = len(empty_subs)
            preview_subs = "\n".join(f"  • {name}" for name in empty_subs[:6])
            if count > 6:
                rem = count - 6
                rem_str = "1 subsesión más" if rem == 1 else f"{rem} subsesiones más"
                preview_subs += f"\n  ...y {rem_str}"

            if count == 1:
                dlg_title = "Subsesión sin archivos"
                msg_header = "Se detectó 1 subsesión sin archivos asignados:"
            else:
                dlg_title = "Subsesiones sin archivos"
                msg_header = f"Se detectaron {count} subsesiones sin archivos asignados:"
            
            msg = (
                f"{msg_header}\n\n"
                f"{preview_subs}\n\n"
                "¿Deseas omitir este aviso y continuar con la generación de las demás subsesiones?"
            )
            if not self._ask_yes_no(dlg_title, msg):
                return

        self.processing = True
        self._stop_requested = False

        self.generate_btn.hide()
        self.stop_btn.setIcon(_qta().icon("fa5s.stop", color="#ffffff", color_disabled="#ffffff"))
        self.stop_btn.setText("  DETENER")
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc2626;
                color: #ffffff;
                border: 1px solid #b91c1c;
                border-radius: 6px;
                font-weight: 700;
                font-size: 13px;
                min-height: 40px;
            }
            QPushButton:hover {
                background-color: #ef4444;
            }
        """)
        self.stop_btn.show()
        self.stop_btn.setEnabled(True)

        self.progress_bar.setValue(0)
        self.progress_lbl.setText("0%")
        self.status.setText("Procesando...")
        self._set_status_color("orange")
        self._log("\U0001f680 Iniciando generaci\u00f3n masiva...")

        self._worker = GenerateWorker(
            self.template_path,
            self.sessions,
            self.grade_combo.currentText(),
            self.period_combo.currentText(),
            self.name_template,
            self.title_template,
            self.output_dir,
        )
        self._worker_thread = QThread()
        self._worker.moveToThread(self._worker_thread)
        self._worker_thread.started.connect(self._worker.run)
        self._worker.progress_updated.connect(self._on_progress)
        self._worker.log_message.connect(self._log)
        self._worker.generation_finished.connect(self._on_generation_finished)
        self._worker.generation_finished.connect(self._worker_thread.quit)
        self._worker_thread.finished.connect(self._cleanup_thread)
        self._worker_thread.start()

    def _on_progress(self, ratio, pct_str):
        self.progress_bar.setValue(int(ratio * 100))
        self.progress_lbl.setText(pct_str)

    def _on_generation_finished(self, data):
        self.processing = False

        ok = sum(1 for k, _ in data if k == "ok")
        err = sum(1 for k, _ in data if k == "error")
        was_stopped = any(k == "stopped" for k, _ in data)

        for k, msg in data:
            if k == "ok":
                self._log(f"\u2705 {msg}")
            elif k == "error":
                self._log(f"\u274c {msg}")
            elif k == "stopped":
                self._log(f"\U0001f6d1 {msg}")

        if was_stopped:
            self.status.setText("Detenido")
            self._set_status_color("orange")
            self._show_warning("Proceso Detenido", "La generaci\u00f3n fue cancelada por el usuario.")
        elif err == 0:
            self.progress_bar.setValue(100)
            self.progress_lbl.setText("100%")
            self.status.setText("Completado")
            self._set_status_color("#10b981")
            doc_word = "documento" if ok == 1 else "documentos"
            gen_word = "generado" if ok == 1 else "generados"
            self._show_info("Completado", f"{ok} {doc_word} {gen_word} correctamente.")
        else:
            self.status.setText(f"Errores ({err})")
            self._set_status_color("#ef4444")
            sub_word = "subsesi\u00f3n" if err == 1 else "subsesiones"
            fall_word = "fall\u00f3" if err == 1 else "fallaron"
            self._show_error("Errores", f"{err} {sub_word} {fall_word}. Revisa el log.")

    def _cleanup_thread(self):
        self.stop_btn.hide()
        self.generate_btn.show()
        self._worker = None
        self._worker_thread = None

    def closeEvent(self, event):
        if self._worker_thread and self._worker_thread.isRunning():
            if self._worker:
                self._worker.stop_requested = True
            self._worker_thread.quit()
            self._worker_thread.wait(3000)
        event.accept()


def get_effective_theme_is_dark():
    try:
        settings = QSettings("GESA", "AcademicManager")
        t = str(settings.value("theme", "system"))
        if t == "dark":
            return True
        elif t == "light":
            return False
        else:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return val == 0
    except Exception:
        return True


def global_excepthook(exctype, value, traceback_obj):
    import traceback
    tb_str = "".join(traceback.format_exception(exctype, value, traceback_obj))
    sys.stderr.write(tb_str)
    try:
        full_msg = f"Ocurrió un error no controlado:\n\n{value}\n\nDetalles técnicos:\n{tb_str}"
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Error de Ejecución — GESA")
        msg_box.setText(full_msg)
        msg_box.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        try:
            ico = _qta().icon("fa5s.times-circle", color="#cf222e")
            msg_box.setIconPixmap(ico.pixmap(48, 48))
        except Exception:
            msg_box.setIcon(QMessageBox.Icon.Critical)
        btn_copy = msg_box.addButton("Copiar Error", QMessageBox.ButtonRole.ActionRole)
        btn_ok = msg_box.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        msg_box.exec()
        if msg_box.clickedButton() == btn_copy:
            QApplication.clipboard().setText(full_msg)
    except Exception:
        pass

sys.excepthook = global_excepthook


class GesaSplashScreen(QSplashScreen):
    def __init__(self, icon_path=None):
        is_dark = get_effective_theme_is_dark()
        
        bg_color = QColor("#161b22") if is_dark else QColor("#ffffff")
        border_color = QColor("#30363d") if is_dark else QColor("#d0d7de")
        title_color = QColor("#58a6ff") if is_dark else QColor("#0969da")
        subtitle_color = QColor("#8b949e") if is_dark else QColor("#57606a")
        loading_color = QColor("#c9d1d9") if is_dark else QColor("#24292f")
        bar_color = QColor("#238636") if is_dark else QColor("#2da44e")
        badge_bg = QColor("#1f6beb") if is_dark else QColor("#0969da")

        pix = QPixmap(520, 260)
        pix.fill(QColor(0, 0, 0, 0))
        super().__init__(pix, Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Draw rounded background card
        painter.setPen(border_color)
        painter.setBrush(bg_color)
        painter.drawRoundedRect(2, 2, 516, 256, 16, 16)

        # Draw Icon (prefer SVG for crisp vector rendering)
        svg_file = get_resource_path("book-check.svg")
        if os.path.exists(svg_file):
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(badge_bg)
            painter.drawRoundedRect(30, 36, 60, 60, 14, 14)
            
            renderer = QSvgRenderer(svg_file)
            renderer.render(painter, QRectF(42, 48, 36, 36))
        elif icon_path and os.path.exists(icon_path):
            ico_pix = QPixmap(icon_path).scaled(64, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            painter.drawPixmap(30, 36, ico_pix)

        # Title GESA
        painter.setPen(title_color)
        font = painter.font()
        font.setFamily("Century Gothic")
        font.setPointSize(24)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(106, 70, "GESA")

        # Subtitle
        painter.setPen(subtitle_color)
        font.setPointSize(11)
        font.setBold(False)
        painter.setFont(font)
        painter.drawText(106, 95, "Gestor de Evaluaciones de Suficiencia Académica")

        # Loading text
        painter.setPen(loading_color)
        font.setPointSize(10)
        painter.setFont(font)
        painter.drawText(30, 195, "⚡ Cargando interfaz y componentes...")

        # Accent Bar
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bar_color)
        painter.drawRoundedRect(30, 215, 460, 4, 2, 2)

        painter.end()
        self.setPixmap(pix)


if __name__ == "__main__":
    try:
        myappid = "danielrozocom.gesa.academic.v1"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName("GESA")
    app.setOrganizationName("GESA")
    app.setDesktopFileName("GESA.lnk")

    icon_file = get_resource_path("app_icon.ico")
    if not os.path.exists(icon_file):
        icon_file = get_resource_path("app_icon.png")
    if os.path.exists(icon_file):
        app.setWindowIcon(QIcon(icon_file))

    app.setStyle("Fusion")

    # ── Instant Splash Screen & Window Load ───────────────────
    splash = None
    try:
        splash = GesaSplashScreen(icon_file)
        splash.show()
        app.processEvents()

        window = DesktopApp()
        window.show()
        splash.finish(window)
    except Exception as err:
        import traceback
        tb = traceback.format_exc()
        try:
            if splash:
                splash.hide()
        except Exception:
            pass
        full_err_text = f"Ocurrió un error inesperado al iniciar la aplicación:\n\n{err}\n\nDetalles técnicos:\n{tb}"
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Error de inicio en GESA")
        msg_box.setText(full_err_text)
        msg_box.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextSelectableByKeyboard)
        msg_box.setIcon(QMessageBox.Icon.Critical)
        btn_copy = msg_box.addButton("Copiar Error", QMessageBox.ButtonRole.ActionRole)
        btn_ok = msg_box.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        msg_box.exec()
        if msg_box.clickedButton() == btn_copy:
            QApplication.clipboard().setText(full_err_text)
        sys.exit(1)

    sys.exit(app.exec())
