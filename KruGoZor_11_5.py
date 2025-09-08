#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
КругоЗор v11.4 — минималистичная «пузырь-камера» (круг/квадрат) поверх всех окон + независимая виртуальная камера
+ интегрированный калькулятор (сворачивание/разворачивание по NumLock, без убийства процесса) + тёмное меню + работа из трея.

Что нового относительно v11.3:
- «Выключить камеру» теперь дополнительно прячет окно (через синхронизацию с чекбоксом трея).
- При показе окна из трея больше не поднимаем Qt.Window — кнопка в панели задач не появляется (окно остаётся Qt.Tool).
- В виртуальной камере режим «По ширине окна» и «Letterbox» используют ровно квадратный ROI из окна (без растяжений).
"""

import argparse
import ctypes
import json
import os
import sys
import subprocess
import threading
import time
import logging
import logging.handlers
import traceback
from dataclasses import dataclass, asdict
from math import ceil

import psutil
import keyboard
import cv2
import numpy as np
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QImage, QPainter, QRegion, QIcon, QPixmap, QCloseEvent
from PyQt5.QtWidgets import (
    QWidget, QApplication, QMenu, QAction, QDialog, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QSlider, QSpinBox, QCheckBox,
    QGroupBox, QSystemTrayIcon, QWidgetAction, QMessageBox, QShortcut
)

APP_NAME = "КругоЗор"
APP_VERSION = "11.4 - 25.08.2025"
CONTACT = "Андрей Кудлай @akudlay_ru"

# --- Виртуальная камера ----
try:
    import pyvirtualcam
    HAVE_PYVIRTUALCAM = True
except Exception:
    HAVE_PYVIRTUALCAM = False



# ===================== КАЛЬКУЛЯТОР (без убийства процесса) =====================
# Управляем штатным калькулятором Windows (UWP) через WinAPI:
# - Находим верхнее окно класса "ApplicationFrameWindow" с заголовком "Calculator"/"Калькулятор"
# - Сворачиваем/разворачиваем через ShowWindow(SW_MINIMIZE/SW_RESTORE)
# - Если окна нет — запускаем calc.exe
CALC_PROC = "CalculatorApp.exe"
CALC_CMD = "calc.exe"
CALC_CLASS = "ApplicationFrameWindow"
CALC_TITLES = ("Calculator", "Калькулятор")
calc_settings = {"force_numlock": True}

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
GetWindowTextW = user32.GetWindowTextW
GetWindowTextLengthW = user32.GetWindowTextLengthW
IsWindowVisible = user32.IsWindowVisible
GetClassNameW = user32.GetClassNameW
ShowWindow = user32.ShowWindow
IsIconic = user32.IsIconic
SetForegroundWindow = user32.SetForegroundWindow

SW_MINIMIZE = 6
SW_RESTORE = 9
SW_SHOW = 5
SW_HIDE = 0

def _get_window_text(hwnd) -> str:
    length = GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowTextW(hwnd, buff, length + 1)
    return buff.value

def _get_class_name(hwnd) -> str:
    buff = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buff, 256)
    return buff.value

def find_calc_hwnd():
    """Ищем верхнее окно калькулятора (UWP) по классу и заголовку."""
    target = {"hwnd": None}

    def _enum_proc(hwnd, lparam):
        try:
            if not IsWindowVisible(hwnd):
                return True
            cls = _get_class_name(hwnd)
            if cls != CALC_CLASS:
                return True
            title = _get_window_text(hwnd)
            if not title:
                return True
            # Совпадение по подстроке (на случай локализации)
            if any(t.lower() in title.lower() for t in CALC_TITLES):
                target["hwnd"] = hwnd
                return False  # стоп
        except Exception:
            pass
        return True

    user32.EnumWindows(EnumWindowsProc(_enum_proc), 0)
    return target["hwnd"]

def is_calc_running():
    return any(proc.info["name"] == CALC_PROC for proc in psutil.process_iter(attrs=["name"]))

def show_calc():
    hwnd = find_calc_hwnd()
    if hwnd:
        ShowWindow(hwnd, SW_RESTORE)
        ShowWindow(hwnd, SW_SHOW)
        try:
            SetForegroundWindow(hwnd)
        except Exception:
            pass
        return True
    return False

def minimize_calc():
    hwnd = find_calc_hwnd()
    if hwnd:
        ShowWindow(hwnd, SW_MINIMIZE)
        return True
    return False

def toggle_calc():
    """
    Поведение:
    - нет процесса/окна → запустить calc.exe
    - есть окно и оно не свёрнуто → свернуть
    - есть окно и свёрнуто → развернуть и активировать
    """
    hwnd = find_calc_hwnd()
    if not hwnd:
        # Возможно процесс ещё стартует — просто запускаем
        try:
            subprocess.Popen([CALC_CMD], shell=True)
        except Exception:
            pass
        return

    iconic = bool(IsIconic(hwnd))
    if iconic:
        show_calc()
    else:
        # Окно есть и не свёрнуто — сворачиваем
        minimize_calc()

def force_numlock_on():
    if not calc_settings.get("force_numlock", True):
        return
    VK_NUMLOCK = 0x90
    state = user32.GetKeyState(VK_NUMLOCK)
    if not (state & 1):
        user32.keybd_event(VK_NUMLOCK, 0, 0, 0)
        user32.keybd_event(VK_NUMLOCK, 0, 2, 0)

def numlock_listener():
    # Фоновый слушатель NumLock → сворачивать/разворачивать калькулятор + форсировать NumLock=ON
    while True:
        try:
            keyboard.wait("num lock")
            toggle_calc()
            time.sleep(0.3)
            force_numlock_on()
        except Exception:
            time.sleep(1.0)
# ==============================================================================



# --- Общие утилиты/пути ---
def _is_frozen() -> bool:
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")

def app_exe_dir() -> str:
    """Папка, где лежит exe/скрипт (portable-режим)."""
    if _is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_data_dir() -> str:
    """
    Пользовательская папка данных:
    - Windows: %APPDATA%\\KrugoZor
    - Иначе:   ~/.KrugoZor
    """
    base = None
    if os.name == "nt":
        base = os.environ.get("APPDATA")
        if base:
            base = os.path.join(base, "KrugoZor")
    if not base:
        base = os.path.join(os.path.expanduser("~"), ".KrugoZor")
    try:
        os.makedirs(base, exist_ok=True)
        for sub in ("Logs", "Snapshots"):
            os.makedirs(os.path.join(base, sub), exist_ok=True)
    except Exception:
        pass
    return base

DATA_DIR = get_data_dir()
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")
LOG_DIR = os.path.join(DATA_DIR, "Logs")
LOG_FILE = os.path.join(LOG_DIR, "krugozor.log")

def ensure_first_run_files():
    """Создаём пустой config.json при первом запуске."""
    try:
        if not os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                f.write("{}")
    except Exception:
        pass

def hide_console_window():
    """Скрыть консольное окно Windows, если есть (для сборки EXE без консоли уберите это)."""
    try:
        if os.name == "nt":
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass

# --- Логирование ---
_logger = None

def setup_logger(debug: bool = False):
    global _logger
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("KrugoZor")
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(threadName)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    # Ротация файла 1 МБ x 3
    fh = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    fh.setLevel(logging.DEBUG if debug else logging.INFO)
    logger.addHandler(fh)
    # В консоль только при --debug
    if debug:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        ch.setLevel(logging.DEBUG)
        logger.addHandler(ch)

    logger.info("=== %s %s старт ===", APP_NAME, APP_VERSION)
    logger.info("Python %s | OpenCV %s | Qt %s", sys.version.split()[0], cv2.__version__, QtCore.QT_VERSION_STR)
    logger.info("Data dir: %s", DATA_DIR)
    _logger = logger
    return logger

def log_exc(msg: str, exc: BaseException):
    if _logger:
        _logger.error("%s: %s\n%s", msg, exc, "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))

def log_cap_props(cap: cv2.VideoCapture):
    if not _logger or cap is None:
        return
    try:
        w = cap.get(cv2.CAP_PROP_FRAME_WIDTH)
        h = cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
        fps = cap.get(cv2.CAP_PROP_FPS)
        fmt = cap.get(cv2.CAP_PROP_FORMAT)
        fourcc = int(cap.get(cv2.CAP_PROP_FOURCC)) if hasattr(cv2, "CAP_PROP_FOURCC") else 0
        _logger.info("CAP props: %.0fx%.0f @ %.2f fps | fmt=%s | fourcc=%s", w, h, fps, fmt, fourcc)
    except Exception as e:
        log_exc("Ошибка чтения CAP свойств", e)

# --- Модели/состояние ---
@dataclass
class ChromaConfig:
    enabled: bool = False
    pickA: list = None
    tolA: int = 30
    pickB: list = None
    tolB: int = 30
    use_hsv: bool = False
    h_min: int = 35
    h_max: int = 85
    s_min: int = 40
    s_max: int = 255
    v_min: int = 40
    v_max: int = 255
    feather: int = 15
    ui_opacity: float = 0.85
    persist: bool = True

@dataclass
class CropConfig:
    enabled: bool = False
    rect: list = None  # [x,y,w,h] в координатах исходного кадра

@dataclass
class AppState:
    always_on_top: bool = True
    click_through: bool = False
    window_mirror: bool = False
    vcam_enabled: bool = False
    vcam_mirror: bool = False
    circle_diameter: int = 360
    pos_x: int = 100
    pos_y: int = 100
    window_shape: str = "circle"  # circle | square
    vcam_fit: str = "fill"        # fill | letterbox
    window_opacity: float = 1.0   # прозрачность окна (0.2..1.0)

def clamp(v, lo, hi):
    return max(lo, min(hi, v))

# --- Диалог хромакея ---
class ChromaDialog(QDialog):
    def __init__(self, parent, chroma, get_frame_callable):
        super().__init__(parent)
        self.setWindowTitle("Настройки хромакея")
        # Без прозрачности, но в тёмном стиле, как меню
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setModal(False)
        self.setStyleSheet(self._dark_dialog_stylesheet())

        self.chroma = chroma
        self.get_frame = get_frame_callable

        self.preview_label = QLabel(self)
        self.preview_label.setMinimumSize(360, 202)
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("background:#1e1e1e; color:#eee; border:1px solid #444;")

        # --- Пипетки и допуски
        self.pickA_btn = QPushButton("Пипетка A"); self.pickB_btn = QPushButton("Пипетка B")
        self.tolA_spin = QSpinBox(); self.tolA_spin.setRange(0, 255); self.tolA_spin.setValue(self.chroma.tolA)
        self.tolB_spin = QSpinBox(); self.tolB_spin.setRange(0, 255); self.tolB_spin.setValue(self.chroma.tolB)

        self.rgbA_lbl = QLabel("A: -"); self.rgbA_lbl.setStyleSheet("color:#eee; min-width:120px;")
        self.rgbB_lbl = QLabel("B: -"); self.rgbB_lbl.setStyleSheet("color:#eee; min-width:120px;")

        # Маленькие цветные квадраты для A/B
        self.rgbA_swatch = QLabel(); self.rgbA_swatch.setFixedSize(22, 22); self.rgbA_swatch.setStyleSheet("background:#555; border:1px solid #777;")
        self.rgbB_swatch = QLabel(); self.rgbB_swatch.setFixedSize(22, 22); self.rgbB_swatch.setStyleSheet("background:#555; border:1px solid #777;")

        # --- Режим HSV + значения (слайдер + спинбокс)
        self.use_hsv_chk = QCheckBox("Режим HSV"); self.use_hsv_chk.setChecked(self.chroma.use_hsv)

        self.h_min_slider = QSlider(Qt.Horizontal); self.h_min_slider.setRange(0, 179)
        self.h_max_slider = QSlider(Qt.Horizontal); self.h_max_slider.setRange(0, 179)
        self.s_min_slider = QSlider(Qt.Horizontal); self.s_min_slider.setRange(0, 255)
        self.s_max_slider = QSlider(Qt.Horizontal); self.s_max_slider.setRange(0, 255)
        self.v_min_slider = QSlider(Qt.Horizontal); self.v_min_slider.setRange(0, 255)
        self.v_max_slider = QSlider(Qt.Horizontal); self.v_max_slider.setRange(0, 255)

        self.h_min_spin = QSpinBox(); self.h_min_spin.setRange(0, 179)
        self.h_max_spin = QSpinBox(); self.h_max_spin.setRange(0, 179)
        self.s_min_spin = QSpinBox(); self.s_min_spin.setRange(0, 255)
        self.s_max_spin = QSpinBox(); self.s_max_spin.setRange(0, 255)
        self.v_min_spin = QSpinBox(); self.v_min_spin.setRange(0, 255)
        self.v_max_spin = QSpinBox(); self.v_max_spin.setRange(0, 255)

        # Инициализируем из модели
        self._sync_sliders_from_model()

        # Мягкость и прочее
        self.feather = QSlider(Qt.Horizontal); self.feather.setRange(0, 101); self.feather.setValue(self.chroma.feather)
        self.opacity_slider = QSlider(Qt.Horizontal); self.opacity_slider.setRange(30, 100)
        self.opacity_slider.setValue(int(clamp(self.chroma.ui_opacity, 0.3, 1.0) * 100))
        self.persist_chk = QCheckBox("Сохранять настройки (config.json)"); self.persist_chk.setChecked(self.chroma.persist)

        # Кнопки
        self.apply_btn = QPushButton("Применить")
        self.close_btn = QPushButton("Закрыть")

        # --- Компоновка
        grid = QGridLayout()
        row = 0
        t = QLabel("Превью (клик для пипетки)"); t.setStyleSheet("color:#eee;")
        grid.addWidget(t, row, 0, 1, 4); row += 1
        grid.addWidget(self.preview_label, row, 0, 1, 4); row += 1

        rgb_box = QGroupBox("Пипетки + допуск"); rgb_box.setStyleSheet("QGroupBox{color:#eee;}")
        rgb_layout = QGridLayout()
        # A
        rgb_layout.addWidget(self.pickA_btn, 0, 0)
        rgb_layout.addWidget(QLabel("Допуск A"), 0, 1)
        rgb_layout.addWidget(self.tolA_spin, 0, 2)
        rgb_layout.addWidget(self.rgbA_swatch, 0, 3)
        rgb_layout.addWidget(self.rgbA_lbl, 0, 4)
        # B
        rgb_layout.addWidget(self.pickB_btn, 1, 0)
        rgb_layout.addWidget(QLabel("Допуск B"), 1, 1)
        rgb_layout.addWidget(self.tolB_spin, 1, 2)
        rgb_layout.addWidget(self.rgbB_swatch, 1, 3)
        rgb_layout.addWidget(self.rgbB_lbl, 1, 4)
        rgb_box.setLayout(rgb_layout)
        grid.addWidget(rgb_box, row, 0, 1, 4); row += 1

        hsv_box = QGroupBox("HSV диапазоны"); hsv_box.setStyleSheet("QGroupBox{color:#eee;}")
        hsv_layout = QGridLayout()
        self.use_hsv_chk.setStyleSheet("color:#eee;")
        hsv_layout.addWidget(self.use_hsv_chk, 0, 0, 1, 4)

        def add_row(r, label_text, slider, spin):
            lbl = QLabel(label_text); lbl.setStyleSheet("color:#eee; min-width:52px;")
            hsv_layout.addWidget(lbl, r, 0)
            hsv_layout.addWidget(slider, r, 1)
            spin.setFixedWidth(60)
            hsv_layout.addWidget(spin, r, 2)

        add_row(1, "H min", self.h_min_slider, self.h_min_spin)
        add_row(2, "H max", self.h_max_slider, self.h_max_spin)
        add_row(3, "S min", self.s_min_slider, self.s_min_spin)
        add_row(4, "S max", self.s_max_slider, self.s_max_spin)
        add_row(5, "V min", self.v_min_slider, self.v_min_spin)
        add_row(6, "V max", self.v_max_slider, self.v_max_spin)
        hsv_box.setLayout(hsv_layout)
        grid.addWidget(hsv_box, row, 0, 1, 4); row += 1

        # feather + opacity of UI + persist
        grid.addWidget(QLabel("Мягкость (feather)"), row, 0)
        grid.addWidget(self.feather, row, 1, 1, 3); row += 1
        grid.addWidget(QLabel("Прозрачность меню/настроек, %"), row, 0)
        grid.addWidget(self.opacity_slider, row, 1, 1, 3); row += 1
        grid.addWidget(self.persist_chk, row, 0, 1, 4); row += 1

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.apply_btn); btn_row.addWidget(self.close_btn)
        grid.addLayout(btn_row, row, 0, 1, 4); row += 1

        self.setLayout(grid)

        # Таймер превью
        self._timer = QTimer(self); self._timer.timeout.connect(self.update_preview); self._timer.start(1000 // 15)

        self.from_model()

        # Сигналы
        self.apply_btn.clicked.connect(self.on_apply)
        self.close_btn.clicked.connect(self.close)
        self.pickA_btn.clicked.connect(lambda: self.start_pick('A'))
        self.pickB_btn.clicked.connect(lambda: self.start_pick('B'))
        self.preview_label.mousePressEvent = self.on_preview_click

        # Синхронизация слайдер ↔ спинбокс
        pairs = [
            (self.h_min_slider, self.h_min_spin, "h_min"),
            (self.h_max_slider, self.h_max_spin, "h_max"),
            (self.s_min_slider, self.s_min_spin, "s_min"),
            (self.s_max_slider, self.s_max_spin, "s_max"),
            (self.v_min_slider, self.v_min_spin, "v_min"),
            (self.v_max_slider, self.v_max_spin, "v_max"),
        ]
        for slider, spin, attr in pairs:
            slider.valueChanged.connect(lambda v, s=spin: s.setValue(v))
            spin.valueChanged.connect(lambda v, a=attr: setattr(self.chroma, a, int(v)))

        self.tolA_spin.valueChanged.connect(lambda v: setattr(self.chroma, "tolA", int(v)))
        self.tolB_spin.valueChanged.connect(lambda v: setattr(self.chroma, "tolB", int(v)))
        self.use_hsv_chk.toggled.connect(lambda v: setattr(self.chroma, "use_hsv", bool(v)))
        self.feather.valueChanged.connect(lambda v: setattr(self.chroma, "feather", int(v)))
        self.opacity_slider.valueChanged.connect(lambda v: setattr(self.chroma, "ui_opacity", v / 100.0))
        self.persist_chk.toggled.connect(lambda v: setattr(self.chroma, "persist", bool(v)))

        # Позиции кликов пипеток в координатах исходного кадра
        self.pickA_pos = None
        self.pickB_pos = None

    def _dark_dialog_stylesheet(self) -> str:
        # Непрозрачный стиль, гармонирующий с меню
        return """
        QDialog { background: #1e1e1e; color: #eee; }
        QLabel { color: #eee; }
        QGroupBox { color: #eee; border: 1px solid #444; margin-top: 8px; }
        QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
        QSlider::groove:horizontal { height: 6px; background: #333; border: 1px solid #444; }
        QSlider::handle:horizontal { width: 12px; background: #777; border: 1px solid #999; margin: -6px 0; }
        QSpinBox { background: #262626; border: 1px solid #555; color: #eee; }
        QCheckBox { color: #eee; }
        QPushButton { background: #2a2a2a; color: #eee; border: 1px solid #555; padding: 4px 10px; }
        QPushButton:hover { background: #333; }
        """

    def _sync_sliders_from_model(self):
        c = self.chroma
        self.h_min_slider.setValue(c.h_min); self.h_min_spin.setValue(c.h_min)
        self.h_max_slider.setValue(c.h_max); self.h_max_spin.setValue(c.h_max)
        self.s_min_slider.setValue(c.s_min); self.s_min_spin.setValue(c.s_min)
        self.s_max_slider.setValue(c.s_max); self.s_max_spin.setValue(c.s_max)
        self.v_min_slider.setValue(c.v_min); self.v_min_spin.setValue(c.v_min)
        self.v_max_slider.setValue(c.v_max); self.v_max_spin.setValue(c.v_max)

    def from_model(self):
        c = self.chroma
        self.tolA_spin.setValue(c.tolA); self.tolB_spin.setValue(c.tolB)
        self.use_hsv_chk.setChecked(c.use_hsv)
        self._sync_sliders_from_model()
        self.feather.setValue(c.feather)
        self.opacity_slider.setValue(int(clamp(c.ui_opacity, 0.3, 1.0) * 100))
        self.persist_chk.setChecked(c.persist)
        self._update_rgb_labels()

    def _update_rgb_labels(self):
        def css(rgb):
            if not rgb: return "background:#555; border:1px solid #777;"
            r, g, b = rgb
            return f"background: rgb({r},{g},{b}); border:1px solid #777;"
        self.rgbA_lbl.setText(f"A: {tuple(self.chroma.pickA) if self.chroma.pickA else '-'}")
        self.rgbB_lbl.setText(f"B: {tuple(self.chroma.pickB) if self.chroma.pickB else '-'}")
        self.rgbA_swatch.setStyleSheet(css(self.chroma.pickA))
        self.rgbB_swatch.setStyleSheet(css(self.chroma.pickB))

    def on_apply(self):
        # Прозрачность диалога НЕ меняем (по ТЗ)
        self.parent().apply_menu_opacity(self.chroma.ui_opacity)
        self.parent().request_repaint()

    def start_pick(self, which):
        self._pick_target = which

    def on_preview_click(self, event):
        if not hasattr(self, "_pick_target"):
            return
        frame = self.get_frame()
        if frame is None:
            return

        label_size = self.preview_label.size()
        h, w = frame.shape[:2]
        scale = min(label_size.width() / w, label_size.height() / h)
        disp_w = int(w * scale); disp_h = int(h * scale)
        off_x = (label_size.width() - disp_w) // 2
        off_y = (label_size.height() - disp_h) // 2
        x = event.pos().x() - off_x; y = event.pos().y() - off_y
        if x < 0 or y < 0 or x >= disp_w or y >= disp_h:
            return
        fx = int(x / scale); fy = int(y / scale)

        b, g, r = frame[fy, fx]
        rgb = [int(r), int(g), int(b)]
        if self._pick_target == 'A':
            self.chroma.pickA = rgb
            self.pickA_pos = (fx, fy)
        else:
            self.chroma.pickB = rgb
            self.pickB_pos = (fx, fy)
        self._update_rgb_labels()
        delattr(self, "_pick_target")

    def update_preview(self):
        frame = self.get_frame()
        if frame is None:
            self.preview_label.setText("Нет сигнала")
            return
        h, w = frame.shape[:2]
        label_size = self.preview_label.size()
        scale = min(label_size.width() / w, label_size.height() / h)
        disp_w = int(w * scale); disp_h = int(h * scale)
        off_x = (label_size.width() - disp_w) // 2
        off_y = (label_size.height() - disp_h) // 2

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        qimg = QImage(rgb.data, rgb.shape[1], rgb.shape[0], QImage.Format_RGB888)
        pix = QPixmap.fromImage(qimg).scaled(disp_w, disp_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        canvas = QPixmap(label_size.width(), label_size.height())
        canvas.fill(Qt.black)
        painter = QPainter(canvas)
        painter.drawPixmap(off_x, off_y, pix)

        # Рисуем точки пипеток
        def draw_point(pos, rgb_color, label_char):
            if not pos or not rgb_color:
                return
            fx, fy = pos
            dx = int(fx * scale) + off_x
            dy = int(fy * scale) + off_y
            r, g, b = rgb_color
            pen = QtGui.QPen(Qt.white); pen.setWidth(2)
            painter.setPen(pen)
            brush = QtGui.QBrush(QtGui.QColor(r, g, b))
            painter.setBrush(brush)
            painter.drawEllipse(QtCore.QPoint(dx, dy), 6, 6)
            painter.setPen(QtGui.QPen(QtGui.QColor(r, g, b)))
            painter.drawText(dx + 8, dy - 8, label_char)

        draw_point(self.pickA_pos, self.chroma.pickA, "A")
        draw_point(self.pickB_pos, self.chroma.pickB, "B")

        painter.end()
        self.preview_label.setPixmap(canvas)


class CropDialog(QDialog):
    def __init__(self, parent, get_frame_callable, initial_rect=None):
        super().__init__(parent)
        self.setWindowTitle("Кроп: выберите область")
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        self.setStyleSheet("""
            QDialog { background: #1e1e1e; color: #eee; }
            QLabel { color: #eee; }
            QPushButton { background:#2a2a2a; color:#eee; border:1px solid #555; padding:4px 10px; }
            QPushButton:hover { background:#333; }
        """)
        self.get_frame = get_frame_callable
        self.preview = QLabel(self); self.preview.setMinimumSize(640, 360)
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setStyleSheet("background:#111; border:1px solid #444; color:#eee;")
        self.info = QLabel("ЛКМ — тянуть прямоугольник; ПКМ — сброс; Enter — применить; Esc — отмена")
        self.btn_apply = QPushButton("Применить"); self.btn_cancel = QPushButton("Отмена"); self.btn_reset = QPushButton("Сброс")
        lay = QVBoxLayout(self); lay.addWidget(self.preview); lay.addWidget(self.info)
        row = QHBoxLayout(); row.addWidget(self.btn_apply); row.addWidget(self.btn_reset); row.addWidget(self.btn_cancel)
        lay.addLayout(row)

        self.timer = QTimer(self); self.timer.timeout.connect(self.render); self.timer.start(1000 // 30)
        self.preview.installEventFilter(self)

        self.dragging = False
        self.start_pt = None
        self.end_pt = None
        self.frame_cache = None
        self.initial_rect = initial_rect

        self.btn_apply.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_reset.clicked.connect(self.reset_rect)

    def reset_rect(self):
        self.start_pt = None; self.end_pt = None; self.initial_rect = None

    def eventFilter(self, obj, ev):
        if obj is self.preview:
            if ev.type() == QtCore.QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
                self.dragging = True; self.start_pt = ev.pos(); self.end_pt = ev.pos(); return True
            if ev.type() == QtCore.QEvent.MouseMove and self.dragging:
                self.end_pt = ev.pos(); return True
            if ev.type() == QtCore.QEvent.MouseButtonRelease and ev.button() == Qt.LeftButton:
                self.dragging = False; return True
            if ev.type() == QtCore.QEvent.MouseButtonPress and ev.button() == Qt.RightButton:
                self.reset_rect(); return True
        return super().eventFilter(obj, ev)

    def keyPressEvent(self, ev):
        if ev.key() in (Qt.Key_Return, Qt.Key_Enter): self.accept(); return
        if ev.key() == Qt.Key_Escape: self.reject(); return
        return super().keyPressEvent(ev)

    def render(self):
        frame = self.get_frame()
        if frame is None:
            self.preview.setText("Нет сигнала"); return
        self.frame_cache = frame
        h, w = frame.shape[:2]
        lbl_w = self.preview.width(); lbl_h = self.preview.height()
        scale = min(lbl_w / w, lbl_h / h)
        disp_w = int(w * scale); disp_h = int(h * scale)
        off_x = (lbl_w - disp_w) // 2; off_y = (lbl_h - disp_h) // 2
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        qimg = QImage(rgb.data, w, h, QImage.Format_RGB888)
        pix = QPixmap.fromImage(qimg).scaled(disp_w, disp_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        canvas = QPixmap(lbl_w, lbl_h); canvas.fill(Qt.transparent)
        painter = QPainter(canvas); painter.fillRect(0, 0, lbl_w, lbl_h, Qt.black); painter.drawPixmap(off_x, off_y, pix)
        if self.start_pt and self.end_pt:
            rect = QtCore.QRect(self.start_pt, self.end_pt).normalized()
            pen = QtGui.QPen(Qt.green); pen.setWidth(2); painter.setPen(pen); painter.drawRect(rect)
        elif self.initial_rect is not None:
            x, y, w0, h0 = self.initial_rect
            rx = int(x * scale) + off_x; ry = int(y * scale) + off_y; rw = int(w0 * scale); rh = int(h0 * scale)
            pen = QtGui.QPen(Qt.yellow); pen.setWidth(2); painter.setPen(pen); painter.drawRect(QtCore.QRect(rx, ry, rw, rh))
        painter.end()
        self.preview.setPixmap(canvas)

    def selected_rect_source_coords(self):
        frame = self.frame_cache
        if frame is None: return None
        h, w = frame.shape[:2]
        lbl_w = self.preview.width(); lbl_h = self.preview.height()
        scale = min(lbl_w / w, lbl_h / h)
        off_x = (lbl_w - int(w * scale)) // 2; off_y = (lbl_h - int(h * scale)) // 2
        if self.start_pt and self.end_pt:
            rect = QtCore.QRect(self.start_pt, self.end_pt).normalized()
            x = clamp(rect.x() - off_x, 0, lbl_w) / scale
            y = clamp(rect.y() - off_y, 0, lbl_h) / scale
            rw = clamp(rect.width(), 1, lbl_w) / scale
            rh = clamp(rect.height(), 1, lbl_h) / scale
            x = int(clamp(x, 0, w - 1)); y = int(clamp(y, 0, h - 1))
            rw = int(clamp(rw, 1, w - x)); rh = int(clamp(rh, 1, h - y))
            return [x, y, rw, rh]
        return self.initial_rect


# --- Основное окно ---
class RoundCamWindow(QWidget):
    def __init__(self, args):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        # Стартуем без кнопки в панели задач (Qt.Tool)
        self.base_flags = Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint | Qt.Tool
        self.setWindowFlags(self.base_flags)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        # Иконка
        icon_path = os.path.join(app_exe_dir(), "icon.ico")
        self.icon = QIcon(icon_path) if os.path.exists(icon_path) else self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon)
        self.setWindowIcon(self.icon)

        # Состояние
        self.state = AppState(); self.chroma = ChromaConfig(); self.crop = CropConfig()
        self.camera_index = args.camera; self.vcam_w, self.vcam_h = args.vcam_res; self.fps = args.fps
        self.req_width = args.width; self.req_height = args.height
        self.vcam = None; self.last_frame_bgr = None

        self.load_config()
        self.setWindowOpacity(self.state.window_opacity)

        # Камера
        backend = cv2.CAP_DSHOW if os.name == 'nt' else cv2.CAP_ANY
        _logger.info("Открываем камеру index=%s backend=%s", self.camera_index, "CAP_DSHOW" if backend == cv2.CAP_DSHOW else "CAP_ANY")
        self.cap = cv2.VideoCapture(self.camera_index, backend)
        if not self.cap.isOpened():
            _logger.error("Не удалось открыть камеру index=%s", self.camera_index)
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть камеру с индексом {self.camera_index}")
            sys.exit(1)
        if self.req_width and self.req_height:
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.req_width)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.req_height)
        if self.fps:
            self.cap.set(cv2.CAP_PROP_FPS, self.fps)
        log_cap_props(self.cap)

        self.resize(self.state.circle_diameter, self.state.circle_diameter)
        self.move(self.state.pos_x, self.state.pos_y)
        self._update_mask()

        # Таймер кадра
        self.timer = QTimer(self); self.timer.timeout.connect(self.on_tick); self.timer.start(int(1000 / (self.fps or 30)))

        # Счётчик ошибок чтения кадра
        self._no_frame_count = 0

        # Горячие клавиши
        QShortcut(QtGui.QKeySequence("Ctrl+Q"), self, activated=self.close)
        QShortcut(QtGui.QKeySequence("V"), self, activated=lambda: self.set_vcam_enabled(not self.state.vcam_enabled))
        QShortcut(QtGui.QKeySequence("B"), self, activated=lambda: self.set_vcam_mirror(not self.state.vcam_mirror))
        QShortcut(QtGui.QKeySequence("M"), self, activated=lambda: self.set_window_mirror(not self.state.window_mirror))
        QShortcut(QtGui.QKeySequence("K"), self, activated=lambda: self.set_chroma_enabled(not self.chroma.enabled))
        QShortcut(QtGui.QKeySequence("+"), self, activated=lambda: self.scale_circle(1))
        QShortcut(QtGui.QKeySequence("="), self, activated=lambda: self.scale_circle(1))
        QShortcut(QtGui.QKeySequence("-"), self, activated=lambda: self.scale_circle(-1))

        # Контекстное меню (окно)
        self.menu = QMenu(self); self._apply_dark_menu_style(self.menu)
        self.menu.aboutToShow.connect(lambda: self.menu.setWindowOpacity(self.chroma.ui_opacity))

        # --- Окно (подменю)
        window_menu = QMenu("Окно", self); self._apply_dark_menu_style(window_menu)

        # Пункт «Выключить/Включить камеру»
        self.act_cam_toggle = QAction("Выключить камеру", self)
        self.act_cam_toggle.triggered.connect(self.toggle_camera)
        window_menu.addAction(self.act_cam_toggle)
        window_menu.addSeparator()

        self.act_on_top = QAction("Всегда поверх", self, checkable=True)
        self.act_mirror_window = QAction("Отразить зеркально", self, checkable=True)
        self.act_click_through = QAction("Прокликиваемый", self, checkable=True)
        window_menu.addAction(self.act_on_top); window_menu.addAction(self.act_mirror_window); window_menu.addAction(self.act_click_through)

        # --- Форма окна
        shape_menu = QMenu("Форма окна", self); self._apply_dark_menu_style(shape_menu)
        self.act_shape_circle = QAction("Круг", self, checkable=True)
        self.act_shape_square = QAction("Квадрат", self, checkable=True)
        grp = QtWidgets.QActionGroup(self)
        for a in (self.act_shape_circle, self.act_shape_square):
            a.setActionGroup(grp); a.setCheckable(True)
        shape_menu.addAction(self.act_shape_circle); shape_menu.addAction(self.act_shape_square)
        window_menu.addMenu(shape_menu)

        # --- Кроп
        crop_menu = QMenu("Кроп", self); self._apply_dark_menu_style(crop_menu)
        self.act_crop_enable = QAction("Вкл/выкл", self, checkable=True)
        self.act_crop_pick = QAction("Выбрать область…", self)
        crop_menu.addAction(self.act_crop_enable); crop_menu.addAction(self.act_crop_pick)
        window_menu.addMenu(crop_menu)

        # --- Хромакей
        chroma_menu = QMenu("Хромакей", self); self._apply_dark_menu_style(chroma_menu)
        self.act_chroma = QAction("Вкл/выкл", self, checkable=True)
        self.act_chroma_settings = QAction("Настройки…", self)
        chroma_menu.addAction(self.act_chroma); chroma_menu.addAction(self.act_chroma_settings)
        window_menu.addMenu(chroma_menu)

        # Добавляем «Окно» в основное меню
        self.menu.addMenu(window_menu)

        # --- Виртуальная камера
        vcam_menu = QMenu("Виртуальная камера", self); self._apply_dark_menu_style(vcam_menu)
        self.act_vcam_enable = QAction("Вкл/выкл", self, checkable=True)
        self.act_mirror_vcam = QAction("Отразить зеркально", self, checkable=True)
        self.act_vcam_fit_fill = QAction("По ширине окна", self, checkable=True)
        vcam_menu.addAction(self.act_vcam_enable); vcam_menu.addAction(self.act_mirror_vcam); vcam_menu.addAction(self.act_vcam_fit_fill)
        self.menu.addMenu(vcam_menu)

        # --- Прозрачность (ползунок) в меню окна
        self.menu.addSeparator()
        self.opacity_label = QLabel(f"Прозрачность: {int(self.state.window_opacity * 100)}%"); self.opacity_label.setStyleSheet("color: rgba(255,255,255,0.92);")
        self.opacity_slider = QSlider(Qt.Horizontal); self.opacity_slider.setRange(20, 100); self.opacity_slider.setValue(int(clamp(self.state.window_opacity, 0.2, 1.0) * 100))
        w = QtWidgets.QWidget(self.menu); lay = QHBoxLayout(w); lay.setContentsMargins(10, 6, 10, 6); lay.addWidget(self.opacity_label); lay.addWidget(self.opacity_slider)
        self.opacity_widget_action = QWidgetAction(self.menu); self.opacity_widget_action.setDefaultWidget(w)
        self.menu.addAction(self.opacity_widget_action)
        self.opacity_slider.valueChanged.connect(self._on_opacity_changed)

        # --- О программе / Выход
        self.menu.addSeparator()
        self.act_about = QAction("О программе…", self)
        self.act_exit = QAction("Выход", self)
        self.menu.addAction(self.act_about); self.menu.addAction(self.act_exit)

        # --- Трей
        self.tray = QSystemTrayIcon(self.icon, self)
        self.tray.setToolTip(APP_NAME)
        self.tray.setIcon(self.icon)
        self.tray_menu = QMenu(self); self._apply_dark_menu_style(self.tray_menu)

        self.tray_act_show = QAction("Показать окно", self); self.tray_act_show.setCheckable(True); self.tray_act_show.setChecked(True)
        self.tray_act_show.toggled.connect(self._on_tray_show_toggled)
        self.tray_menu.addAction(self.tray_act_show); self.tray_menu.addSeparator()

        window_tray_menu = self.tray_menu.addMenu("Окно"); self._apply_dark_menu_style(window_tray_menu)
        window_tray_menu.addAction(self.act_on_top)
        window_tray_menu.addAction(self.act_mirror_window)
        window_tray_menu.addAction(self.act_click_through)

        shape_tray_menu = window_tray_menu.addMenu("Форма окна"); self._apply_dark_menu_style(shape_tray_menu)
        shape_tray_menu.addAction(self.act_shape_circle); shape_tray_menu.addAction(self.act_shape_square)

        crop_tray_menu = window_tray_menu.addMenu("Кроп"); self._apply_dark_menu_style(crop_tray_menu)
        crop_tray_menu.addAction(self.act_crop_enable); crop_tray_menu.addAction(self.act_crop_pick)

        chroma_tray_menu = window_tray_menu.addMenu("Хромакей"); self._apply_dark_menu_style(chroma_tray_menu)
        chroma_tray_menu.addAction(self.act_chroma); chroma_tray_menu.addAction(self.act_chroma_settings)

        vcam_tray_menu = self.tray_menu.addMenu("Виртуальная камера"); self._apply_dark_menu_style(vcam_tray_menu)
        vcam_tray_menu.addAction(self.act_vcam_enable); vcam_tray_menu.addAction(self.act_mirror_vcam); vcam_tray_menu.addAction(self.act_vcam_fit_fill)

        # --- Прозрачность (ползунок) в ТРЕЕ
        self.tray_menu.addSeparator()
        self.tray_opacity_label = QLabel(f"Прозрачность: {int(self.state.window_opacity * 100)}%"); self.tray_opacity_label.setStyleSheet("color: rgba(255,255,255,0.92);")
        self.tray_opacity_slider = QSlider(Qt.Horizontal); self.tray_opacity_slider.setRange(20, 100); self.tray_opacity_slider.setValue(int(clamp(self.state.window_opacity, 0.2, 1.0) * 100))
        tw = QtWidgets.QWidget(self.tray_menu); tlay = QHBoxLayout(tw); tlay.setContentsMargins(10, 6, 10, 6); tlay.addWidget(self.tray_opacity_label); tlay.addWidget(self.tray_opacity_slider)
        self.tray_opacity_action = QWidgetAction(self.tray_menu); self.tray_opacity_action.setDefaultWidget(tw)
        self.tray_menu.addAction(self.tray_opacity_action)
        self.tray_opacity_slider.valueChanged.connect(self._on_tray_opacity_changed)

        # --- Калькулятор в трее
        self.tray_menu.addSeparator()
        calc_menu = self.tray_menu.addMenu("Калькулятор"); self._apply_dark_menu_style(calc_menu)
        self.act_calc_toggle = QAction("Показать/Свернуть", self)
        self.act_calc_numlock = QAction("NumLock всегда включён", self, checkable=True)
        self.act_calc_numlock.setChecked(calc_settings["force_numlock"])
        self.act_calc_toggle.triggered.connect(toggle_calc)
        self.act_calc_numlock.toggled.connect(lambda v: calc_settings.update({"force_numlock": v}))
        calc_menu.addAction(self.act_calc_toggle); calc_menu.addAction(self.act_calc_numlock)

        # --- О программе / Выход
        self.tray_menu.addSeparator(); self.tray_menu.addAction(self.act_about); self.tray_menu.addAction(self.act_exit)
        self.tray.setContextMenu(self.tray_menu)
        # ЛКМ (Trigger) и ПКМ (Context) — показываем меню у курсора
        self.tray.activated.connect(self._on_tray_activated)
        self.tray.show()

        # Коннекты действий
        self.act_on_top.toggled.connect(self.set_always_on_top)
        self.act_click_through.toggled.connect(self.set_click_through)
        self.act_mirror_window.toggled.connect(self.set_window_mirror)
        self.act_vcam_enable.toggled.connect(self.set_vcam_enabled)
        self.act_mirror_vcam.toggled.connect(self.set_vcam_mirror)
        self.act_chroma.toggled.connect(self.set_chroma_enabled)
        self.act_chroma_settings.triggered.connect(self.open_chroma_settings)
        self.act_crop_enable.toggled.connect(self.set_crop_enabled)
        self.act_crop_pick.triggered.connect(self.open_crop_dialog)
        self.act_shape_circle.toggled.connect(lambda v: v and self.set_window_shape("circle"))
        self.act_shape_square.toggled.connect(lambda v: v and self.set_window_shape("square"))
        self.act_vcam_fit_fill.toggled.connect(self.set_vcam_fit_fill)
        self.act_about.triggered.connect(self.show_about)
        self.act_exit.triggered.connect(self.force_quit)

        # Применяем сохранённое состояние
        self.set_always_on_top(self.state.always_on_top)
        self.set_click_through(self.state.click_through)
        self.set_window_mirror(self.state.window_mirror)
        self.set_vcam_enabled(self.state.vcam_enabled)
        self.set_vcam_mirror(self.state.vcam_mirror)
        self.set_chroma_enabled(self.chroma.enabled)
        self.set_crop_enabled(self.crop.enabled)
        self.set_window_shape(self.state.window_shape)
        self.act_vcam_fit_fill.setChecked(self.state.vcam_fit == "fill")

        self.dragging = False; self.drag_pos = None
        self._update_cam_toggle_caption()
        _logger.info("Инициализация окна завершена")

    # --- Единый тёмный стиль меню ---
    def _dark_qmenu_stylesheet(self) -> str:
        alpha = int(clamp(self.chroma.ui_opacity, 0.3, 1.0) * 255)
        bg = f"rgba(30,30,30,{alpha})"
        sel = f"rgba(255,255,255,{int(alpha * 0.25)})"
        txt = "rgba(255,255,255,230)"
        border = "rgba(255,255,255,80)"
        return (
            f"QMenu {{ background-color: {bg}; color:{txt}; border:1px solid {border}; }}"
            f"QMenu::item:selected {{ background-color:{sel}; }}"
            f"QMenu::separator {{ height:1px; background: rgba(255,255,255,45); margin:4px 8px; }}"
        )

    def _apply_dark_menu_style(self, menu: QMenu):
        menu.setStyleSheet(self._dark_qmenu_stylesheet())
        orig_add_menu = menu.addMenu
        def wrapped_add_menu(*args, **kwargs):
            m = orig_add_menu(*args, **kwargs)
            if isinstance(m, QMenu):
                m.setStyleSheet(self._dark_qmenu_stylesheet())
            return m
        menu.addMenu = wrapped_add_menu

    # --- Трей ---
    def _on_tray_show_toggled(self, checked: bool):
        _logger.info("Трей: показать окно = %s", checked)
        if checked:
            # >>> ИЗМЕНЕНО: показываем как Tool (без кнопки в таскбаре)
            flags = self.base_flags
            if self.state.always_on_top:
                flags |= Qt.WindowStaysOnTopHint
            self.setWindowFlags(flags)
            self.show()         # showNormal() не нужен для Tool-окна
            self.raise_(); self.activateWindow()

            if not (self.cap and self.cap.isOpened()):
                _logger.info("Камера закрыта — повторное открытие index=%s", self.camera_index)
                backend = cv2.CAP_DSHOW if os.name == 'nt' else cv2.CAP_ANY
                self.cap = cv2.VideoCapture(self.camera_index, backend)
                if self.req_width and self.req_height:
                    self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.req_width)
                    self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.req_height)
                if self.fps:
                    self.cap.set(cv2.CAP_PROP_FPS, self.fps)
                log_cap_props(self.cap)
                if not self.timer.isActive():
                    self.timer.start(int(1000 / (self.fps or 30)))
            if self.state.vcam_enabled:
                self.start_vcam()
        else:
            # Скрыть в трей и убрать кнопку с панели задач
            try:
                if getattr(self, "vcam", None):
                    self.vcam.close()
                    _logger.info("Виртуальная камера остановлена при скрытии окна")
            except Exception as e:
                log_exc("Ошибка закрытия виртуальной камеры", e)
            self.vcam = None
            try:
                if getattr(self, "cap", None) and self.cap.isOpened():
                    self.cap.release()
                    _logger.info("Камера освобождена при скрытии окна")
            except Exception as e:
                log_exc("Ошибка освобождения камеры", e)
            if self.timer.isActive():
                self.timer.stop()
            flags = self.base_flags
            if self.state.always_on_top:
                flags |= Qt.WindowStaysOnTopHint
            self.setWindowFlags(flags)
            self.hide()
            self._update_cam_toggle_caption()

    def _on_tray_activated(self, reason):
        pos = QtGui.QCursor.pos()
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.Context):
            self.tray.contextMenu().popup(pos)

    # --- Прозрачность ---
    def _on_opacity_changed(self, v: int):
        self.opacity_label.setText(f"Прозрачность: {v}%")
        self.set_window_opacity(v / 100.0)
        self._apply_dark_menu_style(self.menu); self._apply_dark_menu_style(self.tray_menu)

    def _on_tray_opacity_changed(self, v: int):
        self.tray_opacity_label.setText(f"Прозрачность: {v}%")
        self.set_window_opacity(v / 100.0)
        self._apply_dark_menu_style(self.menu); self._apply_dark_menu_style(self.tray_menu)

    def set_window_opacity(self, val: float):
        val = clamp(val, 0.2, 1.0)
        self.state.window_opacity = val
        self.setWindowOpacity(val)
        if self.chroma.persist:
            self.save_config()

    # --- Конфиг ---
    def load_config(self):
        ensure_first_run_files()
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if "state" in data:
                    for k in vars(AppState()).keys():
                        if k in data["state"]:
                            setattr(self.state, k, data["state"][k])
                if "chroma" in data:
                    for k in vars(ChromaConfig()).keys():
                        if k in data["chroma"]:
                            setattr(self.chroma, k, data["chroma"][k])
                if "crop" in data:
                    for k in vars(CropConfig()).keys():
                        if k in data["crop"]:
                            setattr(self.crop, k, data["crop"][k])
            _logger.info("Конфиг загружен")
        except Exception as e:
            log_exc("Ошибка загрузки конфига", e)

    def save_config(self):
        try:
            data = {"state": asdict(self.state), "chroma": asdict(self.chroma), "crop": asdict(self.crop)}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            _logger.info("Конфиг сохранён")
        except Exception as e:
            log_exc("Ошибка сохранения конфига", e)

    # --- Режимы окна/виртуалки ---
    def set_always_on_top(self, enabled: bool):
        self.state.always_on_top = enabled; self.act_on_top.setChecked(enabled)
        flags = self.windowFlags()
        if enabled: flags |= Qt.WindowStaysOnTopHint
        else: flags &= ~Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags); self.show()
        _logger.info("Флаг 'всегда поверх' = %s", enabled)

    def _win_set_exstyle_flag(self, add: bool, flag: int):
        try:
            if os.name != "nt":
                return
            hwnd = int(self.winId())
            GWL_EXSTYLE = -20
            user32 = ctypes.windll.user32
            get_long = user32.GetWindowLongW
            set_long = user32.SetWindowLongW
            exstyle = get_long(hwnd, GWL_EXSTYLE)
            if add:
                exstyle |= flag
            else:
                exstyle &= ~flag
            set_long(hwnd, GWL_EXSTYLE, exstyle)
        except Exception as e:
            log_exc("Ошибка установки exstyle окна", e)

    def _apply_click_through_windows(self, enabled: bool):
        if os.name != "nt":
            return
        WS_EX_LAYERED = 0x00080000
        WS_EX_TRANSPARENT = 0x00000020
        self._win_set_exstyle_flag(True, WS_EX_LAYERED)
        self._win_set_exstyle_flag(enabled, WS_EX_TRANSPARENT)

    def set_click_through(self, enabled: bool):
        self.state.click_through = enabled; self.act_click_through.setChecked(enabled)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, enabled)
        self._apply_click_through_windows(enabled)
        self.show()
        _logger.info("Прокликивание окна = %s", enabled)

    def set_window_mirror(self, enabled: bool):
        self.state.window_mirror = enabled; self.act_mirror_window.setChecked(enabled)
        _logger.info("Зеркало окна = %s", enabled)

    def set_vcam_enabled(self, enabled: bool):
        self.state.vcam_enabled = enabled; self.act_vcam_enable.setChecked(enabled)
        _logger.info("Виртуальная камера включена = %s", enabled)
        if enabled: self.start_vcam()
        else: self.stop_vcam()

    def set_vcam_mirror(self, enabled: bool):
        self.state.vcam_mirror = enabled; self.act_mirror_vcam.setChecked(enabled)
        _logger.info("Зеркало виртуальной камеры = %s", enabled)

    def set_chroma_enabled(self, enabled: bool):
        self.chroma.enabled = enabled; self.act_chroma.setChecked(enabled)
        _logger.info("Хромакей = %s", enabled)

    def set_crop_enabled(self, enabled: bool):
        self.crop.enabled = enabled; self.act_crop_enable.setChecked(enabled)
        _logger.info("Кроп = %s", enabled)

    def set_window_shape(self, shape: str):
        self.state.window_shape = shape
        self.act_shape_circle.setChecked(shape == "circle"); self.act_shape_square.setChecked(shape == "square")
        self._update_mask(); self.update()
        _logger.info("Форма окна = %s", shape)

    def set_vcam_fit_fill(self, enabled: bool):
        self.state.vcam_fit = "fill" if enabled else "letterbox"
        _logger.info("VCAM fit mode = %s", self.state.vcam_fit)

    # --- Диалоги/меню вспомогательные ---
    def open_chroma_settings(self):
        _logger.info("Открыт диалог настроек хромакея")
        dlg = ChromaDialog(self, self.chroma, self.get_last_frame)
        dlg.show()

    def open_crop_dialog(self):
        _logger.info("Открыт диалог кропа")
        dlg = CropDialog(self, self.get_last_frame, initial_rect=self.crop.rect)
        if dlg.exec_() == QDialog.Accepted:
            rect = dlg.selected_rect_source_coords()
            self.crop.rect = rect; self.crop.enabled = rect is not None
            _logger.info("Кроп выбран: %s", rect)
            self.request_repaint()

    def apply_menu_opacity(self, opacity: float):
        self._apply_dark_menu_style(self.menu); self._apply_dark_menu_style(self.tray_menu)

    # --- О программе ---
    def show_about(self):
        channel_html = '<a href="https://t.me/RoundCam">КругоЗор</a>'
        contact_html = '<a href="https://t.me/AKudlay_ru">Андрей Кудлай</a>'
        text = f"""<b>{APP_NAME}</b>
        <br>Версия: {APP_VERSION}<br>
<h3>Горячие клавиши</h3>
NumLock — показать/свернуть калькулятор<br>
"+/-" Масштаб окна<br>
"K"     Хромакей окна<br>
"M"     Зеркало окна<br>
"V"     Вкл/Выкл виртуальную<br>
"Ctrl+Q" Выход<br>
<h3>Связь</h3> Автор: {contact_html}<br>Канал: {channel_html}"""
        box = QMessageBox(self)
        box.setWindowTitle("О программе")
        box.setIcon(QMessageBox.Information)
        box.setTextFormat(QtCore.Qt.RichText)
        box.setTextInteractionFlags(Qt.TextBrowserInteraction | Qt.LinksAccessibleByMouse)
        box.setStandardButtons(QMessageBox.Ok)
        box.setText(text)
        box.exec_()
        _logger.info("Показано окно 'О программе'")

    # --- Рендер/кадры ---
    def request_repaint(self):
        self.update()

    def on_tick(self):
        try:
            if not (self.cap and self.cap.isOpened()):
                # Камера закрыта — просто выходим
                return
            ok, frame = self.cap.read()
            if not ok or frame is None:
                self._no_frame_count += 1
                if self._no_frame_count in (1, 10, 50) or (self._no_frame_count % 100 == 0):
                    _logger.warning("Нет кадра от камеры: ok=%s frame=%s count=%d", ok, "None" if frame is None else "ndarray", self._no_frame_count)
                    log_cap_props(self.cap)
                return
            if self._no_frame_count:
                _logger.info("Кадры восстановились после %d пустых чтений", self._no_frame_count)
                self._no_frame_count = 0

            self.last_frame_bgr = frame
            src = frame
            if self.crop.enabled and self.crop.rect:
                x, y, w, h = self.crop.rect
                x2 = clamp(x, 0, src.shape[1] - 1); y2 = clamp(y, 0, src.shape[0] - 1)
                w2 = clamp(w, 1, src.shape[1] - x2); h2 = clamp(h, 1, src.shape[0] - y2)
                src = src[y2:y2 + h2, x2:x2 + w2].copy()

            disp = src.copy()
            if self.state.window_mirror:
                disp = cv2.flip(disp, 1)

            target = max(1, self.size().width())
            h, w = disp.shape[:2]
            scale = max(target / max(1, w), target / max(1, h))
            nw, nh = int(ceil(w * scale)), int(ceil(h * scale))
            if nw < 1: nw = 1
            if nh < 1: nh = 1
            rs = cv2.resize(disp, (nw, nh), interpolation=cv2.INTER_LINEAR)
            start_x = max(0, (rs.shape[1] - target) // 2); start_y = max(0, (rs.shape[0] - target) // 2)
            end_x = min(rs.shape[1], start_x + target); end_y = min(rs.shape[0], start_y + target)
            roi = rs[start_y:end_y, start_x:end_x].copy()
            if roi.shape[0] != target or roi.shape[1] != target:
                pad = np.zeros((target, target, 3), dtype=roi.dtype)
                pad[:roi.shape[0], :roi.shape[1]] = roi
                roi = pad

            if self.chroma.enabled:
                alpha = self.build_alpha_mask(roi)
            else:
                alpha = np.full((roi.shape[0], roi.shape[1]), 255, dtype=np.uint8)

            if self.state.window_shape == "circle":
                shape_mask = self.circle_alpha_mask(roi.shape[1])
            else:
                shape_mask = np.full((roi.shape[0], roi.shape[1]), 255, dtype=np.uint8)

            alpha = (alpha.astype(np.uint16) * shape_mask.astype(np.uint16) // 255).astype(np.uint8)
            rgb = cv2.cvtColor(roi, cv2.COLOR_BGR2RGB)
            argb = np.dstack((rgb, alpha))
            qimg = QImage(argb.data, argb.shape[1], argb.shape[0], QImage.Format_RGBA8888)
            self._current_qpix = QPixmap.fromImage(qimg)
            self.update()

            # >>> ИЗМЕНЕНО: для виртуальной камеры берём именно квадратный ROI из окна
            if self.vcam is not None:
                vframe = roi.copy()  # квадратный BGR из окна
                if self.state.vcam_mirror:
                    vframe = cv2.flip(vframe, 1)
                vframe_rgb = cv2.cvtColor(vframe, cv2.COLOR_BGR2RGB)
                if self.state.vcam_fit == "fill":
                    # Заполнение по ширине виртуалки с центр-кропом по высоте (без искажений)
                    out = self.center_crop_fit(vframe_rgb, self.vcam_w, self.vcam_h)
                else:
                    # Вписывание с полями (letterbox/pillarbox), без искажений
                    out = self.letterbox(vframe_rgb, self.vcam_w, self.vcam_h)
                try:
                    self.vcam.send(out); self.vcam.sleep_until_next_frame()
                except Exception as e:
                    log_exc("Ошибка отправки кадра в виртуальную камеру", e)
                    self.stop_vcam()
        except Exception as e:
            log_exc("Исключение в on_tick", e)

    def build_alpha_mask(self, bgr_img):
        h, w = bgr_img.shape[:2]
        if self.chroma.use_hsv:
            hsv = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2HSV)
            lower = np.array([self.chroma.h_min, self.chroma.s_min, self.chroma.v_min], dtype=np.uint8)
            upper = np.array([self.chroma.h_max, self.chroma.s_max, self.chroma.v_max], dtype=np.uint8)
            mask = cv2.inRange(hsv, lower, upper)
        else:
            mask = np.zeros((h, w), dtype=np.uint8)
            for pick, tol in [(self.chroma.pickA, self.chroma.tolA), (self.chroma.pickB, self.chroma.tolB)]:
                if pick is None: continue
                r, g, b = pick; target = np.array([b, g, r], dtype=np.int16)
                diff = np.abs(bgr_img.astype(np.int16) - target[None, None, :])
                sq = (diff.astype(np.int32) * diff.astype(np.int32)).sum(axis=2)
                dist = np.sqrt(sq).astype(np.float32)
                dist = np.nan_to_num(dist, nan=0.0, posinf=255.0, neginf=0.0)
                m = (dist <= tol).astype(np.uint8) * 255
                mask = np.maximum(mask, m)
        k = max(0, int(self.chroma.feather))
        if k % 2 == 0: k = k + 1 if k > 0 else 0
        if k >= 3: mask = cv2.GaussianBlur(mask, (k, k), 0)
        alpha = cv2.subtract(np.full_like(mask, 255), mask)
        return alpha

    def circle_alpha_mask(self, size):
        y, x = np.ogrid[:size, :size]
        c = (size - 1) / 2.0; r = size / 2.0
        dist = np.sqrt((x - c) ** 2 + (y - c) ** 2)
        return ((dist <= r).astype(np.uint8) * 255)

    def center_crop_fit(self, rgb, W, H):
        h, w = rgb.shape[:2]
        scale = max(W / w, H / h)
        nw, nh = int(ceil(w * scale)), int(ceil(h * scale))
        rs = cv2.resize(rgb, (nw, nh), interpolation=cv2.INTER_AREA)
        x = max(0, (nw - W) // 2); y = max(0, (nh - H) // 2)
        x2 = min(nw, x + W); y2 = min(nh, y + H)
        out = rs[y:y2, x:x2].copy()
        if out.shape[0] != H or out.shape[1] != W:
            pad = np.zeros((H, W, 3), dtype=out.dtype)
            pad[:out.shape[0], :out.shape[1]] = out
            out = pad
        return out

    def letterbox(self, rgb, W, H):
        h, w = rgb.shape[:2]
        scale = min(W / w, H / h)
        nw, nh = int(w * scale), int(h * scale)
        resized = cv2.resize(rgb, (nw, nh), interpolation=cv2.INTER_AREA)
        out = np.zeros((H, W, 3), dtype=np.uint8)
        x = (W - nw) // 2; y = (H - nh) // 2
        out[y:y + nh, x:x + nw] = resized
        return out

    # --- Маска формы окна/события окна ---
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.SmoothPixmapTransform, True)
        if hasattr(self, "_current_qpix"): p.drawPixmap(0, 0, self.width(), self.height(), self._current_qpix)

    def _update_mask(self):
        if self.state.window_shape == "circle":
            region = QRegion(self.rect(), QRegion.Ellipse)
        else:
            region = QRegion(self.rect())
        self.setMask(region)

    def resizeEvent(self, event):
        self._update_mask(); return super().resizeEvent(event)

    def mousePressEvent(self, event):
        if self.state.click_through:
            event.ignore(); return
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.drag_pos = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
        else:
            event.ignore()

    def mouseMoveEvent(self, event):
        if self.dragging and not self.state.click_through:
            self.move(event.globalPos() - self.drag_pos); event.accept()
        else:
            event.ignore()

    def mouseReleaseEvent(self, event):
        if self.dragging:
            self.dragging = False
            g = self.geometry(); self.state.pos_x, self.state.pos_y = g.left(), g.top()
        else:
            if self.state.click_through:
                event.ignore()

    def contextMenuEvent(self, event):
        if self.state.click_through: return
        self.act_on_top.setChecked(self.state.always_on_top)
        self.act_click_through.setChecked(self.state.click_through)
        self.act_mirror_window.setChecked(self.state.window_mirror)
        self.act_vcam_enable.setChecked(self.state.vcam_enabled)
        self.act_mirror_vcam.setChecked(self.state.vcam_mirror)
        self.act_chroma.setChecked(self.chroma.enabled)
        self.act_crop_enable.setChecked(self.crop.enabled)
        self.act_shape_circle.setChecked(self.state.window_shape == "circle")
        self.act_shape_square.setChecked(self.state.window_shape == "square")
        self.act_vcam_fit_fill.setChecked(self.state.vcam_fit == "fill")
        self.opacity_slider.setValue(int(clamp(self.state.window_opacity, 0.2, 1.0) * 100))
        self._update_cam_toggle_caption()
        self.menu.exec_(event.globalPos())

    # --- Камера вкл/выкл для пункта меню окна ---
    def _update_cam_toggle_caption(self):
        if self.cap and self.cap.isOpened():
            self.act_cam_toggle.setText("Выключить камеру")
        else:
            self.act_cam_toggle.setText("Включить камеру")

    def toggle_camera(self):
        if self.cap and self.cap.isOpened():
            _logger.info("Отключение камеры пользователем")
            try:
                self.cap.release()
                _logger.info("Камера освобождена")
            except Exception as e:
                log_exc("Ошибка освобождения камеры", e)
            if self.timer.isActive():
                self.timer.stop()
            # >>> ДОБАВЛЕНО: синхронизация с треем — спрятать окно
            if self.tray_act_show.isChecked():
                self.tray_act_show.setChecked(False)  # триггерит _on_tray_show_toggled(False)
        else:
            _logger.info("Попытка включить камеру index=%s", self.camera_index)
            backend = cv2.CAP_DSHOW if os.name == 'nt' else cv2.CAP_ANY
            self.cap = cv2.VideoCapture(self.camera_index, backend)
            if not self.cap.isOpened():
                _logger.error("Не удалось открыть камеру при повторном включении")
                QMessageBox.critical(self, "Ошибка", "Не удалось открыть камеру.")
            else:
                if self.req_width and self.req_height:
                    self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, self.req_width)
                    self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, self.req_height)
                if self.fps:
                    self.cap.set(cv2.CAP_PROP_FPS, self.fps)
                log_cap_props(self.cap)
                if not self.timer.isActive():
                    self.timer.start(int(1000 / (self.fps or 30)))
        self._update_cam_toggle_caption()

    def scale_circle(self, direction: int):
        d = clamp(self.state.circle_diameter + (20 * direction), 120, 1080)
        self.state.circle_diameter = int(d); self.resize(self.state.circle_diameter, self.state.circle_diameter)

    def get_last_frame(self):
        return self.last_frame_bgr

    def start_vcam(self):
        if not HAVE_PYVIRTUALCAM:
            _logger.warning("pyvirtualcam не установлен, виртуалка недоступна")
            QMessageBox.warning(self, "pyvirtualcam не установлен", "pyvirtualcam не найден. Установите зависимости и драйвер.")
            self.state.vcam_enabled = False; self.act_vcam_enable.setChecked(False); return
        try:
            if self.vcam is None:
                self.vcam = pyvirtualcam.Camera(width=self.vcam_w, height=self.vcam_h, fps=self.fps or 30, print_fps=False)
                _logger.info("Виртуальная камера запущена %dx%d @%s", self.vcam_w, self.vcam_h, self.fps or 30)
        except Exception as e:
            log_exc("Ошибка запуска виртуальной камеры", e)
            QMessageBox.critical(self, "Ошибка виртуальной камеры", f"Не удалось запустить виртуальную камеру:\n{e}")
            self.vcam = None; self.state.vcam_enabled = False; self.act_vcam_enable.setChecked(False)

    def stop_vcam(self):
        if self.vcam is not None:
            try:
                self.vcam.close()
                _logger.info("Виртуальная камера остановлена")
            except Exception as e:
                log_exc("Ошибка остановки виртуальной камеры", e)
            self.vcam = None

    def _cleanup(self):
        _logger.info("Очистка ресурсов...")
        try:
            if hasattr(self, "timer") and self.timer:
                self.timer.stop()
        except Exception as e: log_exc("Ошибка остановки таймера", e)
        try:
            if getattr(self, "vcam", None):
                self.vcam.close()
        except Exception as e: log_exc("Ошибка закрытия виртуальной камеры", e)
        try:
            if getattr(self, "cap", None):
                self.cap.release()
        except Exception as e: log_exc("Ошибка release камеры", e)
        try:
            cv2.destroyAllWindows()
        except Exception as e: log_exc("Ошибка destroyAllWindows", e)

    def force_quit(self):
        try:
            if hasattr(self, "tray"):
                self.tray.hide(); self.tray.deleteLater()
        except Exception as e:
            log_exc("Ошибка скрытия трея", e)
        self._cleanup()
        _logger.info("Завершение приложения по команде пользователя")
        self.close()
        QApplication.quit()

    def closeEvent(self, event: QCloseEvent):
        if self.chroma.persist:
            self.save_config()
        self._cleanup()
        try:
            if hasattr(self, "tray"):
                self.tray.hide(); self.tray.deleteLater()
        except Exception as e:
            log_exc("Ошибка закрытия трея", e)
        _logger.info("Окно закрыто")
        event.accept()

# --- CLI / main ---
def parse_size(s: str):
    try:
        parts = s.lower().split('x'); return int(parts[0]), int(parts[1])
    except Exception:
        raise argparse.ArgumentTypeError("Ожидается формат WxH, напр. 1920x1080")

def main():
    parser = argparse.ArgumentParser(description=f"{APP_NAME} – камера в круге/квадрате + виртуальная камера + калькулятор")
    parser.add_argument("--camera", type=int, default=0, help="Индекс веб-камеры (по умолчанию 0)")
    parser.add_argument("--fps", type=int, default=30, help="Кадров в секунду (по умолчанию 30)")
    parser.add_argument("--width", type=int, default=0, help="Ширина запроса к камере (0 = по умолчанию)")
    parser.add_argument("--height", type=int, default=0, help="Высота запроса к камере (0 = по умолчанию)")
    parser.add_argument("--vcam_res", type=parse_size, default=(1920, 1080), help="Разрешение виртуальной камеры WxH")
    parser.add_argument("--start-hidden", action="store_true", help="Запуск свёрнутым в трей")
    parser.add_argument("--debug", action="store_true", help="Подробные логи и вывод в консоль")
    args = parser.parse_args()

    _ = DATA_DIR
    ensure_first_run_files()
    setup_logger(debug=args.debug)

    # В режиме обычного запуска прячем консоль. Для --debug оставляем, чтобы видеть поток логов.
    if not args.debug:
        hide_console_window()

    _logger.info("CLI args: camera=%s fps=%s size=%sx%s vcam=%sx%s start_hidden=%s debug=%s",
                 args.camera, args.fps, args.width, args.height, args.vcam_res[0], args.vcam_res[1], args.start_hidden, args.debug)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    w = RoundCamWindow(args)
    if args.start_hidden:
        # Стартуем в трее без кнопки в панели задач
        w.tray_act_show.setChecked(False)  # вызовет _on_tray_show_toggled(False)
    else:
        # Показать окно (Qt.Tool — без кнопки в панели задач)
        w.tray_act_show.setChecked(True)
        w.show(); w.raise_(); w.activateWindow()

    # Слушатель NumLock → сворачивание/разворачивание калькулятора
    threading.Thread(target=numlock_listener, daemon=True, name="NumLockListener").start()

    rc = app.exec_()
    _logger.info("QApplication завершён кодом %s", rc)
    sys.exit(rc)

if __name__ == "__main__":
    main()
