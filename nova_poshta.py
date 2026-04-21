#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Звірка накладних Нової Пошти
Версія: 1.0
Залежності: pip install PyQt5 PyMuPDF
"""

import sys
import os
import re
import csv
from collections import OrderedDict
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QGroupBox, QLineEdit, QSpinBox, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QTabWidget, QDialog, QPlainTextEdit, QFileDialog, QMessageBox,
    QFrame, QScrollArea, QSizePolicy
)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor, QIcon

# ── Звук (тільки Windows) ────────────────────────────────────────────────────
try:
    import winsound
    def beep_ok():  winsound.Beep(1050, 70)
    def beep_dup(): winsound.Beep(400, 280)
    def beep_err(): winsound.Beep(300, 420)
except ImportError:
    def beep_ok(): pass
    def beep_dup(): pass
    def beep_err(): pass

# ── PDF підтримка ────────────────────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ═══════════════════════════════════════════════════════════════════════════════
#  Утилітарні функції
# ═══════════════════════════════════════════════════════════════════════════════

def norm(s: str) -> str:
    """Лише цифри."""
    return re.sub(r'[^\d]', '', str(s))

def fmt_en(en: str) -> str:
    """Форматування ЕН для відображення."""
    if len(en) == 14:
        return f"{en[:2]} {en[2:8]} {en[8:]}"
    if len(en) == 18:
        return f"{en[:2]} {en[2:8]} {en[8:14]} · {en[14:]}"
    return en

def expand(base: str, places: int) -> list:
    """
    Розкладає базову ЕН у список усіх місць.
    Повертає [(en, label), ...]
    """
    result = [(base, f"мамка ({places}м)" if places > 1 else "")]
    for i in range(1, places):
        child = base + str(i).zfill(4)
        result.append((child, f"місце {i+1}/{places}"))
    return result

def parse_lines(lines: list, target: str, reg: OrderedDict, fact: OrderedDict) -> int:
    """
    Парсить список рядків і додає до reg або fact.
    Повертає кількість доданих записів.
    """
    added = 0
    for line in lines:
        line = line.strip()
        if not line:
            continue
        parts = re.split(r'[\t,; ]+', line)
        base = norm(parts[0])
        if len(base) < 8:
            continue
        if target == 'reg':
            places = 1
            if len(parts) > 1:
                m = re.search(r'(\d+)', parts[1])
                if m:
                    places = max(1, int(m.group(1)))
            if base not in reg:
                reg[base] = places
                added += 1
        else:
            if base not in fact:
                fact[base] = True
                added += 1
    return added

# ═══════════════════════════════════════════════════════════════════════════════
#  Стилі
# ═══════════════════════════════════════════════════════════════════════════════

# ── Western Bid кольори ──────────────────────────────────────────────────────
WB_DARK    = "#0E3D30"   # темно-зелений — хедер, головні кнопки
WB_MID     = "#1A5C45"   # середній зелений — акценти
WB_MINT    = "#3DD9B3"   # м'ятний — яскравий акцент
WB_MINT_L  = "#E8FAF5"   # світло-м'ятний — фон бейджів
WB_MINT_M  = "#B2EEE0"   # середній м'ятний — бордери
WB_BG      = "#F2F1EF"   # основний фон (теплий кремовий)
WB_CARD    = "#FFFFFF"   # фон карток
WB_BORDER  = "#E0DDD8"   # бордер карток
WB_TEXT    = "#1A1A1A"   # основний текст
WB_TEXT2   = "#5A6B66"   # другорядний текст

APP_STYLE = f"""
* {{
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
    color: {WB_TEXT};
}}

QMainWindow {{ background: {WB_BG}; }}

/* ── Панелі ── */
QGroupBox {{
    background: {WB_CARD};
    border: 1px solid {WB_BORDER};
    border-top: 3px solid {WB_MINT};
    border-radius: 8px;
    margin-top: 12px;
    padding: 10px 8px 8px 8px;
    font-weight: 700;
    font-size: 13px;
    color: {WB_DARK};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: {WB_DARK};
    font-weight: 700;
}}

/* ── Поля вводу ── */
QLineEdit, QSpinBox {{
    border: 1px solid {WB_BORDER};
    border-radius: 6px;
    padding: 6px 10px;
    background: {WB_CARD};
    min-height: 22px;
    color: {WB_TEXT};
}}
QLineEdit:focus, QSpinBox:focus {{
    border-color: {WB_MINT};
    background: {WB_MINT_L};
}}
QLineEdit::placeholder {{ color: #AAAAAA; }}
QSpinBox::up-button, QSpinBox::down-button {{
    width: 16px;
    border: none;
    background: transparent;
}}

/* ── Кнопки — базові ── */
QPushButton {{
    border: 1px solid {WB_BORDER};
    border-radius: 6px;
    padding: 5px 14px;
    background: {WB_CARD};
    color: {WB_TEXT};
    min-height: 28px;
    font-weight: 500;
}}
QPushButton:hover   {{ background: {WB_MINT_L}; border-color: {WB_MINT_M}; }}
QPushButton:pressed {{ background: {WB_MINT_M}; }}
QPushButton:disabled {{ color: #BBBBBB; background: #F5F5F5; }}

/* Основна — темно-зелена */
QPushButton#btn_primary {{
    background: {WB_DARK};
    border-color: {WB_DARK};
    color: #FFFFFF;
    font-weight: 700;
}}
QPushButton#btn_primary:hover  {{ background: {WB_MID}; border-color: {WB_MID}; }}
QPushButton#btn_primary:pressed {{ background: #0A2E22; }}

/* Успіх — м'ятна */
QPushButton#btn_success {{
    background: {WB_MINT};
    border-color: {WB_MINT};
    color: {WB_DARK};
    font-weight: 700;
}}
QPushButton#btn_success:hover  {{ background: #5EE4C0; border-color: #5EE4C0; }}
QPushButton#btn_success:pressed {{ background: #2EC9A0; }}

/* Небезпека */
QPushButton#btn_danger {{
    background: #FFF0F0;
    border-color: #FFBBBB;
    color: #8B1F1F;
}}
QPushButton#btn_danger:hover {{ background: #FFDADA; }}

/* Головна кнопка Звірити */
QPushButton#btn_reconcile {{
    background: {WB_DARK};
    border: none;
    color: {WB_MINT};
    font-weight: 700;
    font-size: 15px;
    padding: 12px 60px;
    border-radius: 8px;
    letter-spacing: 0.03em;
}}
QPushButton#btn_reconcile:hover  {{ background: {WB_MID}; }}
QPushButton#btn_reconcile:pressed {{ background: #0A2E22; }}

/* ── Таблиці ── */
QTableWidget {{
    border: 1px solid {WB_BORDER};
    border-radius: 6px;
    background: {WB_CARD};
    gridline-color: #F0EDEA;
    selection-background-color: {WB_MINT_L};
    selection-color: {WB_DARK};
    alternate-background-color: #FAFAF8;
}}
QTableWidget::item {{ padding: 4px 6px; }}
QTableWidget::item:selected {{
    background: {WB_MINT_L};
    color: {WB_DARK};
}}

QHeaderView::section {{
    background: {WB_DARK};
    border: none;
    border-right: 1px solid {WB_MID};
    padding: 6px 8px;
    font-weight: 700;
    color: {WB_MINT_M};
    font-size: 11px;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}}
QHeaderView::section:first {{ border-radius: 6px 0 0 0; }}
QHeaderView::section:last  {{ border-right: none; border-radius: 0 6px 0 0; }}

/* ── Вкладки результатів ── */
QTabWidget::pane {{
    border: 1px solid {WB_BORDER};
    border-radius: 0 8px 8px 8px;
    background: {WB_CARD};
    margin-top: -1px;
}}
QTabBar::tab {{
    padding: 8px 20px;
    border: 1px solid {WB_BORDER};
    border-bottom: none;
    border-radius: 7px 7px 0 0;
    background: #E8E6E2;
    color: {WB_TEXT2};
    margin-right: 3px;
    font-size: 12px;
    font-weight: 500;
}}
QTabBar::tab:selected {{
    background: {WB_CARD};
    color: {WB_DARK};
    font-weight: 700;
    border-bottom: 1px solid {WB_CARD};
}}
QTabBar::tab:hover {{ background: {WB_MINT_L}; color: {WB_DARK}; }}

/* ── Скролбар ── */
QScrollBar:vertical {{
    width: 7px;
    background: transparent;
    border-radius: 4px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: {WB_MINT_M};
    border-radius: 4px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {WB_MINT}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}

/* ── Текстове поле (вставка списку) ── */
QPlainTextEdit {{
    border: 1px solid {WB_BORDER};
    border-radius: 6px;
    background: {WB_CARD};
    padding: 4px;
    color: {WB_TEXT};
}}
QPlainTextEdit:focus {{ border-color: {WB_MINT}; background: {WB_MINT_L}; }}

/* ── Діалоги ── */
QDialog {{ background: {WB_BG}; }}
QMessageBox {{ background: {WB_BG}; }}

/* ── Бейджі ── */
QLabel#badge_reg {{
    color: {WB_DARK};
    background: {WB_MINT_L};
    border: 1px solid {WB_MINT_M};
    border-radius: 10px;
    padding: 2px 10px;
    font-size: 11px;
    font-weight: 700;
}}
QLabel#badge_fact {{
    color: {WB_DARK};
    background: {WB_MINT_L};
    border: 1px solid {WB_MINT_M};
    border-radius: 10px;
    padding: 2px 10px;
    font-size: 11px;
    font-weight: 700;
}}

/* ── Метрики ── */
QFrame#metric_box {{
    background: {WB_CARD};
    border: 1px solid {WB_BORDER};
    border-bottom: 3px solid {WB_MINT};
    border-radius: 8px;
}}

/* ── Підказки ── */
QLabel#hint_label {{
    color: #9AADA8;
    font-size: 11px;
}}

/* ── Сплітер ── */
QSplitter::handle {{
    background: {WB_BORDER};
    width: 1px;
}}
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  Діалог "Вставити список"
# ═══════════════════════════════════════════════════════════════════════════════

class PasteDialog(QDialog):
    def __init__(self, parent, target: str):
        super().__init__(parent)
        self.setWindowTitle("Вставити список ЕН")
        self.setMinimumWidth(400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        lay = QVBoxLayout(self)
        lay.setSpacing(10)
        lay.setContentsMargins(16, 16, 16, 16)

        if target == 'reg':
            title = "Реєстр: ЕН + кількість місць"
            hint = ("По одному ЕН на рядку.\n"
                    "Місця вказувати через пробіл, Tab або кому:\n"
                    "  59001630827897 3\n"
                    "  20451419750910     ← 1 місце за замовчуванням")
            placeholder = "20451419750910\n20451418532444\n59001630827897 3\n59001630990706 2\n..."
        else:
            title = "Посилки: список штрих-кодів"
            hint = "По одному ЕН/штрих-коду на рядку."
            placeholder = "590016308278970001\n590016308278970002\n20451419750910\n..."

        lbl_title = QLabel(f"<b>{title}</b>")
        lbl_title.setStyleSheet("font-size:14px;")
        lay.addWidget(lbl_title)

        lbl_hint = QLabel(hint)
        lbl_hint.setStyleSheet("color:#666;font-size:12px;")
        lbl_hint.setWordWrap(True)
        lay.addWidget(lbl_hint)

        self.ta = QPlainTextEdit()
        self.ta.setFont(QFont("Consolas", 11))
        self.ta.setPlaceholderText(placeholder)
        self.ta.setMinimumHeight(200)
        lay.addWidget(self.ta)

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Скасувати")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Додати")
        btn_ok.setObjectName("btn_primary")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)
        lay.addLayout(btns)

    def get_lines(self) -> list:
        return self.ta.toPlainText().strip().splitlines()

# ═══════════════════════════════════════════════════════════════════════════════
#  Головне вікно
# ═══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Звірка накладних Нової Пошти")
        self.setMinimumSize(1060, 720)

        # ── Дані ──────────────────────────────────────────────────────────────
        self.reg: OrderedDict  = OrderedDict()   # base_en -> places (int)
        self.fact: OrderedDict = OrderedDict()   # en -> True

        self.last_missing: list = []
        self.last_extra:   list = []

        self._building = False   # прапор, щоб ігнорувати сигнали під час побудови

        self.setup_ui()
        self.setStyleSheet(APP_STYLE)

    # ──────────────────────────────────────────────────────────────────────────
    #  Побудова UI
    # ──────────────────────────────────────────────────────────────────────────

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Темний хедер у стилі Western Bid ──────────────────────────────────
        header = QWidget()
        header.setFixedHeight(56)
        header.setStyleSheet(f"background: {WB_DARK};")
        h_lay = QHBoxLayout(header)
        h_lay.setContentsMargins(18, 0, 18, 0)
        h_lay.setSpacing(10)

        logo_box = QWidget()
        logo_box.setFixedSize(28, 28)
        logo_box.setStyleSheet(f"background: {WB_MINT}; border-radius: 5px;")
        logo_lbl = QLabel("WB")
        logo_lbl.setAlignment(Qt.AlignCenter)
        logo_lbl.setStyleSheet(f"color: {WB_DARK}; font-weight: 900; font-size: 11px; background: transparent;")
        ll = QVBoxLayout(logo_box)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.addWidget(logo_lbl)

        app_name = QLabel("Звірка накладних")
        app_name.setStyleSheet("color: #FFFFFF; font-size: 15px; font-weight: 700; letter-spacing: 0.02em;")
        company = QLabel("Western Bid")
        company.setStyleSheet(f"color: {WB_MINT}; font-size: 11px; font-weight: 600; letter-spacing: 0.06em;")
        names = QVBoxLayout()
        names.setSpacing(1)
        names.addWidget(app_name)
        names.addWidget(company)

        h_lay.addWidget(logo_box)
        h_lay.addLayout(names)
        h_lay.addStretch()
        version = QLabel("v2.0")
        version.setStyleSheet("color: #5A9E89; font-size: 11px;")
        h_lay.addWidget(version)
        outer.addWidget(header)

        accent_line = QFrame()
        accent_line.setFixedHeight(3)
        accent_line.setStyleSheet(f"background: {WB_MINT};")
        outer.addWidget(accent_line)

        # ── Основний контент ───────────────────────────────────────────────────
        content = QWidget()
        root = QVBoxLayout(content)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)
        outer.addWidget(content, stretch=1)

        # Два панелі
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(8)
        splitter.addWidget(self._build_reg_panel())
        splitter.addWidget(self._build_fact_panel())
        splitter.setSizes([530, 530])
        root.addWidget(splitter, stretch=1)

        # Кнопка Звірити
        btn_rec = QPushButton("⚖   Звірити")
        btn_rec.setObjectName("btn_reconcile")
        btn_rec.clicked.connect(self.reconcile)
        row_rec = QHBoxLayout()
        row_rec.addStretch()
        row_rec.addWidget(btn_rec)
        row_rec.addStretch()
        root.addLayout(row_rec)

        # Блок результатів
        self.res_widget = QWidget()
        self.res_widget.setVisible(False)
        root.addWidget(self.res_widget)
        self._build_result_panel()

    # ── Панель реєстру ────────────────────────────────────────────────────────

    def _build_reg_panel(self) -> QGroupBox:
        grp = QGroupBox("Реєстр (закриті посилки)")
        lay = QVBoxLayout(grp)
        lay.setSpacing(6)

        self.reg_badge = QLabel("0 ЕН · 0 місць")
        self.reg_badge.setObjectName("badge_reg")
        self.reg_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        lay.addWidget(self.reg_badge)

        # Рядок введення
        row1 = QHBoxLayout()
        self.reg_en_input = QLineEdit()
        self.reg_en_input.setPlaceholderText("Скануйте або введіть ЕН...")
        self.reg_en_input.returnPressed.connect(self._reg_enter)
        row1.addWidget(self.reg_en_input)

        row1.addWidget(QLabel("місць:"))
        self.reg_places = QSpinBox()
        self.reg_places.setRange(1, 999)
        self.reg_places.setValue(1)
        self.reg_places.setFixedWidth(64)
        self.reg_places.lineEdit().returnPressed.connect(self.add_to_reg)
        row1.addWidget(self.reg_places)

        btn_add = QPushButton("+")
        btn_add.setObjectName("btn_primary")
        btn_add.setFixedWidth(36)
        btn_add.setStyleSheet(btn_add.styleSheet())
        btn_add.clicked.connect(self.add_to_reg)
        row1.addWidget(btn_add)
        lay.addLayout(row1)

        hint = QLabel("Enter → перехід до «місць» → Enter додає")
        hint.setObjectName("hint_label")
        lay.addWidget(hint)

        # Інструменти
        row2 = QHBoxLayout()
        row2.setSpacing(5)
        for label, slot in [
            ("📂 Файл",          lambda: self.load_file('reg')),
            ("📋 Вставити список",  lambda: self.paste_list('reg')),
            ("📄 PDF",            self.load_pdf),
        ]:
            btn = QPushButton(label)
            if label == "📄 PDF" and not PDF_OK:
                btn.setEnabled(False)
                btn.setToolTip("pip install PyMuPDF")
            btn.clicked.connect(slot)
            row2.addWidget(btn)

        btn_clear = QPushButton("🗑 Очистити")
        btn_clear.setObjectName("btn_danger")
        btn_clear.clicked.connect(self.clear_reg)
        row2.addWidget(btn_clear)
        lay.addLayout(row2)

        # Пошук
        self.reg_search = QLineEdit()
        self.reg_search.setPlaceholderText("🔍 Пошук по 4 цифрах або повному ЕН...")
        self.reg_search.textChanged.connect(self.render_reg)
        lay.addWidget(self.reg_search)

        # Таблиця реєстру
        self.reg_table = QTableWidget()
        self.reg_table.setColumnCount(4)
        self.reg_table.setHorizontalHeaderLabels(["№", "ЕН", "Місць", ""])
        h = self.reg_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Fixed)
        h.setSectionResizeMode(1, QHeaderView.Stretch)
        h.setSectionResizeMode(2, QHeaderView.Fixed)
        h.setSectionResizeMode(3, QHeaderView.Fixed)
        self.reg_table.setColumnWidth(0, 36)
        self.reg_table.setColumnWidth(2, 68)
        self.reg_table.setColumnWidth(3, 30)
        self.reg_table.verticalHeader().setVisible(False)
        self.reg_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.reg_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.reg_table.setAlternatingRowColors(True)
        lay.addWidget(self.reg_table)

        return grp

    # ── Панель сканування ─────────────────────────────────────────────────────

    def _build_fact_panel(self) -> QGroupBox:
        grp = QGroupBox("Фактичне сканування")
        lay = QVBoxLayout(grp)
        lay.setSpacing(6)

        self.fact_badge = QLabel("0 штрих-кодів")
        self.fact_badge.setObjectName("badge_fact")
        self.fact_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        lay.addWidget(self.fact_badge)

        row1 = QHBoxLayout()
        self.fact_en_input = QLineEdit()
        self.fact_en_input.setPlaceholderText("Скануйте штрих-код...")
        self.fact_en_input.returnPressed.connect(self.add_to_fact)
        row1.addWidget(self.fact_en_input)

        btn_add = QPushButton("+")
        btn_add.setObjectName("btn_success")
        btn_add.setFixedWidth(36)
        btn_add.clicked.connect(self.add_to_fact)
        row1.addWidget(btn_add)
        lay.addLayout(row1)

        hint = QLabel("Автододавання після Enter або сканування")
        hint.setObjectName("hint_label")
        lay.addWidget(hint)

        row2 = QHBoxLayout()
        row2.setSpacing(5)
        for label, slot in [
            ("📂 Файл",          lambda: self.load_file('fact')),
            ("📋 Вставити список",  lambda: self.paste_list('fact')),
        ]:
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            row2.addWidget(btn)

        btn_clear = QPushButton("🗑 Очистити")
        btn_clear.setObjectName("btn_danger")
        btn_clear.clicked.connect(self.clear_fact)
        row2.addWidget(btn_clear)
        lay.addLayout(row2)

        self.fact_search = QLineEdit()
        self.fact_search.setPlaceholderText("🔍 Пошук по 4 цифрах або повному ЕН...")
        self.fact_search.textChanged.connect(self.render_fact)
        lay.addWidget(self.fact_search)

        self.fact_table = QTableWidget()
        self.fact_table.setColumnCount(3)
        self.fact_table.setHorizontalHeaderLabels(["№", "ЕН", ""])
        h = self.fact_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.Fixed)
        h.setSectionResizeMode(1, QHeaderView.Stretch)
        h.setSectionResizeMode(2, QHeaderView.Fixed)
        self.fact_table.setColumnWidth(0, 36)
        self.fact_table.setColumnWidth(2, 30)
        self.fact_table.verticalHeader().setVisible(False)
        self.fact_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.fact_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.fact_table.setAlternatingRowColors(True)
        lay.addWidget(self.fact_table)

        return grp

    # ── Панель результатів ────────────────────────────────────────────────────

    def _build_result_panel(self):
        lay = QVBoxLayout(self.res_widget)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        # Метрики
        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(8)
        self.met_vals = []
        for label in ["Очікується місць", "Відскановано", "Відсутніх", "Зайвих"]:
            box = QFrame()
            box.setObjectName("metric_box")
            box.setMinimumHeight(70)
            bl = QVBoxLayout(box)
            val = QLabel("0")
            val.setStyleSheet("font-size:26px;font-weight:700;")
            val.setAlignment(Qt.AlignCenter)
            lbl = QLabel(label)
            lbl.setStyleSheet("font-size:11px;color:#888;")
            lbl.setAlignment(Qt.AlignCenter)
            bl.addWidget(val)
            bl.addWidget(lbl)
            metrics_row.addWidget(box)
            self.met_vals.append(val)
        lay.addLayout(metrics_row)

        # Вкладки
        self.tabs = QTabWidget()
        for title, attr in [
            ("Відсутні (0)", "tbl_ms"),
            ("Зайві (0)",    "tbl_ex"),
            ("Всі (0)",      "tbl_all"),
        ]:
            tab_widget = QWidget()
            tl = QVBoxLayout(tab_widget)
            tl.setContentsMargins(6, 6, 6, 6)
            tl.setSpacing(6)

            # Рядок дій
            arow = QHBoxLayout()
            btn_copy   = QPushButton("📋 Скопіювати ЕН")
            btn_export = QPushButton("💾 Експорт CSV")
            btn_copy.clicked.connect(lambda _, a=attr: self._copy_results(a))
            btn_export.clicked.connect(lambda _, a=attr: self._export_csv(a))
            arow.addWidget(btn_copy)
            arow.addWidget(btn_export)
            arow.addStretch()
            tl.addLayout(arow)

            tbl = QTableWidget()
            tbl.setColumnCount(4)
            tbl.setHorizontalHeaderLabels(["№", "ЕН", "Примітка", "Статус"])
            h = tbl.horizontalHeader()
            h.setSectionResizeMode(0, QHeaderView.Fixed)
            h.setSectionResizeMode(1, QHeaderView.Stretch)
            h.setSectionResizeMode(2, QHeaderView.ResizeToContents)
            h.setSectionResizeMode(3, QHeaderView.Fixed)
            tbl.setColumnWidth(0, 36)
            tbl.setColumnWidth(3, 90)
            tbl.verticalHeader().setVisible(False)
            tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
            tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
            tbl.setAlternatingRowColors(True)
            tl.addWidget(tbl)
            setattr(self, attr, tbl)

            self.tabs.addTab(tab_widget, title)

        lay.addWidget(self.tabs)

    # ──────────────────────────────────────────────────────────────────────────
    #  Дії — Реєстр
    # ──────────────────────────────────────────────────────────────────────────

    def _reg_enter(self):
        """Enter у полі ЕН → переходимо до поля «місць»."""
        self.reg_places.setFocus()
        self.reg_places.selectAll()

    def add_to_reg(self):
        base = norm(self.reg_en_input.text())
        places = self.reg_places.value()
        if len(base) < 8:
            self.reg_en_input.setFocus()
            return
        if base not in self.reg:
            self.reg[base] = places
            beep_ok()
        else:
            beep_dup()
        self.reg_en_input.clear()
        self.reg_places.setValue(1)
        self.render_reg()
        self.reg_en_input.setFocus()

    def render_reg(self):
        q = norm(self.reg_search.text())
        total_places = sum(self.reg.values())
        self.reg_badge.setText(f"{len(self.reg)} ЕН · {total_places} місць")

        entries = list(self.reg.items())
        if q:
            entries = [(en, p) for en, p in entries
                       if en.endswith(q) or q in en]

        self._building = True
        self.reg_table.setRowCount(0)
        for idx, (en, places) in enumerate(entries):
            row = self.reg_table.rowCount()
            self.reg_table.insertRow(row)

            # №
            n_item = QTableWidgetItem(str(idx + 1))
            n_item.setForeground(QColor("#AAAAAA"))
            n_item.setTextAlignment(Qt.AlignCenter)

            # ЕН
            en_item = QTableWidgetItem(fmt_en(en))
            en_item.setFont(QFont("Consolas", 11))
            en_item.setData(Qt.UserRole, en)

            # Підсвічуємо якщо відскановано
            exp = expand(en, places)
            scanned = sum(1 for child_en, _ in exp if child_en in self.fact)
            if scanned == len(exp):
                en_item.setForeground(QColor(WB_MINT))       # всі відскановані
            elif scanned > 0:
                en_item.setForeground(QColor("#E8A020"))     # частково

            # Spin для кількості місць
            spin = QSpinBox()
            spin.setRange(1, 999)
            spin.setValue(places)
            spin.setFrame(False)
            spin.setStyleSheet(
                "QSpinBox{border:none;background:transparent;font-size:12px;}"
                "QSpinBox::up-button,QSpinBox::down-button{width:14px;}"
            )
            spin.valueChanged.connect(lambda val, e=en: self._set_places(e, val))

            # Кнопка видалення
            del_btn = QPushButton("✕")
            del_btn.setStyleSheet(
                "QPushButton{border:none;color:#CCCCCC;font-size:11px;"
                "background:transparent;padding:0;}"
                "QPushButton:hover{color:#791F1F;background:#FCEBEB;}"
            )
            del_btn.setFixedWidth(28)
            del_btn.setCursor(Qt.PointingHandCursor)
            del_btn.clicked.connect(lambda _, e=en: self._del_reg(e))

            self.reg_table.setItem(row, 0, n_item)
            self.reg_table.setItem(row, 1, en_item)
            self.reg_table.setCellWidget(row, 2, spin)
            self.reg_table.setCellWidget(row, 3, del_btn)
            self.reg_table.setRowHeight(row, 30)

        self._building = False

    def _set_places(self, en: str, val: int):
        if self._building:
            return
        if en in self.reg:
            self.reg[en] = max(1, val)
            self.reg_badge.setText(
                f"{len(self.reg)} ЕН · {sum(self.reg.values())} місць"
            )

    def _del_reg(self, en: str):
        self.reg.pop(en, None)
        self.render_reg()

    def clear_reg(self):
        if self.reg and QMessageBox.question(
            self, "Підтвердження", "Очистити весь реєстр?",
            QMessageBox.Yes | QMessageBox.No
        ) != QMessageBox.Yes:
            return
        self.reg.clear()
        self.render_reg()

    # ──────────────────────────────────────────────────────────────────────────
    #  Дії — Сканування
    # ──────────────────────────────────────────────────────────────────────────

    def add_to_fact(self):
        en = norm(self.fact_en_input.text())
        if len(en) < 8:
            return
        if en not in self.fact:
            self.fact[en] = True
            beep_ok()
        else:
            beep_dup()
        self.fact_en_input.clear()
        self.render_fact()
        self.fact_en_input.setFocus()

    def render_fact(self):
        q = norm(self.fact_search.text())
        self.fact_badge.setText(f"{len(self.fact)} штрих-кодів")

        entries = list(self.fact.keys())
        if q:
            entries = [en for en in entries if en.endswith(q) or q in en]

        # Будуємо розгорнутий реєстр для перевірки
        expanded = self._expand_registry()

        self.fact_table.setRowCount(0)
        for idx, en in enumerate(entries):
            row = self.fact_table.rowCount()
            self.fact_table.insertRow(row)

            n_item = QTableWidgetItem(str(idx + 1))
            n_item.setForeground(QColor("#AAAAAA"))
            n_item.setTextAlignment(Qt.AlignCenter)

            en_item = QTableWidgetItem(fmt_en(en))
            en_item.setFont(QFont("Consolas", 11))
            en_item.setData(Qt.UserRole, en)

            # Помаранчевий якщо не в реєстрі
            if en not in expanded:
                en_item.setForeground(QColor("#CC7700"))

            del_btn = QPushButton("✕")
            del_btn.setStyleSheet(
                "QPushButton{border:none;color:#CCCCCC;font-size:11px;"
                "background:transparent;padding:0;}"
                "QPushButton:hover{color:#791F1F;background:#FCEBEB;}"
            )
            del_btn.setFixedWidth(28)
            del_btn.setCursor(Qt.PointingHandCursor)
            del_btn.clicked.connect(lambda _, e=en: self._del_fact(e))

            self.fact_table.setItem(row, 0, n_item)
            self.fact_table.setItem(row, 1, en_item)
            self.fact_table.setCellWidget(row, 2, del_btn)
            self.fact_table.setRowHeight(row, 30)

    def _del_fact(self, en: str):
        self.fact.pop(en, None)
        self.render_fact()

    def clear_fact(self):
        if self.fact and QMessageBox.question(
            self, "Підтвердження", "Очистити список сканування?",
            QMessageBox.Yes | QMessageBox.No
        ) != QMessageBox.Yes:
            return
        self.fact.clear()
        self.render_fact()

    # ──────────────────────────────────────────────────────────────────────────
    #  Завантаження файлів
    # ──────────────────────────────────────────────────────────────────────────

    def load_file(self, target: str):
        path, _ = QFileDialog.getOpenFileName(
            self, "Відкрити файл", "",
            "Текстові файли (*.txt *.csv);;Всі файли (*)"
        )
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8-sig', errors='ignore') as f:
                lines = f.read().splitlines()
            added = parse_lines(lines, target, self.reg, self.fact)
            self.render_reg() if target == 'reg' else self.render_fact()
            QMessageBox.information(self, "Завантажено", f"Додано: {added} записів")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", str(e))

    def paste_list(self, target: str):
        dlg = PasteDialog(self, target)
        if dlg.exec_() == QDialog.Accepted:
            added = parse_lines(dlg.get_lines(), target, self.reg, self.fact)
            self.render_reg() if target == 'reg' else self.render_fact()
            if added:
                QMessageBox.information(self, "Додано", f"Додано: {added} записів")

    def load_pdf(self):
        if not PDF_OK:
            QMessageBox.warning(self, "PDF", "Встановіть PyMuPDF:\npip install PyMuPDF")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Відкрити PDF реєстр", "", "PDF файли (*.pdf)"
        )
        if not path:
            return
        try:
            doc = fitz.open(path)
            added = 0
            for page in doc:
                # Нормалізуємо текст: замінюємо переноси рядків на пробіли
                # щоб рядок таблиці став одним рядком тексту
                text = ' '.join(page.get_text().split())

                # Формат рядка реєстру НП:
                # [№] [ЕН14] [дата] [ім'я] [тел.відправника] [одержувач]
                # [тел.одержувача] [к-сть місць] [вага] [опис] ...
                #
                # Стратегія: знаходимо ЕН, потім у наступних ~500 символах
                # шукаємо: телефон одержувача → к-сть місць → вага (десяткова)

                for m in re.finditer(r'\b(\d{14})\b', text):
                    en = m.group(1)
                    # Беремо фрагмент тексту після ЕН (до 500 символів)
                    chunk = text[m.start(): m.start() + 500]

                    # Шукаємо: +380XXXXXXXXX  [1-3 цифри]  [число з комою/крапкою]
                    # Тобто: телефон одержувача → місця → вага
                    pm = re.search(
                        r'\+380\d{9}\s+(\d{1,3})\s+\d+[.,]\d+',
                        chunk
                    )
                    if pm:
                        places = max(1, int(pm.group(1)))
                    else:
                        places = 1

                    if en not in self.reg:
                        self.reg[en] = places
                        added += 1

            doc.close()
            self.render_reg()
            QMessageBox.information(
                self, "PDF імпорт",
                f"Знайдено та додано: {added} ЕН\n\n"
                f"Перевірте кількість місць у списку — "
                f"для багатомісних посилок значення підсвічено синім."
            )
        except Exception as e:
            QMessageBox.critical(self, "Помилка PDF", str(e))

    # ──────────────────────────────────────────────────────────────────────────
    #  Звірка
    # ──────────────────────────────────────────────────────────────────────────

    def _expand_registry(self) -> dict:
        """Повертає словник: en -> {'base', 'label'}"""
        expanded = {}
        for base, places in self.reg.items():
            for en, label in expand(base, places):
                expanded[en] = {'base': base, 'label': label}
        return expanded

    def reconcile(self):
        expanded = self._expand_registry()
        missing = []
        extra   = []

        for en, info in expanded.items():
            if en not in self.fact:
                missing.append({'en': en, 'label': info['label']})

        for en in self.fact:
            if en not in expanded:
                extra.append({'en': en, 'label': ''})

        self.last_missing = missing
        self.last_extra   = extra

        # Метрики
        counts = [len(expanded), len(self.fact), len(missing), len(extra)]
        colors = [WB_DARK, WB_MID,
                  "#C0392B" if missing else WB_MINT,
                  "#E8A020" if extra   else WB_MINT]
        for lbl, val, color in zip(self.met_vals, counts, colors):
            lbl.setText(str(val))
            lbl.setStyleSheet(f"font-size:26px;font-weight:700;color:{color};")

        # Заголовки вкладок
        self.tabs.setTabText(0, f"Відсутні ({len(missing)})")
        self.tabs.setTabText(1, f"Зайві ({len(extra)})")
        self.tabs.setTabText(2, f"Всі ({len(missing) + len(extra)})")

        # Заповнення таблиць
        self._fill_result_table(
            self.tbl_ms,
            [{'en': i['en'], 'label': i['label'], 'type': 'missing'} for i in missing]
        )
        self._fill_result_table(
            self.tbl_ex,
            [{'en': i['en'], 'label': '', 'type': 'extra'} for i in extra]
        )
        self._fill_result_table(
            self.tbl_all,
            [{'en': i['en'], 'label': i['label'], 'type': 'missing'} for i in missing] +
            [{'en': i['en'], 'label': '', 'type': 'extra'} for i in extra]
        )

        # Оновлюємо підсвічення реєстру
        self.render_reg()

        self.res_widget.setVisible(True)
        beep_ok() if not missing and not extra else beep_err()

    def _fill_result_table(self, tbl: QTableWidget, items: list):
        tbl.setRowCount(0)
        for idx, item in enumerate(items):
            row = tbl.rowCount()
            tbl.insertRow(row)

            n = QTableWidgetItem(str(idx + 1))
            n.setForeground(QColor("#AAAAAA"))
            n.setTextAlignment(Qt.AlignCenter)

            en_item = QTableWidgetItem(fmt_en(item['en']))
            en_item.setFont(QFont("Consolas", 11))
            en_item.setData(Qt.UserRole, item['en'])

            note = QTableWidgetItem(item.get('label', ''))
            note.setForeground(QColor("#888888"))
            note.setFont(QFont("Segoe UI", 11))

            is_missing = item['type'] == 'missing'
            status_text = "Відсутня" if is_missing else "Зайва"
            status = QTableWidgetItem(status_text)
            status.setTextAlignment(Qt.AlignCenter)
            if is_missing:
                status.setForeground(QColor("#8B1F1F"))
                status.setBackground(QColor("#FFF0F0"))
            else:
                status.setForeground(QColor(WB_DARK))
                status.setBackground(QColor(WB_MINT_L))

            tbl.setItem(row, 0, n)
            tbl.setItem(row, 1, en_item)
            tbl.setItem(row, 2, note)
            tbl.setItem(row, 3, status)
            tbl.setRowHeight(row, 28)

    # ──────────────────────────────────────────────────────────────────────────
    #  Копіювання та Експорт
    # ──────────────────────────────────────────────────────────────────────────

    def _copy_results(self, attr: str):
        tbl: QTableWidget = getattr(self, attr)
        lines = []
        for row in range(tbl.rowCount()):
            item = tbl.item(row, 1)
            if item:
                raw = item.data(Qt.UserRole) or item.text()
                lines.append(raw)
        QApplication.clipboard().setText('\n'.join(lines))
        QMessageBox.information(self, "Скопійовано",
                                f"{len(lines)} ЕН скопійовано в буфер обміну")

    def _export_csv(self, attr: str):
        tbl: QTableWidget = getattr(self, attr)
        path, _ = QFileDialog.getSaveFileName(
            self, "Зберегти CSV", "result.csv", "CSV файли (*.csv)"
        )
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["№", "ЕН", "Примітка", "Статус"])
                for row in range(tbl.rowCount()):
                    num    = tbl.item(row, 0).text() if tbl.item(row, 0) else ''
                    en     = tbl.item(row, 1).data(Qt.UserRole) if tbl.item(row, 1) else ''
                    note   = tbl.item(row, 2).text() if tbl.item(row, 2) else ''
                    status = tbl.item(row, 3).text() if tbl.item(row, 3) else ''
                    writer.writerow([num, en, note, status])
            QMessageBox.information(self, "Збережено", f"Файл збережено:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", str(e))


# ═══════════════════════════════════════════════════════════════════════════════
#  Точка входу
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
