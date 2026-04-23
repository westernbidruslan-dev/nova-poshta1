#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Звірка накладних Нової Пошти  —  Western Bid
Версія: 3.0
pip install PyQt5 PyMuPDF
"""

import sys, os, re, csv, sqlite3, platform
from collections import OrderedDict
from datetime import date, datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QGroupBox, QLineEdit, QSpinBox, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QTabWidget, QDialog, QPlainTextEdit, QFileDialog, QMessageBox,
    QFrame, QSizePolicy, QDateEdit, QCheckBox
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor

# ── Звук ─────────────────────────────────────────────────────────────────────
try:
    import winsound
    def beep_ok():  winsound.Beep(1050, 70)
    def beep_dup(): winsound.Beep(400,  280)
    def beep_err(): winsound.Beep(300,  420)
except ImportError:
    def beep_ok():  pass
    def beep_dup(): pass
    def beep_err(): pass

# ── PDF ───────────────────────────────────────────────────────────────────────
try:
    import fitz
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ── Шрифт ────────────────────────────────────────────────────────────────────
if platform.system() == "Darwin":
    UI_FONT = "-apple-system, 'Helvetica Neue', Arial, sans-serif"
elif platform.system() == "Windows":
    UI_FONT = "'Segoe UI', Arial, sans-serif"
else:
    UI_FONT = "Arial, sans-serif"

# ═══════════════════════════════════════════════════════════════════════════════
#  Утиліти
# ═══════════════════════════════════════════════════════════════════════════════

def norm(s):
    return re.sub(r'[^\d]', '', str(s))

def fmt_en(en):
    if len(en) == 14: return f"{en[:2]} {en[2:8]} {en[8:]}"
    if len(en) == 18: return f"{en[:2]} {en[2:8]} {en[8:14]}·{en[14:]}"
    return en

def expand(base, places):
    result = [(base, f"мамка ({places}м)" if places > 1 else "")]
    for i in range(1, places):
        result.append((base + str(i).zfill(4), f"місце {i+1}/{places}"))
    return result

def parse_lines(lines, target, reg, fact):
    added = 0
    for line in lines:
        line = line.strip()
        if not line: continue
        parts = re.split(r'[\t,; ]+', line)
        base = norm(parts[0])
        if len(base) < 8: continue
        if target == 'reg':
            places = 1
            if len(parts) > 1:
                m = re.search(r'(\d+)', parts[1])
                if m: places = max(1, int(m.group(1)))
            if base not in reg:
                reg[base] = places; added += 1
        else:
            if base not in fact:
                fact[base] = True; added += 1
    return added

# ═══════════════════════════════════════════════════════════════════════════════
#  База даних SQLite
# ═══════════════════════════════════════════════════════════════════════════════

DB_PATH = os.path.join(os.path.expanduser("~"), ".nova_poshta_history.db")

class DB:
    def __init__(self):
        self.con = sqlite3.connect(DB_PATH)
        self.con.row_factory = sqlite3.Row
        self._init()

    def _init(self):
        self.con.executescript("""
        CREATE TABLE IF NOT EXISTS sessions (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            date    TEXT NOT NULL,
            created TEXT NOT NULL,
            note    TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS reg_items (
            session_id INTEGER, en TEXT, places INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS fact_items (
            session_id INTEGER, en TEXT
        );
        CREATE TABLE IF NOT EXISTS results (
            session_id INTEGER, en TEXT, label TEXT, status TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_fi ON fact_items(en);
        CREATE INDEX IF NOT EXISTS ix_ri ON reg_items(en);
        CREATE INDEX IF NOT EXISTS ix_sd ON sessions(date);
        """)
        self.con.commit()

    def save(self, session_date, reg, fact, missing, extra):
        cur = self.con.cursor()
        cur.execute("INSERT INTO sessions(date,created) VALUES(?,?)",
                    (session_date, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        sid = cur.lastrowid
        cur.executemany("INSERT INTO reg_items VALUES(?,?,?)",
                        [(sid, en, pl) for en, pl in reg.items()])
        cur.executemany("INSERT INTO fact_items VALUES(?,?)",
                        [(sid, en) for en in fact])
        rows = ([(sid, i['en'], i.get('label',''), 'missing') for i in missing] +
                [(sid, i['en'], '', 'extra') for i in extra])
        if rows:
            cur.executemany("INSERT INTO results VALUES(?,?,?,?)", rows)
        self.con.commit()
        return sid

    def find_fact_en(self, en_list, exclude_date=None):
        """Повертає {en: date} — де ЕН було відскановано раніше."""
        result = {}
        for en in en_list:
            if exclude_date:
                row = self.con.execute(
                    "SELECT s.date FROM fact_items fi "
                    "JOIN sessions s ON s.id=fi.session_id "
                    "WHERE fi.en=? AND s.date < ? ORDER BY s.date DESC LIMIT 1",
                    (en, exclude_date)
                ).fetchone()
            else:
                row = self.con.execute(
                    "SELECT s.date FROM fact_items fi "
                    "JOIN sessions s ON s.id=fi.session_id "
                    "WHERE fi.en=? ORDER BY s.date DESC LIMIT 1",
                    (en,)
                ).fetchone()
            if row:
                result[en] = row['date']
        return result

    def sessions(self, date_from=None, date_to=None):
        q = ("SELECT s.id, s.date, s.created, s.note, "
             "COUNT(DISTINCT r.rowid) as issues "
             "FROM sessions s LEFT JOIN results r ON r.session_id=s.id ")
        w, p = [], []
        if date_from: w.append("s.date>=?"); p.append(date_from)
        if date_to:   w.append("s.date<=?"); p.append(date_to)
        if w: q += "WHERE " + " AND ".join(w) + " "
        q += "GROUP BY s.id ORDER BY s.date DESC, s.created DESC"
        return self.con.execute(q, p).fetchall()

    def session_full(self, sid):
        reg   = self.con.execute("SELECT en,places FROM reg_items   WHERE session_id=?", (sid,)).fetchall()
        fact  = self.con.execute("SELECT en          FROM fact_items  WHERE session_id=?", (sid,)).fetchall()
        res   = self.con.execute("SELECT en,label,status FROM results WHERE session_id=?", (sid,)).fetchall()
        info  = self.con.execute("SELECT date,created,note FROM sessions WHERE id=?", (sid,)).fetchone()
        return info, reg, fact, res

    def delete_range(self, date_from, date_to):
        ids = [r[0] for r in self.con.execute(
            "SELECT id FROM sessions WHERE date>=? AND date<=?", (date_from, date_to)
        ).fetchall()]
        if not ids: return 0
        ph = ",".join("?" * len(ids))
        for tbl in ("reg_items", "fact_items", "results"):
            self.con.execute(f"DELETE FROM {tbl} WHERE session_id IN ({ph})", ids)
        self.con.execute(f"DELETE FROM sessions WHERE id IN ({ph})", ids)
        self.con.execute("VACUUM")
        self.con.commit()
        return len(ids)

    def db_size_kb(self):
        return round(os.path.getsize(DB_PATH) / 1024, 1) if os.path.exists(DB_PATH) else 0

# ═══════════════════════════════════════════════════════════════════════════════
#  Стилі Western Bid
# ═══════════════════════════════════════════════════════════════════════════════

WB_DARK   = "#0E3D30"
WB_MID    = "#1A5C45"
WB_MINT   = "#3DD9B3"
WB_MINT_L = "#E8FAF5"
WB_MINT_M = "#B2EEE0"
WB_BG     = "#F2F1EF"
WB_CARD   = "#FFFFFF"
WB_BORDER = "#E0DDD8"
WB_TEXT   = "#1A1A1A"
WB_TEXT2  = "#5A6B66"

APP_STYLE = f"""
* {{ font-family: {UI_FONT}; font-size: 13px; color: {WB_TEXT}; }}
QMainWindow {{ background: {WB_BG}; }}
QDialog     {{ background: {WB_BG}; }}

QGroupBox {{
    background:{WB_CARD}; border:1px solid {WB_BORDER};
    border-top:3px solid {WB_MINT}; border-radius:8px;
    margin-top:12px; padding:10px 8px 8px 8px;
    font-weight:700; font-size:13px; color:{WB_DARK};
}}
QGroupBox::title {{
    subcontrol-origin:margin; subcontrol-position:top left;
    left:12px; padding:0 6px; color:{WB_DARK}; font-weight:700;
}}
QLineEdit, QSpinBox, QDateEdit {{
    border:1px solid {WB_BORDER}; border-radius:6px;
    padding:6px 10px; background:{WB_CARD}; min-height:22px;
}}
QLineEdit:focus, QSpinBox:focus, QDateEdit:focus {{
    border-color:{WB_MINT}; background:{WB_MINT_L};
}}
QSpinBox::up-button, QSpinBox::down-button {{ width:16px; border:none; background:transparent; }}
QDateEdit::drop-down {{ width:20px; border:none; }}

QPushButton {{
    border:1px solid {WB_BORDER}; border-radius:6px;
    padding:5px 14px; background:{WB_CARD}; color:{WB_TEXT};
    min-height:28px; font-weight:500;
}}
QPushButton:hover   {{ background:{WB_MINT_L}; border-color:{WB_MINT_M}; }}
QPushButton:pressed {{ background:{WB_MINT_M}; }}
QPushButton:disabled {{ color:#BBB; background:#F5F5F5; }}
QPushButton#btn_primary {{
    background:{WB_DARK}; border-color:{WB_DARK}; color:#FFF; font-weight:700;
}}
QPushButton#btn_primary:hover  {{ background:{WB_MID}; border-color:{WB_MID}; }}
QPushButton#btn_success {{
    background:{WB_MINT}; border-color:{WB_MINT}; color:{WB_DARK}; font-weight:700;
}}
QPushButton#btn_success:hover {{ background:#5EE4C0; }}
QPushButton#btn_danger {{
    background:#FFF0F0; border-color:#FFBBBB; color:#8B1F1F;
}}
QPushButton#btn_danger:hover {{ background:#FFDADA; }}
QPushButton#btn_reconcile {{
    background:{WB_DARK}; border:none; color:{WB_MINT};
    font-weight:700; font-size:15px; padding:12px 60px;
    border-radius:8px; letter-spacing:0.03em;
}}
QPushButton#btn_reconcile:hover {{ background:{WB_MID}; }}
QPushButton#btn_history {{
    background:rgba(61,217,179,0.18); border:1px solid {WB_MINT_M};
    color:{WB_MINT}; font-weight:600; font-size:12px;
    padding:5px 14px; border-radius:6px;
}}
QPushButton#btn_history:hover {{ background:rgba(61,217,179,0.30); }}

QTableWidget {{
    border:1px solid {WB_BORDER}; border-radius:6px; background:{WB_CARD};
    gridline-color:#F0EDEA; alternate-background-color:#FAFAF8;
    selection-background-color:{WB_MINT_L}; selection-color:{WB_DARK};
}}
QTableWidget::item {{ padding:4px 6px; }}
QTableWidget::item:selected {{ background:{WB_MINT_L}; color:{WB_DARK}; }}
QHeaderView::section {{
    background:{WB_DARK}; border:none; border-right:1px solid {WB_MID};
    padding:6px 8px; font-weight:700; color:{WB_MINT_M};
    font-size:11px; letter-spacing:0.04em;
}}
QHeaderView::section:last {{ border-right:none; }}

QTabWidget::pane {{
    border:1px solid {WB_BORDER}; border-radius:0 8px 8px 8px;
    background:{WB_CARD}; margin-top:-1px;
}}
QTabBar::tab {{
    padding:8px 18px; border:1px solid {WB_BORDER}; border-bottom:none;
    border-radius:7px 7px 0 0; background:#E8E6E2; color:{WB_TEXT2};
    margin-right:3px; font-size:12px; font-weight:500;
}}
QTabBar::tab:selected {{
    background:{WB_CARD}; color:{WB_DARK}; font-weight:700;
    border-bottom:1px solid {WB_CARD};
}}
QTabBar::tab:hover {{ background:{WB_MINT_L}; color:{WB_DARK}; }}

QScrollBar:vertical {{
    width:7px; background:transparent; border-radius:4px; margin:2px;
}}
QScrollBar::handle:vertical {{
    background:{WB_MINT_M}; border-radius:4px; min-height:24px;
}}
QScrollBar::handle:vertical:hover {{ background:{WB_MINT}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height:0; }}

QPlainTextEdit {{
    border:1px solid {WB_BORDER}; border-radius:6px;
    background:{WB_CARD}; padding:4px;
}}
QPlainTextEdit:focus {{ border-color:{WB_MINT}; background:{WB_MINT_L}; }}

QLabel#badge_reg, QLabel#badge_fact {{
    color:{WB_DARK}; background:{WB_MINT_L}; border:1px solid {WB_MINT_M};
    border-radius:10px; padding:2px 10px; font-size:11px; font-weight:700;
}}
QFrame#metric_box {{
    background:{WB_CARD}; border:1px solid {WB_BORDER};
    border-bottom:3px solid {WB_MINT}; border-radius:8px;
}}
QLabel#hint_label {{ color:#9AADA8; font-size:11px; }}
QSplitter::handle {{ background:{WB_BORDER}; width:1px; }}
QCheckBox {{ color:{WB_TEXT2}; font-size:12px; spacing:6px; }}
QCheckBox::indicator {{
    width:16px; height:16px; border:1px solid {WB_BORDER};
    border-radius:4px; background:{WB_CARD};
}}
QCheckBox::indicator:checked {{
    background:{WB_MINT}; border-color:{WB_MINT};
}}
"""

# ═══════════════════════════════════════════════════════════════════════════════
#  Діалог: Вставити список
# ═══════════════════════════════════════════════════════════════════════════════

class PasteDialog(QDialog):
    def __init__(self, parent, target):
        super().__init__(parent)
        self.setWindowTitle("Вставити список ЕН")
        self.setMinimumWidth(400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        lay = QVBoxLayout(self); lay.setSpacing(10); lay.setContentsMargins(16,16,16,16)
        if target == 'reg':
            title = "Реєстр: ЕН + кількість місць"
            hint  = "По одному ЕН на рядку. Місця через пробіл:\n  59001630827897 3"
            ph    = "20451419750910\n59001630827897 3\n..."
        else:
            title = "Посилки: список штрих-кодів"
            hint  = "По одному ЕН/штрих-коду на рядку."
            ph    = "590016308278970001\n20451419750910\n..."
        lbl = QLabel(f"<b>{title}</b>"); lbl.setStyleSheet("font-size:14px;")
        hl  = QLabel(hint); hl.setStyleSheet("color:#666;font-size:12px;"); hl.setWordWrap(True)
        self.ta = QPlainTextEdit(); self.ta.setFont(QFont("Courier",11))
        self.ta.setPlaceholderText(ph); self.ta.setMinimumHeight(200)
        btns = QHBoxLayout(); btns.addStretch()
        bc = QPushButton("Скасувати"); bc.clicked.connect(self.reject)
        bo = QPushButton("Додати"); bo.setObjectName("btn_primary")
        bo.setDefault(True); bo.clicked.connect(self.accept)
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addWidget(lbl); lay.addWidget(hl); lay.addWidget(self.ta); lay.addLayout(btns)

    def get_lines(self):
        return self.ta.toPlainText().strip().splitlines()

# ═══════════════════════════════════════════════════════════════════════════════
#  Діалог: Архів звірок
# ═══════════════════════════════════════════════════════════════════════════════

class HistoryDialog(QDialog):
    def __init__(self, parent, db: DB):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Архів звірок")
        self.setMinimumSize(980, 640)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._build()
        self._load_sessions()

    def _build(self):
        lay = QVBoxLayout(self); lay.setContentsMargins(14,14,14,14); lay.setSpacing(10)

        # ── Фільтр дат ──
        frow = QHBoxLayout(); frow.setSpacing(8)
        frow.addWidget(QLabel("Період:"))
        self.d_from = QDateEdit(calendarPopup=True)
        self.d_from.setDate(QDate.currentDate().addDays(-30))
        self.d_to = QDateEdit(calendarPopup=True)
        self.d_to.setDate(QDate.currentDate())
        for de in (self.d_from, self.d_to): de.setFixedWidth(120)
        frow.addWidget(self.d_from); frow.addWidget(QLabel("—")); frow.addWidget(self.d_to)
        btn_f = QPushButton("Показати"); btn_f.clicked.connect(self._load_sessions)
        frow.addWidget(btn_f); frow.addStretch()
        self.lbl_size = QLabel(); self.lbl_size.setStyleSheet("color:#888;font-size:12px;")
        frow.addWidget(self.lbl_size)
        lay.addLayout(frow)

        # ── Список сесій ──
        lbl_s = QLabel("Збережені звірки:"); lbl_s.setStyleSheet("font-weight:600;")
        lay.addWidget(lbl_s)
        self.sess_tbl = QTableWidget()
        self.sess_tbl.setColumnCount(5)
        self.sess_tbl.setHorizontalHeaderLabels(["ID","Дата","Час збереження","Розбіжностей","Примітка"])
        h = self.sess_tbl.horizontalHeader()
        for col, w in enumerate([50,100,150,120,None]):
            if w: h.setSectionResizeMode(col, QHeaderView.Fixed); self.sess_tbl.setColumnWidth(col, w)
            else: h.setSectionResizeMode(col, QHeaderView.Stretch)
        self.sess_tbl.verticalHeader().setVisible(False)
        self.sess_tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.sess_tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.sess_tbl.setAlternatingRowColors(True)
        self.sess_tbl.setFixedHeight(170)
        self.sess_tbl.itemSelectionChanged.connect(self._on_select)
        lay.addWidget(self.sess_tbl)

        # ── Деталь ──
        self.detail_tabs = QTabWidget()
        specs = [
            ("Реєстр",     "dtbl_reg",  ["№","ЕН","Місць"],       [40,None,70]),
            ("Факт",       "dtbl_fact", ["№","ЕН"],               [40,None]),
            ("Відсутні",   "dtbl_ms",   ["№","ЕН","Примітка","Статус"],[40,None,0,85]),
            ("Зайві",      "dtbl_ex",   ["№","ЕН","Примітка","Статус"],[40,None,0,85]),
            ("Всі",        "dtbl_all",  ["№","ЕН","Примітка","Статус"],[40,None,0,85]),
        ]
        for title, attr, headers, widths in specs:
            w = QWidget(); wl = QVBoxLayout(w); wl.setContentsMargins(6,6,6,6)
            tbl = QTableWidget(); tbl.setColumnCount(len(headers))
            tbl.setHorizontalHeaderLabels(headers)
            hh = tbl.horizontalHeader()
            for i, ww in enumerate(widths):
                if ww: hh.setSectionResizeMode(i, QHeaderView.Fixed); tbl.setColumnWidth(i, ww)
                else:  hh.setSectionResizeMode(i, QHeaderView.Stretch)
            tbl.verticalHeader().setVisible(False)
            tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
            tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
            tbl.setAlternatingRowColors(True)
            wl.addWidget(tbl)
            self.detail_tabs.addTab(w, title)
            setattr(self, attr, tbl)
        lay.addWidget(self.detail_tabs, stretch=1)

        # ── Кнопки ──
        brow = QHBoxLayout(); brow.setSpacing(8)
        b1 = QPushButton("💾 Експорт поточної сесії")
        b1.clicked.connect(self._export_session)
        b2 = QPushButton("📦 Експорт за період (повний CSV)")
        b2.clicked.connect(self._export_range)
        b2.setObjectName("btn_primary")
        bd = QPushButton("🗑 Видалити за період")
        bd.setObjectName("btn_danger"); bd.clicked.connect(self._delete_range)
        bc = QPushButton("Закрити"); bc.clicked.connect(self.accept)
        brow.addWidget(b1); brow.addWidget(b2); brow.addStretch()
        brow.addWidget(bd); brow.addWidget(bc)
        lay.addLayout(brow)

    def _load_sessions(self):
        df = self.d_from.date().toString("yyyy-MM-dd")
        dt = self.d_to.date().toString("yyyy-MM-dd")
        rows = self.db.sessions(df, dt)
        self.sess_tbl.setRowCount(0)
        for r in rows:
            row = self.sess_tbl.rowCount(); self.sess_tbl.insertRow(row)
            vals = [str(r['id']), r['date'], r['created'], str(r['issues']), r['note'] or '']
            for col, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setData(Qt.UserRole, r['id'])
                if col == 3 and int(r['issues']) > 0:
                    item.setForeground(QColor("#C0392B"))
                self.sess_tbl.setItem(row, col, item)
            self.sess_tbl.setRowHeight(row, 28)
        self.lbl_size.setText(f"Розмір БД: {self.db.db_size_kb()} КБ")

    def _on_select(self):
        items = self.sess_tbl.selectedItems()
        if not items: return
        sid = items[0].data(Qt.UserRole)
        info, reg, fact, res = self.db.session_full(sid)
        # Реєстр
        self.dtbl_reg.setRowCount(0)
        for i, r in enumerate(reg):
            self.dtbl_reg.insertRow(i)
            for col, val in enumerate([i+1, fmt_en(r['en']), r['places']]):
                item = QTableWidgetItem(str(val))
                if col == 1: item.setFont(QFont("Courier",11))
                if col == 2: item.setTextAlignment(Qt.AlignCenter)
                self.dtbl_reg.setItem(i, col, item)
            self.dtbl_reg.setRowHeight(i, 26)
        # Факт
        self.dtbl_fact.setRowCount(0)
        for i, r in enumerate(fact):
            self.dtbl_fact.insertRow(i)
            ni = QTableWidgetItem(str(i+1)); ni.setForeground(QColor("#AAA"))
            ei = QTableWidgetItem(fmt_en(r['en'])); ei.setFont(QFont("Courier",11))
            self.dtbl_fact.setItem(i,0,ni); self.dtbl_fact.setItem(i,1,ei)
            self.dtbl_fact.setRowHeight(i,26)
        # Результати
        ms   = [r for r in res if r['status']=='missing']
        ex   = [r for r in res if r['status']=='extra']
        all_ = list(res)
        n_ms = len(ms); n_ex = len(ex)
        self.detail_tabs.setTabText(2, f"Відсутні ({n_ms})")
        self.detail_tabs.setTabText(3, f"Зайві ({n_ex})")
        self.detail_tabs.setTabText(4, f"Всі ({n_ms+n_ex})")
        for tbl, items in [(self.dtbl_ms,ms),(self.dtbl_ex,ex),(self.dtbl_all,all_)]:
            tbl.setRowCount(0)
            for i, r in enumerate(items):
                tbl.insertRow(i)
                ni = QTableWidgetItem(str(i+1)); ni.setForeground(QColor("#AAA"))
                ei = QTableWidgetItem(fmt_en(r['en'])); ei.setFont(QFont("Courier",11))
                li = QTableWidgetItem(r['label'] or ''); li.setForeground(QColor("#888"))
                si = QTableWidgetItem("Відсутня" if r['status']=='missing' else "Зайва")
                si.setTextAlignment(Qt.AlignCenter)
                if r['status']=='missing':
                    si.setForeground(QColor("#8B1F1F")); si.setBackground(QColor("#FFF0F0"))
                else:
                    si.setForeground(QColor(WB_DARK));  si.setBackground(QColor(WB_MINT_L))
                for col, it in enumerate([ni,ei,li,si]): tbl.setItem(i,col,it)
                tbl.setRowHeight(i,26)

    def _selected_sid(self):
        items = self.sess_tbl.selectedItems()
        if not items: QMessageBox.warning(self,"Увага","Оберіть сесію зі списку"); return None
        return items[0].data(Qt.UserRole)

    def _write_full_csv(self, path, data_list):
        with open(path,'w',newline='',encoding='utf-8-sig') as f:
            w = csv.writer(f)
            for info, reg, fact, res in data_list:
                w.writerow([f"=== Звірка {info['date']}  |  {info['created']} ==="])
                w.writerow([])
                w.writerow(["--- РЕЄСТР ---"])
                w.writerow(["№","ЕН","К-сть місць"])
                for i,r in enumerate(reg,1): w.writerow([i, r['en'], r['places']])
                w.writerow([])
                w.writerow(["--- ФАКТИЧНО ВІДСКАНОВАНІ ---"])
                w.writerow(["№","ЕН"])
                for i,r in enumerate(fact,1): w.writerow([i, r['en']])
                w.writerow([])
                w.writerow(["--- РЕЗУЛЬТАТИ ЗВІРКИ ---"])
                w.writerow(["№","ЕН","Примітка","Статус"])
                for i,r in enumerate(res,1):
                    w.writerow([i, r['en'], r['label'] or '',
                                 "Відсутня" if r['status']=='missing' else "Зайва"])
                w.writerow([]); w.writerow([])

    def _export_session(self):
        sid = self._selected_sid()
        if sid is None: return
        info, reg, fact, res = self.db.session_full(sid)
        path, _ = QFileDialog.getSaveFileName(self,"Зберегти CSV",
            f"звірка_{info['date']}.csv","CSV (*.csv)")
        if not path: return
        self._write_full_csv(path, [(info,reg,fact,res)])
        QMessageBox.information(self,"Збережено",f"Файл збережено:\n{path}")

    def _export_range(self):
        df = self.d_from.date().toString("yyyy-MM-dd")
        dt = self.d_to.date().toString("yyyy-MM-dd")
        sessions = self.db.sessions(df, dt)
        if not sessions:
            QMessageBox.information(self,"Немає даних","За обраний період немає збережених звірок")
            return
        path, _ = QFileDialog.getSaveFileName(self,"Зберегти CSV",
            f"звірки_{df}_{dt}.csv","CSV (*.csv)")
        if not path: return
        data = [self.db.session_full(s['id']) for s in sessions]
        self._write_full_csv(path, data)
        QMessageBox.information(self,"Збережено",
            f"Експортовано {len(data)} звірок:\n{path}")

    def _delete_range(self):
        df = self.d_from.date().toString("yyyy-MM-dd")
        dt = self.d_to.date().toString("yyyy-MM-dd")
        reply = QMessageBox.question(self,"Підтвердження",
            f"Видалити всі звірки з {df} по {dt}?\nЦю дію неможливо скасувати.",
            QMessageBox.Yes|QMessageBox.No)
        if reply != QMessageBox.Yes: return
        deleted = self.db.delete_range(df, dt)
        QMessageBox.information(self,"Видалено",
            f"Видалено {deleted} звірок.\nРозмір БД: {self.db.db_size_kb()} КБ")
        self._load_sessions()

# ═══════════════════════════════════════════════════════════════════════════════
#  Головне вікно
# ═══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Звірка накладних — Western Bid")
        self.setMinimumSize(1080, 740)
        self.reg:  OrderedDict = OrderedDict()
        self.fact: OrderedDict = OrderedDict()
        self.last_missing = []
        self.last_extra   = []
        self._building    = False
        self.db = DB()
        self.setup_ui()
        self.setStyleSheet(APP_STYLE)

    # ── Побудова UI ──────────────────────────────────────────────────────────

    def setup_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        outer = QVBoxLayout(central); outer.setContentsMargins(0,0,0,0); outer.setSpacing(0)

        # Хедер
        header = QWidget(); header.setFixedHeight(56)
        header.setStyleSheet(f"background:{WB_DARK};")
        hl = QHBoxLayout(header); hl.setContentsMargins(18,0,18,0); hl.setSpacing(10)
        logo_box = QWidget(); logo_box.setFixedSize(28,28)
        logo_box.setStyleSheet(f"background:{WB_MINT};border-radius:5px;")
        ll = QVBoxLayout(logo_box); ll.setContentsMargins(0,0,0,0)
        lw = QLabel("WB"); lw.setAlignment(Qt.AlignCenter)
        lw.setStyleSheet(f"color:{WB_DARK};font-weight:900;font-size:11px;background:transparent;")
        ll.addWidget(lw)
        names = QVBoxLayout(); names.setSpacing(1)
        n1 = QLabel("Звірка накладних")
        n1.setStyleSheet("color:#FFF;font-size:15px;font-weight:700;")
        n2 = QLabel("Western Bid")
        n2.setStyleSheet(f"color:{WB_MINT};font-size:11px;font-weight:600;letter-spacing:0.06em;")
        names.addWidget(n1); names.addWidget(n2)
        hl.addWidget(logo_box); hl.addLayout(names); hl.addStretch()

        # Дата звірки в хедері
        lbl_d = QLabel("Дата звірки:")
        lbl_d.setStyleSheet(f"color:{WB_MINT_M};font-size:12px;")
        hl.addWidget(lbl_d)
        self.date_edit = QDateEdit(calendarPopup=True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setFixedWidth(128)
        self.date_edit.setStyleSheet(
            f"background:rgba(255,255,255,0.12);border:1px solid {WB_MINT_M};"
            f"border-radius:5px;color:#FFF;padding:4px 8px;"
        )
        hl.addWidget(self.date_edit)

        btn_hist = QPushButton("📋 Архів")
        btn_hist.setObjectName("btn_history"); btn_hist.clicked.connect(self.open_history)
        hl.addWidget(btn_hist)
        ver = QLabel("v3.0"); ver.setStyleSheet("color:#5A9E89;font-size:11px;")
        hl.addWidget(ver)
        outer.addWidget(header)

        accent = QFrame(); accent.setFixedHeight(3)
        accent.setStyleSheet(f"background:{WB_MINT};"); outer.addWidget(accent)

        # Контент
        content = QWidget(); root = QVBoxLayout(content)
        root.setContentsMargins(12,10,12,10); root.setSpacing(8)
        outer.addWidget(content, stretch=1)

        splitter = QSplitter(Qt.Horizontal); splitter.setHandleWidth(6)
        splitter.addWidget(self._build_reg_panel())
        splitter.addWidget(self._build_fact_panel())
        splitter.setSizes([540,540])
        root.addWidget(splitter, stretch=1)

        # Рядок кнопки звірити
        btn_row = QHBoxLayout(); btn_row.setSpacing(14); btn_row.addStretch()
        self.chk_prev = QCheckBox("Шукати відсутні у минулих скануваннях")
        self.chk_prev.setChecked(True)
        btn_row.addWidget(self.chk_prev)
        btn_rec = QPushButton("⚖   Звірити")
        btn_rec.setObjectName("btn_reconcile"); btn_rec.clicked.connect(self.reconcile)
        btn_row.addWidget(btn_rec); btn_row.addStretch()
        root.addLayout(btn_row)

        self.res_widget = QWidget(); self.res_widget.setVisible(False)
        root.addWidget(self.res_widget)
        self._build_result_panel()

    def _make_table(self, headers, widths):
        tbl = QTableWidget(); tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        h = tbl.horizontalHeader()
        for i, w in enumerate(widths):
            if w: h.setSectionResizeMode(i, QHeaderView.Fixed); tbl.setColumnWidth(i, w)
            else: h.setSectionResizeMode(i, QHeaderView.Stretch)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        tbl.setAlternatingRowColors(True)
        return tbl

    def _build_reg_panel(self):
        grp = QGroupBox("Реєстр (закриті посилки)")
        lay = QVBoxLayout(grp); lay.setSpacing(6)
        self.reg_badge = QLabel("0 ЕН · 0 місць")
        self.reg_badge.setObjectName("badge_reg")
        self.reg_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        lay.addWidget(self.reg_badge)
        row1 = QHBoxLayout()
        self.reg_en = QLineEdit(); self.reg_en.setPlaceholderText("Скануйте або введіть ЕН...")
        self.reg_en.returnPressed.connect(lambda: (self.reg_pl.setFocus(), self.reg_pl.selectAll()))
        row1.addWidget(self.reg_en)
        row1.addWidget(QLabel("місць:"))
        self.reg_pl = QSpinBox(); self.reg_pl.setRange(1,999); self.reg_pl.setValue(1)
        self.reg_pl.setFixedWidth(64)
        self.reg_pl.lineEdit().returnPressed.connect(self.add_to_reg)
        row1.addWidget(self.reg_pl)
        ba = QPushButton("+"); ba.setObjectName("btn_primary")
        ba.setFixedWidth(36); ba.clicked.connect(self.add_to_reg)
        row1.addWidget(ba); lay.addLayout(row1)
        hint = QLabel("Enter → «місць» → Enter додає"); hint.setObjectName("hint_label")
        lay.addWidget(hint)
        tools = QHBoxLayout(); tools.setSpacing(5)
        for lbl, fn in [("📂 Файл", lambda:self.load_file('reg')),
                         ("📋 Список", lambda:self.paste_list('reg')),
                         ("📄 PDF", self.load_pdf)]:
            b = QPushButton(lbl)
            if lbl == "📄 PDF" and not PDF_OK:
                b.setEnabled(False); b.setToolTip("pip install PyMuPDF")
            b.clicked.connect(fn); tools.addWidget(b)
        bc = QPushButton("🗑 Очистити"); bc.setObjectName("btn_danger")
        bc.clicked.connect(self.clear_reg); tools.addWidget(bc)
        lay.addLayout(tools)
        self.reg_search = QLineEdit()
        self.reg_search.setPlaceholderText("🔍 Пошук по 4 цифрах або повному ЕН...")
        self.reg_search.textChanged.connect(self.render_reg); lay.addWidget(self.reg_search)
        self.reg_table = self._make_table(["№","ЕН","Місць",""], [36,None,68,30])
        lay.addWidget(self.reg_table)
        return grp

    def _build_fact_panel(self):
        grp = QGroupBox("Фактичне сканування")
        lay = QVBoxLayout(grp); lay.setSpacing(6)
        self.fact_badge = QLabel("0 штрих-кодів")
        self.fact_badge.setObjectName("badge_fact")
        self.fact_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        lay.addWidget(self.fact_badge)
        row1 = QHBoxLayout()
        self.fact_en = QLineEdit(); self.fact_en.setPlaceholderText("Скануйте штрих-код...")
        self.fact_en.returnPressed.connect(self.add_to_fact)
        row1.addWidget(self.fact_en)
        ba = QPushButton("+"); ba.setObjectName("btn_success")
        ba.setFixedWidth(36); ba.clicked.connect(self.add_to_fact)
        row1.addWidget(ba); lay.addLayout(row1)
        hint = QLabel("Автододавання після Enter або сканування"); hint.setObjectName("hint_label")
        lay.addWidget(hint)
        tools = QHBoxLayout(); tools.setSpacing(5)
        for lbl, fn in [("📂 Файл", lambda:self.load_file('fact')),
                         ("📋 Список", lambda:self.paste_list('fact'))]:
            b = QPushButton(lbl); b.clicked.connect(fn); tools.addWidget(b)
        bc = QPushButton("🗑 Очистити"); bc.setObjectName("btn_danger")
        bc.clicked.connect(self.clear_fact); tools.addWidget(bc)
        lay.addLayout(tools)
        self.fact_search = QLineEdit()
        self.fact_search.setPlaceholderText("🔍 Пошук по 4 цифрах або повному ЕН...")
        self.fact_search.textChanged.connect(self.render_fact); lay.addWidget(self.fact_search)
        self.fact_table = self._make_table(["№","ЕН",""], [36,None,30])
        lay.addWidget(self.fact_table)
        return grp

    def _build_result_panel(self):
        lay = QVBoxLayout(self.res_widget); lay.setContentsMargins(0,0,0,0); lay.setSpacing(8)
        mr = QHBoxLayout(); mr.setSpacing(8)
        self.met_vals = []
        for lbl in ["Очікується місць","Відскановано","Відсутніх","Зайвих"]:
            box = QFrame(); box.setObjectName("metric_box"); box.setMinimumHeight(70)
            bl = QVBoxLayout(box)
            v = QLabel("0"); v.setStyleSheet("font-size:26px;font-weight:700;")
            v.setAlignment(Qt.AlignCenter)
            l = QLabel(lbl); l.setStyleSheet("font-size:10px;color:#888;")
            l.setAlignment(Qt.AlignCenter)
            bl.addWidget(v); bl.addWidget(l); mr.addWidget(box); self.met_vals.append(v)
        lay.addLayout(mr)

        # Банер "знайдено у минулих"
        self.lbl_prev = QLabel()
        self.lbl_prev.setStyleSheet(
            f"background:{WB_MINT_L};border:1px solid {WB_MINT_M};"
            f"border-radius:6px;padding:8px 14px;color:{WB_DARK};font-size:12px;"
        )
        self.lbl_prev.setWordWrap(True); self.lbl_prev.setVisible(False)
        lay.addWidget(self.lbl_prev)

        self.tabs = QTabWidget()
        for title, attr in [("Відсутні (0)","tbl_ms"),("Зайві (0)","tbl_ex"),("Всі (0)","tbl_all")]:
            w = QWidget(); tl = QVBoxLayout(w); tl.setContentsMargins(6,6,6,6); tl.setSpacing(6)
            ar = QHBoxLayout()
            for blbl, slot in [("📋 Скопіювати", lambda a=attr: self._copy(a)),
                                ("💾 Експорт CSV", lambda a=attr: self._export_csv(a))]:
                b = QPushButton(blbl); b.clicked.connect(slot); ar.addWidget(b)
            ar.addStretch(); tl.addLayout(ar)
            tbl = self._make_table(["№","ЕН","Примітка","Статус"],[36,None,0,90])
            tl.addWidget(tbl); setattr(self, attr, tbl)
            self.tabs.addTab(w, title)
        lay.addWidget(self.tabs)

    # ── Реєстр ───────────────────────────────────────────────────────────────

    def add_to_reg(self):
        base = norm(self.reg_en.text()); places = self.reg_pl.value()
        if len(base) < 8: self.reg_en.setFocus(); return
        if base not in self.reg: self.reg[base] = places; beep_ok()
        else: beep_dup()
        self.reg_en.clear(); self.reg_pl.setValue(1); self.render_reg(); self.reg_en.setFocus()

    def _expand_registry(self):
        exp = {}
        for base, places in self.reg.items():
            for en, label in expand(base, places):
                exp[en] = {'base': base, 'label': label}
        return exp

    def render_reg(self):
        q = norm(self.reg_search.text())
        tot = sum(self.reg.values())
        self.reg_badge.setText(f"{len(self.reg)} ЕН · {tot} місць")
        entries = list(self.reg.items())
        if q: entries = [(en,p) for en,p in entries if en.endswith(q) or q in en]
        self._building = True
        self.reg_table.setRowCount(0)
        for idx, (en, places) in enumerate(entries):
            row = self.reg_table.rowCount(); self.reg_table.insertRow(row)
            ni = QTableWidgetItem(str(idx+1)); ni.setForeground(QColor("#AAA")); ni.setTextAlignment(Qt.AlignCenter)
            ei = QTableWidgetItem(fmt_en(en)); ei.setFont(QFont("Courier",11)); ei.setData(Qt.UserRole, en)
            exp_items = expand(en, places)
            scanned = sum(1 for ce,_ in exp_items if ce in self.fact)
            if scanned == len(exp_items): ei.setForeground(QColor(WB_MINT))
            elif scanned > 0:             ei.setForeground(QColor("#E8A020"))
            spin = QSpinBox(); spin.setRange(1,999); spin.setValue(places); spin.setFrame(False)
            spin.setStyleSheet("QSpinBox{border:none;background:transparent;font-size:12px;}")
            spin.valueChanged.connect(lambda v, e=en: self._set_places(e, v))
            db = QPushButton("✕")
            db.setStyleSheet("QPushButton{border:none;color:#CCC;background:transparent;padding:0;}"
                             "QPushButton:hover{color:#8B1F1F;background:#FFF0F0;}")
            db.setFixedWidth(28); db.setCursor(Qt.PointingHandCursor)
            db.clicked.connect(lambda _,e=en: self._del_reg(e))
            self.reg_table.setItem(row,0,ni); self.reg_table.setItem(row,1,ei)
            self.reg_table.setCellWidget(row,2,spin); self.reg_table.setCellWidget(row,3,db)
            self.reg_table.setRowHeight(row,30)
        self._building = False

    def _set_places(self, en, val):
        if self._building or en not in self.reg: return
        self.reg[en] = max(1, val)
        self.reg_badge.setText(f"{len(self.reg)} ЕН · {sum(self.reg.values())} місць")

    def _del_reg(self, en): self.reg.pop(en, None); self.render_reg()

    def clear_reg(self):
        if self.reg and QMessageBox.question(self,"Очистити?","Очистити весь реєстр?",
            QMessageBox.Yes|QMessageBox.No) != QMessageBox.Yes: return
        self.reg.clear(); self.render_reg()

    # ── Сканування ───────────────────────────────────────────────────────────

    def add_to_fact(self):
        en = norm(self.fact_en.text())
        if len(en) < 8: return
        if en not in self.fact: self.fact[en] = True; beep_ok()
        else: beep_dup()
        self.fact_en.clear(); self.render_fact(); self.fact_en.setFocus()

    def render_fact(self):
        q = norm(self.fact_search.text())
        self.fact_badge.setText(f"{len(self.fact)} штрих-кодів")
        expanded = self._expand_registry()
        entries = list(self.fact.keys())
        if q: entries = [en for en in entries if en.endswith(q) or q in en]
        self.fact_table.setRowCount(0)
        for idx, en in enumerate(entries):
            row = self.fact_table.rowCount(); self.fact_table.insertRow(row)
            ni = QTableWidgetItem(str(idx+1)); ni.setForeground(QColor("#AAA"))
            ei = QTableWidgetItem(fmt_en(en)); ei.setFont(QFont("Courier",11))
            if en not in expanded: ei.setForeground(QColor("#E8A020"))
            db = QPushButton("✕")
            db.setStyleSheet("QPushButton{border:none;color:#CCC;background:transparent;padding:0;}"
                             "QPushButton:hover{color:#8B1F1F;background:#FFF0F0;}")
            db.setFixedWidth(28); db.setCursor(Qt.PointingHandCursor)
            db.clicked.connect(lambda _,e=en: self._del_fact(e))
            self.fact_table.setItem(row,0,ni); self.fact_table.setItem(row,1,ei)
            self.fact_table.setCellWidget(row,2,db); self.fact_table.setRowHeight(row,30)

    def _del_fact(self, en): self.fact.pop(en, None); self.render_fact()

    def clear_fact(self):
        if self.fact and QMessageBox.question(self,"Очистити?","Очистити список сканування?",
            QMessageBox.Yes|QMessageBox.No) != QMessageBox.Yes: return
        self.fact.clear(); self.render_fact()

    # ── Файли / PDF ──────────────────────────────────────────────────────────

    def load_file(self, target):
        path, _ = QFileDialog.getOpenFileName(self,"Відкрити файл","",
            "Текстові файли (*.txt *.csv);;Всі файли (*)")
        if not path: return
        try:
            with open(path,'r',encoding='utf-8-sig',errors='ignore') as f:
                lines = f.read().splitlines()
            added = parse_lines(lines, target, self.reg, self.fact)
            self.render_reg() if target=='reg' else self.render_fact()
            QMessageBox.information(self,"Завантажено",f"Додано: {added} записів")
        except Exception as e: QMessageBox.critical(self,"Помилка",str(e))

    def paste_list(self, target):
        dlg = PasteDialog(self, target)
        if dlg.exec_() == QDialog.Accepted:
            added = parse_lines(dlg.get_lines(), target, self.reg, self.fact)
            self.render_reg() if target=='reg' else self.render_fact()
            if added: QMessageBox.information(self,"Додано",f"Додано: {added} записів")

    def load_pdf(self):
        if not PDF_OK: QMessageBox.warning(self,"PDF","pip install PyMuPDF"); return
        path, _ = QFileDialog.getOpenFileName(self,"Відкрити PDF","","PDF (*.pdf)")
        if not path: return
        try:
            doc = fitz.open(path); added = 0
            for page in doc:
                text = ' '.join(page.get_text().split())
                for m in re.finditer(r'\b(\d{14})\b', text):
                    en = m.group(1)
                    chunk = text[m.start():m.start()+500]
                    pm = re.search(r'\+380\d{9}\s+(\d{1,3})\s+\d+[.,]\d+', chunk)
                    places = max(1, int(pm.group(1))) if pm else 1
                    if en not in self.reg: self.reg[en] = places; added += 1
            doc.close(); self.render_reg()
            QMessageBox.information(self,"PDF імпорт",
                f"Додано: {added} ЕН\nПеревірте кількість місць у списку.")
        except Exception as e: QMessageBox.critical(self,"Помилка PDF",str(e))

    # ── Звірка ───────────────────────────────────────────────────────────────

    def reconcile(self):
        expanded = self._expand_registry()
        missing, extra = [], []
        for en, info in expanded.items():
            if en not in self.fact:
                missing.append({'en': en, 'label': info['label']})
        for en in self.fact:
            if en not in expanded:
                extra.append({'en': en, 'label': ''})

        # Пошук відсутніх у минулих скануваннях
        found_prev = {}
        session_date = self.date_edit.date().toString("yyyy-MM-dd")
        if self.chk_prev.isChecked() and missing:
            found_prev = self.db.find_fact_en(
                [i['en'] for i in missing], exclude_date=session_date
            )

        # Зберігаємо в БД
        self.db.save(session_date, self.reg, self.fact, missing, extra)

        self.last_missing = missing; self.last_extra = extra

        # Метрики
        for lbl, val, color in zip(self.met_vals,
            [len(expanded), len(self.fact), len(missing), len(extra)],
            [WB_DARK, WB_MID,
             "#C0392B" if missing else WB_MINT,
             "#E8A020" if extra   else WB_MINT]):
            lbl.setText(str(val))
            lbl.setStyleSheet(f"font-size:26px;font-weight:700;color:{color};")

        # Банер минулих
        if found_prev:
            lines = [f"✅ Знайдено у скануваннях минулих днів: {len(found_prev)} шт."]
            for en, d in list(found_prev.items())[:6]:
                lines.append(f"   {fmt_en(en)}  →  відскановано {d}")
            if len(found_prev) > 6:
                lines.append(f"   ... ще {len(found_prev)-6} шт.")
            self.lbl_prev.setText("\n".join(lines)); self.lbl_prev.setVisible(True)
        else:
            self.lbl_prev.setVisible(False)

        # Вкладки
        self.tabs.setTabText(0, f"Відсутні ({len(missing)})")
        self.tabs.setTabText(1, f"Зайві ({len(extra)})")
        self.tabs.setTabText(2, f"Всі ({len(missing)+len(extra)})")

        def enrich_ms(items):
            res = []
            for i in items:
                lbl = i.get('label','')
                if i['en'] in found_prev:
                    lbl = (lbl + f"  ← відскановано {found_prev[i['en']]}").strip()
                res.append({**i, 'label': lbl})
            return res

        ms_e = enrich_ms(missing)
        self._fill_result_table(self.tbl_ms,
            [{'en':i['en'],'label':i['label'],'type':'missing'} for i in ms_e])
        self._fill_result_table(self.tbl_ex,
            [{'en':i['en'],'label':'','type':'extra'} for i in extra])
        self._fill_result_table(self.tbl_all,
            [{'en':i['en'],'label':i['label'],'type':'missing'} for i in ms_e] +
            [{'en':i['en'],'label':'','type':'extra'} for i in extra])

        self.render_reg()
        self.res_widget.setVisible(True)
        beep_ok() if not missing and not extra else beep_err()

    def _fill_result_table(self, tbl, items):
        tbl.setRowCount(0)
        for idx, item in enumerate(items):
            row = tbl.rowCount(); tbl.insertRow(row)
            ni = QTableWidgetItem(str(idx+1)); ni.setForeground(QColor("#AAA")); ni.setTextAlignment(Qt.AlignCenter)
            ei = QTableWidgetItem(fmt_en(item['en'])); ei.setFont(QFont("Courier",11))
            ei.setData(Qt.UserRole, item['en'])
            li = QTableWidgetItem(item.get('label','')); li.setForeground(QColor("#555"))
            is_ms = item['type'] == 'missing'
            si = QTableWidgetItem("Відсутня" if is_ms else "Зайва")
            si.setTextAlignment(Qt.AlignCenter)
            if is_ms: si.setForeground(QColor("#8B1F1F")); si.setBackground(QColor("#FFF0F0"))
            else:     si.setForeground(QColor(WB_DARK));  si.setBackground(QColor(WB_MINT_L))
            for col, it in enumerate([ni,ei,li,si]): tbl.setItem(row,col,it)
            tbl.setRowHeight(row, 28)

    # ── Копіювання / Експорт ─────────────────────────────────────────────────

    def _copy(self, attr):
        tbl = getattr(self, attr)
        lines = [tbl.item(r,1).data(Qt.UserRole) or tbl.item(r,1).text()
                 for r in range(tbl.rowCount()) if tbl.item(r,1)]
        QApplication.clipboard().setText('\n'.join(lines))
        QMessageBox.information(self,"Скопійовано",f"{len(lines)} ЕН скопійовано")

    def _export_csv(self, attr):
        tbl = getattr(self, attr)
        path, _ = QFileDialog.getSaveFileName(self,"Зберегти CSV","result.csv","CSV (*.csv)")
        if not path: return
        try:
            with open(path,'w',newline='',encoding='utf-8-sig') as f:
                w = csv.writer(f); w.writerow(["№","ЕН","Примітка","Статус"])
                for row in range(tbl.rowCount()):
                    w.writerow([
                        tbl.item(row,0).text() if tbl.item(row,0) else '',
                        tbl.item(row,1).data(Qt.UserRole) if tbl.item(row,1) else '',
                        tbl.item(row,2).text() if tbl.item(row,2) else '',
                        tbl.item(row,3).text() if tbl.item(row,3) else '',
                    ])
            QMessageBox.information(self,"Збережено",f"Файл збережено:\n{path}")
        except Exception as e: QMessageBox.critical(self,"Помилка",str(e))

    # ── Архів ────────────────────────────────────────────────────────────────

    def open_history(self):
        HistoryDialog(self, self.db).exec_()


# ═══════════════════════════════════════════════════════════════════════════════
#  Точка входу
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
