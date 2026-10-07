"""Fogli di stile (QSS) centralizzati della suite.

Tutti gli stili dell'interfaccia sono definiti qui, così palette, spaziature e
correzioni grafiche restano coerenti tra Home, Analyzer e Checklist.
"""

from __future__ import annotations

from PySide6.QtWidgets import QComboBox, QStyledItemDelegate

# Palette condivisa (contrasti verificati WCAG AA per il testo sui rispettivi sfondi).
SIDEBAR_BG = "#2c3e50"
SIDEBAR_FIELD_BG = "#34495e"
SIDEBAR_FIELD_BORDER = "#60758a"
ACCENT_BLUE = "#2878a8"
TEXT_ON_DARK = "#ffffff"

SUITE_QSS = """
QMainWindow { background: #f4f6f8; }
QTabWidget::pane { border: 0; }
QTabBar::tab {
    min-width: 150px; padding: 8px 16px; background: #e7ebef; color: #334;
    border: 1px solid #d0d6dc; border-bottom: none;
}
QTabBar::tab:selected { background: white; font-weight: 800; border-top: 3px solid #2878a8; }
QStatusBar { background: white; border-top: 1px solid #dfe5ea; }
"""

HOME_QSS = """
#homeTitle { font-size: 30px; font-weight: 800; color: #22313f; }
#homeSubtitle { font-size: 15px; color: #52616b; }
#workflowCard { background: white; border: 1px solid #dfe5ea; border-radius: 8px; }
#workflowCard QLabel { font-size: 13px; color: #2d3e4f; }
#primaryButton, #secondaryButton { min-height: 40px; padding: 0 20px; border-radius: 8px; font-weight: 700; }
#primaryButton { background: #2878a8; color: white; border: none; }
#primaryButton:hover { background: #3490c4; }
#secondaryButton { background: #2f855a; color: white; border: none; }
#secondaryButton:hover { background: #38a169; }
#versionLabel { color: #6b7782; }
"""

ANALYZER_QSS = f"""
QMainWindow {{ background: #f5f6fa; }}
#sidebar, #sidebarControls, #sidebarScroll, #sidebarScroll > QWidget > QWidget {{ background: {SIDEBAR_BG}; border: none; }}
#sidebar QLabel {{ color: #ecf0f1; }}
#sidebarTitle {{ color: #67b7f7; font-size: 16px; font-weight: 800; }}
#sectionTitle {{ color: #9fd1ff; font-weight: 800; margin-top: 4px; }}
#fileLabel {{ color: #d0d7de; font-style: italic; }}
#separator {{ background: #4b6175; max-height: 1px; }}
#sidebar QLineEdit, #sidebar QSpinBox, #sidebar QComboBox {{
    min-height: 28px; padding: 4px 8px; border: 1px solid {SIDEBAR_FIELD_BORDER}; border-radius: 4px;
    background: {SIDEBAR_FIELD_BG}; color: {TEXT_ON_DARK};
}}
#sidebar QComboBox QAbstractItemView {{
    background: {SIDEBAR_FIELD_BG}; color: {TEXT_ON_DARK};
    border: 1px solid {SIDEBAR_FIELD_BORDER};
    selection-background-color: {ACCENT_BLUE}; selection-color: {TEXT_ON_DARK};
    outline: none;
}}
#sidebar QComboBox {{ combobox-popup: 0; }}
#sidebar QComboBox QAbstractItemView::item {{ min-height: 28px; padding: 0 8px; color: {TEXT_ON_DARK}; }}
#sidebar QComboBox QAbstractItemView::item:hover,
#sidebar QComboBox QAbstractItemView::item:selected {{ background: {ACCENT_BLUE}; color: {TEXT_ON_DARK}; }}
#sidebar QCheckBox {{ color: white; }}
#sidebar QPushButton {{ min-height: 32px; background: #405a73; color: white; border: 0; border-radius: 4px; }}
#sidebar QPushButton:hover {{ background: #4d6d8b; }}
#btnStart {{ background: #268c4f; font-weight: 700; }}
#btnStop {{ background: #a63a32; font-weight: 700; }}
#btnExport {{ background: #2878a8; font-weight: 700; }}
#btnExit {{ background: #6d7478; font-weight: 700; }}
#btnMiniHelp {{ min-width: 28px; max-width: 28px; border-radius: 14px; background: #2878a8; font-weight: 800; }}
#contentTitle {{ background: white; border: 1px solid #e0e4e8; border-radius: 8px; padding: 8px; font-size: 16px; font-weight: 800; }}
QTextEdit, QTableWidget {{ background: white; border: 1px solid #d9dee3; border-radius: 4px; }}
QTabWidget::pane {{ border: 1px solid #d9dee3; background: white; }}
QTabBar::tab {{ padding: 8px 12px; background: #e8ebef; }}
QTabBar::tab:selected {{ background: white; font-weight: 700; border-bottom: 2px solid #2878a8; }}
"""

CHECKLIST_TABLE_QSS = """
#checklistSummary {
    background: #ffffff;
    color: #263746;
    border: 1px solid #dfe5ea;
    border-left: 4px solid #2878a8;
    border-radius: 6px;
    padding: 10px 12px;
    font-size: 13px;
    font-weight: 600;
}
#checklistPrimaryButton {
    min-height: 38px;
    background: #2878a8;
    color: #ffffff;
    border: none;
    border-radius: 5px;
    padding: 0 12px;
    font-weight: 700;
}
#checklistPrimaryButton:hover { background: #3490c4; }
QTableWidget {
    background-color: #ffffff;
    gridline-color: #d9d9d9;
    selection-background-color: #e6e6e6;
    selection-color: #000000;
}
QTableWidget::item {
    padding: 4px;
}
QTableWidget::item:hover {
    background-color: #eeeeee;
    color: #000000;
}
QTableWidget::item:selected {
    background-color: #dcdcdc;
    color: #000000;
}
QHeaderView::section {
    background-color: #f3f3f3;
    border: 1px solid #d0d0d0;
    padding: 4px;
    font-weight: 600;
}
"""


def enable_styled_combo_popup(combo: QComboBox) -> None:
    """Rende il popup della combo governabile dal foglio di stile.

    Con lo stile Fusion Qt disegna le voci del popup con un delegate "menu" che
    usa il colore testo della combo stessa: su una combo con testo bianco le
    opzioni diventavano invisibili sul fondo chiaro del popup. Un
    QStyledItemDelegate rispetta invece le regole ``QAbstractItemView::item``.
    """
    combo.setItemDelegate(QStyledItemDelegate(combo))
