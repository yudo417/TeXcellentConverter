import sys
import os
import platform

from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, 
                             QVBoxLayout, QWidget, QStatusBar, QMenuBar, QAction)
from PyQt5.QtCore import Qt, QCoreApplication, QLibraryInfo
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor

# Tab 
from table_latex.table_latex import TableLatexTab
from tikz_plot.tikz_plot import TikZPlotTab
from info.info import InfoTab

def setup_qt_plugins():
    try:
        plugin_path = QLibraryInfo.location(QLibraryInfo.PluginsPath)
        if plugin_path and os.path.exists(plugin_path):
            os.environ['QT_PLUGIN_PATH'] = plugin_path
            return True
    except Exception:
        pass
    
    # Windows固有、開発環境でのみ
    if platform.system() == 'Windows' and not getattr(sys, 'frozen', False):
        try:
            import PyQt5
            pyqt5_dir = os.path.dirname(PyQt5.__file__)
            potential_plugin_paths = [
                os.path.join(pyqt5_dir, 'Qt5', 'plugins'),
                os.path.join(pyqt5_dir, 'Qt', 'plugins'),
                os.path.join(pyqt5_dir, 'plugins'),
            ]
            
            for path in potential_plugin_paths:
                if os.path.exists(os.path.join(path, 'platforms')):
                    os.environ['QT_PLUGIN_PATH'] = path
                    return True
        except Exception:
            pass
    
    return False

def apply_simple_dark_theme(app):
    dark_palette = QPalette()
    
    dark_palette.setColor(QPalette.Window, QColor(33, 37, 43))           # #21252b - 深めのダークブルーグレー
    dark_palette.setColor(QPalette.WindowText, QColor(230, 232, 236))    # #e6e8ec - 暖かみのある白
    dark_palette.setColor(QPalette.Base, QColor(40, 44, 52))             # #282c34 - 入力欄背景
    dark_palette.setColor(QPalette.AlternateBase, QColor(50, 54, 62))    # #32363e - 交互背景
    dark_palette.setColor(QPalette.ToolTipBase, QColor(20, 22, 26))      # #14161a - ツールチップ
    dark_palette.setColor(QPalette.ToolTipText, QColor(255, 255, 255))   # #ffffff - ツールチップテキスト
    dark_palette.setColor(QPalette.Text, QColor(230, 232, 236))          # #e6e8ec - メインテキスト
    dark_palette.setColor(QPalette.Button, QColor(45, 50, 58))           # #2d323a - ボタン背景
    dark_palette.setColor(QPalette.ButtonText, QColor(230, 232, 236))    # #e6e8ec - ボタンテキスト
    dark_palette.setColor(QPalette.BrightText, QColor(255, 255, 255))    # #ffffff - 強調テキスト
    dark_palette.setColor(QPalette.Link, QColor(97, 175, 239))           # #61afef - リンク色
    dark_palette.setColor(QPalette.Highlight, QColor(97, 175, 239))      # #61afef - ハイライト
    dark_palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))     # #000000 - ハイライトテキスト
    
    dark_palette.setColor(QPalette.Disabled, QPalette.WindowText, QColor(130, 135, 145))
    dark_palette.setColor(QPalette.Disabled, QPalette.Text, QColor(130, 135, 145))
    dark_palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(130, 135, 145))
    
    app.setPalette(dark_palette)
    
    app.setStyleSheet("""
        QWidget {
            background-color: #21252b;
            color: #e6e8ec;
            font-family: "Segoe UI", "Yu Gothic UI", "Meiryo UI", sans-serif;
        }
        
        QMainWindow {
            background-color: #21252b;
            border: none;
        }
        
        QTabWidget::pane {
            background-color: #282c34;
            border: 1px solid #3e4147;
            border-radius: 6px;
            margin-top: -1px;
        }
        
        QTabBar::tab {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #3c414a, stop:1 #32363e);
            color: #abb2bf;
            padding: 10px 18px;
            border: 1px solid #3e4147;
            border-bottom: none;
            border-top-left-radius: 6px;
            border-top-right-radius: 6px;
            margin-right: 2px;
            min-width: 80px;
        }
        
        QTabBar::tab:selected {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
            color: #ffffff;
            border: 1px solid #61afef;
            border-bottom: none;
            font-weight: bold;
        }
        
        QTabBar::tab:hover:!selected {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #4a5059, stop:1 #3c414a);
            color: #e6e8ec;
            border: 1px solid #61afef;
            border-bottom: none;
        }
        
        QLineEdit {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            padding: 8px 12px;
            font-size: 13px;
            selection-background-color: #61afef;
            selection-color: #ffffff;
        }
        
        QLineEdit:hover {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QLineEdit:focus {
            background-color: #32363e;
            border: 2px solid #61afef;
            box-shadow: 0 0 0 3px rgba(97, 175, 239, 0.1);
        }
        
        QLineEdit:disabled {
            background-color: #1e2126;
            color: #828791;
            border: 2px solid #2c3036;
        }
        
        QTextEdit, QPlainTextEdit {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            padding: 8px;
            selection-background-color: #61afef;
            selection-color: #ffffff;
        }
        
        QTextEdit:hover, QPlainTextEdit:hover {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QTextEdit:focus, QPlainTextEdit:focus {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #4a5059, stop:1 #3c414a);
            color: #e6e8ec;
            border: 1px solid #5a6069;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 500;
            min-height: 18px;
        }
        
        QPushButton:hover:enabled {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #5a6069, stop:1 #4a5059);
            border: 1px solid #61afef;
            color: #ffffff;
        }
        
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
            border: 1px solid #528bcc;
            color: #ffffff;
        }
        
        QPushButton:disabled {
            background-color: #2c3036;
            color: #828791;
            border: 1px solid #3e4147;
        }
        
        QPushButton:focus {
            border: 2px solid #61afef;
        }
        
        QGroupBox {
            color: #e6e8ec;
            background-color: #32363e;
            border: 2px solid #3e4147;
            border-radius: 8px;
            margin-top: 12px;
            padding-top: 12px;
            font-weight: 600;
        }
        
        QGroupBox:hover {
            border: 2px solid #61afef;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 8px 0 8px;
            color: #61afef;
            background-color: #32363e;
            font-weight: bold;
        }
        
        QLabel {
            color: #e6e8ec;
            background-color: transparent;
        }
        
        QLabel:disabled {
            color: #828791;
        }
        
        QComboBox {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            padding: 6px 12px;
            selection-background-color: #61afef;
            min-height: 20px;
        }
        
        QComboBox:hover:enabled {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QComboBox:focus {
            border: 2px solid #61afef;
            background-color: #32363e;
        }
        
        QComboBox:disabled {
            background-color: #1e2126;
            color: #828791;
            border: 2px solid #2c3036;
        }
        
        QComboBox::drop-down {
            border: none;
            width: 20px;
            border-top-right-radius: 6px;
            border-bottom-right-radius: 6px;
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
        }
        
        QComboBox::drop-down:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #528bcc, stop:1 #4a7bc8);
        }
        
        QComboBox::down-arrow {
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 8px solid #ffffff;
        }
        
        QComboBox QAbstractItemView {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #61afef;
            border-radius: 6px;
            selection-background-color: #61afef;
            selection-color: #ffffff;
            outline: none;
        }
        
        QComboBox QAbstractItemView::item {
            padding: 8px 12px;
            border: none;
        }
        
        QComboBox QAbstractItemView::item:hover {
            background-color: #3e4147;
            color: #ffffff;
        }
        
        QComboBox QAbstractItemView::item:selected {
            background-color: #61afef;
            color: #ffffff;
        }
        
        QCheckBox, QRadioButton {
            color: #e6e8ec;
            background-color: transparent;
            spacing: 8px;
            padding: 4px;
        }
        
        QCheckBox:hover, QRadioButton:hover {
            color: #ffffff;
        }
        
        QCheckBox:disabled, QRadioButton:disabled {
            color: #828791;
        }
        
        QCheckBox::indicator, QRadioButton::indicator {
            width: 18px;
            height: 18px;
            background-color: #282c34;
            border: 2px solid #3e4147;
            border-radius: 4px;
        }
        
        QCheckBox::indicator:hover, QRadioButton::indicator:hover {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QCheckBox::indicator:focus, QRadioButton::indicator:focus {
            border: 2px solid #61afef;
            box-shadow: 0 0 0 3px rgba(97, 175, 239, 0.2);
        }
        
        QCheckBox::indicator:checked, QRadioButton::indicator:checked {
            background-color: #61afef;
            border: 2px solid #61afef;
        }
        
        QCheckBox::indicator:checked:hover, QRadioButton::indicator:checked:hover {
            background-color: #528bcc;
            border: 2px solid #528bcc;
        }
        
        QCheckBox::indicator:disabled, QRadioButton::indicator:disabled {
            background-color: #1e2126;
            border: 2px solid #2c3036;
        }
        
        QRadioButton::indicator {
            border-radius: 9px;
        }
        
        QTableWidget, QTableView {
            background-color: #282c34;
            alternate-background-color: #32363e;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            gridline-color: #3e4147;
            selection-background-color: #61afef;
            selection-color: #ffffff;
        }
        
        QTableWidget::item:hover, QTableView::item:hover {
            background-color: #3e4147;
        }
        
        QTableWidget::item:selected, QTableView::item:selected {
            background-color: #61afef;
            color: #ffffff;
        }
        
        QHeaderView::section {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
            color: #ffffff;
            border: 1px solid #528bcc;
            padding: 8px;
            font-weight: bold;
        }
        
        QHeaderView::section:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #528bcc, stop:1 #4a7bc8);
        }
        
        QHeaderView::section:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #4a7bc8, stop:1 #3d6bb0);
        }
        
        QScrollBar:vertical, QScrollBar:horizontal {
            background-color: #32363e;
            border: 1px solid #3e4147;
            border-radius: 6px;
        }
        
        QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
            border: 1px solid #528bcc;
            border-radius: 5px;
            min-height: 20px;
        }
        
        QScrollBar::handle:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #528bcc, stop:1 #4a7bc8);
        }
        
        QScrollBar::handle:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #4a7bc8, stop:1 #3d6bb0);
        }
        
        QScrollBar::add-line:hover, QScrollBar::sub-line:hover {
            background-color: #4a5059;
        }
        
        QSpinBox, QDoubleSpinBox {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            padding: 6px 8px;
            font-size: 13px;
            selection-background-color: #61afef;
            selection-color: #ffffff;
        }
        
        QSpinBox:hover, QDoubleSpinBox:hover {
            background-color: #32363e;
            border: 2px solid #61afef;
        }
        
        QSpinBox:focus, QDoubleSpinBox:focus {
            border: 2px solid #61afef;
            box-shadow: 0 0 0 3px rgba(97, 175, 239, 0.1);
        }
        
        QSpinBox:disabled, QDoubleSpinBox:disabled {
            background-color: #1e2126;
            color: #828791;
            border: 2px solid #2c3036;
        }
        
        QProgressBar {
            background-color: #e9ecef;
            color: #2c3e50;
            border: 2px solid #dee2e6;
            border-radius: 6px;
            text-align: center;
            font-weight: bold;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #61afef, stop:1 #528bcc);
            border-radius: 4px;
        }
        
        QStatusBar {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #32363e, stop:1 #21252b);
            color: #e6e8ec;
            border-top: 1px solid #3e4147;
        }
        
        QListWidget {
            background-color: #282c34;
            color: #e6e8ec;
            border: 2px solid #3e4147;
            border-radius: 6px;
            selection-background-color: #61afef;
            selection-color: #ffffff;
            outline: none;
        }
        
        QListWidget::item {
            padding: 8px;
            border-bottom: 1px solid #3e4147;
        }
        
        QListWidget::item:hover {
            background-color: #3e4147;
        }
        
        QListWidget::item:selected {
            background-color: #61afef;
            color: #ffffff;
        }
        
        QMenuBar {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                      stop:0 #32363e, stop:1 #21252b);
            color: #e6e8ec;
            border-bottom: 1px solid #3e4147;
        }
        
        QMenuBar::item {
            background-color: transparent;
            padding: 6px 12px;
            border-radius: 4px;
        }
        
        QMenuBar::item:hover {
            background-color: #4a5059;
            color: #ffffff;
        }
        
        QMenuBar::item:selected {
            background-color: #61afef;
            color: #ffffff;
        }
        
        QMenu {
            background-color: #ffffff;
            color: #2c3e50;
            border: 2px solid #61afef;
            border-radius: 6px;
        }
        
        QMenu::item {
            padding: 8px 20px;
        }
        
        QMenu::item:hover {
            background-color: #e3f2fd;
        }
        
        QMenu::item:selected {
            background-color: #61afef;
            color: #ffffff;
        }
        
        QMenu::item:disabled {
            color: #6c757d;
        }
    """)

class TeXcellentConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('TeXcellentConverter')
        screen = QApplication.primaryScreen().geometry()
        max_width = int(screen.width() * 0.95)
        max_height = int(screen.height() * 0.95)
        initial_height = int(screen.height() * 0.8)
        initial_width = int(screen.width() * 0.4)
        self.resize(min(initial_width, max_width), min(initial_height, max_height))
        self.setMinimumSize(600, 600)
        self.move(50, 50)

        # メインウィジェット
        mainWidget = QWidget()
        mainLayout = QVBoxLayout()
        
        # タブウィジェット
        self.tabWidget = QTabWidget()
        
        # 3タブ
        self.excelTab = TableLatexTab()
        self.tikzTab = TikZPlotTab()
        self.infoTab = InfoTab()
        
        self.tabWidget.addTab(self.excelTab, "Excel → LaTeX")
        self.tabWidget.addTab(self.tikzTab, "グラフ生成")
        self.tabWidget.addTab(self.infoTab, "この情報・ライセンス")
        
        mainLayout.addWidget(self.tabWidget)
        
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        self.excelTab.set_status_bar(self.statusBar)
        self.tikzTab.set_status_bar(self.statusBar)
        
        mainWidget.setLayout(mainLayout)
        self.setCentralWidget(mainWidget)


if __name__ == '__main__':
    setup_qt_plugins()
    
    app = QApplication(sys.argv)
    
    # Win版のみ
    if platform.system() == 'Windows':
        apply_simple_dark_theme(app)
    
    ex = TeXcellentConverterApp()
    ex.show()
    
    sys.exit(app.exec_())