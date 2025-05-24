#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TexThon - Excel & Data to LaTeX Converter
===========================================

Version 2.0.0 - Complete Independent Architecture

A modern application for converting Excel tables and data to beautiful LaTeX format.
Each tab is completely independent to prevent cross-tab interference and bugs.

Author: Kazuki Hayashi
License: MIT License
"""

import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, 
                             QVBoxLayout, QWidget, QStatusBar, QMenuBar, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont

# Import all-in-one tab implementations
from excel_latex.excel_latex import ExcelToLatexTab
from tikz_plot.tikz import TikZPlotTab
from info.info import InfoTab


class TexThonMainWindow(QMainWindow):
    """
    TexThon Main Application Window
    
    Features:
    - Complete Independent Architecture
    - No cross-tab interference
    - Safe and reliable conversion
    - Modern UI design
    """
    
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        """Initialize the main application UI"""
        # Window properties
        self.setWindowTitle("TexThon v2.0.0 - Excel & Data to LaTeX Converter")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(800, 600)
        
        # Set application icon if available
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'icon.png')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except:
            pass
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        self.tab_widget.setMovable(False)  # Prevent tab reordering for stability
        
        # Create tabs (each completely independent)
        self.create_tabs()
        
        # Add tab widget to layout
        main_layout.addWidget(self.tab_widget)
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("TexThon v2.0.0 起動完了 - 完全独立アーキテクチャ", 5000)
        
        # Apply modern styling
        self.apply_styling()
        
    def create_menu_bar(self):
        """Create the application menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu('ファイル(&F)')
        
        exit_action = QAction('終了(&X)', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Help menu
        help_menu = menubar.addMenu('ヘルプ(&H)')
        
        about_action = QAction('TexThonについて(&A)', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
    def create_tabs(self):
        """Create all application tabs with complete independence"""
        
        # 1. Excel to LaTeX Tab (完全独立)
        try:
            self.excel_tab = ExcelToLatexTab()
            self.excel_tab.set_status_bar(self.status_bar)
            self.tab_widget.addTab(self.excel_tab, "📊 Excel → LaTeX")
        except Exception as e:
            print(f"Error creating Excel tab: {e}")
            
        # 2. TikZ Plot Tab (完全独立)  
        try:
            self.tikz_tab = TikZPlotTab()
            self.tikz_tab.set_status_bar(self.status_bar)
            self.tab_widget.addTab(self.tikz_tab, "📈 TikZ グラフ")
        except Exception as e:
            print(f"Error creating TikZ tab: {e}")
            
        # 3. Info Tab (完全独立)
        try:
            self.info_tab = InfoTab()
            self.info_tab.set_status_bar(self.status_bar)
            self.tab_widget.addTab(self.info_tab, "ℹ️ 情報")
        except Exception as e:
            print(f"Error creating Info tab: {e}")
    
    def apply_styling(self):
        """Apply modern styling to the application"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                background-color: white;
                border-radius: 8px;
            }
            
            QTabWidget::tab-bar {
                alignment: center;
            }
            
            QTabBar::tab {
                background-color: #e9ecef;
                color: #495057;
                padding: 12px 20px;
                margin-right: 2px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-weight: bold;
                font-size: 14px;
            }
            
            QTabBar::tab:selected {
                background-color: #007bff;
                color: white;
            }
            
            QTabBar::tab:hover:!selected {
                background-color: #f1f3f4;
            }
            
            QStatusBar {
                background-color: #343a40;
                color: white;
                font-size: 12px;
                padding: 5px;
            }
            
            QMenuBar {
                background-color: #fff;
                color: #212529;
                border-bottom: 1px solid #dee2e6;
            }
            
            QMenuBar::item:selected {
                background-color: #007bff;
                color: white;
            }
        """)
    
    def show_about(self):
        """Show about dialog"""
        from PyQt5.QtWidgets import QMessageBox
        
        QMessageBox.about(self, "TexThonについて", 
                         "TexThon v2.0.0\n\n"
                         "Excel表とデータを美しいLaTeX形式に変換\n"
                         "完全独立アーキテクチャで安全・確実\n\n"
                         "作成者: Kazuki Hayashi\n"
                         "ライセンス: MIT License\n\n"
                         "特徴:\n"
                         "🔥 完全独立アーキテクチャ\n"
                         "🛡️ タブ間の干渉なし\n"
                         "📊 Excel表変換\n"
                         "📈 TikZグラフ生成\n"
                         "⚡ 高速・安全処理")
    
    def closeEvent(self, event):
        """Handle application close event"""
        self.status_bar.showMessage("TexThon を終了しています...", 1000)
        event.accept()


def main():
    """Main application entry point"""
    # Create QApplication
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("TexThon")
    app.setApplicationVersion("2.0.0")
    app.setOrganizationName("Kazuki Hayashi")
    app.setOrganizationDomain("github.com/hayashikazuki")
    
    # Set application font
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Create and show main window
    window = TexThonMainWindow()
    window.show()
    
    # Start event loop
    try:
        sys.exit(app.exec_())
    except SystemExit:
        pass


if __name__ == "__main__":
    main() 