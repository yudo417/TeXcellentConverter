
import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, 
                             QVBoxLayout, QWidget, QStatusBar, QMenuBar, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont

# Tab 
from table_latex.table_latex import TableLatexTab
from tikz_plot.tikz_plot import TikZPlotTab
from info.info import InfoTab

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
    app = QApplication(sys.argv)
    ex = TeXcellentConverterApp()
    ex.show()
    sys.exit(app.exec_())