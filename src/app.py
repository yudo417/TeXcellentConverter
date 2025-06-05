import sys
import os
import platform

from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, 
                             QVBoxLayout, QWidget, QStatusBar, QMenuBar, QAction)
from PyQt5.QtCore import Qt, QCoreApplication, QLibraryInfo
from PyQt5.QtGui import QIcon, QFont

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
    # Qtプラグインの設定をQApplicationの作成前に実行
    setup_qt_plugins()
    
    app = QApplication(sys.argv)
    
    ex = TeXcellentConverterApp()
    ex.show()
    sys.exit(app.exec_())