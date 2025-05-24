import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QTabWidget, QStatusBar)
from excel_latex.ui import ExcelToLatexTab
from info.ui import InfoTab


class TexcellentConverterApp(QMainWindow):
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

        # メインウィジェットとタブウィジェット
        mainWidget = QWidget()
        mainLayout = QVBoxLayout()
        
        # タブウィジェットを作成
        self.tabWidget = QTabWidget()
        
        # 3つのタブを作成
        self.excelTab = ExcelToLatexTab()
        # self.tikzTab = TikZPlotTab()  # 一時的にコメントアウト
        self.infoTab = InfoTab()
        
        # タブをタブウィジェットに追加
        self.tabWidget.addTab(self.excelTab, "Excel → LaTeX 表")
        # self.tabWidget.addTab(self.tikzTab, "TikZ グラフ")  # 一時的にコメントアウト
        self.tabWidget.addTab(self.infoTab, "情報・ライセンス")
        
        # メインレイアウトに追加
        mainLayout.addWidget(self.tabWidget)
        
        # ステータスバー
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # タブからステータスバーにアクセスできるように設定
        self.excelTab.set_status_bar(self.statusBar)
        # self.tikzTab.set_status_bar(self.statusBar)  # 一時的にコメントアウト
        
        # メインウィジェット設定
        mainWidget.setLayout(mainLayout)
        self.setCentralWidget(mainWidget)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TexcellentConverterApp()
    ex.show()
    sys.exit(app.exec_()) 