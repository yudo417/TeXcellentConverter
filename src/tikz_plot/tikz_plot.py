import os
import numpy as np
import pandas as pd
import sys
import math  # 数学関数を使用するためにインポート
import traceback
import openpyxl
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QLabel, QLineEdit, QPushButton,
                            QFileDialog, QComboBox, QMessageBox, QTextEdit,
                            QCheckBox, QGridLayout, QGroupBox, QSplitter,
                            QStatusBar, QFrame, QTabWidget, QTableWidget,
                            QTableWidgetItem, QColorDialog, QSpinBox, 
                            QDoubleSpinBox, QHeaderView, QListWidget,
                            QFormLayout, QRadioButton, QButtonGroup, QInputDialog,
                            QSizePolicy, QScrollArea)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
import re


class TikZPlotTab(QWidget):
    """TikZ plot converter tab"""
    
    def __init__(self):
        super().__init__()
        
        self.datasets = []  
        """
        全データセット格納
        ### 格納予定セット
            dataset = {
                'name': final_name, # Always a string
                'data_source_type': 'measured',  # 'measured' または 'formula'
                'data_x': [],
                'data_y': [],
                'color': QColor('blue'),  # QColorオブジェクトを新規作成
                'line_width': 1.0,
                'marker_style': '*',
                'marker_size': 2.0,
                'plot_type': "line",
                'legend_label': final_name, # Always a string, initialized with name
                'show_legend': True,
                'equation': '',
                'domain_min': 0,
                'domain_max': 10,
                'samples': 200,
                'special_points': [],  # [(x, y, color, show_coords), ...]
                'annotations': [],     # [(x, y, text, color, pos), ...]
                'special_points_enabled': False,  # 特殊点チェックボックスの状態
                'annotations_enabled': False,    # 注釈チェックボックスの状態
                ### ファイル読み込み関連の設定
                'file_path': '',
                'file_type': 'csv',  # 'csv' or 'excel' or 'manual'
                'sheet_name': '',
                'x_column': '',
                'y_column': ''
            }
        """
        self.current_dataset_index = -1
        """現在選択されているデータセットのインデックス（初期値-1)"""
        
        self.global_settings = {
            'x_label': 'x軸',
            'y_label': 'y軸',
            'x_min': 0,
            'x_max': 10,
            'y_min': 0,
            'y_max': 10,
            'grid': True,
            'show_legend': True,
            'legend_pos': 'north east',  
            'width': 0.8,
            'height': 0.6,
            'caption': 'グラフのキャプション',
            'label': 'fig:tikz_plot',
            'position': 'H',
            'scale_type': 'normal'  
        }
        """
        グラフ全体の設定\n
        x_labe：x軸ラベル | y_label：y軸ラベル\n
        x_min：x軸の最小値 | x_max：x軸の最大値\n
        y_min：y軸の最小値 | y_max：y軸の最大値\n
        grid：グリッド表示\n
        show_legend：凡例表示Bool | legend_pos：凡例の位置\n
        width：グラフの幅 | height：グラフの高さ\n
        caption：グラフのキャプション | label：グラフのラベル |position：グラフの位置(H)\n
        scale_type：軸の目盛りタイプ(対数メモリとか)
        """
        
        self.initUI()
        
        # 初期データセット追加
        QTimer.singleShot(0, lambda: self.add_dataset("データセット1"))

    def set_status_bar(self, status_bar):
        self.statusBar = status_bar

    def color_to_tikz_rgb(self, color):
        if color == QColor('red'):
            return 'red'
        elif color == QColor('green'):
            return 'green'
        elif color == QColor('blue'):
            return 'blue'
        elif color == QColor('black'):
            return 'black'
        elif color == QColor('yellow'):
            return 'yellow'
        elif color == QColor('cyan'):
            return 'cyan'
        elif color == QColor('magenta'):
            return 'magenta'
        elif color == QColor('orange'):
            return 'orange'
        elif color == QColor('purple'):
            return 'purple'
        elif color == QColor('brown'):
            return 'brown'
        elif color == QColor('gray'):
            return 'gray'
        else:
            r = color.red() / 255.0
            g = color.green() / 255.0
            b = color.blue() / 255.0
            return f"color = {{rgb,255:red,{color.red()};green,{color.green()};blue,{color.blue()}}}"
            

    def initUI(self):
        self.setWindowTitle('TikZPlot Converter')
        
        screen = QApplication.primaryScreen().geometry()
        max_width = int(screen.width() * 0.95)
        max_height = int(screen.height() * 0.95)
        initial_height = int(screen.height() * 0.9)
        initial_width = int(screen.width() * 0.4)
        self.resize(min(initial_width, max_width), min(initial_height, max_height))
        self.setMinimumSize(600, 600)
        self.move(50, 50)
        
        mainWidget = QWidget()
        mainLayout = QVBoxLayout()
        
        splitter = QSplitter(Qt.Vertical)
        
        # --- 上部：設定部分 ---
        #* ======================settingWidget========================= 
        settingsWidget = QWidget()
        settingsLayout = QVBoxLayout()
        settingsWidget.setLayout(settingsLayout)
        
        infoLayout = QVBoxLayout()
        
        infoLabel = QLabel("注意: このグラフを使用するには、LaTeXドキュメントのプリアンブルに以下を追加してください:")
        infoLabel.setStyleSheet("color: #cc0000; font-weight: bold;") 
        
        packageLayout = QHBoxLayout()
        
        packageLabel = QLabel("\\usepackage{pgfplots}\\usepackage{float}\\pgfplotsset{compat=1.18}")
        packageLabel.setStyleSheet("font-family: monospace;")
        
        copyPackageButton = QPushButton("コピー")
        copyPackageButton.setFixedWidth(60)
        def copy_package():
            clipboard = QApplication.clipboard()
            clipboard.setText("\\usepackage{pgfplots}\n\\usepackage{float}\n\\pgfplotsset{compat=1.18}")
            self.statusBar.showMessage("パッケージ名をコピーしました", 3000)
        copyPackageButton.clicked.connect(copy_package)
        
        packageLayout.addWidget(packageLabel)
        packageLayout.addWidget(copyPackageButton)
        packageLayout.addStretch()
        
        infoLayout.addWidget(infoLabel)
        infoLayout.addLayout(packageLayout)
        
        datasetGroup = QGroupBox("データセット管理")
        datasetLayout = QVBoxLayout()
        
        self.datasetList = QListWidget()
        self.datasetList.currentRowChanged.connect(self.on_dataset_selected)
        self.datasetList.setMinimumHeight(100) 
        self.datasetList.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        datasetLayout.addWidget(QLabel("データセット:"))
        datasetLayout.addWidget(self.datasetList)
        
        datasetButtonLayout = QHBoxLayout()
        addDatasetButton = QPushButton("データセット追加")
        addDatasetButton.clicked.connect(self.add_dataset)
        removeDatasetButton = QPushButton("データセット削除")
        removeDatasetButton.clicked.connect(self.remove_dataset)
        renameDatasetButton = QPushButton("名前変更")
        renameDatasetButton.clicked.connect(self.rename_dataset)
        datasetButtonLayout.addWidget(addDatasetButton)
        datasetButtonLayout.addWidget(removeDatasetButton)
        datasetButtonLayout.addWidget(renameDatasetButton)
        
        datasetLayout.addLayout(datasetButtonLayout)
        datasetGroup.setLayout(datasetLayout)
        
        #* ======================TabWidget========================= 
        tabWidget = QTabWidget()
        
        #* ======================データ入力タブ========================= 
        dataTab = QWidget()
        dataTabLayout = QVBoxLayout()
        
        dataSourceTypeGroup = QGroupBox("データソースタイプ")
        dataSourceTypeLayout = QVBoxLayout()
        
        self.measuredRadio = QRadioButton("実測値データ（CSV/Excel/手入力）")
        self.formulaRadio = QRadioButton("数式によるグラフ生成")
        self.measuredRadio.setChecked(True)  # デフォルトは実測値
        
        dataSourceTypeButtonGroup = QButtonGroup(self)
        dataSourceTypeButtonGroup.addButton(self.measuredRadio)
        dataSourceTypeButtonGroup.addButton(self.formulaRadio)
        
        radioStyle = "QRadioButton { font-weight: bold; font-size: 14px; padding: 5px; }"
        self.measuredRadio.setStyleSheet(radioStyle)
        self.formulaRadio.setStyleSheet(radioStyle)
        
        measuredLabel = QLabel("実験などで取得した実測データを使用してグラフを生成します")
        measuredLabel.setStyleSheet("color: gray; font-style: italic; padding-left: 20px; min-height: 20px;")
        formulaLabel = QLabel("入力した数式に基づいて理論曲線を生成します")
        formulaLabel.setStyleSheet("color: gray; font-style: italic; padding-left: 20px; min-height: 20px;")
        
        dataSourceTypeLayout.addWidget(self.measuredRadio)
        dataSourceTypeLayout.addWidget(measuredLabel)
        dataSourceTypeLayout.addWidget(self.formulaRadio)
        dataSourceTypeLayout.addWidget(formulaLabel)
        
        dataSourceTypeGroup.setFixedHeight(150)
        
        self.measuredRadio.toggled.connect(self.on_data_source_type_changed)
        self.formulaRadio.toggled.connect(self.on_data_source_type_changed)
        
        dataSourceTypeGroup.setLayout(dataSourceTypeLayout)
        dataTabLayout.addWidget(dataSourceTypeGroup)
        
        #* ======================実測値========================= 
        self.measuredContainer = QWidget()
        """実測データ入力用のコンテナ"""
        measuredLayout = QVBoxLayout(self.measuredContainer)
        """実測値固有レイアウト"""
        
        dataSourceGroup = QGroupBox("データソース")
        dataSourceLayout = QVBoxLayout()
        
        csvLayout = QHBoxLayout()
        self.csvRadio = QRadioButton("CSVファイル:")
        self.fileEntry = QLineEdit()
        browseButton = QPushButton('参照...')
        browseButton.clicked.connect(self.browse_csv_file)
        csvLayout.addWidget(self.csvRadio)
        csvLayout.addWidget(self.fileEntry)
        csvLayout.addWidget(browseButton)
        
        excelLayout = QHBoxLayout()
        self.excelRadio = QRadioButton("Excelファイル:")
        self.excelEntry = QLineEdit()
        excelBrowseButton = QPushButton('参照...')
        excelBrowseButton.clicked.connect(self.browse_excel_file)
        excelLayout.addWidget(self.excelRadio)
        excelLayout.addWidget(self.excelEntry)
        excelLayout.addWidget(excelBrowseButton)
        
        sheetLayout = QHBoxLayout()
        sheetLabel = QLabel('シート名:')
        sheetLabel.setIndent(30)  
        self.sheetCombobox = QComboBox()
        self.sheetCombobox.setEnabled(False)  
        sheetLayout.addWidget(sheetLabel)
        sheetLayout.addWidget(self.sheetCombobox)
        
        self.manualRadio = QRadioButton("データを直接入力:")
        self.manualRadio.setChecked(True)  
        
        sourceGroup = QButtonGroup(self)
        sourceGroup.addButton(self.csvRadio)
        sourceGroup.addButton(self.excelRadio)
        sourceGroup.addButton(self.manualRadio)
        sourceGroup.buttonClicked.connect(self.toggle_source_fields)
        
        self.dataTable = QTableWidget(10, 2)  # 行数, 列数
        self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
        self.dataTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dataTable.setEnabled(True)  
        self.dataTable.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        tableButtonLayout = QHBoxLayout()
        self.addRowButton = QPushButton('行を追加')
        self.addRowButton.clicked.connect(self.add_table_row)
        self.removeRowButton = QPushButton('選択行を削除')
        self.removeRowButton.clicked.connect(self.remove_table_row)
        self.addRowButton.setEnabled(False)
        self.removeRowButton.setEnabled(False)
        tableButtonLayout.addWidget(self.addRowButton)
        tableButtonLayout.addWidget(self.removeRowButton)
        
        dataSourceLayout.addLayout(csvLayout)
        dataSourceLayout.addLayout(excelLayout)
        dataSourceLayout.addLayout(sheetLayout)
        dataSourceLayout.addWidget(self.manualRadio)
        dataSourceLayout.addWidget(self.dataTable)
        dataSourceLayout.addLayout(tableButtonLayout)
        dataSourceGroup.setLayout(dataSourceLayout)
        
        columnGroup = QGroupBox("データ選択")
        columnLayout = QGridLayout()
        
        usageGuideLabel = QLabel("【データ選択ガイド】\n"
                              "■ CSVファイル/Excelファイル: セル範囲で直接指定します\n"
                              "■ セル範囲の例: X軸「A2:A10」Y軸「B2:B10」または X軸「A2:E2」Y軸「A3:E3」\n"
                              "■ A1形式のセル指定で、列と行の範囲を指定してください")
        usageGuideLabel.setStyleSheet(
            "background-color: #546e7a; " +  # ブルーグレー（中間色調）
            "color: #ffffff; " +             # 白テキスト（どちらのモードでも視認性良好）
            "padding: 10px; " + 
            "border: 2px solid #90a4ae; " +  # 明るめのブルーグレー境界線
            "border-radius: 6px; " +
            "font-weight: bold; " +          # テキストを太字に
            "margin: 5px;"                   # 周囲に余白を追加
        )
        usageGuideLabel.setWordWrap(True)
        columnLayout.addWidget(usageGuideLabel, 0, 0, 1, 3)
        
        cellRangeSectionLabel = QLabel("セル範囲指定:")
        cellRangeSectionLabel.setStyleSheet("font-weight: bold;")
        columnLayout.addWidget(cellRangeSectionLabel, 1, 0, 1, 3)
        
        xRangeLabel = QLabel('X軸セル範囲:')
        self.xRangeEntry = QLineEdit()
        self.xRangeEntry.setPlaceholderText("例: A2:A10")
        columnLayout.addWidget(xRangeLabel, 2, 0)
        columnLayout.addWidget(self.xRangeEntry, 2, 1, 1, 2)
        
        yRangeLabel = QLabel('Y軸セル範囲:')
        self.yRangeEntry = QLineEdit()
        self.yRangeEntry.setPlaceholderText("例: B2:B10")
        columnLayout.addWidget(yRangeLabel, 3, 0)
        columnLayout.addWidget(self.yRangeEntry, 3, 1, 1, 2)
        
        columnHelpLabel = QLabel('※ セル範囲は「始点:終点」の形式で指定してください。\n　　例: 「A2:E2」は A2セルから E2セルまで、「A2:A10」は A2セルから A10セルまで')
        columnHelpLabel.setStyleSheet('color: gray; font-style: italic;')
        columnHelpLabel.setWordWrap(True)
        columnLayout.addWidget(columnHelpLabel, 4, 0, 1, 3)

        
        columnGroup.setLayout(columnLayout)
        
        dataActionGroup = QGroupBox("データ確定")
        dataActionLayout = QVBoxLayout()
        
        actionNoteLabel = QLabel("※ CSVファイル、Excelファイル、手入力のいずれの方法でも、データを保存するには下のボタンを押してください")
        actionNoteLabel.setStyleSheet("color: #cc0000; font-weight: bold;")
        actionNoteLabel.setWordWrap(True)
        dataActionLayout.addWidget(actionNoteLabel)
        
        # ホバーテキスト
        self.loadDataButton = QPushButton('データを確定・保存')
        self.loadDataButton.setToolTip('入力したデータを現在のデータセットに保存します')
        self.loadDataButton.clicked.connect(self.load_data)
        self.loadDataButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        dataActionLayout.addWidget(self.loadDataButton)
        
        dataActionGroup.setLayout(dataActionLayout)
        
        # 実測値コンテナに追加
        measuredLayout.addWidget(dataSourceGroup)
        measuredLayout.addWidget(columnGroup)
        measuredLayout.addWidget(dataActionGroup) 
        
        #* ======================数式============================================ 
        self.formulaContainer = QWidget()
        """数式固有コンテナ"""
        formulaLayout = QVBoxLayout(self.formulaContainer)
        """数式固有レイアウト"""
        
        formulaFormGroup = QGroupBox("数式入力")
        formulaFormLayout = QVBoxLayout()
        
        equationLayout = QHBoxLayout()
        equationLabel = QLabel('数式 (x変数を使用):')
        self.equationEntry = QLineEdit('x^2')
        self.equationEntry.setPlaceholderText('例: x^2, sin(x), exp(-x/2)')
        equationLayout.addWidget(equationLabel)
        equationLayout.addWidget(self.equationEntry)
        formulaFormLayout.addLayout(equationLayout)
        
        domainLayout = QHBoxLayout()
        domainLabel = QLabel('x軸範囲:')
        self.domainMinSpin = QDoubleSpinBox()
        self.domainMinSpin.setRange(-10000, 10000)
        self.domainMinSpin.setValue(0)
        self.domainMaxSpin = QDoubleSpinBox()
        self.domainMaxSpin.setRange(-10000, 10000)
        self.domainMaxSpin.setValue(1)
        domainLayout.addWidget(domainLabel)
        domainLayout.addWidget(self.domainMinSpin)
        domainLayout.addWidget(QLabel('〜'))
        domainLayout.addWidget(self.domainMaxSpin)
        formulaFormLayout.addLayout(domainLayout)
        
        samplesLayout = QHBoxLayout()
        samplesLabel = QLabel('プロットの計算数:')
        self.samplesSpin = QSpinBox()
        self.samplesSpin.setRange(1, 10000)
        self.samplesSpin.setValue(200)
        samplesLayout.addWidget(samplesLabel)
        samplesLayout.addWidget(self.samplesSpin)
        formulaFormLayout.addLayout(samplesLayout)

        samplesSupplementLabel = QLabel('＊計算数は実行速度に影響します\n単純な関数であれば50~200程度を推奨\n複雑な関数であれば200~1000程度を推奨')
        samplesSupplementLabel.setStyleSheet("color: gray; font-style: italic; padding-left: 20px; min-height: 50px;")
        formulaFormLayout.addWidget(samplesSupplementLabel)

        formulaInfoLabel = QLabel(
            '※ <span style="color:black;">掛け算は必ず <b><span style="color:red;">*</span></b> を明示してください（例: 2*x, (x+1)*y）</span><br>'
            '数式内では以下などの関数と演算が使用可能です(詳しくは下記の関数リストを参照)：<br>'
            'sin, cos, tan, exp, ln, log, sqrt, ^（累乗）, +, -, *, /'
        )
        formulaInfoLabel.setStyleSheet(
            'background-color: #fffbe6; color: black; border: 1px solid #ffe082; border-radius: 4px; padding: 8px;'
        )
        formulaFormLayout.addWidget(formulaInfoLabel)
        
        formulaFormGroup.setLayout(formulaFormLayout)
        formulaLayout.addWidget(formulaFormGroup)
        
        formulaOptionsGroup = QGroupBox("数式オプション設定")
        formulaOptionsLayout = QGridLayout()
        
        self.showTangentCheck = QCheckBox('特定点における接線を表示')
        self.showTangentCheck.setChecked(False)
        formulaOptionsLayout.addWidget(self.showTangentCheck, 0, 0, 1, 2)
        
        tangentXLabel = QLabel('接線のx座標:')
        self.tangentXSpin = QDoubleSpinBox()
        self.tangentXSpin.setRange(-10000, 10000)
        self.tangentXSpin.setValue(1)  
        formulaOptionsLayout.addWidget(tangentXLabel, 1, 0)
        formulaOptionsLayout.addWidget(self.tangentXSpin, 1, 1)
        
        tangentLengthLabel = QLabel('接線の長さ(0.1から設定可能):')
        self.tangentLengthSpin = QDoubleSpinBox()
        self.tangentLengthSpin.setRange(0.1, 10000)
        self.tangentLengthSpin.setValue(2.0)  
        formulaOptionsLayout.addWidget(tangentLengthLabel, 2, 0)
        formulaOptionsLayout.addWidget(self.tangentLengthSpin, 2, 1)
        
        tangentColorLabel = QLabel('接線の色:')
        self.tangentColorButton = QPushButton()
        self.tangentColorButton.setStyleSheet('background-color: red;')
        self.tangentColor = QColor('red')
        self.tangentColorButton.clicked.connect(self.select_tangent_color)
        formulaOptionsLayout.addWidget(tangentColorLabel, 3, 0)
        formulaOptionsLayout.addWidget(self.tangentColorButton, 3, 1)
        
        tangentStyleLabel = QLabel('接線のスタイル:')
        self.tangentStyleCombo = QComboBox()
        self.tangentStyleCombo.addItems(['実線', '点線', '破線', '一点鎖線'])
        formulaOptionsLayout.addWidget(tangentStyleLabel, 4, 0)
        formulaOptionsLayout.addWidget(self.tangentStyleCombo, 4, 1)
        
        self.showTangentEquationCheck = QCheckBox('接線の方程式を表示')
        self.showTangentEquationCheck.setChecked(False)
        self.showTangentEquationCheck.setToolTip('グラフ上に接線の方程式（y = ax + b）を表示します')
        formulaOptionsLayout.addWidget(self.showTangentEquationCheck, 5, 0, 1, 2)
        
        formulaOptionsGroup.setLayout(formulaOptionsLayout)
        formulaLayout.addWidget(formulaOptionsGroup)
        
        tikzGuideGroup = QGroupBox("TikZ数式ガイド")
        tikzGuideLayout = QVBoxLayout()
        
        guideLabel = QLabel('TikZで使える特殊な数学関数:')
        guideLabel.setStyleSheet("font-weight: bold;")
        tikzGuideLayout.addWidget(guideLabel)
        
        self.tikzFunctionsTable = QTableWidget(0, 2)  # 行数は動的に設定
        self.tikzFunctionsTable.setHorizontalHeaderLabels(['関数名', '説明と使用例'])
        self.tikzFunctionsTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tikzFunctionsTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tikzFunctionsTable.verticalHeader().setVisible(False)  
        self.tikzFunctionsTable.setEditTriggers(QTableWidget.NoEditTriggers)  
        
        tikz_functions = [
            ("add(x,y)", "x+y(x+yでも可)"),
            ("subtract(x,y)", "x-y(x-yでも可)"),
            ("multiply(x,y)", "x×y(x*yでも可)"),
            ("divide(x,y)", "x÷y(x/yでも可)"),
            ("sin(x)", "正弦関数:(度数法で入力)"),
            ("cos(x)", "余弦関数:(度数法で入力)"),
            ("tan(x)", "正接関数:(度数法で入力)"),
            ("pi", "円周率: (3.14159...)"),
            ("e", "自然対数の底:（2.71828...）"),
            ("exp(x)", "指数関数: exp(x)（eの累乗）"),
            ("ln(x)", "自然対数: ln(x)（底はe）"),
            ("log10(x)", "常用対数: log10(x)（底は10）"),
            ("log2(x)", "二進対数: log2(x)（底は2）"),
            ("sqrt(x)", "平方根: xの平方根"),
            ("pow(x,y)", "累乗: x^y"),
            ("factorial(x)", "階乗: x!（xは整数）"),
            ("abs(x)", "絶対値: xの絶対値"),
            ("asin(x)", "逆正弦関数(度数法で出力)"),
            ("acos(x)", "逆余弦関数(度数法で出力)"),
            ("atan(x)", "逆正接関数(度数法で出力)"),
            ("atan2(y,x)", "2引数の逆正接関数(度数法で出力)"),
            ("sinh(x)", "双曲線正弦関数: sinh(x)"),
            ("cosh(x)", "双曲線余弦関数: cosh(x)"),
            ("tanh(x)", "双曲線正接関数: tanh(x)"),
            ("min(x,y)", "xとyの小さい方"),
            ("max(x,y)", "xとyの大きい方"),
            ("deg(x)", "ラジアンから度に変換"),
            ("rad(x)", "度からラジアンに変換"),
            ("mod(x,y)", "剰余:xをyで割った余り"),
            ("floor(x)", "床関数: xを超えない最大の整数"),
            ("ceil(x)", "天井関数: x以上の最小の整数"),
            ("round(x)", "四捨五入: 小数点以下を四捨五入"),
            ("equal(x,y)", "等号: x=yなら1, そうでないなら0"),
            ("greater(x,y)", "不等号: x>yなら1, そうでないなら0"),
            ("less(x,y)", "不等号: x<yなら1, そうでないなら0"),
            ("rand", "乱数: 0から1の乱数を生成"),
        ]
        
        self.tikzFunctionsTable.setRowCount(len(tikz_functions))
        for i, (func, desc) in enumerate(tikz_functions):
            self.tikzFunctionsTable.setItem(i, 0, QTableWidgetItem(func))
            self.tikzFunctionsTable.setItem(i, 1, QTableWidgetItem(desc))
            
        table_height = self.tikzFunctionsTable.horizontalHeader().height()
        for i in range(len(tikz_functions)):
            table_height += self.tikzFunctionsTable.rowHeight(i)
        self.tikzFunctionsTable.setMinimumHeight(min(250, table_height))  # 最大高さを制限
        
        tikzGuideLayout.addWidget(self.tikzFunctionsTable)
        
        tableHelpLabel = QLabel('※ 関数をダブルクリックすると数式に挿入できます')
        tableHelpLabel.setStyleSheet("color: #666666; font-style: italic;")
        tikzGuideLayout.addWidget(tableHelpLabel)
        
        # 関数をダブルクリックで挿入
        self.tikzFunctionsTable.cellDoubleClicked.connect(self.insert_function_from_table)
        
        tikzGuideGroup.setLayout(tikzGuideLayout)
        formulaLayout.addWidget(tikzGuideGroup)
        
        formulaDataActionGroup = QGroupBox("データ確定")
        formulaDataActionLayout = QVBoxLayout()
        
        formulaActionNoteLabel = QLabel("※ 数式に基づくグラフを生成するには、下のボタンを押して数式データを保存してください")
        formulaActionNoteLabel.setStyleSheet("color: #cc0000; font-weight: bold;")
        formulaActionNoteLabel.setWordWrap(True)
        formulaDataActionLayout.addWidget(formulaActionNoteLabel)
        
        self.formulaDataButton = QPushButton('数式データを確定・保存')
        self.formulaDataButton.setToolTip('入力した数式に基づいてデータを生成し、現在のデータセットに保存します')
        self.formulaDataButton.clicked.connect(self.apply_formula)
        self.formulaDataButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        formulaDataActionLayout.addWidget(self.formulaDataButton)
        
        formulaDataActionGroup.setLayout(formulaDataActionLayout)
        formulaLayout.addWidget(formulaDataActionGroup)
        
        # 初期状態では実測値データコンテナを表示
        dataTabLayout.addWidget(self.measuredContainer)
        dataTabLayout.addWidget(self.formulaContainer)
        self.measuredContainer.setVisible(True)
        self.formulaContainer.setVisible(False)
        
        # タブに設定
        dataTab.setLayout(dataTabLayout)
        
        #* ======================グラフ設定タブ======================================= 
        plotTab = QWidget()
        plotTabLayout = QVBoxLayout()
        
        #* ======================個別設定============================================ 
        plotTypeGroup = QGroupBox("データセット個別設定 - グラフタイプ")
        plotTypeLayout = QHBoxLayout()
        
        self.lineRadio = QRadioButton("線グラフ")
        self.lineRadio.setChecked(True)
        self.scatterRadio = QRadioButton("散布図")
        self.lineScatterRadio = QRadioButton("線と点")
        self.barRadio = QRadioButton("棒グラフ")
        
        plotTypeLayout.addWidget(self.lineRadio)
        plotTypeLayout.addWidget(self.scatterRadio)
        plotTypeLayout.addWidget(self.lineScatterRadio)
        plotTypeLayout.addWidget(self.barRadio)
        plotTypeGroup.setLayout(plotTypeLayout)
        
        styleGroup = QGroupBox("データセット個別設定 - スタイル")
        styleLayout = QGridLayout()
        
        dataSourceTypeLabel = QLabel("選択中のデータセットタイプ:")
        self.dataSourceTypeDisplayLabel = QLabel("実測データ")
        self.dataSourceTypeDisplayLabel.setStyleSheet("font-weight: bold;")
        styleLayout.addWidget(dataSourceTypeLabel, 0, 0)
        styleLayout.addWidget(self.dataSourceTypeDisplayLabel, 0, 1)
        
        colorLabel = QLabel('線/点の色:')
        self.colorButton = QPushButton()
        self.colorButton.setStyleSheet('background-color: blue;')
        self.currentColor = QColor('blue')
        self.colorButton.clicked.connect(self.select_color)
        styleLayout.addWidget(colorLabel, 1, 0)
        styleLayout.addWidget(self.colorButton, 1, 1)
        
        lineWidthLabel = QLabel('線の太さ:')
        self.lineWidthSpin = QDoubleSpinBox()
        self.lineWidthSpin.setRange(0.1, 5.0)
        self.lineWidthSpin.setSingleStep(0.1)
        self.lineWidthSpin.setValue(1.0)
        styleLayout.addWidget(lineWidthLabel, 2, 0)
        styleLayout.addWidget(self.lineWidthSpin, 2, 1)
        
        markerLabel = QLabel('マーカースタイル:')
        self.markerCombo = QComboBox()
        self.markerCombo.addItems(['*', 'o', 'square', 'triangle', 'diamond', '+', 'x'])
        styleLayout.addWidget(markerLabel, 3, 0)
        styleLayout.addWidget(self.markerCombo, 3, 1)
        
        markerSizeLabel = QLabel('マーカーサイズ:')
        self.markerSizeSpin = QDoubleSpinBox()
        self.markerSizeSpin.setRange(0.5, 10.0)
        self.markerSizeSpin.setSingleStep(0.5)
        self.markerSizeSpin.setValue(3.0)
        styleLayout.addWidget(markerSizeLabel, 4, 0)
        styleLayout.addWidget(self.markerSizeSpin, 4, 1)
        
        styleGroup.setLayout(styleLayout)

        legendLabelGroup = QGroupBox("データセット個別設定 - 凡例ラベル")
        legendLabelLayout = QFormLayout()
        
        self.legendLabel = QLineEdit('データ')
        self.legendLabel.setReadOnly(True)
        legendLabelLayout.addRow('凡例ラベル(データセット名):', self.legendLabel)
        
        legendLabelGroup.setLayout(legendLabelLayout)
        
        #* ======================全体設定============================================ 
        axisGroup = QGroupBox("グラフ全体設定 - 軸")
        axisLayout = QGridLayout()
        
        xLabelLabel = QLabel('X軸ラベル:')
        self.xLabelEntry = QLineEdit(self.global_settings['x_label'])
        axisLayout.addWidget(xLabelLabel, 0, 0)
        axisLayout.addWidget(self.xLabelEntry, 0, 1)
        
        xTickStepLabel = QLabel('X軸目盛り間隔:')
        self.xTickStepSpin = QDoubleSpinBox()
        self.xTickStepSpin.setRange(0.01, 10000)
        self.xTickStepSpin.setSingleStep(1.0)
        self.xTickStepSpin.setValue(1.0)
        axisLayout.addWidget(xTickStepLabel, 0, 2)
        axisLayout.addWidget(self.xTickStepSpin, 0, 3)
        
        yLabelLabel = QLabel('Y軸ラベル:')
        self.yLabelEntry = QLineEdit(self.global_settings['y_label'])
        axisLayout.addWidget(yLabelLabel, 1, 0)
        axisLayout.addWidget(self.yLabelEntry, 1, 1)
        
        yTickStepLabel = QLabel('Y軸目盛り間隔:')
        self.yTickStepSpin = QDoubleSpinBox()
        self.yTickStepSpin.setRange(0.01, 10000)
        self.yTickStepSpin.setSingleStep(1.0)
        self.yTickStepSpin.setValue(1.0)
        axisLayout.addWidget(yTickStepLabel, 1, 2)
        axisLayout.addWidget(self.yTickStepSpin, 1, 3)

        xRangeLabel = QLabel('X軸範囲:')
        xRangeLayout = QHBoxLayout()
        self.xMinSpin = QDoubleSpinBox()
        self.xMinSpin.setRange(-10000, 10000)
        self.xMinSpin.setValue(self.global_settings['x_min'])
        self.xMinSpin.setFixedWidth(100)
        xRangeLayout.addWidget(self.xMinSpin)
        xRangeLayout.addStretch()
        xRangeLayout.addWidget(QLabel('〜'))
        xRangeLayout.addStretch() 
        self.xMaxSpin = QDoubleSpinBox()
        self.xMaxSpin.setRange(-10000, 10000)
        self.xMaxSpin.setValue(self.global_settings['x_max'])
        self.xMaxSpin.setFixedWidth(100)
        xRangeLayout.addWidget(self.xMaxSpin)
        axisLayout.addWidget(xRangeLabel, 2, 0)
        axisLayout.addLayout(xRangeLayout, 2, 1)
        
        yRangeLabel = QLabel('Y軸範囲:')
        yRangeLayout = QHBoxLayout()
        self.yMinSpin = QDoubleSpinBox()
        self.yMinSpin.setRange(-10000, 10000)
        self.yMinSpin.setValue(self.global_settings['y_min'])
        self.yMinSpin.setFixedWidth(100)
        yRangeLayout.addWidget(self.yMinSpin)
        yRangeLayout.addStretch()
        yRangeLayout.addWidget(QLabel('〜'))
        yRangeLayout.addStretch()
        self.yMaxSpin = QDoubleSpinBox()
        self.yMaxSpin.setRange(-10000, 10000)
        self.yMaxSpin.setValue(self.global_settings['y_max'])
        self.yMaxSpin.setFixedWidth(100)
        yRangeLayout.addWidget(self.yMaxSpin)
        axisLayout.addWidget(yRangeLabel, 3, 0)
        axisLayout.addLayout(yRangeLayout, 3, 1)

        self.x_tick_step = 1.0
        self.y_tick_step = 1.0

        self.gridCheck = QCheckBox('グリッド表示')
        self.gridCheck.setChecked(self.global_settings['grid'])
        axisLayout.addWidget(self.gridCheck, 4, 0, 1, 2)
        
        self.legendCheck = QCheckBox('凡例を表示')
        self.legendCheck.setChecked(self.global_settings.get('show_legend', True))
        
        legendPosLabel = QLabel('凡例の位置:')
        self.legendPosCombo = QComboBox()
        self.legendPosCombo.addItems(['左上', '右上', '左下', '右下'])
        # なんか英語の表記と実際の位置が一致しないのでしかたなくマッピングしないといけないぽい
        self.legend_pos_mapping = {
            '左上': 'north west',
            '右上': 'north east',
            '左下': 'south west',
            '右下': 'south east'
        }
        # キーと値の入れ替え
        reverse_mapping = {v: k for k, v in self.legend_pos_mapping.items()}
        current_pos = self.global_settings['legend_pos']
        if current_pos in reverse_mapping:
            self.legendPosCombo.setCurrentText(reverse_mapping[current_pos])
        else:
            self.legendPosCombo.setCurrentText('右上')
        
        axisLayout.addWidget(self.legendCheck, 5, 0, 1, 2)
        
        axisLayout.addWidget(legendPosLabel, 6, 0)
        axisLayout.addWidget(self.legendPosCombo, 6, 1)
        
        scaleTypeLabel = QLabel('軸の目盛りタイプ:')
        scaleTypeLayout = QHBoxLayout()
        self.scaleTypeGroup = QButtonGroup(self)
        
        self.normalScaleRadio = QRadioButton('通常')
        self.logXScaleRadio = QRadioButton('X軸対数')
        self.logYScaleRadio = QRadioButton('Y軸対数')
        self.logLogScaleRadio = QRadioButton('両軸対数')
        
        # グループ設定しているのでオンオフ連動する
        self.scaleTypeGroup.addButton(self.normalScaleRadio)
        self.scaleTypeGroup.addButton(self.logXScaleRadio)
        self.scaleTypeGroup.addButton(self.logYScaleRadio)
        self.scaleTypeGroup.addButton(self.logLogScaleRadio)
        
        scale_type = self.global_settings.get('scale_type', 'normal')
        if scale_type == 'logx':
            self.logXScaleRadio.setChecked(True)
        elif scale_type == 'logy':
            self.logYScaleRadio.setChecked(True)
        elif scale_type == 'loglog':
            self.logLogScaleRadio.setChecked(True)
        else:
            self.normalScaleRadio.setChecked(True)
        
        scaleTypeLayout.addWidget(self.normalScaleRadio)
        scaleTypeLayout.addWidget(self.logXScaleRadio)
        scaleTypeLayout.addWidget(self.logYScaleRadio)
        scaleTypeLayout.addWidget(self.logLogScaleRadio)
        
        axisLayout.addWidget(scaleTypeLabel, 7, 0)
        axisLayout.addLayout(scaleTypeLayout, 7, 1)
        
        axisGroup.setLayout(axisLayout)
        
        figureGroup = QGroupBox("グラフ全体設定 - 図")
        figureLayout = QFormLayout()
        
        self.captionEntry = QLineEdit(self.global_settings['caption'])
        
        self.labelEntry = QLineEdit(self.global_settings['label'])
        
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        self.positionCombo.setCurrentText(self.global_settings['position'])
        
        widthLayout = QHBoxLayout()
        self.widthSpin = QDoubleSpinBox()
        self.widthSpin.setRange(0.1, 1.0)
        self.widthSpin.setSingleStep(0.1)
        self.widthSpin.setValue(self.global_settings['width'])
        self.widthSpin.setSuffix('\\textwidth')
        
        heightLayout = QHBoxLayout()
        self.heightSpin = QDoubleSpinBox()
        self.heightSpin.setRange(0.1, 1.0)
        self.heightSpin.setSingleStep(0.1)
        self.heightSpin.setValue(self.global_settings['height'])
        self.heightSpin.setSuffix('\\textwidth')
        
        figureLayout.addRow('キャプション:', self.captionEntry)
        figureLayout.addRow('ラベル:', self.labelEntry)
        figureLayout.addRow('位置:', self.positionCombo)
        figureLayout.addRow('幅:', self.widthSpin)
        figureLayout.addRow('高さ:', self.heightSpin)
        
        figureGroup.setLayout(figureLayout)
        
        #* 全体設定レイアウト
        plotTabLayout.addWidget(QLabel("【グラフ全体の設定】"))
        plotTabLayout.addWidget(axisGroup)
        plotTabLayout.addWidget(figureGroup)
        
        # 区切り線
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        plotTabLayout.addWidget(separator)
        
        #* 個別設定レイアウト
        plotTabLayout.addWidget(QLabel("【データセット個別の設定】"))
        plotTabLayout.addWidget(plotTypeGroup)
        plotTabLayout.addWidget(styleGroup)
        plotTabLayout.addWidget(legendLabelGroup)
        
        plotTab.setLayout(plotTabLayout)
        
        #* ======================特殊点・注釈タブ============================================ 
        annotationTab = QWidget()
        annotationTabLayout = QVBoxLayout()
        
        specialPointsGroup = QGroupBox("特殊点")
        specialPointsLayout = QVBoxLayout()
        
        self.specialPointsCheck = QCheckBox('特殊点を表示')
        specialPointsLayout.addWidget(self.specialPointsCheck)
        
        self.specialPointsTable = QTableWidget(0, 4)  
        self.specialPointsTable.setHorizontalHeaderLabels(['X', 'Y', '色', '座標表示'])
        self.specialPointsTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.specialPointsTable.setEnabled(False)
        specialPointsLayout.addWidget(self.specialPointsTable)
        
        spButtonLayout = QHBoxLayout()
        addSpButton = QPushButton('特殊点を追加')
        addSpButton.clicked.connect(self.add_special_point)
        removeSpButton = QPushButton('選択した特殊点を削除')
        removeSpButton.clicked.connect(self.remove_special_point)
        assignToDatasetBtn = QPushButton('データセットに割り当て')
        assignToDatasetBtn.clicked.connect(self.assign_special_points_to_dataset)
        assignToDatasetBtn.setStyleSheet('background-color: #007BFF; border-radius: 3px; padding-top: 1px; padding-bottom: 1px;')
        spButtonLayout.addWidget(addSpButton)
        spButtonLayout.addWidget(removeSpButton)
        spButtonLayout.addWidget(assignToDatasetBtn)
        specialPointsLayout.addLayout(spButtonLayout)
        
        #　特殊点Widgetオンオフ連動
        self.specialPointsCheck.toggled.connect(lambda checked: [
            self.specialPointsTable.setEnabled(checked),
            addSpButton.setEnabled(checked),
            removeSpButton.setEnabled(checked),
            assignToDatasetBtn.setEnabled(checked)
        ])
        
        specialPointsGroup.setLayout(specialPointsLayout)
        
        annotationsGroup = QGroupBox("注釈")
        annotationsLayout = QVBoxLayout()
        
        self.annotationsCheck = QCheckBox('注釈を表示')
        annotationsLayout.addWidget(self.annotationsCheck)
        
        self.annotationsTable = QTableWidget(0, 5) 
        self.annotationsTable.setHorizontalHeaderLabels(['X', 'Y', 'テキスト', '色', '位置'])
        self.annotationsTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.annotationsTable.setEnabled(False)
        annotationsLayout.addWidget(self.annotationsTable)
        
        annButtonLayout = QHBoxLayout()
        addAnnButton = QPushButton('注釈を追加')
        addAnnButton.clicked.connect(self.add_annotation)
        removeAnnButton = QPushButton('選択した注釈を削除')
        removeAnnButton.clicked.connect(self.remove_annotation)
        assignAnnToDatasetBtn = QPushButton('データセットに割り当て')
        assignAnnToDatasetBtn.clicked.connect(self.assign_annotations_to_dataset)
        assignAnnToDatasetBtn.setStyleSheet('background-color: #007BFF; border-radius: 3px; padding-top: 1px; padding-bottom: 1px;')
        annButtonLayout.addWidget(addAnnButton)
        annButtonLayout.addWidget(removeAnnButton)
        annButtonLayout.addWidget(assignAnnToDatasetBtn)
        annotationsLayout.addLayout(annButtonLayout)
        
        # 特異点Widgetオンオフ連動
        self.annotationsCheck.toggled.connect(lambda checked: [
            self.annotationsTable.setEnabled(checked),
            addAnnButton.setEnabled(checked),
            removeAnnButton.setEnabled(checked),
            assignAnnToDatasetBtn.setEnabled(checked)
        ])
        
        annotationsGroup.setLayout(annotationsLayout)
        
        annotationTabLayout.addWidget(specialPointsGroup)
        annotationTabLayout.addWidget(annotationsGroup)
        annotationTab.setLayout(annotationTabLayout)
        
        #* ======================タブ追加============================================ 
        tabWidget.addTab(dataTab, "データ入力")
        tabWidget.addTab(plotTab, "グラフ設定")
        tabWidget.addTab(annotationTab, "特殊点・注釈設定")
        
        # 現在のデータセットの値を変更したときに現在のデータセットの値を保存
        self.lineWidthSpin.valueChanged.connect(self.update_current_dataset)
        self.markerCombo.currentIndexChanged.connect(self.update_current_dataset)
        self.markerSizeSpin.valueChanged.connect(self.update_current_dataset)
        self.colorButton.clicked.connect(self.update_current_dataset)
        self.lineRadio.toggled.connect(self.update_current_dataset)
        self.scatterRadio.toggled.connect(self.update_current_dataset)
        self.lineScatterRadio.toggled.connect(self.update_current_dataset)
        self.barRadio.toggled.connect(self.update_current_dataset)

        # タブ切り替え時にも保存
        tabWidget.currentChanged.connect(lambda _: self.update_current_dataset())
        
        #* ======================Tikz_plot_Layout============================================ 
        settingsLayout.addLayout(infoLayout)
        settingsLayout.addWidget(datasetGroup)
        settingsLayout.addWidget(tabWidget)

        self.dataTable.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        convertButton = QPushButton("LaTeXコードに変換")
        convertButton.clicked.connect(self.convert_to_tikz)
        convertButton.setStyleSheet('background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;')
        convertButton.setFixedHeight(32)
        settingsLayout.addWidget(convertButton)
        
        # --- 下部：結果表示部分 ---
        resultWidget = QWidget()
        resultLayout = QVBoxLayout(resultWidget)
        
        resultLabel = QLabel("LaTeX コード:")
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(100)  # 最小高さを設定
        self.resultText.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # 縦方向も拡大可能に
        
        copyButton = QPushButton("クリップボードにコピー")
        copyButton.setStyleSheet("background-color: #007BFF; color: white; font-size: 16px; padding: 12px; font-weight: 900; border-radius: 8px; text-transform: uppercase; letter-spacing: 1px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);")
        copyButton.clicked.connect(self.copy_to_clipboard)
        
        resultLayout.addWidget(resultLabel)
        resultLayout.addWidget(self.resultText)
        resultLayout.addWidget(copyButton)
        
        #* ======================main_Container============================================ 
        splitter.addWidget(settingsWidget)
        splitter.addWidget(resultWidget)
        splitter.setSizes([600, 200])  
        
        mainLayout.addWidget(splitter)
        
        self.statusBar = QStatusBar()
        mainLayout.addWidget(self.statusBar)
        
        mainWidget.setLayout(mainLayout)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(mainWidget)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        self.setLayout(QVBoxLayout())
        self.layout().addWidget(scroll)
        
        scrollHint = QLabel('画面に収まらない場合はスクロールバーで全体を表示できます')
        scrollHint.setStyleSheet('background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fffbe6, stop:1 #ffe082); color: #d32f2f; font-size: 13px; font-weight: bold; border: 1px solid #ffe082; border-radius: 4px; padding: 6px;')
        scrollHint.setAlignment(Qt.AlignHCenter)
        mainLayout.insertWidget(0, scrollHint)
        
        # 初期状態のデータソース選択を反映
        self.toggle_source_fields()

        #* ======================処理============================================ 


    def browse_csv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "CSVファイルを選択", "", "CSV Files (*.csv);;All Files (*)")
        if file_path:
            self.fileEntry.setText(file_path)
            self.csvRadio.setChecked(True)
            self.toggle_source_fields()

    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        if file_path:
            self.excelEntry.setText(file_path)
            self.excelRadio.setChecked(True)
            self.toggle_source_fields()
            self.update_sheet_names(file_path)

    def update_sheet_names(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            self.sheetCombobox.clear()
            self.sheetCombobox.addItems(xls.sheet_names)
            self.statusBar.showMessage(f"ファイル '{os.path.basename(file_path)}' を読み込みました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"シート名の取得に失敗しました: {str(e)}")
            self.statusBar.showMessage("ファイル読み込みエラー")
            
    def toggle_source_fields(self, button=None):
        if self.csvRadio.isChecked():
            self.fileEntry.setEnabled(True)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(False)
            # セル範囲入力欄を有効化
            self.xRangeEntry.setEnabled(True)
            self.yRangeEntry.setEnabled(True)
            self.addRowButton.setEnabled(False)
            self.removeRowButton.setEnabled(False)
            # CSVファイル用のツールチップ
            self.loadDataButton.setToolTip("CSVファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.excelRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(True)
            self.sheetCombobox.setEnabled(True)
            self.dataTable.setEnabled(False)
            # セル範囲入力欄を有効化
            self.xRangeEntry.setEnabled(True)
            self.yRangeEntry.setEnabled(True)
            self.addRowButton.setEnabled(False)
            self.removeRowButton.setEnabled(False)
            # Excelファイル用のツールチップ
            self.loadDataButton.setToolTip("Excelファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.manualRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(True)
            # セル範囲入力欄を無効化
            self.xRangeEntry.setEnabled(False)
            self.yRangeEntry.setEnabled(False)
            self.addRowButton.setEnabled(True)
            self.removeRowButton.setEnabled(True)
            # 手入力データ用のツールチップ
            self.loadDataButton.setToolTip("テーブルに入力したデータを現在のデータセットに保存します")

    def load_data(self):
        """
        現在のデータセットから実測値データ入力タブ保存ボタン\n
        この場でdatasetsにdata_x, data_yを追加するので戻り値なし
        """
        try:
            if self.current_dataset_index < 0:
                QMessageBox.warning(self, "警告", "データを読み込むデータセットを選択してください")
                return
                
            data_x = []
            data_y = []
            
            x_range = self.xRangeEntry.text().strip()
            y_range = self.yRangeEntry.text().strip()
            
            if self.csvRadio.isChecked():
                file_path = self.fileEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なCSVファイルを選択してください")
                    return
                
                if not x_range or not y_range:
                    QMessageBox.warning(self, "警告", "X軸とY軸のセル範囲を指定してください")
                    return
                
                # CSVファイルの読み込み処理（以下略）...
                try:
                    # UTF-8
                    df = pd.read_csv(file_path, encoding='utf-8', sep=None, engine='python')
                except UnicodeDecodeError:
                    try:
                        # Shift-JIS
                        df = pd.read_csv(file_path, encoding='shift_jis', sep=None, engine='python')
                    except UnicodeDecodeError:
                        try:
                            # CP932
                            df = pd.read_csv(file_path, encoding='cp932', sep=None, engine='python')
                        except UnicodeDecodeError:
                            try:
                                # EUC-JP
                                df = pd.read_csv(file_path, encoding='euc_jp', sep=None, engine='python')
                            except UnicodeDecodeError:
                                # その他のエラー
                                QMessageBox.warning(self, "エンコーディングエラー", 
                                                  "CSVファイルのエンコーディングを自動判別できませんでした。\n"
                                                  "ファイルが破損しているか、サポートされていないエンコーディングの可能性があります。\n"
                                                  "UTF-8、Shift-JIS、CP932、EUC-JPのいずれかで保存し直してください。")
                                return
                except Exception as e:
                    QMessageBox.warning(self, "CSVファイル読み込みエラー", 
                                      f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}\n"
                                      "ファイル形式を確認してください。")
                    return
                
                try:
                    # A1:A10形式のセル範囲を解析してデータを取得
                    result = self.extract_data_from_range(df, x_range, y_range)
                    
                    # 警告文があるかどうか
                    if len(result) == 3:
                        data_x, data_y, warnings = result
                        if warnings:
                            safe_warnings = []
                            for warning in warnings:
                                if isinstance(warning, (list, tuple)):
                                    safe_warnings.append(str(warning))
                                else:
                                    safe_warnings.append(str(warning))
                            QMessageBox.information(self, "データ読み込み情報", 
                                              "データは正常に読み込まれました。\n参考情報:\n- " + 
                                              "\n- ".join(safe_warnings))
                    else:
                        data_x, data_y = result
                except Exception as e:
                    error_msg = str(e)
                    if "セル範囲" in error_msg:
                        QMessageBox.warning(self, "セル範囲エラー", f"{error_msg}\n\n正しいセル範囲の例:\n- X軸「A2:A10」Y軸「B2:B10」（同じ行数の異なる列）\n- X軸「A2:E2」Y軸「A3:E3」（同じ列数の異なる行）")
                    else:
                        QMessageBox.warning(self, "警告", f"CSVセル範囲からのデータ抽出中にエラーが発生しました: {error_msg}")
                    return
                
            elif self.excelRadio.isChecked():
                file_path = self.excelEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なExcelファイルを選択してください")
                    return
                
                sheet_name = self.sheetCombobox.currentText()
                if not sheet_name:
                    QMessageBox.warning(self, "警告", "シート名を選択してください")
                    return
                
                if not x_range or not y_range:
                    QMessageBox.warning(self, "警告", "X軸とY軸のセル範囲を指定してください")
                    return
                
                try:
                    result = self.extract_data_from_excel_range(file_path, sheet_name, x_range, y_range)
                    
                    # 警告文があるかどうか
                    if len(result) == 3:
                        data_x, data_y, warnings = result
                        if warnings:
                            safe_warnings = []
                            for warning in warnings:
                                if isinstance(warning, (list, tuple)):
                                    safe_warnings.append(str(warning))
                                else:
                                    safe_warnings.append(str(warning))
                            QMessageBox.information(self, "データ読み込み情報", 
                                              "データは正常に読み込まれました。\n参考情報:\n- " + 
                                              "\n- ".join(safe_warnings))
                    else:
                        data_x, data_y = result
                except Exception as e:
                    error_msg = str(e)
                    if "セル範囲" in error_msg or "有効なデータ" in error_msg:
                        # セル範囲関連のエラー
                        QMessageBox.warning(self, "Excelセル範囲エラー", f"{error_msg}\n\n正しいセル範囲の例:\n- X軸「A2:A10」Y軸「B2:B10」（同じ行数の異なる列）\n- X軸「A2:E2」Y軸「A3:E3」（同じ列数の異なる行）")
                    else:
                        # その他のエラー
                        QMessageBox.warning(self, "警告", f"Excelセル範囲からのデータ抽出中にエラーが発生しました: {error_msg}")
                    return
                
            elif self.manualRadio.isChecked():
                for row in range(self.dataTable.rowCount()):
                    x_item = self.dataTable.item(row, 0)
                    y_item = self.dataTable.item(row, 1)
                    
                    if x_item and y_item and x_item.text() and y_item.text():
                        try:
                            x_val = float(x_item.text())
                            y_val = float(y_item.text())
                            data_x.append(x_val)
                            data_y.append(y_val)
                        except ValueError:
                            pass  
                
                if not data_x:
                    QMessageBox.warning(self, "警告", "有効なデータポイントがありません")
                    return
            else:
                QMessageBox.warning(self, "警告", "データソース（CSV/Excel/手入力）を選択してください")
                return
            
            if len(data_x) != len(data_y):
                QMessageBox.warning(self, "警告", f"X軸とY軸のデータ長が一致しません (X: {len(data_x)}, Y: {len(data_y)})")
                return
                
            if not data_x:
                QMessageBox.warning(self, "警告", "有効なデータがありません")
                return
                
            self.datasets[self.current_dataset_index]['data_x'] = data_x
            self.datasets[self.current_dataset_index]['data_y'] = data_y
            
            if self.csvRadio.isChecked() or self.excelRadio.isChecked():
                self.datasets[self.current_dataset_index]['x_range'] = x_range
                self.datasets[self.current_dataset_index]['y_range'] = y_range
            
            if self.manualRadio.isChecked():
                self.update_data_table_from_dataset(self.datasets[self.current_dataset_index])
            
            dataset_name = self.datasets[self.current_dataset_index]['name']
            QMessageBox.information(self, "読み込み成功 ✓", f"データセット '{dataset_name}' に{len(data_x)}個のデータポイントを正常に読み込みました")
            self.statusBar.showMessage(f"データセット '{dataset_name}' にデータを読み込みました: {len(data_x)}ポイント")
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データ読み込み中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("データ読み込みエラー")

    def extract_data_from_range(self, df, x_range, y_range):
        """
        -> data_x, data_y, warnings\n
        - csvファイルから直接セル範囲のデータを抽出する\n
        - 有効なx,yが返る\n
        - イレギュラーがない場合戻り値2，警告がある場合戻り値3\n
        """
        
        def parse_range(range_str):
            """S行idx, S列idx, E行idx, E列idx"""
            try:
                parts = range_str.split(':')
                if len(parts) != 2:
                    raise ValueError(f"セル範囲の形式が正しくありません: {range_str} (例: A1:A10)")
                
                start_cell, end_cell = parts
                
                start_col = ''.join(c for c in start_cell if c.isalpha())
                start_row_idx = int(''.join(c for c in start_cell if c.isdigit())) - 1 
                
                end_col = ''.join(c for c in end_cell if c.isalpha())
                end_row_idx = int(''.join(c for c in end_cell if c.isdigit())) - 1  
                
                def col_to_index(col_str):
                    """
                    #* 26進数
                    UnicodeのA-Zを0-25に変換\n
                    BAの場合は 2 * 26^1 + 1 * 26^0 = 53\n
                    indexなので-1\n
                    """
                    index = 0
                    for i, char in enumerate(reversed(col_str)):
                        index += (ord(char.upper()) - ord('A') + 1) * (26 ** i)
                    return index - 1  
                
                start_col_idx = col_to_index(start_col)
                end_col_idx = col_to_index(end_col)
                
                return (start_row_idx, start_col_idx, end_row_idx, end_col_idx)
            except Exception as e:
                raise ValueError(f"セル範囲の解析中にエラーが発生しました ({range_str}): {str(e)}")
        
        try:
            x_start_row, x_start_col, x_end_row, x_end_col = parse_range(x_range)
            y_start_row, y_start_col, y_end_row, y_end_col = parse_range(y_range)
            
            data_x = []
            data_y = []
            warnings = []
            
            x_is_row = (x_start_row == x_end_row)  # 横方向
            x_is_column = (x_start_col == x_end_col)  # 縦方向
            
            y_is_row = (y_start_row == y_end_row)  # 横方向
            y_is_column = (y_start_col == y_end_col)  # 縦方向
            
            if not (x_is_row or x_is_column):
                raise ValueError("X軸のセル範囲は、1行の複数セルまたは1列の複数セルである必要があります")
            
            if x_is_row:
                # 横方向
                for col in range(x_start_col, x_end_col + 1):
                    if col < len(df.columns):
                        val = df.iloc[x_start_row, col]
                        data_x.append(val)
                    else:
                        warnings.append(f"指定されたX軸の列インデックス {col} はデータフレームの範囲外です")
            else:
                # 縦方向
                for row in range(x_start_row, x_end_row + 1):
                    if row < len(df):
                        val = df.iloc[row, x_start_col]
                        data_x.append(val)
                    else:
                        warnings.append(f"指定されたX軸の行インデックス {row} はデータフレームの範囲外です")
            
            if not (y_is_row or y_is_column):
                raise ValueError("Y軸のセル範囲は、1行の複数セルまたは1列の複数セルである必要があります")
            
            if y_is_row:
                # 横方向
                for col in range(y_start_col, y_end_col + 1):
                    if col < len(df.columns):
                        val = df.iloc[y_start_row, col]
                        data_y.append(val)
                    else:
                        warnings.append(f"指定されたY軸の列インデックス {col} はデータフレームの範囲外です")
            else:
                # 縦方向
                for row in range(y_start_row, y_end_row + 1):
                    if row < len(df):
                        val = df.iloc[row, y_start_col]
                        data_y.append(val)
                    else:
                        warnings.append(f"指定されたY軸の行インデックス {row} はデータフレームの範囲外です")
            
            if (x_is_row and y_is_column) or (x_is_column and y_is_row):
                # 縦横異なるとデータ数が一致しない可能性あり
                warnings.append(f"データの向きの情報:\n X軸は{'横方向（行に沿って）' if x_is_row else '縦方向（列に沿って）'}、Y軸は{'横方向（行に沿って）' if y_is_row else '縦方向（列に沿って）'}です．これは問題ありません．\n")
                
                # x_debug = f"X軸データ({len(data_x)}個): {str(data_x[:5])}{'...' if len(data_x) > 5 else ''}"
                # y_debug = f"Y軸データ({len(data_y)}個): {str(data_y[:5])}{'...' if len(data_y) > 5 else ''}"
                # warnings.append(x_debug)
                # warnings.append(y_debug)
                
            if len(data_x) != len(data_y):
                warnings.append(f"X軸({len(data_x)}個)とY軸({len(data_y)}個)のデータ数が一致しません")
                
                if len(data_x) < len(data_y):
                    data_y = data_y[:len(data_x)]
                    warnings.append(f"Y軸データを先頭から{len(data_x)}個使用します")
                else:
                    data_x = data_x[:len(data_y)]
                    warnings.append(f"X軸データを先頭から{len(data_y)}個使用します")
            
            min_len = min(len(data_x), len(data_y))
            if min_len == 0:
                raise ValueError("有効なデータがありません。セル範囲を確認してください。")
                
            if len(data_x) > min_len:
                warnings.append(f"データ数を調整しました: X軸のデータを{min_len}個に揃えました。")
                data_x = data_x[:min_len]
            elif len(data_y) > min_len:
                warnings.append(f"データ数を調整しました: Y軸のデータを{min_len}個に揃えました。")
                data_y = data_y[:min_len]
            
            #! pd使わずinstanceでやろうとするとなんかバグる…？
            data_x = pd.Series(data_x).apply(lambda x: float('nan') if pd.isna(x) else float(x)).tolist()
            data_y = pd.Series(data_y).apply(lambda y: float('nan') if pd.isna(y) else float(y)).tolist()
            
            valid_indices = [i for i, (x, y) in enumerate(zip(data_x, data_y)) 
                            if not (math.isnan(x) or math.isnan(y))]
            
            if not valid_indices:
                raise ValueError("有効なデータポイントがありません。セル範囲に数値データが含まれているか確認してください。")
                
            if len(valid_indices) < min_len:
                warnings.append(f"数値に変換できないデータが{min_len - len(valid_indices)}個あったため無視されました")
                
            data_x = [data_x[i] for i in valid_indices]
            data_y = [data_y[i] for i in valid_indices]
            
            return data_x, data_y, warnings if warnings else (data_x, data_y)
        except Exception as e:
            raise e

    def extract_data_from_excel_range(self, file_path, sheet_name, x_range, y_range):
        """
        -> data_x, data_y, warnings\n
        - Excelファイルから直接セル範囲のデータを抽出する\n
        - 有効なx,yが返る\n
        - イレギュラーがない場合戻り値2，警告がある場合戻り値3\n
        """
        
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb[sheet_name]
            
            try:
                x_cells = list(sheet[x_range])
            except Exception as e:
                raise ValueError(f"X軸のセル範囲「{x_range}」が無効です: {str(e)}")
                
            try:
                y_cells = list(sheet[y_range])
            except Exception as e:
                raise ValueError(f"Y軸のセル範囲「{y_range}」が無効です: {str(e)}")
            
            data_x = []
            data_y = []
            warnings = []
            
            #* 行ごとの配列なので1であれば1行で列方向(左右)
            x_is_row = len(x_cells) == 1 
            #*　全ての行の要素数が1であれば行方向(上下)
            x_is_column = all(len(row) == 1 for row in x_cells)  
            y_is_row = len(y_cells) == 1  
            y_is_column = all(len(row) == 1 for row in y_cells) 
            
            if not (x_is_row or x_is_column):
                raise ValueError("X軸のセル範囲は、1行の複数セルまたは1列の複数セルである必要があります\n")
            
            if x_is_row:
                data_x = [cell.value for cell in x_cells[0]]
            else:
                data_x = [row[0].value for row in x_cells]
            
            # yも同様
            if not (y_is_row or y_is_column):
                raise ValueError("Y軸のセル範囲は、1行の複数セルまたは1列の複数セルである必要があります\n")
            
            if y_is_row:
                data_y = [cell.value for cell in y_cells[0]]
            else:
                data_y = [row[0].value for row in y_cells]
            
            def safe_str_value(value):
                if value is None:
                    return "None"
                elif isinstance(value, (list, tuple)):
                    return f"リスト型データ({len(value)}個)"
                else:
                    return str(value)
            
            x_samples = [safe_str_value(x) for x in data_x[:3]]
            y_samples = [safe_str_value(y) for y in data_y[:3]]
            
            x_types = set(type(x).__name__ for x in data_x if x is not None)
            y_types = set(type(y).__name__ for y in data_y if y is not None)
            
            # x_type_info = ", ".join(x_types) if x_types else "None"
            # y_type_info = ", ".join(y_types) if y_types else "None"
            
            if (x_is_row and y_is_column) or (x_is_column and y_is_row):
                # 行列不一致
                warnings.append(f"データの向きの情報: X軸は{'横方向（行に沿って）' if x_is_row else '縦方向（列に沿って）'}、Y軸は{'横方向（行に沿って）' if y_is_row else '縦方向（列に沿って）'}です。これは問題ありません。\n")
                
                # warnings.append(f"X軸データタイプ: {x_type_info}, サンプル: {', '.join(x_samples)}\n")
                # warnings.append(f"Y軸データタイプ: {y_type_info}, サンプル: {', '.join(y_samples)}\n")
                
                # x,y軸個数不一致
                if len(data_x) != len(data_y):
                    warnings.append(f"X軸({len(data_x) - 1}個)とY軸({len(data_y) - 1}個)のデータ数を自動的に調整しました。\n")
                    
                    if len(data_x) < len(data_y):
                        data_y = data_y[:len(data_x)]
                        warnings.append(f"Y軸データは先頭から{len(data_x) - 1}個を使用します。\n")
                    else:
                        data_x = data_x[:len(data_y)]
                        warnings.append(f"X軸データは先頭から{len(data_y) - 1}個を使用します。\n")

            #* ===================processed_data_x=============================
            
            processed_data_x = []
            for x in data_x:
                try:
                    if x is None:
                        # Noneは欠損値として扱う
                        processed_data_x.append(float('nan'))
                    elif isinstance(x, (int, float)):
                        processed_data_x.append(float(x))
                    elif isinstance(x, str):
                        x = x.strip().replace(',', '')  
                        if x:
                            try:
                                processed_data_x.append(float(x))
                            except ValueError:
                                processed_data_x.append(float('nan'))
                                warnings.append(f"数値に変換できない文字列「{x}」を欠損値として処理しました\n")
                        else:
                            processed_data_x.append(float('nan'))
                    elif isinstance(x, (list, tuple)):# ROW関数等使われた時用
                        if len(x) > 0:
                            first_item = x[0]
                            if first_item is None:
                                processed_data_x.append(float('nan'))
                            elif isinstance(first_item, (int, float)):
                                processed_data_x.append(float(first_item))
                            elif isinstance(first_item, str):
                                try:
                                    processed_data_x.append(float(first_item.strip().replace(',', '')))
                                except ValueError:
                                    processed_data_x.append(float('nan'))
                                    warnings.append(f"複合データの数値変換に失敗しました: {safe_str_value(first_item)}")
                            else:
                                processed_data_x.append(float('nan'))
                                warnings.append(f"未対応の複合データ型: {type(first_item).__name__}")
                        else:
                            processed_data_x.append(float('nan'))
                    else:
                        try:
                            processed_data_x.append(float(str(x)))
                        except:
                            processed_data_x.append(float('nan'))
                            warnings.append(f"未対応のデータ型: {type(x).__name__}")
                except Exception as e:
                    processed_data_x.append(float('nan'))
                    warnings.append(f"データ処理中にエラー: {str(e)}")

            #* ===================processed_data_y=============================
            
            processed_data_y = []
            for y in data_y:
                try:
                    if y is None:
                        # Noneは欠損値として扱う
                        processed_data_y.append(float('nan'))
                    elif isinstance(y, (int, float)):
                        processed_data_y.append(float(y))
                    elif isinstance(y, str):
                        y = y.strip().replace(',', '')  
                        if y:
                            try:
                                processed_data_y.append(float(y))
                            except ValueError:
                                processed_data_y.append(float('nan'))
                                warnings.append(f"数値に変換できない文字列「{y}」を欠損値として処理しました\n")
                        else:
                            processed_data_y.append(float('nan'))
                    elif isinstance(y, (list, tuple)):# ROW関数等使われた時用
                        if len(y) > 0:
                            first_item = y[0]
                            if first_item is None:
                                processed_data_y.append(float('nan'))
                            elif isinstance(first_item, (int, float)):
                                processed_data_y.append(float(first_item))
                            elif isinstance(first_item, str):
                                try:
                                    processed_data_y.append(float(first_item.strip().replace(',', '')))
                                except ValueError:
                                    processed_data_y.append(float('nan'))
                                    warnings.append(f"複合データの数値変換に失敗しました: {safe_str_value(first_item)}")
                            else:
                                processed_data_y.append(float('nan'))
                                warnings.append(f"未対応の複合データ型: {type(first_item).__name__}")
                        else:
                            processed_data_y.append(float('nan'))
                    else:
                        try:
                            processed_data_y.append(float(str(y)))
                        except:
                            processed_data_y.append(float('nan'))
                            warnings.append(f"未対応のデータ型: {type(y).__name__}")
                except Exception as e:
                    processed_data_y.append(float('nan'))
                    warnings.append(f"データ処理中にエラー: {str(e)}")

            
            data_x = processed_data_x
            data_y = processed_data_y
            
            min_len = min(len(data_x), len(data_y))
            if min_len == 0:
                raise ValueError("有効なデータがありません。セル範囲を確認してください。\n")
                
            if len(data_x) > min_len:
                data_x = data_x[:min_len]
                warnings.append(f"データ数を調整しました: X軸のデータを{min_len}個に揃えました。\n")
            elif len(data_y) > min_len:
                data_y = data_y[:min_len]
                warnings.append(f"データ数を調整しました: Y軸のデータを{min_len}個に揃えました。\n")
            
            valid_indices = [i for i, (x, y) in enumerate(zip(data_x, data_y)) if not (math.isnan(x) or math.isnan(y))]
            
            if not valid_indices:
                raise ValueError("有効なデータポイントがありません。セル範囲に数値データが含まれているか確認してください。\n")
                
            if len(valid_indices) < min_len:
                warnings.append(f"数値に変換できないデータが{min_len - len(valid_indices)}個あったため無視されました\n")
                
            data_x = [data_x[i] for i in valid_indices]
            data_y = [data_y[i] for i in valid_indices]
            
            
            return data_x, data_y, warnings if warnings else (data_x, data_y)
            
        except Exception as e:
            raise ValueError(f"Excelセル範囲からのデータ抽出中にエラーが発生しました: {str(e)}")

    def add_special_point(self):
        row_position = self.specialPointsTable.rowCount()
        self.specialPointsTable.insertRow(row_position)
        
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        
        color_combo = QComboBox()
        color_combo.addItems(['red', 'blue', 'green', 'black', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("red")
        
        coord_display_combo = QComboBox()
        coord_display_combo.addItems([
            'なし', 
            'X座標のみ（線のみ）', 
            'X座標のみ（値も表示）', 
            'Y座標のみ（線のみ）', 
            'Y座標のみ（値も表示）', 
            'X,Y座標（線のみ）', 
            'X,Y座標（値も表示）'
        ])
        coord_display_combo.setCurrentText("なし")
        
        self.specialPointsTable.setItem(row_position, 0, x_item)
        self.specialPointsTable.setItem(row_position, 1, y_item)
        self.specialPointsTable.setCellWidget(row_position, 2, color_combo)
        self.specialPointsTable.setCellWidget(row_position, 3, coord_display_combo)

    def remove_special_point(self):
        selected_rows = set(index.row() for index in self.specialPointsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):# 後ろから削除
            self.specialPointsTable.removeRow(row)

    def add_annotation(self):
        row_position = self.annotationsTable.rowCount()
        self.annotationsTable.insertRow(row_position)
        
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        text_item = QTableWidgetItem("注釈テキスト")
        
        color_combo = QComboBox()
        color_combo.addItems(['black', 'red', 'blue', 'green', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("black")
        
        pos_combo = QComboBox()
        pos_combo.addItems(['上', '右上', '右', '右下', '下', '左下', '左', '左上'])
        pos_combo.setCurrentText("右上")
        
        self.annotationsTable.setItem(row_position, 0, x_item)
        self.annotationsTable.setItem(row_position, 1, y_item)
        self.annotationsTable.setItem(row_position, 2, text_item)
        self.annotationsTable.setCellWidget(row_position, 3, color_combo)
        self.annotationsTable.setCellWidget(row_position, 4, pos_combo)

    def remove_annotation(self):
        selected_rows = set(index.row() for index in self.annotationsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.annotationsTable.removeRow(row)
    
    def copy_to_clipboard(self):
        latex_code = self.resultText.toPlainText()
        if latex_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(latex_code)
            self.statusBar.showMessage("LaTeXコードをクリップボードにコピーしました", 3000)  # 3秒間表示

    def update_global_settings(self):
        """
        グラフの全体設定の更新\n
        LaTeX変換時に発火
        """
        try:
            self.global_settings['x_label'] = self.xLabelEntry.text()
            self.global_settings['y_label'] = self.yLabelEntry.text()
            
            self.global_settings['x_min'] = self.xMinSpin.value()
            self.global_settings['x_max'] = self.xMaxSpin.value()
            self.global_settings['y_min'] = self.yMinSpin.value()
            self.global_settings['y_max'] = self.yMaxSpin.value()
            
            self.global_settings['grid'] = self.gridCheck.isChecked()
            
            self.global_settings['show_legend'] = self.legendCheck.isChecked()
            legend_pos_jp = self.legendPosCombo.currentText()
            if legend_pos_jp in self.legend_pos_mapping:
                self.global_settings['legend_pos'] = self.legend_pos_mapping[legend_pos_jp]
            else:
                # デフォルトは右上
                self.global_settings['legend_pos'] = 'north east'
            
            if self.logXScaleRadio.isChecked():
                self.global_settings['scale_type'] = 'logx'
            elif self.logYScaleRadio.isChecked():
                self.global_settings['scale_type'] = 'logy'
            elif self.logLogScaleRadio.isChecked():
                self.global_settings['scale_type'] = 'loglog'
            else:
                self.global_settings['scale_type'] = 'normal'
            
            self.global_settings['caption'] = self.captionEntry.text()
            self.global_settings['label'] = self.labelEntry.text()
            self.global_settings['position'] = self.positionCombo.currentText()
            self.global_settings['width'] = self.widthSpin.value()
            self.global_settings['height'] = self.heightSpin.value()
            
            self.x_tick_step = self.xTickStepSpin.value()
            self.y_tick_step = self.yTickStepSpin.value()
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"グラフ全体設定の更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def convert_to_tikz(self):
        # データセットsが空か
        if not self.datasets or all(not dataset.get('data_x') for dataset in self.datasets):
            QMessageBox.warning(self, "警告", "データが読み込まれていません。先にデータを読み込んでください。")
            return
        
        try:
            self.update_global_settings()
            
            latex_code = self.generate_tikz_code_multi_datasets()
            
            self.resultText.setPlainText(latex_code)
            self.statusBar.showMessage("TikZコードが生成されました")
                        
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"変換中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("変換エラー")


    def add_dataset(self, name_arg=None):
        """新しいデータセットを追加する"""
        try:
            # 現在のデータセットの状態を保存
            if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
                self.update_current_dataset()
                
            final_name = ""
            # ボタンからの追加のみ
            if name_arg is None:
                dataset_count = self.datasetList.count() + 1
                text_from_dialog, ok = QInputDialog.getText(self, "データセット名", "新しいデータセット名を入力してください:",
                                                          QLineEdit.Normal, f"データセット{dataset_count}")
                if not ok or not text_from_dialog.strip(): # Ensure name is not empty or just whitespace
                    self.statusBar.showMessage("データセットの追加がキャンセルされました。", 3000)
                    return
                final_name = text_from_dialog.strip()
            else:
                if name_arg is False:  # Falseが渡された場合の対策
                    dataset_count = self.datasetList.count() + 1
                    final_name = f"データセット{dataset_count}"
                else:
                    final_name = str(name_arg).strip() 

            if not final_name: 
                dataset_count = self.datasetList.count() + 1
                final_name = f"データセット{dataset_count}"
                self.statusBar.showMessage("データセット名が空のため、デフォルト名を使用します。", 3000)

            dataset = {
                'name': final_name, # Always a string
                'data_source_type': 'measured',  # 'measured' または 'formula'
                'data_x': [],
                'data_y': [],
                'color': QColor('blue'),  # QColorオブジェクトを新規作成
                'line_width': 1.0,
                'marker_style': '*',
                'marker_size': 2.0,
                'plot_type': "line",
                'legend_label': final_name, # Always a string, initialized with name
                'show_legend': True,
                'equation': '',
                'domain_min': 0,
                'domain_max': 10,
                'samples': 200,
                'special_points': [],  # [(x, y, color, show_coords), ...]
                'annotations': [],     # [(x, y, text, color, pos), ...]
                'special_points_enabled': False,  # 特殊点チェックボックスの状態
                'annotations_enabled': False,    # 注釈チェックボックスの状態
                # ファイル読み込み関連の設定
                'file_path': '',
                'file_type': 'csv',  # 'csv' or 'excel' or 'manual'
                'sheet_name': '',
                'x_column': '',
                'y_column': ''
            }
            
            self.datasets.append(dataset)
            # additemはあくまでリストに機能の持たない名前を追加するだけ．UIはindexに基づいて別処理
            self.datasetList.addItem(final_name) 
            
            # 新しく追加したデータセットを選択（これがon_dataset_selectedを呼び出す）
            self.datasetList.setCurrentRow(len(self.datasets) - 1)
            self.statusBar.showMessage(f"データセット '{final_name}' を追加しました", 3000)

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データセット追加中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def remove_dataset(self):
        """選択されたデータセットを削除する"""
        try:
            if not self.datasets:
                QMessageBox.warning(self, "警告", "削除するデータセットがありません")
                return
            
            current_row = self.datasetList.currentRow()
            if current_row < 0:
                QMessageBox.warning(self, "警告", "削除するデータセットを選択してください")
                return
            
            dataset_name = str(self.datasets[current_row]['name'])
            
            msgBox, delete_button, cancel_button = self.create_styled_delete_confirmation(
                "確認", 
                f"データセット '{dataset_name}' を削除してもよろしいですか？\n\nこの操作は元に戻せません。",
                "削除"
            )
            
            msgBox.exec_()
            
            if msgBox.clickedButton() == delete_button:
                self.datasets.pop(current_row)
                item = self.datasetList.takeItem(current_row)
                if item:
                    del item 
                
                if self.datasets: 
                    new_index = max(0, min(current_row, len(self.datasets) - 1))
                    self.datasetList.setCurrentRow(new_index)
                else: 
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets() 
                    self.add_dataset("データセット1") 
                
                self.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました", 3000)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データセット削除中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def update_ui_for_no_datasets(self):
        """データセットない時（防御)"""
        self.legendLabel.setText("")
        pass # Placeholder

    def rename_dataset(self):
        """選択されたデータセットの名前を変更する"""
        try:
            if self.current_dataset_index < 0 or not self.datasets:
                QMessageBox.warning(self, "警告", "名前を変更するデータセットが選択されていません。")
                return
            
            current_row = self.datasetList.currentRow() 
            if current_row != self.current_dataset_index:
                 # 通常起こらない．current_dataset_index に合わせる
                 current_row = self.current_dataset_index
                 self.datasetList.setCurrentRow(current_row)

            current_name = str(self.datasets[current_row]['name'])
            current_legend = str(self.datasets[current_row].get('legend_label', current_name))

            new_name_text, ok = QInputDialog.getText(self, "データセット名の変更", \
                                                      "新しいデータセット名を入力してください:",
                                                      QLineEdit.Normal, current_name)# current_nameはデフォルト値
            
            if ok and new_name_text.strip():
                actual_new_name = new_name_text.strip()
                if not actual_new_name:
                    QMessageBox.warning(self, "警告", "データセット名は空にできません。")
                    return

                self.datasets[current_row]['name'] = actual_new_name
                if current_legend == current_name:
                    self.datasets[current_row]['legend_label'] = actual_new_name
                    self.legendLabel.setText(actual_new_name) 
                
                item = self.datasetList.item(current_row)
                if item:
                    # リストの表示名更新
                    item.setText(actual_new_name)
                self.statusBar.showMessage(f"データセット名を '{actual_new_name}' に変更しました", 3000)
            elif ok and not new_name_text.strip():
                 QMessageBox.warning(self, "警告", "データセット名は空にできません。")

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データセット名の変更中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def on_dataset_selected(self, row):
        """行選択時，rowは選択した行のindexを放り込む"""
        try:
            # 以前のデータセットの状態を保存（存在する場合）
            old_index = self.current_dataset_index
            if old_index >= 0 and old_index < len(self.datasets):
                self.update_current_dataset()
            if row < 0 or row >= len(self.datasets):  #　防御
                if not self.datasets:
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets()
                return
            self.current_dataset_index = row
            dataset = self.datasets[row]
            #!  UIのみの更新、保存処理は呼ばない（保存処理と分離）
            self.update_ui_from_dataset(dataset)
            self.statusBar.showMessage(f"データセット '{dataset['name']}' を選択しました", 3000)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def update_current_dataset(self):
        """現在のデータセット値を保存(selfの値をdataset[self.current_dataset_index]に)"""
        try:
            if self.current_dataset_index < 0 or not self.datasets or self.current_dataset_index >= len(self.datasets):
                return
            
            #! 代入されるのは辞書
            dataset = self.datasets[self.current_dataset_index]
            
            # 色の設定（初回1回のみ）
            if not isinstance(self.currentColor, QColor):
                self.currentColor = QColor(self.currentColor)

            dataset['color'] = QColor(self.currentColor)
            dataset['line_width'] = self.lineWidthSpin.value()
            dataset['marker_style'] = self.markerCombo.currentText()
            dataset['marker_size'] = self.markerSizeSpin.value()
            dataset['show_legend'] = self.legendCheck.isChecked()
            dataset['legend_label'] = self.legendLabel.text()
            
            # プロットタイプの設定
            if self.lineRadio.isChecked():
                dataset['plot_type'] = "line" #　線グラフ
            elif self.scatterRadio.isChecked():
                dataset['plot_type'] = "scatter" #　散布図
            elif self.lineScatterRadio.isChecked():
                dataset['plot_type'] = "line_scatter" #　線と点
            else:
                dataset['plot_type'] = "bar" #　棒グラフ
            
            # 特殊点
            special_points_table_data = []
            for row in range(self.specialPointsTable.rowCount()):
                x_item = self.specialPointsTable.item(row, 0)
                y_item = self.specialPointsTable.item(row, 1)
                color_combo = self.specialPointsTable.cellWidget(row, 2)
                coord_combo = self.specialPointsTable.cellWidget(row, 3)
                
                if x_item and y_item and color_combo and coord_combo:
                    special_points_table_data.append({
                        'x': x_item.text(),
                        'y': y_item.text(),
                        'color': color_combo.currentText(),
                        'coord_display': coord_combo.currentText()
                    })
            dataset['special_points_table_data'] = special_points_table_data
            
            # 注釈
            annotations_table_data = []
            for row in range(self.annotationsTable.rowCount()):
                x_item = self.annotationsTable.item(row, 0)
                y_item = self.annotationsTable.item(row, 1)
                text_item = self.annotationsTable.item(row, 2)
                color_combo = self.annotationsTable.cellWidget(row, 3)
                pos_combo = self.annotationsTable.cellWidget(row, 4)
                
                if x_item and y_item and text_item and color_combo and pos_combo:
                    annotations_table_data.append({
                        'x': x_item.text(),
                        'y': y_item.text(),
                        'text': text_item.text(),
                        'color': color_combo.currentText(),
                        'position': pos_combo.currentText()
                    })
            dataset['annotations_table_data'] = annotations_table_data
            
            if hasattr(self, 'measuredRadio') and hasattr(self, 'formulaRadio'):
                dataset['data_source_type'] = 'measured' if self.measuredRadio.isChecked() else 'formula'
                
                # 測定データの場合
                if self.measuredRadio.isChecked():
                    if self.csvRadio.isChecked():
                        dataset['file_type'] = 'csv'
                        dataset['file_path'] = self.fileEntry.text()
                    elif self.excelRadio.isChecked():
                        dataset['file_type'] = 'excel'
                        dataset['file_path'] = self.excelEntry.text()
                        dataset['sheet_name'] = self.sheetCombobox.currentText()
                    else:
                        dataset['file_type'] = 'manual'
                    
                    dataset['x_range'] = self.xRangeEntry.text().strip()
                    dataset['y_range'] = self.yRangeEntry.text().strip()                    
                    if self.manualRadio.isChecked():
                        data_x, data_y = [], []
                        for row in range(self.dataTable.rowCount()):
                            x_item = self.dataTable.item(row, 0)
                            y_item = self.dataTable.item(row, 1)
                            if x_item and y_item and x_item.text() and y_item.text():
                                try:
                                    data_x.append(float(x_item.text()))
                                    data_y.append(float(y_item.text()))
                                except ValueError:
                                    pass
                        dataset['data_x'] = data_x.copy()
                        dataset['data_y'] = data_y.copy()
                else:
                    dataset['equation'] = self.equationEntry.text()
                    dataset['domain_min'] = self.domainMinSpin.value()
                    dataset['domain_max'] = self.domainMaxSpin.value()
                    dataset['samples'] = self.samplesSpin.value()
                    dataset['show_tangent'] = self.showTangentCheck.isChecked()
                    dataset['tangent_x'] = self.tangentXSpin.value()
                    dataset['tangent_length'] = self.tangentLengthSpin.value()
                    
                    if hasattr(self, 'tangentColor'):
                        dataset['tangent_color'] = QColor(self.tangentColor)
                        
                        if hasattr(self, 'tangentColorButton'):
                            self.tangentColorButton.setStyleSheet(f'background-color: {self.tangentColor.name()};')
                    
                    if hasattr(self, 'tangentStyleCombo'):
                        dataset['tangent_style'] = self.tangentStyleCombo.currentText()  # 永続化のために保存
                        tangent_style = dataset.get('tangent_style', '実線')
                        index = self.tangentStyleCombo.findText(tangent_style)
                        if index >= 0:
                            self.tangentStyleCombo.setCurrentIndex(index)
                    
                    if hasattr(self, 'showTangentEquationCheck'):
                        dataset['show_tangent_equation'] = self.showTangentEquationCheck.isChecked()  # 永続化のために保存
                        self.showTangentEquationCheck.setChecked(dataset.get('show_tangent_equation', False))
                        #pythonは基本参照なので気にしなくておけ(Swift基本触ってるから頭バグるﾝｺﾞ〜泣)
            
            
            # 特殊点・注釈のチェックボックス状態を保存
            dataset['special_points_enabled'] = self.specialPointsCheck.isChecked()
            dataset['annotations_enabled'] = self.annotationsCheck.isChecked()
            
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データセット更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def update_ui_from_dataset(self, dataset):
        """新しい選択行を保存時点でのデータセットに基づいてUIを更新する（apply_fomula,on_dataset_selected）"""
        try:
            self.block_signals_temporarily(True)
            
            data_source_type = dataset.get('data_source_type', 'measured')
            
            self.dataSourceTypeDisplayLabel.setText("実測データ" if data_source_type == 'measured' else "数式データ")
            
            color = dataset.get('color', QColor('blue'))
            if not isinstance(color, QColor):
                color = QColor(color)
            self.currentColor = QColor(color)  
            self.colorButton.setStyleSheet(f'background-color: {self.currentColor.name()};')
            
            self.lineWidthSpin.setValue(dataset.get('line_width', 1.0))
            
            marker_style = str(dataset.get('marker_style', '*'))
            marker_index = self.markerCombo.findText(marker_style)
            if marker_index >= 0:
                self.markerCombo.setCurrentIndex(marker_index)
            
            self.markerSizeSpin.setValue(dataset.get('marker_size', 2.0))
            self.gridCheck.setChecked(dataset.get('grid',True))
            self.legendCheck.setChecked(dataset.get('show_legend', True))
            
            legend_text = str(dataset.get('legend_label', dataset.get('name', '')))
            self.legendLabel.setText(legend_text)
            
            plot_type = str(dataset.get('plot_type', 'line'))
            if plot_type == "line": 
                self.lineRadio.setChecked(True)
            elif plot_type == "scatter": 
                self.scatterRadio.setChecked(True)
            elif plot_type == "line_scatter": 
                self.lineScatterRadio.setChecked(True)
            else: 
                self.barRadio.setChecked(True)
                
            if data_source_type == 'measured':
                if not self.measuredRadio.isChecked():
                    self.measuredRadio.setChecked(True)
            else:
                if not self.formulaRadio.isChecked():
                    self.formulaRadio.setChecked(True)
            
            if data_source_type == 'measured':
                file_path = dataset.get('file_path', '')
                file_type = dataset.get('file_type', 'csv')
                
                self.xRangeEntry.setText(dataset.get('x_range', ''))
                self.yRangeEntry.setText(dataset.get('y_range', ''))
                
                if file_type == 'csv':
                    self.csvRadio.setChecked(True)
                    self.fileEntry.setText(file_path)
                elif file_type == 'excel':
                    self.excelRadio.setChecked(True)
                    self.excelEntry.setText(file_path)
                    
                    sheet_name = dataset.get('sheet_name', '')
                    if sheet_name:
                        index = self.sheetCombobox.findText(sheet_name)
                        if index >= 0:
                            self.sheetCombobox.setCurrentIndex(index)
                elif file_type == 'manual':
                    self.manualRadio.setChecked(True)
                
                self.toggle_source_fields()
                
                self.dataTable.setRowCount(0)
                if dataset.get('data_x') and len(dataset.get('data_x')) > 0:
                    self.update_data_table_from_dataset(dataset)
            else:
                self.equationEntry.setText(dataset.get('equation', 'x^2'))
                self.domainMinSpin.setValue(dataset.get('domain_min', 0))
                self.domainMaxSpin.setValue(dataset.get('domain_max', 10))
                self.samplesSpin.setValue(dataset.get('samples', 200))
            
                if hasattr(self, 'showTangentCheck'):
                    self.showTangentCheck.setChecked(dataset.get('show_tangent', False))
                
                if hasattr(self, 'tangentXSpin'):
                    self.tangentXSpin.setValue(dataset.get('tangent_x', 1))
                
                if hasattr(self, 'tangentLengthSpin'):
                    self.tangentLengthSpin.setValue(dataset.get('tangent_length', 2))
                
                if hasattr(self, 'tangentColorButton'):
                    tangent_color = dataset.get('tangent_color', QColor('red'))
                    self.tangentColor = QColor(tangent_color) 
                    self.tangentColorButton.setStyleSheet(f'background-color: {self.tangentColor.name()};')
                
                if hasattr(self, 'tangentStyleCombo'):
                    tangent_style = dataset.get('tangent_style', '実線')
                    index = self.tangentStyleCombo.findText(tangent_style)
                    if index >= 0:
                        self.tangentStyleCombo.setCurrentIndex(index)
                
                if hasattr(self, 'showTangentEquationCheck'):
                    self.showTangentEquationCheck.setChecked(dataset.get('show_tangent_equation', False))
            
            # データソースタイプ変更処理を明示的に呼び出し(最後にupdate_ui_based_on_data_source_typeを呼ぶ)
            self.on_data_source_type_changed(True)
            
            self.specialPointsCheck.setChecked(dataset.get('special_points_enabled', False))
            self.annotationsCheck.setChecked(dataset.get('annotations_enabled', False))
            
            # 特殊点
            self.specialPointsTable.setRowCount(0)
            special_points_table_data = dataset.get('special_points_table_data', [])
            for point_data in special_points_table_data:
                row_position = self.specialPointsTable.rowCount()
                self.specialPointsTable.insertRow(row_position)
                
                x_item = QTableWidgetItem(point_data.get('x', '0.0'))
                y_item = QTableWidgetItem(point_data.get('y', '0.0'))
                self.specialPointsTable.setItem(row_position, 0, x_item)
                self.specialPointsTable.setItem(row_position, 1, y_item)
                
                color_combo = QComboBox()
                color_combo.addItems(['red', 'blue', 'green', 'black', 'purple', 'orange', 'brown', 'gray'])
                color_combo.setCurrentText(point_data.get('color', 'red'))
                self.specialPointsTable.setCellWidget(row_position, 2, color_combo)
                
                coord_display_combo = QComboBox()
                coord_display_combo.addItems([
                    'なし', 
                    'X座標のみ（線のみ）', 
                    'X座標のみ（値も表示）', 
                    'Y座標のみ（線のみ）', 
                    'Y座標のみ（値も表示）', 
                    'X,Y座標（線のみ）', 
                    'X,Y座標（値も表示）'
                ])
                coord_display_combo.setCurrentText(point_data.get('coord_display', 'なし'))
                self.specialPointsTable.setCellWidget(row_position, 3, coord_display_combo)
            
            # 注釈
            self.annotationsTable.setRowCount(0)
            annotations_table_data = dataset.get('annotations_table_data', [])
            for ann_data in annotations_table_data:
                row_position = self.annotationsTable.rowCount()
                self.annotationsTable.insertRow(row_position)
                
                x_item = QTableWidgetItem(ann_data.get('x', '0.0'))
                y_item = QTableWidgetItem(ann_data.get('y', '0.0'))
                text_item = QTableWidgetItem(ann_data.get('text', '注釈テキスト'))
                self.annotationsTable.setItem(row_position, 0, x_item)
                self.annotationsTable.setItem(row_position, 1, y_item)
                self.annotationsTable.setItem(row_position, 2, text_item)
                
                color_combo = QComboBox()
                color_combo.addItems(['black', 'red', 'blue', 'green', 'purple', 'orange', 'brown', 'gray'])
                color_combo.setCurrentText(ann_data.get('color', 'black'))
                self.annotationsTable.setCellWidget(row_position, 3, color_combo)
                
                pos_combo = QComboBox()
                pos_combo.addItems(['上', '右上', '右', '右下', '下', '左下', '左', '左上'])
                pos_combo.setCurrentText(ann_data.get('position', '右上'))
                self.annotationsTable.setCellWidget(row_position, 4, pos_combo)
            
            self.block_signals_temporarily(False)
            
            
        except Exception as e:
            self.block_signals_temporarily(False)
            QMessageBox.critical(self, "エラー", f"UI更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def block_signals_temporarily(self, block):
        """値の更新による挙動を防ぐ"""
        # 個別設定
        ui_elements = [
            self.lineWidthSpin,
            self.markerCombo,
            self.markerSizeSpin,
            self.lineRadio,
            self.scatterRadio,
            self.lineScatterRadio,
            self.barRadio,
            self.legendCheck,
            self.legendLabel
        ]
        
        if hasattr(self, 'measuredRadio'):
            ui_elements.append(self.measuredRadio)
        if hasattr(self, 'formulaRadio'):
            ui_elements.append(self.formulaRadio)
            
        if hasattr(self, 'equationEntry'):
            ui_elements.append(self.equationEntry)
        if hasattr(self, 'domainMinSpin'):
            ui_elements.append(self.domainMinSpin)
        if hasattr(self, 'domainMaxSpin'):
            ui_elements.append(self.domainMaxSpin)
        if hasattr(self, 'samplesSpin'):
            ui_elements.append(self.samplesSpin)
        if hasattr(self, 'showTangentCheck'):
            ui_elements.append(self.showTangentCheck)
            
        # すべての要素のシグナルをブロック/解除
        for element in ui_elements:
            element.blockSignals(block)

    def update_ui_based_on_data_source_type(self):
        """データソースタイプに基づいて手入力，数式入力のUI表示更新"""
        if not hasattr(self, 'measuredRadio') or not hasattr(self, 'formulaRadio'):
            return 
            
        if self.measuredRadio.isChecked():
            self.csvRadio.setEnabled(True)
            self.excelRadio.setEnabled(True)
            self.fileEntry.setEnabled(self.csvRadio.isChecked())
            self.excelEntry.setEnabled(self.excelRadio.isChecked())
            self.sheetCombobox.setEnabled(self.excelRadio.isChecked())
            
            is_csv_or_excel = self.csvRadio.isChecked() or self.excelRadio.isChecked()
            
            self.xRangeEntry.setEnabled(is_csv_or_excel)
            self.yRangeEntry.setEnabled(is_csv_or_excel)
            
            self.manualRadio.setEnabled(True)
            self.dataTable.setEnabled(self.manualRadio.isChecked())
            
            self.equationEntry.setEnabled(False)
            self.domainMinSpin.setEnabled(False)
            self.domainMaxSpin.setEnabled(False)
            self.samplesSpin.setEnabled(False)
            
        else:
            self.csvRadio.setEnabled(False)
            self.excelRadio.setEnabled(False)
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.xRangeEntry.setEnabled(False)
            self.yRangeEntry.setEnabled(False)
            self.manualRadio.setEnabled(False)
            self.dataTable.setEnabled(False)
            
            self.equationEntry.setEnabled(True)
            self.domainMinSpin.setEnabled(True)
            self.domainMaxSpin.setEnabled(True)
            self.samplesSpin.setEnabled(True)

    def update_data_table_from_dataset(self, dataset):
        """データテーブルにデータセットの内容を反映する"""
        try:
            self.dataTable.setRowCount(0)
            
            if not dataset.get('data_x') or not dataset.get('data_y'):
                return
                
            data_x = dataset.get('data_x', [])
            data_y = dataset.get('data_y', [])
            
            if not data_x or not data_y:
                return
            
            # データを表示
            for i, (x, y) in enumerate(zip(data_x, data_y)):
                self.dataTable.insertRow(i)
                self.dataTable.setItem(i, 0, QTableWidgetItem(str(x)))
                self.dataTable.setItem(i, 1, QTableWidgetItem(str(y)))
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データテーブル更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    #* 重要
    def generate_tikz_code_multi_datasets(self):
        """最終的なLaTeXコードを生成する"""
        latex = []
        
        latex.append("\\begin{figure}[" + self.global_settings['position'] + "]")
        latex.append("  \\centering")
        width = self.global_settings['width']
        height = self.global_settings['height']
        latex.append("  \\begin{tikzpicture}")
        #軸の設定
        x_min = self.global_settings['x_min']
        x_max = self.global_settings['x_max']
        y_min = self.global_settings['y_min']
        y_max = self.global_settings['y_max']
        
        all_x_values = []
        all_y_values = []
        for dataset in self.datasets:
            if dataset.get('data_x') and dataset.get('data_y'):
                all_x_values.extend(dataset['data_x'])
                all_y_values.extend(dataset['data_y'])
        
        if all_x_values and all_y_values:
            data_x_min = min(all_x_values)
            data_x_max = max(all_x_values)
            data_y_min = min(all_y_values)
            data_y_max = max(all_y_values)
            
            # データ数1
            if abs(data_x_max - data_x_min) < 1e-10:
                data_x_min -= 0.5 if abs(data_x_min) > 1 else 0.1
                data_x_max += 0.5 if abs(data_x_max) > 1 else 0.1
            
            if abs(data_y_max - data_y_min) < 1e-10:
                data_y_min -= 0.5 if abs(data_y_min) > 1 else 0.1
                data_y_max += 0.5 if abs(data_y_max) > 1 else 0.1
            
            # 軸の範囲を越える時の調整(10%余裕持たせる)
            if x_min > data_x_min or x_min == 0:
                x_min = data_x_min - abs(data_x_min) * 0.1 - 0.1
            if x_max < data_x_max or x_max == 0:
                x_max = data_x_max + abs(data_x_max) * 0.1 + 0.1
            if y_min > data_y_min or y_min == 0:
                y_min = data_y_min - abs(data_y_min) * 0.1 - 0.1
            if y_max < data_y_max or y_max == 0:
                y_max = data_y_max + abs(data_y_max) * 0.1 + 0.1
            
            # 小さすぎる値を補正
            if abs(y_min) < 1e-10:
                y_min = -0.1
            if abs(y_max) < 1e-10:
                y_max = 0.1
            if abs(x_min) < 1e-10:
                x_min = -0.1
            if abs(x_max) < 1e-10:
                x_max = 0.1
            
            if abs(x_min - x_max) < 1e-10:
                x_min -= 0.5
                x_max += 0.5
            if abs(y_min - y_max) < 1e-10:
                y_min -= 0.5
                y_max += 0.5
                
            # 軸の範囲が小さい時
            if abs(x_max - x_min) < 1e-3:
                margin = abs(x_min) * 0.2 if abs(x_min) > 1e-10 else 0.1
                x_min -= margin
                x_max += margin
            if abs(y_max - y_min) < 1e-3:
                margin = abs(y_min) * 0.2 if abs(y_min) > 1e-10 else 0.1
                y_min -= margin
                y_max += margin
        
        # 軸の表示点
        xtick_values = []
        ytick_values = []

        tick_min = math.ceil(x_min / self.x_tick_step) * self.x_tick_step
        tick_max = math.floor(x_max / self.x_tick_step) * self.x_tick_step
        
        current = tick_min
        while current <= tick_max:
            xtick_values.append(current)
            current += self.x_tick_step
    
        tick_min = math.ceil(y_min / self.y_tick_step) * self.y_tick_step
        tick_max = math.floor(y_max / self.y_tick_step) * self.y_tick_step
        
        current = tick_min
        while current <= tick_max:
            ytick_values.append(current)
            current += self.y_tick_step
    
        axis_options = []
        axis_options.append(f"width={width}\\textwidth")
        axis_options.append(f"height={height}\\textwidth")
        axis_options.append(f"xlabel={{{self.global_settings['x_label']}}}")
        axis_options.append(f"ylabel={{{self.global_settings['y_label']}}}")
        
        #　原則起こらない
        if x_min != x_max:
            axis_options.append(f"xmin={x_min}, xmax={x_max}")
        if y_min != y_max:
            axis_options.append(f"ymin={y_min}, ymax={y_max}")
        
        scale_type = self.global_settings.get('scale_type', 'normal')
        is_xlog = (scale_type == 'logx' or scale_type == 'loglog')
        is_ylog = (scale_type == 'logy' or scale_type == 'loglog')
        
        if is_xlog and x_min <= 0:
            self.statusBar.showMessage(f"X軸の対数スケールには正の値のみ有効です。X軸の最小値を0.1に自動調整しました。", 5000)
            if all_x_values:
                positive_x = [x for x in all_x_values if x > 0]
                if positive_x:
                    x_min = min(positive_x) * 0.5  
                else:
                    x_min = 0.1  
            else:
                x_min = 0.1  
        
        if is_ylog and y_min <= 0:
            self.statusBar.showMessage(f"Y軸の対数スケールには正の値のみ有効です。Y軸の最小値を0.1に自動調整しました。", 5000)
            if all_y_values:
                positive_y = [y for y in all_y_values if y > 0]
                if positive_y:
                    y_min = min(positive_y) * 0.5  
                else:
                    y_min = 0.1  
            else:
                y_min = 0.1  
        
        if is_xlog:
            axis_options.append("xmode=log")
            axis_options.append("log basis x=10")
        if is_ylog:
            axis_options.append("ymode=log")
            axis_options.append("log basis y=10")
        
        if is_xlog:
            # 指数部分
            log_x_min = math.floor(math.log10(max(x_min, 1e-10)))
            log_x_max = math.ceil(math.log10(max(x_max, 1e-10)))
            xticks = ','.join(str(10**i) for i in range(log_x_min, log_x_max + 1))
            axis_options.append("xminorticks=true")
        else:
            if xtick_values:
                xticks = ','.join(str(round(tick, 8)) for tick in xtick_values)
            else:
                if abs(x_max - x_min) < 1e-10 or self.x_tick_step < 1e-10:
                    xticks = str(round(x_min, 8))
                else:
                    steps = max(1, min(20, int((x_max - x_min) / self.x_tick_step) + 1))  
                    xticks = ','.join(str(round(x_min + i * self.x_tick_step, 8)) for i in range(steps))
        
        if is_ylog:
            log_y_min = math.floor(math.log10(max(y_min, 1e-10)))
            log_y_max = math.ceil(math.log10(max(y_max, 1e-10)))
            yticks = ','.join(str(10**i) for i in range(log_y_min, log_y_max + 1))
            axis_options.append("yminorticks=true")
        else:
            if ytick_values:
                yticks = ','.join(str(round(tick, 8)) for tick in ytick_values)
            else:
                if abs(y_max - y_min) < 1e-10 or self.y_tick_step < 1e-10:
                    yticks = str(round(y_min, 8))
                else:
                    steps = max(1, min(20, int((y_max - y_min) / self.y_tick_step) + 1))  # 最大20ステップに制限
                    yticks = ','.join(str(round(y_min + i * self.y_tick_step, 8)) for i in range(steps))
        
        axis_options.append(f"xtick={{{xticks}}}")
        axis_options.append(f"ytick={{{yticks}}}")
        
        axis_options.append("tick align=outside")
        axis_options.append("minor tick num=1")
        axis_options.append("tick label style={font=\\small}")
        axis_options.append("every node near coord/.style={font=\\footnotesize}")
        axis_options.append("clip=true")  
        
        if self.global_settings['grid']:
            axis_options.append("grid=both")
        
        if self.global_settings.get('show_legend', True):
            axis_options.append(f"legend pos={self.global_settings['legend_pos']}")
        
        #* ========================axis開始====================================
        latex.append(f"        \\begin{{axis}}[")
        latex.append(f"            {','.join(axis_options)}")
        latex.append("        ]")
        
        for i, dataset in enumerate(self.datasets):
            if dataset.get('data_source_type') == 'formula' and not dataset.get('equation'):
                continue  
                
            if dataset.get('data_source_type') == 'measured' and (not dataset.get('data_x') or not dataset.get('data_y')):
                continue  
            
            # 個別設定
            plot_type = dataset.get('plot_type', 'line')
            color = dataset.get('color', QColor('blue')).name()
            line_width = dataset.get('line_width', 1.0)
            marker_style = dataset.get('marker_style', '*')
            marker_size = dataset.get('marker_size', 2.0)
            
            show_legend = self.global_settings.get('show_legend', True) and dataset.get('show_legend', True)
            legend_label = dataset.get('legend_label', dataset.get('name', ''))
            
            # 対数かつ測定データ
            if (is_xlog or is_ylog) and dataset.get('data_source_type') == 'measured':
                data_x = dataset.get('data_x', [])
                data_y = dataset.get('data_y', [])
                filtered_data_x = []
                filtered_data_y = []
                invalid_points = 0
                
                for x, y in zip(data_x, data_y):
                    if (is_xlog and x <= 0) or (is_ylog and y <= 0):
                        invalid_points += 1
                        continue
                    filtered_data_x.append(x)
                    filtered_data_y.append(y)
                
                if invalid_points > 0:
                    warning_msg = f"データセット '{dataset.get('name', '')}' の {invalid_points}個の点が対数スケールに適さないため除外されました"
                    latex.append(f"        % 警告: {warning_msg}")
                    self.statusBar.showMessage(warning_msg, 5000)
                
                if not filtered_data_x:
                    latex.append(f"        % データセット名[ {dataset.get('name', '')} ] - 対数スケールに適した点がありません")
                    continue
                
                filtered_dataset = dataset.copy()
                filtered_dataset['data_x'] = filtered_data_x
                filtered_dataset['data_y'] = filtered_data_y
                
                #* データセット
                self.add_dataset_to_latex(latex, filtered_dataset, i, plot_type, color, line_width, 
                                        marker_style, marker_size, show_legend, legend_label)
            else:
                # 通常のデータセットを処理
                self.add_dataset_to_latex(latex, dataset, i, plot_type, color, line_width, 
                                        marker_style, marker_size, show_legend, legend_label)
            
            #*==========================特殊点の追加===============================
            #* 特殊点 -> 垂線 -> 注釈
            special_points = dataset.get('special_points', [])
            for point in special_points:
                x, y, point_color, coord_display = point  
                latex.append(f"        % 特殊点 [ {dataset.get('name', '')} ]")
                latex.append(f"        \\addplot[only marks, mark=*, {point_color}] coordinates {{")
                latex.append(f"            ({x}, {y})")
                latex.append("        };")
                
                def is_tick_value(val, tick_min, tick_max, tick_step, tol=1e-6):
                    if tick_step is None or tick_step <= 0:
                        return False
                    n = round((val - tick_min) / tick_step) # roundなので整数値になりnとtick_valは同じはず
                    tick_val = tick_min + n * tick_step
                    return abs(val - tick_val) < tol and tick_min <= val <= tick_max

                if coord_display.startswith('X座標') or coord_display.startswith('X,Y座標'):
                    latex.append(f"        % X軸への垂線")
                    latex.append(f"        \\draw[{point_color}, dashed] (axis cs:{x},{y}) -- (axis cs:{x},{y_min});")
                    if '値も表示' in coord_display and not is_tick_value(x, x_min, x_max, self.x_tick_step):
                        formatted_x = '{:g}'.format(x)
                        latex.append(f"        % X座標値を表示")
                        latex.append(f"        \\node[{point_color}, below, yshift=-2pt, font=\\small] at (axis cs:{x},{y_min}) {{{formatted_x}}};")

                if coord_display.startswith('Y座標') or coord_display.startswith('X,Y座標'):
                    latex.append(f"        % Y軸への垂線")
                    latex.append(f"        \\draw[{point_color}, dashed] (axis cs:{x},{y}) -- (axis cs:{x_min},{y});")
                    if '値も表示' in coord_display and not is_tick_value(y, y_min, y_max, self.y_tick_step):
                        formatted_y = '{:g}'.format(y)
                        latex.append(f"        % Y座標値を表示")
                        latex.append(f"        \\node[{point_color}, left, xshift=-2pt, font=\\small] at (axis cs:{x_min},{y}) {{{formatted_y}}};")
        
        # 注釈の追加
        for i, dataset in enumerate(self.datasets):
            annotations = dataset.get('annotations', [])
            for ann in annotations:
                x, y, text, color, pos = ann
                latex.append(f"        % 注釈 [ {dataset.get('name', '')} ]")
                latex.append(f"        \\node at (axis cs:{x},{y}) [anchor={pos}, font=\\small, text={color}] {{{text}}};")
        
        # axis終了
        latex.append("        \\end{axis}")
        
        # tikzpicture終了
        latex.append("    \\end{tikzpicture}")
        
        latex.append(f"    \\caption{{{self.global_settings['caption']}}}")
        latex.append(f"    \\label{{{self.global_settings['label']}}}")
        
        latex.append("\\end{figure}")
        
        return '\n'.join(latex)
    
    def add_dataset_to_latex(self, latex, dataset, index, plot_type, color, line_width, 
                             marker_style, marker_size, show_legend, legend_label):
        """
        LaTeXコードにデータセットを追加する\n
        latex引数は参照なので戻り値はいらない\n
        formulaの場合は途中でreturn
        """
        data_source_type = dataset.get('data_source_type', 'measured')
        
        scale_type = self.global_settings.get('scale_type', 'normal')
        is_xlog = scale_type == 'logx' or scale_type == 'loglog'
        is_ylog = scale_type == 'logy' or scale_type == 'loglog'
        
        if data_source_type == 'formula':
            equation = dataset.get('equation', 'x^2')
            tikz_equation = equation
            domain_min = dataset.get('domain_min', 0)
            domain_max = dataset.get('domain_max', 10)
            samples = dataset.get('samples', 200)
            
            if is_xlog and domain_min <= 0:
                latex.append(f"        % 警告: X軸が対数スケールのため、数式のxの範囲が調整されました")
                domain_min = max(0.01, domain_min)  
            
            if not equation.strip():
                return
            
            tikz_color = self.color_to_tikz_rgb(dataset.get('color', QColor('blue')))
            
            # 理論曲線のオプション
            theory_options = []
            theory_options.append(f"domain={domain_min}:{domain_max}")
            theory_options.append(f"samples={samples}")
            theory_options.append("smooth")
            theory_options.append("thick")
            theory_options.append(tikz_color)  
            theory_options.append(f"line width={dataset.get('line_width', 1.0)}pt")
            
            latex.append(f"        % データセット[ {dataset.get('name', '')} ] （数式: {equation}）")
            latex.append(f"        \\addplot[{', '.join(theory_options)}] {{")
            latex.append(f"            {tikz_equation}")
            latex.append("        };")
            
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")

            if dataset.get('show_tangent', False):
                tangent_x = dataset.get('tangent_x', 1)
                tangent_length = dataset.get('tangent_length', 2)
                
                tangent_color = self.color_to_tikz_rgb(dataset.get('tangent_color', QColor('red')))
                
                line_style = ""
                tangent_style = dataset.get('tangent_style', '実線')
                if tangent_style == '点線':
                    line_style = "dotted"
                elif tangent_style == '破線':
                    line_style = "dashed"
                elif tangent_style == '一点鎖線':
                    line_style = "dashdotted"
                
                #*==========================接線===============================
                tangent_options = []
                if line_style:
                    tangent_options.append(line_style)
                tangent_options.append("thick")
                tangent_options.append(tangent_color)  
                tangent_options.append(f"line width={dataset.get('line_width', 1.5)}pt")

                latex.append(f"        % 接線 [ {dataset.get('name', '')} ]")
                
                try:
                    x_val = tangent_x
                    formula = dataset.get('equation', 'x^2') 
                    
                    python_formula = formula.replace('^', '**')
                    
                    y_min = self.global_settings['y_min']
                    y_max = self.global_settings['y_max']
                    
                    #　math以外を封じる
                    y_val = eval(python_formula.replace('x', str(x_val)), {"__builtins__": {}}, {"math": math})

                    if y_val < y_min or y_val > y_max:
                        latex.append(f"        % 計算されたy値 ({round(y_val,3)}) がグラフ範囲外です．")
                        self.statusBar.showMessage(f"警告: 点 (x={round(x_val,3)}) の計算値 (y={round(y_val,3)}) がグラフ範囲外です", 5000)
                    
                    # 接点
                    point_code = f"""        % 接点をマーク
        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({x_val}, {y_val})}};\n"""
                    latex.append(point_code)
                    
                    # 中心差分法で傾き計算
                    dx = 0.0001
                    globals_dict = {"__builtins__": {}}
                    locals_dict = {"math": math}
                    
                    left_x = x_val - dx
                    right_x = x_val + dx
                    
                    try:
                        # y_1:左  y_2:右
                        y1 = eval(python_formula.replace('x', str(left_x)), globals_dict, locals_dict)
                        y2 = eval(python_formula.replace('x', str(right_x)), globals_dict, locals_dict)
                        slope = (y2 - y1) / (2 * dx)
                    except Exception as inner_e:
                        latex.append(f"        % 接線の微分計算でエラー: {str(inner_e)}. 中央差分法を使用できないため、前方差分法を試みます。")
                        # 前方差分法を試みる
                        try:
                            forward_x = x_val + dx
                            y_current = y_val  
                            y_forward = eval(python_formula.replace('x', str(forward_x)), globals_dict, locals_dict)
                            slope = (y_forward - y_current) / dx
                        except Exception as e2:
                            # 基本的に起こらない
                            latex.append(f"        % 前方差分も失敗。デフォルトの傾き 1 を使用します。エラー: {str(e2)}")
                            slope = 1.0
                            
                    #  y = mx + b 
                    slope_rounded = round(slope, 3) #傾き項(m)
                    intercept = round(y_val - slope * x_val, 3) # 切片(b)
                    equation_text = f"y = {slope_rounded}x + {intercept}" if intercept >= 0 else f"y = {slope_rounded}x - {abs(intercept)}"
                    
                    latex.append(f"        % x={round(x_val,3)}, y={round(y_val,3)}, 傾き={slope_rounded}, 切片={intercept}")
                    
                    test_y = slope_rounded * x_val + intercept
                    if abs(test_y - y_val) > 0.1:  # 許容誤差
                        latex.append(f"        % 警告: 接線方程式が元の点を通りません。式による計算値: {test_y}, 元の点: {y_val}")
                    
                    tangent_equation = f"{round(y_val,3)} + {round(slope,3)}*(x - {round(x_val,3)})"
                    
                    # 接線のプロット
                    tangent_code = f"""        % 接線のプロット
        \\addplot[{', '.join(tangent_options)}, domain={round(x_val-tangent_length/2,3)}:{round(x_val+tangent_length/2,3)}] {{
            {tangent_equation}
        }};\n"""
                    latex.append(tangent_code)
                    
                    if dataset.get('show_tangent_equation', False):
                        equation_pos_x = x_val
                        # 表示位置を調整
                        if y_val < y_min:
                            equation_pos_y = y_min + 0.5  # 上
                        elif y_val > y_max:
                            equation_pos_y = y_max - 0.5  # 下
                        else:
                            equation_pos_y = y_val + 0.5  # 通常は点の少し上
                        
                        # 接線ラベル
                        equation_display = f"""        % 接線の方程式を表示\n        %下の(axis cs:{{x座標の値}}，{{y座標の値}})の値を調整すると接線のラベルの表示位置を調整できます．
        \\node[anchor=south, font=\\small, {tangent_color}] at (axis cs:{equation_pos_x}, {equation_pos_y}) {{{equation_text}}};\n"""
                        latex.append(equation_display)
                    
                    # 凡例
                    latex.append(f"        \\addlegendimage{{{', '.join(tangent_options)}}};")
                    
                    # 凡例エントリを追加
                    if show_legend:
                        latex.append(f"        \\addlegendentry{{{legend_label}の接線 (x={tangent_x})}}")
                except Exception as e:
                    error_msg = f"接線の計算でエラーが発生しました。式: {formula}, エラー: {str(e)}"
                    self.statusBar.showMessage(f"警告: {error_msg}", 5000)
                    
                    latex.append(f"        % {error_msg}")
                    
                    try:
                        python_formula = formula.replace('^', '**')
                        point_y = eval(python_formula.replace('x', str(x_val)), {"__builtins__": {}}, {"math": math})
                        latex.append(f"        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({tangent_x}, {point_y})}}; % 接線計算エラーだが点は表示")
                    except Exception as point_error:
                        # それでも失敗したら原点にプロット
                        latex.append(f"        % 点の計算も失敗: {str(point_error)}")
                        latex.append(f"        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({tangent_x}, 0)}}; % エラーのため原点にプロット")
                    
                    latex.append(f"        \\addplot[{', '.join(tangent_options)}, dashed] coordinates {{({tangent_x-0.5}, 0) ({tangent_x+0.5}, 0)}}; % エラーのため水平線を表示")
                    
                    latex.append(f"        \\addlegendimage{{{', '.join(tangent_options)}, dashed}};")
                    if show_legend:
                        latex.append(f"        \\addlegendentry{{{legend_label}の接線 (計算エラー)}}")
            
            # 理論曲線の場合はここで終了し脱出
            return
            
        # 実測値
        coordinates = []
        for x, y in zip(dataset.get('data_x', []), dataset.get('data_y', [])):
            coordinates.append(f"({round(x,3)}, {round(y,3)})")
        
        if not coordinates:
            return
        
        tikz_color = self.color_to_tikz_rgb(QColor(color))
        
        # 線
        if plot_type == "line" or plot_type == "line_scatter":
            plot_options = []
            plot_options.append(tikz_color)  
            plot_options.append(f"line width={line_width}pt")
            plot_options.append("thick")
            
            latex.append(f"        % データセット[ {dataset.get('name', '')} ] （線）")
            latex.append(f"        \\addplot[{', '.join(plot_options)}] coordinates {{")
            
            # 1行に20個ずつ
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
        
        # 点
        if plot_type == "scatter" or plot_type == "line_scatter":
            scatter_options = []
            scatter_options.append("only marks")
            scatter_options.append(f"mark={marker_style}")
            scatter_options.append(tikz_color) 
            scatter_options.append(f"mark size={marker_size}")
            
            latex.append(f"        % データセット{index+1}: {dataset.get('name', '')} （点）")
            latex.append(f"        \\addplot[{', '.join(scatter_options)}] coordinates {{")
            
            # 1行に20個ずつ
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            if show_legend and plot_type == "scatter": # line_scatterの場合は既出
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
                
        # 棒グラフ
        if plot_type == "bar":
            bar_options = []
            bar_options.append(f"ybar")
            bar_options.append(tikz_color)  
            bar_options.append(f"fill={tikz_color.replace('color = ','')}")  # fill=color = ... ではなく fill=...
            
            latex.append(f"        % データセット[ {dataset.get('name', '')} ] （棒グラフ）")
            latex.append(f"        \\addplot[{', '.join(bar_options)}] coordinates {{")
            
            # 1行に20個ずつ
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")

    def assign_special_points_to_dataset(self):
        if self.current_dataset_index < 0:
            QMessageBox.warning(self, "警告", "特殊点を割り当てるデータセットを選択してください")
            return
        
        special_points = []
        for row in range(self.specialPointsTable.rowCount()):
            x_item = self.specialPointsTable.item(row, 0)
            y_item = self.specialPointsTable.item(row, 1)
            color_combo = self.specialPointsTable.cellWidget(row, 2)
            coord_display_combo = self.specialPointsTable.cellWidget(row, 3)
            
            if x_item and y_item and color_combo and coord_display_combo:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    color_val = color_combo.currentText()
                    coord_display = coord_display_combo.currentText()
                    special_points.append((x_val, y_val, color_val, coord_display))
                except ValueError:
                    pass
        
        dataset = self.datasets[self.current_dataset_index]
        dataset['special_points'] = special_points
        
        QMessageBox.information(self, "成功", 
                              f"データセット '{dataset['name']}' に {len(special_points)} 個の特殊点を割り当てました")
    
    def assign_annotations_to_dataset(self):
        if self.current_dataset_index < 0:
            QMessageBox.warning(self, "警告", "注釈を割り当てるデータセットを選択してください")
            return
        
        annotations = []
        for row in range(self.annotationsTable.rowCount()):
            x_item = self.annotationsTable.item(row, 0)
            y_item = self.annotationsTable.item(row, 1)
            text_item = self.annotationsTable.item(row, 2)
            color_combo = self.annotationsTable.cellWidget(row, 3)
            pos_combo = self.annotationsTable.cellWidget(row, 4)
            
            if x_item and y_item and text_item and color_combo and pos_combo:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    text_val = text_item.text()
                    color_val = color_combo.currentText()
                    pos_val = pos_combo.currentText()
                    
                    tikz_pos = self.convert_position_to_tikz(pos_val)
                    
                    annotations.append((x_val, y_val, text_val, color_val, tikz_pos))
                except ValueError:
                    pass
        
        dataset = self.datasets[self.current_dataset_index]
        dataset['annotations'] = annotations
        
        QMessageBox.information(self, "成功", 
                              f"データセット '{dataset['name']}' に {len(annotations)} 個の注釈を割り当てました")
    
    def convert_position_to_tikz(self, jp_position):
        """日本語の位置表記をTikZの位置表記に変換する\n
        TikZのアンカー指定はなぜか直感と逆になる\n
        例: 「右上」は「左下(south west)」になる\n
        """
        position_map = {
            '上': 'south',      # 下アンカー → 上に表示
            '右上': 'south west', # 左下アンカー → 右上に表示
            '右': 'west',       # 左アンカー → 右に表示
            '右下': 'north west', # 左上アンカー → 右下に表示
            '下': 'north',      # 上アンカー → 下に表示
            '左下': 'north east', # 右上アンカー → 左下に表示
            '左': 'east',       # 右アンカー → 左に表示
            '左上': 'south east'  # 右下アンカー → 左上に表示
        }
        return position_map.get(jp_position, 'south east')  # デフォルトは左上に表示

    def apply_formula(self):
        """数式を適用してデータを生成する"""
        self.update_current_dataset()
        if self.current_dataset_index < 0:
            QMessageBox.warning(self, "警告", "数式を適用するデータセットを選択してください")
            return
            
        equation = self.equationEntry.text().strip()
        if not equation:
            QMessageBox.warning(self, "警告", "有効な数式を入力してください")
            return
            
        try:
            domain_min = self.domainMinSpin.value()
            domain_max = self.domainMaxSpin.value()
            samples = self.samplesSpin.value()
            
            dataset = self.datasets[self.current_dataset_index]
            dataset['data_source_type'] = 'formula'
            dataset['equation'] = equation
            dataset['domain_min'] = domain_min
            dataset['domain_max'] = domain_max
            dataset['samples'] = samples
            
            dataset['show_tangent'] = self.showTangentCheck.isChecked()
            dataset['tangent_x'] = self.tangentXSpin.value()
            dataset['tangent_length'] = self.tangentLengthSpin.value()
            dataset['tangent_color'] = QColor(self.tangentColor)
            dataset['tangent_style'] = self.tangentStyleCombo.currentText()
            dataset['show_tangent_equation'] = self.showTangentEquationCheck.isChecked()
            
            # xの範囲を分割
            x_values = np.linspace(domain_min, domain_max, samples).tolist()
            dataset['data_x'] = x_values
            
            # 実際の値はTikZコード生成時に計算されるため空
            dataset['data_y'] = []
            
            python_formula = equation.replace('^', '**')
            
            y_min = self.global_settings['y_min']
            y_max = self.global_settings['y_max']
            
            # テスト点でチェック（最初、中央、最後）
            test_points = [domain_min, (domain_min + domain_max) / 2, domain_max]
            out_of_range_points = []
            
            for test_x in test_points:
                try:
                    # math以外の組み込み関数を封じる
                    test_y = eval(python_formula.replace('x', str(test_x)), {"__builtins__": {}}, {"math": math})
                    if test_y < y_min or test_y > y_max:
                        out_of_range_points.append((test_x, test_y))
                except Exception as e:
                    pass
            
            if out_of_range_points:
                point_info = ", ".join([f"(x={x:.2f}, y={y:.2f})" for x, y in out_of_range_points])
                QMessageBox.warning(self, "表示範囲外の値", 
                    f"計算された一部の点がグラフ表示範囲外です: {point_info}など\n\n"
                    f"グラフ表示範囲: Y軸 [{y_min}, {y_max}]\n\n"
                    "グラフの一部が表示されない可能性があります。X軸，Y軸の範囲を調整することを検討してください。")
            
            self.update_ui_from_dataset(dataset)
            self.statusBar.showMessage("数式データを適用しました", 3000)
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"数式の適用中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("数式適用エラー", 3000)

    def on_data_source_type_changed(self, checked):
        """データソースタイプが変更されたときに呼ばれる"""
        if not checked:  
            return
            
        if self.current_dataset_index >= 0:
            self.update_current_dataset()
            
        is_measured = self.measuredRadio.isChecked()
        
        self.measuredContainer.setVisible(is_measured)
        self.formulaContainer.setVisible(not is_measured)
        
        if self.current_dataset_index >= 0:
            # 選択した新しいindexが放り込まれる(on_dataset_selected参照）
            dataset = self.datasets[self.current_dataset_index]
            dataset['data_source_type'] = 'measured' if is_measured else 'formula'
            
            self.dataSourceTypeDisplayLabel.setText("実測データ" if is_measured else "数式データ")
            
        self.update_ui_based_on_data_source_type()

    def add_table_row(self):
        row_position = self.dataTable.rowCount()
        self.dataTable.insertRow(row_position)
    
    def remove_table_row(self):
        selected_rows = set(index.row() for index in self.dataTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.dataTable.removeRow(row)
    
    def select_color(self):
        """線の色を選択するダイアログを表示"""
        color = QColorDialog.getColor(self.currentColor, self, "線の色を選択")
        if color.isValid():
            self.currentColor = color
            self.colorButton.setStyleSheet(f'background-color: {color.name()};')
    
    def select_tangent_color(self):
        """接線の色を選択するダイアログを表示"""
        color = QColorDialog.getColor(self.tangentColor, self, "接線の色を選択")
        if color.isValid():
            self.tangentColor = color
            self.tangentColorButton.setStyleSheet(f'background-color: {color.name()};')
            self.statusBar.showMessage("接線の色を設定しました", 2000)

    def insert_function_from_table(self, row):
        """関数テーブルの関数をダブルクリックして挿入"""
        try:
            function_item = self.tikzFunctionsTable.item(row, 0)
            if function_item:
                function_text = function_item.text()
                
                if function_text in ["pi", "e"]:
                    self.insert_into_equation(function_text)
                else:
                    if "(" in function_text and ")" in function_text:
                        bracket_content = function_text[function_text.find("(")+1:function_text.find(")")]
                        # ()内の内容を削除
                        insert_text = function_text.replace(bracket_content, "")
                        self.insert_into_equation(insert_text)
                    else:
                        # 括弧のない関数の場合はそのまま挿入
                        self.insert_into_equation(function_text)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"関数の挿入中にエラーが発生しました: {str(e)}")

    def insert_function_from_table(self, row):
        """関数テーブルの関数をダブルクリックして挿入"""
        try:
            function_item = self.tikzFunctionsTable.item(row, 0)
            if function_item:
                function_text = function_item.text()
                
                if function_text in ["pi", "e"]:
                    self.insert_into_equation(function_text)
                else:
                    if "(" in function_text and ")" in function_text:
                        bracket_content = function_text[function_text.find("(")+1:function_text.find(")")]
                        # ()内の内容を削除
                        insert_text = function_text.replace(bracket_content, "")
                        self.insert_into_equation(insert_text)
                    else:
                        # 括弧のない関数の場合はそのまま挿入
                        self.insert_into_equation(function_text)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"関数の挿入中にエラーが発生しました: {str(e)}")

    def create_styled_delete_confirmation(self, title, message, dangerous_action_text="削除"):
        """スタイル付き削除確認ダイアログを作成（QMessageBox使用）"""
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle(title)
        msgBox.setText(message)
        msgBox.setIcon(QMessageBox.Warning)
        
        # カスタムボタンを追加
        delete_button = msgBox.addButton(dangerous_action_text, QMessageBox.YesRole)
        cancel_button = msgBox.addButton("キャンセル", QMessageBox.NoRole)
        
        # 削除ボタンを赤色に設定（破壊的操作）
        delete_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: 2px solid #e74c3c;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #c0392b;
                border-color: #c0392b;
                transform: translateY(-1px);
            }
            QPushButton:pressed {
                background-color: #a93226;
                border-color: #a93226;
            }
        """)
        
        # キャンセルボタンを安全な色に設定
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: 2px solid #95a5a6;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
                border-color: #7f8c8d;
            }
            QPushButton:pressed {
                background-color: #6c7b7d;
                border-color: #6c7b7d;
            }
        """)
        
        # デフォルトボタンを安全な選択肢（キャンセル）に設定
        msgBox.setDefaultButton(cancel_button)
        
        return msgBox, delete_button, cancel_button


# For standalone execution
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TikZPlotTab()
    window.show()
    sys.exit(app.exec_())
