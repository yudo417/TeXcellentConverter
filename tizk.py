import os
import numpy as np
import pandas as pd
import sys
import math  # 数学関数を使用するためにインポート
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


class TikZPlotConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # データと状態の初期化
        self.datasets = []  # 複数のデータセットを格納するリスト
        self.current_dataset_index = -1  # 現在選択されているデータセットのインデックス
        
        # グラフ全体の設定（すべてのデータセットに共通）
        self.global_settings = {
            'x_label': 'x',          # X軸ラベル
            'y_label': 'y',          # Y軸ラベル
            'x_min': 0,              # X軸最小値
            'x_max': 10,             # X軸最大値
            'y_min': 0,              # Y軸最小値
            'y_max': 10,             # Y軸最大値
            'grid': True,            # グリッド表示
            'legend_pos': 'north east', # 凡例位置
            'caption': 'グラフのタイトル', # キャプション
            'label': 'fig:plot',     # ラベル
            'position': 'H',         # 図の位置
            'width': 0.8,            # 図の幅
            'height': 0.5            # 図の高さ
        }
        
        # UIを初期化
        self.initUI()
        
        # 初期データセットを追加（UIの初期化後に呼び出す）
        QTimer.singleShot(0, lambda: self.add_dataset("データセット1"))
        
    # QColorオブジェクトをTikZ互換のRGB形式に変換するヘルパー関数
    def color_to_tikz_rgb(self, color):
        """QColorオブジェクトをTikZ互換のRGB形式に変換する"""
        # red, green, blueの一般的な色名はそのまま使用
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
        # それ以外の色はRGB値として変換（0-255から0-1へ）
        else:
            r = color.red() / 255.0
            g = color.green() / 255.0
            b = color.blue() / 255.0
            return f"color = {{rgb,255:red,{color.red()};green,{color.green()};blue,{color.blue()}}}"
            
    def initUI(self):
        self.setWindowTitle('TikZPlot Converter')
        
        # 画面サイズを取得
        screen = QApplication.primaryScreen().geometry()
        max_width = int(screen.width() * 0.95)
        max_height = int(screen.height() * 0.95)
        initial_height = int(screen.height() * 0.9)
        initial_width = int(screen.width() * 0.4)
        self.resize(min(initial_width, max_width), min(initial_height, max_height))
        self.setMinimumSize(600, 600)
        self.move(50, 50)
        
        # メインウィジェットとレイアウト
        mainWidget = QWidget()
        mainLayout = QVBoxLayout()
        
        # 上下に分割するスプリッター
        splitter = QSplitter(Qt.Vertical)
        
        # --- 上部：設定部分 ---
        settingsWidget = QWidget()
        settingsLayout = QVBoxLayout(settingsWidget)
        
        # 注意書き用のレイアウト
        infoLayout = QVBoxLayout()
        
        # 注意書き用のラベル - 赤色テキスト
        infoLabel = QLabel("注意: このグラフを使用するには、LaTeXドキュメントのプリアンブルに以下を追加してください:")
        infoLabel.setStyleSheet("color: #cc0000; font-weight: bold;") # 赤色のテキスト、太字
        
        # パッケージ名と複コピーボタンの水平レイアウト
        packageLayout = QHBoxLayout()
        
        # パッケージ名用のラベル - 普通の色
        packageLabel = QLabel("\\usepackage{pgfplots}\\usepackage{float}\\pgfplotsset{compat=1.18}")
        packageLabel.setStyleSheet("font-family: monospace;") # 等幅フォント
        
        # パッケージ名をコピーするボタン
        copyPackageButton = QPushButton("コピー")
        copyPackageButton.setFixedWidth(60)
        def copy_package():
            clipboard = QApplication.clipboard()
            clipboard.setText("\\usepackage{pgfplots}\n\\usepackage{float}\n\\pgfplotsset{compat=1.18}")
            self.statusBar.showMessage("パッケージ名をコピーしました", 3000)
        copyPackageButton.clicked.connect(copy_package)
        
        # パッケージレイアウトに追加
        packageLayout.addWidget(packageLabel)
        packageLayout.addWidget(copyPackageButton)
        packageLayout.addStretch()
        
        # 注意レイアウトに追加
        infoLayout.addWidget(infoLabel)
        infoLayout.addLayout(packageLayout)
        
        # データセット管理セクション
        datasetGroup = QGroupBox("データセット管理")
        datasetLayout = QVBoxLayout()
        
        # データセットリスト
        self.datasetList = QListWidget()
        self.datasetList.currentRowChanged.connect(self.on_dataset_selected)
        self.datasetList.setMinimumHeight(100)  # 高さを150pxに縮小
        self.datasetList.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        datasetLayout.addWidget(QLabel("データセット:"))
        datasetLayout.addWidget(self.datasetList)
        
        # データセット操作ボタン
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
        
        # タブウィジェットの作成
        tabWidget = QTabWidget()
        
        # タブ1: データ入力
        dataTab = QWidget()
        dataTabLayout = QVBoxLayout()
        
        # データソースタイプの選択（実測値か数式か）
        dataSourceTypeGroup = QGroupBox("データソースタイプ")
        dataSourceTypeLayout = QVBoxLayout()
        
        # ラジオボタン
        self.measuredRadio = QRadioButton("実測値データ（CSV/Excel/手入力）")
        self.formulaRadio = QRadioButton("数式によるグラフ生成")
        self.measuredRadio.setChecked(True)  # デフォルトは実測値
        
        # ラジオボタングループ
        dataSourceTypeButtonGroup = QButtonGroup(self)
        dataSourceTypeButtonGroup.addButton(self.measuredRadio)
        dataSourceTypeButtonGroup.addButton(self.formulaRadio)
        
        # ラジオボタンのスタイル設定
        radioStyle = "QRadioButton { font-weight: bold; font-size: 14px; padding: 5px; }"
        self.measuredRadio.setStyleSheet(radioStyle)
        self.formulaRadio.setStyleSheet(radioStyle)
        
        # 説明ラベル
        measuredLabel = QLabel("実験などで取得した実測データを使用してグラフを生成します")
        measuredLabel.setStyleSheet("color: gray; font-style: italic; padding-left: 20px; min-height: 20px;")
        formulaLabel = QLabel("入力した数式に基づいて理論曲線を生成します")
        formulaLabel.setStyleSheet("color: gray; font-style: italic; padding-left: 20px; min-height: 20px;")
        
        # レイアウトに追加
        dataSourceTypeLayout.addWidget(self.measuredRadio)
        dataSourceTypeLayout.addWidget(measuredLabel)
        dataSourceTypeLayout.addWidget(self.formulaRadio)
        dataSourceTypeLayout.addWidget(formulaLabel)
        
        # 高さを固定化
        dataSourceTypeGroup.setFixedHeight(150)
        
        # イベントハンドラーの接続
        self.measuredRadio.toggled.connect(self.on_data_source_type_changed)
        self.formulaRadio.toggled.connect(self.on_data_source_type_changed)
        
        dataSourceTypeGroup.setLayout(dataSourceTypeLayout)
        dataTabLayout.addWidget(dataSourceTypeGroup)
        
        # データソースコンテナ（実測値）
        self.measuredContainer = QWidget()
        measuredLayout = QVBoxLayout(self.measuredContainer)
        
        # データソース選択
        dataSourceGroup = QGroupBox("データソース")
        dataSourceLayout = QVBoxLayout()
        
        # CSVファイル選択
        csvLayout = QHBoxLayout()
        self.csvRadio = QRadioButton("CSVファイル:")
        self.csvRadio.setChecked(True)
        self.fileEntry = QLineEdit()
        browseButton = QPushButton('参照...')
        browseButton.clicked.connect(self.browse_csv_file)
        csvLayout.addWidget(self.csvRadio)
        csvLayout.addWidget(self.fileEntry)
        csvLayout.addWidget(browseButton)
        
        # Excelファイル選択
        excelLayout = QHBoxLayout()
        self.excelRadio = QRadioButton("Excelファイル:")
        self.excelEntry = QLineEdit()
        excelBrowseButton = QPushButton('参照...')
        excelBrowseButton.clicked.connect(self.browse_excel_file)
        excelLayout.addWidget(self.excelRadio)
        excelLayout.addWidget(self.excelEntry)
        excelLayout.addWidget(excelBrowseButton)
        
        # シート選択（Excelの場合）
        sheetLayout = QHBoxLayout()
        sheetLabel = QLabel('シート名:')
        sheetLabel.setIndent(30)  # インデント追加
        self.sheetCombobox = QComboBox()
        self.sheetCombobox.setEnabled(False)  # 初期状態では無効
        sheetLayout.addWidget(sheetLabel)
        sheetLayout.addWidget(self.sheetCombobox)
        
        # データ直接入力
        self.manualRadio = QRadioButton("データを直接入力:")
        
        # ボタングループの設定
        sourceGroup = QButtonGroup(self)
        sourceGroup.addButton(self.csvRadio)
        sourceGroup.addButton(self.excelRadio)
        sourceGroup.addButton(self.manualRadio)
        sourceGroup.buttonClicked.connect(self.toggle_source_fields)
        
        # データテーブル
        self.dataTable = QTableWidget(10, 2)  # 行数, 列数
        self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
        self.dataTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dataTable.setEnabled(False)  # 初期状態では無効
        self.dataTable.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        # データテーブル操作ボタン
        tableButtonLayout = QHBoxLayout()
        addRowButton = QPushButton('行を追加')
        addRowButton.clicked.connect(self.add_table_row)
        removeRowButton = QPushButton('選択行を削除')
        removeRowButton.clicked.connect(self.remove_table_row)
        tableButtonLayout.addWidget(addRowButton)
        tableButtonLayout.addWidget(removeRowButton)
        
        # データソースレイアウトに追加
        dataSourceLayout.addLayout(csvLayout)
        dataSourceLayout.addLayout(excelLayout)
        dataSourceLayout.addLayout(sheetLayout)
        dataSourceLayout.addWidget(self.manualRadio)
        dataSourceLayout.addWidget(self.dataTable)
        dataSourceLayout.addLayout(tableButtonLayout)
        dataSourceGroup.setLayout(dataSourceLayout)
        
        # 列選択グループ
        columnGroup = QGroupBox("CSV/Excel列の選択")
        columnLayout = QGridLayout()
        
        xColLabel = QLabel('X軸データ列:')
        self.xColCombo = QComboBox()
        columnLayout.addWidget(xColLabel, 0, 0)
        columnLayout.addWidget(self.xColCombo, 0, 1)
        
        yColLabel = QLabel('Y軸データ列:')
        self.yColCombo = QComboBox()
        columnLayout.addWidget(yColLabel, 1, 0)
        columnLayout.addWidget(self.yColCombo, 1, 1)
        
        columnHelpLabel = QLabel('※ CSVまたはExcelファイルからデータを読み込む際の列を選択します')
        columnHelpLabel.setStyleSheet('color: gray; font-style: italic;')
        columnLayout.addWidget(columnHelpLabel, 2, 0, 1, 2)
        
        # 列選択グループのレイアウトを設定
        columnGroup.setLayout(columnLayout)
        
        # データ確定ボタン（すべての入力方法で共通）- 別の場所に配置
        dataActionGroup = QGroupBox("データ確定")
        dataActionLayout = QVBoxLayout()
        
        # 注意書きラベルを追加
        actionNoteLabel = QLabel("※ CSVファイル、Excelファイル、手入力のいずれの方法でも、データを保存するには下のボタンを押してください")
        actionNoteLabel.setStyleSheet("color: #cc0000; font-weight: bold;")
        actionNoteLabel.setWordWrap(True)
        dataActionLayout.addWidget(actionNoteLabel)
        
        # ボタンのテキストはどの入力方法でも統一
        self.loadDataButton = QPushButton('データを確定・保存')
        self.loadDataButton.setToolTip('入力したデータを現在のデータセットに保存します')
        self.loadDataButton.clicked.connect(self.load_data)
        self.loadDataButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        dataActionLayout.addWidget(self.loadDataButton)
        
        dataActionGroup.setLayout(dataActionLayout)
        
        # 実測値コンテナに追加
        measuredLayout.addWidget(dataSourceGroup)
        measuredLayout.addWidget(columnGroup)
        measuredLayout.addWidget(dataActionGroup)  # 新しいグループを追加
        
        # データソースコンテナ（数式）
        self.formulaContainer = QWidget()
        formulaLayout = QVBoxLayout(self.formulaContainer)
        
        # 数式フォームグループ
        formulaFormGroup = QGroupBox("数式入力")
        formulaFormLayout = QVBoxLayout()
        
        # 数式入力フィールド
        equationLayout = QHBoxLayout()
        equationLabel = QLabel('数式 (x変数を使用):')
        self.equationEntry = QLineEdit('x^2')
        self.equationEntry.setPlaceholderText('例: x^2, sin(x), exp(-x/2)')
        equationLayout.addWidget(equationLabel)
        equationLayout.addWidget(self.equationEntry)
        formulaFormLayout.addLayout(equationLayout)
        
        # ドメイン範囲
        domainLayout = QHBoxLayout()
        domainLabel = QLabel('x軸範囲:')
        self.domainMinSpin = QDoubleSpinBox()
        self.domainMinSpin.setRange(-1000, 1000)
        self.domainMinSpin.setValue(0)
        self.domainMaxSpin = QDoubleSpinBox()
        self.domainMaxSpin.setRange(-1000, 1000)
        self.domainMaxSpin.setValue(10)
        domainLayout.addWidget(domainLabel)
        domainLayout.addWidget(self.domainMinSpin)
        domainLayout.addWidget(QLabel('〜'))
        domainLayout.addWidget(self.domainMaxSpin)
        formulaFormLayout.addLayout(domainLayout)
        
        # サンプル数
        samplesLayout = QHBoxLayout()
        samplesLabel = QLabel('サンプル数:')
        self.samplesSpin = QSpinBox()
        self.samplesSpin.setRange(10, 1000)
        self.samplesSpin.setValue(200)
        samplesLayout.addWidget(samplesLabel)
        samplesLayout.addWidget(self.samplesSpin)
        formulaFormLayout.addLayout(samplesLayout)
        
        # 数式説明
        formulaInfoLabel = QLabel(
            '※ <span style="color:black;">掛け算は必ず <b><span style="color:red;">*</span></b> を明示してください（例: 2*x, (x+1)*y）</span><br>'
            '数式内では以下の関数と演算が使用可能です：<br>'
            'sin, cos, tan, exp, ln, log, sqrt, ^（累乗）, +, -, *, /'
        )
        formulaInfoLabel.setStyleSheet(
            'background-color: #fffbe6; color: black; border: 1px solid #ffe082; border-radius: 4px; padding: 8px;'
        )
        formulaFormLayout.addWidget(formulaInfoLabel)
        
        formulaFormGroup.setLayout(formulaFormLayout)
        formulaLayout.addWidget(formulaFormGroup)
        
        # 新規追加: 数式オプションセクション
        formulaOptionsGroup = QGroupBox("数式オプション設定")
        formulaOptionsLayout = QGridLayout()
        
        # 微分・積分機能を削除してシンプルにする
        # 接線表示オプションのみ残す
        self.showTangentCheck = QCheckBox('特定点における接線を表示')
        self.showTangentCheck.setChecked(False)
        formulaOptionsLayout.addWidget(self.showTangentCheck, 0, 0, 1, 2)
        
        # 接線のx座標
        tangentXLabel = QLabel('接線のx座標:')
        self.tangentXSpin = QDoubleSpinBox()
        self.tangentXSpin.setRange(-1000, 1000)
        self.tangentXSpin.setValue(5)  # デフォルト値
        formulaOptionsLayout.addWidget(tangentXLabel, 1, 0)
        formulaOptionsLayout.addWidget(self.tangentXSpin, 1, 1)
        
        # 接線の長さ
        tangentLengthLabel = QLabel('接線の長さ:')
        self.tangentLengthSpin = QDoubleSpinBox()
        self.tangentLengthSpin.setRange(0.1, 20.0)
        self.tangentLengthSpin.setValue(2.0)  # デフォルト値
        formulaOptionsLayout.addWidget(tangentLengthLabel, 2, 0)
        formulaOptionsLayout.addWidget(self.tangentLengthSpin, 2, 1)
        
        # 接線の色
        tangentColorLabel = QLabel('接線の色:')
        self.tangentColorCombo = QComboBox()
        common_colors = ['赤', '青', '緑', '紫', '橙', '黒', '水色', 'マゼンタ', '茶色', '灰色']
        self.tangentColorCombo.addItems(common_colors)
        self.tangentColorCombo.setCurrentText('紫')
        self.tangentColorButton = QPushButton('カスタム色...')
        self.tangentColorButton.clicked.connect(self.select_tangent_color)
        tangentColorLayout = QHBoxLayout()
        tangentColorLayout.addWidget(self.tangentColorCombo)
        tangentColorLayout.addWidget(self.tangentColorButton)
        formulaOptionsLayout.addWidget(tangentColorLabel, 3, 0)
        formulaOptionsLayout.addLayout(tangentColorLayout, 3, 1)
        self.tangentColor = QColor('purple')  # デフォルト値
        self.tangentColorCombo.currentTextChanged.connect(self.on_tangent_color_changed)
        
        # 接線の線スタイル
        tangentStyleLabel = QLabel('接線のスタイル:')
        self.tangentStyleCombo = QComboBox()
        self.tangentStyleCombo.addItems(['実線', '点線', '破線', '一点鎖線'])
        formulaOptionsLayout.addWidget(tangentStyleLabel, 4, 0)
        formulaOptionsLayout.addWidget(self.tangentStyleCombo, 4, 1)
        
        # 接線の式表示設定
        self.showTangentEquationCheck = QCheckBox('接線の方程式を表示')
        self.showTangentEquationCheck.setChecked(False)
        self.showTangentEquationCheck.setToolTip('グラフ上に接線の方程式（y = ax + b）を表示します')
        formulaOptionsLayout.addWidget(self.showTangentEquationCheck, 5, 0, 1, 2)
        
        formulaOptionsGroup.setLayout(formulaOptionsLayout)
        formulaLayout.addWidget(formulaOptionsGroup)
        
        # 新規追加: TikZ数式ガイドセクション
        tikzGuideGroup = QGroupBox("TikZ数式ガイド")
        tikzGuideLayout = QVBoxLayout()
        
        # ガイド説明
        guideLabel = QLabel('TikZで使える特殊な数学関数:')
        guideLabel.setStyleSheet("font-weight: bold;")
        tikzGuideLayout.addWidget(guideLabel)
        
        # 関数一覧と説明をテーブルで表示
        self.tikzFunctionsTable = QTableWidget(0, 2)  # 行数は動的に設定
        self.tikzFunctionsTable.setHorizontalHeaderLabels(['関数名', '説明と使用例'])
        self.tikzFunctionsTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tikzFunctionsTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tikzFunctionsTable.verticalHeader().setVisible(False)  # 行番号を非表示
        self.tikzFunctionsTable.setEditTriggers(QTableWidget.NoEditTriggers)  # 編集不可
        
        # TikZで使える関数のリスト
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
        
        # テーブルに関数を追加
        self.tikzFunctionsTable.setRowCount(len(tikz_functions))
        for i, (func, desc) in enumerate(tikz_functions):
            self.tikzFunctionsTable.setItem(i, 0, QTableWidgetItem(func))
            self.tikzFunctionsTable.setItem(i, 1, QTableWidgetItem(desc))
            
        # テーブルの高さを調整（全ての行を表示）
        table_height = self.tikzFunctionsTable.horizontalHeader().height()
        for i in range(len(tikz_functions)):
            table_height += self.tikzFunctionsTable.rowHeight(i)
        self.tikzFunctionsTable.setMinimumHeight(min(250, table_height))  # 最大高さを制限
        
        tikzGuideLayout.addWidget(self.tikzFunctionsTable)
        
        # テーブルの説明ラベルを追加
        tableHelpLabel = QLabel('※ 関数をダブルクリックすると数式に挿入できます')
        tableHelpLabel.setStyleSheet("color: #666666; font-style: italic;")
        tikzGuideLayout.addWidget(tableHelpLabel)
        
        # 関数をダブルクリックで挿入
        self.tikzFunctionsTable.cellDoubleClicked.connect(self.insert_function_from_table)
        
        tikzGuideGroup.setLayout(tikzGuideLayout)
        formulaLayout.addWidget(tikzGuideGroup)
        
        # データ確定ボタン（数式入力モード用）
        formulaDataActionGroup = QGroupBox("データ確定")
        formulaDataActionLayout = QVBoxLayout()
        
        # 注意書きラベル
        formulaActionNoteLabel = QLabel("※ 数式に基づくグラフを生成するには、下のボタンを押して数式データを保存してください")
        formulaActionNoteLabel.setStyleSheet("color: #cc0000; font-weight: bold;")
        formulaActionNoteLabel.setWordWrap(True)
        formulaDataActionLayout.addWidget(formulaActionNoteLabel)
        
        # 数式データ確定ボタン
        self.formulaDataButton = QPushButton('数式データを確定・保存')
        self.formulaDataButton.setToolTip('入力した数式に基づいてデータを生成し、現在のデータセットに保存します')
        self.formulaDataButton.clicked.connect(self.apply_formula)
        self.formulaDataButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        formulaDataActionLayout.addWidget(self.formulaDataButton)
        
        formulaDataActionGroup.setLayout(formulaDataActionLayout)
        formulaLayout.addWidget(formulaDataActionGroup)
        
        # 初期状態ではコンテナの表示/非表示を設定
        dataTabLayout.addWidget(self.measuredContainer)
        dataTabLayout.addWidget(self.formulaContainer)
        self.measuredContainer.setVisible(True)
        self.formulaContainer.setVisible(False)
        
        # タブに設定
        dataTab.setLayout(dataTabLayout)
        
        # タブ2: グラフ設定
        plotTab = QWidget()
        plotTabLayout = QVBoxLayout()
        
        # グラフタイプの選択
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
        
        # グラフスタイル設定
        styleGroup = QGroupBox("データセット個別設定 - スタイル")
        styleLayout = QGridLayout()
        
        # データソースタイプ説明ラベル
        dataSourceTypeLabel = QLabel("選択中のデータセットタイプ:")
        self.dataSourceTypeDisplayLabel = QLabel("実測データ")
        self.dataSourceTypeDisplayLabel.setStyleSheet("font-weight: bold;")
        styleLayout.addWidget(dataSourceTypeLabel, 0, 0)
        styleLayout.addWidget(self.dataSourceTypeDisplayLabel, 0, 1)
        
        # 色選択
        colorLabel = QLabel('線/点の色:')
        self.colorButton = QPushButton()
        self.colorButton.setStyleSheet('background-color: blue;')
        self.currentColor = QColor('blue')
        self.colorButton.clicked.connect(self.select_color)
        styleLayout.addWidget(colorLabel, 1, 0)
        styleLayout.addWidget(self.colorButton, 1, 1)
        
        # 線の太さ
        lineWidthLabel = QLabel('線の太さ:')
        self.lineWidthSpin = QDoubleSpinBox()
        self.lineWidthSpin.setRange(0.1, 5.0)
        self.lineWidthSpin.setSingleStep(0.1)
        self.lineWidthSpin.setValue(1.0)
        styleLayout.addWidget(lineWidthLabel, 2, 0)
        styleLayout.addWidget(self.lineWidthSpin, 2, 1)
        
        # マーカースタイル
        markerLabel = QLabel('マーカースタイル:')
        self.markerCombo = QComboBox()
        self.markerCombo.addItems(['*', 'o', 'square', 'triangle', 'diamond', '+', 'x'])
        styleLayout.addWidget(markerLabel, 3, 0)
        styleLayout.addWidget(self.markerCombo, 3, 1)
        
        # マーカーサイズ
        markerSizeLabel = QLabel('マーカーサイズ:')
        self.markerSizeSpin = QDoubleSpinBox()
        self.markerSizeSpin.setRange(0.5, 10.0)
        self.markerSizeSpin.setSingleStep(0.5)
        self.markerSizeSpin.setValue(3.0)
        styleLayout.addWidget(markerSizeLabel, 4, 0)
        styleLayout.addWidget(self.markerSizeSpin, 4, 1)
        
        # データ点をマークで表示するオプション
        self.showDataPointsCheck = QCheckBox('データ点をマークで表示（線グラフでも点を表示）')
        self.showDataPointsCheck.setChecked(False)
        styleLayout.addWidget(self.showDataPointsCheck, 5, 0, 1, 2)
        
        styleGroup.setLayout(styleLayout)
        
        # 軸設定
        axisGroup = QGroupBox("グラフ全体設定 - 軸")
        axisLayout = QGridLayout()
        
        # X軸ラベル
        xLabelLabel = QLabel('X軸ラベル:')
        self.xLabelEntry = QLineEdit(self.global_settings['x_label'])
        axisLayout.addWidget(xLabelLabel, 0, 0)
        axisLayout.addWidget(self.xLabelEntry, 0, 1)
        
        # X軸目盛り間隔
        xTickStepLabel = QLabel('X軸目盛り間隔:')
        self.xTickStepSpin = QDoubleSpinBox()
        self.xTickStepSpin.setRange(0.01, 1000)
        self.xTickStepSpin.setSingleStep(0.1)
        self.xTickStepSpin.setValue(1.0)
        axisLayout.addWidget(xTickStepLabel, 0, 2)
        axisLayout.addWidget(self.xTickStepSpin, 0, 3)
        
        # Y軸ラベル
        yLabelLabel = QLabel('Y軸ラベル:')
        self.yLabelEntry = QLineEdit(self.global_settings['y_label'])
        axisLayout.addWidget(yLabelLabel, 1, 0)
        axisLayout.addWidget(self.yLabelEntry, 1, 1)
        
        # Y軸目盛り間隔
        yTickStepLabel = QLabel('Y軸目盛り間隔:')
        self.yTickStepSpin = QDoubleSpinBox()
        self.yTickStepSpin.setRange(0.01, 1000)
        self.yTickStepSpin.setSingleStep(0.1)
        self.yTickStepSpin.setValue(1.0)
        axisLayout.addWidget(yTickStepLabel, 1, 2)
        axisLayout.addWidget(self.yTickStepSpin, 1, 3)

        # X軸範囲
        xRangeLabel = QLabel('X軸範囲:')
        xRangeLayout = QHBoxLayout()
        self.xMinSpin = QDoubleSpinBox()
        self.xMinSpin.setRange(-1000, 1000)
        self.xMinSpin.setValue(self.global_settings['x_min'])
        self.xMaxSpin = QDoubleSpinBox()
        self.xMaxSpin.setRange(-1000, 1000)
        self.xMaxSpin.setValue(self.global_settings['x_max'])
        xRangeLayout.addWidget(self.xMinSpin)
        xRangeLayout.addWidget(QLabel('〜'))
        xRangeLayout.addWidget(self.xMaxSpin)
        axisLayout.addWidget(xRangeLabel, 2, 0)
        axisLayout.addLayout(xRangeLayout, 2, 1)
        
        # Y軸範囲
        yRangeLabel = QLabel('Y軸範囲:')
        yRangeLayout = QHBoxLayout()
        self.yMinSpin = QDoubleSpinBox()
        self.yMinSpin.setRange(-1000, 1000)
        self.yMinSpin.setValue(self.global_settings['y_min'])
        self.yMaxSpin = QDoubleSpinBox()
        self.yMaxSpin.setRange(-1000, 1000)
        self.yMaxSpin.setValue(self.global_settings['y_max'])
        yRangeLayout.addWidget(self.yMinSpin)
        yRangeLayout.addWidget(QLabel('〜'))
        yRangeLayout.addWidget(self.yMaxSpin)
        axisLayout.addWidget(yRangeLabel, 3, 0)
        axisLayout.addLayout(yRangeLayout, 3, 1)

        # 目盛り間隔の値を保持する変数を初期化
        self.x_tick_step = 1.0
        self.y_tick_step = 1.0

        # グリッド表示
        self.gridCheck = QCheckBox('グリッド表示')
        self.gridCheck.setChecked(self.global_settings['grid'])
        axisLayout.addWidget(self.gridCheck, 4, 0, 1, 2)
        
        axisGroup.setLayout(axisLayout)
        
        # 凡例設定
        legendGroup = QGroupBox("データセット個別設定 - 凡例")
        legendLayout = QFormLayout()
        
        self.legendCheck = QCheckBox('凡例を表示')
        self.legendCheck.setChecked(True)
        
        self.legendLabel = QLineEdit('データ')
        
        legendPosLabel = QLabel('凡例の位置:')
        self.legendPosCombo = QComboBox()
        self.legendPosCombo.addItems(['north west', 'north east', 'south west', 'south east', 'north', 'south', 'east', 'west'])
        self.legendPosCombo.setCurrentText(self.global_settings['legend_pos'])
        
        legendLayout.addRow(self.legendCheck)
        legendLayout.addRow('凡例ラベル:', self.legendLabel)
        legendLayout.addRow(legendPosLabel, self.legendPosCombo)
        
        legendGroup.setLayout(legendLayout)
        
        # 図の設定
        figureGroup = QGroupBox("グラフ全体設定 - 図")
        figureLayout = QFormLayout()
        
        # キャプション
        self.captionEntry = QLineEdit(self.global_settings['caption'])
        
        # ラベル
        self.labelEntry = QLineEdit(self.global_settings['label'])
        
        # 位置
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        self.positionCombo.setCurrentText(self.global_settings['position'])
        
        # 幅と高さ
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
        
        # プロットタブレイアウトに追加
        plotTabLayout.addWidget(QLabel("【グラフ全体の設定】"))
        plotTabLayout.addWidget(axisGroup)
        plotTabLayout.addWidget(figureGroup)
        
        # 区切り線
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        plotTabLayout.addWidget(separator)
        
        plotTabLayout.addWidget(QLabel("【データセット個別の設定】"))
        plotTabLayout.addWidget(plotTypeGroup)
        plotTabLayout.addWidget(styleGroup)
        plotTabLayout.addWidget(legendGroup)
        
        plotTab.setLayout(plotTabLayout)
        
        # タブ4: 特殊点・注釈
        annotationTab = QWidget()
        annotationTabLayout = QVBoxLayout()
        
        # 特殊点グループ
        specialPointsGroup = QGroupBox("特殊点")
        specialPointsLayout = QVBoxLayout()
        
        self.specialPointsCheck = QCheckBox('特殊点を表示')
        specialPointsLayout.addWidget(self.specialPointsCheck)
        
        # 特殊点テーブル
        self.specialPointsTable = QTableWidget(0, 4)  # 行数, 列数 (X, Y, 色, 座標表示)
        self.specialPointsTable.setHorizontalHeaderLabels(['X', 'Y', '色', '座標表示'])
        self.specialPointsTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.specialPointsTable.setEnabled(False)
        specialPointsLayout.addWidget(self.specialPointsTable)
        
        # 特殊点操作ボタン
        spButtonLayout = QHBoxLayout()
        addSpButton = QPushButton('特殊点を追加')
        addSpButton.clicked.connect(self.add_special_point)
        removeSpButton = QPushButton('選択した特殊点を削除')
        removeSpButton.clicked.connect(self.remove_special_point)
        assignToDatasetBtn = QPushButton('データセットに割り当て')
        assignToDatasetBtn.clicked.connect(self.assign_special_points_to_dataset)
        spButtonLayout.addWidget(addSpButton)
        spButtonLayout.addWidget(removeSpButton)
        spButtonLayout.addWidget(assignToDatasetBtn)
        specialPointsLayout.addLayout(spButtonLayout)
        
        # イベント接続
        self.specialPointsCheck.toggled.connect(lambda checked: [
            self.specialPointsTable.setEnabled(checked),
            addSpButton.setEnabled(checked),
            removeSpButton.setEnabled(checked),
            assignToDatasetBtn.setEnabled(checked)
        ])
        
        specialPointsGroup.setLayout(specialPointsLayout)
        
        # 注釈グループ
        annotationsGroup = QGroupBox("注釈")
        annotationsLayout = QVBoxLayout()
        
        self.annotationsCheck = QCheckBox('注釈を表示')
        annotationsLayout.addWidget(self.annotationsCheck)
        
        # 注釈テーブル
        self.annotationsTable = QTableWidget(0, 5)  # 行数, 列数 (X, Y, テキスト, 色, 位置)
        self.annotationsTable.setHorizontalHeaderLabels(['X', 'Y', 'テキスト', '色', '位置'])
        self.annotationsTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.annotationsTable.setEnabled(False)
        annotationsLayout.addWidget(self.annotationsTable)
        
        # 注釈操作ボタン
        annButtonLayout = QHBoxLayout()
        addAnnButton = QPushButton('注釈を追加')
        addAnnButton.clicked.connect(self.add_annotation)
        removeAnnButton = QPushButton('選択した注釈を削除')
        removeAnnButton.clicked.connect(self.remove_annotation)
        assignAnnToDatasetBtn = QPushButton('データセットに割り当て')
        assignAnnToDatasetBtn.clicked.connect(self.assign_annotations_to_dataset)
        annButtonLayout.addWidget(addAnnButton)
        annButtonLayout.addWidget(removeAnnButton)
        annButtonLayout.addWidget(assignAnnToDatasetBtn)
        annotationsLayout.addLayout(annButtonLayout)
        
        # イベント接続
        self.annotationsCheck.toggled.connect(lambda checked: [
            self.annotationsTable.setEnabled(checked),
            addAnnButton.setEnabled(checked),
            removeAnnButton.setEnabled(checked),
            assignAnnToDatasetBtn.setEnabled(checked)
        ])
        
        annotationsGroup.setLayout(annotationsLayout)
        
        # 注釈タブレイアウトに追加
        annotationTabLayout.addWidget(specialPointsGroup)
        annotationTabLayout.addWidget(annotationsGroup)
        annotationTab.setLayout(annotationTabLayout)
        
        # タブ追加
        tabWidget.addTab(dataTab, "データ入力")
        tabWidget.addTab(plotTab, "グラフ設定")
        tabWidget.addTab(annotationTab, "特殊点・注釈設定")
        
        # --- グラフ設定タブの値が変わったら即保存 ---
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
        
        # 設定部分のレイアウトに追加
        settingsLayout.addLayout(infoLayout)
        settingsLayout.addWidget(datasetGroup)
        settingsLayout.addWidget(tabWidget)

        # データテーブルの高さを減らす
        self.dataTable.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        
        # LaTeXコードに変換ボタン
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
        copyButton.clicked.connect(self.copy_to_clipboard)
        
        resultLayout.addWidget(resultLabel)
        resultLayout.addWidget(self.resultText)
        resultLayout.addWidget(copyButton)
        
        # スプリッターに追加
        splitter.addWidget(settingsWidget)
        splitter.addWidget(resultWidget)
        splitter.setSizes([600, 200])  # 初期サイズ比率を調整
        
        # メインレイアウトに追加
        mainLayout.addWidget(splitter)
        
        # ステータスバー
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # メインウィジェット設定
        mainWidget.setLayout(mainLayout)
        # QScrollAreaでラップ
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(mainWidget)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setCentralWidget(scroll)
        
        # スクロールヒントラベルの上に大きな下向き矢印を表示
        # arrowLabel = QLabel('⬇️')
        # arrowLabel.setAlignment(Qt.AlignHCenter)
        # arrowLabel.setStyleSheet('font-size: 32px; color: #d32f2f; margin-bottom: 0px;')
        # arrowLabel.setFixedHeight(40)
        # mainLayout.addWidget(arrowLabel)

        scrollHint = QLabel('画面に収まらない場合はスクロールバーで全体を表示できます')
        scrollHint.setStyleSheet('background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fffbe6, stop:1 #ffe082); color: #d32f2f; font-size: 13px; font-weight: bold; border: 1px solid #ffe082; border-radius: 4px; padding: 6px;')
        scrollHint.setAlignment(Qt.AlignHCenter)
        mainLayout.insertWidget(0, scrollHint)
        
    # CSVファイル選択ダイアログ
    def browse_csv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "CSVファイルを選択", "", "CSV Files (*.csv);;All Files (*)")
        if file_path:
            self.fileEntry.setText(file_path)
            self.csvRadio.setChecked(True)
            self.toggle_source_fields()
            self.update_column_names(file_path)
    
    # Excelファイル選択ダイアログ
    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        if file_path:
            self.excelEntry.setText(file_path)
            self.excelRadio.setChecked(True)
            self.toggle_source_fields()
            self.update_sheet_names(file_path)
    
    # シート名の更新（Excelファイル用）
    def update_sheet_names(self, file_path):
        try:
            xls = pd.ExcelFile(file_path)
            self.sheetCombobox.clear()
            self.sheetCombobox.addItems(xls.sheet_names)
            self.statusBar.showMessage(f"ファイル '{os.path.basename(file_path)}' を読み込みました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"シート名の取得に失敗しました: {str(e)}")
            self.statusBar.showMessage("ファイル読み込みエラー")
    
    # データソース選択に応じてUIを切り替え
    def toggle_source_fields(self):
        if self.csvRadio.isChecked():
            self.fileEntry.setEnabled(True)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(False)
            # CSVファイル用のツールチップ
            self.loadDataButton.setToolTip("CSVファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.excelRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(True)
            self.sheetCombobox.setEnabled(True)
            self.dataTable.setEnabled(False)
            # Excelファイル用のツールチップ
            self.loadDataButton.setToolTip("Excelファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.manualRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(True)
            # 手入力データ用のツールチップ
            self.loadDataButton.setToolTip("テーブルに入力したデータを現在のデータセットに保存します")
    
    # データを読み込む
    def load_data(self):
        """選択されたデータソースからデータを読み込む"""
        # 入力値を保存
        self.update_current_dataset()
        try:
            if self.current_dataset_index < 0:
                QMessageBox.warning(self, "警告", "データを読み込むデータセットを選択してください")
                return
                
            data_x = []
            data_y = []
            
            if self.csvRadio.isChecked():
                file_path = self.fileEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なCSVファイルを選択してください")
                    return
                
                df = pd.read_csv(file_path)
                x_col = self.xColCombo.currentText()
                y_col = self.yColCombo.currentText()
                
                if x_col in df.columns and y_col in df.columns:
                    data_x = df[x_col].values.tolist()
                    data_y = df[y_col].values.tolist()
                else:
                    QMessageBox.warning(self, "警告", "選択した列名が不正です")
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
                
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                x_col = self.xColCombo.currentText()
                y_col = self.yColCombo.currentText()
                
                if x_col in df.columns and y_col in df.columns:
                    data_x = df[x_col].values.tolist()
                    data_y = df[y_col].values.tolist()
                else:
                    QMessageBox.warning(self, "警告", "選択した列名が不正です")
                    return
                
            elif self.manualRadio.isChecked():
                data_x = []
                data_y = []
                
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
                            pass  # 数値に変換できない場合はスキップ
                
                if not data_x:
                    QMessageBox.warning(self, "警告", "有効なデータポイントがありません")
                    return
                
            # 現在のデータセットにデータを設定
            self.datasets[self.current_dataset_index]['data_x'] = data_x
            self.datasets[self.current_dataset_index]['data_y'] = data_y
            
            # データテーブルを更新（手動入力モードの場合）
            if self.manualRadio.isChecked():
                self.update_data_table_from_dataset(self.datasets[self.current_dataset_index])
            
            dataset_name = self.datasets[self.current_dataset_index]['name']
            QMessageBox.information(self, "成功", f"データセット '{dataset_name}' に{len(data_x)}個のデータポイントを読み込みました")
            self.statusBar.showMessage(f"データセット '{dataset_name}' にデータを読み込みました: {len(data_x)}ポイント")
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データ読み込み中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("データ読み込みエラー")
    
    # 特殊点を追加
    def add_special_point(self):
        row_position = self.specialPointsTable.rowCount()
        self.specialPointsTable.insertRow(row_position)
        
        # デフォルト値の設定
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        
        # 色選択コンボボックス
        color_combo = QComboBox()
        color_combo.addItems(['red', 'blue', 'green', 'black', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("red")
        
        # 座標表示コンボボックス（チェックボックスから変更）
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
    
    # 特殊点を削除
    def remove_special_point(self):
        selected_rows = set(index.row() for index in self.specialPointsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.specialPointsTable.removeRow(row)
    
    # 注釈を追加
    def add_annotation(self):
        row_position = self.annotationsTable.rowCount()
        self.annotationsTable.insertRow(row_position)
        
        # デフォルト値の設定
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        text_item = QTableWidgetItem("注釈テキスト")
        
        # 色選択コンボボックス
        color_combo = QComboBox()
        color_combo.addItems(['black', 'red', 'blue', 'green', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("black")
        
        # 位置選択コンボボックス - より直感的な表現に変更
        pos_combo = QComboBox()
        pos_combo.addItems(['上', '右上', '右', '右下', '下', '左下', '左', '左上'])
        pos_combo.setCurrentText("右上")
        
        self.annotationsTable.setItem(row_position, 0, x_item)
        self.annotationsTable.setItem(row_position, 1, y_item)
        self.annotationsTable.setItem(row_position, 2, text_item)
        self.annotationsTable.setCellWidget(row_position, 3, color_combo)
        self.annotationsTable.setCellWidget(row_position, 4, pos_combo)
    
    # 注釈を削除
    def remove_annotation(self):
        selected_rows = set(index.row() for index in self.annotationsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.annotationsTable.removeRow(row)
    
    # パラメータ値を追加
    def add_param_value(self):
        param_name = self.paramNameEntry.text() or "a"
        min_val = self.paramMinSpin.value()
        max_val = self.paramMaxSpin.value()
        step = self.paramStepSpin.value()
        
        values = np.arange(min_val, max_val + step/2, step).tolist()
        
        for val in values:
            item_text = f"{param_name} = {val:.4g}"
            if item_text not in [self.sweepCurvesList.item(i).text() for i in range(self.sweepCurvesList.count())]:
                self.sweepCurvesList.addItem(item_text)
    
    # パラメータ値を削除
    def remove_param_value(self):
        for item in self.sweepCurvesList.selectedItems():
            self.sweepCurvesList.takeItem(self.sweepCurvesList.row(item))
    
    # LaTeXコードをクリップボードにコピー
    def copy_to_clipboard(self):
        latex_code = self.resultText.toPlainText()
        if latex_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(latex_code)
            self.statusBar.showMessage("LaTeXコードをクリップボードにコピーしました", 3000)  # 3秒間表示

    def update_global_settings(self):
        """UIからグラフ全体の設定を更新する"""
        try:
            # 軸ラベル
            self.global_settings['x_label'] = self.xLabelEntry.text()
            self.global_settings['y_label'] = self.yLabelEntry.text()
            
            # 軸範囲
            self.global_settings['x_min'] = self.xMinSpin.value()
            self.global_settings['x_max'] = self.xMaxSpin.value()
            self.global_settings['y_min'] = self.yMinSpin.value()
            self.global_settings['y_max'] = self.yMaxSpin.value()
            
            # グリッド
            self.global_settings['grid'] = self.gridCheck.isChecked()
            
            # 凡例位置
            self.global_settings['legend_pos'] = self.legendPosCombo.currentText()
            
            # 図の設定
            self.global_settings['caption'] = self.captionEntry.text()
            self.global_settings['label'] = self.labelEntry.text()
            self.global_settings['position'] = self.positionCombo.currentText()
            self.global_settings['width'] = self.widthSpin.value()
            self.global_settings['height'] = self.heightSpin.value()
            
            # 目盛り間隔も反映
            self.x_tick_step = self.xTickStepSpin.value()
            self.y_tick_step = self.yTickStepSpin.value()
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"グラフ全体設定の更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")

    def convert_to_tikz(self):
        # データセットが空かチェック
        if not self.datasets or all(not dataset.get('data_x') for dataset in self.datasets):
            QMessageBox.warning(self, "警告", "データが読み込まれていません。先にデータを読み込んでください。")
            return
        
        try:
            # グラフ全体の設定を更新
            self.update_global_settings()
            
            # 各データセットを処理
            latex_code = self.generate_tikz_code_multi_datasets()
            
            self.resultText.setPlainText(latex_code)
            self.statusBar.showMessage("TikZコードが生成されました")
            
            # ポップアップを削除
            # QMessageBox.information(self, "TikZコード生成完了", 
            #                      "TikZコードが生成されました。クリップボードにコピーボタンを押して利用できます。")
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"変換中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("変換エラー")

    # データセット管理の関数
    def add_dataset(self, name_arg=None):
        """新しいデータセットを追加する"""
        try:
            # 既存のデータセットの状態を保存（存在する場合）
            if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
                self.update_current_dataset()
                
            final_name = ""
            if name_arg is None:
                dataset_count = self.datasetList.count() + 1
                # QInputDialog.getText returns (text: str, ok: bool)
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
                    final_name = str(name_arg).strip() # Ensure argument is also a string and stripped

            if not final_name: # Double check if somehow final_name is empty
                dataset_count = self.datasetList.count() + 1
                final_name = f"データセット{dataset_count}"
                self.statusBar.showMessage("データセット名が空のため、デフォルト名を使用します。", 3000)

            # 明示的に空のデータと初期設定を持つデータセットを作成
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
                # ファイル読み込み関連の設定
                'file_path': '',
                'file_type': 'csv',  # 'csv' or 'excel' or 'manual'
                'sheet_name': '',
                'x_column': '',
                'y_column': ''
                # パラメータスイープ関連の設定を削除
            }
            
            self.datasets.append(dataset)
            self.datasetList.addItem(final_name) # addItem directly with the guaranteed string
            
            # 新しく追加したデータセットを選択（これがon_dataset_selectedを呼び出す）
            self.datasetList.setCurrentRow(len(self.datasets) - 1)
            self.statusBar.showMessage(f"データセット '{final_name}' を追加しました", 3000)

        except Exception as e:
            import traceback
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
            reply = QMessageBox.question(self, "確認",
                                       f"データセット '{dataset_name}' を削除してもよろしいですか？",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.datasets.pop(current_row)
                item = self.datasetList.takeItem(current_row)
                if item:
                    del item #明示的に削除
                
                if self.datasets: # 他のデータセットがある場合
                    new_index = max(0, min(current_row, len(self.datasets) - 1))
                    self.datasetList.setCurrentRow(new_index)
                    # on_dataset_selectedが呼ばれるので、current_dataset_indexはそこで更新される
                else: # データセットがなくなった場合
                    self.current_dataset_index = -1
                    # 新しいデフォルトデータセットを追加するか、UIをクリア状態にする
                    # ここではクリア状態にし、ユーザーに手動で追加させるか、あるいは自動で追加する
                    self.update_ui_for_no_datasets() # UIを空のデータセット状態に更新するヘルパー関数を想定
                    self.add_dataset("データセット1") # Or add a new default one
                
                self.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット削除中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_for_no_datasets(self):
        """データセットがない場合にUIをリセット/クリアする"""
        # 例: 関連する入力フィールドをクリアまたは無効化
        self.legendLabel.setText("")
        # 他のUI要素も必要に応じてリセット
        # (この関数は具体的にどのようなUI状態が望ましいかによって実装を調整)
        pass # Placeholder
    
    def rename_dataset(self):
        """選択されたデータセットの名前を変更する"""
        try:
            if self.current_dataset_index < 0 or not self.datasets:
                QMessageBox.warning(self, "警告", "名前を変更するデータセットが選択されていません。")
                return
            
            current_row = self.datasetList.currentRow() # currentRowが選択されていれば正しいはず
            # current_dataset_index と current_row の一貫性を確認
            if current_row != self.current_dataset_index:
                 # 予期せぬ状態。current_dataset_index に合わせるかエラー表示
                 current_row = self.current_dataset_index
                 self.datasetList.setCurrentRow(current_row)

            current_name = str(self.datasets[current_row]['name'])
            current_legend = str(self.datasets[current_row].get('legend_label', current_name))

            new_name_text, ok = QInputDialog.getText(self, "データセット名の変更", \
                                                      "新しいデータセット名を入力してください:",
                                                      QLineEdit.Normal, current_name)
            
            if ok and new_name_text.strip():
                actual_new_name = new_name_text.strip()
                if not actual_new_name:
                    QMessageBox.warning(self, "警告", "データセット名は空にできません。")
                    return

                self.datasets[current_row]['name'] = actual_new_name
                # 凡例ラベルが元の名前と同じだった場合、凡例ラベルも更新
                if current_legend == current_name:
                    self.datasets[current_row]['legend_label'] = actual_new_name
                    self.legendLabel.setText(actual_new_name) # UIも即時更新
                
                item = self.datasetList.item(current_row)
                if item:
                    item.setText(actual_new_name)
                self.statusBar.showMessage(f"データセット名を '{actual_new_name}' に変更しました", 3000)
            elif ok and not new_name_text.strip():
                 QMessageBox.warning(self, "警告", "データセット名は空にできません。")

        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット名の変更中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def on_dataset_selected(self, row):
        try:
            # 以前のデータセットの状態を保存（存在する場合）
            old_index = self.current_dataset_index
            if old_index >= 0 and old_index < len(self.datasets):
                self.update_current_dataset()
            if row < 0 or row >= len(self.datasets):
                if not self.datasets:
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets()
                return
            # 現在のインデックスを更新
            self.current_dataset_index = row
            # UIを更新（この中でデータテーブルも更新される）
            dataset = self.datasets[row]
            # --- ここでUIのみを更新し、保存処理は呼ばない ---
            self.update_ui_from_dataset(dataset)
            self.statusBar.showMessage(f"データセット '{dataset['name']}' を選択しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_current_dataset(self):
        try:
            if self.current_dataset_index < 0 or not self.datasets or self.current_dataset_index >= len(self.datasets):
                return
            dataset = self.datasets[self.current_dataset_index]
            if not isinstance(self.currentColor, QColor):
                self.currentColor = QColor(self.currentColor)
            dataset['color'] = QColor(self.currentColor)
            dataset['color'] = QColor(self.currentColor)
            dataset['line_width'] = self.lineWidthSpin.value()
            dataset['marker_style'] = self.markerCombo.currentText()
            dataset['marker_size'] = self.markerSizeSpin.value()
            dataset['show_legend'] = self.legendCheck.isChecked()
            dataset['legend_label'] = self.legendLabel.text()
            if self.lineRadio.isChecked():
                dataset['plot_type'] = "line"
            elif self.scatterRadio.isChecked():
                dataset['plot_type'] = "scatter"
            elif self.lineScatterRadio.isChecked():
                dataset['plot_type'] = "line_scatter"
            else:
                dataset['plot_type'] = "bar"
            if hasattr(self, 'measuredRadio') and hasattr(self, 'formulaRadio'):
                dataset['data_source_type'] = 'measured' if self.measuredRadio.isChecked() else 'formula'
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
                    dataset['x_column'] = self.xColCombo.currentText()
                    dataset['y_column'] = self.yColCombo.currentText()
                    # 手入力データの保存は「データ入力タブがアクティブ」かつ「手入力モード」のときだけ
                    if (
                        hasattr(self, 'tabWidget')
                        and self.tabWidget.currentIndex() == 0  # データ入力タブ
                        and self.manualRadio.isChecked()
                    ):
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
                        dataset['data_x'] = data_x
                        dataset['data_y'] = data_y
                else:
                    dataset['equation'] = self.equationEntry.text()
                    dataset['domain_min'] = self.domainMinSpin.value()
                    dataset['domain_max'] = self.domainMaxSpin.value()
                    dataset['samples'] = self.samplesSpin.value()
                    dataset['show_tangent'] = self.showTangentCheck.isChecked()
                    dataset['tangent_x'] = self.tangentXSpin.value()
                    dataset['tangent_length'] = self.tangentLengthSpin.value()
                    tangent_color = dataset.get('tangent_color', QColor('purple'))
                    self.tangentColor = QColor(tangent_color)
                    tangent_color_name = dataset.get('tangent_color_name', '紫')
                    index = self.tangentColorCombo.findText(tangent_color_name)
                    if index >= 0:
                        self.tangentColorCombo.setCurrentIndex(index)
                    else:
                        if 'tangent_color' in dataset:
                            custom_index = self.tangentColorCombo.findText('カスタム')
                            if custom_index == -1:
                                self.tangentColorCombo.addItem('カスタム')
                                custom_index = self.tangentColorCombo.findText('カスタム')
                            self.tangentColorCombo.setCurrentIndex(custom_index)
                    self.tangentColorButton.setStyleSheet(f'background-color: {self.tangentColor.name()};')
                    tangent_style = dataset.get('tangent_style', '実線')
                    index = self.tangentStyleCombo.findText(tangent_style)
                    if index >= 0:
                        self.tangentStyleCombo.setCurrentIndex(index)
                    self.showTangentEquationCheck.setChecked(dataset.get('show_tangent_equation', False))
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_from_dataset(self, dataset):
        """現在のデータセットに基づいてUIを更新する"""
        try:
            # データソースタイプの設定
            data_source_type = dataset.get('data_source_type', 'measured')
            
            # データソースタイプ表示ラベルを更新
            self.dataSourceTypeDisplayLabel.setText("実測データ" if data_source_type == 'measured' else "数式データ")
            
            # 色の設定（共通項目）
            color = dataset.get('color', QColor('blue'))
            if not isinstance(color, QColor):
                color = QColor(color)
            self.currentColor = color
            self.colorButton.setStyleSheet(f'background-color: {self.currentColor.name()};')
            
            # 線の太さ
            self.lineWidthSpin.setValue(dataset.get('line_width', 1.0))
            
            # マーカースタイル
            marker_style = str(dataset.get('marker_style', '*'))
            marker_index = self.markerCombo.findText(marker_style)
            if marker_index >= 0:
                self.markerCombo.setCurrentIndex(marker_index)
            
            # マーカーサイズ
            self.markerSizeSpin.setValue(dataset.get('marker_size', 3.0))
            
            # 凡例表示
            self.legendCheck.setChecked(dataset.get('show_legend', True))
            
            # 凡例ラベル
            legend_text = str(dataset.get('legend_label', dataset.get('name', '')))
            self.legendLabel.setText(legend_text)
            
            # プロットタイプ
            plot_type = str(dataset.get('plot_type', 'line'))
            if plot_type == "line": 
                self.lineRadio.setChecked(True)
            elif plot_type == "scatter": 
                self.scatterRadio.setChecked(True)
            elif plot_type == "line_scatter": 
                self.lineScatterRadio.setChecked(True)
            else: 
                self.barRadio.setChecked(True)
                
            # ラジオボタンを更新（イベントハンドラが発生）
            # 注：ここでイベントが発生するため、下記のコードはイベントハンドラー内で上書きされる可能性がある
            if data_source_type == 'measured':
                if not self.measuredRadio.isChecked():
                    self.measuredRadio.setChecked(True)
            else:
                if not self.formulaRadio.isChecked():
                    self.formulaRadio.setChecked(True)
            
            # データソースタイプに応じた設定の更新
            if data_source_type == 'measured':
                # 実測値の場合
                # CSVファイルパス
                file_path = dataset.get('file_path', '')
                file_type = dataset.get('file_type', 'csv')
                
                if file_type == 'csv':
                    self.csvRadio.setChecked(True)
                    self.fileEntry.setText(file_path)
                    self.toggle_source_fields()  # UIの有効/無効を更新
                elif file_type == 'excel':
                    self.excelRadio.setChecked(True)
                    self.excelEntry.setText(file_path)
                    self.toggle_source_fields()  # UIの有効/無効を更新
                    
                    # シート名を設定
                    sheet_name = dataset.get('sheet_name', '')
                    if sheet_name:
                        index = self.sheetCombobox.findText(sheet_name)
                        if index >= 0:
                            self.sheetCombobox.setCurrentIndex(index)
                elif file_type == 'manual':
                    self.manualRadio.setChecked(True)
                    self.toggle_source_fields()
                
                # 列名が設定されている場合は選択
                x_column = dataset.get('x_column', '')
                y_column = dataset.get('y_column', '')
                if x_column:
                    index = self.xColCombo.findText(x_column)
                    if index >= 0:
                        self.xColCombo.setCurrentIndex(index)
                if y_column:
                    index = self.yColCombo.findText(y_column)
                    if index >= 0:
                        self.yColCombo.setCurrentIndex(index)
                
                # データテーブルを必ずクリアして更新
                self.dataTable.setRowCount(0)
                # 既存のデータがある場合のみテーブルに表示
                if dataset.get('data_x') and len(dataset.get('data_x')) > 0:
                    self.update_data_table_from_dataset(dataset)
            else:
                # 数式モードの場合
                # 数式設定を更新
                self.equationEntry.setText(dataset.get('equation', 'x^2'))
                self.domainMinSpin.setValue(dataset.get('domain_min', 0))
                self.domainMaxSpin.setValue(dataset.get('domain_max', 10))
                self.samplesSpin.setValue(dataset.get('samples', 200))
            
                # 微分・積分設定を削除し、接線設定のみ残す
                self.showTangentCheck.setChecked(dataset.get('show_tangent', False))
                self.tangentXSpin.setValue(dataset.get('tangent_x', 5))
                self.tangentLengthSpin.setValue(dataset.get('tangent_length', 2))
                
                # 色設定の更新
                tangent_color = dataset.get('tangent_color', QColor('purple'))
                self.tangentColor = QColor(tangent_color)
                
                # 色名の更新（保存されていれば）
                tangent_color_name = dataset.get('tangent_color_name', '紫')
                index = self.tangentColorCombo.findText(tangent_color_name)
                if index >= 0:
                    self.tangentColorCombo.setCurrentIndex(index)
                else:
                    # 保存された色名がリストにない場合（カスタム色の可能性）
                    if 'tangent_color' in dataset:
                        # カスタム項目を追加/選択
                        custom_index = self.tangentColorCombo.findText('カスタム')
                        if custom_index == -1:
                            self.tangentColorCombo.addItem('カスタム')
                            custom_index = self.tangentColorCombo.findText('カスタム')
                        self.tangentColorCombo.setCurrentIndex(custom_index)
                
                # ボタンの背景色を更新
                self.tangentColorButton.setStyleSheet(f'background-color: {self.tangentColor.name()};')
                
                tangent_style = dataset.get('tangent_style', '実線')
                index = self.tangentStyleCombo.findText(tangent_style)
                if index >= 0:
                    self.tangentStyleCombo.setCurrentIndex(index)
                
                # 接線の式表示設定も更新
                self.showTangentEquationCheck.setChecked(dataset.get('show_tangent_equation', False))
                
            # データソースに応じたUIの表示/非表示を更新
            self.update_ui_based_on_data_source_type()
                
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"UI更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_based_on_data_source_type(self):
        """データソースタイプに基づいてUIの表示/非表示を更新する"""
        if not hasattr(self, 'measuredRadio') or not hasattr(self, 'formulaRadio'):
            return  # UIがまだ初期化されていない場合
            
        # 測定値モードの場合
        if self.measuredRadio.isChecked():
            # 測定値関連のUIを表示
            self.csvRadio.setEnabled(True)
            self.excelRadio.setEnabled(True)
            self.fileEntry.setEnabled(self.csvRadio.isChecked())
            self.excelEntry.setEnabled(self.excelRadio.isChecked())
            self.sheetCombobox.setEnabled(self.excelRadio.isChecked())
            self.xColCombo.setEnabled(True)
            self.yColCombo.setEnabled(True)
            self.manualRadio.setEnabled(True)
            self.dataTable.setEnabled(self.manualRadio.isChecked())
            
            # 数式関連のUIを非表示または無効化
            self.equationEntry.setEnabled(False)
            self.domainMinSpin.setEnabled(False)
            self.domainMaxSpin.setEnabled(False)
            self.samplesSpin.setEnabled(False)
            
        # 数式モードの場合
        else:
            # 測定値関連のUIを非表示または無効化
            self.csvRadio.setEnabled(False)
            self.excelRadio.setEnabled(False)
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.xColCombo.setEnabled(False)
            self.yColCombo.setEnabled(False)
            self.manualRadio.setEnabled(False)
            self.dataTable.setEnabled(False)
            
            # 数式関連のUIを表示
            self.equationEntry.setEnabled(True)
            self.domainMinSpin.setEnabled(True)
            self.domainMaxSpin.setEnabled(True)
            self.samplesSpin.setEnabled(True)
    
    def update_data_table_from_dataset(self, dataset):
        """データテーブルにデータセットの内容を反映する"""
        try:
            # テーブルをクリア
            self.dataTable.setRowCount(0)
            
            # データがない場合は何もしない
            if not dataset.get('data_x') or not dataset.get('data_y'):
                return
                
            data_x = dataset.get('data_x', [])
            data_y = dataset.get('data_y', [])
            
            # 空のリストの場合も処理せずに終了
            if not data_x or not data_y:
                return
            
            # データを表示
            for i, (x, y) in enumerate(zip(data_x, data_y)):
                self.dataTable.insertRow(i)
                self.dataTable.setItem(i, 0, QTableWidgetItem(str(x)))
                self.dataTable.setItem(i, 1, QTableWidgetItem(str(y)))
                
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データテーブル更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def generate_tikz_code_multi_datasets(self):
        """複数のデータセットを持つTikZコードを生成する"""
        # 結果のLaTeXコード
        latex = []
        
        # figure環境の開始
        latex.append(f"\\begin{{figure}}[{self.global_settings['position']}]")
        latex.append("    \\centering")
        
        # tikzpictureの開始
        latex.append("    \\begin{tikzpicture}")
        
        # axis環境の設定
        axis_options = []
        axis_options.append(f"width={self.global_settings['width']}\\textwidth")
        axis_options.append(f"height={self.global_settings['height']}\\textwidth")
        axis_options.append(f"xlabel={{{self.global_settings['x_label']}}}")
        axis_options.append(f"ylabel={{{self.global_settings['y_label']}}}")
        
        x_min = self.global_settings['x_min']
        x_max = self.global_settings['x_max']
        y_min = self.global_settings['y_min']
        y_max = self.global_settings['y_max']
        
        # 目盛り間隔
        x_tick_step = getattr(self, 'x_tick_step', 1.0)
        y_tick_step = getattr(self, 'y_tick_step', 1.0)
        
        if x_min != x_max:
            axis_options.append(f"xmin={x_min}, xmax={x_max}")
        if y_min != y_max:
            axis_options.append(f"ymin={y_min}, ymax={y_max}")
        
        # xtick, ytickを自動生成
        xticks = ','.join(str(round(x_min + i * x_tick_step, 8)) for i in range(int((x_max - x_min) / x_tick_step) + 1))
        yticks = ','.join(str(round(y_min + i * y_tick_step, 8)) for i in range(int((y_max - y_min) / y_tick_step) + 1))
        axis_options.append(f"xtick={{{xticks}}}")
        axis_options.append(f"ytick={{{yticks}}}")
        
        # 目盛りを見やすくする設定を追加
        axis_options.append("tick align=outside")
        axis_options.append("minor tick num=1")
        axis_options.append("tick label style={font=\\small}")
        axis_options.append("every node near coord/.style={font=\\footnotesize}")
        axis_options.append("clip=false")  # これを追加
        
        if self.global_settings['grid']:
            axis_options.append("grid=both")
        
        if self.legendCheck.isChecked():
            axis_options.append(f"legend pos={self.global_settings['legend_pos']}")
        
        # axis環境の開始
        latex.append(f"        \\begin{{axis}}[")
        latex.append(f"            {','.join(axis_options)}")
        latex.append("        ]")
        
        # 各データセットのプロット処理
        for i, dataset in enumerate(self.datasets):
            if dataset.get('data_source_type') == 'formula' and not dataset.get('equation'):
                continue  # 数式が空の場合はスキップ
                
            if dataset.get('data_source_type') == 'measured' and (not dataset.get('data_x') or not dataset.get('data_y')):
                continue  # 測定データがない場合はスキップ
            
            # データセットの設定を取得
            plot_type = dataset.get('plot_type', 'line')
            color = dataset.get('color', QColor('blue')).name()
            line_width = dataset.get('line_width', 1.0)
            marker_style = dataset.get('marker_style', '*')
            marker_size = dataset.get('marker_size', 3.0)
            show_legend = dataset.get('show_legend', True)
            legend_label = dataset.get('legend_label', dataset.get('name', ''))
            
            # データセットを処理
            self.add_dataset_to_latex(latex, dataset, i, plot_type, color, line_width, 
                                     marker_style, marker_size, show_legend, legend_label)
            
            # 特殊点の追加
            special_points = dataset.get('special_points', [])
            for point in special_points:
                x, y, point_color, coord_display = point  # 座標表示の種類を取得
                latex.append(f"        % 特殊点 (データセット{i+1}: {dataset.get('name', '')})")
                latex.append(f"        \\addplot[only marks, mark=*, {point_color}] coordinates {{")
                latex.append(f"            ({x}, {y})")
                latex.append("        };")
                
                # 目盛りと重複する値には座標値を表示しないようにする関数を追加
                def is_tick_value(val, tick_min, tick_max, tick_step, tol=1e-6):
                    if tick_step is None or tick_step <= 0:
                        return False
                    n = round((val - tick_min) / tick_step)
                    tick_val = tick_min + n * tick_step
                    return abs(val - tick_val) < tol and tick_min <= val <= tick_max

                # X軸への垂線
                if coord_display.startswith('X座標') or coord_display.startswith('X,Y座標'):
                    latex.append(f"        % X軸への垂線")
                    latex.append(f"        \\draw[{point_color}, dashed] (axis cs:{x},{y}) -- (axis cs:{x},{y_min});")
                    # X座標値の表示（値も表示が選択されている場合）
                    if '値も表示' in coord_display and not is_tick_value(x, x_min, x_max, x_tick_step):
                        formatted_x = '{:g}'.format(x)
                        latex.append(f"        % X座標値を表示")
                        latex.append(f"        \\node[{point_color}, below, yshift=-2pt, font=\\small] at (axis cs:{x},{y_min}) {{{formatted_x}}};")

                # Y軸への垂線
                if coord_display.startswith('Y座標') or coord_display.startswith('X,Y座標'):
                    latex.append(f"        % Y軸への垂線")
                    latex.append(f"        \\draw[{point_color}, dashed] (axis cs:{x},{y}) -- (axis cs:{x_min},{y});")
                    # Y座標値の表示（値も表示が選択されている場合）
                    if '値も表示' in coord_display and not is_tick_value(y, y_min, y_max, y_tick_step):
                        formatted_y = '{:g}'.format(y)
                        latex.append(f"        % Y座標値を表示")
                        latex.append(f"        \\node[{point_color}, left, xshift=-2pt, font=\\small] at (axis cs:{x_min},{y}) {{{formatted_y}}};")
        
        # 注釈の追加
        for i, dataset in enumerate(self.datasets):
            annotations = dataset.get('annotations', [])
            for ann in annotations:
                x, y, text, color, pos = ann
                latex.append(f"        % 注釈 (データセット{i+1}: {dataset.get('name', '')})")
                latex.append(f"        \\node at (axis cs:{x},{y}) [anchor={pos}, font=\\small, text={color}] {{{text}}};")
        
        # axis環境の終了
        latex.append("        \\end{axis}")
        
        # tikzpictureの終了
        latex.append("    \\end{tikzpicture}")
        
        # キャプションとラベル
        latex.append(f"    \\caption{{{self.global_settings['caption']}}}")
        latex.append(f"    \\label{{{self.global_settings['label']}}}")
        
        # figure環境の終了
        latex.append("\\end{figure}")
        
        return '\n'.join(latex)
    
    def add_dataset_to_latex(self, latex, dataset, index, plot_type, color, line_width, 
                             marker_style, marker_size, show_legend, legend_label):
        """LaTeXコードにデータセットを追加する"""
        data_source_type = dataset.get('data_source_type', 'measured')
        
        if data_source_type == 'formula':
            # 数式の場合は理論曲線として描画
            equation = dataset.get('equation', 'x^2')
            # TikZ互換の数式に変換
            tikz_equation = self.format_equation_for_tikz(equation)
            domain_min = dataset.get('domain_min', 0)
            domain_max = dataset.get('domain_max', 10)
            samples = dataset.get('samples', 200)
            
            # 数式がない場合はスキップ
            if not equation.strip():
                return
            
            # QColorオブジェクトをTikZ互換形式に変換
            tikz_color = self.color_to_tikz_rgb(dataset.get('color', QColor('blue')))
            
            # 理論曲線のオプション
            theory_options = []
            theory_options.append(f"domain={domain_min}:{domain_max}")
            theory_options.append(f"samples={samples}")
            theory_options.append("smooth")
            theory_options.append("thick")
            theory_options.append(tikz_color)  # TikZ互換のRGB値
            theory_options.append(f"line width={dataset.get('line_width', 1.0)}pt")
            
            latex.append(f"        % データセット{index+1}: {dataset.get('name', '')} （数式: {equation}）")
            latex.append(f"        \\addplot[{', '.join(theory_options)}] {{")
            latex.append(f"            {tikz_equation}")
            latex.append("        };")
            
            # 凡例エントリを追加
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
            
            # 微分曲線の追加（オプションが有効な場合）
            if dataset.get('show_derivative', False):
                # 微分曲線の色をTikZ互換形式に変換
                deriv_color = self.color_to_tikz_rgb(dataset.get('derivative_color', QColor('red')))
                
                # 線のスタイル設定
                line_style = ""
                deriv_style = dataset.get('derivative_style', '実線')
                if deriv_style == '点線':
                    line_style = "dotted"
                elif deriv_style == '破線':
                    line_style = "dashed"
                elif deriv_style == '一点鎖線':
                    line_style = "dashdotted"
                
                # 微分曲線のオプション
                deriv_options = []
                deriv_options.append(f"domain={domain_min}:{domain_max}")
                deriv_options.append(f"samples={samples}")
                deriv_options.append("smooth")
                if line_style:
                    deriv_options.append(line_style)
                deriv_options.append(deriv_color)  # TikZ互換のRGB値
                deriv_options.append(f"line width={dataset.get('line_width', 1.0)}pt")
                
                latex.append(f"        % 微分曲線 （データセット{index+1}: {dataset.get('name', '')}）")
                latex.append(f"        \\addplot[{', '.join(deriv_options)}] {{")
                # 微分を計算するための数式変換（簡易的な実装）
                latex.append(f"            derivative({tikz_equation}, x)")
                latex.append("        };")
                
                # 凡例エントリを追加
                if show_legend:
                    latex.append(f"        \\addlegendentry{{{legend_label}の微分}}")
            
            # 積分曲線の追加（オプションが有効な場合）
            if dataset.get('show_integral', False):
                # 積分曲線の色をTikZ互換形式に変換
                integral_color = self.color_to_tikz_rgb(dataset.get('integral_color', QColor('green')))
                
                # 線のスタイル設定
                line_style = ""
                integral_style = dataset.get('integral_style', '点線')
                if integral_style == '点線':
                    line_style = "dotted"
                elif integral_style == '破線':
                    line_style = "dashed"
                elif integral_style == '一点鎖線':
                    line_style = "dashdotted"
                
                # 積分定数
                integral_const = dataset.get('integral_const', 0)
                
                # 積分曲線のオプション
                integral_options = []
                integral_options.append(f"domain={domain_min}:{domain_max}")
                integral_options.append(f"samples={samples}")
                integral_options.append("smooth")
                if line_style:
                    integral_options.append(line_style)
                integral_options.append(integral_color)  # TikZ互換のRGB値
                integral_options.append(f"line width={dataset.get('line_width', 1.0)}pt")
                
                latex.append(f"        % 積分曲線 （データセット{index+1}: {dataset.get('name', '')}）")
                latex.append(f"        \\addplot[{', '.join(integral_options)}] {{")
                # 積分を計算するための数式変換（簡易的な実装）
                latex.append(f"            integral({tikz_equation}, x) + {integral_const}")
                latex.append("        };")
                
                # 凡例エントリを追加
                if show_legend:
                    latex.append(f"        \\addlegendentry{{{legend_label}の積分}}")
            
            # 接線表示部分を完全に書き直し
            # TikZ コードを完全に書き換え - 接線の明示的な計算
            if dataset.get('show_tangent', False):
                # 接線のx座標と長さ
                tangent_x = dataset.get('tangent_x', 5)
                tangent_length = dataset.get('tangent_length', 2)
                
                # 接線の色をTikZ互換形式に変換
                tangent_color = self.color_to_tikz_rgb(dataset.get('tangent_color', QColor('purple')))
                
                # 線のスタイル設定
                line_style = ""
                tangent_style = dataset.get('tangent_style', '実線')
                if tangent_style == '点線':
                    line_style = "dotted"
                elif tangent_style == '破線':
                    line_style = "dashed"
                elif tangent_style == '一点鎖線':
                    line_style = "dashdotted"
                
                # 接線のオプション
                tangent_options = []
                if line_style:
                    tangent_options.append(line_style)
                tangent_options.append("thick")
                tangent_options.append(tangent_color)  # TikZ互換のRGB値
                tangent_options.append(f"line width={dataset.get('line_width', 1.5)}pt")

                # コメント行
                latex.append(f"        % 接線 （データセット{index+1}: {dataset.get('name', '')}）")
                
                # まず点をプロット - この部分もPython側で計算して値を直接指定
                try:
                    # 数式の文字列置換 'x' を tangent_x に置き換えて評価
                    x_val = tangent_x
                    formula = dataset.get('equation', 'x^2')  # データセットから式を取得
                    
                    # TikZ式をPython式に変換（^ を ** に置換）
                    python_formula = formula.replace('^', '**')
                    # 数式形式を整える（乗算記号の追加など）- 改良したヘルパー関数を使用
                    python_formula = self.format_equation_for_tikz(python_formula)
                    
                    # グラフの表示範囲を取得
                    y_min = self.global_settings['y_min']
                    y_max = self.global_settings['y_max']
                    
                    # 式を評価 - 数学関数を使用できるようにmath名前空間も提供
                    y_val = eval(python_formula.replace('x', str(x_val)), {"__builtins__": {}}, {"math": math})
                    
                    # y値が異常に大きくないかチェック - 表示範囲外なら警告
                    if y_val < y_min or y_val > y_max:
                        latex.append(f"        % 警告: 計算されたy値 ({y_val}) がグラフ範囲外です。範囲: [{y_min}, {y_max}]")
                        self.statusBar.showMessage(f"警告: 点 (x={x_val}) の計算値 (y={y_val}) がグラフ範囲外です", 5000)
                        # 範囲外の場合でも計算は続行するが、後で表示位置を調整する
                    
                    # 点の表示 - 実際の計算値を使用
                    point_code = f"""        % 接線の点をマーク
        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({x_val}, {y_val})}};\n"""
                    latex.append(point_code)
                    
                    # 数値微分で接線の傾きを計算 - 精度改善
                    dx = 0.0001
                    # 安全な評価のためmathを提供
                    globals_dict = {"__builtins__": {}}
                    locals_dict = {"math": math}
                    
                    # 左右の点を計算
                    left_x = x_val - dx
                    right_x = x_val + dx
                    
                    # 左右の点のy値を計算（数式にxを代入）
                    try:
                        y1 = eval(python_formula.replace('x', str(left_x)), globals_dict, locals_dict)
                        y2 = eval(python_formula.replace('x', str(right_x)), globals_dict, locals_dict)
                        # 傾きを計算（中心差分法）
                        slope = (y2 - y1) / (2 * dx)
                    except Exception as inner_e:
                        # 微分計算中のエラーを詳細に記録（デバッグ用）
                        latex.append(f"        % 接線の微分計算でエラー: {str(inner_e)}. 中央差分法を使用できないため、前方差分法を試みます。")
                        # 前方差分法を試みる
                        try:
                            forward_x = x_val + dx
                            y_current = y_val  # すでに計算済み
                            y_forward = eval(python_formula.replace('x', str(forward_x)), globals_dict, locals_dict)
                            slope = (y_forward - y_current) / dx
                        except Exception as e2:
                            # それでも失敗した場合はデフォルト値を使用
                            latex.append(f"        % 前方差分も失敗。デフォルトの傾き 1 を使用します。エラー: {str(e2)}")
                            slope = 1.0
                            
                    # 接線の方程式: y = y0 + slope*(x - x0)
                    # もっとシンプルに y = mx + b の形に変換
                    slope_rounded = round(slope, 3)
                    intercept = round(y_val - slope * x_val, 3)
                    equation_text = f"y = {slope_rounded}x + {intercept}" if intercept >= 0 else f"y = {slope_rounded}x - {abs(intercept)}"
                    
                    # デバッグ情報を追加
                    latex.append(f"        % デバッグ情報: x={x_val}, y={y_val}, 傾き={slope_rounded}, 切片={intercept}")
                    
                    # 接線方程式のチェック - 計算式で元の点を通るか検証
                    test_y = slope_rounded * x_val + intercept
                    if abs(test_y - y_val) > 0.1:  # 許容誤差
                        latex.append(f"        % 警告: 接線方程式が元の点を通りません。式による計算値: {test_y}, 元の点: {y_val}")
                    
                    # 関数形式での接線の方程式
                    tangent_equation = f"{y_val} + {slope}*(x - {x_val})"
                    
                    # 接線を関数形式でプロット
                    tangent_code = f"""        % 接線を関数形式でプロット
        \\addplot[{', '.join(tangent_options)}, domain={x_val-tangent_length/2}:{x_val+tangent_length/2}] {{
            {tangent_equation}
        }};\n"""
                    latex.append(tangent_code)
                    
                    # 接線の式を表示（オプションがオンの場合）
                    if dataset.get('show_tangent_equation', False):
                        # 式の表示位置を調整（接線の少し上に配置）- 範囲外対応
                        equation_pos_x = x_val
                        # y値が範囲外の場合、表示位置を調整
                        if y_val < y_min:
                            equation_pos_y = y_min + 0.5  # 下限より少し上
                        elif y_val > y_max:
                            equation_pos_y = y_max - 0.5  # 上限より少し下
                        else:
                            equation_pos_y = y_val + 0.5  # 通常は点の少し上
                        
                        # 式の表示コード
                        equation_display = f"""        % 接線の方程式を表示
        \\node[anchor=south, font=\\small, {tangent_color}] at (axis cs:{equation_pos_x}, {equation_pos_y}) {{{equation_text}}};\n"""
                        latex.append(equation_display)
                    
                    # 凡例用のエントリ
                    latex.append(f"        \\addlegendimage{{{', '.join(tangent_options)}}};")
                    
                    # 凡例エントリを追加
                    if show_legend:
                        latex.append(f"        \\addlegendentry{{{legend_label}の接線 (x={tangent_x})}}")
                except Exception as e:
                    # 詳細なエラー情報を提供
                    error_msg = f"接線の計算でエラーが発生しました。式: {formula}, エラー: {str(e)}"
                    print(error_msg)  # コンソールにも出力
                    self.statusBar.showMessage(f"警告: {error_msg}", 5000)
                    
                    # エラーが発生した場合は警告コメントを追加
                    latex.append(f"        % {error_msg}")
                    
                    # 点の表示を試みる - 式の評価ができなければ原点を使用
                    try:
                        # TikZ式をPython式に変換し、安全に評価
                        python_formula = formula.replace('^', '**')
                        python_formula = self.format_equation_for_tikz(python_formula)
                        point_y = eval(python_formula.replace('x', str(x_val)), {"__builtins__": {}}, {"math": math})
                        latex.append(f"        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({tangent_x}, {point_y})}}; % 接線計算エラーだが点は表示")
                    except Exception as point_error:
                        # それでも失敗したら原点にプロット
                        latex.append(f"        % 点の計算も失敗: {str(point_error)}")
                        latex.append(f"        \\addplot[only marks, mark=*, {tangent_color}, mark size=3] coordinates {{({tangent_x}, 0)}}; % エラーのため原点にプロット")
                    
                    # 接線の代わりに水平線を表示（目印として）
                    latex.append(f"        \\addplot[{', '.join(tangent_options)}, dashed] coordinates {{({tangent_x-0.5}, 0) ({tangent_x+0.5}, 0)}}; % エラーのため水平線を表示")
                    
                    # 凡例用のエントリ（エラーが発生していることを示す）
                    latex.append(f"        \\addlegendimage{{{', '.join(tangent_options)}, dashed}};")
                    if show_legend:
                        latex.append(f"        \\addlegendentry{{{legend_label}の接線 (計算エラー)}}")
            
            return
            
        # 実測値の場合（以下は既存のコード）
        # データポイントのフォーマット
        coordinates = []
        for x, y in zip(dataset.get('data_x', []), dataset.get('data_y', [])):
            coordinates.append(f"({x}, {y})")
        
        if not coordinates:
            return
        
        # QColorオブジェクトをTikZ互換形式に変換
        tikz_color = self.color_to_tikz_rgb(QColor(color))
        
        # 線グラフまたは線と点の組み合わせの場合の線プロット
        if plot_type == "line" or plot_type == "line_scatter":
            # 線プロット
            plot_options = []
            plot_options.append(tikz_color)  # TikZ互換のRGB値
            plot_options.append(f"line width={line_width}pt")
            plot_options.append("thick")
            
            latex.append(f"        % データセット{index+1}: {dataset.get('name', '')} （線）")
            latex.append(f"        \\addplot[{', '.join(plot_options)}] coordinates {{")
            
            # 座標リストを1行に20個ずつ分割
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            # 凡例エントリを追加
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
        
        # 散布図または線と点の組み合わせの場合の点プロット
        if plot_type == "scatter" or plot_type == "line_scatter" or self.showDataPointsCheck.isChecked():
            # 点プロット
            scatter_options = []
            scatter_options.append("only marks")
            scatter_options.append(f"mark={marker_style}")
            scatter_options.append(tikz_color)  # TikZ互換のRGB値
            scatter_options.append(f"mark size={marker_size}")
            
            latex.append(f"        % データセット{index+1}: {dataset.get('name', '')} （点）")
            latex.append(f"        \\addplot[{', '.join(scatter_options)}] coordinates {{")
            
            # 座標リストを1行に20個ずつ分割
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            # 凡例エントリを追加（線と点の組み合わせの場合は表示しない）
            if show_legend and plot_type == "scatter":
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
                
        # 棒グラフの場合
        if plot_type == "bar":
            # 棒グラフ
            bar_options = []
            bar_options.append(f"ybar")
            bar_options.append(tikz_color)  # TikZ互換のRGB値
            bar_options.append(f"fill={tikz_color.replace('color = ','')}")  # fill=color = ... ではなく fill=...
            
            latex.append(f"        % データセット{index+1}: {dataset.get('name', '')} （棒グラフ）")
            latex.append(f"        \\addplot[{', '.join(bar_options)}] coordinates {{")
            
            # 座標リストを1行に20個ずつ分割
            for i in range(0, len(coordinates), 20):
                chunk = coordinates[i:i+20]
                latex.append(f"            {' '.join(chunk)}")
            
            latex.append("        };")
            
            # 凡例エントリを追加
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
    
    def add_theory_curve_to_latex(self, latex, dataset, index):
        """LaTeXコードに理論曲線を追加する"""
        try:
            # 理論曲線の設定
            equation = dataset.get('equation', 'x^2')
            domain_min = dataset.get('domain_min', 0)
            domain_max = dataset.get('domain_max', 10)
            samples = dataset.get('samples', 200)
            color = dataset.get('color', QColor('green')).name()
            line_width = dataset.get('line_width', 1.0)
            show_legend = dataset.get('show_legend', True)
            legend_label = dataset.get('legend_label', "理論曲線")
            
            # 理論曲線のオプション
            theory_options = []
            theory_options.append(f"domain={domain_min}:{domain_max}")
            theory_options.append(f"samples={samples}")
            theory_options.append("smooth")
            theory_options.append("thick")
            theory_options.append(color)
            theory_options.append(f"line width={line_width}pt")
            
            latex.append(f"        % 理論曲線{index+1}: {dataset.get('name', '')}")
            latex.append(f"        \\addplot[{', '.join(theory_options)}] {{")
            latex.append(f"            {equation}")
            latex.append("        };")
            
            # 凡例エントリを追加
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label}}}")
        except Exception as e:
            # エラー発生時はコメントとして記録
            latex.append(f"        % Error in theory curve {index+1}: {str(e)}")
            # 安全な代替プロット
            latex.append(f"        \\addplot[{color}, dashed] coordinates {{")
            latex.append(f"            ({domain_min}, 0) ({domain_max}, 0)")
            latex.append("        };")
            if show_legend:
                latex.append(f"        \\addlegendentry{{{legend_label} (ERROR)}}")

    # 特殊点をデータセットに割り当て
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
    
    # 注釈をデータセットに割り当て
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
                    
                    # 日本語位置表記をTikZ形式に変換
                    tikz_pos = self.convert_position_to_tikz(pos_val)
                    
                    annotations.append((x_val, y_val, text_val, color_val, tikz_pos))
                except ValueError:
                    pass
        
        dataset = self.datasets[self.current_dataset_index]
        dataset['annotations'] = annotations
        
        QMessageBox.information(self, "成功", 
                              f"データセット '{dataset['name']}' に {len(annotations)} 個の注釈を割り当てました")
    
    # 日本語位置表記をTikZ形式に変換するヘルパーメソッドを追加
    def convert_position_to_tikz(self, jp_position):
        """日本語の位置表記をTikZの位置表記に変換する
        注意: TikZのアンカー指定は直感と逆になる
        例: 「右上」に表示したい場合は「左下(south west)」をアンカーに指定する
        """
        # 直感的な位置とTikZアンカーの対応関係（逆転させる）
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

    def on_data_source_type_changed(self, checked):
        """データソースタイプが変更されたときに呼ばれる"""
        if not checked:  # イベントはtoggleで発生するため、チェックされたラジオボタンのみ処理
            return
            
        # 現在のデータセットがあれば状態を保存
        if self.current_dataset_index >= 0:
            self.update_current_dataset()
            
        # UIの表示/非表示を切り替え
        is_measured = self.measuredRadio.isChecked()
        
        self.measuredContainer.setVisible(is_measured)
        self.formulaContainer.setVisible(not is_measured)
        
        # 現在のデータセットの状態を更新
        if self.current_dataset_index >= 0:
            dataset = self.datasets[self.current_dataset_index]
            dataset['data_source_type'] = 'measured' if is_measured else 'formula'
            
            # データソースタイプ表示ラベルを更新
            self.dataSourceTypeDisplayLabel.setText("実測データ" if is_measured else "数式データ")
            
        # UIの要素の有効/無効を更新
        self.update_ui_based_on_data_source_type()

    def apply_formula(self):
        """数式を適用してデータを生成する"""
        # 入力値を保存
        self.update_current_dataset()
        if self.current_dataset_index < 0:
            QMessageBox.warning(self, "警告", "数式を適用するデータセットを選択してください")
            return
            
        equation = self.equationEntry.text().strip()
        if not equation:
            QMessageBox.warning(self, "警告", "有効な数式を入力してください")
            return
            
        try:
            # 数式、ドメイン、サンプル数を取得
            domain_min = self.domainMinSpin.value()
            domain_max = self.domainMaxSpin.value()
            samples = self.samplesSpin.value()
            
            # データセットを更新
            dataset = self.datasets[self.current_dataset_index]
            dataset['data_source_type'] = 'formula'
            dataset['equation'] = equation
            dataset['domain_min'] = domain_min
            dataset['domain_max'] = domain_max
            dataset['samples'] = samples
            
            # 接線設定も更新
            dataset['show_tangent'] = self.showTangentCheck.isChecked()
            dataset['tangent_x'] = self.tangentXSpin.value()
            dataset['tangent_length'] = self.tangentLengthSpin.value()
            dataset['tangent_color'] = QColor(self.tangentColor)
            dataset['tangent_style'] = self.tangentStyleCombo.currentText()
            dataset['show_tangent_equation'] = self.showTangentEquationCheck.isChecked()
            
            # X値（ドメイン）の生成
            x_values = np.linspace(domain_min, domain_max, samples).tolist()
            dataset['data_x'] = x_values
            
            # 実際の値はTikZコード生成時に計算されるため、空のリストを設定
            dataset['data_y'] = []
            
            # 数式をPythonで評価して試験的にY値を計算（範囲外チェック用）
            python_formula = equation.replace('^', '**')
            python_formula = self.format_equation_for_tikz(python_formula)
            
            # グラフ表示範囲を取得
            y_min = self.global_settings['y_min']
            y_max = self.global_settings['y_max']
            
            # テスト点でチェック（最初、中央、最後）
            test_points = [domain_min, (domain_min + domain_max) / 2, domain_max]
            out_of_range_points = []
            
            for test_x in test_points:
                try:
                    test_y = eval(python_formula.replace('x', str(test_x)), {"__builtins__": {}}, {"math": math})
                    if test_y < y_min or test_y > y_max:
                        out_of_range_points.append((test_x, test_y))
                except Exception as e:
                    # 評価エラーは無視（実際のプロット時に処理される）
                    pass
            
            # 範囲外の点があれば警告
            if out_of_range_points:
                point_info = ", ".join([f"(x={x:.2f}, y={y:.2f})" for x, y in out_of_range_points])
                QMessageBox.warning(self, "表示範囲外の値", 
                    f"計算された一部の点がグラフ表示範囲外です: {point_info}\n\n"
                    f"グラフ表示範囲: Y軸 [{y_min}, {y_max}]\n\n"
                    "グラフの一部が表示されない可能性があります。Y軸の範囲を調整することを検討してください。")
            
            # データセットの変更後にUI更新
            self.update_ui_from_dataset(dataset)
            self.statusBar.showMessage("数式データを適用しました", 3000)
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"数式の適用中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("数式適用エラー", 3000)

    # CSVファイルの列名を更新
    def update_column_names(self, file_path):
        try:
            df = pd.read_csv(file_path)
            self.xColCombo.clear()
            self.yColCombo.clear()
            self.xColCombo.addItems(df.columns)
            self.yColCombo.addItems(df.columns)
            self.statusBar.showMessage(f"CSV列名を読み込みました: {len(df.columns)}列")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVファイルの読み込みに失敗しました: {str(e)}")
            self.statusBar.showMessage("ファイル読み込みエラー")
    
    # データテーブルに行を追加
    def add_table_row(self):
        row_position = self.dataTable.rowCount()
        self.dataTable.insertRow(row_position)
    
    # データテーブルから選択行を削除
    def remove_table_row(self):
        selected_rows = set(index.row() for index in self.dataTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.dataTable.removeRow(row)
    
    # データテーブルに列を追加
    def add_table_column(self):
        col_position = self.dataTable.columnCount()
        self.dataTable.insertColumn(col_position)
        self.dataTable.setHorizontalHeaderItem(col_position, QTableWidgetItem(f"列{col_position+1}"))
    
    # 色選択ダイアログ
    def select_color(self):
        color = QColorDialog.getColor(self.currentColor, self, "線の色を選択")
        if color.isValid():
            self.currentColor = color
            self.colorButton.setStyleSheet(f'background-color: {color.name()};')
    
    # 理論曲線の色選択ダイアログは削除（パラメータスイープ機能の廃止に伴い）
    
    # データを読み込む
    def load_data(self):
        """選択されたデータソースからデータを読み込む"""
        try:
            if self.current_dataset_index < 0:
                QMessageBox.warning(self, "警告", "データを読み込むデータセットを選択してください")
                return
                
            data_x = []
            data_y = []
            
            if self.csvRadio.isChecked():
                file_path = self.fileEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なCSVファイルを選択してください")
                    return
                
                df = pd.read_csv(file_path)
                x_col = self.xColCombo.currentText()
                y_col = self.yColCombo.currentText()
                
                if x_col in df.columns and y_col in df.columns:
                    data_x = df[x_col].values.tolist()
                    data_y = df[y_col].values.tolist()
                else:
                    QMessageBox.warning(self, "警告", "選択した列名が不正です")
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
                
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                x_col = self.xColCombo.currentText()
                y_col = self.yColCombo.currentText()
                
                if x_col in df.columns and y_col in df.columns:
                    data_x = df[x_col].values.tolist()
                    data_y = df[y_col].values.tolist()
                else:
                    QMessageBox.warning(self, "警告", "選択した列名が不正です")
                    return
                
            elif self.manualRadio.isChecked():
                data_x = []
                data_y = []
                
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
                            pass  # 数値に変換できない場合はスキップ
                
                if not data_x:
                    QMessageBox.warning(self, "警告", "有効なデータポイントがありません")
                    return
                
            # 現在のデータセットにデータを設定
            self.datasets[self.current_dataset_index]['data_x'] = data_x
            self.datasets[self.current_dataset_index]['data_y'] = data_y
            
            # データテーブルを更新（手動入力モードの場合）
            if self.manualRadio.isChecked():
                self.update_data_table_from_dataset(self.datasets[self.current_dataset_index])
            
            dataset_name = self.datasets[self.current_dataset_index]['name']
            QMessageBox.information(self, "成功", f"データセット '{dataset_name}' に{len(data_x)}個のデータポイントを読み込みました")
            self.statusBar.showMessage(f"データセット '{dataset_name}' にデータを読み込みました: {len(data_x)}ポイント")
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データ読み込み中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("データ読み込みエラー")
    
    # 特殊点を追加
    def add_special_point(self):
        row_position = self.specialPointsTable.rowCount()
        self.specialPointsTable.insertRow(row_position)
        
        # デフォルト値の設定
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        
        # 色選択コンボボックス
        color_combo = QComboBox()
        color_combo.addItems(['red', 'blue', 'green', 'black', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("red")
        
        # 座標表示コンボボックス（チェックボックスから変更）
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
    
    # 特殊点を削除
    def remove_special_point(self):
        selected_rows = set(index.row() for index in self.specialPointsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.specialPointsTable.removeRow(row)
    
    # 注釈を追加
    def add_annotation(self):
        row_position = self.annotationsTable.rowCount()
        self.annotationsTable.insertRow(row_position)
        
        # デフォルト値の設定
        x_item = QTableWidgetItem("0.0")
        y_item = QTableWidgetItem("0.0")
        text_item = QTableWidgetItem("注釈テキスト")
        
        # 色選択コンボボックス
        color_combo = QComboBox()
        color_combo.addItems(['black', 'red', 'blue', 'green', 'purple', 'orange', 'brown', 'gray'])
        color_combo.setCurrentText("black")
        
        # 位置選択コンボボックス - より直感的な表現に変更
        pos_combo = QComboBox()
        pos_combo.addItems(['上', '右上', '右', '右下', '下', '左下', '左', '左上'])
        pos_combo.setCurrentText("右上")
        
        self.annotationsTable.setItem(row_position, 0, x_item)
        self.annotationsTable.setItem(row_position, 1, y_item)
        self.annotationsTable.setItem(row_position, 2, text_item)
        self.annotationsTable.setCellWidget(row_position, 3, color_combo)
        self.annotationsTable.setCellWidget(row_position, 4, pos_combo)
    
    # 注釈を削除
    def remove_annotation(self):
        selected_rows = set(index.row() for index in self.annotationsTable.selectedIndexes())
        for row in sorted(selected_rows, reverse=True):
            self.annotationsTable.removeRow(row)
    
    # パラメータスイープ関連のメソッドを削除（パラメータスイープ機能の廃止に伴い）
    
    # LaTeXコードをクリップボードにコピー
    def copy_to_clipboard(self):
        latex_code = self.resultText.toPlainText()
        if latex_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(latex_code)
            self.statusBar.showMessage("LaTeXコードをクリップボードにコピーしました", 3000)  # 3秒間表示

    # 微分曲線の色選択ダイアログは削除
    
    # 積分曲線の色選択ダイアログは削除
            
    # 接線の色選択ダイアログ
    def select_tangent_color(self):
        """接線の色を選択するダイアログを表示"""
        color = QColorDialog.getColor(self.tangentColor, self, "接線の色を選択")
        if color.isValid():
            self.tangentColor = color
            # ボタンの背景色を更新
            self.tangentColorButton.setStyleSheet(f'background-color: {color.name()};')
            # カスタム色選択時はコンボボックスのカスタムオプションを選択
            index = self.tangentColorCombo.findText('カスタム')
            if index == -1:  # 'カスタム'項目がなければ追加
                self.tangentColorCombo.addItem('カスタム')
                index = self.tangentColorCombo.findText('カスタム')
            self.tangentColorCombo.setCurrentIndex(index)
            self.statusBar.showMessage("接線の色を設定しました", 2000)

    # 数式グラフをプレビュー
    def preview_formula_graph(self):
        try:
            equation = self.equationEntry.text().strip()
            if not equation:
                QMessageBox.warning(self, "警告", "有効な数式を入力してください")
                return

            # Matplotlibを使用してグラフをプレビュー表示
            # 注：この機能を実装するにはMatplotlibをインポートする必要があります
            QMessageBox.information(self, "プレビュー機能", 
                "グラフのプレビュー機能が準備されています。\n\n"
                f"数式：{equation}\n"
                f"範囲：{self.domainMinSpin.value()} 〜 {self.domainMaxSpin.value()}\n"
                f"サンプル数：{self.samplesSpin.value()}\n\n"
                "接線表示：" + ("はい" if self.showTangentCheck.isChecked() else "いいえ") + 
                (f" (x={self.tangentXSpin.value()})" if self.showTangentCheck.isChecked() else "") + "\n\n"
                "完全なプレビュー機能を有効にするには、Matplotlibをインポートして実装してください。")
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"数式プレビュー中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("プレビューエラー")

    # まず、クラスのトップレベルに数式を変換するヘルパーメソッドを追加
    def format_equation_for_tikz(self, equation):
        # 自動補間をやめて、入力そのまま返す
        return equation

    def insert_formula_preset(self, formula):
        """数式プリセットをクリックしたときに呼ばれる関数"""
        try:
            # 現在のカーソル位置に数式を挿入するか現在の内容を置き換える
            current_text = self.equationEntry.text()
            if current_text.strip() == "":
                # テキストが空の場合は単純に設定
                self.equationEntry.setText(formula)
            else:
                # 確認ダイアログを表示
                reply = QMessageBox.question(
                    self, 
                    "数式の挿入", 
                    f"現在の数式を\n「{formula}」\nに置き換えますか？\n\n「いいえ」を選択すると現在の数式に追加します。",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                )
                
                if reply == QMessageBox.Yes:
                    # 内容を置き換え
                    self.equationEntry.setText(formula)
                elif reply == QMessageBox.No:
                    # 現在のテキストの末尾に追加
                    self.equationEntry.setText(f"{current_text} + {formula}")
                # Cancelの場合は何もしない
            
            # フォーカスを数式入力欄に戻す
            self.equationEntry.setFocus()
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"数式の挿入中にエラーが発生しました: {str(e)}")

    def insert_function_from_table(self, row, column):
        """関数テーブルの関数をダブルクリックして挿入"""
        try:
            # 選択された関数を取得（常に0列目の項目を使用）
            function_item = self.tikzFunctionsTable.item(row, 0)
            if function_item:
                function_text = function_item.text()
                
                # 単一の定数（pi, e）の場合は直接挿入
                if function_text in ["pi", "e"]:
                    self.insert_into_equation(function_text)
                else:
                    # 括弧を含む関数の場合、カーソル位置を括弧内に設定するために
                    # 括弧部分を抽出し、カーソルを括弧内に配置
                    if "(" in function_text and ")" in function_text:
                        # 括弧内の内容を抽出
                        bracket_content = function_text[function_text.find("(")+1:function_text.find(")")]
                        # 括弧内にカーソルを配置するため、内容を削除
                        insert_text = function_text.replace(bracket_content, "")
                        self.insert_into_equation(insert_text)
                    else:
                        # 括弧のない関数の場合はそのまま挿入
                        self.insert_into_equation(function_text)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"関数の挿入中にエラーが発生しました: {str(e)}")

    def insert_into_equation(self, text):
        """数式入力欄にテキストを挿入"""
        current_text = self.equationEntry.text()
        current_pos = self.equationEntry.cursorPosition()
        
        # 現在のテキストが空の場合は単純に設定
        if current_text.strip() == "":
            self.equationEntry.setText(text)
            # 括弧がある場合はカーソルを括弧の中に配置
            if "(" in text and ")" in text:
                self.equationEntry.setCursorPosition(text.find("(")+1)
        else:
            # テキストの挿入
            new_text = current_text[:current_pos] + text + current_text[current_pos:]
            self.equationEntry.setText(new_text)
            # 括弧がある場合はカーソルを括弧の中に配置
            if "(" in text and ")" in text:
                new_cursor_pos = current_pos + text.find("(")+1
                self.equationEntry.setCursorPosition(new_cursor_pos)
            else:
                # 括弧がない場合は挿入したテキストの後にカーソルを配置
                self.equationEntry.setCursorPosition(current_pos + len(text))
        
        # フォーカスを数式入力欄に戻す
        self.equationEntry.setFocus()

    def show_tikz_function_help(self):
        """TikZ数式の詳細ガイドを表示"""
        help_text = """
<h2>TikZ数式の使い方ガイド</h2>

<h3>基本構文</h3>
<p>TikZでは、数式は以下のように記述します：</p>
<pre>\\addplot[domain=0:10] {数式};</pre>

<h3>数式の例</h3>
<ul>
    <li><b>基本演算:</b> <code>x^2 + 3*x - 5</code></li>
    <li><b>三角関数:</b> <code>sin(pi*x)</code> （引数はラジアン）</li>
    <li><b>指数関数:</b> <code>exp(-x^2/2)</code> （正規分布）</li>
    <li><b>条件分岐:</b> <code>x>0 ? x^2 : -x^2</code> （三項演算子）</li>
</ul>

<h3>TikZ数式の特徴</h3>
<ul>
    <li>変数と定数の積は <code>2*x</code> のように<b>明示的に乗算記号を記述</b>してください</li>
    <li>べき乗は <code>x^2</code> のように <code>^</code> を使います</li>
    <li>角度はラジアンで指定します（piを使用: <code>sin(pi/2)</code>）</li>
    <li>関数は入れ子にできます（<code>sin(cos(x))</code>）</li>
</ul>

<h3>よく使う関数の組み合わせ</h3>
<ul>
    <li><b>正規分布:</b> <code>exp(-((x-5)^2)/(2*2^2))</code></li>
    <li><b>対数スケール:</b> <code>ln(1+x)</code></li>
    <li><b>信号処理:</b> <code>sin(2*pi*x)*exp(-x/5)</code>（減衰振動）</li>
    <li><b>ロジスティック成長:</b> <code>10/(1+exp(-x+5))</code></li>
</ul>

<h3>注意点</h3>
<p>TikZの数式はTeX環境で評価されるため、Pythonなど他の言語とは一部記法が異なります。</p>
<p>特に以下の点に注意してください：</p>
<ul>
    <li>数字の後に変数が来る場合は必ず <code>*</code> を使う（<code>2x</code> ではなく <code>2*x</code>）</li>
    <li>ラジアンと度の変換に注意（<code>sin(90)</code> ではなく <code>sin(pi/2)</code>）</li>
    <li>関数名は英語表記（<code>tan</code>, <code>exp</code> など）</li>
</ul>
    """
        
        # 詳細ガイドウィンドウの作成
        help_dialog = QMessageBox(self)
        help_dialog.setWindowTitle("TikZ数式詳細ガイド")
        help_dialog.setTextFormat(Qt.RichText)
        help_dialog.setText(help_text)
        help_dialog.setStandardButtons(QMessageBox.Ok)
        help_dialog.setMinimumWidth(600)
        help_dialog.exec_()

    def on_tangent_color_changed(self, color_text):
        """日本語カラー名から英語カラー名に変換して設定"""
        color_map = {
            '赤': 'red',
            '青': 'blue',
            '緑': 'green',
            '紫': 'purple',
            '橙': 'orange',
            '黒': 'black',
            '水色': 'cyan',
            'マゼンタ': 'magenta',
            '茶色': 'brown',
            '灰色': 'gray'
        }
        
        # マッピングから英語名を取得（なければそのまま使用）
        color_name = color_map.get(color_text, color_text)
        self.tangentColor = QColor(color_name)
        self.tangentColorButton.setStyleSheet(f'background-color: {self.tangentColor.name()};')
        self.statusBar.showMessage(f"接線の色を {color_text} に設定しました", 2000)

    def save_manual_data(self):
        if self.current_dataset_index < 0 or not self.datasets:
            return
        if not self.manualRadio.isChecked():
            return
        dataset = self.datasets[self.current_dataset_index]
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
        dataset['data_x'] = data_x
        dataset['data_y'] = data_y
        self.statusBar.showMessage(f"データセット '{dataset['name']}' の手入力データを保存しました", 3000)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TikZPlotConverter()
    ex.show()
    sys.exit(app.exec_())