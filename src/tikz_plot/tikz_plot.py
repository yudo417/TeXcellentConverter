import os
import numpy as np
import pandas as pd
import sys
import math  # 数学関数を使用するためにインポート
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout,
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
        
        # データと状態の初期化
        self.datasets = []  # 複数のデータセットを格納するリスト
        self.current_dataset_index = -1  # 現在選択されているデータセットのインデックス
        self.statusBar = None
        
        # グラフ全体の設定
        self.global_settings = {
            'x_label': 'x軸',
            'y_label': 'y軸',
            'x_min': 0,
            'x_max': 10,
            'y_min': 0,
            'y_max': 10,
            'grid': True,
            'show_legend': True,
            'legend_pos': 'north east',  # 凡例の位置
            'width': 0.8,
            'height': 0.6,
            'caption': 'グラフのキャプション',
            'label': 'fig:tikz_plot',
            'position': 'htbp',
            'scale_type': 'normal'  # normal, logx, logy, loglog のいずれか
        }
        
        # UIを初期化
        self.initUI()
        
        # 初期データセットを追加（UIの初期化後に呼び出す）
        QTimer.singleShot(0, lambda: self.add_dataset("データセット1"))

    def set_status_bar(self, status_bar):
        """Set the status bar for displaying messages"""
        self.statusBar = status_bar

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
        # メインレイアウト
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
            if self.statusBar:
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
        self.datasetList.setMinimumHeight(100)
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
        sheetLayout.addStretch()
        
        # 手動入力
        manualLayout = QHBoxLayout()
        self.manualRadio = QRadioButton("手動入力:")
        manualLayout.addWidget(self.manualRadio)
        manualLayout.addStretch()
        
        # ラジオボタングループ
        self.sourceButtonGroup = QButtonGroup(self)
        self.sourceButtonGroup.addButton(self.csvRadio)
        self.sourceButtonGroup.addButton(self.excelRadio)
        self.sourceButtonGroup.addButton(self.manualRadio)
        self.csvRadio.setChecked(True)
        
        # イベントハンドラー
        self.csvRadio.toggled.connect(self.toggle_source_fields)
        self.excelRadio.toggled.connect(self.toggle_source_fields)
        self.manualRadio.toggled.connect(self.toggle_source_fields)
        
        dataSourceLayout.addLayout(csvLayout)
        dataSourceLayout.addLayout(excelLayout)
        dataSourceLayout.addLayout(sheetLayout)
        dataSourceLayout.addLayout(manualLayout)
        dataSourceGroup.setLayout(dataSourceLayout)
        
        # 列・範囲指定
        rangeGroup = QGroupBox("データ範囲")
        rangeLayout = QGridLayout()
        
        rangeLayout.addWidget(QLabel("X列・範囲:"), 0, 0)
        self.xRangeEntry = QLineEdit()
        self.xRangeEntry.setPlaceholderText("A:A または A1:A10")
        rangeLayout.addWidget(self.xRangeEntry, 0, 1)
        
        rangeLayout.addWidget(QLabel("Y列・範囲:"), 1, 0)
        self.yRangeEntry = QLineEdit()
        self.yRangeEntry.setPlaceholderText("B:B または B1:B10")
        rangeLayout.addWidget(self.yRangeEntry, 1, 1)
        
        rangeGroup.setLayout(rangeLayout)
        
        # データ読み込みボタン
        loadDataButton = QPushButton("データを読み込み")
        loadDataButton.clicked.connect(self.load_data)
        loadDataButton.setStyleSheet("background-color: #17a2b8; color: white; font-weight: bold; padding: 8px;")
        
        # データプレビューテーブル
        self.dataTable = QTableWidget()
        self.dataTable.setMaximumHeight(150)
        
        # ヘッダーサイズを内容に合わせる
        header = self.dataTable.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # 手動入力用のボタン
        manualButtonLayout = QHBoxLayout()
        addRowButton = QPushButton("行追加")
        addRowButton.clicked.connect(self.add_table_row)
        removeRowButton = QPushButton("行削除")
        removeRowButton.clicked.connect(self.remove_table_row)
        addColButton = QPushButton("列追加")
        addColButton.clicked.connect(self.add_table_column)
        manualButtonLayout.addWidget(addRowButton)
        manualButtonLayout.addWidget(removeRowButton)
        manualButtonLayout.addWidget(addColButton)
        manualButtonLayout.addStretch()
        
        # レイアウトに追加
        measuredLayout.addWidget(dataSourceGroup)
        measuredLayout.addWidget(rangeGroup)
        measuredLayout.addWidget(loadDataButton)
        measuredLayout.addWidget(QLabel("データプレビュー:"))
        measuredLayout.addWidget(self.dataTable)
        measuredLayout.addLayout(manualButtonLayout)
        
        # 数式コンテナ
        self.formulaContainer = QWidget()
        formulaLayout = QVBoxLayout(self.formulaContainer)
        
        # 数式入力部分
        formulaGroup = QGroupBox("数式設定")
        formulaGroupLayout = QVBoxLayout()
        
        # 数式入力
        formulaInputLayout = QHBoxLayout()
        formulaInputLayout.addWidget(QLabel("y = "))
        self.formulaEntry = QLineEdit()
        self.formulaEntry.setPlaceholderText("例: sin(x), x^2, exp(-x^2)")
        formulaInputLayout.addWidget(self.formulaEntry)
        
        # x範囲設定
        xRangeFormLayout = QHBoxLayout()
        xRangeFormLayout.addWidget(QLabel("x範囲:"))
        self.xMinFormEntry = QDoubleSpinBox()
        self.xMinFormEntry.setRange(-1000, 1000)
        self.xMinFormEntry.setValue(0)
        xRangeFormLayout.addWidget(self.xMinFormEntry)
        xRangeFormLayout.addWidget(QLabel("～"))
        self.xMaxFormEntry = QDoubleSpinBox()
        self.xMaxFormEntry.setRange(-1000, 1000)
        self.xMaxFormEntry.setValue(10)
        xRangeFormLayout.addWidget(self.xMaxFormEntry)
        
        # ポイント数設定
        pointsLayout = QHBoxLayout()
        pointsLayout.addWidget(QLabel("ポイント数:"))
        self.pointsSpinBox = QSpinBox()
        self.pointsSpinBox.setRange(10, 1000)
        self.pointsSpinBox.setValue(100)
        pointsLayout.addWidget(self.pointsSpinBox)
        pointsLayout.addStretch()
        
        # 数式適用ボタン
        applyFormulaButton = QPushButton("数式を適用")
        applyFormulaButton.clicked.connect(self.apply_formula)
        applyFormulaButton.setStyleSheet("background-color: #28a745; color: white; font-weight: bold; padding: 8px;")
        
        formulaGroupLayout.addLayout(formulaInputLayout)
        formulaGroupLayout.addLayout(xRangeFormLayout)
        formulaGroupLayout.addLayout(pointsLayout)
        formulaGroupLayout.addWidget(applyFormulaButton)
        formulaGroup.setLayout(formulaGroupLayout)
        formulaLayout.addWidget(formulaGroup)
        
        # 数式コンテナは初期状態で非表示
        self.formulaContainer.setVisible(False)
        
        # データタブに両方のコンテナを追加
        dataTabLayout.addWidget(self.measuredContainer)
        dataTabLayout.addWidget(self.formulaContainer)
        dataTabLayout.addStretch()
        dataTab.setLayout(dataTabLayout)
        
        # タブ2: プロット設定
        plotTab = QWidget()
        plotTabLayout = QVBoxLayout()
        
        # プロットタイプ
        plotTypeGroup = QGroupBox("プロットタイプ")
        plotTypeLayout = QHBoxLayout()
        
        plotTypeLayout.addWidget(QLabel("タイプ:"))
        self.plotTypeCombo = QComboBox()
        self.plotTypeCombo.addItems(['線グラフ', '散布図', '線+マーカー'])
        plotTypeLayout.addWidget(self.plotTypeCombo)
        plotTypeLayout.addStretch()
        
        plotTypeGroup.setLayout(plotTypeLayout)
        
        # スタイル設定
        styleGroup = QGroupBox("スタイル設定")
        styleLayout = QGridLayout()
        
        # 色設定
        styleLayout.addWidget(QLabel("色:"), 0, 0)
        self.colorButton = QPushButton()
        self.colorButton.setFixedSize(50, 30)
        self.colorButton.setStyleSheet("background-color: blue;")
        self.colorButton.clicked.connect(self.select_color)
        self.current_color = QColor('blue')
        styleLayout.addWidget(self.colorButton, 0, 1)
        
        # 線の太さ
        styleLayout.addWidget(QLabel("線の太さ:"), 0, 2)
        self.lineWidthSpinBox = QDoubleSpinBox()
        self.lineWidthSpinBox.setRange(0.1, 5.0)
        self.lineWidthSpinBox.setSingleStep(0.1)
        self.lineWidthSpinBox.setValue(1.0)
        styleLayout.addWidget(self.lineWidthSpinBox, 0, 3)
        
        # マーカースタイル
        styleLayout.addWidget(QLabel("マーカー:"), 1, 0)
        self.markerCombo = QComboBox()
        self.markerCombo.addItems(['なし', '○', '□', '△', '×', '+', '*'])
        styleLayout.addWidget(self.markerCombo, 1, 1)
        
        # マーカーサイズ
        styleLayout.addWidget(QLabel("マーカーサイズ:"), 1, 2)
        self.markerSizeSpinBox = QDoubleSpinBox()
        self.markerSizeSpinBox.setRange(0.1, 10.0)
        self.markerSizeSpinBox.setSingleStep(0.1)
        self.markerSizeSpinBox.setValue(2.0)
        styleLayout.addWidget(self.markerSizeSpinBox, 1, 3)
        
        styleGroup.setLayout(styleLayout)
        
        # 凡例設定
        legendGroup = QGroupBox("凡例設定")
        legendLayout = QHBoxLayout()
        
        self.showLegendCheck = QCheckBox("凡例を表示")
        self.showLegendCheck.setChecked(True)
        legendLayout.addWidget(self.showLegendCheck)
        
        legendLayout.addWidget(QLabel("ラベル:"))
        self.legendEntry = QLineEdit("データ系列")
        legendLayout.addWidget(self.legendEntry)
        
        legendGroup.setLayout(legendLayout)
        
        plotTabLayout.addWidget(plotTypeGroup)
        plotTabLayout.addWidget(styleGroup)
        plotTabLayout.addWidget(legendGroup)
        plotTabLayout.addStretch()
        plotTab.setLayout(plotTabLayout)
        
        # タブ3: 軸設定
        axisTab = QWidget()
        axisTabLayout = QVBoxLayout()
        
        # 軸ラベル
        axisLabelGroup = QGroupBox("軸ラベル")
        axisLabelLayout = QGridLayout()
        
        axisLabelLayout.addWidget(QLabel("X軸ラベル:"), 0, 0)
        self.xLabelEntry = QLineEdit("x軸")
        axisLabelLayout.addWidget(self.xLabelEntry, 0, 1)
        
        axisLabelLayout.addWidget(QLabel("Y軸ラベル:"), 1, 0)
        self.yLabelEntry = QLineEdit("y軸")
        axisLabelLayout.addWidget(self.yLabelEntry, 1, 1)
        
        axisLabelGroup.setLayout(axisLabelLayout)
        
        # 軸範囲
        axisRangeGroup = QGroupBox("軸範囲")
        axisRangeLayout = QGridLayout()
        
        axisRangeLayout.addWidget(QLabel("X軸最小値:"), 0, 0)
        self.xMinSpinBox = QDoubleSpinBox()
        self.xMinSpinBox.setRange(-1000, 1000)
        self.xMinSpinBox.setValue(0)
        axisRangeLayout.addWidget(self.xMinSpinBox, 0, 1)
        
        axisRangeLayout.addWidget(QLabel("X軸最大値:"), 0, 2)
        self.xMaxSpinBox = QDoubleSpinBox()
        self.xMaxSpinBox.setRange(-1000, 1000)
        self.xMaxSpinBox.setValue(10)
        axisRangeLayout.addWidget(self.xMaxSpinBox, 0, 3)
        
        axisRangeLayout.addWidget(QLabel("Y軸最小値:"), 1, 0)
        self.yMinSpinBox = QDoubleSpinBox()
        self.yMinSpinBox.setRange(-1000, 1000)
        self.yMinSpinBox.setValue(0)
        axisRangeLayout.addWidget(self.yMinSpinBox, 1, 1)
        
        axisRangeLayout.addWidget(QLabel("Y軸最大値:"), 1, 2)
        self.yMaxSpinBox = QDoubleSpinBox()
        self.yMaxSpinBox.setRange(-1000, 1000)
        self.yMaxSpinBox.setValue(10)
        axisRangeLayout.addWidget(self.yMaxSpinBox, 1, 3)
        
        axisRangeGroup.setLayout(axisRangeLayout)
        
        # グリッド設定
        gridGroup = QGroupBox("グリッド・その他")
        gridLayout = QVBoxLayout()
        
        self.gridCheck = QCheckBox("グリッドを表示")
        self.gridCheck.setChecked(True)
        gridLayout.addWidget(self.gridCheck)
        
        # スケールタイプ
        scaleLayout = QHBoxLayout()
        scaleLayout.addWidget(QLabel("スケール:"))
        self.scaleCombo = QComboBox()
        self.scaleCombo.addItems(['通常', '対数X軸', '対数Y軸', '両対数'])
        scaleLayout.addWidget(self.scaleCombo)
        scaleLayout.addStretch()
        gridLayout.addLayout(scaleLayout)
        
        gridGroup.setLayout(gridLayout)
        
        axisTabLayout.addWidget(axisLabelGroup)
        axisTabLayout.addWidget(axisRangeGroup)
        axisTabLayout.addWidget(gridGroup)
        axisTabLayout.addStretch()
        axisTab.setLayout(axisTabLayout)
        
        # タブ4: 図の設定
        figureTab = QWidget()
        figureTabLayout = QVBoxLayout()
        
        # 図のサイズ
        figSizeGroup = QGroupBox("図のサイズ")
        figSizeLayout = QGridLayout()
        
        figSizeLayout.addWidget(QLabel("幅:"), 0, 0)
        self.widthSpinBox = QDoubleSpinBox()
        self.widthSpinBox.setRange(0.1, 2.0)
        self.widthSpinBox.setSingleStep(0.1)
        self.widthSpinBox.setValue(0.8)
        figSizeLayout.addWidget(self.widthSpinBox, 0, 1)
        
        figSizeLayout.addWidget(QLabel("高さ:"), 0, 2)
        self.heightSpinBox = QDoubleSpinBox()
        self.heightSpinBox.setRange(0.1, 2.0)
        self.heightSpinBox.setSingleStep(0.1)
        self.heightSpinBox.setValue(0.6)
        figSizeLayout.addWidget(self.heightSpinBox, 0, 3)
        
        figSizeGroup.setLayout(figSizeLayout)
        
        # キャプション・ラベル
        captionGroup = QGroupBox("キャプション・ラベル")
        captionLayout = QGridLayout()
        
        captionLayout.addWidget(QLabel("キャプション:"), 0, 0)
        self.captionEntry = QLineEdit("グラフのキャプション")
        captionLayout.addWidget(self.captionEntry, 0, 1)
        
        captionLayout.addWidget(QLabel("ラベル:"), 1, 0)
        self.labelEntry = QLineEdit("fig:tikz_plot")
        captionLayout.addWidget(self.labelEntry, 1, 1)
        
        captionLayout.addWidget(QLabel("位置:"), 2, 0)
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['htbp', 'h', 't', 'b', 'p', 'H'])
        captionLayout.addWidget(self.positionCombo, 2, 1)
        
        captionGroup.setLayout(captionLayout)
        
        figureTabLayout.addWidget(figSizeGroup)
        figureTabLayout.addWidget(captionGroup)
        figureTabLayout.addStretch()
        figureTab.setLayout(figureTabLayout)
        
        # タブをタブウィジェットに追加
        tabWidget.addTab(dataTab, "データ")
        tabWidget.addTab(plotTab, "プロット")
        tabWidget.addTab(axisTab, "軸設定")
        tabWidget.addTab(figureTab, "図の設定")
        
        # 変換ボタン
        convertButton = QPushButton('TikZコードを生成')
        convertButton.clicked.connect(self.convert_to_tikz)
        convertButton.setStyleSheet('background-color: #4CAF50; color: white; font-size: 14px; padding: 10px; font-weight: bold;')
        
        # 設定部分のレイアウトに追加
        settingsLayout.addLayout(infoLayout)
        settingsLayout.addWidget(datasetGroup)
        settingsLayout.addWidget(tabWidget)
        settingsLayout.addWidget(convertButton)
        
        # --- 下部：結果表示部分 ---
        resultWidget = QWidget()
        resultLayout = QVBoxLayout()
        resultWidget.setLayout(resultLayout)
        
        resultLabel = QLabel("TikZ コード:")
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(200)
        
        copyButton = QPushButton("クリップボードにコピー")
        copyButton.clicked.connect(self.copy_to_clipboard)
        copyButton.setStyleSheet("background-color: #007BFF; color: white; font-size: 16px; padding: 12px; font-weight: bold; border-radius: 8px;")
        
        resultLayout.addWidget(resultLabel)
        resultLayout.addWidget(self.resultText)
        resultLayout.addWidget(copyButton)
        
        # スプリッターに追加
        splitter.addWidget(settingsWidget)
        splitter.addWidget(resultWidget)
        splitter.setSizes([400, 300])  # 初期サイズ比率
        
        # メインレイアウトに追加
        mainLayout.addWidget(splitter)
        
        # メインウィジェット設定
        self.setLayout(mainLayout)
        
        # 初期状態設定
        self.toggle_source_fields()

    def browse_csv_file(self):
        """CSVファイルを参照"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "CSVファイルを選択", "", "CSV Files (*.csv)")
        if file_path:
            self.fileEntry.setText(file_path)

    def browse_excel_file(self):
        """Excelファイルを参照"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excelEntry.setText(file_path)
            self.update_sheet_names(file_path)

    def update_sheet_names(self, file_path):
        """Excelファイルのシート名を更新"""
        try:
            from pandas import ExcelFile
            xls = ExcelFile(file_path)
            self.sheetCombobox.clear()
            self.sheetCombobox.addItems(xls.sheet_names)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"シート名の取得に失敗しました: {str(e)}")

    def toggle_source_fields(self, button=None):
        """データソースに応じてフィールドの有効/無効を切り替え"""
        csv_enabled = self.csvRadio.isChecked()
        excel_enabled = self.excelRadio.isChecked()
        manual_enabled = self.manualRadio.isChecked()
        
        # CSV関連
        self.fileEntry.setEnabled(csv_enabled)
        
        # Excel関連
        self.excelEntry.setEnabled(excel_enabled)
        self.sheetCombobox.setEnabled(excel_enabled)
        
        # 範囲設定は手動入力以外で有効
        range_enabled = not manual_enabled
        self.xRangeEntry.setEnabled(range_enabled)
        self.yRangeEntry.setEnabled(range_enabled)

    def load_data(self):
        """データを読み込んでテーブルに表示"""
        try:
            data = None
            
            if self.csvRadio.isChecked():
                file_path = self.fileEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なCSVファイルを選択してください")
                    return
                data = pd.read_csv(file_path)
                
            elif self.excelRadio.isChecked():
                file_path = self.excelEntry.text()
                sheet_name = self.sheetCombobox.currentText()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なExcelファイルを選択してください")
                    return
                data = pd.read_excel(file_path, sheet_name=sheet_name)
                
            elif self.manualRadio.isChecked():
                # 手動入力の場合は空のテーブルを作成
                self.dataTable.setRowCount(5)
                self.dataTable.setColumnCount(2)
                self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
                return
            
            if data is not None:
                # 範囲指定がある場合は該当データを抽出
                x_range = self.xRangeEntry.text().strip()
                y_range = self.yRangeEntry.text().strip()
                
                if x_range and y_range:
                    x_data, y_data = self.extract_data_from_range(data, x_range, y_range)
                    if x_data is not None and y_data is not None:
                        display_data = pd.DataFrame({'X': x_data, 'Y': y_data})
                    else:
                        return
                else:
                    display_data = data.head(100)  # 最初の100行のみ表示
                
                # テーブルに表示
                self.dataTable.setRowCount(len(display_data))
                self.dataTable.setColumnCount(len(display_data.columns))
                self.dataTable.setHorizontalHeaderLabels(display_data.columns.astype(str))
                
                for i in range(len(display_data)):
                    for j in range(len(display_data.columns)):
                        item = QTableWidgetItem(str(display_data.iloc[i, j]))
                        self.dataTable.setItem(i, j, item)
                
                if self.statusBar:
                    self.statusBar.showMessage("データが読み込まれました", 3000)
                    
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データの読み込みに失敗しました: {str(e)}")

    def extract_data_from_range(self, df, x_range, y_range):
        """指定された範囲からデータを抽出"""
        def parse_range(range_str):
            # A1:A10 のような形式を解析
            if ':' in range_str:
                start, end = range_str.split(':')
                # 列名のみの場合（A:A）
                if start.isalpha() and end.isalpha():
                    return start, None, end, None
                # セル範囲の場合（A1:A10）
                else:
                    # 列名と行番号を分離
                    def col_to_index(col_str):
                        col_letters = ''.join(filter(str.isalpha, col_str))
                        col_index = 0
                        for char in col_letters:
                            col_index = col_index * 26 + (ord(char.upper()) - ord('A') + 1)
                        return col_index - 1
                    
                    start_col = ''.join(filter(str.isalpha, start))
                    start_row = ''.join(filter(str.isdigit, start))
                    end_col = ''.join(filter(str.isalpha, end))
                    end_row = ''.join(filter(str.isdigit, end))
                    
                    return start_col, int(start_row) if start_row else None, end_col, int(end_row) if end_row else None
            else:
                # 単一列の場合
                return range_str, None, range_str, None
        
        try:
            # X範囲の解析
            x_start_col, x_start_row, x_end_col, x_end_row = parse_range(x_range)
            # Y範囲の解析
            y_start_col, y_start_row, y_end_col, y_end_row = parse_range(y_range)
            
            # データフレームから該当列を取得
            if x_start_col.isalpha():
                # Excel列名の場合、数値インデックスに変換
                def excel_col_to_index(col):
                    index = 0
                    for char in col.upper():
                        index = index * 26 + (ord(char) - ord('A') + 1)
                    return index - 1
                
                x_col_idx = excel_col_to_index(x_start_col)
                y_col_idx = excel_col_to_index(y_start_col)
                
                if x_col_idx < len(df.columns) and y_col_idx < len(df.columns):
                    x_data = df.iloc[:, x_col_idx]
                    y_data = df.iloc[:, y_col_idx]
                else:
                    QMessageBox.warning(self, "警告", "指定された列が存在しません")
                    return None, None
            else:
                # 列名で指定
                if x_start_col in df.columns and y_start_col in df.columns:
                    x_data = df[x_start_col]
                    y_data = df[y_start_col]
                else:
                    QMessageBox.warning(self, "警告", "指定された列名が存在しません")
                    return None, None
            
            # 行範囲の適用
            if x_start_row is not None and x_end_row is not None:
                x_data = x_data.iloc[x_start_row-1:x_end_row]
            if y_start_row is not None and y_end_row is not None:
                y_data = y_data.iloc[y_start_row-1:y_end_row]
            
            # NaN値を除去
            combined = pd.DataFrame({'x': x_data, 'y': y_data}).dropna()
            
            return combined['x'].values, combined['y'].values
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"範囲指定の解析に失敗しました: {str(e)}")
            return None, None

    def add_table_row(self):
        """テーブルに行を追加"""
        current_rows = self.dataTable.rowCount()
        self.dataTable.setRowCount(current_rows + 1)

    def remove_table_row(self):
        """テーブルから行を削除"""
        current_row = self.dataTable.currentRow()
        if current_row >= 0:
            self.dataTable.removeRow(current_row)

    def add_table_column(self):
        """テーブルに列を追加"""
        current_cols = self.dataTable.columnCount()
        self.dataTable.setColumnCount(current_cols + 1)
        self.dataTable.setHorizontalHeaderItem(current_cols, QTableWidgetItem(f"Column {current_cols + 1}"))

    def select_color(self):
        """色を選択"""
        color = QColorDialog.getColor(self.current_color, self, "色を選択")
        if color.isValid():
            self.current_color = color
            self.colorButton.setStyleSheet(f"background-color: {color.name()};")

    def on_data_source_type_changed(self, checked):
        """データソースタイプの変更処理"""
        if checked:  # チェックされた時のみ処理
            if self.measuredRadio.isChecked():
                self.measuredContainer.setVisible(True)
                self.formulaContainer.setVisible(False)
            else:  # formula
                self.measuredContainer.setVisible(False)
                self.formulaContainer.setVisible(True)

    def apply_formula(self):
        """数式を適用してデータを生成"""
        try:
            formula = self.formulaEntry.text().strip()
            if not formula:
                QMessageBox.warning(self, "警告", "数式を入力してください")
                return
            
            x_min = self.xMinFormEntry.value()
            x_max = self.xMaxFormEntry.value()
            num_points = self.pointsSpinBox.value()
            
            if x_min >= x_max:
                QMessageBox.warning(self, "警告", "x範囲が無効です（最小値 < 最大値）")
                return
            
            # x値を生成
            x_values = np.linspace(x_min, x_max, num_points)
            
            # 数式を評価
            formula = formula.replace('^', '**')  # べき乗を変換
            formula = formula.replace('sin', 'np.sin')
            formula = formula.replace('cos', 'np.cos')
            formula = formula.replace('tan', 'np.tan')
            formula = formula.replace('exp', 'np.exp')
            formula = formula.replace('log', 'np.log')
            formula = formula.replace('sqrt', 'np.sqrt')
            
            # 安全な評価のための環境
            safe_dict = {"x": x_values, "np": np, "math": math}
            
            try:
                y_values = eval(formula, {"__builtins__": {}}, safe_dict)
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"数式の評価に失敗しました: {str(e)}")
                return
            
            # テーブルに表示
            self.dataTable.setRowCount(len(x_values))
            self.dataTable.setColumnCount(2)
            self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
            
            for i in range(len(x_values)):
                self.dataTable.setItem(i, 0, QTableWidgetItem(f"{x_values[i]:.6f}"))
                self.dataTable.setItem(i, 1, QTableWidgetItem(f"{y_values[i]:.6f}"))
            
            if self.statusBar:
                self.statusBar.showMessage("数式データが生成されました", 3000)
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"数式の適用に失敗しました: {str(e)}")

    def add_dataset(self, name_arg=None):
        """新しいデータセットを追加"""
        if name_arg is None:
            name, ok = QInputDialog.getText(self, 'データセット追加', 'データセット名を入力:')
            if not ok or not name.strip():
                return
        else:
            name = name_arg
        
        # データセットを作成
        dataset = {
            'name': name,
            'data_source_type': 'measured',  # measured or formula
            'file_path': '',
            'sheet_name': '',
            'x_range': '',
            'y_range': '',
            'x_data': [],
            'y_data': [],
            'plot_type': '線グラフ',
            'color': QColor('blue'),
            'line_width': 1.0,
            'marker_style': 'なし',
            'marker_size': 2.0,
            'show_legend': True,
            'legend_label': name,
            'formula': '',
            'x_min_form': 0,
            'x_max_form': 10,
            'num_points': 100
        }
        
        self.datasets.append(dataset)
        self.datasetList.addItem(name)
        self.datasetList.setCurrentRow(len(self.datasets) - 1)

    def remove_dataset(self):
        """選択されたデータセットを削除"""
        current_row = self.datasetList.currentRow()
        if current_row >= 0:
            reply = QMessageBox.question(self, '確認', 
                                       f'データセット "{self.datasets[current_row]["name"]}" を削除しますか？',
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                del self.datasets[current_row]
                self.datasetList.takeItem(current_row)
                
                if len(self.datasets) == 0:
                    self.current_dataset_index = -1
                else:
                    # 削除後のインデックス調整
                    if current_row >= len(self.datasets):
                        self.datasetList.setCurrentRow(len(self.datasets) - 1)
                    else:
                        self.datasetList.setCurrentRow(current_row)

    def rename_dataset(self):
        """データセット名を変更"""
        current_row = self.datasetList.currentRow()
        if current_row >= 0:
            old_name = self.datasets[current_row]['name']
            name, ok = QInputDialog.getText(self, 'データセット名変更', 
                                          'データセット名を入力:', text=old_name)
            if ok and name.strip():
                self.datasets[current_row]['name'] = name.strip()
                self.datasets[current_row]['legend_label'] = name.strip()
                self.datasetList.item(current_row).setText(name.strip())

    def on_dataset_selected(self, row):
        """データセットが選択された時の処理"""
        if row >= 0 and row < len(self.datasets):
            # 現在のデータセットの設定を保存
            if self.current_dataset_index >= 0:
                self.update_current_dataset()
            
            # 新しいデータセットの設定をUIに反映
            self.current_dataset_index = row
            self.update_ui_from_dataset(self.datasets[row])

    def update_current_dataset(self):
        """現在のUIの設定を現在のデータセットに保存"""
        if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
            dataset = self.datasets[self.current_dataset_index]
            
            # データソースタイプ
            dataset['data_source_type'] = 'measured' if self.measuredRadio.isChecked() else 'formula'
            
            # ファイル情報
            if self.csvRadio.isChecked():
                dataset['file_path'] = self.fileEntry.text()
            elif self.excelRadio.isChecked():
                dataset['file_path'] = self.excelEntry.text()
                dataset['sheet_name'] = self.sheetCombobox.currentText()
            
            dataset['x_range'] = self.xRangeEntry.text()
            dataset['y_range'] = self.yRangeEntry.text()
            
            # プロット設定
            dataset['plot_type'] = self.plotTypeCombo.currentText()
            dataset['color'] = self.current_color
            dataset['line_width'] = self.lineWidthSpinBox.value()
            dataset['marker_style'] = self.markerCombo.currentText()
            dataset['marker_size'] = self.markerSizeSpinBox.value()
            dataset['show_legend'] = self.showLegendCheck.isChecked()
            dataset['legend_label'] = self.legendEntry.text()
            
            # 数式設定
            dataset['formula'] = self.formulaEntry.text()
            dataset['x_min_form'] = self.xMinFormEntry.value()
            dataset['x_max_form'] = self.xMaxFormEntry.value()
            dataset['num_points'] = self.pointsSpinBox.value()
            
            # テーブルデータの保存
            x_data = []
            y_data = []
            for i in range(self.dataTable.rowCount()):
                x_item = self.dataTable.item(i, 0)
                y_item = self.dataTable.item(i, 1)
                if x_item and y_item:
                    try:
                        x_val = float(x_item.text())
                        y_val = float(y_item.text())
                        x_data.append(x_val)
                        y_data.append(y_val)
                    except:
                        pass
            dataset['x_data'] = x_data
            dataset['y_data'] = y_data

    def update_ui_from_dataset(self, dataset):
        """データセットの設定をUIに反映"""
        # データソースタイプ
        if dataset['data_source_type'] == 'measured':
            self.measuredRadio.setChecked(True)
        else:
            self.formulaRadio.setChecked(True)
        
        # ファイル情報
        self.fileEntry.setText(dataset.get('file_path', ''))
        self.excelEntry.setText(dataset.get('file_path', ''))
        if dataset.get('sheet_name'):
            index = self.sheetCombobox.findText(dataset['sheet_name'])
            if index >= 0:
                self.sheetCombobox.setCurrentIndex(index)
        
        self.xRangeEntry.setText(dataset.get('x_range', ''))
        self.yRangeEntry.setText(dataset.get('y_range', ''))
        
        # プロット設定
        plot_index = self.plotTypeCombo.findText(dataset.get('plot_type', '線グラフ'))
        if plot_index >= 0:
            self.plotTypeCombo.setCurrentIndex(plot_index)
        
        color = dataset.get('color', QColor('blue'))
        self.current_color = color
        self.colorButton.setStyleSheet(f"background-color: {color.name()};")
        
        self.lineWidthSpinBox.setValue(dataset.get('line_width', 1.0))
        
        marker_index = self.markerCombo.findText(dataset.get('marker_style', 'なし'))
        if marker_index >= 0:
            self.markerCombo.setCurrentIndex(marker_index)
        
        self.markerSizeSpinBox.setValue(dataset.get('marker_size', 2.0))
        self.showLegendCheck.setChecked(dataset.get('show_legend', True))
        self.legendEntry.setText(dataset.get('legend_label', dataset['name']))
        
        # 数式設定
        self.formulaEntry.setText(dataset.get('formula', ''))
        self.xMinFormEntry.setValue(dataset.get('x_min_form', 0))
        self.xMaxFormEntry.setValue(dataset.get('x_max_form', 10))
        self.pointsSpinBox.setValue(dataset.get('num_points', 100))
        
        # テーブルデータの復元
        x_data = dataset.get('x_data', [])
        y_data = dataset.get('y_data', [])
        
        if x_data and y_data:
            self.dataTable.setRowCount(len(x_data))
            self.dataTable.setColumnCount(2)
            self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
            
            for i in range(len(x_data)):
                self.dataTable.setItem(i, 0, QTableWidgetItem(str(x_data[i])))
                self.dataTable.setItem(i, 1, QTableWidgetItem(str(y_data[i])))

    def convert_to_tikz(self):
        """TikZコードを生成"""
        try:
            # 現在のデータセットを更新
            if self.current_dataset_index >= 0:
                self.update_current_dataset()
            
            # グローバル設定を更新
            self.update_global_settings()
            
            # TikZコードを生成
            tikz_code = self.generate_tikz_code_multi_datasets()
            
            self.resultText.setPlainText(tikz_code)
            
            if self.statusBar:
                self.statusBar.showMessage("TikZコードが生成されました", 3000)
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"TikZコードの生成に失敗しました: {str(e)}")

    def update_global_settings(self):
        """グローバル設定を更新"""
        self.global_settings['x_label'] = self.xLabelEntry.text()
        self.global_settings['y_label'] = self.yLabelEntry.text()
        self.global_settings['x_min'] = self.xMinSpinBox.value()
        self.global_settings['x_max'] = self.xMaxSpinBox.value()
        self.global_settings['y_min'] = self.yMinSpinBox.value()
        self.global_settings['y_max'] = self.yMaxSpinBox.value()
        self.global_settings['grid'] = self.gridCheck.isChecked()
        self.global_settings['width'] = self.widthSpinBox.value()
        self.global_settings['height'] = self.heightSpinBox.value()
        self.global_settings['caption'] = self.captionEntry.text()
        self.global_settings['label'] = self.labelEntry.text()
        self.global_settings['position'] = self.positionCombo.currentText()
        
        # スケールタイプの設定
        scale_text = self.scaleCombo.currentText()
        if scale_text == '対数X軸':
            self.global_settings['scale_type'] = 'logx'
        elif scale_text == '対数Y軸':
            self.global_settings['scale_type'] = 'logy'
        elif scale_text == '両対数':
            self.global_settings['scale_type'] = 'loglog'
        else:
            self.global_settings['scale_type'] = 'normal'

    def generate_tikz_code_multi_datasets(self):
        """複数のデータセットに対応したTikZコードを生成"""
        if not self.datasets:
            return "% データセットがありません"
        
        latex = []
        
        # 図の開始
        latex.append(f"\\begin{{figure}}[{self.global_settings['position']}]")
        latex.append("\\centering")
        
        # TikZの開始
        latex.append("\\begin{tikzpicture}")
        latex.append("\\begin{axis}[")
        
        # 軸設定
        latex.append(f"    xlabel={{{self.global_settings['x_label']}}},")
        latex.append(f"    ylabel={{{self.global_settings['y_label']}}},")
        latex.append(f"    xmin={self.global_settings['x_min']}, xmax={self.global_settings['x_max']},")
        latex.append(f"    ymin={self.global_settings['y_min']}, ymax={self.global_settings['y_max']},")
        latex.append(f"    width={self.global_settings['width']}\\textwidth,")
        latex.append(f"    height={self.global_settings['height']}\\textwidth,")
        
        # グリッド設定
        if self.global_settings['grid']:
            latex.append("    grid=both,")
            latex.append("    grid style={line width=.1pt, draw=gray!10},")
            latex.append("    major grid style={line width=.2pt,draw=gray!50},")
        
        # スケール設定
        scale_type = self.global_settings['scale_type']
        if scale_type == 'logx':
            latex.append("    xmode=log,")
        elif scale_type == 'logy':
            latex.append("    ymode=log,")
        elif scale_type == 'loglog':
            latex.append("    xmode=log,")
            latex.append("    ymode=log,")
        
        # 凡例設定
        show_any_legend = any(dataset.get('show_legend', False) for dataset in self.datasets)
        if show_any_legend:
            latex.append("    legend pos=north east,")
        
        latex.append("]")
        
        # 各データセットをプロット
        for i, dataset in enumerate(self.datasets):
            if dataset.get('x_data') and dataset.get('y_data'):
                self.add_dataset_to_latex(latex, dataset, i)
        
        # TikZの終了
        latex.append("\\end{axis}")
        latex.append("\\end{tikzpicture}")
        
        # キャプションとラベル
        latex.append(f"\\caption{{{self.global_settings['caption']}}}")
        latex.append(f"\\label{{{self.global_settings['label']}}}")
        latex.append("\\end{figure}")
        
        return "\n".join(latex)

    def add_dataset_to_latex(self, latex, dataset, index):
        """データセットをLaTeXコードに追加"""
        x_data = dataset['x_data']
        y_data = dataset['y_data']
        
        if not x_data or not y_data:
            return
        
        # プロットタイプによる設定
        plot_type = dataset.get('plot_type', '線グラフ')
        color = dataset.get('color', QColor('blue'))
        line_width = dataset.get('line_width', 1.0)
        marker_style = dataset.get('marker_style', 'なし')
        marker_size = dataset.get('marker_size', 2.0)
        
        # TikZ色名に変換
        tikz_color = self.color_to_tikz_rgb(color)
        
        # プロットオプション
        plot_options = []
        
        # 色設定
        if tikz_color in ['red', 'green', 'blue', 'black', 'yellow', 'cyan', 'magenta', 'orange', 'purple', 'brown', 'gray']:
            plot_options.append(f"color={tikz_color}")
        else:
            plot_options.append(tikz_color)
        
        # 線の設定
        if plot_type in ['線グラフ', '線+マーカー']:
            plot_options.append(f"line width={line_width}pt")
        else:
            plot_options.append("only marks")
        
        # マーカーの設定
        if plot_type in ['散布図', '線+マーカー'] and marker_style != 'なし':
            marker_map = {
                '○': 'o',
                '□': 'square',
                '△': 'triangle',
                '×': 'x',
                '+': '+',
                '*': 'asterisk'
            }
            if marker_style in marker_map:
                plot_options.append(f"mark={marker_map[marker_style]}")
                plot_options.append(f"mark size={marker_size}pt")
        
        # addplotコマンド
        options_str = ", ".join(plot_options)
        latex.append(f"\\addplot[{options_str}] coordinates {{")
        
        # データ点
        for x, y in zip(x_data, y_data):
            latex.append(f"    ({x}, {y})")
        
        latex.append("};")
        
        # 凡例
        if dataset.get('show_legend', False):
            legend_label = dataset.get('legend_label', dataset['name'])
            latex.append(f"\\addlegendentry{{{legend_label}}}")

    def copy_to_clipboard(self):
        """結果をクリップボードにコピー"""
        tikz_code = self.resultText.toPlainText()
        if tikz_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(tikz_code)
            if self.statusBar:
                self.statusBar.showMessage("TikZコードをクリップボードにコピーしました", 3000)
        else:
            QMessageBox.warning(self, "警告", "コピーするコードがありません")


# For standalone execution
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TikZPlotTab()
    window.show()
    sys.exit(app.exec_())
