import os
import numpy as np
import pandas as pd
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QLabel, QLineEdit, QPushButton,
                            QFileDialog, QComboBox, QMessageBox, QTextEdit,
                            QCheckBox, QGridLayout, QGroupBox, QSplitter,
                            QStatusBar, QFrame, QTabWidget, QTableWidget,
                            QTableWidgetItem, QColorDialog, QSpinBox, 
                            QDoubleSpinBox, QHeaderView, QListWidget,
                            QFormLayout, QRadioButton, QButtonGroup, QInputDialog)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor


class TikZPlotConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # データと状態の初期化
        self.datasets = []  # 複数のデータセットを格納するリスト
        self.current_dataset_index = -1  # 現在選択されているデータセットのインデックス
        
        # UIを初期化
        self.initUI()
        
        # 初期データセットを追加（UIの初期化後に呼び出す）
        QTimer.singleShot(0, lambda: self.add_dataset("データセット1"))
        
    def initUI(self):
        self.setWindowTitle('TikZPlot Converter')
        self.setGeometry(100, 100, 1000, 700)
        
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
        self.datasetList.setMinimumHeight(100)  # 高さを増やす
        self.datasetList.setSelectionMode(QListWidget.SingleSelection)  # 単一選択モード
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
        
        # ボタングループの設定
        sourceGroup = QButtonGroup(self)
        sourceGroup.addButton(self.csvRadio)
        sourceGroup.addButton(self.excelRadio)
        sourceGroup.buttonClicked.connect(self.toggle_source_fields)
        
        # データ直接入力
        self.manualRadio = QRadioButton("データを直接入力:")
        sourceGroup.addButton(self.manualRadio)
        
        # データテーブル
        self.dataTable = QTableWidget(10, 2)  # 行数, 列数
        self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
        self.dataTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dataTable.setEnabled(False)  # 初期状態では無効
        
        # データテーブル操作ボタン
        tableButtonLayout = QHBoxLayout()
        addRowButton = QPushButton('行を追加')
        addRowButton.clicked.connect(self.add_table_row)
        removeRowButton = QPushButton('選択行を削除')
        removeRowButton.clicked.connect(self.remove_table_row)
        addColumnButton = QPushButton('列を追加')
        addColumnButton.clicked.connect(self.add_table_column)
        tableButtonLayout.addWidget(addRowButton)
        tableButtonLayout.addWidget(removeRowButton)
        tableButtonLayout.addWidget(addColumnButton)
        
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
        
        loadDataButton = QPushButton('データ読み込み')
        loadDataButton.clicked.connect(self.load_data)
        loadDataButton.setStyleSheet('background-color: #4CAF50; color: white;')
        columnLayout.addWidget(loadDataButton, 3, 0, 1, 2)
        
        columnGroup.setLayout(columnLayout)
        
        # データタブレイアウトに追加
        dataTabLayout.addWidget(dataSourceGroup)
        dataTabLayout.addWidget(columnGroup)
        dataTab.setLayout(dataTabLayout)
        
        # タブ2: グラフ設定
        plotTab = QWidget()
        plotTabLayout = QVBoxLayout()
        
        # グラフタイプの選択
        plotTypeGroup = QGroupBox("グラフタイプ")
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
        styleGroup = QGroupBox("スタイル設定")
        styleLayout = QGridLayout()
        
        # 色選択
        colorLabel = QLabel('線の色:')
        self.colorButton = QPushButton()
        self.colorButton.setStyleSheet('background-color: blue;')
        self.currentColor = QColor('blue')
        self.colorButton.clicked.connect(self.select_color)
        styleLayout.addWidget(colorLabel, 0, 0)
        styleLayout.addWidget(self.colorButton, 0, 1)
        
        # 線の太さ
        lineWidthLabel = QLabel('線の太さ:')
        self.lineWidthSpin = QDoubleSpinBox()
        self.lineWidthSpin.setRange(0.1, 5.0)
        self.lineWidthSpin.setSingleStep(0.1)
        self.lineWidthSpin.setValue(1.0)
        styleLayout.addWidget(lineWidthLabel, 1, 0)
        styleLayout.addWidget(self.lineWidthSpin, 1, 1)
        
        # マーカースタイル
        markerLabel = QLabel('マーカースタイル:')
        self.markerCombo = QComboBox()
        self.markerCombo.addItems(['*', 'o', 'square', 'triangle', 'diamond', '+', 'x'])
        styleLayout.addWidget(markerLabel, 2, 0)
        styleLayout.addWidget(self.markerCombo, 2, 1)
        
        # マーカーサイズ
        markerSizeLabel = QLabel('マーカーサイズ:')
        self.markerSizeSpin = QDoubleSpinBox()
        self.markerSizeSpin.setRange(0.5, 10.0)
        self.markerSizeSpin.setSingleStep(0.5)
        self.markerSizeSpin.setValue(3.0)
        styleLayout.addWidget(markerSizeLabel, 3, 0)
        styleLayout.addWidget(self.markerSizeSpin, 3, 1)
        
        # データ点をマークで表示するオプション
        self.showDataPointsCheck = QCheckBox('データ点をマークで表示（線グラフでも点を表示）')
        self.showDataPointsCheck.setChecked(False)
        styleLayout.addWidget(self.showDataPointsCheck, 4, 0, 1, 2)
        
        styleGroup.setLayout(styleLayout)
        
        # 軸設定
        axisGroup = QGroupBox("軸設定")
        axisLayout = QGridLayout()
        
        # X軸ラベル
        xLabelLabel = QLabel('X軸ラベル:')
        self.xLabelEntry = QLineEdit('x')
        axisLayout.addWidget(xLabelLabel, 0, 0)
        axisLayout.addWidget(self.xLabelEntry, 0, 1)
        
        # Y軸ラベル
        yLabelLabel = QLabel('Y軸ラベル:')
        self.yLabelEntry = QLineEdit('y')
        axisLayout.addWidget(yLabelLabel, 1, 0)
        axisLayout.addWidget(self.yLabelEntry, 1, 1)
        
        # X軸範囲
        xRangeLabel = QLabel('X軸範囲:')
        xRangeLayout = QHBoxLayout()
        self.xMinSpin = QDoubleSpinBox()
        self.xMinSpin.setRange(-1000, 1000)
        self.xMinSpin.setValue(0)
        self.xMaxSpin = QDoubleSpinBox()
        self.xMaxSpin.setRange(-1000, 1000)
        self.xMaxSpin.setValue(10)
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
        self.yMinSpin.setValue(0)
        self.yMaxSpin = QDoubleSpinBox()
        self.yMaxSpin.setRange(-1000, 1000)
        self.yMaxSpin.setValue(10)
        yRangeLayout.addWidget(self.yMinSpin)
        yRangeLayout.addWidget(QLabel('〜'))
        yRangeLayout.addWidget(self.yMaxSpin)
        axisLayout.addWidget(yRangeLabel, 3, 0)
        axisLayout.addLayout(yRangeLayout, 3, 1)
        
        # グリッド表示
        self.gridCheck = QCheckBox('グリッド表示')
        self.gridCheck.setChecked(True)
        axisLayout.addWidget(self.gridCheck, 4, 0, 1, 2)
        
        axisGroup.setLayout(axisLayout)
        
        # 凡例設定
        legendGroup = QGroupBox("凡例設定")
        legendLayout = QFormLayout()
        
        self.legendCheck = QCheckBox('凡例を表示')
        self.legendCheck.setChecked(True)
        
        self.legendLabel = QLineEdit('データ')
        
        legendPosLabel = QLabel('凡例の位置:')
        self.legendPosCombo = QComboBox()
        self.legendPosCombo.addItems(['north west', 'north east', 'south west', 'south east', 'north', 'south', 'east', 'west'])
        self.legendPosCombo.setCurrentText('north east')
        
        legendLayout.addRow(self.legendCheck)
        legendLayout.addRow('凡例ラベル:', self.legendLabel)
        legendLayout.addRow(legendPosLabel, self.legendPosCombo)
        
        legendGroup.setLayout(legendLayout)
        
        # 図の設定
        figureGroup = QGroupBox("図の設定")
        figureLayout = QFormLayout()
        
        # キャプション
        self.captionEntry = QLineEdit('グラフのタイトル')
        
        # ラベル
        self.labelEntry = QLineEdit('fig:plot')
        
        # 位置
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        self.positionCombo.setCurrentText('H')
        
        # 幅と高さ
        widthLayout = QHBoxLayout()
        self.widthSpin = QDoubleSpinBox()
        self.widthSpin.setRange(0.1, 1.0)
        self.widthSpin.setSingleStep(0.1)
        self.widthSpin.setValue(0.8)
        self.widthSpin.setSuffix('\\textwidth')
        
        heightLayout = QHBoxLayout()
        self.heightSpin = QDoubleSpinBox()
        self.heightSpin.setRange(0.1, 1.0)
        self.heightSpin.setSingleStep(0.1)
        self.heightSpin.setValue(0.5)
        self.heightSpin.setSuffix('\\textwidth')
        
        figureLayout.addRow('キャプション:', self.captionEntry)
        figureLayout.addRow('ラベル:', self.labelEntry)
        figureLayout.addRow('位置:', self.positionCombo)
        figureLayout.addRow('幅:', self.widthSpin)
        figureLayout.addRow('高さ:', self.heightSpin)
        
        figureGroup.setLayout(figureLayout)
        
        # プロットタブレイアウトに追加
        plotTabLayout.addWidget(plotTypeGroup)
        plotTabLayout.addWidget(styleGroup)
        plotTabLayout.addWidget(axisGroup)
        plotTabLayout.addWidget(legendGroup)
        plotTabLayout.addWidget(figureGroup)
        plotTab.setLayout(plotTabLayout)
        
        # タブ3: 理論曲線
        theoryTab = QWidget()
        theoryTabLayout = QVBoxLayout()
        
        # 理論曲線追加
        theoryGroup = QGroupBox("理論曲線")
        theoryLayout = QVBoxLayout()
        
        self.theoryCurveCheck = QCheckBox('理論曲線を追加')
        theoryLayout.addWidget(self.theoryCurveCheck)
        
        # 方程式入力
        equationLayout = QHBoxLayout()
        equationLabel = QLabel('方程式:')
        self.equationEntry = QLineEdit('x^2')
        self.equationEntry.setEnabled(False)
        self.equationEntry.setPlaceholderText('例: x^2 or sin(x) or exp(-x/2)')
        equationLayout.addWidget(equationLabel)
        equationLayout.addWidget(self.equationEntry)
        theoryLayout.addLayout(equationLayout)
        
        # 理論曲線パラメータ設定
        paramsGroup = QGroupBox("パラメータ設定")
        paramsLayout = QVBoxLayout()
        
        # ドメイン範囲
        domainLayout = QHBoxLayout()
        domainLabel = QLabel('ドメイン範囲:')
        self.domainMinSpin = QDoubleSpinBox()
        self.domainMinSpin.setRange(-1000, 1000)
        self.domainMinSpin.setValue(0)
        self.domainMinSpin.setEnabled(False)
        self.domainMaxSpin = QDoubleSpinBox()
        self.domainMaxSpin.setRange(-1000, 1000)
        self.domainMaxSpin.setValue(10)
        self.domainMaxSpin.setEnabled(False)
        domainLayout.addWidget(domainLabel)
        domainLayout.addWidget(self.domainMinSpin)
        domainLayout.addWidget(QLabel('〜'))
        domainLayout.addWidget(self.domainMaxSpin)
        paramsLayout.addLayout(domainLayout)
        
        # サンプル数
        samplesLayout = QHBoxLayout()
        samplesLabel = QLabel('サンプル数:')
        self.samplesSpin = QSpinBox()
        self.samplesSpin.setRange(10, 1000)
        self.samplesSpin.setValue(200)
        self.samplesSpin.setEnabled(False)
        samplesLayout.addWidget(samplesLabel)
        samplesLayout.addWidget(self.samplesSpin)
        paramsLayout.addLayout(samplesLayout)
        
        # パラメータスイープ
        self.paramSweepCheck = QCheckBox('パラメータスイープ')
        self.paramSweepCheck.setEnabled(False)
        paramsLayout.addWidget(self.paramSweepCheck)
        
        # パラメータ名
        paramNameLayout = QHBoxLayout()
        paramNameLabel = QLabel('パラメータ名:')
        self.paramNameEntry = QLineEdit('a')
        self.paramNameEntry.setEnabled(False)
        paramNameLayout.addWidget(paramNameLabel)
        paramNameLayout.addWidget(self.paramNameEntry)
        paramsLayout.addLayout(paramNameLayout)
        
        # パラメータ範囲
        paramRangeLayout = QHBoxLayout()
        paramRangeLabel = QLabel('パラメータ範囲:')
        self.paramMinSpin = QDoubleSpinBox()
        self.paramMinSpin.setRange(-1000, 1000)
        self.paramMinSpin.setValue(0.01)
        self.paramMinSpin.setEnabled(False)
        self.paramMaxSpin = QDoubleSpinBox()
        self.paramMaxSpin.setRange(-1000, 1000)
        self.paramMaxSpin.setValue(0.1)
        self.paramMaxSpin.setEnabled(False)
        self.paramStepSpin = QDoubleSpinBox()
        self.paramStepSpin.setRange(0.001, 100)
        self.paramStepSpin.setValue(0.01)
        self.paramStepSpin.setEnabled(False)
        paramRangeLayout.addWidget(paramRangeLabel)
        paramRangeLayout.addWidget(self.paramMinSpin)
        paramRangeLayout.addWidget(QLabel('〜'))
        paramRangeLayout.addWidget(self.paramMaxSpin)
        paramRangeLayout.addWidget(QLabel('ステップ:'))
        paramRangeLayout.addWidget(self.paramStepSpin)
        paramsLayout.addLayout(paramRangeLayout)
        
        # パラメータスイープのための曲線リスト
        self.sweepCurvesList = QListWidget()
        self.sweepCurvesList.setEnabled(False)
        paramsLayout.addWidget(self.sweepCurvesList)
        
        # 追加ボタン
        addSweepButtonLayout = QHBoxLayout()
        addSweepButton = QPushButton('パラメータ値を追加')
        addSweepButton.setEnabled(False)
        addSweepButton.clicked.connect(self.add_param_value)
        removeSweepButton = QPushButton('選択した値を削除')
        removeSweepButton.setEnabled(False)
        removeSweepButton.clicked.connect(self.remove_param_value)
        addSweepButtonLayout.addWidget(addSweepButton)
        addSweepButtonLayout.addWidget(removeSweepButton)
        paramsLayout.addLayout(addSweepButtonLayout)
        
        # イベント接続
        self.theoryCurveCheck.toggled.connect(lambda checked: [
            self.equationEntry.setEnabled(checked),
            self.domainMinSpin.setEnabled(checked),
            self.domainMaxSpin.setEnabled(checked),
            self.samplesSpin.setEnabled(checked),
            self.paramSweepCheck.setEnabled(checked)
        ])
        
        self.paramSweepCheck.toggled.connect(lambda checked: [
            self.paramNameEntry.setEnabled(checked),
            self.paramMinSpin.setEnabled(checked),
            self.paramMaxSpin.setEnabled(checked),
            self.paramStepSpin.setEnabled(checked),
            self.sweepCurvesList.setEnabled(checked),
            addSweepButton.setEnabled(checked),
            removeSweepButton.setEnabled(checked)
        ])
        
        paramsGroup.setLayout(paramsLayout)
        theoryLayout.addWidget(paramsGroup)
        theoryGroup.setLayout(theoryLayout)
        
        # 理論曲線のスタイル設定
        theoryStyleGroup = QGroupBox("理論曲線のスタイル")
        theoryStyleLayout = QGridLayout()
        
        # 色選択
        theoryColorLabel = QLabel('線の色:')
        self.theoryColorButton = QPushButton()
        self.theoryColorButton.setStyleSheet('background-color: green;')
        self.theoryCurrentColor = QColor('green')
        self.theoryColorButton.clicked.connect(self.select_theory_color)
        self.theoryColorButton.setEnabled(False)
        theoryStyleLayout.addWidget(theoryColorLabel, 0, 0)
        theoryStyleLayout.addWidget(self.theoryColorButton, 0, 1)
        
        # 線の太さ
        theoryLineWidthLabel = QLabel('線の太さ:')
        self.theoryLineWidthSpin = QDoubleSpinBox()
        self.theoryLineWidthSpin.setRange(0.1, 5.0)
        self.theoryLineWidthSpin.setSingleStep(0.1)
        self.theoryLineWidthSpin.setValue(1.0)
        self.theoryLineWidthSpin.setEnabled(False)
        theoryStyleLayout.addWidget(theoryLineWidthLabel, 1, 0)
        theoryStyleLayout.addWidget(self.theoryLineWidthSpin, 1, 1)
        
        # 凡例ラベル
        theoryLegendLabel = QLabel('凡例ラベル:')
        self.theoryLegendEntry = QLineEdit('理論曲線')
        self.theoryLegendEntry.setEnabled(False)
        theoryStyleLayout.addWidget(theoryLegendLabel, 2, 0)
        theoryStyleLayout.addWidget(self.theoryLegendEntry, 2, 1)
        
        theoryStyleGroup.setLayout(theoryStyleLayout)
        
        # イベント接続
        self.theoryCurveCheck.toggled.connect(lambda checked: [
            self.theoryColorButton.setEnabled(checked),
            self.theoryLineWidthSpin.setEnabled(checked),
            self.theoryLegendEntry.setEnabled(checked)
        ])
        
        # 理論曲線タブレイアウトに追加
        theoryTabLayout.addWidget(theoryGroup)
        theoryTabLayout.addWidget(theoryStyleGroup)
        theoryTab.setLayout(theoryTabLayout)
        
        # タブ4: 特殊点・注釈
        annotationTab = QWidget()
        annotationTabLayout = QVBoxLayout()
        
        # 特殊点グループ
        specialPointsGroup = QGroupBox("特殊点")
        specialPointsLayout = QVBoxLayout()
        
        self.specialPointsCheck = QCheckBox('特殊点を表示')
        specialPointsLayout.addWidget(self.specialPointsCheck)
        
        # 特殊点テーブル
        self.specialPointsTable = QTableWidget(0, 3)  # 行数, 列数
        self.specialPointsTable.setHorizontalHeaderLabels(['X', 'Y', '色'])
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
        self.annotationsTable = QTableWidget(0, 4)  # 行数, 列数
        self.annotationsTable.setHorizontalHeaderLabels(['X', 'Y', 'テキスト', '位置'])
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
        tabWidget.addTab(theoryTab, "理論曲線")
        tabWidget.addTab(annotationTab, "特殊点・注釈")
        
        # 設定部分のレイアウトに追加
        settingsLayout.addLayout(infoLayout)
        settingsLayout.addWidget(datasetGroup)
        settingsLayout.addWidget(tabWidget)

        # データテーブルの高さを増やす
        self.dataTable.setMinimumHeight(200)
        
        # 変換ボタン
        convertButton = QPushButton('LaTeXコードに変換')
        convertButton.clicked.connect(self.convert_to_tikz)
        convertButton.setStyleSheet('background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;')
        settingsLayout.addWidget(convertButton)
        
        # --- 下部：結果表示部分 ---
        resultWidget = QWidget()
        resultLayout = QVBoxLayout(resultWidget)
        
        resultLabel = QLabel("LaTeX コード:")
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(200)
        
        copyButton = QPushButton("クリップボードにコピー")
        copyButton.clicked.connect(self.copy_to_clipboard)
        
        resultLayout.addWidget(resultLabel)
        resultLayout.addWidget(self.resultText)
        resultLayout.addWidget(copyButton)
        
        # スプリッターに追加
        splitter.addWidget(settingsWidget)
        splitter.addWidget(resultWidget)
        splitter.setSizes([500, 200])  # 初期サイズ比率
        
        # メインレイアウトに追加
        mainLayout.addWidget(splitter)
        
        # ステータスバー
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # メインウィジェット設定
        mainWidget.setLayout(mainLayout)
        self.setCentralWidget(mainWidget)
        
        # データと状態の初期化
        self.data_x = []
        self.data_y = []
        self.special_points = []
        self.annotations = []
        self.param_values = []
        
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
        elif self.excelRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(True)
            self.sheetCombobox.setEnabled(True)
            self.dataTable.setEnabled(False)
        elif self.manualRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(True)
    
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
    
    # 理論曲線の色選択ダイアログ
    def select_theory_color(self):
        color = QColorDialog.getColor(self.theoryCurrentColor, self, "理論曲線の色を選択")
        if color.isValid():
            self.theoryCurrentColor = color
            self.theoryColorButton.setStyleSheet(f'background-color: {color.name()};')
    
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
        color_item = QTableWidgetItem("red")
        
        self.specialPointsTable.setItem(row_position, 0, x_item)
        self.specialPointsTable.setItem(row_position, 1, y_item)
        self.specialPointsTable.setItem(row_position, 2, color_item)
    
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
        pos_item = QTableWidgetItem("north east")
        
        self.annotationsTable.setItem(row_position, 0, x_item)
        self.annotationsTable.setItem(row_position, 1, y_item)
        self.annotationsTable.setItem(row_position, 2, text_item)
        self.annotationsTable.setItem(row_position, 3, pos_item)
    
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

    # TikZコードに変換する
    def convert_to_tikz(self):
        # データセットが空かチェック
        if not self.datasets or all(not dataset.get('data_x') for dataset in self.datasets):
            QMessageBox.warning(self, "警告", "データが読み込まれていません。先にデータを読み込んでください。")
            return
        
        try:
            # 各データセットを処理
            latex_code = self.generate_tikz_code_multi_datasets()
            
            self.resultText.setPlainText(latex_code)
            self.statusBar.showMessage("TikZコードが生成されました")
            
            # コードが生成されたらコピーするように促す
            QMessageBox.information(self, "TikZコード生成完了", 
                                   "TikZコードが生成されました。クリップボードにコピーボタンを押して利用できます。")
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"変換中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            self.statusBar.showMessage("変換エラー")

    # データセット管理の関数
    def add_dataset(self, name_arg=None):
        """新しいデータセットを追加する"""
        try:
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

            dataset = {
                'name': final_name, # Always a string
                'data_x': [],
                'data_y': [],
                'color': QColor('blue'),
                'line_width': 1.0,
                'marker_style': '*',
                'marker_size': 3.0,
                'plot_type': "line",
                'legend_label': final_name, # Always a string, initialized with name
                'show_legend': True,
                'is_theory': False,
                'equation': '',
                'special_points': [],
                'annotations': []
            }
            
            self.datasets.append(dataset)
            self.datasetList.addItem(final_name) # addItem directly with the guaranteed string
            self.datasetList.setCurrentRow(len(self.datasets) - 1) # This will trigger on_dataset_selected
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
        """データセットが選択されたときに呼ばれる"""
        try:
            if row < 0 or row >= len(self.datasets):
                # 選択が無効になった場合（例：最後のアイテム削除直後など）
                # current_dataset_index を無効な値に設定し、UIを適切に処理
                if not self.datasets: # リストが空の場合
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets()
                return
            
            self.current_dataset_index = row
            dataset = self.datasets[row]
            self.update_ui_from_dataset(dataset)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_from_dataset(self, dataset):
        """現在のデータセットに基づいてUIを更新する"""
        try:
            # UIを更新
            # 凡例ラベル
            legend_text = str(dataset.get('legend_label', dataset.get('name', '')))
            self.legendLabel.setText(legend_text)
            
            # 色
            color = dataset.get('color', QColor('blue'))
            self.currentColor = color
            self.colorButton.setStyleSheet(f'background-color: {color.name()};')
            
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
            
            # データテーブルの更新
            if self.manualRadio.isChecked():
                self.update_data_table_from_dataset(dataset)
                
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"UI更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
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
            
            # データを表示
            for i, (x, y) in enumerate(zip(data_x, data_y)):
                self.dataTable.insertRow(i)
                self.dataTable.setItem(i, 0, QTableWidgetItem(str(x)))
                self.dataTable.setItem(i, 1, QTableWidgetItem(str(y)))
                
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データテーブル更新中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_current_dataset(self):
        """UIの値を現在のデータセットに反映する"""
        if self.current_dataset_index < 0 or not self.datasets:
            return
            
        dataset = self.datasets[self.current_dataset_index]
        
        # UIから値を取得してデータセットを更新
        dataset['color'] = self.currentColor
        dataset['line_width'] = self.lineWidthSpin.value()
        dataset['marker_style'] = self.markerCombo.currentText()
        dataset['marker_size'] = self.markerSizeSpin.value()
        dataset['show_legend'] = self.legendCheck.isChecked()
        dataset['legend_label'] = self.legendLabel.text()
        
        # プロットタイプ
        if self.lineRadio.isChecked():
            dataset['plot_type'] = "line"
        elif self.scatterRadio.isChecked():
            dataset['plot_type'] = "scatter"
        elif self.lineScatterRadio.isChecked():
            dataset['plot_type'] = "line_scatter"
        else:
            dataset['plot_type'] = "bar"

    def generate_tikz_code_multi_datasets(self):
        """複数のデータセットを持つTikZコードを生成する"""
        # 結果のLaTeXコード
        latex = []
        
        # figure環境の開始
        latex.append(f"\\begin{{figure}}[{self.positionCombo.currentText()}]")
        latex.append("    \\centering")
        
        # tikzpictureの開始
        latex.append("    \\begin{tikzpicture}")
        
        # axis環境の設定
        axis_options = []
        axis_options.append(f"width={self.widthSpin.value()}\\textwidth")
        axis_options.append(f"height={self.heightSpin.value()}\\textwidth")
        axis_options.append(f"xlabel={{{self.xLabelEntry.text()}}}")
        axis_options.append(f"ylabel={{{self.yLabelEntry.text()}}}")
        
        x_min = self.xMinSpin.value()
        x_max = self.xMaxSpin.value()
        y_min = self.yMinSpin.value()
        y_max = self.yMaxSpin.value()
        
        if x_min != x_max:
            axis_options.append(f"xmin={x_min}, xmax={x_max}")
        if y_min != y_max:
            axis_options.append(f"ymin={y_min}, ymax={y_max}")
        
        if self.gridCheck.isChecked():
            axis_options.append("grid=both")
        
        if self.legendCheck.isChecked():
            axis_options.append(f"legend pos={self.legendPosCombo.currentText()}")
        
        # axis環境の開始
        latex.append(f"        \\begin{{axis}}[")
        latex.append(f"            {','.join(axis_options)}")
        latex.append("        ]")
        
        # 各データセットのプロット処理
        for i, dataset in enumerate(self.datasets):
            if not dataset.get('data_x') or not dataset.get('data_y'):
                continue  # データがないデータセットはスキップ
            
            # データセットの設定を取得
            plot_type = dataset.get('plot_type', 'line')
            color = dataset.get('color', QColor('blue')).name()
            line_width = dataset.get('line_width', 1.0)
            marker_style = dataset.get('marker_style', '*')
            marker_size = dataset.get('marker_size', 3.0)
            show_legend = dataset.get('show_legend', True)
            legend_label = dataset.get('legend_label', dataset.get('name', ''))
            
            # データセットがグラフタイプに応じたプロット処理
            if dataset.get('is_theory', False):
                # 理論曲線の処理
                self.add_theory_curve_to_latex(latex, dataset, i)
            else:
                # 通常のデータセットの処理
                self.add_dataset_to_latex(latex, dataset, i, plot_type, color, line_width, 
                                         marker_style, marker_size, show_legend, legend_label)
            
            # 特殊点の追加
            special_points = dataset.get('special_points', [])
            for x, y, point_color in special_points:
                latex.append(f"        % 特殊点 (データセット{i+1}: {dataset.get('name', '')})")
                latex.append(f"        \\addplot[only marks, mark=*, {point_color}] coordinates {{")
                latex.append(f"            ({x}, {y})")
                latex.append("        };")
            
            # 注釈の追加
            annotations = dataset.get('annotations', [])
            for x, y, text, pos in annotations:
                latex.append(f"        % 注釈 (データセット{i+1}: {dataset.get('name', '')})")
                latex.append(f"        \\node at (axis cs:{x},{y}) [anchor={pos}, font=\\small] {{{text}}};")
        
        # 理論曲線の追加 (パラメータスイープによる複数曲線)
        show_theory = self.theoryCurveCheck.isChecked()
        if show_theory and self.paramSweepCheck.isChecked():
            equation = self.equationEntry.text()
            domain_min = self.domainMinSpin.value()
            domain_max = self.domainMaxSpin.value()
            samples = self.samplesSpin.value()
            theory_color = self.theoryCurrentColor.name()
            theory_line_width = self.theoryLineWidthSpin.value()
            theory_legend = self.theoryLegendEntry.text()
            
            # パラメータスイープ設定の取得
            param_name = self.paramNameEntry.text()
            param_values = []
            for i in range(self.sweepCurvesList.count()):
                item_text = self.sweepCurvesList.item(i).text()
                param_val = float(item_text.split('=')[1].strip())
                param_values.append((param_name, param_val))
            
            # 理論曲線のオプション
            theory_options = []
            theory_options.append(f"domain={domain_min}:{domain_max}")
            theory_options.append(f"samples={samples}")
            theory_options.append("smooth")
            theory_options.append("thick")
            theory_options.append(theory_color)
            theory_options.append(f"line width={theory_line_width}pt")
            
            # 色のリスト（理論曲線ごとに色を変える）
            colors = ["red", "green", "blue", "orange", "purple", "cyan", "magenta", "brown", "gray", "darkgray"]
            
            # パラメータ値ごとに理論曲線を追加
            for i, (param_name, param_val) in enumerate(param_values):
                # 色の設定
                theory_options_copy = theory_options.copy()
                if i > 0:  # 最初の曲線以外は色を変える
                    # 色の部分を置き換え
                    for j, opt in enumerate(theory_options_copy):
                        if opt.startswith("rgb") or QColor(opt).isValid():
                            theory_options_copy[j] = colors[i % len(colors)]
                            break
                
                eq_with_param = equation.replace(param_name, str(param_val))
                
                latex.append(f"        % 理論曲線 ({param_name}={param_val})")
                latex.append(f"        \\addplot[{', '.join(theory_options_copy)}] {{")
                latex.append(f"            {eq_with_param}")
                latex.append("        };")
                
                # 凡例エントリを別途追加
                if self.legendCheck.isChecked():
                    latex.append(f"        \\addlegendentry{{{theory_legend} ($\\{param_name}={param_val}$)}}")
        
        # axis環境の終了
        latex.append("        \\end{axis}")
        
        # tikzpictureの終了
        latex.append("    \\end{tikzpicture}")
        
        # キャプションとラベル
        latex.append(f"    \\caption{{{self.captionEntry.text()}}}")
        latex.append(f"    \\label{{{self.labelEntry.text()}}}")
        
        # figure環境の終了
        latex.append("\\end{figure}")
        
        return '\n'.join(latex)
    
    def add_dataset_to_latex(self, latex, dataset, index, plot_type, color, line_width, 
                             marker_style, marker_size, show_legend, legend_label):
        """LaTeXコードにデータセットを追加する"""
        # データポイントのフォーマット
        coordinates = []
        for x, y in zip(dataset.get('data_x', []), dataset.get('data_y', [])):
            coordinates.append(f"({x}, {y})")
        
        if not coordinates:
            return
        
        # 線グラフまたは線と点の組み合わせの場合の線プロット
        if plot_type == "line" or plot_type == "line_scatter":
            # 線プロット
            plot_options = []
            plot_options.append(color)
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
            scatter_options.append(color)
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
            bar_options.append(color)
            bar_options.append(f"fill={color}")
            
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
            color_item = self.specialPointsTable.item(row, 2)
            
            if x_item and y_item and color_item:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    color_val = color_item.text()
                    special_points.append((x_val, y_val, color_val))
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
            pos_item = self.annotationsTable.item(row, 3)
            
            if x_item and y_item and text_item and pos_item:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    text_val = text_item.text()
                    pos_val = pos_item.text()
                    annotations.append((x_val, y_val, text_val, pos_val))
                except ValueError:
                    pass
        
        dataset = self.datasets[self.current_dataset_index]
        dataset['annotations'] = annotations
        
        QMessageBox.information(self, "成功", 
                              f"データセット '{dataset['name']}' に {len(annotations)} 個の注釈を割り当てました")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TikZPlotConverter()
    ex.show()
    sys.exit(app.exec_())