import sys
import os
import webbrowser
import numpy as np
import pandas as pd
import math  # 数学関数を使用するためにインポート
from pandas import ExcelFile
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
from PyQt5.QtGui import QColor, QPixmap, QFont
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import re


class TexcellentConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('TexcellentConverter - Excel & TikZ to LaTeX Converter')
        # 元のtizk.pyと同じウィンドウ設定
        screen = QApplication.primaryScreen().geometry()
        max_width = int(screen.width() * 0.95)
        max_height = int(screen.height() * 0.95)
        initial_height = int(screen.height() * 0.9)
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
        self.tikzTab = TikZPlotTab()
        self.infoTab = InfoTab()
        
        # タブをタブウィジェットに追加
        self.tabWidget.addTab(self.excelTab, "Excel → LaTeX 表")
        self.tabWidget.addTab(self.tikzTab, "TikZ グラフ")
        self.tabWidget.addTab(self.infoTab, "情報・ライセンス")
        
        # メインレイアウトに追加
        mainLayout.addWidget(self.tabWidget)
        
        # ステータスバー
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # タブからステータスバーにアクセスできるように設定
        self.excelTab.set_status_bar(self.statusBar)
        self.tikzTab.set_status_bar(self.statusBar)
        
        # メインウィジェット設定
        mainWidget.setLayout(mainLayout)
        self.setCentralWidget(mainWidget)


class ExcelToLatexTab(QWidget):
    def __init__(self):
        super().__init__()
        self.statusBar = None
        self.initUI()

    def set_status_bar(self, status_bar):
        self.statusBar = status_bar

    def initUI(self):
        mainLayout = QVBoxLayout()
        
        # 上下に分割するスプリッター
        splitter = QSplitter(Qt.Vertical)
        
        # --- 上部：設定部分 ---
        settingsWidget = QWidget()
        settingsLayout = QVBoxLayout()
        settingsWidget.setLayout(settingsLayout)
        
        # 注意書き用のレイアウト
        infoLayout = QVBoxLayout()
        
        # 注意書き用のラベル - 赤色テキスト
        infoLabel = QLabel("注意: この表を使用するには、LaTeXドキュメントのプリアンブルに以下を追加してください:")
        infoLabel.setStyleSheet("color: #cc0000; font-weight: bold;") # 赤色のテキスト、太字
        
        # パッケージ名と複コピーボタンの水平レイアウト
        packageLayout = QHBoxLayout()
        
        # パッケージ名用のラベル - 普通の色
        packageLabel = QLabel(r"\usepackage{multirow}")
        packageLabel.setStyleSheet("font-family: monospace;") # 等幅フォント
        
        # パッケージ名をコピーするボタン
        copyPackageButton = QPushButton("コピー")
        copyPackageButton.setFixedWidth(60)
        def copy_package():
            clipboard = QApplication.clipboard()
            clipboard.setText(r"\usepackage{multirow}")
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
        
        # Excelファイル選択
        fileLayout = QHBoxLayout()
        fileLabel = QLabel('Excelファイル:')
        self.fileEntry = QLineEdit()
        browseButton = QPushButton('参照...')
        browseButton.clicked.connect(self.browse_excel_file)
        fileLayout.addWidget(fileLabel)
        fileLayout.addWidget(self.fileEntry)
        fileLayout.addWidget(browseButton)
        
        # シート選択
        sheetLayout = QHBoxLayout()
        sheetLabel = QLabel('シート名:')
        self.sheetCombobox = QComboBox()
        sheetLayout.addWidget(sheetLabel)
        sheetLayout.addWidget(self.sheetCombobox)
        
        # セル範囲
        rangeLayout = QHBoxLayout()
        rangeLabel = QLabel('セル範囲:')
        self.rangeEntry = QLineEdit()
        self.rangeEntry.setPlaceholderText('例: A1:E6')  # プレースホルダーテキスト追加

        # ヘルプボタンを追加
        rangeHelpButton = QPushButton('?')
        rangeHelpButton.setFixedSize(20, 20)
        rangeHelpButton.setStyleSheet('border-radius: 10px; background-color: #007AFF; color: white;')
        rangeHelpButton.clicked.connect(lambda: QMessageBox.information(
            self, "セル範囲の指定方法", 
            "表にしたい範囲の左上のセルと右下のセルを「A1:E6」のように指定してください。\n\n"
            "・A1: 左上のセル\n・E6: 右下のセル\n・コロン(:)で区切ります（「:」は小文字）\n\n"
            "結合セルを含む場合は、その結合セル全体が範囲内に入るように選択してください。\n"
            "（結合セルを一部だけ選択すると、正しく値が表示されない場合があります）"
        ))

        rangeLayout.addWidget(rangeLabel)
        rangeLayout.addWidget(self.rangeEntry)
        rangeLayout.addWidget(rangeHelpButton)
        
        # オプション
        optionsGroup = QGroupBox("オプション")
        optionsLayout = QGridLayout()
        
        captionLabel = QLabel('表のキャプション:')
        self.captionEntry = QLineEdit('表のタイトル')
        optionsLayout.addWidget(captionLabel, 0, 0)
        optionsLayout.addWidget(self.captionEntry, 0, 1)
        
        labelLabel = QLabel('表のラベル:')
        self.labelEntry = QLineEdit('tab:data')
        optionsLayout.addWidget(labelLabel, 1, 0)
        optionsLayout.addWidget(self.labelEntry, 1, 1)
        
        positionLabel = QLabel('表の位置:')
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        optionsLayout.addWidget(positionLabel, 2, 0)
        optionsLayout.addWidget(self.positionCombo, 2, 1)
        
        self.showValueCheck = QCheckBox('数式の代わりに値を表示')
        self.showValueCheck.setChecked(True)
        optionsLayout.addWidget(self.showValueCheck, 3, 0)
        
        self.addBordersCheck = QCheckBox('罫線を追加')
        self.addBordersCheck.setChecked(True)
        optionsLayout.addWidget(self.addBordersCheck, 3, 1)
        
        optionsGroup.setLayout(optionsLayout)
        
        # 変換ボタン
        convertButton = QPushButton('LaTeXに変換')
        convertButton.clicked.connect(self.convert_to_latex)
        convertButton.setStyleSheet('background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;')
        
        # 設定部分のレイアウトに追加
        settingsLayout.addLayout(infoLayout)
        settingsLayout.addLayout(fileLayout)
        settingsLayout.addLayout(sheetLayout)
        settingsLayout.addLayout(rangeLayout)
        settingsLayout.addWidget(optionsGroup)
        settingsLayout.addWidget(convertButton)
        
        # --- 下部：結果表示部分 ---
        resultWidget = QWidget()
        resultLayout = QVBoxLayout()
        resultWidget.setLayout(resultLayout)
        
        resultLabel = QLabel("LaTeX コード:")
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(200)
        
        copyButton = QPushButton("クリップボードにコピー")
        copyButton.clicked.connect(self.copy_to_clipboard)
        copyButton.setStyleSheet("background-color: #007BFF; color: white; font-size: 16px; padding: 12px; font-weight: 900; border-radius: 8px; text-transform: uppercase; letter-spacing: 1px;")
        
        resultLayout.addWidget(resultLabel)
        resultLayout.addWidget(self.resultText)
        resultLayout.addWidget(copyButton)
        
        # スプリッターに追加
        splitter.addWidget(settingsWidget)
        splitter.addWidget(resultWidget)
        splitter.setSizes([300, 300])  # 初期サイズ比率
        
        # メインレイアウトに追加
        mainLayout.addWidget(splitter)
        self.setLayout(mainLayout)

    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.fileEntry.setText(file_path)
            self.update_sheet_names(file_path)

    def update_sheet_names(self, file_path):
        try:
            xls = ExcelFile(file_path)
            self.sheetCombobox.clear()
            self.sheetCombobox.addItems(xls.sheet_names)
            if self.statusBar:
                self.statusBar.showMessage(f"ファイル '{os.path.basename(file_path)}' を読み込みました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"シート名の取得に失敗しました: {str(e)}")
            if self.statusBar:
                self.statusBar.showMessage("ファイル読み込みエラー")

    def convert_to_latex(self):
        excel_file = self.fileEntry.text()
        if not excel_file or not os.path.exists(excel_file):
            QMessageBox.critical(self, "エラー", "有効なExcelファイルを選択してください")
            return

        # 選択データの回収
        sheet_name = self.sheetCombobox.currentText()
        cell_range = self.rangeEntry.text()
        caption = self.captionEntry.text()
        label = self.labelEntry.text()
        position = self.positionCombo.currentText()
        show_value = self.showValueCheck.isChecked()
        add_borders = self.addBordersCheck.isChecked()

        try:
            latex_code = self.excel_to_latex_universal(
                excel_file, sheet_name, cell_range, caption, label, position,
                show_value, add_borders
            )
            if latex_code:  # エラー発生時は空文字列が返る
                # 注意書きコメントを削除（GUIに表示しているので）
                latex_lines = latex_code.split('\n')
                filtered_lines = [line for line in latex_lines if not (line.startswith('% 注意:') or line.startswith('% \\usepackage'))]
                # 空行が連続している場合は1つにする
                if filtered_lines and filtered_lines[0] == '':
                    filtered_lines.pop(0)
                
                self.resultText.setPlainText('\n'.join(filtered_lines))
                if self.statusBar:
                    self.statusBar.showMessage("LaTeXコードが生成されました")

        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"変換中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
            if self.statusBar:
                self.statusBar.showMessage("変換エラー")

    def copy_to_clipboard(self):
        latex_code = self.resultText.toPlainText()
        if latex_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(latex_code)
            if self.statusBar:
                self.statusBar.showMessage("LaTeXコードをクリップボードにコピーしました", 3000)  # 3秒間表示

    def excel_to_latex_universal(self, excel_file, sheet_name, cell_range, caption, label,
                                position, show_value, add_borders=True):
        try:
            wb = load_workbook(excel_file, data_only=show_value)
            ws = wb[sheet_name]
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"Excelファイルの読み込みに失敗しました: {e}")
            return ""

        # ここに実際の変換ロジックが必要
        latex = ["% LaTeX code placeholder"]
        return "\n".join(latex)

    def add_dataset(self, name_arg=None):
        """データセットを追加"""
        if name_arg:
            name = name_arg
        else:
            name = f"Dataset {len(self.datasets) + 1}"
        
        # データセット作成
        dataset = {
            'name': name,
            'data_x': [1, 2, 3, 4, 5],
            'data_y': [1, 4, 9, 16, 25]
        }
        
        self.datasets.append(dataset)
        
        if self.statusBar:
            self.statusBar.showMessage(f"データセット '{name}' を追加しました")


class InfoTab(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # アプリケーション情報
        titleLabel = QLabel("TexcellentConverter")
        titleLabel.setFont(QFont("Arial", 24, QFont.Bold))
        titleLabel.setAlignment(Qt.AlignCenter)
        titleLabel.setStyleSheet("color: #2C3E50; margin: 20px;")
        
        versionLabel = QLabel("Version 1.0.0")
        versionLabel.setFont(QFont("Arial", 12))
        versionLabel.setAlignment(Qt.AlignCenter)
        versionLabel.setStyleSheet("color: #7F8C8D; margin-bottom: 30px;")
        
        descriptionLabel = QLabel("Excel表とデータを美しいLaTeX形式に変換するアプリケーション")
        descriptionLabel.setFont(QFont("Arial", 14))
        descriptionLabel.setAlignment(Qt.AlignCenter)
        descriptionLabel.setWordWrap(True)
        descriptionLabel.setStyleSheet("color: #34495E; margin: 20px; padding: 10px;")
        
        # ボタンエリア
        buttonLayout = QVBoxLayout()
        
        # GitHubリポジトリボタン
        githubButton = QPushButton("📂 GitHubリポジトリを開く")
        githubButton.setFont(QFont("Arial", 12))
        githubButton.setStyleSheet("""
            QPushButton {
                background-color: #24292e;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #444d56;
            }
        """)
        githubButton.clicked.connect(lambda: webbrowser.open("https://github.com/yourusername/texcellentconverter"))
        
        # ライセンスボタン
        licenseButton = QPushButton("📄 ライセンス情報")
        licenseButton.setFont(QFont("Arial", 12))
        licenseButton.setStyleSheet("""
            QPushButton {
                background-color: #0366d6;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0256cc;
            }
        """)
        licenseButton.clicked.connect(self.show_license_info)
        
        # 使い方ボタン
        helpButton = QPushButton("❓ 使い方")
        helpButton.setFont(QFont("Arial", 12))
        helpButton.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        helpButton.clicked.connect(self.show_help_info)
        
        buttonLayout.addWidget(githubButton)
        buttonLayout.addWidget(licenseButton)
        buttonLayout.addWidget(helpButton)
        buttonLayout.setSpacing(15)
        
        # ライセンス情報テキスト
        licenseText = QLabel("""
<b>ライセンス:</b> MIT License<br><br>
<b>作成者:</b> Your Name<br><br>
<b>使用ライブラリ:</b><br>
• PyQt5 - GUI フレームワーク<br>
• openpyxl - Excel ファイル処理<br>
• pandas - データ処理<br>
• numpy - 数値計算<br><br>
<b>連絡先:</b> your.email@example.com
        """)
        licenseText.setAlignment(Qt.AlignLeft)
        licenseText.setWordWrap(True)
        licenseText.setStyleSheet("""
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 20px;
            color: #495057;
            line-height: 1.5;
        """)
        
        # レイアウト構成
        layout.addWidget(titleLabel)
        layout.addWidget(versionLabel)
        layout.addWidget(descriptionLabel)
        layout.addLayout(buttonLayout)
        layout.addWidget(licenseText)
        layout.addStretch()
        
        self.setLayout(layout)

    def show_license_info(self):
        QMessageBox.information(self, "ライセンス情報", 
                               "MIT License\n\n"
                               "Copyright (c) 2024 Your Name\n\n"
                               "Permission is hereby granted, free of charge, to any person obtaining a copy "
                               "of this software and associated documentation files (the \"Software\"), to deal "
                               "in the Software without restriction, including without limitation the rights "
                               "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell "
                               "copies of the Software, and to permit persons to whom the Software is "
                               "furnished to do so, subject to the following conditions:\n\n"
                               "The above copyright notice and this permission notice shall be included in all "
                               "copies or substantial portions of the Software.")

    def show_help_info(self):
        QMessageBox.information(self, "使い方", 
                               "【Excel → LaTeX 表】\n"
                               "1. Excelファイルを選択\n"
                               "2. シート名を選択\n"
                               "3. セル範囲を指定（例: A1:E6）\n"
                               "4. オプションを設定\n"
                               "5. 「LaTeXに変換」をクリック\n"
                               "6. 生成されたコードをコピー\n\n"
                               "【TikZ グラフ】\n"
                               "1. データソースを選択（CSV/Excel/手動入力）\n"
                               "2. ファイルを選択またはデータを手動入力\n"
                               "3. データセットを追加・管理\n"
                               "4. プロット設定（色、マーカー、線の太さなど）\n"
                               "5. 軸設定（タイトル、ラベル、範囲）\n"
                               "6. 特殊ポイント・アノテーション・理論曲線の追加\n"
                               "7. 「TikZコードを生成」をクリック\n"
                               "8. 生成されたコードをコピー")


class TikZPlotTab(QWidget):
    def __init__(self):
        super().__init__()
        self.statusBar = None
        self.datasets = []
        self.current_dataset_index = -1
        
        # 初期化
        self.special_points = []
        self.annotations = []
        self.param_values = []
        
        self.initUI()

    def set_status_bar(self, status_bar):
        self.statusBar = status_bar

    def initUI(self):
        # スクロールエリアで全体をラップ
        scrollArea = QScrollArea()
        scrollArea.setWidgetResizable(True)
        scrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # メインウィジェット
        mainWidget = QWidget()
        mainLayout = QVBoxLayout(mainWidget)
        
        # スクロールのヒント表示
        scrollHintLabel = QLabel("画面サイズを変更可能・縦スクロール可能です")
        scrollHintLabel.setAlignment(Qt.AlignCenter)
        scrollHintLabel.setStyleSheet("color: #666; font-style: italic; margin: 5px;")
        mainLayout.addWidget(scrollHintLabel)
        
        # パッケージ情報
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
            if self.statusBar:
                self.statusBar.showMessage("パッケージ名をコピーしました", 3000)
        copyPackageButton.clicked.connect(copy_package)
        
        packageLayout.addWidget(packageLabel)
        packageLayout.addWidget(copyPackageButton)
        packageLayout.addStretch()
        
        infoLayout.addWidget(infoLabel)
        infoLayout.addLayout(packageLayout)
        mainLayout.addLayout(infoLayout)
        
        # データセット管理セクション
        datasetGroup = QGroupBox("データセット管理")
        datasetLayout = QVBoxLayout()
        
        self.datasetList = QListWidget()
        self.datasetList.currentRowChanged.connect(self.on_dataset_selected)
        self.datasetList.setMinimumHeight(100)
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
        mainLayout.addWidget(datasetGroup)
        
        # タブウィジェット作成
        tabWidget = QTabWidget()
        
        # データ入力タブを作成
        self.create_data_input_tab(tabWidget)
        
        # グラフ設定タブを作成（全体設定も含む）
        self.create_graph_settings_tab(tabWidget)
        
        # 特殊点・注釈設定タブを作成
        self.create_annotation_tab(tabWidget)
        
        mainLayout.addWidget(tabWidget)
        
        # LaTeXコードに変換ボタン（タブの外側）
        convertButton = QPushButton("LaTeXコードに変換")
        convertButton.clicked.connect(self.convert_to_tikz)
        convertButton.setStyleSheet('background-color: #4CAF50; color: white; font-size: 14px; padding: 10px;')
        mainLayout.addWidget(convertButton)
        
        # 結果表示エリア
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(200)
        mainLayout.addWidget(QLabel("生成されたTikZコード:"))
        mainLayout.addWidget(self.resultText)
        
        # コピーボタン
        copyButton = QPushButton("クリップボードにコピー")
        copyButton.clicked.connect(self.copy_to_clipboard)
        copyButton.setStyleSheet("background-color: #007BFF; color: white; font-size: 16px; padding: 12px;")
        mainLayout.addWidget(copyButton)
        
        # スクロールエリアにメインウィジェットを設定
        scrollArea.setWidget(mainWidget)
        
        # 最終レイアウト
        finalLayout = QVBoxLayout()
        finalLayout.addWidget(scrollArea)
        self.setLayout(finalLayout)
        
        # 初期データセット追加
        self.add_dataset("Dataset 1")

    def create_data_input_tab(self, tabWidget):
        """データ入力タブを作成"""
        dataTab = QWidget()
        dataTabLayout = QVBoxLayout()
        
        # データソースタイプの選択
        dataSourceTypeGroup = QGroupBox("データソースタイプ")
        dataSourceTypeLayout = QVBoxLayout()
        
        self.measuredRadio = QRadioButton("実測値データ（CSV/Excel/手入力）")
        self.formulaRadio = QRadioButton("数式によるグラフ生成")
        self.measuredRadio.setChecked(True)
        
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
        
        # 実測値データコンテナ
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
        
        # シート選択
        sheetLayout = QHBoxLayout()
        sheetLabel = QLabel('シート名:')
        sheetLabel.setIndent(30)
        self.sheetCombobox = QComboBox()
        self.sheetCombobox.setEnabled(False)
        sheetLayout.addWidget(sheetLabel)
        sheetLayout.addWidget(self.sheetCombobox)
        
        # データ直接入力
        self.manualRadio = QRadioButton("データを直接入力:")
        self.manualRadio.setChecked(True)
        
        # ボタングループ設定
        sourceGroup = QButtonGroup(self)
        sourceGroup.addButton(self.csvRadio)
        sourceGroup.addButton(self.excelRadio)
        sourceGroup.addButton(self.manualRadio)
        sourceGroup.buttonClicked.connect(self.toggle_source_fields)
        
        # データテーブル
        self.dataTable = QTableWidget(10, 2)
        self.dataTable.setHorizontalHeaderLabels(['X', 'Y'])
        self.dataTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dataTable.setEnabled(True)
        
        # テーブル操作ボタン
        tableButtonLayout = QHBoxLayout()
        addRowButton = QPushButton('行を追加')
        addRowButton.clicked.connect(self.add_table_row)
        removeRowButton = QPushButton('選択行を削除')
        removeRowButton.clicked.connect(self.remove_table_row)
        tableButtonLayout.addWidget(addRowButton)
        tableButtonLayout.addWidget(removeRowButton)
        
        dataSourceLayout.addLayout(csvLayout)
        dataSourceLayout.addLayout(excelLayout)
        dataSourceLayout.addLayout(sheetLayout)
        dataSourceLayout.addWidget(self.manualRadio)
        dataSourceLayout.addWidget(self.dataTable)
        dataSourceLayout.addLayout(tableButtonLayout)
        dataSourceGroup.setLayout(dataSourceLayout)
        
        # セル範囲指定
        columnGroup = QGroupBox("データ選択")
        columnLayout = QGridLayout()
        
        usageGuideLabel = QLabel("【データ選択ガイド】\n"
                              "■ CSVファイル/Excelファイル: セル範囲で直接指定します\n"
                              "■ セル範囲の例: X軸「A2:A10」Y軸「B2:B10」\n"
                              "■ A1形式のセル指定で、列と行の範囲を指定してください")
        usageGuideLabel.setStyleSheet(
            "background-color: #546e7a; color: #ffffff; padding: 10px; "
            "border: 2px solid #90a4ae; border-radius: 6px; font-weight: bold; margin: 5px;"
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
        
        columnHelpLabel = QLabel('※ セル範囲はA1形式で指定してください\n'
                               '※ 例: 同じ行なら「A2:E2」と「A3:E3」、同じ列なら「A2:A10」と「B2:B10」')
        columnHelpLabel.setStyleSheet('color: gray; font-style: italic;')
        columnHelpLabel.setWordWrap(True)
        columnLayout.addWidget(columnHelpLabel, 4, 0, 1, 3)
        
        columnGroup.setLayout(columnLayout)
        
        # データ確定ボタン
        dataActionGroup = QGroupBox("データ確定")
        dataActionLayout = QVBoxLayout()
        
        actionNoteLabel = QLabel("※ CSVファイル、Excelファイル、手入力のいずれの方法でも、データを保存するには下のボタンを押してください")
        actionNoteLabel.setStyleSheet("color: #cc0000; font-weight: bold;")
        actionNoteLabel.setWordWrap(True)
        dataActionLayout.addWidget(actionNoteLabel)
        
        self.loadDataButton = QPushButton('データを確定・保存')
        self.loadDataButton.setToolTip('入力したデータを現在のデータセットに保存します')
        self.loadDataButton.clicked.connect(self.load_data)
        self.loadDataButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        dataActionLayout.addWidget(self.loadDataButton)
        
        dataActionGroup.setLayout(dataActionLayout)
        
        measuredLayout.addWidget(dataSourceGroup)
        measuredLayout.addWidget(columnGroup)
        measuredLayout.addWidget(dataActionGroup)
        
        # 数式データコンテナ
        self.formulaContainer = QWidget()
        formulaLayout = QVBoxLayout(self.formulaContainer)
        
        # 数式入力
        formulaFormGroup = QGroupBox("数式入力")
        formulaFormLayout = QVBoxLayout()
        
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
            '※ 掛け算は必ず * を明示してください（例: 2*x, (x+1)*y）<br>'
            '数式内では以下などの関数と演算が使用可能です:<br>'
            'sin, cos, tan, exp, ln, log, sqrt, ^（累乗）, +, -, *, /'
        )
        formulaInfoLabel.setStyleSheet(
            'background-color: #fffbe6; color: black; border: 1px solid #ffe082; border-radius: 4px; padding: 8px;'
        )
        formulaFormLayout.addWidget(formulaInfoLabel)
        
        formulaFormGroup.setLayout(formulaFormLayout)
        formulaLayout.addWidget(formulaFormGroup)
        
        # 数式適用ボタン
        applyFormulaButton = QPushButton('数式を適用してデータ生成')
        applyFormulaButton.clicked.connect(self.apply_formula)
        applyFormulaButton.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; padding: 8px;')
        formulaLayout.addWidget(applyFormulaButton)
        
        # コンテナの初期表示設定
        self.formulaContainer.setVisible(False)
        
        dataTabLayout.addWidget(self.measuredContainer)
        dataTabLayout.addWidget(self.formulaContainer)
        
        dataTab.setLayout(dataTabLayout)
        tabWidget.addTab(dataTab, "データ入力")

    def create_graph_settings_tab(self, tabWidget):
        """グラフ設定タブを作成（元のtizk.pyと同じ構造）"""
        plotTab = QWidget()
        plotTabLayout = QVBoxLayout()
        
        # グラフ全体の設定ヘッダー
        plotTabLayout.addWidget(QLabel("【グラフ全体の設定】"))
        
        # 軸設定
        axisGroup = QGroupBox("グラフ全体設定 - 軸")
        axisLayout = QGridLayout()
        
        # X軸ラベル
        xLabelLabel = QLabel('X軸ラベル:')
        self.xLabelEntry = QLineEdit('X')
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
        self.yLabelEntry = QLineEdit('Y')
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
        
        # 凡例設定
        self.legendCheck = QCheckBox('凡例を表示')
        self.legendCheck.setChecked(True)
        
        legendPosLabel = QLabel('凡例の位置:')
        self.legendPosCombo = QComboBox()
        self.legendPosCombo.addItems(['左上', '右上', '左下', '右下'])
        
        axisLayout.addWidget(self.legendCheck, 5, 0, 1, 2)
        axisLayout.addWidget(legendPosLabel, 6, 0)
        axisLayout.addWidget(self.legendPosCombo, 6, 1)
        
        axisGroup.setLayout(axisLayout)
        plotTabLayout.addWidget(axisGroup)
        
        # 図の設定
        figureGroup = QGroupBox("グラフ全体設定 - 図")
        figureLayout = QFormLayout()
        
        # キャプション
        self.captionEntry = QLineEdit('グラフのキャプション')
        
        # ラベル
        self.labelEntry = QLineEdit('fig:graph')
        
        # 位置
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        
        # 幅と高さ
        self.widthSpin = QDoubleSpinBox()
        self.widthSpin.setRange(0.1, 1.0)
        self.widthSpin.setSingleStep(0.1)
        self.widthSpin.setValue(0.8)
        self.widthSpin.setSuffix('\\textwidth')
        
        self.heightSpin = QDoubleSpinBox()
        self.heightSpin.setRange(0.1, 1.0)
        self.heightSpin.setSingleStep(0.1)
        self.heightSpin.setValue(0.6)
        self.heightSpin.setSuffix('\\textwidth')
        
        figureLayout.addRow('キャプション:', self.captionEntry)
        figureLayout.addRow('ラベル:', self.labelEntry)
        figureLayout.addRow('位置:', self.positionCombo)
        figureLayout.addRow('幅:', self.widthSpin)
        figureLayout.addRow('高さ:', self.heightSpin)
        
        figureGroup.setLayout(figureLayout)
        plotTabLayout.addWidget(figureGroup)
        
        # 区切り線
        from PyQt5.QtWidgets import QFrame
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        plotTabLayout.addWidget(separator)
        
        # データセット個別の設定ヘッダー
        plotTabLayout.addWidget(QLabel("【データセット個別の設定】"))
        
        # プロットタイプグループ
        plotTypeGroup = QGroupBox("データセット個別設定 - プロットタイプ")
        plotTypeLayout = QVBoxLayout()
        
        self.lineRadio = QRadioButton('線グラフ')
        self.scatterRadio = QRadioButton('散布図')
        self.lineScatterRadio = QRadioButton('線＋散布図')
        self.barRadio = QRadioButton('棒グラフ')
        self.scatterRadio.setChecked(True)  # デフォルト
        
        plotTypeLayout.addWidget(self.lineRadio)
        plotTypeLayout.addWidget(self.scatterRadio)
        plotTypeLayout.addWidget(self.lineScatterRadio)
        plotTypeLayout.addWidget(self.barRadio)
        
        plotTypeGroup.setLayout(plotTypeLayout)
        plotTabLayout.addWidget(plotTypeGroup)
        
        # スタイル設定
        styleGroup = QGroupBox("データセット個別設定 - スタイル")
        styleLayout = QGridLayout()
        
        # 色選択
        colorLabel = QLabel('線/点の色:')
        self.colorButton = QPushButton()
        self.colorButton.setStyleSheet('background-color: blue;')
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
        
        # データ点表示オプション
        self.showDataPointsCheck = QCheckBox('データ点をマークで表示（線グラフでも点を表示）')
        self.showDataPointsCheck.setChecked(False)
        styleLayout.addWidget(self.showDataPointsCheck, 5, 0, 1, 2)
        
        styleGroup.setLayout(styleLayout)
        plotTabLayout.addWidget(styleGroup)
        
        # 凡例ラベル（データセット個別）
        legendLabelGroup = QGroupBox("データセット個別設定 - 凡例ラベル")
        legendLabelLayout = QFormLayout()
        
        self.legendLabel = QLineEdit('データ')
        legendLabelLayout.addRow('凡例ラベル:', self.legendLabel)
        
        legendLabelGroup.setLayout(legendLabelLayout)
        plotTabLayout.addWidget(legendLabelGroup)
        
        plotTab.setLayout(plotTabLayout)
        tabWidget.addTab(plotTab, "グラフ設定")

    def create_annotation_tab(self, tabWidget):
        """特殊点・注釈設定タブを作成"""
        annotationTab = QWidget()
        annotationTabLayout = QVBoxLayout()
        
        # 特殊点グループ
        specialPointsGroup = QGroupBox("特殊点")
        specialPointsLayout = QVBoxLayout()
        
        self.specialPointsCheck = QCheckBox('特殊点を表示')
        specialPointsLayout.addWidget(self.specialPointsCheck)
        
        # 特殊点テーブル
        self.specialPointsTable = QTableWidget(0, 4)
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
        self.annotationsTable = QTableWidget(0, 5)
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
        
        annotationTabLayout.addWidget(specialPointsGroup)
        annotationTabLayout.addWidget(annotationsGroup)
        annotationTab.setLayout(annotationTabLayout)
        
        tabWidget.addTab(annotationTab, "特殊点・注釈設定")

    def browse_csv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "CSVファイルを選択", "", "CSV Files (*.csv)")
        if file_path:
            self.fileEntry.setText(file_path)
            self.update_column_names(file_path)

    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excelEntry.setText(file_path)
            self.update_sheet_names(file_path)

    def update_sheet_names(self, file_path):
        try:
            xls = ExcelFile(file_path)
            self.sheetCombobox.clear()
            self.sheetCombobox.addItems(xls.sheet_names)
            if self.statusBar:
                self.statusBar.showMessage(f"ファイル '{os.path.basename(file_path)}' を読み込みました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"シート名の取得に失敗しました: {str(e)}")
            if self.statusBar:
                self.statusBar.showMessage("ファイル読み込みエラー")

    def toggle_source_fields(self, button=None):
        if self.csvRadio.isChecked():
            self.fileEntry.setEnabled(True)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(False)
            self.xRangeEntry.setEnabled(True)
            self.yRangeEntry.setEnabled(True)
            self.loadDataButton.setToolTip("CSVファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.excelRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(True)
            self.sheetCombobox.setEnabled(True)
            self.dataTable.setEnabled(False)
            self.xRangeEntry.setEnabled(True)
            self.yRangeEntry.setEnabled(True)
            self.loadDataButton.setToolTip("Excelファイルからデータを読み込み、現在のデータセットに保存します")
        elif self.manualRadio.isChecked():
            self.fileEntry.setEnabled(False)
            self.excelEntry.setEnabled(False)
            self.sheetCombobox.setEnabled(False)
            self.dataTable.setEnabled(True)
            self.xRangeEntry.setEnabled(False)
            self.yRangeEntry.setEnabled(False)
            self.loadDataButton.setToolTip("テーブルに入力したデータを現在のデータセットに保存します")

    def on_data_source_type_changed(self, checked=None):
        """データソースタイプ変更時の処理"""
        if self.measuredRadio.isChecked():
            self.measuredContainer.setVisible(True)
            self.formulaContainer.setVisible(False)
        else:
            self.measuredContainer.setVisible(False)
            self.formulaContainer.setVisible(True)

    def update_column_names(self, file_path):
        """CSVの列名情報を更新"""
        try:
            df = pd.read_csv(file_path, nrows=1)
            columns = list(df.columns)
            if self.statusBar:
                self.statusBar.showMessage(f"CSV列: {', '.join(map(str, columns))}")
        except Exception as e:
            if self.statusBar:
                self.statusBar.showMessage(f"CSV読み込みエラー: {e}")

    def load_data(self):
        """選択されたデータソースからデータを読み込む"""
        try:
            if self.current_dataset_index < 0:
                QMessageBox.warning(self, "警告", "データを読み込むデータセットを選択してください")
                return
                
            data_x = []
            data_y = []
            
            # セル範囲の取得
            x_range = self.xRangeEntry.text().strip()
            y_range = self.yRangeEntry.text().strip()
            
            if self.csvRadio.isChecked():
                # CSVファイル選択時のチェックと処理
                file_path = self.fileEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なCSVファイルを選択してください")
                    return
                
                # セル範囲のチェック
                if not x_range or not y_range:
                    QMessageBox.warning(self, "警告", "X軸とY軸のセル範囲を指定してください")
                    return
                
                try:
                    # まずUTF-8で試す
                    df = pd.read_csv(file_path, encoding='utf-8', sep=None, engine='python')
                except UnicodeDecodeError:
                    try:
                        # 次にShift-JISで試す
                        df = pd.read_csv(file_path, encoding='shift_jis', sep=None, engine='python')
                    except UnicodeDecodeError:
                        QMessageBox.warning(self, "エンコーディングエラー", 
                                          "CSVファイルのエンコーディングを自動判別できませんでした。\n"
                                          "UTF-8またはShift-JISで保存し直してください。")
                        return
                except Exception as e:
                    QMessageBox.warning(self, "CSVファイル読み込みエラー", 
                                      f"CSVファイルの読み込み中にエラーが発生しました: {str(e)}")
                    return
                
                try:
                    # A1:A10形式のセル範囲を解析してデータを取得
                    data_x, data_y = self.extract_data_from_range(df, x_range, y_range)
                except Exception as e:
                    QMessageBox.warning(self, "セル範囲エラー", f"データ抽出中にエラーが発生しました: {str(e)}")
                    return
                
            elif self.excelRadio.isChecked():
                # Excelファイル選択時のチェックと処理
                file_path = self.excelEntry.text()
                if not file_path or not os.path.exists(file_path):
                    QMessageBox.warning(self, "警告", "有効なExcelファイルを選択してください")
                    return
                
                sheet_name = self.sheetCombobox.currentText()
                if not sheet_name:
                    QMessageBox.warning(self, "警告", "シート名を選択してください")
                    return
                
                # セル範囲のチェック
                if not x_range or not y_range:
                    QMessageBox.warning(self, "警告", "X軸とY軸のセル範囲を指定してください")
                    return
                
                try:
                    data_x, data_y = self.extract_data_from_excel_range(file_path, sheet_name, x_range, y_range)
                except Exception as e:
                    QMessageBox.warning(self, "警告", f"Excelデータ抽出中にエラーが発生しました: {str(e)}")
                    return
                
            elif self.manualRadio.isChecked():
                # 手入力データの場合
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
            
            # データの長さチェック
            if len(data_x) != len(data_y):
                QMessageBox.warning(self, "警告", f"X軸とY軸のデータ長が一致しません (X: {len(data_x)}, Y: {len(data_y)})")
                return
                
            if not data_x:
                QMessageBox.warning(self, "警告", "有効なデータがありません")
                return
                
            # 現在のデータセットにデータを設定
            self.datasets[self.current_dataset_index]['data_x'] = data_x
            self.datasets[self.current_dataset_index]['data_y'] = data_y
            
            dataset_name = self.datasets[self.current_dataset_index]['name']
            QMessageBox.information(self, "読み込み成功 ✓", f"データセット '{dataset_name}' に{len(data_x)}個のデータポイントを正常に読み込みました")
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset_name}' にデータを読み込みました: {len(data_x)}ポイント")
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"データ読み込み中にエラーが発生しました: {str(e)}")

    def extract_data_from_range(self, df, x_range, y_range):
        """データフレームからセル範囲を使ってデータを抽出"""
        def parse_range(range_str):
            # A1:A10 のような形式を解析
            if ':' in range_str:
                start, end = range_str.split(':')
                start_col = ''.join(filter(str.isalpha, start))
                start_row = int(''.join(filter(str.isdigit, start))) - 1  # 0-based index
                end_col = ''.join(filter(str.isalpha, end))
                end_row = int(''.join(filter(str.isdigit, end))) - 1
                
                def col_to_index(col_str):
                    """列文字を数値インデックスに変換 (A=0, B=1, ...)"""
                    result = 0
                    for char in col_str:
                        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
                    return result - 1
                
                col_idx = col_to_index(start_col)
                return df.iloc[start_row:end_row+1, col_idx].tolist()
            else:
                # 単一の列指定の場合
                if range_str.isdigit():
                    col_idx = int(range_str)
                    return df.iloc[:, col_idx].tolist()
                else:
                    def col_to_index(col_str):
                        result = 0
                        for char in col_str:
                            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
                        return result - 1
                    col_idx = col_to_index(range_str)
                    return df.iloc[:, col_idx].tolist()
        
        try:
            x_data = parse_range(x_range)
            y_data = parse_range(y_range)
            
            # 数値に変換
            def safe_float_conversion(value):
                if pd.isna(value):
                    return None
                try:
                    return float(value)
                except (ValueError, TypeError):
                    return None
            
            x_data = [safe_float_conversion(x) for x in x_data]
            y_data = [safe_float_conversion(y) for y in y_data]
            
            # Noneを除去
            valid_pairs = [(x, y) for x, y in zip(x_data, y_data) if x is not None and y is not None]
            if valid_pairs:
                x_data, y_data = zip(*valid_pairs)
                return list(x_data), list(y_data)
            else:
                return [], []
                
        except Exception as e:
            raise Exception(f"セル範囲の解析に失敗しました: {e}")

    def extract_data_from_excel_range(self, file_path, sheet_name, x_range, y_range):
        """Excelの範囲からデータを抽出"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            return self.extract_data_from_range(df, x_range, y_range)
        except Exception as e:
            raise Exception(f"Excel範囲の解析に失敗しました: {e}")

    def add_table_row(self):
        """テーブルに行を追加"""
        current_rows = self.dataTable.rowCount()
        self.dataTable.insertRow(current_rows)

    def remove_table_row(self):
        """テーブルから行を削除"""
        current_row = self.dataTable.currentRow()
        if current_row >= 0:
            self.dataTable.removeRow(current_row)

    def apply_formula(self):
        """数式からデータを生成"""
        if self.current_dataset_index < 0:
            QMessageBox.warning(self, "警告", "データを生成するデータセットを選択してください")
            return
        
        try:
            equation = self.equationEntry.text().strip()
            if not equation:
                QMessageBox.warning(self, "警告", "数式を入力してください")
                return
            
            x_min = self.domainMinSpin.value()
            x_max = self.domainMaxSpin.value()
            samples = self.samplesSpin.value()
            
            if x_min >= x_max:
                QMessageBox.warning(self, "警告", "x軸の範囲が正しくありません")
                return
            
            # x値の配列を生成
            x_values = np.linspace(x_min, x_max, samples)
            y_values = []
            
            # 数式を評価
            for x in x_values:
                try:
                    # 簡単な数式評価（安全でない場合があるので本番では注意）
                    y = eval(equation.replace('^', '**'))
                    y_values.append(y)
                except Exception:
                    y_values.append(0)
            
            # データセットに保存
            self.datasets[self.current_dataset_index]['data_x'] = x_values.tolist()
            self.datasets[self.current_dataset_index]['data_y'] = y_values
            
            dataset_name = self.datasets[self.current_dataset_index]['name']
            QMessageBox.information(self, "生成成功 ✓", f"データセット '{dataset_name}' に数式データを生成しました")
            if self.statusBar:
                self.statusBar.showMessage(f"数式データを生成しました: {equation}")
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"数式データ生成中にエラーが発生しました: {str(e)}")

    def add_dataset(self, name_arg=None):
        """データセットを追加"""
        if name_arg:
            name = name_arg
        else:
            name = f"Dataset {len(self.datasets) + 1}"
        
        # データセット作成
        dataset = {
            'name': name,
            'data_x': [1, 2, 3, 4, 5],
            'data_y': [1, 4, 9, 16, 25]
        }
        
        self.datasets.append(dataset)
        self.datasetList.addItem(name)
        self.datasetList.setCurrentRow(len(self.datasets) - 1)
        
        if self.statusBar:
            self.statusBar.showMessage(f"データセット '{name}' を追加しました")

    def remove_dataset(self):
        """選択されたデータセットを削除"""
        current_row = self.datasetList.currentRow()
        if current_row >= 0:
            dataset_name = self.datasets[current_row]['name']
            self.datasets.pop(current_row)
            self.datasetList.takeItem(current_row)
            self.current_dataset_index = -1
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました")

    def rename_dataset(self):
        """データセット名を変更"""
        current_row = self.datasetList.currentRow()
        if current_row >= 0:
            old_name = self.datasets[current_row]['name']
            new_name, ok = QInputDialog.getText(self, 'データセット名変更', '新しい名前を入力してください:', text=old_name)
            if ok and new_name.strip():
                self.datasets[current_row]['name'] = new_name
                self.datasetList.item(current_row).setText(new_name)
                if self.statusBar:
                    self.statusBar.showMessage(f"データセット名を '{old_name}' から '{new_name}' に変更しました")

    def on_dataset_selected(self, row):
        """データセットが選択された時の処理"""
        if row >= 0:
            self.current_dataset_index = row
            self.update_ui_from_dataset(self.datasets[row])

    def update_ui_from_dataset(self, dataset):
        """データセットの内容でUIを更新"""
        try:
            self.block_signals_temporarily(True)
            
            # UI要素が存在するかチェックしてから更新
            if hasattr(self, 'plot_type_combo'):
                self.plot_type_combo.setCurrentText(dataset.get('plot_type', 'scatter'))
            if hasattr(self, 'line_width_spin'):
                self.line_width_spin.setValue(dataset.get('line_width', 1.0))
            if hasattr(self, 'marker_combo'):
                self.marker_combo.setCurrentText(dataset.get('marker_style', '*'))
            if hasattr(self, 'marker_size_spin'):
                self.marker_size_spin.setValue(dataset.get('marker_size', 3))
            if hasattr(self, 'legend_check'):
                self.legend_check.setChecked(dataset.get('show_legend', True))
            if hasattr(self, 'legend_entry'):
                self.legend_entry.setText(dataset.get('legend_label', dataset['name']))
            
            # 色ボタンの更新
            if hasattr(self, 'color_btn'):
                color = dataset.get('color', 'blue')
                if isinstance(color, str):
                    if color.startswith('#'):
                        self.color_btn.setStyleSheet(f"background-color: {color}")
                    else:
                        self.color_btn.setStyleSheet(f"background-color: {color}")
            
            self.block_signals_temporarily(False)
        except Exception as e:
            print(f"Error updating UI from dataset: {e}")

    def block_signals_temporarily(self, block):
        """シグナルを一時的にブロック"""
        widgets = []
        if hasattr(self, 'plot_type_combo'):
            widgets.append(self.plot_type_combo)
        if hasattr(self, 'line_width_spin'):
            widgets.append(self.line_width_spin)
        if hasattr(self, 'marker_combo'):
            widgets.append(self.marker_combo)
        if hasattr(self, 'marker_size_spin'):
            widgets.append(self.marker_size_spin)
        if hasattr(self, 'legend_check'):
            widgets.append(self.legend_check)
        if hasattr(self, 'legend_entry'):
            widgets.append(self.legend_entry)
        
        for widget in widgets:
            widget.blockSignals(block)

    def update_current_dataset(self):
        """現在のデータセットを更新"""
        if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
            dataset = self.datasets[self.current_dataset_index]
            if hasattr(self, 'plot_type_combo'):
                dataset['plot_type'] = self.plot_type_combo.currentText()
            if hasattr(self, 'line_width_spin'):
                dataset['line_width'] = self.line_width_spin.value()
            if hasattr(self, 'marker_combo'):
                dataset['marker_style'] = self.marker_combo.currentText()
            if hasattr(self, 'marker_size_spin'):
                dataset['marker_size'] = self.marker_size_spin.value()
            if hasattr(self, 'legend_check'):
                dataset['show_legend'] = self.legend_check.isChecked()
            if hasattr(self, 'legend_entry'):
                dataset['legend_label'] = self.legend_entry.text()

    def select_color(self):
        """色を選択"""
        color = QColorDialog.getColor()
        if color.isValid():
            if hasattr(self, 'color_btn'):
                self.color_btn.setStyleSheet(f"background-color: {color.name()}")
            if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
                self.datasets[self.current_dataset_index]['color'] = color.name()

    def select_tangent_color(self):
        """理論曲線の色を選択"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.theory_color_btn.setStyleSheet(f"background-color: {color.name()}")

    def add_special_point(self):
        """特殊ポイントを追加"""
        x, ok1 = QInputDialog.getDouble(self, '特殊ポイント', 'X座標を入力:')
        if not ok1:
            return
        y, ok2 = QInputDialog.getDouble(self, '特殊ポイント', 'Y座標を入力:')
        if not ok2:
            return
        label, ok3 = QInputDialog.getText(self, '特殊ポイント', 'ラベルを入力:')
        if not ok3:
            return
        
        special_point = {'x': x, 'y': y, 'label': label}
        self.special_points.append(special_point)
        self.special_points_list.addItem(f"({x}, {y}): {label}")

    def remove_special_point(self):
        """特殊ポイントを削除"""
        current_row = self.special_points_list.currentRow()
        if current_row >= 0:
            self.special_points.pop(current_row)
            self.special_points_list.takeItem(current_row)

    def add_annotation(self):
        """アノテーションを追加"""
        x, ok1 = QInputDialog.getDouble(self, 'アノテーション', 'X座標を入力:')
        if not ok1:
            return
        y, ok2 = QInputDialog.getDouble(self, 'アノテーション', 'Y座標を入力:')
        if not ok2:
            return
        text, ok3 = QInputDialog.getText(self, 'アノテーション', 'テキストを入力:')
        if not ok3:
            return
        
        annotation = {'x': x, 'y': y, 'text': text}
        self.annotations.append(annotation)
        self.annotations_list.addItem(f"({x}, {y}): {text}")

    def remove_annotation(self):
        """アノテーションを削除"""
        current_row = self.annotations_list.currentRow()
        if current_row >= 0:
            self.annotations.pop(current_row)
            self.annotations_list.takeItem(current_row)

    def add_param_value(self):
        """パラメータ値を追加"""
        name, ok1 = QInputDialog.getText(self, 'パラメータ', 'パラメータ名を入力:')
        if not ok1 or not name.strip():
            return
        value, ok2 = QInputDialog.getDouble(self, 'パラメータ', f'{name}の値を入力:')
        if not ok2:
            return
        
        param = {'name': name, 'value': value}
        self.param_values.append(param)
        self.param_list.addItem(f"{name} = {value}")

    def remove_param_value(self):
        """パラメータ値を削除"""
        current_row = self.param_list.currentRow()
        if current_row >= 0:
            self.param_values.pop(current_row)
            self.param_list.takeItem(current_row)

    def insert_into_equation(self, text):
        """数式に挿入"""
        current_cursor_pos = self.theory_formula_entry.cursorPosition()
        current_text = self.theory_formula_entry.text()
        new_text = current_text[:current_cursor_pos] + text + current_text[current_cursor_pos:]
        self.theory_formula_entry.setText(new_text)
        self.theory_formula_entry.setCursorPosition(current_cursor_pos + len(text))

    def show_tikz_function_help(self):
        """TikZ関数ヘルプを表示"""
        help_text = """
TikZ/pgfplotsで使用可能な数学関数:

基本関数:
• sin(x), cos(x), tan(x) - 三角関数
• asin(x), acos(x), atan(x) - 逆三角関数
• exp(x) - 指数関数 (e^x)
• ln(x) - 自然対数
• log10(x) - 常用対数
• sqrt(x) - 平方根
• abs(x) - 絶対値
• floor(x), ceil(x) - 床関数、天井関数

演算子:
• + - * / - 四則演算
• ^ - べき乗
• () - カッコ

定数:
• pi - 円周率 π
• e - ネイピア数

例:
• sin(x)
• x^2 + 2*x + 1
• exp(-x^2)
• sqrt(x^2 + 1)
        """
        QMessageBox.information(self, "TikZ数学関数ヘルプ", help_text)

    def convert_to_tikz(self):
        """TikZコードを生成"""
        if not self.datasets:
            QMessageBox.warning(self, "警告", "データセットがありません。先にデータセットを追加してください。")
            return
        
        try:
            tikz_code = self.generate_tikz_code_multi_datasets()
            self.result_text.setPlainText(tikz_code)
            if self.statusBar:
                self.statusBar.showMessage("TikZコードが生成されました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"TikZコード生成中にエラーが発生しました: {str(e)}")
            if self.statusBar:
                self.statusBar.showMessage("コード生成エラー")

    def generate_tikz_code_multi_datasets(self):
        """複数のデータセットからTikZコードを生成"""
        latex = []
        
        # begin{tikzpicture}
        latex.append("\\begin{tikzpicture}")
        
        # figure size設定
        size_options = {
            "small": "width=6cm,height=4cm",
            "normal": "width=8cm,height=6cm", 
            "large": "width=12cm,height=8cm"
        }
        size_setting = size_options.get(self.figure_size_combo.currentText(), size_options["normal"])
        
        latex.append("\\begin{axis}[")
        
        # 軸設定
        title = self.title_entry.text() or "My Plot"
        xlabel = self.xlabel_entry.text() or "X"
        ylabel = self.ylabel_entry.text() or "Y"
        
        latex.append(f"    {size_setting},")
        latex.append(f"    title={{{title}}},")
        latex.append(f"    xlabel={{{xlabel}}},")
        latex.append(f"    ylabel={{{ylabel}}},")
        
        if self.grid_check.isChecked():
            latex.append("    grid=major,")
        
        legend_pos = self.legend_pos_combo.currentText()
        latex.append(f"    legend pos={legend_pos},")
        
        # 軸範囲設定
        if self.xmin_entry.text() and self.xmax_entry.text():
            latex.append(f"    xmin={self.xmin_entry.text()}, xmax={self.xmax_entry.text()},")
        if self.ymin_entry.text() and self.ymax_entry.text():
            latex.append(f"    ymin={self.ymin_entry.text()}, ymax={self.ymax_entry.text()},")
        
        latex.append("]")
        
        # 各データセットをプロット
        for i, dataset in enumerate(self.datasets):
            self.add_dataset_to_latex(latex, dataset, i)
        
        # 理論曲線
        if self.theory_check.isChecked() and self.theory_formula_entry.text():
            self.add_theory_curve_to_latex(latex, None, 0)
        
        # 特殊ポイント
        for point in self.special_points:
            latex.append(f"\\node[circle,fill=red,inner sep=2pt] at (axis cs:{point['x']},{point['y']}) {{}};")
            latex.append(f"\\node[above right] at (axis cs:{point['x']},{point['y']}) {{{point['label']}}};")
        
        # アノテーション
        for ann in self.annotations:
            latex.append(f"\\node at (axis cs:{ann['x']},{ann['y']}) {{{ann['text']}}};")
        
        latex.append("\\end{axis}")
        latex.append("\\end{tikzpicture}")
        
        return "\n".join(latex)

    def add_dataset_to_latex(self, latex, dataset, index):
        """データセットをLaTeXコードに追加"""
        # 座標データの生成
        coordinates = []
        for x, y in zip(dataset['x_data'], dataset['y_data']):
            coordinates.append(f"({x},{y})")
        coords_str = " ".join(coordinates)
        
        # プロットオプション
        options = []
        
        if dataset.get('color'):
            color = dataset['color']
            if isinstance(color, QColor):
                color = self.color_to_tikz_rgb(color)
            options.append(f"color={color}")
        
        if dataset.get('plot_type') == 'line':
            pass  # デフォルトで線が引かれる
        elif dataset.get('plot_type') == 'scatter':
            options.append("only marks")
        elif dataset.get('plot_type') == 'both':
            pass  # 線とマーカー両方
        
        if dataset.get('marker_style'):
            marker = dataset['marker_style']
            options.append(f"mark={marker}")
        
        if dataset.get('marker_size'):
            size = dataset['marker_size']
            options.append(f"mark size={size}pt")
        
        if dataset.get('line_width'):
            width = dataset['line_width']
            options.append(f"line width={width}pt")
        
        options_str = ", ".join(options) if options else ""
        
        # addplotコマンド
        if options_str:
            latex.append(f"\\addplot[{options_str}] coordinates {{{coords_str}}};")
        else:
            latex.append(f"\\addplot coordinates {{{coords_str}}};")
        
        # 凡例
        if dataset.get('show_legend', True):
            legend_label = dataset.get('legend_label', dataset['name'])
            latex.append(f"\\addlegendentry{{{legend_label}}}")

    def add_theory_curve_to_latex(self, latex, dataset, index):
        """理論曲線をLaTeXコードに追加"""
        formula = self.theory_formula_entry.text()
        range_min = float(self.theory_range_min.text())
        range_max = float(self.theory_range_max.text())
        samples = self.theory_samples.value()
        
        # 色設定
        color_style = self.theory_color_btn.styleSheet()
        if "background-color:" in color_style:
            color = color_style.split("background-color:")[1].split(";")[0].strip()
        else:
            color = "red"
        
        options = [f"color={color}", f"samples={samples}", f"domain={range_min}:{range_max}"]
        options_str = ", ".join(options)
        
        # パラメータ置換
        formula_with_params = formula
        for param in self.param_values:
            formula_with_params = formula_with_params.replace(param['name'], str(param['value']))
        
        latex.append(f"\\addplot[{options_str}] {{{formula_with_params}}};")
        latex.append("\\addlegendentry{Theory}")

    def copy_to_clipboard(self):
        """結果をクリップボードにコピー"""
        tikz_code = self.result_text.toPlainText()
        if tikz_code:
            clipboard = QApplication.clipboard()
            clipboard.setText(tikz_code)
            if self.statusBar:
                self.statusBar.showMessage("TikZコードをクリップボードにコピーしました", 3000)


class InfoTab(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # アプリケーション情報
        titleLabel = QLabel("TexcellentConverter")
        titleLabel.setFont(QFont("Arial", 24, QFont.Bold))
        titleLabel.setAlignment(Qt.AlignCenter)
        titleLabel.setStyleSheet("color: #2C3E50; margin: 20px;")
        
        versionLabel = QLabel("Version 1.0.0")
        versionLabel.setFont(QFont("Arial", 12))
        versionLabel.setAlignment(Qt.AlignCenter)
        versionLabel.setStyleSheet("color: #7F8C8D; margin-bottom: 30px;")
        
        descriptionLabel = QLabel("Excel表とデータを美しいLaTeX形式に変換するアプリケーション")
        descriptionLabel.setFont(QFont("Arial", 14))
        descriptionLabel.setAlignment(Qt.AlignCenter)
        descriptionLabel.setWordWrap(True)
        descriptionLabel.setStyleSheet("color: #34495E; margin: 20px; padding: 10px;")
        
        # ボタンエリア
        buttonLayout = QVBoxLayout()
        
        # GitHubリポジトリボタン
        githubButton = QPushButton("📂 GitHubリポジトリを開く")
        githubButton.setFont(QFont("Arial", 12))
        githubButton.setStyleSheet("""
            QPushButton {
                background-color: #24292e;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #444d56;
            }
        """)
        githubButton.clicked.connect(lambda: webbrowser.open("https://github.com/yourusername/texcellentconverter"))
        
        # ライセンスボタン
        licenseButton = QPushButton("📄 ライセンス情報")
        licenseButton.setFont(QFont("Arial", 12))
        licenseButton.setStyleSheet("""
            QPushButton {
                background-color: #0366d6;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0256cc;
            }
        """)
        licenseButton.clicked.connect(self.show_license_info)
        
        # 使い方ボタン
        helpButton = QPushButton("❓ 使い方")
        helpButton.setFont(QFont("Arial", 12))
        helpButton.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        helpButton.clicked.connect(self.show_help_info)
        
        buttonLayout.addWidget(githubButton)
        buttonLayout.addWidget(licenseButton)
        buttonLayout.addWidget(helpButton)
        buttonLayout.setSpacing(15)
        
        # ライセンス情報テキスト
        licenseText = QLabel("""
<b>ライセンス:</b> MIT License<br><br>
<b>作成者:</b> Your Name<br><br>
<b>使用ライブラリ:</b><br>
• PyQt5 - GUI フレームワーク<br>
• openpyxl - Excel ファイル処理<br>
• pandas - データ処理<br>
• numpy - 数値計算<br><br>
<b>連絡先:</b> your.email@example.com
        """)
        licenseText.setAlignment(Qt.AlignLeft)
        licenseText.setWordWrap(True)
        licenseText.setStyleSheet("""
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 20px;
            color: #495057;
            line-height: 1.5;
        """)
        
        # レイアウト構成
        layout.addWidget(titleLabel)
        layout.addWidget(versionLabel)
        layout.addWidget(descriptionLabel)
        layout.addLayout(buttonLayout)
        layout.addWidget(licenseText)
        layout.addStretch()
        
        self.setLayout(layout)

    def show_license_info(self):
        QMessageBox.information(self, "ライセンス情報", 
                               "MIT License\n\n"
                               "Copyright (c) 2024 Your Name\n\n"
                               "Permission is hereby granted, free of charge, to any person obtaining a copy "
                               "of this software and associated documentation files (the \"Software\"), to deal "
                               "in the Software without restriction, including without limitation the rights "
                               "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell "
                               "copies of the Software, and to permit persons to whom the Software is "
                               "furnished to do so, subject to the following conditions:\n\n"
                               "The above copyright notice and this permission notice shall be included in all "
                               "copies or substantial portions of the Software.")

    def show_help_info(self):
        QMessageBox.information(self, "使い方", 
                               "【Excel → LaTeX 表】\n"
                               "1. Excelファイルを選択\n"
                               "2. シート名を選択\n"
                               "3. セル範囲を指定（例: A1:E6）\n"
                               "4. オプションを設定\n"
                               "5. 「LaTeXに変換」をクリック\n"
                               "6. 生成されたコードをコピー\n\n"
                               "【TikZ グラフ】\n"
                               "1. データソースを選択（CSV/Excel/手動入力）\n"
                               "2. ファイルを選択またはデータを手動入力\n"
                               "3. データセットを追加・管理\n"
                               "4. プロット設定（色、マーカー、線の太さなど）\n"
                               "5. 軸設定（タイトル、ラベル、範囲）\n"
                               "6. 特殊ポイント・アノテーション・理論曲線の追加\n"
                               "7. 「TikZコードを生成」をクリック\n"
                               "8. 生成されたコードをコピー")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TexcellentConverterApp()
    ex.show()
    sys.exit(app.exec_()) 