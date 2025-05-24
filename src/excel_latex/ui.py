import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
                             QPushButton, QFileDialog, QComboBox, QMessageBox, QTextEdit,
                             QCheckBox, QGridLayout, QGroupBox, QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
from pandas import ExcelFile
from excel_latex.processor import ExcelToLatexProcessor


class ExcelToLatexTab(QWidget):
    def __init__(self):
        super().__init__()
        self.statusBar = None
        self.processor = ExcelToLatexProcessor()
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
            latex_code = self.processor.excel_to_latex_universal(
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