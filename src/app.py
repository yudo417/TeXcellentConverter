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
                    if self.statusBar:
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
                if self.statusBar:
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
            }
            
            self.datasets.append(dataset)
            self.datasetList.addItem(final_name) # addItem directly with the guaranteed string
            
            # 新しく追加したデータセットを選択（これがon_dataset_selectedを呼び出す）
            self.datasetList.setCurrentRow(len(self.datasets) - 1)
            if self.statusBar:
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
                    self.update_ui_for_no_datasets() 
                    self.add_dataset("データセット1") # Or add a new default one
                
                if self.statusBar:
                    self.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット削除中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_for_no_datasets(self):
        """データセットがない場合にUIをリセット/クリアする"""
        # 例: 関連する入力フィールドをクリアまたは無効化
        if hasattr(self, 'legendLabel'):
            self.legendLabel.setText("")
        # 他のUI要素も必要に応じてリセット
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
                    if hasattr(self, 'legendLabel'):
                        self.legendLabel.setText(actual_new_name) # UIも即時更新
                
                item = self.datasetList.item(current_row)
                if item:
                    item.setText(actual_new_name)
                if self.statusBar:
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
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset['name']}' からUIを更新しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_current_dataset(self):
        try:
            if self.current_dataset_index < 0 or not self.datasets or self.current_dataset_index >= len(self.datasets):
                return
            
            # 現在のデータセットを取得
            dataset = self.datasets[self.current_dataset_index]
            
            # 色の設定（1回のみ）
            if not isinstance(self.currentColor, QColor):
                self.currentColor = QColor(self.currentColor)
            # QColorオブジェクトのコピーを作成して代入
            dataset['color'] = QColor(self.currentColor)
            
            # 基本設定の更新
            if hasattr(self, 'lineWidthSpin'):
                dataset['line_width'] = self.lineWidthSpin.value()
            if hasattr(self, 'markerCombo'):
                dataset['marker_style'] = self.markerCombo.currentText()
            if hasattr(self, 'markerSizeSpin'):
                dataset['marker_size'] = self.markerSizeSpin.value()
            if hasattr(self, 'legendCheck'):
                dataset['show_legend'] = self.legendCheck.isChecked()
            if hasattr(self, 'legendLabel'):
                dataset['legend_label'] = self.legendLabel.text()
            
            # プロットタイプの更新
            if hasattr(self, 'lineRadio') and self.lineRadio.isChecked():
                dataset['plot_type'] = 'line'
            elif hasattr(self, 'scatterRadio') and self.scatterRadio.isChecked():
                dataset['plot_type'] = 'scatter'
            elif hasattr(self, 'lineScatterRadio') and self.lineScatterRadio.isChecked():
                dataset['plot_type'] = 'line+scatter'
            elif hasattr(self, 'barRadio') and self.barRadio.isChecked():
                dataset['plot_type'] = 'bar'
            
            # データソースタイプの更新
            if hasattr(self, 'measuredRadio') and self.measuredRadio.isChecked():
                dataset['data_source_type'] = 'measured'
            elif hasattr(self, 'formulaRadio') and self.formulaRadio.isChecked():
                dataset['data_source_type'] = 'formula'
            
            # 数式関連の設定
            if hasattr(self, 'equationEntry'):
                dataset['equation'] = self.equationEntry.text()
            if hasattr(self, 'domainMinSpin'):
                dataset['domain_min'] = self.domainMinSpin.value()
            if hasattr(self, 'domainMaxSpin'):
                dataset['domain_max'] = self.domainMaxSpin.value()
            if hasattr(self, 'samplesSpin'):
                dataset['samples'] = self.samplesSpin.value()
            
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset['name']}' が更新されました", 3000)
                
        except Exception as e:
            print(f"Error updating current dataset: {e}")
    
    def update_ui_from_dataset(self, dataset):
        """データセットの内容でUIを更新"""
        try:
            self.block_signals_temporarily(True)
            
            # 色の更新
            color = dataset.get('color', QColor('blue'))
            if isinstance(color, QColor):
                self.currentColor = QColor(color)
                if hasattr(self, 'colorButton'):
                    self.colorButton.setStyleSheet(f"background-color: {color.name()}")
            
            # 基本設定の更新
            if hasattr(self, 'lineWidthSpin'):
                self.lineWidthSpin.setValue(dataset.get('line_width', 1.0))
            if hasattr(self, 'markerCombo'):
                self.markerCombo.setCurrentText(dataset.get('marker_style', '*'))
            if hasattr(self, 'markerSizeSpin'):
                self.markerSizeSpin.setValue(dataset.get('marker_size', 2.0))
            if hasattr(self, 'legendCheck'):
                self.legendCheck.setChecked(dataset.get('show_legend', True))
            if hasattr(self, 'legendLabel'):
                self.legendLabel.setText(dataset.get('legend_label', dataset['name']))
            
            # プロットタイプの更新
            plot_type = dataset.get('plot_type', 'line')
            if hasattr(self, 'lineRadio'):
                self.lineRadio.setChecked(plot_type == 'line')
            if hasattr(self, 'scatterRadio'):
                self.scatterRadio.setChecked(plot_type == 'scatter')
            if hasattr(self, 'lineScatterRadio'):
                self.lineScatterRadio.setChecked(plot_type == 'line+scatter')
            if hasattr(self, 'barRadio'):
                self.barRadio.setChecked(plot_type == 'bar')
            
            # データソースタイプの更新
            data_source_type = dataset.get('data_source_type', 'measured')
            if hasattr(self, 'measuredRadio'):
                self.measuredRadio.setChecked(data_source_type == 'measured')
            if hasattr(self, 'formulaRadio'):
                self.formulaRadio.setChecked(data_source_type == 'formula')
            
            # 数式関連の設定
            if hasattr(self, 'equationEntry'):
                self.equationEntry.setText(dataset.get('equation', ''))
            if hasattr(self, 'domainMinSpin'):
                self.domainMinSpin.setValue(dataset.get('domain_min', 0))
            if hasattr(self, 'domainMaxSpin'):
                self.domainMaxSpin.setValue(dataset.get('domain_max', 10))
            if hasattr(self, 'samplesSpin'):
                self.samplesSpin.setValue(dataset.get('samples', 200))
            
            # データテーブルの更新
            self.update_data_table_from_dataset(dataset)
            
            # データソースタイプに基づくUIの更新
            self.update_ui_based_on_data_source_type()
            
            self.block_signals_temporarily(False)
        except Exception as e:
            print(f"Error updating UI from dataset: {e}")
    
    def block_signals_temporarily(self, block):
        """シグナルを一時的にブロック"""
        widgets = []
        if hasattr(self, 'lineWidthSpin'):
            widgets.append(self.lineWidthSpin)
        if hasattr(self, 'markerCombo'):
            widgets.append(self.markerCombo)
        if hasattr(self, 'markerSizeSpin'):
            widgets.append(self.markerSizeSpin)
        if hasattr(self, 'legendCheck'):
            widgets.append(self.legendCheck)
        if hasattr(self, 'legendLabel'):
            widgets.append(self.legendLabel)
        if hasattr(self, 'lineRadio'):
            widgets.append(self.lineRadio)
        if hasattr(self, 'scatterRadio'):
            widgets.append(self.scatterRadio)
        if hasattr(self, 'lineScatterRadio'):
            widgets.append(self.lineScatterRadio)
        if hasattr(self, 'barRadio'):
            widgets.append(self.barRadio)
        
        for widget in widgets:
            widget.blockSignals(block)
    
    def update_data_table_from_dataset(self, dataset):
        """データセットからデータテーブルを更新"""
        try:
            if hasattr(self, 'dataTable'):
                # テーブルをクリア
                self.dataTable.setRowCount(0)
                
                # データがある場合はテーブルに設定
                data_x = dataset.get('data_x', [])
                data_y = dataset.get('data_y', [])
                
                if data_x and data_y:
                    max_rows = max(len(data_x), len(data_y))
                    self.dataTable.setRowCount(max_rows)
                    
                    for i in range(max_rows):
                        # X値の設定
                        if i < len(data_x):
                            x_item = QTableWidgetItem(str(data_x[i]))
                            self.dataTable.setItem(i, 0, x_item)
                        
                        # Y値の設定
                        if i < len(data_y):
                            y_item = QTableWidgetItem(str(data_y[i]))
                            self.dataTable.setItem(i, 1, y_item)
                else:
                    # データがない場合は10行の空テーブルを作成
                    self.dataTable.setRowCount(10)
        except Exception as e:
            print(f"Error updating data table: {e}")
    
    def update_ui_based_on_data_source_type(self):
        """データソースタイプに基づいてUIの表示を更新"""
        try:
            if hasattr(self, 'measuredRadio') and hasattr(self, 'formulaRadio'):
                if self.measuredRadio.isChecked():
                    if hasattr(self, 'measuredContainer'):
                        self.measuredContainer.setVisible(True)
                    if hasattr(self, 'formulaContainer'):
                        self.formulaContainer.setVisible(False)
                else:
                    if hasattr(self, 'measuredContainer'):
                        self.measuredContainer.setVisible(False)
                    if hasattr(self, 'formulaContainer'):
                        self.formulaContainer.setVisible(True)
        except Exception as e:
            print(f"Error updating UI based on data source type: {e}")


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
        
        # データと状態の初期化
        self.datasets = []  # 複数のデータセットを格納するリスト
        self.current_dataset_index = -1  # 現在選択されているデータセットのインデックス
        
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
        
        # 色の初期化
        self.currentColor = QColor('blue')
        
        # 目盛り間隔の値を保持する変数を初期化
        self.x_tick_step = 1.0
        self.y_tick_step = 1.0
        
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
        convertButton.setFixedHeight(32)
        mainLayout.addWidget(convertButton)
        
        # 結果表示エリア
        self.resultText = QTextEdit()
        self.resultText.setReadOnly(True)
        self.resultText.setMinimumHeight(100)
        self.resultText.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
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
        self.xLabelEntry = QLineEdit('x軸')
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
        self.yLabelEntry = QLineEdit('y軸')
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
        self.legendPosCombo.setCurrentText('右上')
        
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
        self.labelEntry = QLineEdit('fig:tikz_plot')
        
        # 位置
        self.positionCombo = QComboBox()
        self.positionCombo.addItems(['h', 'htbp', 't', 'b', 'p', 'H'])
        self.positionCombo.setCurrentText('htbp')
        
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
        self.lineRadio.setChecked(True)  # デフォルト
        
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
        self.markerSizeSpin.setValue(2.0)
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
                    if self.statusBar:
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
                if self.statusBar:
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
            }
            
            self.datasets.append(dataset)
            self.datasetList.addItem(final_name) # addItem directly with the guaranteed string
            
            # 新しく追加したデータセットを選択（これがon_dataset_selectedを呼び出す）
            self.datasetList.setCurrentRow(len(self.datasets) - 1)
            if self.statusBar:
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
                    self.update_ui_for_no_datasets() 
                    self.add_dataset("データセット1") # Or add a new default one
                
                if self.statusBar:
                    self.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット削除中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_for_no_datasets(self):
        """データセットがない場合にUIをリセット/クリアする"""
        # 例: 関連する入力フィールドをクリアまたは無効化
        if hasattr(self, 'legendLabel'):
            self.legendLabel.setText("")
        # 他のUI要素も必要に応じてリセット
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
                    if hasattr(self, 'legendLabel'):
                        self.legendLabel.setText(actual_new_name) # UIも即時更新
                
                item = self.datasetList.item(current_row)
                if item:
                    item.setText(actual_new_name)
                if self.statusBar:
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
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset['name']}' からUIを更新しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_current_dataset(self):
        try:
            if self.current_dataset_index < 0 or not self.datasets or self.current_dataset_index >= len(self.datasets):
                return
            
            # 現在のデータセットを取得
            dataset = self.datasets[self.current_dataset_index]
            
            # 色の設定（1回のみ）
            if not isinstance(self.currentColor, QColor):
                self.currentColor = QColor(self.currentColor)
            # QColorオブジェクトのコピーを作成して代入
            dataset['color'] = QColor(self.currentColor)
            
            # 基本設定の更新
            if hasattr(self, 'lineWidthSpin'):
                dataset['line_width'] = self.lineWidthSpin.value()
            if hasattr(self, 'markerCombo'):
                dataset['marker_style'] = self.markerCombo.currentText()
            if hasattr(self, 'markerSizeSpin'):
                dataset['marker_size'] = self.markerSizeSpin.value()
            if hasattr(self, 'legendCheck'):
                dataset['show_legend'] = self.legendCheck.isChecked()
            if hasattr(self, 'legendLabel'):
                dataset['legend_label'] = self.legendLabel.text()
            
            # プロットタイプの更新
            if hasattr(self, 'lineRadio') and self.lineRadio.isChecked():
                dataset['plot_type'] = 'line'
            elif hasattr(self, 'scatterRadio') and self.scatterRadio.isChecked():
                dataset['plot_type'] = 'scatter'
            elif hasattr(self, 'lineScatterRadio') and self.lineScatterRadio.isChecked():
                dataset['plot_type'] = 'line+scatter'
            elif hasattr(self, 'barRadio') and self.barRadio.isChecked():
                dataset['plot_type'] = 'bar'
            
            # データソースタイプの更新
            if hasattr(self, 'measuredRadio') and self.measuredRadio.isChecked():
                dataset['data_source_type'] = 'measured'
            elif hasattr(self, 'formulaRadio') and self.formulaRadio.isChecked():
                dataset['data_source_type'] = 'formula'
            
            # 数式関連の設定
            if hasattr(self, 'equationEntry'):
                dataset['equation'] = self.equationEntry.text()
            if hasattr(self, 'domainMinSpin'):
                dataset['domain_min'] = self.domainMinSpin.value()
            if hasattr(self, 'domainMaxSpin'):
                dataset['domain_max'] = self.domainMaxSpin.value()
            if hasattr(self, 'samplesSpin'):
                dataset['samples'] = self.samplesSpin.value()
            
            if self.statusBar:
                self.statusBar.showMessage(f"データセット '{dataset['name']}' が更新されました", 3000)
                
        except Exception as e:
            print(f"Error updating current dataset: {e}")
    
    def update_ui_from_dataset(self, dataset):
        """データセットの内容でUIを更新"""
        try:
            self.block_signals_temporarily(True)
            
            # 色の更新
            color = dataset.get('color', QColor('blue'))
            if isinstance(color, QColor):
                self.currentColor = QColor(color)
                if hasattr(self, 'colorButton'):
                    self.colorButton.setStyleSheet(f"background-color: {color.name()}")
            
            # 基本設定の更新
            if hasattr(self, 'lineWidthSpin'):
                self.lineWidthSpin.setValue(dataset.get('line_width', 1.0))
            if hasattr(self, 'markerCombo'):
                self.markerCombo.setCurrentText(dataset.get('marker_style', '*'))
            if hasattr(self, 'markerSizeSpin'):
                self.markerSizeSpin.setValue(dataset.get('marker_size', 2.0))
            if hasattr(self, 'legendCheck'):
                self.legendCheck.setChecked(dataset.get('show_legend', True))
            if hasattr(self, 'legendLabel'):
                self.legendLabel.setText(dataset.get('legend_label', dataset['name']))
            
            # プロットタイプの更新
            plot_type = dataset.get('plot_type', 'line')
            if hasattr(self, 'lineRadio'):
                self.lineRadio.setChecked(plot_type == 'line')
            if hasattr(self, 'scatterRadio'):
                self.scatterRadio.setChecked(plot_type == 'scatter')
            if hasattr(self, 'lineScatterRadio'):
                self.lineScatterRadio.setChecked(plot_type == 'line+scatter')
            if hasattr(self, 'barRadio'):
                self.barRadio.setChecked(plot_type == 'bar')
            
            # データソースタイプの更新
            data_source_type = dataset.get('data_source_type', 'measured')
            if hasattr(self, 'measuredRadio'):
                self.measuredRadio.setChecked(data_source_type == 'measured')
            if hasattr(self, 'formulaRadio'):
                self.formulaRadio.setChecked(data_source_type == 'formula')
            
            # 数式関連の設定
            if hasattr(self, 'equationEntry'):
                self.equationEntry.setText(dataset.get('equation', ''))
            if hasattr(self, 'domainMinSpin'):
                self.domainMinSpin.setValue(dataset.get('domain_min', 0))
            if hasattr(self, 'domainMaxSpin'):
                self.domainMaxSpin.setValue(dataset.get('domain_max', 10))
            if hasattr(self, 'samplesSpin'):
                self.samplesSpin.setValue(dataset.get('samples', 200))
            
            # データテーブルの更新
            self.update_data_table_from_dataset(dataset)
            
            # データソースタイプに基づくUIの更新
            self.update_ui_based_on_data_source_type()
            
            self.block_signals_temporarily(False)
        except Exception as e:
            print(f"Error updating UI from dataset: {e}")
    
    def block_signals_temporarily(self, block):
        """シグナルを一時的にブロック"""
        widgets = []
        if hasattr(self, 'lineWidthSpin'):
            widgets.append(self.lineWidthSpin)
        if hasattr(self, 'markerCombo'):
            widgets.append(self.markerCombo)
        if hasattr(self, 'markerSizeSpin'):
            widgets.append(self.markerSizeSpin)
        if hasattr(self, 'legendCheck'):
            widgets.append(self.legendCheck)
        if hasattr(self, 'legendLabel'):
            widgets.append(self.legendLabel)
        if hasattr(self, 'lineRadio'):
            widgets.append(self.lineRadio)
        if hasattr(self, 'scatterRadio'):
            widgets.append(self.scatterRadio)
        if hasattr(self, 'lineScatterRadio'):
            widgets.append(self.lineScatterRadio)
        if hasattr(self, 'barRadio'):
            widgets.append(self.barRadio)
        
        for widget in widgets:
            widget.blockSignals(block)
    
    def update_data_table_from_dataset(self, dataset):
        """データセットからデータテーブルを更新"""
        try:
            if hasattr(self, 'dataTable'):
                # テーブルをクリア
                self.dataTable.setRowCount(0)
                
                # データがある場合はテーブルに設定
                data_x = dataset.get('data_x', [])
                data_y = dataset.get('data_y', [])
                
                if data_x and data_y:
                    max_rows = max(len(data_x), len(data_y))
                    self.dataTable.setRowCount(max_rows)
                    
                    for i in range(max_rows):
                        # X値の設定
                        if i < len(data_x):
                            x_item = QTableWidgetItem(str(data_x[i]))
                            self.dataTable.setItem(i, 0, x_item)
                        
                        # Y値の設定
                        if i < len(data_y):
                            y_item = QTableWidgetItem(str(data_y[i]))
                            self.dataTable.setItem(i, 1, y_item)
                else:
                    # データがない場合は10行の空テーブルを作成
                    self.dataTable.setRowCount(10)
        except Exception as e:
            print(f"Error updating data table: {e}")
    
    def update_ui_based_on_data_source_type(self):
        """データソースタイプに基づいてUIの表示を更新"""
        try:
            if hasattr(self, 'measuredRadio') and hasattr(self, 'formulaRadio'):
                if self.measuredRadio.isChecked():
                    if hasattr(self, 'measuredContainer'):
                        self.measuredContainer.setVisible(True)
                    if hasattr(self, 'formulaContainer'):
                        self.formulaContainer.setVisible(False)
                else:
                    if hasattr(self, 'measuredContainer'):
                        self.measuredContainer.setVisible(False)
                    if hasattr(self, 'formulaContainer'):
                        self.formulaContainer.setVisible(True)
        except Exception as e:
            print(f"Error updating UI based on data source type: {e}")

    def select_color(self):
        """色を選択"""
        color = QColorDialog.getColor(self.currentColor)
        if color.isValid():
            self.currentColor = color
            self.colorButton.setStyleSheet(f"background-color: {color.name()}")
            # 現在のデータセットにも反映
            if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
                self.datasets[self.current_dataset_index]['color'] = QColor(color)

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
        # データセットが空かチェック
        if not self.datasets:
            QMessageBox.warning(self, "警告", "データセットがありません。先にデータセットを追加してください。")
            return
        
        # データが存在するデータセットがあるかチェック
        has_data = False
        for dataset in self.datasets:
            if dataset.get('data_x') and dataset.get('data_y'):
                has_data = True
                break
        
        if not has_data:
            QMessageBox.warning(self, "警告", "データが入力されているデータセットがありません。\n"
                              "先に「データ入力」タブでデータを読み込んでください。")
            return
        
        try:
            # 現在のデータセットの設定を保存
            if self.current_dataset_index >= 0:
                self.update_current_dataset()
            
            # グローバル設定を更新
            self.update_global_settings()
            
            tikz_code = self.generate_tikz_code_multi_datasets()
            self.resultText.setPlainText(tikz_code)
            if self.statusBar:
                self.statusBar.showMessage("TikZコードが生成されました")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"TikZコード生成中にエラーが発生しました: {str(e)}")
            if self.statusBar:
                self.statusBar.showMessage("コード生成エラー")

    def update_global_settings(self):
        """グローバル設定を更新"""
        try:
            if hasattr(self, 'xLabelEntry'):
                self.global_settings['x_label'] = self.xLabelEntry.text()
            if hasattr(self, 'yLabelEntry'):
                self.global_settings['y_label'] = self.yLabelEntry.text()
            if hasattr(self, 'xMinSpin'):
                self.global_settings['x_min'] = self.xMinSpin.value()
            if hasattr(self, 'xMaxSpin'):
                self.global_settings['x_max'] = self.xMaxSpin.value()
            if hasattr(self, 'yMinSpin'):
                self.global_settings['y_min'] = self.yMinSpin.value()
            if hasattr(self, 'yMaxSpin'):
                self.global_settings['y_max'] = self.yMaxSpin.value()
            if hasattr(self, 'gridCheck'):
                self.global_settings['grid'] = self.gridCheck.isChecked()
            if hasattr(self, 'legendCheck'):
                self.global_settings['show_legend'] = self.legendCheck.isChecked()
            if hasattr(self, 'legendPosCombo'):
                legend_pos_map = {
                    '左上': 'north west',
                    '右上': 'north east',
                    '左下': 'south west',
                    '右下': 'south east'
                }
                jp_pos = self.legendPosCombo.currentText()
                self.global_settings['legend_pos'] = legend_pos_map.get(jp_pos, 'north east')
            if hasattr(self, 'widthSpin'):
                self.global_settings['width'] = self.widthSpin.value()
            if hasattr(self, 'heightSpin'):
                self.global_settings['height'] = self.heightSpin.value()
            if hasattr(self, 'captionEntry'):
                self.global_settings['caption'] = self.captionEntry.text()
            if hasattr(self, 'labelEntry'):
                self.global_settings['label'] = self.labelEntry.text()
            if hasattr(self, 'positionCombo'):
                self.global_settings['position'] = self.positionCombo.currentText()
            
            # 目盛り間隔の更新
            if hasattr(self, 'xTickStepSpin'):
                self.x_tick_step = self.xTickStepSpin.value()
            if hasattr(self, 'yTickStepSpin'):
                self.y_tick_step = self.yTickStepSpin.value()
                
        except Exception as e:
            print(f"Error updating global settings: {e}")

    def generate_tikz_code_multi_datasets(self):
        """複数のデータセットを含むTikZコードを生成する"""
        # LaTeXコード生成
        latex = []
        
        # 図の開始
        latex.append("\\begin{figure}[" + self.global_settings['position'] + "]")
        latex.append("  \\centering")
        
        # 図の幅と高さ設定
        width = self.global_settings['width']
        height = self.global_settings['height']
        
        # TikZ図の開始
        latex.append("  \\begin{tikzpicture}")
        
        # 軸設定
        x_min = self.global_settings['x_min']
        x_max = self.global_settings['x_max']
        y_min = self.global_settings['y_min']
        y_max = self.global_settings['y_max']
        
        # 軸の範囲チェックと補正
        # データの実際の最小値と最大値を計算
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
            
            # データポイントが1つしかない場合や、min/maxが同じ場合の処理
            if abs(data_x_max - data_x_min) < 1e-10:
                data_x_min -= 0.5 if abs(data_x_min) > 1 else 0.1
                data_x_max += 0.5 if abs(data_x_max) > 1 else 0.1
            
            if abs(data_y_max - data_y_min) < 1e-10:
                data_y_min -= 0.5 if abs(data_y_min) > 1 else 0.1
                data_y_max += 0.5 if abs(data_y_max) > 1 else 0.1
            
            # 軸の範囲がデータの範囲を含んでいるか確認し、必要に応じて調整
            if x_min > data_x_min or x_min == 0:
                x_min = data_x_min - abs(data_x_min) * 0.1 - 0.1
            if x_max < data_x_max or x_max == 0:
                x_max = data_x_max + abs(data_x_max) * 0.1 + 0.1
            if y_min > data_y_min or y_min == 0:
                y_min = data_y_min - abs(data_y_min) * 0.1 - 0.1
            if y_max < data_y_max or y_max == 0:
                y_max = data_y_max + abs(data_y_max) * 0.1 + 0.1
            
            # 小さすぎる値は0に近い値に設定
            if abs(y_min) < 1e-10:
                y_min = -0.1
            if abs(y_max) < 1e-10:
                y_max = 0.1
            if abs(x_min) < 1e-10:
                x_min = -0.1
            if abs(x_max) < 1e-10:
                x_max = 0.1
            
            # 範囲が同じ値の場合（データが全て同じ値の場合）
            if abs(x_min - x_max) < 1e-10:
                x_min -= 0.5
                x_max += 0.5
            if abs(y_min - y_max) < 1e-10:
                y_min -= 0.5
                y_max += 0.5
                
            # 範囲が極端に小さい場合も調整
            if abs(x_max - x_min) < 1e-3:
                margin = abs(x_min) * 0.2 if abs(x_min) > 1e-10 else 0.1
                x_min -= margin
                x_max += margin
            if abs(y_max - y_min) < 1e-3:
                margin = abs(y_min) * 0.2 if abs(y_min) > 1e-10 else 0.1
                y_min -= margin
                y_max += margin
        
        # 目盛りの設定
        xtick_values = []
        ytick_values = []
        
        # X軸の目盛りを設定
        if self.x_tick_step > 0:
            tick_min = math.ceil(x_min / self.x_tick_step) * self.x_tick_step
            tick_max = math.floor(x_max / self.x_tick_step) * self.x_tick_step
            
            current = tick_min
            while current <= tick_max:
                xtick_values.append(current)
                current += self.x_tick_step
        
        # Y軸の目盛りを設定
        if self.y_tick_step > 0:
            tick_min = math.ceil(y_min / self.y_tick_step) * self.y_tick_step
            tick_max = math.floor(y_max / self.y_tick_step) * self.y_tick_step
            
            current = tick_min
            while current <= tick_max:
                ytick_values.append(current)
                current += self.y_tick_step
        
        # axis環境の設定
        axis_options = []
        axis_options.append(f"width={width}\\textwidth")
        axis_options.append(f"height={height}\\textwidth")
        axis_options.append(f"xlabel={{{self.global_settings['x_label']}}}")
        axis_options.append(f"ylabel={{{self.global_settings['y_label']}}}")
        
        if x_min != x_max:
            axis_options.append(f"xmin={x_min}, xmax={x_max}")
        if y_min != y_max:
            axis_options.append(f"ymin={y_min}, ymax={y_max}")
        
        # xtick, ytickを自動生成
        if xtick_values:
            xticks = ','.join(str(round(tick, 8)) for tick in xtick_values)
        else:
            # ゼロ除算を防ぐ
            if abs(x_max - x_min) < 1e-10 or self.x_tick_step < 1e-10:
                xticks = str(round(x_min, 8))
            else:
                steps = max(1, min(20, int((x_max - x_min) / self.x_tick_step) + 1))  # 最大20ステップに制限
                xticks = ','.join(str(round(x_min + i * self.x_tick_step, 8)) for i in range(steps))
        
        if ytick_values:
            yticks = ','.join(str(round(tick, 8)) for tick in ytick_values)
        else:
            # ゼロ除算を防ぐ
            if abs(y_max - y_min) < 1e-10 or self.y_tick_step < 1e-10:
                yticks = str(round(y_min, 8))
            else:
                steps = max(1, min(20, int((y_max - y_min) / self.y_tick_step) + 1))  # 最大20ステップに制限
                yticks = ','.join(str(round(y_min + i * self.y_tick_step, 8)) for i in range(steps))
        
        axis_options.append(f"xtick={{{xticks}}}")
        axis_options.append(f"ytick={{{yticks}}}")
        
        # グリッドの設定
        if self.global_settings['grid']:
            axis_options.append("grid=major")
        
        # 凡例の設定
        if self.global_settings['show_legend']:
            axis_options.append(f"legend pos={self.global_settings['legend_pos']}")
        
        # axis環境の開始
        latex.append("    \\begin{axis}[")
        for i, option in enumerate(axis_options):
            if i == len(axis_options) - 1:
                latex.append(f"      {option}")
            else:
                latex.append(f"      {option},")
        latex.append("    ]")
        
        # 各データセットをプロット
        for i, dataset in enumerate(self.datasets):
            if dataset.get('data_x') and dataset.get('data_y'):
                self.add_dataset_to_latex(latex, dataset, i)
        
        # axis環境の終了
        latex.append("    \\end{axis}")
        latex.append("  \\end{tikzpicture}")
        
        # キャプションとラベル
        latex.append(f"  \\caption{{{self.global_settings['caption']}}}")
        latex.append(f"  \\label{{{self.global_settings['label']}}}")
        
        # 図の終了
        latex.append("\\end{figure}")
        
        return "\n".join(latex)

    def add_dataset_to_latex(self, latex, dataset, index):
        """データセットをLaTeXコードに追加"""
        # 座標データの生成
        coordinates = []
        data_x = dataset.get('data_x', [])
        data_y = dataset.get('data_y', [])
        
        for x, y in zip(data_x, data_y):
            coordinates.append(f"({x},{y})")
        coords_str = " ".join(coordinates)
        
        # プロットオプション
        options = []
        
        # 色の設定
        color = dataset.get('color', QColor('blue'))
        if isinstance(color, QColor):
            color_str = self.color_to_tikz_rgb(color)
            if color_str != 'blue':  # デフォルト色でない場合のみ追加
                options.append(color_str)
        
        # プロットタイプの設定
        plot_type = dataset.get('plot_type', 'line')
        if plot_type == 'scatter':
            options.append("only marks")
        elif plot_type == 'line':
            pass  # デフォルトで線が引かれる
        elif plot_type == 'line+scatter':
            pass  # 線とマーカー両方
        elif plot_type == 'bar':
            options.append("ybar")
        
        # マーカーの設定
        marker = dataset.get('marker_style', '*')
        if marker and plot_type in ['scatter', 'line+scatter']:
            options.append(f"mark={marker}")
        
        # マーカーサイズの設定
        marker_size = dataset.get('marker_size', 2.0)
        if marker and plot_type in ['scatter', 'line+scatter']:
            options.append(f"mark size={marker_size}pt")
        
        # 線の太さの設定
        line_width = dataset.get('line_width', 1.0)
        if plot_type in ['line', 'line+scatter']:
            options.append(f"line width={line_width}pt")
        
        options_str = ", ".join(options) if options else ""
        
        # addplotコマンド
        if options_str:
            latex.append(f"      \\addplot[{options_str}] coordinates {{{coords_str}}};")
        else:
            latex.append(f"      \\addplot coordinates {{{coords_str}}};")
        
        # 凡例
        if dataset.get('show_legend', True) and self.global_settings['show_legend']:
            legend_label = dataset.get('legend_label', dataset.get('name', f'Dataset {index+1}'))
            latex.append(f"      \\addlegendentry{{{legend_label}}}")

    def copy_to_clipboard(self):
        """結果をクリップボードにコピー"""
        tikz_code = self.resultText.toPlainText()
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