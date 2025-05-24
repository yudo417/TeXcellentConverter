from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QInputDialog, QLineEdit, QMessageBox


class DatasetManager:
    def __init__(self, parent):
        self.parent = parent
        self.datasets = []
        self.current_dataset_index = -1
        
    def add_dataset(self, name_arg=None):
        """新しいデータセットを追加する"""
        try:
            # 既存のデータセットの状態を保存（存在する場合）
            if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
                self.parent.update_current_dataset()
                
            final_name = ""
            if name_arg is None:
                dataset_count = self.parent.datasetList.count() + 1
                text_from_dialog, ok = QInputDialog.getText(self.parent, "データセット名", "新しいデータセット名を入力してください:",
                                                          QLineEdit.Normal, f"データセット{dataset_count}")
                if not ok or not text_from_dialog.strip():
                    if self.parent.statusBar:
                        self.parent.statusBar.showMessage("データセットの追加がキャンセルされました。", 3000)
                    return
                final_name = text_from_dialog.strip()
            else:
                if name_arg is False:
                    dataset_count = self.parent.datasetList.count() + 1
                    final_name = f"データセット{dataset_count}"
                else:
                    final_name = str(name_arg).strip()

            if not final_name:
                dataset_count = self.parent.datasetList.count() + 1
                final_name = f"データセット{dataset_count}"
                if self.parent.statusBar:
                    self.parent.statusBar.showMessage("データセット名が空のため、デフォルト名を使用します。", 3000)

            # 明示的に空のデータと初期設定を持つデータセットを作成
            dataset = {
                'name': final_name,
                'data_source_type': 'measured',
                'data_x': [],
                'data_y': [],
                'color': QColor('blue'),
                'line_width': 1.0,
                'marker_style': '*',
                'marker_size': 2.0,
                'plot_type': "line",
                'legend_label': final_name,
                'show_legend': True,
                'equation': '',
                'domain_min': 0,
                'domain_max': 10,
                'samples': 200,
                'special_points': [],
                'annotations': [],
                'file_path': '',
                'file_type': 'csv',
                'sheet_name': '',
                'x_column': '',
                'y_column': ''
            }
            
            self.datasets.append(dataset)
            self.parent.datasetList.addItem(final_name)
            
            # 新しく追加したデータセットを選択
            self.parent.datasetList.setCurrentRow(len(self.datasets) - 1)
            if self.parent.statusBar:
                self.parent.statusBar.showMessage(f"データセット '{final_name}' を追加しました", 3000)

        except Exception as e:
            import traceback
            QMessageBox.critical(self.parent, "エラー", f"データセット追加中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def remove_dataset(self):
        """選択されたデータセットを削除する"""
        try:
            if not self.datasets:
                QMessageBox.warning(self.parent, "警告", "削除するデータセットがありません")
                return
            
            current_row = self.parent.datasetList.currentRow()
            if current_row < 0:
                QMessageBox.warning(self.parent, "警告", "削除するデータセットを選択してください")
                return
            
            dataset_name = str(self.datasets[current_row]['name'])
            reply = QMessageBox.question(self.parent, "確認",
                                       f"データセット '{dataset_name}' を削除してもよろしいですか？",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.datasets.pop(current_row)
                item = self.parent.datasetList.takeItem(current_row)
                if item:
                    del item
                
                if self.datasets:
                    new_index = max(0, min(current_row, len(self.datasets) - 1))
                    self.parent.datasetList.setCurrentRow(new_index)
                else:
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets()
                    self.add_dataset("データセット1")
                
                if self.parent.statusBar:
                    self.parent.statusBar.showMessage(f"データセット '{dataset_name}' を削除しました", 3000)
        except Exception as e:
            import traceback
            QMessageBox.critical(self.parent, "エラー", f"データセット削除中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def rename_dataset(self):
        """選択されたデータセットの名前を変更する"""
        try:
            if self.current_dataset_index < 0 or not self.datasets:
                QMessageBox.warning(self.parent, "警告", "名前を変更するデータセットが選択されていません。")
                return
            
            current_row = self.parent.datasetList.currentRow()
            if current_row != self.current_dataset_index:
                current_row = self.current_dataset_index
                self.parent.datasetList.setCurrentRow(current_row)

            current_name = str(self.datasets[current_row]['name'])
            current_legend = str(self.datasets[current_row].get('legend_label', current_name))

            new_name_text, ok = QInputDialog.getText(self.parent, "データセット名の変更",
                                                      "新しいデータセット名を入力してください:",
                                                      QLineEdit.Normal, current_name)
            
            if ok and new_name_text.strip():
                actual_new_name = new_name_text.strip()
                if not actual_new_name:
                    QMessageBox.warning(self.parent, "警告", "データセット名は空にできません。")
                    return

                self.datasets[current_row]['name'] = actual_new_name
                if current_legend == current_name:
                    self.datasets[current_row]['legend_label'] = actual_new_name
                    if hasattr(self.parent, 'legendLabel'):
                        self.parent.legendLabel.setText(actual_new_name)
                
                item = self.parent.datasetList.item(current_row)
                if item:
                    item.setText(actual_new_name)
                if self.parent.statusBar:
                    self.parent.statusBar.showMessage(f"データセット名を '{actual_new_name}' に変更しました", 3000)
            elif ok and not new_name_text.strip():
                 QMessageBox.warning(self.parent, "警告", "データセット名は空にできません。")

        except Exception as e:
            import traceback
            QMessageBox.critical(self.parent, "エラー", f"データセット名の変更中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def on_dataset_selected(self, row):
        """データセット選択時の処理"""
        try:
            # 以前のデータセットの状態を保存
            old_index = self.current_dataset_index
            if old_index >= 0 and old_index < len(self.datasets):
                self.parent.update_current_dataset()
            
            if row < 0 or row >= len(self.datasets):
                if not self.datasets:
                    self.current_dataset_index = -1
                    self.update_ui_for_no_datasets()
                return
            
            # 現在のインデックスを更新
            self.current_dataset_index = row
            dataset = self.datasets[row]
            self.parent.update_ui_from_dataset(dataset)
            
            if self.parent.statusBar:
                self.parent.statusBar.showMessage(f"データセット '{dataset['name']}' からUIを更新しました", 3000)
                
        except Exception as e:
            import traceback
            QMessageBox.critical(self.parent, "エラー", f"データセット選択処理中にエラーが発生しました: {str(e)}\n\n{traceback.format_exc()}")
    
    def update_ui_for_no_datasets(self):
        """データセットがない場合にUIをリセット"""
        if hasattr(self.parent, 'legendLabel'):
            self.parent.legendLabel.setText("")
    
    def get_current_dataset(self):
        """現在選択されているデータセットを取得"""
        if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
            return self.datasets[self.current_dataset_index]
        return None
    
    def update_current_dataset_data(self, data_x, data_y):
        """現在のデータセットのデータを更新"""
        if self.current_dataset_index >= 0 and self.current_dataset_index < len(self.datasets):
            self.datasets[self.current_dataset_index]['data_x'] = data_x
            self.datasets[self.current_dataset_index]['data_y'] = data_y 