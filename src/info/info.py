import webbrowser
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class InfoTab(QWidget):
    """Application info tab - ライセンス等の参照タブ"""
    
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        """Initialize the info tab UI"""
        layout = QVBoxLayout()
        
        # アプリケーション情報
        titleLabel = QLabel("TexThon")
        titleLabel.setFont(QFont("Arial", 24, QFont.Bold))
        titleLabel.setAlignment(Qt.AlignCenter)
        titleLabel.setStyleSheet("color: #2C3E50; margin: 20px;")
        
        versionLabel = QLabel("Version 2.0.0")
        versionLabel.setFont(QFont("Arial", 12))
        versionLabel.setAlignment(Qt.AlignCenter)
        versionLabel.setStyleSheet("color: #7F8C8D; margin-bottom: 30px;")
        
        descriptionLabel = QLabel("Excel表とデータを美しいLaTeX形式に変換するアプリケーション\n完全独立アーキテクチャで安全・確実な変換を実現")
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
        githubButton.clicked.connect(lambda: webbrowser.open("https://github.com/hayashikazuki/texthon"))
        
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
        
        # 新機能ボタン
        featuresButton = QPushButton("✨ 新機能")
        featuresButton.setFont(QFont("Arial", 12))
        featuresButton.setStyleSheet("""
            QPushButton {
                background-color: #6f42c1;
                color: white;
                border: none;
                padding: 15px 30px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5a32a3;
            }
        """)
        featuresButton.clicked.connect(self.show_features_info)
        
        buttonLayout.addWidget(githubButton)
        buttonLayout.addWidget(licenseButton)
        buttonLayout.addWidget(helpButton)
        buttonLayout.addWidget(featuresButton)
        buttonLayout.setSpacing(15)
        
        # ライセンス情報テキスト
        licenseText = QLabel("""
<b>ライセンス:</b> MIT License<br><br>
<b>作成者:</b> Kazuki Hayashi<br><br>
<b>使用ライブラリ:</b><br>
• PyQt5 - GUI フレームワーク<br>
• openpyxl - Excel ファイル処理<br>
• pandas - データ処理<br>
• numpy - 数値計算<br><br>
<b>アーキテクチャ:</b> 完全独立タブ設計<br>
<b>安全性:</b> タブ間の干渉なし<br>
<b>保守性:</b> 各機能の独立実装
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
        """Show license information dialog"""
        QMessageBox.information(self, "ライセンス情報", 
                               "MIT License\n\n"
                               "Copyright (c) 2024 Kazuki Hayashi\n\n"
                               "Permission is hereby granted, free of charge, to any person obtaining a copy "
                               "of this software and associated documentation files (the \"Software\"), to deal "
                               "in the Software without restriction, including without limitation the rights "
                               "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell "
                               "copies of the Software, and to permit persons to whom the Software is "
                               "furnished to do so, subject to the following conditions:\n\n"
                               "The above copyright notice and this permission notice shall be included in all "
                               "copies or substantial portions of the Software.")

    def show_help_info(self):
        """Show help information dialog"""
        QMessageBox.information(self, "使い方", 
                               "【Excel → LaTeX 表】\n"
                               "1. Excelファイルを選択\n"
                               "2. シート名を選択\n"
                               "3. セル範囲を指定（例: A1:E6）\n"
                               "4. オプションを設定\n"
                               "5. 「LaTeXに変換」をクリック\n"
                               "6. 生成されたコードをコピー\n\n"
                               "【TikZ グラフ】\n"
                               "1. データソースを選択（CSV/Excel/手動入力/数式）\n"
                               "2. ファイルを選択またはデータを手動入力\n"
                               "3. データセットを追加・管理\n"
                               "4. プロット設定（色、マーカー、線の太さなど）\n"
                               "5. 軸設定（タイトル、ラベル、範囲）\n"
                               "6. 特殊ポイント・アノテーション設定\n"
                               "7. 「TikZコードを生成」をクリック\n"
                               "8. 生成されたコードをコピー\n\n"
                               "【LaTeXドキュメントでの使用】\n"
                               "・プリアンブルに必要なパッケージを追加\n"
                               "・生成されたコードをドキュメントに貼り付け\n"
                               "・コンパイルして結果を確認")

    def show_features_info(self):
        """Show new features information dialog"""
        QMessageBox.information(self, "新機能・特徴", 
                               "【Version 2.0.0 新機能】\n"
                               "🔥 完全独立アーキテクチャ\n"
                               "　各タブが完全に独立し、バグの相互影響を排除\n\n"
                               "🛡️ 安全性の向上\n"
                               "　UI部品の受け渡しによるバグを根本的に解決\n\n"
                               "📊 Excel表変換\n"
                               "　・結合セル完全対応\n"
                               "　・複雑な表構造の正確な変換\n"
                               "　・multirow パッケージの自動活用\n\n"
                               "📈 TikZ グラフ生成\n"
                               "　・複数データセットサポート\n"
                               "　・数式からのデータ生成\n"
                               "　・豊富なプロットタイプ\n"
                               "　・カスタマイズ可能なスタイル\n\n"
                               "⚡ パフォーマンス\n"
                               "　・高速なデータ処理\n"
                               "　・メモリ効率の最適化\n\n"
                               "🎨 ユーザビリティ\n"
                               "　・直感的なインターフェース\n"
                               "　・リアルタイムプレビュー\n"
                               "　・ワンクリックコピー")

    def set_status_bar(self, status_bar):
        """Set status bar (compatibility method)"""
        # Info tab doesn't use status bar, but keeping for consistency
        pass

