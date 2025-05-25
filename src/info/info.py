import webbrowser
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, QMessageBox, QFrame)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class InfoTab(QWidget):
    
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        titleLabel = QLabel("TeXcellentConverter")
        titleLabel.setFont(QFont("Arial", 30, QFont.Bold))
        titleLabel.setAlignment(Qt.AlignCenter)
        titleLabel.setStyleSheet("color: #0095d9; margin: 20px;")
        
        versionLabel = QLabel("Version 1.0.0")
        versionLabel.setFont(QFont("Arial", 16))
        versionLabel.setAlignment(Qt.AlignCenter)
        versionLabel.setStyleSheet("color: #7F8C8D; margin-bottom: 30px;")
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #BDC3C7; margin: 10px 50px;")
        
        buttonLayout = QVBoxLayout()

        groupLabel = QLabel("🔗 リンク")
        groupLabel.setFont(QFont("Arial", 24 , QFont.Bold))
        groupLabel.setAlignment(Qt.AlignLeft)
        groupLabel.setStyleSheet("color: #0095d9; margin: 20px;")
        
        githubButton = QPushButton("📂 ソースコードを参照")
        githubButton.setFont(QFont("Arial", 12))
        githubButton.setStyleSheet("""
            QPushButton {
                background-color: #24292e;
                color: white;
                border: none;
                padding: 15px 1px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #444d56;
            }
        """)
        githubButton.clicked.connect(lambda: webbrowser.open("https://github.com/yudo417/TeXcellentConverter"))
        
        licenseButton = QPushButton("📄 ライセンス")
        licenseButton.setFont(QFont("Arial", 12))
        licenseButton.setStyleSheet("""
            QPushButton {
                background-color: #0366d6;
                color: white;
                border: none;
                padding: 15px 1px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0256cc;
            }
        """)
        licenseButton.clicked.connect(lambda: webbrowser.open("https://github.com/yudo417/TeXcellentConverter/blob/main/LICENSE"))
        
        helpButton = QPushButton("📖 使い方を参照")
        helpButton.setFont(QFont("Arial", 12))
        helpButton.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                padding: 15px 1px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        helpButton.clicked.connect(lambda: webbrowser.open("https://github.com/yudo417/TeXcellentConverter#%E4%BD%BF%E7%94%A8%E6%96%B9%E6%B3%95"))
        

        
        buttonLayout.addWidget(githubButton)
        buttonLayout.addWidget(licenseButton)
        buttonLayout.addWidget(helpButton)
        buttonLayout.setSpacing(30)
        
        buttonLayout.setContentsMargins(50, 30, 50, 0)
        
        
        layout.addWidget(titleLabel)
        layout.addWidget(versionLabel)
        layout.addWidget(groupLabel)
        layout.addWidget(line)
        layout.addLayout(buttonLayout)
        layout.addStretch()
        
        self.setLayout(layout)

