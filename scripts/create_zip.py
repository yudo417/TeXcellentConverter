#!/usr/bin/env python3
"""
TeXcellentConverter EXE版配布用zipファイル作成スクリプト
"""

import os
import shutil
import zipfile
from pathlib import Path

def create_exe_zip():
    """EXE版の配布用zipファイルを作成"""
    
    print("🚀 TeXcellentConverter EXE版 配布用zipファイルを作成中...")
    
    # ビルドディレクトリの確認
    build_dir = Path("build/exe.win-amd64-3.12")
    if not build_dir.exists():
        print("❌ EXEファイルが見つかりません。先にビルドしてください。")
        return
    
    # zipファイルを作成
    zip_filename = "TeXcellentConverter_EXE.zip"
    
    print("📦 zipファイルを作成中...")
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # buildディレクトリ内のすべてのファイルを含める
        for root, dirs, files in os.walk(build_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # zipファイル内でのパスをTeXcellentConverter/にする
                arc_name = os.path.join("TeXcellentConverter", 
                                       os.path.relpath(file_path, build_dir))
                zipf.write(file_path, arc_name)
    
    # READMEファイルを追加
    readme_content = """# TeXcellentConverter EXE版

## 使用方法
1. このzipファイルを解凍
2. `TeXcellentConverter.exe` をダブルクリックして実行

## 特徴
- Pythonのインストール不要
- スタンドアロンで動作
- Excel表データをLaTeX形式に変換
- グラフ作成機能
- 直感的なUI

## システム要件
- Windows 10以上
- 64bit環境

ご利用ありがとうございます！
"""
    
    with zipfile.ZipFile(zip_filename, 'a') as zipf:
        zipf.writestr("TeXcellentConverter/README.txt", readme_content)
    
    print(f"✅ 作成完了: {zip_filename}")
    print(f"📁 ファイルサイズ: {os.path.getsize(zip_filename) / 1024 / 1024:.1f} MB")
    print("\n🎉 EXE版配布準備完了！このzipファイルを共有してください。")
    print("💡 受け取った人はPythonなしで直接EXEを実行できます！")

if __name__ == "__main__":
    create_exe_zip() 