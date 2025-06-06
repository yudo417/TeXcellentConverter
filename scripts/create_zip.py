#!/usr/bin/env python3

import os
import shutil
import zipfile
from pathlib import Path

def create_exe_zip():
    
    print("TeXcellentConverter win版zipファイル作成")
    
    build_dir = Path("build/exe.win-amd64-3.12")
    if not build_dir.exists():
        print("EXEファイルが見つからない")
        return
    
    #! zipファイル名
    zip_filename = "TeXcellentConverter-v1.0.0-win-x64.zip"
    
    print("zipファイルを作成中")
    
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
    readme_content = """# TeXcellentConverter windows版

## 使用方法
1. このzipファイルを解凍
2. `TeXcellentConverter.exe` をダブルクリックして実行
"""
    
    with zipfile.ZipFile(zip_filename, 'a') as zipf:
        zipf.writestr("TeXcellentConverter/README.txt", readme_content)
    
    print(f"作成完了: {zip_filename}")
    print(f"ファイルサイズ: {os.path.getsize(zip_filename) / 1024 / 1024:.1f} MB")

if __name__ == "__main__":
    create_exe_zip() 