#!/usr/bin/env python3
"""
クロスプラットフォーム対応のビルドスクリプト
"""

import os
import sys
import subprocess
import shutil
import platform

def clean_build():
    """ビルドディレクトリをクリーンアップ"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for directory in dirs_to_clean:
        if os.path.exists(directory):
            print(f"Cleaning {directory}...")
            shutil.rmtree(directory)
    
    # .specファイルもクリーンアップ
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    for spec_file in spec_files:
        if spec_file != 'build_spec.py':
            print(f"Removing {spec_file}...")
            os.remove(spec_file)

def install_dependencies():
    """必要な依存関係をインストール"""
    requirements = ['pyinstaller', 'PyQt5']
    
    for package in requirements:
        print(f"Installing {package}...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', package], check=True)

def build_app():
    """アプリケーションをビルド"""
    current_platform = platform.system().lower()
    print(f"Building for {current_platform}...")
    
    # PyInstallerコマンドを構築
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onedir',  # ワンディレクトリ形式（依存関係を含む）
        '--windowed',  # コンソールウィンドウを表示しない
        '--name', 'TeXcellentConverter',
        '--add-data', 'assets;assets' if current_platform == 'windows' else 'assets:assets',
        'src/app.py'
    ]
    
    # プラットフォーム固有の設定
    if current_platform == 'windows':
        if os.path.exists('assets/icon.ico'):
            cmd.extend(['--icon', 'assets/icon.ico'])
    elif current_platform == 'darwin':
        if os.path.exists('assets/icon.icns'):
            cmd.extend(['--icon', 'assets/icon.icns'])
        # macOS用のapp bundleを作成
        cmd.append('--osx-bundle-identifier=com.example.texcellentconverter')
    
    # ビルド実行
    print("Running PyInstaller...")
    print(" ".join(cmd))
    subprocess.run(cmd, check=True)
    
    print(f"Build completed! Check the 'dist' directory.")

def main():
    """メイン関数"""
    print("TeXcellentConverter Build Script")
    print("=" * 40)
    
    # 引数処理
    if len(sys.argv) > 1 and sys.argv[1] == 'clean':
        clean_build()
        print("Clean completed.")
        return
    
    try:
        # 依存関係のインストール
        install_dependencies()
        
        # ビルドディレクトリのクリーンアップ
        clean_build()
        
        # アプリケーションのビルド
        build_app()
        
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nBuild cancelled by user.")
        sys.exit(1)

if __name__ == '__main__':
    main() 