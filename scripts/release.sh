#!/bin/bash

echo "TeXcellentConverter release build"


echo "古いビルドファイルを削除"
rm -rf build dist


echo "リリース用ビルド"
pyinstaller --windowed --icon=assets/icon/favicon.ico --name TeXcellentConverter src/main.py

if [ ! -d "dist/TeXcellentConverter.app" ]; then
    echo "エラー: ビルドに失敗。dist/TeXcellentConverter.appが見つかりません。"
    exit 1
fi

echo "ZIPファイル生成"
cd dist
zip -r TeXcellentConverter-v1.0.0-macOS-arm64.zip TeXcellentConverter.app
cd ..

if [ -f "dist/TeXcellentConverter-v1.0.0-macOS-arm64.zip" ]; then
    SIZE=$(du -h "dist/TeXcellentConverter-v1.0.0-macOS-arm64.zip" | cut -f1)
    echo "ビルド終了: dist/TeXcellentConverter-v1.0.0-macOS-arm64.zip (サイズ: $SIZE)"

    echo ".appファイル削除"
    rm -rf dist/TeXcellentConverter.app
    
    # 削除確認
    if [ ! -d "dist/TeXcellentConverter.app" ]; then
        echo ".appファイル削除完了"
    else
        echo "警告: .appファイルの削除に失敗"
    fi
else
    echo "警告: ZIPファイル生成失敗。"
fi 