#!/bin/bash

echo "TexcellentConverter release build"


echo "古いビルドファイルを削除"
rm -rf build dist


echo "リリース用ビルド"
pyinstaller --windowed --icon=favicon.ico --name TexcellentConverter main.py

if [ ! -d "dist/TexcellentConverter.app" ]; then
    echo "エラー: ビルドに失敗しました。dist/TexcellentConverter.appが見つかりません。"
    exit 1
fi

echo "ZIPファイル生成"
cd dist
zip -r TexcellentConverter_macOS.zip TexcellentConverter.app
cd ..

if [ -f "dist/TexcellentConverter_macOS.zip" ]; then
    SIZE=$(du -h "dist/TexcellentConverter_macOS.zip" | cut -f1)
    echo "ビルド終了: dist/TexcellentConverter_macOS.zip (サイズ: $SIZE)"

    echo ".appファイル削除"
    rm -rf dist/TexcellentConverter.app
    
    # 削除確認
    if [ ! -d "dist/TexcellentConverter.app" ]; then
        echo ".appファイル削除完了"
    else
        echo "警告: .appファイルの削除に失敗しました"
    fi
else
    echo "警告: ZIPファイル生成失敗。"
fi 