#!/bin/bash
# リリース用アプリケーションのビルド

echo "TexcellentConverter 配布用ビルド開始..."

# クリーンアップ
echo "古いビルドファイルを削除中..."
rm -rf build dist
# 古いZIPファイルも削除
rm -f dist/TexcellentConverter_macOS.zip

# 必要なライブラリの確認（必要時にコメントを外す）
# echo "依存ライブラリを確認中..."
# pip install -r requirements.txt

# PyInstallerでビルド
echo "アプリケーションをビルド中..."
pyinstaller --windowed --name TexcellentConverter main.py

# ビルド確認
if [ ! -d "dist/TexcellentConverter.app" ]; then
    echo "エラー: ビルドに失敗しました。dist/TexcellentConverter.appが見つかりません。"
    exit 1
fi

# ZIP作成
echo "配布用ZIPファイルを作成中..."
cd dist
zip -r TexcellentConverter_macOS.zip TexcellentConverter.app
cd ..

# 完了メッセージとファイルサイズの表示
if [ -f "dist/TexcellentConverter_macOS.zip" ]; then
    SIZE=$(du -h "dist/TexcellentConverter_macOS.zip" | cut -f1)
    echo "ビルド完了! 配布ファイル: dist/TexcellentConverter_macOS.zip (サイズ: $SIZE)"
    
    # ZIPファイルができたら元の.appを削除（容量節約）
    echo "不要な.appファイルを削除中..."
    rm -rf dist/TexcellentConverter.app
    
    echo "次のリリースに備えました！"
else
    echo "警告: ZIPファイルの作成に問題がありました。"
fi 