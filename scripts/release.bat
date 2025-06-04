@echo off
setlocal

echo TeXcellentConverter release build

echo 古いビルドファイルを削除
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo リリース用ビルド
pyinstaller --windowed --icon=assets\icon\favicon.ico --name TeXcellentConverter src\main.py

if not exist "dist\TeXcellentConverter.exe" (
    echo エラー: ビルドに失敗。dist\TeXcellentConverter.exeが見つかりません。
    pause
    exit /b 1
)

echo ZIPファイル生成
cd dist
powershell -Command "Compress-Archive -Path TeXcellentConverter.exe -DestinationPath TeXcellentConverter-v1.0.0-Windows-x64.zip -Force"
cd ..

if exist "dist\TeXcellentConverter-v1.0.0-Windows-x64.zip" (
    for %%A in ("dist\TeXcellentConverter-v1.0.0-Windows-x64.zip") do set SIZE=%%~zA
    echo ビルド終了: dist\TeXcellentConverter-v1.0.0-Windows-x64.zip (サイズ: !SIZE! bytes)

    echo .exeファイル削除
    del /q dist\TeXcellentConverter.exe
    
    if not exist "dist\TeXcellentConverter.exe" (
        echo .exeファイル削除完了
    ) else (
        echo 警告: .exeファイルの削除に失敗
    )
) else (
    echo 警告: ZIPファイル生成失敗。
)

pause 