@echo off
setlocal enabledelayedexpansion

echo TeXcellentConverter Release Build Script
echo ==========================================

echo [1/5] 依存関係をチェック中...
python -c "import PyQt5; print('PyQt5: OK')" 2>nul
if errorlevel 1 (
    echo エラー: PyQt5がインストールされていません
    echo pip install PyQt5 を実行してください
    pause
    exit /b 1
)

python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo エラー: PyInstallerがインストールされていません  
    echo pip install pyinstaller を実行してください
    pause
    exit /b 1
)

echo [2/5] 古いビルドファイルを削除中...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo [3/5] アイコンファイルを確認中...
set ICON_PARAM=
if exist "assets\icon\favicon.ico" (
    set ICON_PARAM=--icon=assets\icon\favicon.ico
    echo アイコンファイル発見: assets\icon\favicon.ico
) else (
    echo 警告: アイコンファイルが見つかりません
)

echo [4/5] PyInstallerでビルド中...
echo コマンド: python -m PyInstaller --onedir --windowed !ICON_PARAM! --name TeXcellentConverter --add-data "assets;assets" src\app.py

python -m PyInstaller --onedir --windowed !ICON_PARAM! --name TeXcellentConverter --add-data "assets;assets" src\app.py

if not exist "dist\TeXcellentConverter\TeXcellentConverter.exe" (
    echo エラー: ビルドに失敗しました
    echo 予想されるファイル: dist\TeXcellentConverter\TeXcellentConverter.exe
    echo.
    echo 代替案: 以下のコマンドを手動で実行してください
    echo python -m PyInstaller --onedir --windowed !ICON_PARAM! --name TeXcellentConverter src\app.py
    pause
    exit /b 1
)

echo [5/5] ビルド結果を確認中...
for %%A in ("dist\TeXcellentConverter\TeXcellentConverter.exe") do set EXE_SIZE=%%~zA
echo 実行ファイル作成完了: !EXE_SIZE! bytes

echo.
echo ZIPファイルを作成しますか？ (Y/N)
set /p CREATE_ZIP=
if /i "!CREATE_ZIP!"=="Y" (
    echo ZIPファイル生成中...
    cd dist
    powershell -Command "Compress-Archive -Path TeXcellentConverter -DestinationPath TeXcellentConverter-v1.0.0-Windows-x64.zip -Force"
    cd ..
    
    if exist "dist\TeXcellentConverter-v1.0.0-Windows-x64.zip" (
        for %%A in ("dist\TeXcellentConverter-v1.0.0-Windows-x64.zip") do set ZIP_SIZE=%%~zA
        echo ZIPファイル作成完了: !ZIP_SIZE! bytes
    )
)

echo.
echo ==========================================
echo ビルド完了!
echo ==========================================
echo 実行ファイル: dist\TeXcellentConverter\TeXcellentConverter.exe
if exist "dist\TeXcellentConverter-v1.0.0-Windows-x64.zip" (
    echo 配布ファイル: dist\TeXcellentConverter-v1.0.0-Windows-x64.zip
)
echo ==========================================
echo.
echo テスト実行しますか？ (Y/N)
set /p TEST_RUN=
if /i "!TEST_RUN!"=="Y" (
    echo アプリケーションを起動中...
    start "" "dist\TeXcellentConverter\TeXcellentConverter.exe"
)

pause 