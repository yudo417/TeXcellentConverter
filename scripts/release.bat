@echo off
setlocal enabledelayedexpansion

echo TeXcellentConverter Release Build Script (Windows)

python -c "import PyQt5; print('PyQt5: OK')" 2>nul
if errorlevel 1 (
    echo エラー: PyQt5が見つからない
    pause
    exit /b 1
)

python -c "import cx_Freeze; print('cx_Freeze: OK')" 2>nul
if errorlevel 1 (
    echo エラー: cx_Freezeが見つからない
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

if exist "assets\icon\favicon.ico" (
    echo アイコンファイル確認: OK
) else (
    echo 警告: アイコンファイルが見つからない
)

REM cx_Freezeでbuildが必要
python scripts\create_exe.py build

if not exist "build\exe.win-amd64-3.12\TeXcellentConverter.exe" (
    echo エラー: ビルド失敗，exeの生成が失敗
    pause
    exit /b 1
)

for %%A in ("build\exe.win-amd64-3.12\TeXcellentConverter.exe") do set EXE_SIZE=%%~zA
echo exe作成完了: !EXE_SIZE! bytes

python scripts\create_zip.py

if exist "TeXcellentConverter_EXE.zip" (
    for %%A in ("TeXcellentConverter_EXE.zip") do set ZIP_SIZE=%%~zA
    echo ZIPファイル作成完了: !ZIP_SIZE! bytes
) else (
    echo 警告: ZIP作成に失敗
)

echo.
echo ビルド完了
echo 実行ファイル: build\exe.win-amd64-3.12\TeXcellentConverter.exe
if exist "TeXcellentConverter_EXE.zip" (
    echo 配布ファイル: TeXcellentConverter_EXE.zip
)


pause 