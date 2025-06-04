@echo off
setlocal

echo TeXcellentConverter dev mode running

python src\app.py

if %errorlevel% equ 0 (
    echo 正常に終了
) else (
    echo エラー．終了コード: %errorlevel%
)

pause 