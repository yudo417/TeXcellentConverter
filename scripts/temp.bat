@echo off
setlocal

set "name=src\app.py"

echo TexcellentConverter %name% running

python src\%name%

if %errorlevel% equ 0 (
    echo 正常に終了
) else (
    echo エラー．終了コード: %errorlevel%
)

pause 