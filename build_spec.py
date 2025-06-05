# PyInstaller用の設定ファイル
# 使用方法: pyinstaller build_spec.py

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# アプリケーションの基本情報
app_name = 'TeXcellentConverter'
main_script = os.path.join('src', 'app.py')

# データファイルの収集
datas = []

# PyQt5のプラグインを含める
pyqt5_datas = collect_data_files('PyQt5.Qt5.plugins', include_py_files=False)
datas.extend(pyqt5_datas)

# アセットディレクトリがあれば追加
if os.path.exists('assets'):
    datas.append(('assets', 'assets'))

# 隠しインポート（必要に応じて追加）
hiddenimports = []
hiddenimports.extend(collect_submodules('PyQt5'))

# プラットフォーム固有の設定
if sys.platform == 'win32':
    # Windows用設定
    icon = 'assets/icon.ico' if os.path.exists('assets/icon.ico') else None
    console = False
elif sys.platform == 'darwin':
    # macOS用設定
    icon = 'assets/icon.icns' if os.path.exists('assets/icon.icns') else None
    console = False
else:
    # Linux用設定
    icon = None
    console = False

# PyInstaller設定
a = Analysis(
    [main_script],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=console,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)

# macOS用のAppバンドル作成
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name=f'{app_name}.app',
        icon=icon,
        bundle_identifier=f'com.example.{app_name.lower()}',
        info_plist={
            'CFBundleName': app_name,
            'CFBundleDisplayName': app_name,
            'CFBundleVersion': '1.0.0',
            'CFBundleShortVersionString': '1.0.0',
            'NSHighResolutionCapable': True,
        },
    ) 