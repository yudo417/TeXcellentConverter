import sys
import os
from cx_Freeze import setup, Executable

try:
    import PyQt5
    pyqt5_path = os.path.dirname(PyQt5.__file__)
    qt_plugins_path = os.path.join(pyqt5_path, "Qt5", "plugins")
except ImportError:
    pyqt5_path = ""
    qt_plugins_path = ""

if "src" not in sys.path:
    sys.path.insert(0, "src")

build_exe_options = {
    "packages": [
        "PyQt5.QtCore", 
        "PyQt5.QtGui", 
        "PyQt5.QtWidgets",
        "urllib", 
        "urllib.parse", 
        "pathlib", 
        "os", 
        "sys", 
        "platform",
        "json"
    ],
    "include_files": [
        ("assets/", "assets/"),
        ("src/", "lib/"),  # srcディレクトリ全体をlibにコピー
    ],
    "excludes": [
        "tkinter", 
        "unittest", 
        "test",
        "distutils"
    ],
    "zip_include_packages": "*",
    "zip_exclude_packages": []
}

if qt_plugins_path and os.path.exists(qt_plugins_path):
    platforms_path = os.path.join(qt_plugins_path, "platforms")
    if os.path.exists(platforms_path):
        build_exe_options["include_files"].append((platforms_path, "platforms/"))
        print(f"✅ Qtプラットフォームプラグインを追加: {platforms_path}")
    else:
        print("⚠️ Qtプラットフォームプラグインが見つかりません")
else:
    print("⚠️ PyQt5パスが見つかりません")

# EXEファイルの設定
base = None
if sys.platform == "win32": # sys.platformは64bitでもwin32となる
    base = "Win32GUI"  

target = Executable(
    script="src/app.py",
    base=base,
    target_name="TeXcellentConverter.exe",
    icon="assets/icon/favicon.ico",
)

setup(
    name="TeXcellentConverter",
    version="1.0.0",
    description="TeXcel書出しアプリケーション",
    options={"build_exe": build_exe_options},
    executables=[target],
) 