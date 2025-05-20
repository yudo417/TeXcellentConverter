# TexcellentConverter

TexcellentConverterは、Excel表をLaTeXコードに変換したり、TikZグラフをGUIで作成できるPCアプリです．

## 特徴
- Excelファイルから簡単にLaTeX表を生成
- TikZ形式のグラフをGUIで作成・編集

## スクリーンショット
![screenshot](screenshots/sample.png)

## インストール方法
TeXcellentConverterのインストール方法を説明します．

ここでは，インストールするバージョンを ``{version}`` と表記しています．

### Windows版
1. [リリースページ](https://github.com/yudo417/TexcellentConverter/releases)から最新の``TeXcellentConverter-{version}-win-x64.zip``をダウンロードしてください．
2. ダウンロードしたZIPファイルを任意の場所に解凍してください．
3. フォルダ内の``TeXcellentConverter-{version}-win-x64.exe``をダブルクリックすると起動できます．

### macOS版
1. [リリースページ](https://github.com/yudo417/TexcellentConverter/releases)から最新の``TeXcellentConverter-{version}-macOS-arm64.zip``をダウンロード
2. ダウンロードしたZIPファイルを任意の場所に解凍してください．
3. 解凍した``TexcellentConverter.app``をダブルクリックすると起動できます．

## 使い方

### 1. アプリの起動
解凍したフォルダから「TexcellentConverter.exe」（Windowsの場合）または「TexcellentConverter.app」（macOSの場合）をダブルクリックして起動します．

---

### 2. Excel表をLaTeXに変換する手順

1. **Excelファイルの選択**  
　画面上部の「Excelファイル」欄の右にある「参照...」ボタンをクリックし、変換したいExcelファイル（.xlsxまたは.xls）を選択します．

2. **シート名の選択**  
　ファイルを選択すると、自動的にシート名がプルダウンに表示されます．変換したいシートを選びます．

3. **セル範囲の指定**  
　「セル範囲」欄に、変換したい範囲（例：A1:E6）のように入力します．

4. **オプションの設定（任意）**  
　- 表のキャプションやラベルを入力できます．
　- 表の位置（h, t, bなど）を選べます．
　- 「数式の代わりに値を表示」「空の行と列を削除」「数式を$で囲む」「罫線を追加」などのチェックボックスで細かい設定が可能です．

5. **LaTeXコードの生成**  
　「LaTeXに変換」ボタンをクリックすると、下部にLaTeXコードが表示されます．

6. **LaTeXコードのコピー**  
　「クリップボードにコピー」ボタンを押すと、生成されたLaTeXコードをワンクリックでコピーできます．

---

### 3. LaTeX表の利用方法

- 生成されたLaTeXコードを、TeXエディタやOverleafなどに貼り付けて利用できます．
- 注意書きに従い、必要なパッケージ（例：`\usepackage{multirow}`）をプリアンブルに追加してください．

---

### 4. その他の機能

- **パッケージ名のコピー**  
　画面上部の「コピー」ボタンで、`\usepackage{multirow}`をクリップボードにコピーできます．

- **エラー表示**  
　ファイル選択や範囲指定に誤りがある場合は、エラーメッセージが表示されます．

## 動作環境
- Windows 10以降
- macOS 10.13 (High Sierra) 以降
- Python, pip等のインストールは不要です（アプリに全て同梱されています）

## ライセンス
MIT License


