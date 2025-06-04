# TeXcellentConverter

**Excelの表や採集したデータを，表やグラフを生成するLaTeXのコードに変換するソフトです．**



## 📋 特徴

TeXcellentConverterは、主に以下の2つの機能を提供しています

1. **Excel表のLaTeX変換**
   

https://github.com/user-attachments/assets/0dfe4b4f-033a-4a20-b570-0210c4ea2fae


   - Excelで作成した表をtabular形式のLaTeXコードに変換

3. **グラフ生成**
   

https://github.com/user-attachments/assets/e65eeb4d-19db-4e99-b3b5-3d9f344695c5


   - Excel/CSVデータや数式からTikZ形式の美しいグラフを生成
   - 手入力データにも対応

## 🛠️ 主要機能

### 📊 Excel → LaTeX表変換
- **結合セル対応**: 複雑な結合セルを含む表も正確に変換
- **範囲指定**: 必要な部分のみを指定して変換
- **位置調整**: 「H,h,t,b,p,htbp」から位置指定可能

### 📈 TikZグラフ生成
#### データ入力方法
- **実測データ**: CSV/Excelファイルまたは手動入力
- **数式グラフ**: 数学関数から描画
- **複数データセット**: 複数のグラフを同時に描画

#### 基本的なグラフ設定
- **グラフタイプ**: 線グラフ、散布図、線+点、棒グラフから選択可能
- **軸設定**: 通常の軸から片対数，両対数グラフにも対応
- **範囲**: 描画範囲を任意に設定可能

#### 高度な機能
- **特殊点**: 重要なデータポイントをハイライト表示
- **注釈**: グラフ上にテキストによる説明を追加
- **接線**: 数式グラフの指定点における接線を描画
- **座標表示**: 特殊点の座標を表示


## 💻 システム条件

| プラットフォーム | 対応バージョン 
|------------------|----------------
| **Windows** | Windows 10 (64bit) 以降 
| **macOS** | macOS 13.0 (Ventura) 以降

> **重要**:
> - **Windows版**: Windows 10 (64bit) では環境によって動作しない場合がございます．
> - **macOS版**: Apple シリコン搭載のMacを使用してください．


## 💿 インストール方法

[リリースページ](https://github.com/yudo417/TeXcellentConverter/releases)からインストールしてください．



## 📜 ライセンス

[GNU GENERAL PUBLIC LICENSE Version 3](https://github.com/yudo417/TeXcellentConverter/blob/main/LICENSE) 



