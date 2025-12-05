# QR Code Generator YMCK

**QR Code Generator YMCK** は、URLからQRコードを生成するためのWindows用デスクトップアプリケーションです。
直感的なGUI操作で、Web用（JPEG）および印刷用（EPS/CMYK）のQRコードを「単体」または「Excelリストから一括」で作成できます。

## 概要 (Description)

このツールは、一般的なQRコード作成機能に加え、DTPや印刷業務での利用を想定した **CMYKカラーモードのEPS形式** 出力をサポートしている点が特徴です。
Pythonの `customtkinter` を使用したモダンなUIを採用しており、エンジニアでなくても簡単に操作可能です。

## 特徴 (Features)

  * **単体作成モード:** URLとファイル名を入力して素早くQRコードを作成。
  * **一括作成モード:** Excelファイルを読み込み、大量のQRコードを一度に生成。
  * **マルチフォーマット対応:**
      * **JPEG:** Webや画面表示用（RGB）。
      * **EPS:** 印刷入稿用（CMYK変換済み）。
      * **両方:** JPEGとEPSを同時に出力。
  * **保存場所:** デスクトップの `QR-Code` フォルダに自動保存。

## 動作環境 (Requirement)

  * **OS:** Windows 10 / 11
  * **Python:** Python 3.x (ソースコードから実行する場合)

## インストールと実行方法 (Installation & Usage)

このツールは、配布された実行ファイル（.exe）を使用するか、Python環境でソースコードを実行するかの2通りの方法で使用できます。

### A. 実行ファイル（.exe）を使用する場合

1.  配布された `QRコード作成ツール.exe`（または任意の名称）をダブルクリックして起動します。
2.  必要なライブラリやPython環境のインストールは不要です。

### B. Python環境でソースコードから実行する場合

開発者向けの手順です。

1.  **リポジトリのクローンまたはダウンロード**

2.  **依存ライブラリのインストール**
    以下のコマンドで必要なパッケージをインストールしてください。

    ```bash
    pip install customtkinter qrcode openpyxl pillow
    ```

    ※ `tkinter` はPython標準ライブラリに含まれています。

3.  **アプリケーションの起動**

    ```bash
    python app_v2.py
    ```

## 使い方 (Usage)

### 1\. 単体作成 (Single Mode)

1.  **「単体作成」** タブを選択します。
2.  **【URL】** にQRコード化したいアドレスを入力します。
3.  **【ファイル名】** に出力ファイル名を入力します（拡張子は不要）。
4.  **【ファイル形式】** を選択します（JPEG / EPS / 両方）。
5.  **「作成」** ボタンを押すと、デスクトップの `QR-Code` フォルダに保存されます。

### 2\. 一括作成 (Bulk Mode)

Excelファイルを用意することで、複数のQRコードをまとめて作成できます。

#### Excelファイルのフォーマット

1行目をヘッダーとし、以下の列構成で作成してください。

| 行番号 | A列 (ファイル名) | B列 (URL) |
| :--- | :--- | :--- |
| **1** | **ファイル名** | **URL** |
| 2 | google\_qr | [https://google.com](https://google.com) |
| 3 | yahoo\_qr | [https://yahoo.co.jp](https://yahoo.co.jp) |
| ... | ... | ... |

#### 手順

1.  **「一括作成」** タブを選択します。
2.  **「選択」** ボタンを押し、用意したExcelファイル（.xlsx）を指定します。
3.  **【ファイル形式】** を選択します。
4.  **「作成」** ボタンを押すと、一括処理が開始されます。

## ビルド方法 (Build)

PyInstallerを使用してexe化する場合のコマンドは以下の通りです。
※環境に合わせてパス（`add-data`部分）を調整してください。

```bash
pyinstaller app_v2.py --onefile --noconsole --add-data "path\to\site-packages\customtkinter;customtkinter"
```

## 使用ライブラリ (Dependencies)

  * [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - UIフレームワーク
  * [qrcode](https://pypi.org/project/qrcode/) - QRコード生成
  * [openpyxl](https://openpyxl.readthedocs.io/) - Excel操作
  * [Pillow (PIL)](https://www.google.com/search?q=https://python-pillow.org/) - 画像処理・CMYK変換

## ライセンス (License)

本プロジェクトのライセンス情報の記載はありません。
