# Excel Search Tool

## 概要

指定した文字列を複数のExcelファイルから検索するためのPythonツールです。  
フォルダ内すべての `.xlsx` ファイルの全シート・全列を対象に検索を行い、該当箇所を表示します。  
動作環境はWindowsです。

## 使用方法

### 1. 環境構築
以下のコマンドでツールをインストールしてください。
```bash
git clone https://github.com/nao-webdev/excel-search-tool
```

以下のコマンドでpythonの依存ライブラリをインストールしてください。

```bash
cd excel-search-tool
pip install -r requirements.txt
```
### 2. ツールの起動
**start_search.bat** をダブルクリックしてください。


### 3. 検索手順
フォルダ選択ダイアログで検索対象フォルダを選択し、「フォルダの選択」ボタンを押下してください。

キーワード入力ダイアログで検索ワードを入力し、「OK」ボタンを押下してください。

該当箇所があれば、ファイル名・シート名・該当行がコンソールに表示されます。

## ファイル構成
```bash
excel-search-tool/
├── search.py           # 本体スクリプト
├── start_search.bat    # 起動バッチ
├── requirements.txt    # 必要なPythonライブラリ
└── README.md           # この説明ファイル
```
