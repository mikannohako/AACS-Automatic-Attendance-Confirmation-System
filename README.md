# 目次

- [目次](#目次)
  - [このソフトについて](#このソフトについて)
  - [機能](#機能)
  - [インストール方法](#インストール方法)
  - [アンインストール方法](#アンインストール方法)
  - [開発に関すること](#開発に関すること)
    - [開発環境](#開発環境)
    - [使用ライブラリ](#使用ライブラリ)
      - [外部ライブラリ](#外部ライブラリ)
      - [標準ライブラリ](#標準ライブラリ)
    - [VSCode使用拡張機能](#vscode使用拡張機能)
      - [おそらく必須](#おそらく必須)
      - [あるとコードが見やすくなる](#あるとコードが見やすくなる)

## このソフトについて

Python初心者が練習などのために作ったソフトです。  
出席の管理などができます。

- 初心者が書いた低クオリティなコード
- ぐちゃぐちゃでカオスなコード

上記のことを許せる方はどうぞ

## 機能

- Excelに出席データを記録
- 顔認証
- 出席率の記録
- 出席率のファイル出力　（予定）

## インストール方法

1. Python公式サイトから最新のPython 3.12.x（推奨）をダウンロードし、インストールしてください。  
   [Python公式ダウンロードページ](https://www.python.org/downloads/)
2. このレポジトリの[最新のReleases](https://github.com/mikannohako/AACS-Automatic-Attendance-Confirmation-System/releases/latest)から`AACS.zip`をダウンロード
3. 任意のフォルダに解凍
4. `launch.vbs`を実行
5. 機能を選択

※初回起動時は起動に時間がかかる可能性があります。

## アンインストール方法

1. インストールされているフォルダごと消してください
2. （任意）インストール時にインストールしたPythonをアンインストール

## 開発に関すること

### 開発環境

- Windows 10 VSCode
- Python 3.12.5

### 使用ライブラリ

ライブラリのバージョンなどの詳しい情報は`requirements.txt`を参照してください。

#### 外部ライブラリ

- PySimpleGUI
- openpyxl
- requests
- cv2
- face_recognition
- numpy
- pickle

#### 標準ライブラリ

- sys
- os
- datetime
- time
- json
- tkinter（及びそのサブモジュール）
  - simpledialog
  - messagebox
- logging
- hashlib
- tempfile
- msvcrt
- webbrowser

### VSCode使用拡張機能

#### おそらく必須

- [Python](https://marketplace.visualstudio.com/items?itemName=ms-python.python)
- [Pylance](https://marketplace.visualstudio.com/items?itemName=ms-python.vscode-pylance)
- [PySimpleGUI Snippets](https://marketplace.visualstudio.com/items?itemName=Acezx.pysimplegui-snippets)

#### あるとコードが見やすくなる

- [Better Comments](https://marketplace.visualstudio.com/items?itemName=aaron-bond.better-comments)
- [indent-rainbow](https://marketplace.visualstudio.com/items?itemName=oderwat.indent-rainbow)
