■　はじめに

「AACS」は部活などの出欠確認をサポートするソフトです。

「user_list.xlsx」に登録されている名前またはIDを入力することで自動的にExcelに出欠を記録できます。

--------------------
■　主な機能

・出席や欠席の情報をExcelに自動で記録

--------------------
■　動作環境

対応OS

・Windows10
・Windows11

--------------------
■　インストール

1. Python公式サイトから最新のPython 3.12.x（推奨）をダウンロードし、PATHのチェックボックスを有効にしてインストールしてください。  
2. 「launch.vbs」を実行してください。
3. ソフトが正常に起動したら完了です。

※初回起動時は必ずネットワークに接続してください。
※初回起動時は起動に時間がかかる可能性があります。

--------------------
■　アップデート

・更新が必要な場合は起動時に自動的に更新案内が出ます。

アップデート用のファイルはGithubにあります。
URL: https://github.com/mikannohako/AACS-Automatic-Attendance-Confirmation-System/releases

--------------------
■　アンインストール

1. 完全にアンインストールしたい場合はすべて、記録を残したい場合は「Attendance_records.xlsx」を別の場所に保存してください。

--------------------
■　使用方法

（初めての場合）
1. 「user_list.xlsx」を開きID（一列目）と名前（二列目）を設定してください。
2. 「launch.vbs」を実行してください。
3. メニュー画面が出るので使用する機能を選んでください
※起動に時間がかかる場合があります。

（二回目以降）
1. 「launch.vbs」を実行してください。
2. メニュー画面が出るので使用する機能を選んでください

（顔認識を使用する場合）
1. 「face_images」の中に、顔認識対象の人物の名前でフォルダを作成し、そのフォルダ内に以下の顔写真を3枚ほど入れてください。
 ・正面の顔
 ・左斜めの顔
 ・右斜めの顔
例）
face_images/
  ├── 田中/
  │    ├── front.jpg
  │    ├── left.jpg
  │    └── right.jpg
  └── 佐藤/
        ├── front.jpg
        ├── left.jpg
        └── right.jpg

2. 設定から顔認識を有効化してください。
3. 「再起動しますか？」と問われるのでOKを押して再起動してください。
※対象の人物の名前は「user_list.xlsx」に設定されている名前と同じ名前にしてください。

（その他）
・記録は「Attendance_records.xlsx」に保存されています。

--------------------
■　ライセンス・著作権

Copyright (C) 2025 mikannohako

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

このソフトウェアはGNU General Public License (GPL) に基づき、自由に再配布・改変することができます。ただし、使用に伴う損害等については一切の保証を行いません。詳細はライセンス条文をご覧ください。

--------------------
■　その他

上記以外の詳細な情報はGithubで確認することができます。
URL: https://github.com/mikannohako/AACS-Automatic-Attendance-Confirmation-System