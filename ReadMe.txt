■　はじめに

「AACS」は部活などの出欠確認をサポートするソフトです。

「Register.db」に登録されている名前またはIDを入力することで自動的にExcelに出欠を記録できます。

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

・「AACS.zip」を解凍してください。

--------------------
■　アップデート

・更新が必要な場合は起動時に自動的に更新案内が出ます。

アップデート用のファイルはGithubにあります。
URL: https://github.com/mikannohako/AACS-Automatic-Attendance-Confirmation-System/releases
（公開されてない場合は404Not Foundと表示されることがあります。）

--------------------
■　アンインストール

・レジストリなどの変更はないので関連ファイルを削除するだけでアンインストールできます。

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

2. 設定から顔認識を有効化してください。
3. 再起動しますかと問われるのでOKを押して再起動してください。
※対象の人物の名前は「user_list.xlsx」に設定されている名前と同じ名前にしてください。

--------------------
■　免責事項

Disclaimer: This software is provided "as is", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software.

--------------------
■　著作権

Copyright (c) 2024 mikannohako. All rights reserved.