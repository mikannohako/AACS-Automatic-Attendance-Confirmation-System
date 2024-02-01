#? インポート
import PySimpleGUI as sg
import sys
import os
import subprocess

# GUI画面のレイアウト
layout = [
    [sg.Text("起動する機能を選んでください。", font=("Helvetica", 15), justification='center')],  # カンマを追加
    [sg.Button('通常出席', bind_return_key=True, font=("Helvetica", 15)),
        sg.Button('Excel整理', bind_return_key=True, font=("Helvetica", 15)),
        sg.Button('初期起動', bind_return_key=True, font=("Helvetica", 15)),
        sg.Button('終了', bind_return_key=True, font=("Helvetica", 15))]
]

window = sg.Window('出席処理', layout, finalize=True)

while True:  #? 無限ループ
    event, values = window.read()
    
    if event == sg.WIN_CLOSED or event == '終了':  # Xボタンが押されたか、'終了'ボタンが押された場合
        window.close()
        sys.exit(0)
        
    if event == '通常出席':
        window.close()
        subprocess.run(["python", "GeneralAttendance.py"])
        sys.exit(0)
    
    if event == 'Excel整理':
        window.close()
        subprocess.run(["python", "ExcelClean.py"])
        sys.exit(0)
    
    if event == '初期起動':
        window.close()
        subprocess.run(["python", "SpecialAttendance.py"])
        sys.exit(0)