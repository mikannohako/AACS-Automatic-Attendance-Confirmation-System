# Automatic attendance confirmation system

import PySimpleGUI as sg
from tkinter import messagebox

layout = [  [sg.Text('名前を入力してください。')],
            [sg.InputText()],
            [sg.Button('OK')] ]

window = sg.Window('log in', layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    
    if event == 'OK':
        print('あなたが入力した値： ', values[0])
        name = values[0]

        if name == 'test':
            messagebox.showinfo('完了', f'{name} の出席処理は完了しました。')