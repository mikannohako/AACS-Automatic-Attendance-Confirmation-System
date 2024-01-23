import PySimpleGUI as sg

from tkinter import messagebox
import sqlite3
import openpyxl

# 新しいWorkbook（エクセルファイル）を作成し、ファイル名を指定
workbook = openpyxl.Workbook()
sheet = workbook.active

sheet['A1'] = 'ID'
sheet['B1'] = '名前'
sheet['C1'] = '学年'
sheet['D1'] = '出席状況'

conn = sqlite3.connect('Register.db')
cursor = conn.cursor()

cursor.execute('SELECT ID, Name, GradeinSchool FROM Register')
all_data = cursor.fetchall()

# Process each data and update the Excel sheet
for data in all_data:
    row_number = sheet.max_row + 1
    sheet.cell(row=row_number, column=1, value=data[0])  # ID
    sheet.cell(row=row_number, column=2, value=data[1])  # 名前
    sheet.cell(row=row_number, column=3, value=data[2])  # 学年
    sheet.cell(row=row_number, column=4, value='未出席')  # 初めは未出席として設定

workbook.save('temp.xlsx')

layout = [
    [sg.Text('名前を入力してください。', key="-INPUT-")],
    [sg.InputText(key="-NAME-")],
    [sg.Button('OK', bind_return_key=True)]
]

window = sg.Window('log in', layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        conn.close()
        break
    
    if event == 'OK' or event == 'Escape:13':
        name = values["-NAME-"]
        cursor.execute('SELECT ID, GradeinSchool FROM Register WHERE Name = ?', (name,))
        result = cursor.fetchone()
        
        print('入力された値：', name)
        
        if result:
            print(f"{name} はデータベースに存在します。")
            
            # 名前が一致する行を探し、出席を記録
            for row in range(1, sheet.max_row + 1):
                if sheet.cell(row=row, column=2).value == name:
                    sheet.cell(row=row, column=4, value='出席')
                    workbook.save('temp.xlsx')
                    messagebox.showinfo('完了', f'{name} の出席処理は完了しました。')
                    window["-NAME-"].update("")  # 入力フィールドをクリア
                    conn.commit()  # 変更を確定
                    break
        else:
            print(f"{name} はデータベースに存在しません。")
            messagebox.showinfo('失敗', f'{name} はデータベースに存在しません。')
            window["-NAME-"].update("")  # 入力フィールドをクリア