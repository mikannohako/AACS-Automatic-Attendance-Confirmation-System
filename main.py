
#? インポート
import PySimpleGUI as sg
import sys
import os
import subprocess
from datetime import datetime
from tkinter import messagebox
import sqlite3
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import json
import cv2
import time

#? エラー時の処理の作成

def exit_with_error(message):
    print(f"Error: {message}")
    messagebox.showerror("Error:", message)
    sys.exit(1)  # アプリケーションをエラーコード 1 で終了します

# 時間変数の設定
current_date = datetime.now()
current_date_y = current_date.strftime("%Y")
current_date_m = current_date.strftime("%m")
current_date_d = current_date.strftime("%d")
current_date_h = current_date.strftime("%H")
current_date_M = current_date.strftime("%M")
current_date_s = current_date.strftime("%S")

#? 各機能の関数

def SApy(): #? SpecialAttendance.py
    # 月ごとのシートを作成する関数
    def create_month_sheet(workbook, month):
        sheet_name = month
        if month.startswith("0"):
            sheet_name = month[1:]
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)
            # 作成したシートを返す
            return workbook[sheet_name]
    
    # 日付データを入力する関数
    def input_date_data(sheet, month):
        days_in_month = 30 if month in ["04", "06", "09", "11"] else 31 if month != "02" else 29  # 月ごとの日数
        for day in range(1, days_in_month + 1):
            sheet.cell(row=1, column=day + 2).value = day
        # 学年の列を追加
        sheet.cell(row=1, column=1).value = '名前'
        sheet.cell(row=1, column=2).value = '学年'
    
    # 新しいWorkbook（エクセルファイル）を作成して、ファイル名を指定
    workbook = openpyxl.Workbook()
    
    # DBに接続する
    conn = sqlite3.connect('Register.db')
    cursor = conn.cursor()
    
    # すべてのデータを取得
    cursor.execute('SELECT Name, GradeinSchool FROM Register')
    all_data = cursor.fetchall()
    
    # 月ごとのシートを作成して名前と学年の行を追加
    for month in range(1, 13):
        month_str = str(month).zfill(2)  # 1桁の月を2桁の文字列に変換
        sheet = create_month_sheet(workbook, f"{month_str}月")
        input_date_data(sheet, month_str)
        
        # エクセルにすべてのデータを入力
        for data in all_data:
            row_number = sheet.max_row + 1
            sheet.cell(row=row_number, column=1, value=data[0])  # 名前
            sheet.cell(row=row_number, column=2, value=data[1])  # 学年
        
        end_row = len(all_data) + 1 # データの数に基づいて終了行を決定する
        table_range = f"A1:B{end_row}"  # 範囲を変数に格納する
        table = openpyxl.worksheet.table.Table(displayName=f"Table{month}", ref=table_range)
        
        # スタイル設定
        style = TableStyleInfo(
            name="TableStyleMedium9",  # テーブルのスタイル名
            showRowStripes=True,    # 行のストライプを表示する
        )
        
        # テーブルにスタイルを適用
        table.tableStyleInfo = style
        
        # シートにテーブルを追加
        sheet.add_table(table)
    
    # 削除したいシート名を指定
    sheet_name_to_delete = "Sheet"
    
    # シートを削除
    if sheet_name_to_delete in workbook.sheetnames:
        sheet_to_delete = workbook[sheet_name_to_delete]
        workbook.remove(sheet_to_delete)
        print("一時シート削除>>>正常終了")
    else:
        print(f"{sheet_name_to_delete} シートは存在しません。")
    
    # エクセルファイルを保存
    workbook.save(f"{current_date_y}Attendance records.xlsx")

def GApy(): #? GeneralAttendance.py & ExcelClean.py
    #? config設定
    
    # 時間変数の設定
    current_date = datetime.now()
    current_date_y = current_date.strftime("%Y")
    current_date_m = current_date.strftime("%m")
    current_date_d = current_date.strftime("%d")
    current_date_h = current_date.strftime("%H")
    current_date_M = current_date.strftime("%M")
    current_date_s = current_date.strftime("%S")
    
    # 設定ファイルのパス
    config_file_path = 'config.json'
    
    # 設定ファイルの読み込み
    with open(config_file_path, 'r') as config_file:
        config_data = json.load(config_file)
        
    #? 変数の初期設定
    
    name_column_index = None
    date_row_index = None
    name = None
    
    # 記録ファイル名
    ar_filename = f"{current_date_y}Attendance records.xlsx"
    
    #? Excel初期設定
    
    # 新しいWorkbook（エクセルファイル）を作成して、ファイル名を指定
    try:
        workbook = load_workbook(ar_filename)
    except FileNotFoundError:
        messagebox.showerror("Error: File not found", "記録用のファイルがありません。\n初期起動を行ってください。")
        subprocess.run(["python", "Menu.py"])
        sys.exit(0)
    
    # アクティブなシートを開く
    temp_sheet = workbook.create_sheet("temp")  
    day_int = int(current_date_m)
    sheet = workbook[f"{day_int}月"]
    
    # 項目の作成
    temp_sheet['A1'] = '名前'
    temp_sheet['B1'] = '学年'
    temp_sheet['C1'] = '出席状況'
    
    #? DB初期設定
    
    # DBに接続する
    
    conn = sqlite3.connect('Register.db')
    cursor = conn.cursor()    
    
    
    # すべてのデータを取得
    cursor.execute('SELECT Name, GradeinSchool FROM Register')
    all_data = cursor.fetchall()
    
    #? 一時ファイルにDBを入力
    
    # エクセルにすべてのデータを入力
    for data in all_data:
        row_number = temp_sheet.max_row + 1
        temp_sheet.cell(row=row_number, column=1, value=data[0])  # 名前
        temp_sheet.cell(row=row_number, column=2, value=data[1])  # 学年
    
    end_row = len(all_data) + 1 # データの数に基づいて終了行を決定する
    for row_number in range(2, end_row + 1): # 2からend_row + 1までの行数を繰り返し処理
        temp_sheet.cell(row=row_number, column=3, value='欠席') # 初めは欠席として設定
    
    
    # 保存
    workbook.save(ar_filename)
    
    information = '記録なし'
    
    # last_qr_dataの初期化
    last_qr_data = None
    
    capbool = False
    
    def mainwindowshow(): #? メインウィンドウ表示
        
        # シートから欠席のデータを取得してリストに格納
        data = []
        
        for row in temp_sheet.iter_rows(values_only=True):
            if row[2] == '欠席':  # 欠席のデータのみを抽出 
                modified_row = list(row)
                modified_row[2] = '未出席'
                data.append(modified_row)
        
        # ヘッダーを取得
        header = list(temp_sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
        
        left_column = [
            [sg.Text(information, font=("Helvetica", 40))],
            [sg.Text('苗字を入力してください。', key="-INPUT-", font=("Helvetica", 15))],
            [sg.InputText(key="-NAME-", font=("Helvetica", 15))],
            [sg.Button('OK', bind_return_key=True, font=("Helvetica", 15)), sg.Button('終了', bind_return_key=True, font=("Helvetica", 15))],
            [sg.Table(values=data, headings=header, display_row_numbers=False, auto_size_columns=False, num_rows=min(20, len(data)))]
        ]
        
        # GUI画面のレイアウト
        layout = [
            [sg.Column(left_column)]
        ]
        
        window = sg.Window('出席処理', layout, finalize=True)
        window.Maximize()
        return window  # window変数を返す
    
    def get_name_by_id(input_id): #? IDから名前を取得
        # IDに対応するnameをクエリで検索
        cursor.execute("SELECT Name FROM Register WHERE ID=?", (input_id,))
        result = cursor.fetchone()  # 一致する最初の行を取得
        if result:
            return result[0]  # nameを返す
        else:
            return "IDに対応する名前が見つかりません"
    
    window = mainwindowshow()  # mainwindowshow()関数を呼び出して、window変数に格納する
    
    while True: #? 無限ループ
        # イベントとデータの読み込み
        event, values = window.read()
        
        #? ウィンドウを閉じる時の処理
        if event == sg.WIN_CLOSED:
            messagebox.showinfo('警告', 'windowを閉じるのは「終了」ボタンから行ってください。')
            window.close()
            window = mainwindowshow()
        
        #? OKが押されたときの処理
        if event == 'OK' or event == 'Escape:13' or capbool:
            name = values["-NAME-"]
            
            print("id: ", name)
            
            if name.isdigit():
                name = get_name_by_id(int(name))
                result = name
            else:
                cursor.execute('SELECT GradeinSchool FROM Register WHERE Name = ?', (name,))
                result = cursor.fetchone()
            
            
            
            print('名前：', name)
            if result:
                # 名前が一致する行を探し、出席を記録
                for row in range(1, temp_sheet.max_row + 1):
                    if temp_sheet.cell(row=row, column=1).value == name:
                        temp_sheet.cell(row=row, column=3, value='出席')
                        workbook.save(ar_filename)
                        
                        information = f'{name}さんの出席処理は完了しました。'
                        window["-NAME-"].update("")  # 入力フィールドをクリア
                        conn.commit()  # 変更を確定
                        
                        # ウィンドウを閉じてから新しいウィンドウを作成
                        window.close()
                        window = mainwindowshow()
                        break
            else:
                # 失敗処理
                print(f"{name} はデータベースに存在しません。")
                messagebox.showinfo('失敗', f'{name} はデータベースに存在しません。')
                window["-NAME-"].update("")  # 入力フィールドをクリア
        
        #? 閉じられるときの処理
        if event == '終了':
            
            window.close()
            
            # 警告メッセージ表示
            endresult = messagebox.askquestion('警告', '本当に閉じますか？', icon='warning')
            if endresult == 'yes': # yes
                
                start_row = 2
                end_row = len(all_data) + 1
                start_column = 3
                end_column = 3
                
                # コピー先の開始セルの指定
                dest_start_row = 2
                dest_start_column = int(current_date_d) + 2 # 文字列を整数値に変換
                
                # 範囲をコピーしてコピー先のセルに貼り付ける
                for row in range(start_row, end_row + 1):
                    for col in range(start_column, end_column + 1):
                        cell_value = temp_sheet.cell(row=row, column=col).value
                        dest_row = row - start_row + dest_start_row
                        dest_col = col - start_column + dest_start_column
                        dest_cell = sheet.cell(row=dest_row, column=dest_col)
                        # セルの値をコピー
                        dest_cell.value = cell_value
                
                
                workbook.save(ar_filename)
                
                #? 一時シート削除
                
                # 削除したいシート名を指定
                sheet_name_to_delete = "temp"
                
                # シートを削除
                if sheet_name_to_delete in workbook.sheetnames:
                    sheet_to_delete = workbook[sheet_name_to_delete]
                    workbook.remove(sheet_to_delete)
                    print(f"{sheet_name_to_delete} シートを削除しました。")
                else:
                    print(f"{sheet_name_to_delete} シートは存在しません。")
                
                workbook.save(ar_filename)
                
                
                #? excelの色付け
                print("Color Change >>> ", end="")
                
                # 時間変数の設定
                current_date = datetime.now()
                current_date_y = current_date.strftime("%Y")
                
                # 一時ファイル名
                ar_filename = f"{current_date_y}Attendance records.xlsx"
                
                try:
                    workbook = load_workbook(ar_filename)
                except FileNotFoundError:
                    exit_with_error("File not found")
                
                # 出席と欠席のセルの背景色を定義
                absent_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')  # 赤色
                present_fill = PatternFill(start_color='ADFF2F', end_color='ADFF2F', fill_type='solid')  # 緑色
                
                for sheet in workbook.sheetnames:
                    current_sheet = workbook[sheet]
                    
                    # すべてのセルを調べる
                    for row in current_sheet.iter_rows():
                        for cell in row:
                            # セルの値が欠席か出席かを確認し、背景色を変更する
                            if cell.value == '欠席':
                                cell.fill = absent_fill
                            elif cell.value == '出席':
                                cell.fill = present_fill
                                
                # 変更を保存する
                workbook.save(ar_filename)
                
                print("done")
                
                messagebox.showinfo('完了', '記録終了は正常に終了しました。')
                window.close()
                break
            else:
                # ウィンドウを閉じてから新しいウィンドウを作成
                window.close()
                window = mainwindowshow()
                continue


# GUI画面のレイアウト
layout = [
    [sg.Text("起動する機能を選んでください。", font=("Helvetica", 15), justification='center')],  # カンマを追加
    [sg.Button('通常出席', bind_return_key=True, font=("Helvetica", 15)),
        sg.Button('終了', bind_return_key=True, font=("Helvetica", 15))]
]

menu = sg.Window('MENU', layout, finalize=True, keep_on_top=True)

# 記録ファイル名
ar_filename = f"{current_date_y}Attendance records.xlsx"

# ファイルが存在するかチェック
if not os.path.exists(ar_filename):
    # ファイルが存在しない場合の処理
    SApy()

if not os.path.exists("Register.db"):
    exit_with_error("名簿データが存在しません。")


while True:  #? 無限ループ
    event, values = menu.read()
    
    if event == sg.WIN_CLOSED or event == '終了':  # Xボタンが押されたか、'終了'ボタンが押された場合
        menu.close()
        sys.exit(0)
    
    if event == '通常出席':
        menu.close()
        
        GApy()
        messagebox.showinfo('インフォメーション', 'メニューが再度起動します。\nお待ちください。')
        
        python = sys.executable
        os.execl(python, python, * sys.argv)