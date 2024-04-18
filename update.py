#? インポート
import PySimpleGUI as sg
import sys
import os
from tkinter import messagebox
import json
import logging
import requests
import zipfile

#? ログの設定

# ログファイルの出力パス
filename = 'logfile.log'

# ログのメッセージフォーマットを指定
fmt = "%(asctime)s - %(levelname)s - %(message)s - %(module)s - %(funcName)s - %(lineno)d"

# ログの出力レベルを設定
logging.basicConfig(filename=filename, encoding='utf-8', level=logging.INFO, format=fmt)

# 設定ファイルのパス
config_file_path = 'config.json'

# 設定ファイルの読み込み
with open(config_file_path, 'r') as config_file:
    config_data = json.load(config_file)

def json_save(): #? JSONデータを保存
    #
    with open('config.json', 'w') as f:
        json.dump(config_data, f, indent=4)


def update(): #? アップデート
    
    def check_internet_connection():
        try:
            response = requests.get("http://www.google.com", timeout=5)
            response.raise_for_status()  # HTTPエラーコードが返ってきた場合に例外を発生させる
            
            return True
        except requests.RequestException as e:
            logging.warning("インターネット接続の確立に失敗しました。:", e)
            return False
    
    def update_check():
        api_url = f"https://api.github.com/repos/mikannohako/AACS-Automatic-Attendance-Confirmation-System/releases/latest"
        
        # 最新のリリース情報を取得
        response = requests.get(api_url)
        if response.status_code == 200:
            release_info = response.json()
            # バージョン取得
            tag_name = release_info["tag_name"]
            tag_name_int = int(tag_name.replace("v", "").replace(".", ""))
            
            if config_data["version"] < tag_name_int:
                messagebox.showinfo("info", "実行中のAACSがある場合は閉じてください。")
                # 最新のバージョンをダウンロードする
                file_update()
            else:
                messagebox.showinfo("更新", "最新バージョンです。")
    
    def file_update():
        
        BAR_MAX = 70
        
        layout = [
            [sg.Text('更新中')],
            [sg.ProgressBar(BAR_MAX, orientation='h', size=(20, 20), key='-PROG-')],
            ]
        
        window = sg.Window('更新中', layout, keep_on_top=True)
        
        # ここでウィンドウを初期化
        event, values = window.read(timeout=0)
        if event == sg.WINDOW_CLOSED:
            window.close()
            return
        
        window['-PROG-'].update(10)
        
        # GitHubのリポジトリ情報
        api_url = f"https://api.github.com/repos/mikannohako/AACS-Automatic-Attendance-Confirmation-System/releases/latest"
        
        # 最新のリリース情報を取得
        response = requests.get(api_url)
        if response.status_code == 200:
            release_info = response.json()
            # バージョン取得
            tag_name = release_info["tag_name"]
            
            window['-PROG-'].update(20)
            
            assets = release_info["assets"]
            # 最新のZIPファイルのダウンロードURLを取得
            download_url = None  # 初期値を設定
            for asset in assets:
                if asset["name"] == "update.zip":
                    download_url = asset["browser_download_url"]
                    break
                
            window['-PROG-'].update(30)
            
            # ZIPファイルをダウンロード
            if download_url:
                response = requests.get(download_url)
                if response.status_code == 200:
                    
                    window['-PROG-'].update(40)
                    
                    # ZIPファイルを保存
                    with open("update.zip", "wb") as f:
                        f.write(response.content)
                    
                    window['-PROG-'].update(50)
                    
                    try:
                        # ZIPファイルを解凍
                        with zipfile.ZipFile("update.zip", "r") as zip_ref:
                            zip_ref.extractall(".")
                    except Exception:
                        messagebox.showerror("Error", "エラーが発生しました。")
                        logging.warn("ZIPの解凍ができない。")
                        return
                    
                    window['-PROG-'].update(60)
                    
                    # ZIPファイルを削除
                    os.remove("update.zip")
                    
                    
                    window['-PROG-'].update(70)
                    
                    # configをバージョンアップ
                    # バージョン取得
                    tag_name = release_info["tag_name"]
                    tag_name_int = int(tag_name.replace("v", "").replace(".", ""))
                    config_data["version"] = tag_name_int
                    json_save()
                    
                    logging.info(f"{tag_name}に更新しました。")
                    
                    window.close()
                    
                    messagebox.showinfo("完了", "正常に更新されました。")
                    
                    sys.exit(0)  # アプリケーションを正常終了コード 0 で終了します。
                else:
                    logging.warning("ZIPファイルのダウンロードに失敗しました。")
                    
                    messagebox.showwarning("失敗", "更新が失敗しました。\nネットワークの問題の可能性があります。")
                    return
            else:
                logging.warning("リリースにZIPファイルが見つかりません。")
        else:
            logging.warning("リリース情報の取得に失敗しました。")
        
        messagebox.showwarning("失敗", "更新が失敗しました。")
    
    if check_internet_connection():
        update_check()

if __name__ == "__main__":
    update()