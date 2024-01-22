import os
from datetime import datetime
from tkinter import messagebox

# 現在の日付を取得
current_date = datetime.now().strftime("%Y-%m-%d")

# 変更前のファイル名
old_filename = "temp.xlsx"

# 新しいファイル名を作成
new_filename = f"{current_date}.xlsx"

# ファイル名の変更
os.rename(old_filename, new_filename)

print(f"ファイル名を変更しました: {old_filename} → {new_filename}")
messagebox.showinfo('完了', '記録終了は正常に終了しました。')