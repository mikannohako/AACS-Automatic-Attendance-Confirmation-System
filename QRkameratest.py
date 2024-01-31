import cv2
import PySimpleGUI as sg
import time

# カメラを起動
cap = cv2.VideoCapture(1)

# QRコードを検出するためのdetectorを作成
detector = cv2.QRCodeDetector()

# ウィンドウに配置するコントロール
layout = [
    [sg.Image(filename='', key='-IMAGE-')],
    [sg.Text('', key='-QR_DATA-')],
    [sg.Button('Exit')]
]

# ウィンドウの生成
window = sg.Window('Camera Feed', layout, finalize=True)

# 前回のQRコードの内容を記録する変数と前回の読み取り時間を初期化
last_qr_data = None
last_qr_read_time = 0

# メインループ
while True:
    # イベントとデータの読み込み
    event, values = window.read(timeout=20)

    # 'Exit'ボタンが押されたらループを抜ける
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break

    # カメラからフレームを取得
    ret, frame = cap.read()

    # QRコードを検出
    qr_data, _, _ = detector.detectAndDecode(frame)

    # 読み取ったQRコードがあれば
    if qr_data and qr_data != last_qr_data:
        window['-QR_DATA-'].update(f'QRコードの中の数値: {qr_data}')
        last_qr_data = qr_data

    # OpenCVのBGR形式をPySimpleGUIの画像形式に変換してウィンドウに表示
    if ret:
        imgbytes = cv2.imencode('.png', frame)[1].tobytes()
        window['-IMAGE-'].update(data=imgbytes)

# ウィンドウを閉じる
window.close()

# カメラを解放
cap.release()