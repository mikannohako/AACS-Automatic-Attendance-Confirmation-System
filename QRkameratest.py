import cv2

# カメラを起動
cap = cv2.VideoCapture(1)

# QRコードを検出するためのdetectorを作成
detector = cv2.QRCodeDetector()

while True:
    # カメラからフレームを取得
    ret, frame = cap.read()

    # QRコードを検出
    qr_data, _, _ = detector.detectAndDecode(frame)

    # 読み取ったQRコードがあれば
    if qr_data:
        print("QRコードの中の数値:", qr_data)

    # フレームにQRコードの内容を表示
    cv2.putText(frame, qr_data, (20, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

    # フレームを表示
    cv2.imshow('Camera', frame)

    # 'q'キーでループを抜ける
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# カメラを解放
cap.release()
cv2.destroyAllWindows()