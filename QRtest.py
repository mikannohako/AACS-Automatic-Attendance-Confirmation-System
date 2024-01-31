import cv2

# 画像ファイルのパス
image_path = "testid1.png"

# 画像を読み込む
image = cv2.imread(image_path)

# QRコードを検出
detector = cv2.QRCodeDetector()
qr_data, _, _ = detector.detectAndDecode(image)

# 読み取ったQRコードがあれば
if qr_data:
    print("QRコードの中の数値:", qr_data)
else:
    print("QRコードが見つかりませんでした。")