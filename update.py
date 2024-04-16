import requests
import os
import zipfile

def download_and_extract_latest_release():
    # GitHubのリポジトリ情報
    owner = "mikannohako"
    repo = "AACS-Automatic-Attendance-Confirmation-System"
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"

    # 最新のリリース情報を取得
    response = requests.get(api_url)
    if response.status_code == 200:
        release_info = response.json()
        # バージョン取得
        tag_name = release_info["tag_name"]
        print("Release version:", tag_name)
        assets = release_info["assets"]
        # 最新のZIPファイルのダウンロードURLを取得
        download_url = None  # 初期値を設定
        for asset in assets:
            if asset["name"] == "AACS.zip":
                download_url = asset["browser_download_url"]
                break

        # ZIPファイルをダウンロード
        if download_url:
            response = requests.get(download_url)
            if response.status_code == 200:
                # ZIPファイルを保存
                with open("AACS.zip", "wb") as f:
                    f.write(response.content)
                print("Download completed successfully.")

                # AACSディレクトリを作成
                if not os.path.exists("AACS"):
                    os.makedirs("AACS")

                # ZIPファイルを解凍
                with zipfile.ZipFile("AACS.zip", "r") as zip_ref:
                    zip_ref.extractall("AACS")
                print("Extraction completed successfully.")

                # ZIPファイルを削除
                os.remove("AACS.zip")
                print("ZIP file deleted.")
            else:
                print("Failed to download ZIP file.")
        else:
            print("No ZIP file found in the release.")
    else:
        print("Failed to fetch release information.")

if __name__ == "__main__":
    download_and_extract_latest_release()
