@echo off
chcp 65001

REM .venv フォルダが存在するか確認
if not exist .\.venv (
    echo 仮想環境が見つかりません。setupを開始します...

    REM 仮想環境の作成
    echo Creating virtual environment...

    python -m venv .venv
    if %ERRORLEVEL% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo Done.
    echo.

    REM 仮想環境の有効化
    echo Activating virtual environment...

    call ".venv\Scripts\activate.bat"
    if %ERRORLEVEL% neq 0 (
        echo Failed to activate virtual environment.
        pause
        exit /b 1
    )
    echo Done.
    echo.
    
    REM 依存パッケージのインストール
    echo Installing dependencies...
    echo.

    pip install -r requirements.txt

    if %ERRORLEVEL% neq 0 (
        echo Failed to install packages.
        pause
        exit /b 1
    )

    echo.
    echo Done.

    echo.
    echo 処理は正常に完了しました。

) else (
    echo 仮想環境が見つかりました。

    REM 仮想環境を有効化
    call ".venv\Scripts\activate.bat"

    echo 仮想環境が有効化されました。
)

REM Pythonスクリプトを実行
echo Running Python script...
python main.py