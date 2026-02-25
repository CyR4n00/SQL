@echo off
chcp 65001 > nul
echo.
echo ====================================
echo  DB マネージャー セットアップ & 起動
echo ====================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] Python が見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    echo インストール時に「Add Python to PATH」に必ずチェックを入れてください。
    pause
    exit /b 1
)

echo [OK] Python 確認済み
pip install -r requirements.txt -q
echo [OK] パッケージ確認済み
echo.
echo GUI アプリを起動します...
python app.py

pause
