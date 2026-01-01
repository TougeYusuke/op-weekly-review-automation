@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Git環境をセットアップしています...
echo.

if not exist .git (
    echo Gitリポジトリを初期化しています...
    git init
)

echo ファイルをステージングしています...
git add .

echo 初期コミットを作成しています...
git commit -m "Initial commit: 週次レビュー自動化ツールの基本実装"

echo.
echo Git環境のセットアップが完了しました！
echo.
git log --oneline
pause

