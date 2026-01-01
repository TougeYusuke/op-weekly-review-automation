@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Git競合を解決しています...
echo.

echo 競合状態を確認...
git status

echo.
echo 競合を解決する方法を選択してください:
echo 1. リモートの変更を採用（推奨）
echo 2. ローカルの変更を採用
echo 3. 手動で解決
echo.
set /p choice="選択 (1/2/3): "

if "%choice%"=="1" (
    echo リモートの変更を採用しています...
    git checkout --theirs .github/workflows/test.yml
    git add .github/workflows/test.yml
    git commit -m "Resolve merge conflict: use remote version"
    echo.
    echo 競合が解決されました。プッシュします...
    git push origin main
) else if "%choice%"=="2" (
    echo ローカルの変更を採用しています...
    git checkout --ours .github/workflows/test.yml
    git add .github/workflows/test.yml
    git commit -m "Resolve merge conflict: use local version"
    echo.
    echo 競合が解決されました。プッシュします...
    git push origin main
) else (
    echo 手動で解決してください。
    echo 1. 競合ファイルを開く
    echo 2. 競合マーカーを削除
    echo 3. git add .github/workflows/test.yml
    echo 4. git commit -m "Resolve merge conflict"
    echo 5. git push origin main
)

echo.
pause

