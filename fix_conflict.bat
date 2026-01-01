@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Git競合を解決しています...
echo.

echo リモートの変更を採用します...
git checkout --theirs .github/workflows/test.yml

echo.
echo 競合解決をマークします...
git add .github/workflows/test.yml

echo.
echo コミットします...
git commit -m "Resolve merge conflict: use remote version of test.yml"

echo.
echo プッシュします...
git push origin main

echo.
echo 完了しました！
echo.
pause

