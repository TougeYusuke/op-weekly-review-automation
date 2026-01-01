@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo テスト環境をセットアップしています...
echo.

echo 依存関係をインストールしています...
call npm install

if %ERRORLEVEL% NEQ 0 (
    echo エラー: npm install が失敗しました
    pause
    exit /b 1
)

echo.
echo テストを実行しています...
call npm test

if %ERRORLEVEL% NEQ 0 (
    echo エラー: テストが失敗しました
    pause
    exit /b 1
)

echo.
echo テスト環境のセットアップが完了しました！
echo.
pause

