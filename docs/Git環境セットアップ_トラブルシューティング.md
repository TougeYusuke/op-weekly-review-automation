# Git環境セットアップ - トラブルシューティング

## 🔧 PowerShellでの文字エンコーディング問題

### 問題

PowerShellで日本語のディレクトリ名を含むパスを処理する際、文字エンコーディングの問題が発生することがあります。

**エラーメッセージ例:**
```
fatal: not a git repository (or any of the parent directories): .git
```

### 解決方法

#### 方法1: バッチファイルを使用（推奨）

プロジェクトルートにある `setup_git.bat` をダブルクリックして実行してください。

または、コマンドプロンプトで：

```cmd
cd /d "D:\my_mind_project\週次レビュー自動化ツール"
setup_git.bat
```

#### 方法2: コマンドプロンプト（cmd.exe）で直接実行

1. **コマンドプロンプトを開く**
   - Windowsキー + R
   - `cmd` と入力してEnter

2. **以下のコマンドを実行**

```cmd
cd /d "D:\my_mind_project\週次レビュー自動化ツール"
git init
git add .
git commit -m "Initial commit: 週次レビュー自動化ツールの基本実装"
```

#### 方法3: PowerShellでUTF-8エンコーディングを設定

PowerShellで実行する場合：

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:LANG = "ja_JP.UTF-8"
cd "D:\my_mind_project\週次レビュー自動化ツール"
git init
git add .
git commit -m "Initial commit: 週次レビュー自動化ツールの基本実装"
```

---

## 📝 手動セットアップ手順（確実な方法）

### Step 1: コマンドプロンプトを開く

1. Windowsキー + R
2. `cmd` と入力してEnter

### Step 2: プロジェクトディレクトリに移動

```cmd
cd /d "D:\my_mind_project\週次レビュー自動化ツール"
```

### Step 3: Gitリポジトリの初期化

```cmd
git init
```

**期待される出力:**
```
Initialized empty Git repository in D:/my_mind_project/週次レビュー自動化ツール/.git/
```

### Step 4: ファイルをステージング

```cmd
git add .
```

### Step 5: ステータス確認

```cmd
git status
```

**期待される出力:**
```
On branch master

No commits yet

Changes to be committed:
  (use "git rm --cached <file>..." to unstage)
        new file:   .gitignore
        new file:   README_GIT.md
        new file:   main.js
        new file:   docs/Git環境セットアップ.md
        ...
```

### Step 6: 初期コミット

```cmd
git commit -m "Initial commit: 週次レビュー自動化ツールの基本実装"
```

**期待される出力:**
```
[master (root-commit) xxxxxxx] Initial commit: 週次レビュー自動化ツールの基本実装
 X files changed, XXX insertions(+)
 create mode 100644 .gitignore
 create mode 100644 README_GIT.md
 ...
```

### Step 7: コミット確認

```cmd
git log --oneline
```

**期待される出力:**
```
xxxxxxx Initial commit: 週次レビュー自動化ツールの基本実装
```

---

## ❓ よくある質問

### Q1: "fatal: not a git repository" エラーが出る

**A:** Gitリポジトリが初期化されていません。`git init`を実行してください。

### Q2: "nothing to commit" と表示される

**A:** ファイルがステージングされていません。`git add .`を実行してください。

### Q3: PowerShellで文字化けする

**A:** コマンドプロンプト（cmd.exe）を使用するか、上記の方法3でUTF-8エンコーディングを設定してください。

### Q4: ファイルが追加されない

**A:** `.gitignore`で除外されている可能性があります。`git status`で確認してください。

---

## 🔗 関連ドキュメント

- [[Git環境セットアップ|Git環境セットアップ.md]] - 基本的なセットアップ手順
- [[README_GIT|../README_GIT.md]] - クイックスタートガイド

---

**Status:** ✅ 完成

