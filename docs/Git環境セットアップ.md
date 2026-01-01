# Git環境セットアップ

**週次レビュー自動化ツールのGit環境セットアップ手順**

---

## 📋 前提条件

- Gitがインストールされていること
- コマンドプロンプトまたはPowerShellへのアクセス権限

---

## 🚀 セットアップ手順

### Step 1: プロジェクトディレクトリに移動

```bash
cd D:\my_mind_project\週次レビュー自動化ツール
```

---

### Step 2: Gitリポジトリの初期化（既に完了している場合はスキップ）

```bash
git init
```

**注意:** 既にGitリポジトリが初期化されている場合は、このステップをスキップしてください。

---

### Step 3: .gitignoreファイルの確認

`.gitignore`ファイルが存在することを確認してください。

**内容:**
- `node_modules/` - npm依存関係
- `.clasp.json` - clasp設定ファイル（機密情報を含む可能性）
- `appsscript.json` - GAS設定ファイル
- その他の一時ファイルやログファイル

---

### Step 4: ファイルをステージング

```bash
git add .
```

または、個別に追加する場合：

```bash
git add .gitignore
git add main.js
git add docs/
```

---

### Step 5: 初期コミット

```bash
git commit -m "Initial commit: 週次レビュー自動化ツールの基本実装"
```

---

### Step 6: リモートリポジトリの設定（オプション）

GitHubなどのリモートリポジトリを使用する場合：

```bash
# リモートリポジトリを追加
git remote add origin https://github.com/your-username/your-repo.git

# ブランチ名をmainに変更（オプション）
git branch -M main

# リモートにプッシュ
git push -u origin main
```

---

## 📝 今後の作業フロー

### 変更をコミットする場合

```bash
# 変更をステージング
git add .

# コミット
git commit -m "コミットメッセージ"

# リモートにプッシュ（リモートリポジトリを設定している場合）
git push
```

### コミットメッセージの例

- `feat: 新機能を追加`
- `fix: バグを修正`
- `docs: ドキュメントを更新`
- `refactor: コードをリファクタリング`

---

## 🔧 トラブルシューティング

### Q1: "fatal: pathspec '.gitignore' did not match any files" エラー

**A:** `.gitignore`ファイルが存在しない可能性があります。以下のコマンドで確認してください：

```bash
ls .gitignore
```

ファイルが存在しない場合は、`docs/`フォルダ内の`.gitignore`テンプレートをコピーしてください。

### Q2: 文字エンコーディングの問題

**A:** PowerShellで日本語のディレクトリ名が正しく処理されない場合、コマンドプロンプト（cmd.exe）を使用してください。

### Q3: Gitがインストールされていない

**A:** [Git公式サイト](https://git-scm.com/downloads)からGitをダウンロードしてインストールしてください。

---

## 📚 関連ドキュメント

- [[README|README.md]] - プロジェクト概要
- [[開発仕様書|開発仕様書.md]] - 技術仕様、アーキテクチャ

---

**Status:** ✅ 完成

