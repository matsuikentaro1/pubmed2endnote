# 🚀 PubMed2EndNote

PubMedの論文情報をワンクリックでEndNote形式に変換するChrome拡張機能

## ⚡ 簡単インストール

### 必要なもの
- Windows PC
- Google Chrome
- インターネット接続

### 手順
1. **ダウンロード**:「Code → Download ZIP」
2. **install.ps1を実行**: 右クリック → 「PowerShellで実行」
3. **指示に従う**: 拡張機能IDの入力のみ

**完了！** ファイルはユーザーフォルダに安全にインストールされ、ダウンロードフォルダは削除できます。

## 📱 使い方

1. **PubMedで論文検索**: https://pubmed.ncbi.nlm.nih.gov/
2. **青いアイコンをクリック**: 論文ページ右上に表示
3. **Wordに貼り付け**: Ctrl+V でEndNote引用として貼り付け

## 🆘 トラブルシューティング

### インストールエラー
- **PowerShellスクリプトが実行できない**: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
- **Python未インストール**: https://www.python.org/downloads/ （「Add Python to PATH」にチェック）
- **pywin32エラー**: コマンドプロンプトで `pip install pywin32`

### 使用時のエラー  
- **青いアイコンが表示されない**: ページ再読み込み、拡張機能の有効化確認
- **クリップボードエラー**: Chrome拡張機能でメールアドレス設定を確認
- **EndNote認識しない**: Wordで「EndNote」タブ → 「Update Citations」実行

### サポート
問題が解決しない場合は [Issues](https://github.com/matsuikentaro1/pubmed2endnote/issues) で報告してください

---
🤖 このプロジェクトは[Claude Code](https://claude.ai/code)で開発されました