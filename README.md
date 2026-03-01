# 🚀 PubMed2EndNote

PubMedの論文情報をワンクリックでEndNote形式に変換するChrome拡張機能

## ⚡ 簡単インストール

### 必要なもの
- Windows PC
- Google Chrome
- インターネット接続

### 手順
1. **ダウンロード**: 「Code → Download ZIP」で解凍
2. **Chrome拡張機能を読み込む**:
   - `chrome://extensions/` を開く
   - 「デベロッパーモード」をON
   - 「パッケージ化されていない拡張機能を読み込む」で `PubMed2EndNote_2.0.0` フォルダを選択
   - 表示された拡張機能ID（32文字）をコピーしておく
3. **`install.bat` をダブルクリック**: 指示に従って拡張機能IDを入力

**完了！** ファイルはユーザーフォルダに安全にインストールされ、ダウンロードフォルダは削除できます。

## 📱 使い方

1. **PubMedで論文検索**: https://pubmed.ncbi.nlm.nih.gov/
2. **青いアイコンをクリック**: 論文ページ右上に表示
3. **Wordに貼り付け**: Ctrl+V でEndNote引用として貼り付け（黄色ハイライト付き）

## 🔄 アンインストール

`uninstall.bat` をダブルクリックし、Chrome拡張機能を手動で削除

## 🆘 トラブルシューティング

### インストールエラー
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
