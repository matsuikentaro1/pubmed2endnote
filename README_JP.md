# PubMed2EndNote（Chrome拡張機能）

PubMedの論文情報を瞬時にEndNote形式でMicrosoft WordにコピーできるChrome拡張機能です。

## 📋 概要

この拡張機能を使用すると、PubMedのページで論文を閲覧中に、ワンクリックでEndNote互換のRTF形式の引用をクリップボードにコピーできます。コピーした引用は、Microsoft WordでEndNoteフィールドとして直接貼り付けることができます。

## ✨ 機能

- **ワンクリック引用コピー**: PubMedページに表示されるアイコンをクリックするだけ
- **EndNote完全互換**: Word上でEndNoteフィールドとして認識される形式
- **Rich Text Format (RTF)**: WindowsのネイティブクリップボードでCF_RTF形式を使用
- **完全な論文情報**: 著者、タイトル、雑誌名、年、DOI、抄録、キーワードなどすべて含む
- **美しいUI**: モダンなデザインの設定画面とステータス表示

## 🚀 インストール方法

### 1. ファイルのダウンロード
```bash
git clone https://github.com/matsuikentaro1/pubmed2endnote.git
cd pubmed2endnote
```

### 2. Chrome拡張機能のインストール
1. Chromeで `chrome://extensions/` を開く
2. 右上の「デベロッパーモード」をオンにする
3. 「パッケージ化されていない拡張機能を読み込む」をクリック
4. `PubMed2EndNote_1.0.0` フォルダを選択

### 3. 拡張機能IDの設定
1. Chrome拡張機能ページで表示されるIDをコピー
2. `com.pubmed.endnote.json` の `allowed_origins` を実際のIDに更新:
   ```json
   "allowed_origins": [
       "chrome-extension://[あなたの拡張機能ID]/"
   ]
   ```

### 4. Native Hostのセットアップ
1. Python環境の確認（Python 3.6以上）
2. 必要なライブラリのインストール:
   ```bash
   pip install pywin32
   ```
3. レジストリへの登録:
   ```bash
   # 管理者権限でコマンドプロンプトを開き、プロジェクトフォルダで実行
   reg import register_native_host.reg
   ```

## 📖 使い方

### 初回設定
1. 拡張機能のアイコンをクリックして設定画面を開く
2. PubMed API用のメールアドレスを入力して保存

### 論文の引用コピー
1. PubMedで論文ページを開く（例: `https://pubmed.ncbi.nlm.nih.gov/12345678/`）
2. 右上に表示される青いアイコンをクリック
3. 「Successfully copied to clipboard!」の表示を確認
4. Microsoft Wordで `Ctrl+V` で貼り付け
5. EndNoteで「Update Citations and Bibliography」を実行

## 🔧 技術仕様

### アーキテクチャ
- **Chrome拡張機能**: フロントエンド（JavaScript）
- **Native Host**: バックエンド（Python）でクリップボード操作
- **Native Messaging**: Chrome拡張機能とPythonスクリプト間の通信

### なぜPythonが必要？
ブラウザのWebクリップボードAPIでは `text/rtf` MIME タイプしか設定できませんが、WordとEndNoteは Windows ネイティブクリップボードの `CF_RTF` 形式を必要とします。PythonのWin32 APIを使用することで、正しい形式でクリップボードに書き込むことができます。

### ファイル構成
```
PubMed2EndNote/
├── PubMed2EndNote_1.0.0/        # Chrome拡張機能
│   ├── manifest.json
│   ├── background.js
│   ├── content.js
│   ├── options.html
│   ├── options.js
│   └── icons/
├── native_host.py               # Pythonバックエンド
├── native_host.bat              # 実行用バッチファイル
├── com.pubmed.endnote.json      # Native host設定
└── register_native_host.reg     # レジストリ登録用
```

## 🐛 トラブルシューティング

### 「Native host communication error」が表示される
1. Python環境とpywin32の確認:
   ```bash
   python --version
   python -c "import win32clipboard"
   ```
2. レジストリ登録の確認:
   ```bash
   reg query "HKEY_CURRENT_USER\\SOFTWARE\\Google\\Chrome\\NativeMessagingHosts\\com.pubmed.endnote"
   ```
3. 拡張機能IDが `com.pubmed.endnote.json` と一致しているか確認

### EndNote/Wordで引用が認識されない
1. クリップボードにコピーされているか確認（Ctrl+Vでテスト）
2. native_host.logでエラーを確認
3. 拡張機能の再読み込み

### アイコンが表示されない
1. PubMedの正しいページ（PMID付きURL）にいるか確認
2. ページの再読み込み
3. 拡張機能が有効になっているか確認

## 📝 更新履歴

### v1.0.0
- 初回リリース
- PubMedからEndNoteへの基本的な引用コピー機能
- Native Messaging対応
- 日本語ドキュメント

## 🤝 貢献

バグ報告や機能要望は、GitHubのIssuesページでお願いします。

## 📄 ライセンス

MIT License

## 📞 サポート

問題や質問がある場合は、以下を確認してください：
1. このREADMEのトラブルシューティング
2. `native_host.log` のエラーメッセージ
3. GitHubのIssuesページ

---

**注意**: この拡張機能はWindows環境でのみ動作します。macOSやLinuxでの使用には追加の開発が必要です。