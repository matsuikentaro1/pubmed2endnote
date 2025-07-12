# PubMed2EndNote拡張機能 セットアップガイド

このガイドでは、Native Messagingを使用したPubMed2EndNote Chrome拡張機能のセットアップ手順を説明します。

## 前提条件

- Windows 10/11
- Python 3.6以上がインストール済み
- Google Chrome

## セットアップ手順

### 1. Python依存関係のインストール

コマンドプロンプトまたはPowerShellで以下を実行：

```bash
pip install pywin32
```

### 2. 拡張機能IDの取得と設定

1. Chromeで `chrome://extensions/` にアクセス
2. 開発者モードを有効にする
3. 「パッケージ化されていない拡張機能を読み込む」で `PubMed2EndNote_1.0.0` フォルダを選択
4. 表示された拡張機能IDをコピー
5. `com.pubmed.endnote.json` ファイルを開き、`YOUR_EXTENSION_ID_HERE` を実際のIDに置き換える

例：
```json
{
    "name": "com.pubmed.endnote",
    "description": "Native messaging host for PubMed2EndNote Chrome extension",
    "path": "C:\\Users\\Kentaro Matsui\\Dropbox\\__Coding\\PubMed-EndNote-Word Quicklinker\\native_host.py",
    "type": "stdio",
    "allowed_origins": [
        "chrome-extension://abcdefghijklmnopqrstuvwxyz012345/"
    ]
}
```

### 3. レジストリへの登録

1. `register_native_host.reg` ファイルをダブルクリック
2. 「はい」をクリックしてレジストリに追加

### 4. 動作確認

1. PubMedの論文ページ（例：https://pubmed.ncbi.nlm.nih.gov/34886562/）にアクセス
2. 拡張機能アイコンをクリック
3. Microsoft Wordを開き、Ctrl+Vでペースト
4. EndNoteの引用フィールドとして認識されることを確認

## トラブルシューティング

### 拡張機能が動作しない場合

1. Chrome DevToolsで console エラーを確認
2. `native_host.py` へのパスが正しいか確認
3. Python環境で `import win32clipboard` が動作するか確認

### クリップボードにデータが書き込まれない場合

1. `pywin32` が正しくインストールされているか確認
2. ファイルパスにスペースが含まれる場合、引用符で囲む

### レジストリ登録に失敗する場合

管理者権限でコマンドプロンプトを開き、以下を実行：

```cmd
reg add "HKEY_CURRENT_USER\\SOFTWARE\\Google\\Chrome\\NativeMessagingHosts\\com.pubmed.endnote" /ve /d "C:\\Users\\Kentaro Matsui\\Dropbox\\__Coding\\PubMed-EndNote-Word Quicklinker\\com.pubmed.endnote.json"
```

## ファイル構成

```
PubMed-EndNote-Word Quicklinker/
├── PubMed2EndNote_1.0.0/          # Chrome拡張機能
│   ├── manifest.json               # 拡張機能設定（nativeMessaging権限追加済み）
│   ├── background.js               # Native Messaging対応済み
│   ├── content.js                  # 変更なし
│   └── icons/                      # アイコンファイル
├── native_host.py                  # Pythonネイティブホスト
├── com.pubmed.endnote.json         # ホストマニフェスト
├── register_native_host.reg        # レジストリ登録ファイル
└── SETUP.md                        # このファイル
```

## 注意事項

- このセットアップはWindows専用です
- ファイルパスを変更した場合は、`com.pubmed.endnote.json` と `register_native_host.reg` の両方を更新してください
- セキュリティ上、信頼できるディレクトリにファイルを配置することを推奨します