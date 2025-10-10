# 🚀 PubMed2EndNote

PubMedの論文情報をワンクリックでEndNote形式に変換するChrome拡張機能（パッケージ化されていない拡張機能です）。Endnote本体を介した取り込みの手順を使わずに、PubMedのページから直接Wordファイル内に貼り付けができます！（引用すべき論文の貼り付け作業のテンポが相当よくなります！）

## ⚡ インストール手順（それなりに手作業が必要）
Chrome の拡張機能画面での操作、 PowerShell での実行、事前のPubMedアカウント登録など、いくつか手作業が必要です。

### 必要なもの
- Windows PC
- Google Chrome
- インターネット接続
- PubMed アカウント（後続の設定でこのアドレスを使用します）

### 手順
1. **ダウンロードと展開**: [リリースページ](https://github.com/matsuikentaro1/pubmed2endnote/releases)または「Code → Download ZIP」からアーカイブを取得し、展開して `PubMed2EndNote_1.0.0` フォルダを用意します。
2. **Chrome に読み込む**: Chrome で `chrome://extensions/` を開き、「パッケージ化されていない拡張機能を読み込む」から `PubMed2EndNote_1.0.0` フォルダを選びます。読み込み後に表示される拡張機能 ID を必ずメモしてください（後でスクリプトに入力します）。
3. **install.ps1 を実行**: `install.ps1` はダブルクリックでは起動しません。ファイルを右クリックし、「PowerShell で実行」を選んでください。スクリプト内では、手順 2 で控えた拡張機能 ID を入力します。
4. **メールアドレスの登録**: 任意のPubMedの記事を開いてください。拡張機能の読み込みが済んでいれば、右上に丸いアイコンが表示されているはずです。初回のみメール入力画面が別ウインドウで開くので、PubMed アカウントのメールアドレスを入力してください。受理されたらページを閉じて構いません。
5. **クリップボードにコピー**: PubMedの記事に戻り、改めて右上のアイコンをクリックしてください。"Successfully copied to clipboard!"と表記されたら成功です
6. **Wordに貼り付け**: Ctrl+V でWordファイル内の該当箇所に貼り付けしてください。Endnoteが認識してくれたら成功です。

## 🆘 トラブルシューティング

### インストールエラー
- **PowerShellスクリプトが実行できない**: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
- **Python未インストール**: https://www.python.org/downloads/ （「Add Python to PATH」にチェック）
- **pywin32エラー**: コマンドプロンプトで `pip install pywin32`

### 使用時のエラー  
- **青いアイコンが表示されない**: ページ再読み込み、拡張機能の有効化確認
- **クリップボードエラー**: Chrome拡張機能でメールアドレス設定を確認
- **EndNoteが認識したかわからない**: Wordで「EndNote」タブ → 「Update Citations」実行

### 署名関連のエラー
- `install.ps1` を実行した際に「デジタル署名されていません。このスクリプトは現在のシステムでは実行できません。」と表示された場合は、`install.ps1` を右クリック →「プロパティ」→「全般」タブの「セキュリティ：このファイルは他のコンピューターから…」欄にある「許可する（解除）」にチェックを入れて「OK」を押してから再度実行してください。

### その他諸注意
- `PubMed2EndNote_1.0.0` フォルダを別の場所に移動すると動作しなくなります。Chrome に読み込んだフォルダを移動しないでください。
- Chrome で読み込んだ拡張機能 ID が変わった場合も同様に動作しなくなります。その場合は `uninstall.ps1` を右クリック →「PowerShell で実行」でアンインストールし、手順 1 からやり直してください。

### サポート
問題が解決しない場合は [Issues](https://github.com/matsuikentaro1/pubmed2endnote/issues) で報告してください。
