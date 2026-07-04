# 🚀 PubMed2EndNote

PubMedの論文情報をワンクリックでEndNote引用としてWordに貼り付けられるChrome拡張機能

## ✨ 特徴

- PubMedの論文ページにコピー用ボタンを追加
- ワンクリックでEndNote対応の引用がクリップボードにコピーされる
- Wordに貼り付けるだけで、EndNoteが認識できる引用フィールドになる（黄色ハイライト付き）
- 著者名・雑誌名・巻号ページ・DOI・Abstract・MeSHタームなどの書誌情報を自動取得
- ウムラウト等の特殊文字にも対応

## ⚡ インストール

### 必要なもの
- Google Chrome
- インターネット接続
- Microsoft Word + EndNote（Cite While You Write）

### 手順
1. **ダウンロード**: 「Code → Download ZIP」で解凍
2. **Chrome拡張機能を読み込む**:
   - `chrome://extensions/` を開く
   - 「デベロッパーモード」をON
   - 「パッケージ化されていない拡張機能を読み込む」で `PubMed2EndNote_3.0.0` フォルダを選択

**完了！**

## 📱 使い方

1. **PubMedで論文検索**: https://pubmed.ncbi.nlm.nih.gov/
2. **青いアイコンをクリック**: 論文ページ右上に表示（初回はメールアドレス設定画面が開きます）
3. **Wordに貼り付け**: Ctrl+V でEndNote引用として貼り付け（黄色ハイライト付き）
4. EndNoteタブの「Update Citations and Bibliography」で引用と文献リストが整形されます

### ⚙️ 設定画面

拡張機能の設定画面（`chrome://extensions/` → PubMed2EndNote → 「拡張機能のオプション」）で以下を変更できます：

- **フォント / フォントサイズ**: 貼り付ける引用のフォントを執筆中の論文に合わせられます（Century, Times New Roman, Calibri, Arial など）。「Default」のままならWord側の書式になります
- **メールアドレス**: PubMed API（NCBI E-utilities）の利用マナーとしてNCBIにのみ送信されるもので、開発者が収集することはありません

※ メールアドレスはPubMed API（NCBI E-utilities）の利用マナーとして送信されるもので、それ以外には使用されません。

## 🛠️ 仕組み

1. PubMed E-utilities APIから論文の書誌情報（XML）を取得
2. EndNoteのTraveling Library形式（`ADDIN EN.CITE` フィールドコード）を生成
3. Word互換のHTML形式でクリップボードに書き込み
4. Wordが貼り付け時にフィールドを再構築し、EndNoteが引用として認識

すべてブラウザ内で完結し、外部ソフトのインストールは不要です。

## 🔄 アンインストール

`chrome://extensions/` で拡張機能を削除するだけです。

## 🆘 トラブルシューティング

- **青いアイコンが表示されない**: ページを再読み込みし、拡張機能が有効になっているか確認
- **クリックすると設定画面が開く**: メールアドレスを設定して保存してから再度クリック
- **EndNoteが認識しない**: Wordで「EndNote」タブ → 「Update Citations and Bibliography」を実行
- **貼り付けてもただの文字になる**: Wordの貼り付けオプションで「元の書式を保持」を選択
- **黄色ハイライトが付かない**: Wordの「ファイル → オプション → 詳細設定 → 切り取り、コピー、貼り付け」で「**他のプログラムからの貼り付け**」を「**元の書式を保持（既定）**」にしてください。「書式を結合」になっているとハイライトだけが除去されます（引用機能自体は動作します）

### サポート
問題が解決しない場合は [Issues](https://github.com/matsuikentaro1/pubmed2endnote/issues) で報告してください

---
🤖 このプロジェクトは[Claude Code](https://claude.ai/code)で開発されました
