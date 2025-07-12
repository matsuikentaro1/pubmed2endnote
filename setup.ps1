# PubMed2EndNote 自動セットアップスクリプト
# PowerShell実行ポリシーの確認が必要な場合があります

param(
    [switch]$Force
)

# 管理者権限チェック
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# カラー出力関数
function Write-ColorOutput($Message, $Color = "White") {
    Write-Host $Message -ForegroundColor $Color
}

# エラー処理関数
function Write-Error($Message) {
    Write-ColorOutput "❌ エラー: $Message" "Red"
}

function Write-Success($Message) {
    Write-ColorOutput "✅ $Message" "Green"
}

function Write-Info($Message) {
    Write-ColorOutput "ℹ️  $Message" "Cyan"
}

function Write-Warning($Message) {
    Write-ColorOutput "⚠️  $Message" "Yellow"
}

# メイン処理開始
Clear-Host
Write-ColorOutput "🚀 PubMed2EndNote セットアップウィザード" "Magenta"
Write-ColorOutput "=" * 50 "Magenta"
Write-Host ""

# 1. Python環境チェック
Write-Info "Step 1: Python環境をチェック中..."
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Python が見つかりました: $pythonVersion"
    } else {
        throw "Python が見つかりません"
    }
} catch {
    Write-Error "Python がインストールされていません"
    Write-Info "https://www.python.org/downloads/ からPythonをインストールしてください"
    Write-Info "インストール時は「Add Python to PATH」にチェックを入れてください"
    Read-Host "Python インストール後、Enterキーを押して続行..."
    return
}

# 2. pywin32チェック・インストール
Write-Info "Step 2: 必要なライブラリをチェック中..."
try {
    python -c "import win32clipboard" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Success "pywin32 が既にインストールされています"
    } else {
        throw "pywin32 が必要です"
    }
} catch {
    Write-Warning "pywin32 をインストールしています..."
    pip install pywin32
    if ($LASTEXITCODE -eq 0) {
        Write-Success "pywin32 のインストールが完了しました"
    } else {
        Write-Error "pywin32 のインストールに失敗しました"
        return
    }
}

# 3. 現在のディレクトリ確認
$currentPath = Get-Location
Write-Info "Step 3: セットアップディレクトリ: $currentPath"

# 必要ファイルの存在確認
$requiredFiles = @(
    "native_host.py",
    "native_host.bat", 
    "PubMed2EndNote_1.0.0\manifest.json"
)

foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        Write-Error "必要なファイルが見つかりません: $file"
        Write-Info "正しいフォルダでスクリプトを実行してください"
        return
    }
}
Write-Success "必要なファイルが全て確認できました"

# 4. Chrome拡張機能の情報取得
Write-Host ""
Write-ColorOutput "📋 Chrome拡張機能のセットアップ" "Yellow"
Write-Info "以下の手順で拡張機能IDを取得してください:"
Write-Host "1. Google Chromeを開く"
Write-Host "2. chrome://extensions/ にアクセス"
Write-Host "3. 右上の「デベロッパーモード」をON"
Write-Host "4. 「パッケージ化されていない拡張機能を読み込む」をクリック"
Write-Host "5. このフォルダ内の「PubMed2EndNote_1.0.0」フォルダを選択"
Write-Host "6. インストールされた拡張機能の「ID」をコピー"
Write-Host ""

do {
    $extensionId = Read-Host "拡張機能ID を入力してください（例: abcdefghijklmnopqrstuvwxyz123456）"
    
    # IDの形式チェック（32文字の英数字）
    if ($extensionId -match "^[a-z]{32}$") {
        Write-Success "有効な拡張機能IDです: $extensionId"
        break
    } else {
        Write-Error "無効な拡張機能IDです。32文字の小文字英数字である必要があります"
        $retry = Read-Host "再入力しますか？ (y/n)"
        if ($retry -ne "y" -and $retry -ne "Y") {
            return
        }
    }
} while ($true)

# 5. 設定ファイル生成
Write-Info "Step 4: 設定ファイルを生成中..."

# native_host.batのパス（絶対パス）
$nativeHostBatPath = Join-Path $currentPath "native_host.bat"
$nativeHostBatPath = $nativeHostBatPath -replace "\\", "\\"

# com.pubmed.endnote.json生成
$configJson = @{
    name = "com.pubmed.endnote"
    description = "Native messaging host for PubMed2EndNote Chrome extension"
    path = $nativeHostBatPath
    type = "stdio"
    allowed_origins = @("chrome-extension://$extensionId/")
} | ConvertTo-Json -Depth 3

$configJson | Out-File -FilePath "com.pubmed.endnote.json" -Encoding UTF8
Write-Success "設定ファイル com.pubmed.endnote.json を生成しました"

# 6. レジストリファイル生成
Write-Info "Step 5: レジストリファイルを生成中..."

$registryPath = Join-Path $currentPath "com.pubmed.endnote.json"
$registryPath = $registryPath -replace "\\", "\\"

$regContent = @"
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.pubmed.endnote]
@="$registryPath"
"@

$regContent | Out-File -FilePath "register_native_host.reg" -Encoding UTF8
Write-Success "レジストリファイル register_native_host.reg を生成しました"

# 7. レジストリ登録
Write-Info "Step 6: Windowsレジストリに登録中..."

if (-not (Test-Administrator)) {
    Write-Warning "管理者権限が必要です。レジストリファイルを手動で実行してください"
    Write-Info "「register_native_host.reg」をダブルクリックして「はい」を選択してください"
    $manual = Read-Host "手動でレジストリ登録を完了しましたか？ (y/n)"
    if ($manual -ne "y" -and $manual -ne "Y") {
        Write-Warning "セットアップを中断しました"
        return
    }
} else {
    try {
        Start-Process "reg" -ArgumentList "import", "register_native_host.reg" -Wait -NoNewWindow
        Write-Success "レジストリへの登録が完了しました"
    } catch {
        Write-Error "レジストリ登録に失敗しました: $_"
        Write-Info "「register_native_host.reg」を手動で実行してください"
        return
    }
}

# 8. 動作テスト
Write-Info "Step 7: 基本動作テストを実行中..."

try {
    $testResult = python native_host.py --help 2>$null
    Write-Success "Python スクリプトが正常に実行できます"
} catch {
    Write-Warning "Python スクリプトのテストに失敗しました"
}

# 9. セットアップ完了
Write-Host ""
Write-ColorOutput "🎉 セットアップが完了しました！" "Green"
Write-ColorOutput "=" * 50 "Green"
Write-Host ""

Write-Info "次の手順で使用開始してください:"
Write-Host "1. Chrome拡張機能の設定でメールアドレスを入力"
Write-Host "2. PubMedの論文ページで青いアイコンをクリック"
Write-Host "3. Microsoft Wordで Ctrl+V で貼り付け"
Write-Host ""

Write-Info "設定ファイルの場所:"
Write-Host "- com.pubmed.endnote.json: $currentPath"
Write-Host "- レジストリキー: HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.pubmed.endnote"
Write-Host ""

Write-Warning "重要: このフォルダ ($currentPath) を移動・削除しないでください"
Write-Host ""

$openChrome = Read-Host "Chrome拡張機能の設定ページを開きますか？ (y/n)"
if ($openChrome -eq "y" -or $openChrome -eq "Y") {
    Start-Process "chrome://extensions/"
}

Write-ColorOutput "セットアップウィザードを終了します。" "Magenta"