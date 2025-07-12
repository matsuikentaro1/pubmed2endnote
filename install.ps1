# PubMed2EndNote インストーラー
# ファイルを固定場所にコピーしてポータブル化

param(
    [string]$InstallPath = "$env:LOCALAPPDATA\PubMed2EndNote"
)

function Write-ColorOutput($Message, $Color = "White") {
    Write-Host $Message -ForegroundColor $Color
}

Clear-Host
Write-ColorOutput "🚀 PubMed2EndNote インストーラー" "Magenta"
Write-ColorOutput "=" * 50 "Magenta"

# インストール先の確認
Write-ColorOutput "インストール先: $InstallPath" "Cyan"
$confirm = Read-Host "この場所にインストールしますか？ (y/n)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    $InstallPath = Read-Host "インストール先を入力してください"
}

# インストールディレクトリ作成
Write-ColorOutput "インストールディレクトリを作成中..." "Yellow"
if (-not (Test-Path $InstallPath)) {
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
}

# 必要ファイルをコピー
Write-ColorOutput "ファイルをコピー中..." "Yellow"
$filesToCopy = @(
    "native_host.py",
    "native_host.bat",
    "PubMed2EndNote_1.0.0"
)

foreach ($file in $filesToCopy) {
    Copy-Item -Path $file -Destination $InstallPath -Recurse -Force
    Write-ColorOutput "✅ コピー完了: $file" "Green"
}

# 設定ファイル生成
Write-ColorOutput "設定ファイルを生成中..." "Yellow"

# 拡張機能ID取得
Write-Host ""
Write-ColorOutput "Chrome拡張機能のIDを入力してください:" "Cyan"
Write-Host "1. chrome://extensions/ を開く"
Write-Host "2. デベロッパーモードをON"
Write-Host "3. PubMed2EndNote_1.0.0 フォルダを読み込む"
Write-Host "4. 表示されたIDをコピー"
Write-Host ""

do {
    $extensionId = Read-Host "拡張機能ID"
    if ($extensionId -match "^[a-z]{32}$") {
        break
    } else {
        Write-ColorOutput "❌ 無効なIDです。32文字の英数字を入力してください" "Red"
    }
} while ($true)

# 設定ファイル作成
$configPath = Join-Path $InstallPath "com.pubmed.endnote.json"
$batPath = Join-Path $InstallPath "native_host.bat"
$batPath = $batPath -replace "\\", "\\"

$configJson = @{
    name = "com.pubmed.endnote"
    description = "Native messaging host for PubMed2EndNote Chrome extension"
    path = $batPath
    type = "stdio"
    allowed_origins = @("chrome-extension://$extensionId/")
} | ConvertTo-Json -Depth 3

$configJson | Out-File -FilePath $configPath -Encoding UTF8

# レジストリ登録
Write-ColorOutput "Windowsレジストリに登録中..." "Yellow"
$configPath = $configPath -replace "\\", "\\"

try {
    $regKey = "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.pubmed.endnote"
    reg add $regKey /ve /d $configPath /f | Out-Null
    Write-ColorOutput "✅ レジストリ登録完了" "Green"
} catch {
    Write-ColorOutput "❌ レジストリ登録に失敗しました" "Red"
    Write-ColorOutput "手動で以下を実行してください:" "Yellow"
    Write-Host "reg add `"$regKey`" /ve /d `"$configPath`" /f"
    return
}

# 完了メッセージ
Write-Host ""
Write-ColorOutput "🎉 インストール完了！" "Green"
Write-ColorOutput "=" * 30 "Green"
Write-Host ""
Write-ColorOutput "インストール場所: $InstallPath" "Cyan"
Write-ColorOutput "今後このフォルダを移動・削除しないでください" "Yellow"
Write-Host ""
Write-ColorOutput "使用方法:" "Cyan"
Write-Host "1. Chrome拡張機能の設定でメールアドレスを入力"
Write-Host "2. PubMedで論文ページを開く"
Write-Host "3. 青いアイコンをクリック"
Write-Host "4. Wordで Ctrl+V で貼り付け"