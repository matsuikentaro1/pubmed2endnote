# PubMed2EndNote アンインストーラー

function Write-ColorOutput($Message, $Color = "White") {
    Write-Host $Message -ForegroundColor $Color
}

Clear-Host
Write-ColorOutput "🗑️ PubMed2EndNote アンインストーラー" "Red"
Write-ColorOutput "=" * 50 "Red"

$confirm = Read-Host "PubMed2EndNoteを完全にアンインストールしますか？ (y/n)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    Write-ColorOutput "アンインストールをキャンセルしました" "Yellow"
    return
}

# レジストリ削除
Write-ColorOutput "レジストリエントリを削除中..." "Yellow"
try {
    reg delete "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.pubmed.endnote" /f 2>$null
    Write-ColorOutput "✅ レジストリ削除完了" "Green"
} catch {
    Write-ColorOutput "⚠️ レジストリエントリが見つかりませんでした" "Yellow"
}

# インストールフォルダ削除
$installPath = "$env:LOCALAPPDATA\PubMed2EndNote"
if (Test-Path $installPath) {
    Write-ColorOutput "インストールフォルダを削除中..." "Yellow"
    try {
        Remove-Item -Path $installPath -Recurse -Force
        Write-ColorOutput "✅ フォルダ削除完了: $installPath" "Green"
    } catch {
        Write-ColorOutput "❌ フォルダ削除に失敗しました: $_" "Red"
    }
} else {
    Write-ColorOutput "⚠️ インストールフォルダが見つかりませんでした" "Yellow"
}

Write-Host ""
Write-ColorOutput "🎉 アンインストール完了" "Green"
Write-ColorOutput "Chrome拡張機能は手動で削除してください" "Cyan"