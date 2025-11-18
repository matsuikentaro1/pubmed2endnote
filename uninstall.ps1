# PubMed2EndNote Uninstaller

function Write-ColorOutput($Message, $Color = "White") {
    Write-Host $Message -ForegroundColor $Color
}

Clear-Host
Write-ColorOutput "PubMed2EndNote Uninstaller" "Red"
Write-ColorOutput ("=" * 50) "Red"

$confirm = Read-Host "Do you want to completely uninstall PubMed2EndNote? (y/n)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    Write-ColorOutput "Uninstallation cancelled." "Yellow"
    return
}

# Remove registry entry
Write-ColorOutput "Removing registry entry..." "Yellow"
try {
    reg delete "HKEY_CURRENT_USER\SOFTWARE\Google\Chrome\NativeMessagingHosts\com.pubmed.endnote" /f 2>$null
    Write-ColorOutput "Registry entry removed." "Green"
} catch {
    Write-ColorOutput "Registry entry not found." "Yellow"
}

# Remove installation folder
$installPath = "$env:LOCALAPPDATA\PubMed2EndNote"
if (Test-Path $installPath) {
    Write-ColorOutput "Removing installation folder..." "Yellow"
    try {
        Remove-Item -Path $installPath -Recurse -Force
        Write-ColorOutput "Folder removed: $installPath" "Green"
    } catch {
        Write-ColorOutput "Failed to remove folder: $_" "Red"
    }
} else {
    Write-ColorOutput "Installation folder not found." "Yellow"
}

Write-Host ""
Write-ColorOutput "Uninstallation completed." "Green"
Write-ColorOutput "Please remove the Chrome extension manually." "Cyan"
