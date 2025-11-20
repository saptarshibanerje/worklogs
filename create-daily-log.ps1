# create-daily-log-safe-fixed.ps1
# Works on Windows PowerShell 5.1 and PowerShell 7+
# Writes UTF-8 with BOM reliably by writing bytes
# Creates Desktop shortcut, logs errors, pauses at end

$ErrorActionPreference = "Stop"

try {
    # === CONFIG ===
    $basePath = "E:\SAPTARSHI BANERJEE\PERSONAL\PROJECT\OTHERS\WORKLOGS"    # <-- change this path if you want
    $createShortcutOnDesktop = $false
    $iconLocation = ""  # e.g. "C:\Users\You\Pictures\md-icon.ico"

    # === Paths & Names ===
    $today = Get-Date
    $year = $today.ToString("yyyy")
    $month = $today.ToString("MMMM")
    $fileName = $today.ToString("yyyy-MM-dd") + ".md"
    $folderPath = Join-Path $basePath "$year\$month"
    $filePath = Join-Path $folderPath $fileName

    if (-not (Test-Path $folderPath)) {
        New-Item -ItemType Directory -Path $folderPath -Force | Out-Null
    }

    # === Markdown content (with emojis) ===
    $content = @"
# 🗓️ Work Log – $($today.ToString('dd MMM yyyy'))

## ✅ Completed
-

## 🔃 In Progress
-

## 🧠 Learnings / Notes
-

## 🚧 Blockers
-
"@

    # === Write file with UTF-8 BOM in a cross-version way ===
    if (-not (Test-Path $filePath)) {
        # Get UTF8 BOM bytes and content bytes
        $bom = [System.Text.Encoding]::UTF8.GetPreamble()
        $contentBytes = [System.Text.Encoding]::UTF8.GetBytes($content)
        $all = New-Object byte[] ($bom.Length + $contentBytes.Length)
        [Array]::Copy($bom, 0, $all, 0, $bom.Length)
        [Array]::Copy($contentBytes, 0, $all, $bom.Length, $contentBytes.Length)

        # Ensure directory exists then write bytes
        [System.IO.File]::WriteAllBytes($filePath, $all)

        Write-Host "Created log file with UTF-8 BOM: $filePath"
    } else {
        Write-Host "File already exists: $filePath"
    }

    # === Shortcut creation on Desktop ===
    if ($createShortcutOnDesktop) {
        $wShell = New-Object -ComObject WScript.Shell
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutName = "WorkLog - " + $today.ToString("yyyy-MM-dd") + ".lnk"
        $shortcutPath = Join-Path $desktopPath $shortcutName

        $shortcut = $wShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $filePath
        $shortcut.WorkingDirectory = $folderPath
        $shortcut.WindowStyle = 1
        $shortcut.Description = "Daily work log - $($today.ToString('dd MMM yyyy'))"

        if (![string]::IsNullOrWhiteSpace($iconLocation) -and (Test-Path ($iconLocation.Split(",")[0]))) {
            $shortcut.IconLocation = $iconLocation
        } else {
            $notepadPath = Join-Path $Env:SystemRoot "System32\notepad.exe"
            if (Test-Path $notepadPath) {
                $shortcut.IconLocation = "$notepadPath,0"
            } else {
                $explorerPath = Join-Path $Env:SystemRoot "explorer.exe"
                $shortcut.IconLocation = "$explorerPath,0"
            }
        }

        $shortcut.Save()
        Write-Host "Created shortcut on Desktop: $shortcutPath"
    }

    # Start-Process -FilePath $filePath
    Write-Host "`nScript finished successfully."
}
catch {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    if (-not $scriptDir) { $scriptDir = Get-Location }
    $errorLogPath = Join-Path $scriptDir "error.log"

    $errTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorText = @()
    $errorText += "=== ERROR at $errTime ==="
    $errorText += "Message: $($_.Exception.Message)"
    $errorText += "Type: $($_.Exception.GetType().FullName)"
    $errorText += "StackTrace:"
    $errorText += $_.Exception.StackTrace
    $errorText += "`nFull ErrorRecord:"
    $errorText += $_ | Out-String
    $errorText += "`n"

    $errorText | Out-File -FilePath $errorLogPath -Encoding utf8

    Write-Host "An error occurred. Details written to: $errorLogPath" -ForegroundColor Red
    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host ""
    Write-Host "Press ENTER to close this window..."
    Read-Host | Out-Null
}
