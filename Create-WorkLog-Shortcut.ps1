# Create-WorkLog-Shortcut.ps1
# This script creates a Desktop shortcut to run your Work Log mail sender

# --- Configuration ---
$scriptPath   = "E:\SAPTARSHI BANERJEE\PERSONAL\PROJECT\OTHERS\WORKLOGS\email-todays-worklog.ps1"
$recipient    = "siddhesh@ibps.in;bhoopendra.singh@ibps.in"   # Default recipient email
$shortcutName = "Send Work Log"
# ----------------------

# Get Desktop path
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "$shortcutName.lnk"

# PowerShell executable path
$powershellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

# Arguments to run the script silently (no black window pop-up)
$arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -To `"$recipient`""

# Create shortcut
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powershellExe
$shortcut.Arguments = $arguments
$shortcut.WorkingDirectory = (Split-Path $scriptPath)
$shortcut.IconLocation = "powershell.exe,0"
$shortcut.Description = "Click to send today's Work Log email"
$shortcut.Save()

Write-Host "✅ Shortcut created successfully:"
Write-Host "   $shortcutPath"
