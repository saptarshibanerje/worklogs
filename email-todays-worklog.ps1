<#
email-todays-worklog-no-bullets.ps1

Sends today's Work Log email (plain dash-separated lines)
via Outlook with Calibri 11pt styling and static red disclaimer.

Usage:
.\email-todays-worklog-no-bullets.ps1 -WorkLogsBase "E:\SAPTARSHI BANERJEE\WORKLOGS" -To "recipient@domain.com"
#>

param(
    [string]$WorkLogsBase = "E:\SAPTARSHI BANERJEE\WORKLOGS",
    [Parameter(Mandatory = $true)][string]$To,
    [string]$Cc = "",
    [string]$Bcc = "",
    [string]$SubjectPrefix = "Work Log -",
    [string]$LogFile = "$PSScriptRoot\emaillogs.log"
)

function Write-Log {
    param($m)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$ts`t$m" | Out-File -FilePath $LogFile -Append -Encoding utf8
    Write-Host $m
}

try {
    if (Test-Path $LogFile) { Remove-Item $LogFile -ErrorAction SilentlyContinue }
    Write-Log "=== START no-bullets run ==="

    # 1️⃣ Locate today's file
    $today = Get-Date
    $filePath = Join-Path (Join-Path $WorkLogsBase $today.ToString('yyyy')) (Join-Path $today.ToString('MMMM') ($today.ToString('yyyy-MM-dd') + ".md"))
    Write-Log "Looking for: $filePath"

    if (-not (Test-Path $filePath)) {
        throw "Worklog file not found: $filePath"
    }

    # 2️⃣ Read and extract lines starting with '-'
    $content = Get-Content -Raw -LiteralPath $filePath -Encoding UTF8
    $lineRegex = [regex]'^\-\s*(.+)$'
    $lines = @()
    foreach ($line in ($content -split "`r?`n")) {
        $m = $lineRegex.Match($line.Trim())
        if ($m.Success) { $lines += "- " + $m.Groups[1].Value.Trim() }
    }

    Write-Log ("Lines found: {0}" -f $lines.Count)
    if ($lines.Count -eq 0) {
        Write-Log "No dash-prefixed lines found. Nothing to send."
        exit 0
    }

    # Join with <br/> to preserve new lines visually in HTML
    $taskHtml = ($lines -join "<br/>`n")

    # 3️⃣ HTML template (no bullets, uses dash-prefixed lines)
    $finalHtml = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Work Log</title>
</head>
<body style="margin:0;padding:16px;background:#fff;font-family:Calibri,Arial,sans-serif;color:#111;font-size:11pt;line-height:1.45;">
  <div style="max-width:720px;margin:0 auto;">

    <!-- Greeting -->
    <div style="margin-bottom:12px;">
      <p style="margin:0 0 8px 0;">Dear Sir,</p>
      <p style="margin:0 0 8px 0;">Following activity has been done by me today:</p>
    </div>

    <!-- Work Log Lines -->
    <div style="margin-bottom:18px;">
      $taskHtml
    </div>

    <!-- Signature -->
    <hr style="border-top:1px solid #000;margin:16px 0;"></div>
    <div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">
        
       <p style="margin:0;">with regards</p> 

      <p style="margin:0;">Saptarshi Banerjee</p>
      <p style="margin:0;">Deputy Manager</p>
      <p style="margin:0;">Intercom No: 7183</p>        
      <p style="margin-bottom:10px;color:#B00000;font-weight:bold;">Disclaimer:</p>
      <p style="margin:0;color:#B00000;">
        <b>It is triggered from automated mail system created by the above user to automate his workflow.</b>
      </p>     
      <p style="font-size:9pt;color:#B00000;margin:0;line-height:1.4;">
        This e-mail and its attachments may contain IBPS official information. If you are not the intended recipient,
        please notify the sender immediately and delete this e-mail. Any dissemination or use of this information by a
        person other than the intended recipient is unauthorized. The responsibility lies with the recipient to check
        this email and any attachment for the presence of viruses.
      </p>
    </div>

  </div>
</body>
</html>
"@

    # 4️⃣ Save HTML for review
    $tempHtml = Join-Path $Env:TEMP "worklog-composed-no-bullets.html"
    $finalHtml | Out-File -FilePath $tempHtml -Encoding UTF8
    Write-Log ("Composed HTML saved to: {0}" -f $tempHtml)

    # 5️⃣ Send mail
    Write-Log "Creating Outlook COM object..."
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    $mail.To = $To
    if ($Cc) { $mail.CC = $Cc }
    if ($Bcc) { $mail.BCC = $Bcc }
    $mail.Subject = "$SubjectPrefix $($today.ToString('dd MMM yyyy'))"
    $mail.HTMLBody = $finalHtml

    Write-Log "Sending mail to $To..."
    $mail.Send()
    Write-Log "Mail sent successfully. Check Sent Items."

    Write-Log "=== END no-bullets run (success) ==="
    exit 0
}
catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message)
    Write-Log "=== END no-bullets run (failure) ==="
    throw
}
