<#!
email-todays-worklog-no-bullets.ps1

Sends today's Work Log email (plain dash-separated lines)
via Outlook with Calibri 11pt styling and static red disclaimer.

Usage:
.\email-todays-worklog-no-bullets.ps1 -WorkLogsBase "E:\SAPTARSHI BANERJEE\WORKLOGS" -To "recipient@domain.com"
#>

param(
  [string]$WorkLogsBase = "E:\SAPTARSHI BANERJEE\PERSONAL\PROJECT\OTHERS\WORKLOGS",
  [Parameter(Mandatory = $true)][string]$To,
  [string]$Cc = "",
  [string]$Bcc = "",
  [string]$SubjectPrefix = "Work Log -",
  [string]$LogFile = "$PSScriptRoot\emaillogs.log",
  [switch]$PreviewOnly,
  [string]$SummaryPath
)

if (-not $SummaryPath) {
  $SummaryPath = Join-Path $WorkLogsBase 'summary.xlsx'
}

$SummaryPath = [System.IO.Path]::GetFullPath($SummaryPath)

function Write-Log {
  param($m)
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  "$ts`t$m" | Out-File -FilePath $LogFile -Append -Encoding utf8
  Write-Host $m
}

function Update-SummaryWorkbook {
  param(
    [string]$Path,
    [datetime]$LogDate,
    [System.Collections.Generic.List[object]]$Sections
  )

  $excel = $null
  $workbook = $null
  $worksheet = $null
  $entriesWritten = 0
  $xlUp = -4162

  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $createdNew = $false

  if ((Test-Path -LiteralPath $Path -PathType Leaf) -and ((Get-Item -LiteralPath $Path).Length -gt 0)) {
      try {
        $workbook = $excel.Workbooks.Open($Path)
      }
      catch {
        Write-Log ("Failed to open existing summary workbook ({0}); recreating." -f $_.Exception.Message)
      }
    }

    if (-not $workbook) {
      $createdNew = $true
      $workbook = $excel.Workbooks.Add()
    }

    $worksheet = $workbook.Worksheets.Item(1)

    $headers = @('Date','Section','Entry')
    $needsHeader = $createdNew
    if (-not $needsHeader) {
      for ($i = 0; $i -lt $headers.Count; $i++) {
        $currentHeader = $worksheet.Cells.Item(1, $i + 1).Value2
        if ($currentHeader -ne $headers[$i]) {
          $needsHeader = $true
          break
        }
      }
    }

    if ($needsHeader) {
      for ($i = 0; $i -lt $headers.Count; $i++) {
        $cell = $worksheet.Cells.Item(1, $i + 1)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
      }
      $worksheet.Columns.Item(1).ColumnWidth = 18
      $worksheet.Columns.Item(2).ColumnWidth = 28
      $worksheet.Columns.Item(3).ColumnWidth = 80
    }

    $lastRow = $worksheet.Cells.Item($worksheet.Rows.Count, 1).End($xlUp).Row
    if ($lastRow -lt 1) { $lastRow = 1 }
    $currentRow = $lastRow + 1

    foreach ($section in $Sections) {
      foreach ($item in $section.Items) {
        $worksheet.Cells.Item($currentRow, 1).Value2 = $LogDate.ToString('yyyy-MM-dd')
        $worksheet.Cells.Item($currentRow, 2).Value2 = $section.Title
        $worksheet.Cells.Item($currentRow, 3).Value2 = $item
        $currentRow++
        $entriesWritten++
      }
    }

    if ($entriesWritten -eq 0 -and $createdNew) {
      $workbook.SaveAs($Path, 51)
      return 0
    }

    if ($entriesWritten -gt 0) {
      if ($createdNew) {
        $workbook.SaveAs($Path, 51)
      }
      else {
        $workbook.Save()
      }
    }

    return $entriesWritten
  }
  finally {
    if ($worksheet) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    }
    if ($workbook) {
      try { $workbook.Close($false) } catch {}
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel) {
      $excel.Quit()
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}

try {
  if (Test-Path $LogFile) { Remove-Item $LogFile -ErrorAction SilentlyContinue }
  Write-Log "=== START no-bullets run ==="

  $today = Get-Date
  $filePath = Join-Path (Join-Path $WorkLogsBase $today.ToString('yyyy')) (Join-Path $today.ToString('MMMM') ($today.ToString('yyyy-MM-dd') + ".md"))
  Write-Log "Looking for: $filePath"

  if (-not (Test-Path $filePath)) {
    throw "Worklog file not found: $filePath"
  }

  $content = Get-Content -Raw -LiteralPath $filePath -Encoding UTF8
  try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch {}

  $headingRegex = [regex]'^(#{2,6})\s+(.*\S)'
  $bulletRegex = [regex]'^\s*(?:-|(?:\d+\.)|[a-zA-Z]\.)\s*(.+)$'
  $sections = @()
  $currentSection = $null

  foreach ($rawLine in ($content -split "`r?`n")) {
    $trimmed = $rawLine.Trim()

    $headingMatch = $headingRegex.Match($trimmed)
    if ($headingMatch.Success) {
      if ($currentSection) { $sections += $currentSection }
      $currentSection = [PSCustomObject]@{
        Title = $headingMatch.Groups[2].Value.Trim()
        Items = New-Object System.Collections.Generic.List[string]
      }
      continue
    }

    if (-not $currentSection) { continue }

    $bulletMatch = $bulletRegex.Match($rawLine)
    if ($bulletMatch.Success) {
      $itemText = $bulletMatch.Groups[1].Value.Trim()
      if ($itemText) { $currentSection.Items.Add($itemText) }
    }
  }

  if ($currentSection) { $sections += $currentSection }

  $validSections = $sections | Where-Object { $_.Items.Count -gt 0 }

  Write-Log ("Sections with content: {0}" -f ($validSections.Count))
  if ($validSections.Count -eq 0) {
    Write-Log "No sections with bullet content found. Nothing to send."
    exit 0
  }

  Write-Log ("Summary workbook target: {0}" -f $SummaryPath)

  $sectionHtmlParts = foreach ($section in $validSections) {
    $titleHtml = [System.Web.HttpUtility]::HtmlEncode($section.Title)
    $itemsHtml = ($section.Items | ForEach-Object {
      '- ' + [System.Web.HttpUtility]::HtmlEncode($_)
    }) -join "<br/>`n"

    '<div style="margin-bottom:16px;"><p style="margin:0 0 6px 0;font-size:11pt;font-weight:bold;color:#003366;">{0}</p><p style="margin:0 0 12px 24px;font-size:11pt;">{1}</p></div>' -f $titleHtml, $itemsHtml
  }

  $taskHtml = $sectionHtmlParts -join "`n"

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

  $tempHtml = Join-Path $Env:TEMP "worklog-composed-no-bullets.html"
  $finalHtml | Out-File -FilePath $tempHtml -Encoding UTF8
  Write-Log ("Composed HTML saved to: {0}" -f $tempHtml)

  $outlook = $null
  $mail = $null
  try {
    Write-Log "Creating Outlook COM object..."
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    $mail.To = $To
    if ($Cc) { $mail.CC = $Cc }
    if ($Bcc) { $mail.BCC = $Bcc }
    $mail.Subject = "$SubjectPrefix $($today.ToString('dd MMM yyyy'))"
    $mail.HTMLBody = $finalHtml

    if ($PreviewOnly) {
      Write-Log "PreviewOnly mode requested. Skipping send but continuing with summary update."
    }
    else {
      Write-Log ("Sending mail to {0}..." -f $To)
      $mail.Send()
      Write-Log "Mail sent successfully. Check Sent Items."
    }
  }
  finally {
    if ($mail) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
    if ($outlook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
  }

  try {
    Write-Log "Updating summary workbook..."
    $rowsAdded = Update-SummaryWorkbook -Path $SummaryPath -LogDate $today -Sections $validSections
    Write-Log ("Summary workbook updated. Rows added: {0}" -f $rowsAdded)
  }
  catch {
    Write-Log ("ERROR updating summary workbook: {0}" -f $_.Exception.Message)
    throw
  }

  try {
      Write-Log "Checking for Git changes..."
      $gitStatus = git status --porcelain
      if ($gitStatus) {
          Write-Log "Changes detected. Committing and pushing to Git..."
          git add .
          $commitMsg = "Worklog update: $($today.ToString('yyyy-MM-dd'))"
          git commit -m $commitMsg
          git push origin main
          Write-Log "Git push successful."
      } else {
          Write-Log "No changes to commit."
      }
  }
  catch {
      Write-Log ("WARNING: Git operation failed: {0}" -f $_)
  }

  Write-Log "=== END no-bullets run (success) ==="
  exit 0
}
catch {
  Write-Log ("ERROR: {0}" -f $_.Exception.Message)
  Write-Log "=== END no-bullets run (failure) ==="
  throw
}
