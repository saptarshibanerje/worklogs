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

$script:RunId = [Guid]::NewGuid().ToString("N").Substring(0, 8)

function Write-Log {
  param($m)
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $logDir = Split-Path -Parent $LogFile
  if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
  }
  "$ts`t[$script:RunId] $m" | Out-File -FilePath $LogFile -Append -Encoding utf8
  Write-Host $m
}

function ConvertTo-DateKey {
  param([datetime]$Date)
  return $Date.ToString('yyyy-MM-dd')
}

function Parse-WorkLogSections {
  param(
    [string]$Content
  )

  $headingRegex = [regex]'^(#{2,6})\s+(.*\S)'
  $bulletRegex = [regex]'^\s*(?:-|(?:\d+\.)|[a-zA-Z]\.)\s*(.+)$'
  $sections = New-Object System.Collections.Generic.List[object]
  $currentSection = $null

  foreach ($rawLine in ($Content -split "`r?`n")) {
    $trimmed = $rawLine.Trim()

    $headingMatch = $headingRegex.Match($trimmed)
    if ($headingMatch.Success) {
      if ($currentSection) { $sections.Add($currentSection) }
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

  if ($currentSection) { $sections.Add($currentSection) }
  return $sections
}

function Get-SectionsFromWorkLogFile {
  param(
    [string]$Path
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    return $null
  }

  $content = Get-Content -Raw -LiteralPath $Path -Encoding UTF8
  if (-not $content) {
    return @()
  }

  return (Parse-WorkLogSections -Content $content)
}

function Convert-SectionsToRecords {
  param(
    [datetime]$LogDate,
    [System.Collections.Generic.List[object]]$Sections
  )

  $records = New-Object System.Collections.Generic.List[object]
  $dateOnly = $LogDate.Date
  $dateKey = ConvertTo-DateKey -Date $dateOnly

  foreach ($section in $Sections) {
    $title = [string]$section.Title
    foreach ($item in $section.Items) {
      $records.Add([PSCustomObject]@{
        Date = $dateOnly
        DateKey = $dateKey
        Section = $title
        Entry = [string]$item
      })
    }
  }

  return $records
}

function Get-WorkLogFileMap {
  param(
    [string]$BasePath
  )

  $map = @{}
  if (-not (Test-Path -LiteralPath $BasePath)) {
    return $map
  }

  $logFiles = Get-ChildItem -Path $BasePath -Recurse -File -Filter "*.md" -ErrorAction SilentlyContinue
  foreach ($file in $logFiles) {
    if ($file.Name -match '^(?<date>\d{4}-\d{2}-\d{2})\.md$') {
      $dateKey = $matches['date']
      if (-not $map.ContainsKey($dateKey)) {
        $map[$dateKey] = $file.FullName
      }
    }
  }

  return $map
}

function Sync-SummaryWorkbook {
  param(
    [string]$Path,
    [datetime]$Today,
    [System.Collections.Generic.List[object]]$TodaySections,
    [string]$WorkLogsBase
  )

  $excel = $null
  $workbook = $null
  $worksheet = $null
  $xlUp = -4162
  $existingDateKeys = New-Object System.Collections.Generic.HashSet[string]
  $recordsByDate = New-Object 'System.Collections.Generic.SortedDictionary[datetime,System.Collections.Generic.List[pscustomobject]]'

  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $summaryDir = Split-Path -Parent $Path
    if ($summaryDir -and -not (Test-Path -LiteralPath $summaryDir)) {
      New-Item -ItemType Directory -Path $summaryDir -Force | Out-Null
    }

    $createdNew = $false

    if ((Test-Path -LiteralPath $Path -PathType Leaf) -and ((Get-Item -LiteralPath $Path).Length -gt 0)) {
      try {
        Write-Log ("Opening existing summary workbook: {0}" -f $Path)
        $workbook = $excel.Workbooks.Open($Path)
      }
      catch {
        Write-Log ("Failed to open existing summary workbook ({0}); recreating." -f $_.Exception.Message)
      }
    }

    if (-not $workbook) {
      $createdNew = $true
      Write-Log ("Creating new summary workbook at: {0}" -f $Path)
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
      $worksheet.Columns.Item(1).NumberFormat = "yyyy-mm-dd"
    }

    $lastRow = $worksheet.Cells.Item($worksheet.Rows.Count, 1).End($xlUp).Row
    if ($lastRow -gt 1) {
      $endAddress = "C$lastRow"
      $dataRange = $worksheet.Range("A2", $endAddress)
      $data = $dataRange.Value2

      if ($null -ne $data) {
        if ($data -is [System.Array]) {
          if ($data.Rank -eq 1) {
            $cols = $data.Length
            $matrix = New-Object 'object[,]' 1, $cols
            for ($c = 0; $c -lt $cols; $c++) {
              $matrix[0, $c] = $data[$c]
            }
            $data = $matrix
          }
        }
        else {
          $matrix = New-Object 'object[,]' 1, 1
          $matrix[0, 0] = $data
          $data = $matrix
        }

        $rowLower = $data.GetLowerBound(0)
        $rowUpper = $data.GetUpperBound(0)
        $colLower = $data.GetLowerBound(1)

        for ($row = $rowLower; $row -le $rowUpper; $row++) {
          $dateValue = $data.GetValue($row, $colLower)
          if ($null -eq $dateValue) { continue }

          if ($dateValue -is [double]) {
            $dateValue = [DateTime]::FromOADate($dateValue)
          } elseif ($dateValue -is [string]) {
            try {
              $dateValue = [DateTime]::Parse($dateValue, [System.Globalization.CultureInfo]::InvariantCulture)
            }
            catch {
              continue
            }
          } else {
            $dateValue = [datetime]$dateValue
          }

          $sectionValue = $data.GetValue($row, $colLower + 1)
          $entryValue = $data.GetValue($row, $colLower + 2)
          $dateOnly = $dateValue.Date
          $dateKey = ConvertTo-DateKey -Date $dateOnly

          $existingDateKeys.Add($dateKey) | Out-Null
          if (-not $recordsByDate.ContainsKey($dateOnly)) {
            $recordsByDate[$dateOnly] = New-Object System.Collections.Generic.List[pscustomobject]
          }

          $recordsByDate[$dateOnly].Add([pscustomobject]@{
            Date = $dateOnly
            DateKey = $dateKey
            Section = [string]$sectionValue
            Entry = [string]$entryValue
          })
        }
      }
    }

    $todayKey = ConvertTo-DateKey -Date $Today
    if ($recordsByDate.ContainsKey($Today.Date)) {
      $removedCount = $recordsByDate[$Today.Date].Count
      Write-Log ("Removing {0} existing entries for {1} before refresh." -f $removedCount, $todayKey)
      $recordsByDate.Remove($Today.Date) | Out-Null
    }

    $fileMap = Get-WorkLogFileMap -BasePath $WorkLogsBase
    $missingDateKeys = @()
    foreach ($kvp in $fileMap.GetEnumerator()) {
      if (-not $existingDateKeys.Contains($kvp.Key)) {
        $missingDateKeys += $kvp.Key
      }
    }

    $missingDateKeys = $missingDateKeys | Sort-Object
    foreach ($dateKey in $missingDateKeys) {
      $filePath = $fileMap[$dateKey]
      Write-Log ("Backfilling summary for {0} from file {1}" -f $dateKey, $filePath)
      $sections = Get-SectionsFromWorkLogFile -Path $filePath
      if (-not $sections -or $sections.Count -eq 0) {
        Write-Log ("No bullet entries found in {0}; skipping backfill." -f $filePath)
        continue
      }

      $logDate = [DateTime]::ParseExact($dateKey, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
      $records = Convert-SectionsToRecords -LogDate $logDate -Sections $sections
      if ($records.Count -eq 0) {
        Write-Log ("Parsed sections for {0} but found no bullet entries to record." -f $dateKey)
        continue
      }

      if (-not $recordsByDate.ContainsKey($logDate.Date)) {
        $recordsByDate[$logDate.Date] = New-Object System.Collections.Generic.List[pscustomobject]
      }

      foreach ($record in $records) {
        $recordsByDate[$logDate.Date].Add($record)
      }

      $existingDateKeys.Add($dateKey) | Out-Null
    }

    $todayRecords = Convert-SectionsToRecords -LogDate $Today -Sections $TodaySections
    Write-Log ("Adding {0} entries for today's log ({1})." -f $todayRecords.Count, $todayKey)
    $recordsByDate[$Today.Date] = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($record in $todayRecords) {
      $recordsByDate[$Today.Date].Add($record)
    }
    $existingDateKeys.Add($todayKey) | Out-Null

    $totalEntries = 0
    foreach ($kvp in $recordsByDate.GetEnumerator()) {
      $totalEntries += $kvp.Value.Count
    }

    Write-Log ("Writing {0} total entries back to summary workbook." -f $totalEntries)

    if ($lastRow -gt 1) {
      $worksheet.Range("A2", "C$lastRow").ClearContents() | Out-Null
    }

    if ($totalEntries -gt 0) {
      $dataArray = New-Object 'object[,]' $totalEntries, 3
      $rowIndex = 0
      foreach ($kvp in $recordsByDate.GetEnumerator()) {
        foreach ($record in $kvp.Value) {
          $dataArray[$rowIndex, 0] = $record.Date
          $dataArray[$rowIndex, 1] = $record.Section
          $dataArray[$rowIndex, 2] = $record.Entry
          $rowIndex++
        }
      }

      $writeRange = $worksheet.Range("A2").Resize($totalEntries, 3)
      $writeRange.Value2 = $dataArray
      $writeRange.Columns.Item(1).NumberFormat = "yyyy-mm-dd"
    }

    if ($createdNew) {
      $workbook.SaveAs($Path, 51)
    }
    else {
      $workbook.Save()
    }

    return $totalEntries
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
  Write-Log "=== START no-bullets run ==="
  Write-Log ("Run context -> WorkLogsBase: {0}; SummaryPath: {1}" -f $WorkLogsBase, $SummaryPath)

  $today = Get-Date
  $filePath = Join-Path (Join-Path $WorkLogsBase $today.ToString('yyyy')) (Join-Path $today.ToString('MMMM') ($today.ToString('yyyy-MM-dd') + ".md"))
  Write-Log "Looking for: $filePath"

  if (-not (Test-Path $filePath)) {
    throw "Worklog file not found: $filePath"
  }

  $content = Get-Content -Raw -LiteralPath $filePath -Encoding UTF8
  try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch {}

  Write-Log "Parsing worklog content for bullet sections..."
  $sections = Parse-WorkLogSections -Content $content
  $validSections = New-Object System.Collections.Generic.List[object]
  foreach ($section in $sections) {
    if ($section.Items.Count -gt 0) {
      $validSections.Add($section)
    }
  }

  Write-Log ("Sections with content: {0}" -f $validSections.Count)
  $sendMail = $true
  if ($validSections.Count -eq 0) {
    Write-Log "No sections with bullet content found. Email delivery will be skipped, continuing with summary sync."
    $sendMail = $false
  }

  Write-Log ("Summary workbook target: {0}" -f $SummaryPath)

  if ($sendMail) {
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
  }
  else {
    Write-Log "Skipping Outlook email step because no actionable items were found."
  }

  try {
    Write-Log "Synchronizing summary workbook (dedupe + backfill)..."
    $totalEntries = Sync-SummaryWorkbook -Path $SummaryPath -Today $today -TodaySections $validSections -WorkLogsBase $WorkLogsBase
    Write-Log ("Summary workbook synchronization complete. Total entries now: {0}" -f $totalEntries)
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
