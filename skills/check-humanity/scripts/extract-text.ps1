<#
  Read the text of a user-selected .docx, .txt, or .pdf file for check-humanity.
  Writes to stdout only; never creates or modifies files.

  -Mode Outline  headings, paragraph numbers, and size only (no body text)
  -Mode Text     text of the selection (default: whole file)
  -Mode Marks    like -Mode Text (same -Section / -Paragraphs / -Body / -Pages / -Cells), but prints only the
                 paragraphs / table cells (PDF: the lines) that contain an arrow (U+2192), an em dash (U+2014)
                 or an en dash (U+2013), with the same coordinate labels as -Mode Text ([¶N ...], rRcC / pK).
                 A dash with a digit on each side (spaces allowed), for example
                 1–12 or 6–9 ตัวอักษร, is a number/date range and is not counted. The last
                 line is a summary: "สรุป: พบ N จุด ใน M ย่อหน้า/เซลล์ (→ a, — b, – c)". Creates no file. Use it
                 to find every such mark in the selection before building a review list or a rewrite.
  -Mode Locate   (.docx / .txt; needs -Paragraphs "79,831") tells where each listed ¶ is, the way a person finds it in Word:
                 one line per ¶, "¶42: บทที่ 1 › คำตอบรายข้อ Section 1 › 1.3.3 | ค้น: <text>". The path is the headings above
                 the ¶ by the numbers the document shows (and an item number written in the body, e.g. 5.2.25); "ค้น:" is a
                 short text from the start of the paragraph to paste into Word's Ctrl+F, with " (พบ N ที่)" when it is
                 found N times in the document (table of contents included). A table: the path and the table label
                 ("ตาราง “ตารางที่ 3”") and the text of its first cell that has text. A ¶ that is not in the file:
                 "¶N: ไม่พบในไฟล์". Use it to name a place for the user in the chat (the ¶ number alone means nothing in
                 Word). Creates no file. Not for a .pdf (use the page number). -Section / -Body are not used.
  -Section 2     section S2 from the outline (until the next heading of the same or higher level)
  -Paragraphs "5-9,12"   paragraph numbers from the outline
  -Body "2" or "1-3"   body paragraphs by order within -Section, or whole file if no section
  -Pages "3-5"   PDF pages
  .docx Text mode: a picture that sits between two numbered ¶ but has no text of its own (an
                 image-only paragraph) has no ¶ number; a line "[รูป ไม่มีข้อความ]" is printed where it
                 sits (right before the next selected ¶), without a ¶ number. It is not counted as text
                 and no other ¶/body number changes. -Mode Outline is unchanged.
  -Cells         (.docx Text mode only) print rRcC/pK coordinates for each table cell, so review-file
                 items and apply-revisions JSON can target a specific table cell/paragraph. Ignored
                 for .txt/.pdf and in -Mode Outline.
  .docx Text mode: a line break (w:br) is shown as one space. apply-revisions.ps1 edits a paragraph
                 with a line break, a special hyphen, a hyperlink or a field word by word and keeps each of
                 them where it is; it skips the paragraph (with a reason) only when a change would touch one
                 of them. No label is needed for these paragraphs: ¶ numbers and every label are the plain ones.
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [ValidateSet('Outline', 'Text', 'Marks', 'Locate')][string]$Mode = 'Text',
  [string]$Section,
  [string]$Paragraphs,
  [string]$Body,
  [string]$Pages,
  [switch]$Cells
)

. (Join-Path $PSScriptRoot 'doc-lib.ps1')

function Show-Outline($items, [string]$name) {
  $chars = ($items | ForEach-Object { $_.Text.Length } | Measure-Object -Sum).Sum
  if (-not $chars) { $chars = 0 }
  $heads = @($items | Where-Object { $_.Level -gt 0 })
  $tables = @($items | Where-Object { $_.Kind -eq 'table' }).Count
  $captions = @($items | Where-Object { $_.Caption }).Count
  $bodyCount = @($items | Where-Object { Test-Body $_ }).Count
  Write-Output "ไฟล์: $name"
  $tablePart = if ($captions -gt 0) { "(ตาราง $tables, คำบรรยายภาพ $captions)" } else { "(ตาราง $tables)" }
  Write-Output "ย่อหน้า: $($items.Count) ¶ $tablePart | เนื้อความ: $bodyCount | ตัวอักษรประมาณ: $chars | หัวข้อ: $($heads.Count)"
  if ($script:skippedToc) { Write-Output "ข้าม: สารบัญ $($script:skippedToc) บรรทัด" }
  if ($script:skippedTextBoxes) { Write-Output "ข้าม: text box $($script:skippedTextBoxes) กล่อง" }
  if ($script:encodingNote) { Write-Output "หมายเหตุ: $($script:encodingNote)" }
  if (-not $heads.Count) { Write-Output 'ไม่พบหัวข้อ ให้เลือกเป็นเลข ¶ หรือเนื้อความลำดับที่'; return }
  $i = 0
  foreach ($h in $heads) {
    $i++
    $end = Get-SectionEnd $items $h
    $bCount = @($items | Where-Object { $_.N -ge $h.N -and $_.N -le $end -and (Test-Body $_) }).Count
    $title = $h.Text -replace '\s+', ' '
    if ($title.Length -gt 80) { $title = $title.Substring(0, 80) + '…' }
    Write-Output ("[S{0}] {1}H{2} ¶{3}-{4} เนื้อความ {5}  {6}" -f $i, ('  ' * ($h.Level - 1)), $h.Level, $h.N, $end, $bCount, $title)
  }
}

# -Mode Marks: counts arrows (U+2192), em dashes (U+2014) and en dashes (U+2013) in a text. A dash with a digit on each
# side (spaces and tabs allowed between) is a number/date range: it is counted in Range, not in Em/En.
function Get-MarkInfo([string]$text) {
  $arrow = 0; $em = 0; $en = 0; $range = 0
  if ($text) {
    for ($i = 0; $i -lt $text.Length; $i++) {
      $code = [int]$text[$i]
      if ($code -eq 0x2192) { $arrow++; continue }
      if ($code -ne 0x2014 -and $code -ne 0x2013) { continue }
      $l = $i - 1
      while ($l -ge 0 -and ($text[$l] -eq ' ' -or $text[$l] -eq "`t" -or [int]$text[$l] -eq 0xA0)) { $l-- }
      $r = $i + 1
      while ($r -lt $text.Length -and ($text[$r] -eq ' ' -or $text[$r] -eq "`t" -or [int]$text[$r] -eq 0xA0)) { $r++ }
      if ($l -ge 0 -and $r -lt $text.Length -and [char]::IsDigit($text[$l]) -and [char]::IsDigit($text[$r])) { $range++; continue }
      if ($code -eq 0x2014) { $em++ } else { $en++ }
    }
  }
  return [pscustomobject]@{ Arrow = $arrow; Em = $em; En = $en; Range = $range; Total = ($arrow + $em + $en) }
}

# Adds one paragraph/cell/line to the -Mode Marks totals; returns $true when it has at least one mark to show.
function Add-MarkHit($info) {
  $script:markRange += $info.Range
  if ($info.Total -eq 0) { return $false }
  $script:markArrow += $info.Arrow; $script:markEm += $info.Em; $script:markEn += $info.En
  $script:markUnits++
  return $true
}

function Write-MarkSummary {
  $total = $script:markArrow + $script:markEm + $script:markEn
  $arrow = [string][char]0x2192; $em = [string][char]0x2014; $en = [string][char]0x2013
  $note = ''
  if ($script:markRange -gt 0) { $note = " (ไม่นับขีดที่บอกช่วงตัวเลข $($script:markRange) จุด)" }
  if ($total -eq 0) {
    Write-Output "สรุป: ไม่พบ $arrow $em $en ในช่วงที่เลือก$note"
  } else {
    Write-Output "สรุป: พบ $total จุด ใน $($script:markUnits) ย่อหน้า/เซลล์ ($arrow $($script:markArrow), $em $($script:markEm), $en $($script:markEn))$note"
  }
}

function Show-Text($items) {
  if ($Paragraphs -and $Body) { Fail 'ใช้ -Paragraphs หรือ -Body อย่างใดอย่างหนึ่ง' }
  $isMarks = ($Mode -eq 'Marks')
  $script:markArrow = 0; $script:markEm = 0; $script:markEn = 0; $script:markRange = 0; $script:markUnits = 0
  $selected = $items
  $sectionLabel = $null
  $idx = $null
  if ($Section) {
    if ($Section -notmatch '^[Ss]?(\d+)$') { Fail "Section ต้องเป็นเลขจาก outline เช่น 2 หรือ S2" }
    $heads = @($items | Where-Object { $_.Level -gt 0 })
    $idx = [int]$Matches[1]
    if ($idx -lt 1 -or $idx -gt $heads.Count) { Fail "ไม่มีหัวข้อ S$idx (มี $($heads.Count) หัวข้อ)" }
    $h = $heads[$idx - 1]
    $end = Get-SectionEnd $items $h
    $selected = @($items | Where-Object { $_.N -ge $h.N -and $_.N -le $end })
    $sectionLabel = "S$idx"
  }

  # Number body items within the current scope (section, or whole document).
  $bodyMap = @{}
  $bodyList = New-Object System.Collections.ArrayList
  $k = 0
  foreach ($it in $selected) {
    if (Test-Body $it) { $k++; $bodyMap[$it.N] = $k; [void]$bodyList.Add($it.N) }
  }

  if ($Paragraphs) {
    $set = Parse-Ranges $Paragraphs
    $selected = @($selected | Where-Object { $set.Contains($_.N) })
    if (-not @($selected).Count) { Fail 'ไม่พบย่อหน้าในช่วงที่เลือก' }
  } elseif ($Body) {
    $set = Parse-Ranges $Body
    $cnt = $bodyList.Count
    foreach ($bi in $set) {
      if ($bi -lt 1 -or $bi -gt $cnt) {
        if ($sectionLabel -and $cnt -eq 0) {
          Fail "$sectionLabel ไม่มีย่อหน้าเนื้อความ (มีเฉพาะหัวข้อ ตาราง หรือคำบรรยายภาพ) ให้เลือกหัวข้ออื่นหรือใช้ -Paragraphs"
        } elseif ($sectionLabel) {
          $list = ($bodyList | ForEach-Object { "¶$_" }) -join ', '
          Fail "ไม่มีเนื้อความลำดับที่ $bi ใน $sectionLabel (มี $cnt ย่อหน้า: $list)"
        } else {
          Fail "ไม่มีเนื้อความลำดับที่ $bi ในเอกสาร (มี $cnt ย่อหน้า)"
        }
      }
    }
    $chosenNs = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($n in $bodyList) { if ($set.Contains($bodyMap[$n])) { [void]$chosenNs.Add($n) } }
    $selected = @($selected | Where-Object { $chosenNs.Contains($_.N) })
    if (-not @($selected).Count) { Fail 'ไม่พบย่อหน้าในช่วงที่เลือก' }
    $pairs = ($selected | ForEach-Object { "¶$($_.N)" }) -join ', '
    if ($sectionLabel) {
      Write-Output "เลือก: $sectionLabel เนื้อความ $Body = $pairs"
    } else {
      Write-Output "เลือก: เนื้อความ $Body = $pairs"
    }
  }

  if (-not @($selected).Count) { Fail 'ไม่พบย่อหน้าในช่วงที่เลือก' }
  # .docx: pictures that sit between two numbered ¶ but carry no text have no ¶ number of their own.
  # Show where they are with a marker line (no ¶ number, not counted as text, no other number moves).
  # Skipped for the section's own heading under -Section: a picture right before it belongs to the
  # previous part of the document, not to this section.
  $imageBefore = @{}
  if ($ext -eq '.docx' -and $script:DocxNs) {
    foreach ($g in (Get-DocxGapGroups $items $script:DocxNs)) {
      if ($g.HasImage) { $imageBefore[[int]$g.NextN] = 1 }
    }
  }
  foreach ($it in $selected) {
    if (-not $isMarks -and $imageBefore.ContainsKey([int]$it.N) -and -not ($Section -and $it.N -eq $h.N)) {
      Write-Output '[รูป ไม่มีข้อความ]'
      Write-Output ''
    }
    $tag =if ($it.Kind -eq 'table') { "[¶$($it.N) ตาราง]" }
      elseif ($it.Level -gt 0) { "[¶$($it.N) H$($it.Level)]" }
      elseif ($it.Caption) { "[¶$($it.N) คำบรรยายภาพ]" }
      elseif ($bodyMap.ContainsKey($it.N)) { "[¶$($it.N) เนื้อความ $($bodyMap[$it.N])]" }
      else { "[¶$($it.N)]" }
    if ($it.Kind -eq 'table' -and $Cells) {
      if ($isMarks) {
        # print the table tag only when at least one of its cells has a mark; blank line only after that
        Show-Cells $it.Node -Tag $tag -OnlyMarks
        if (-not $script:cellsShown) { continue }
      } else {
        Write-Output $tag
        Show-Cells $it.Node
      }
    } elseif ($isMarks) {
      if (-not (Add-MarkHit (Get-MarkInfo $it.Text))) { continue }
      Write-Output "$tag $($it.Text)"
    } else {
      Write-Output "$tag $($it.Text)"
    }
    Write-Output ''
  }
  if ($isMarks) { Write-MarkSummary }
}

# -Cells: print each table cell's rRcC[pK] coordinate and text (docx Text mode only, see Get-TableCells).
# -OnlyMarks (-Mode Marks): print only the cells that have a mark, preceded by -Tag once. Sets $script:cellsShown to
# $true when any line was printed (the caller then prints the blank line that follows a table).
function Show-Cells($tbl, [string]$Tag, [switch]$OnlyMarks) {
  $lines = New-Object System.Collections.ArrayList
  foreach ($c in (Get-TableCells $tbl)) {
    if ($c.VCont) {
      if (-not $c.Text) { continue }
      [void]$lines.Add(@{ Text = $c.Text; Line = "r$($c.Row)c$($c.Cell) (ผสานแนวตั้ง แก้ไม่ได้): $($c.Text)" })
      continue
    }
    if ($c.Nested) {
      if (-not $c.Text) { continue }
      [void]$lines.Add(@{ Text = $c.Text; Line = "r$($c.Row)c$($c.Cell) (มีตารางซ้อน แก้ในไฟล์ไม่ได้): $($c.Text)" })
      continue
    }
    if ($c.ParaCount -eq 0) { continue }
    $label = "r$($c.Row)c$($c.Cell)"
    if ($c.Para) { $label += "p$($c.Para)" }
    $suffix = ''
    if ($c.Span -gt 1) { $suffix = " (กว้าง $($c.Span) คอลัมน์)" }
    [void]$lines.Add(@{ Text = $c.Text; Line = ("{0}{1}: {2}" -f $label, $suffix, $c.Text) })
  }
  $shown = New-Object System.Collections.ArrayList
  foreach ($l in $lines) {
    if ($OnlyMarks -and -not (Add-MarkHit (Get-MarkInfo $l.Text))) { continue }
    [void]$shown.Add($l.Line)
  }
  $script:cellsShown = ($shown.Count -gt 0)
  if (-not $script:cellsShown) { return }
  if ($OnlyMarks) { Write-Output $Tag }
  foreach ($s in $shown) { Write-Output $s }
}

# -Mode Locate: one line per requested ¶ - the heading path and the Ctrl+F text (doc-lib.ps1 Get-ItemFinder, the same code
# review-file.ps1 writes the review / report with).
function Show-Locate($items) {
  $set = Parse-Ranges $Paragraphs
  if (-not $set.Count) { Fail 'ไม่พบเลข ¶ ใน -Paragraphs (ใช้รูปแบบ 5-9,12)' }
  $corpus = Get-FindCorpus $items
  foreach ($n in (@($set) | Sort-Object)) {
    $f = Get-ItemFinder $items $corpus $n $null $null $null ''
    if ($f.Error) { Write-Output "¶${n}: ไม่พบในไฟล์"; continue }
    $line = "¶${n}: $($f.Path)"
    if ($f.Quote) {
      $line += " | ค้น: $($f.Quote)"
      if ($f.Count -gt 1) { $line += " (พบ $($f.Count) ที่)" }
    }
    Write-Output $line
  }
}

function Run-Pdf([string]$file) {
  if ($Body -and $Mode -ne 'Outline') { Fail 'PDF เลือกได้เฉพาะหน้า (-Pages)' }
  $tool = Get-Command pdftotext -ErrorAction SilentlyContinue
  if (-not $tool) { Fail 'เครื่องนี้ไม่มีตัวอ่าน PDF (pdftotext) ให้คัดลอกข้อความจาก PDF มาวาง หรือเปิด PDF ด้วย Word แล้วบันทึกเป็น .docx' }
  $info = Get-Command pdfinfo -ErrorAction SilentlyContinue
  $pageCount = $null
  if ($info) {
    $line = (& $info.Source $file 2>$null | Where-Object { $_ -match '^Pages:\s+(\d+)' })
    if ($line -and ($line -match '(\d+)')) { $pageCount = [int]$Matches[1] }
  }
  $pdfArgs = @('-enc', 'UTF-8')
  if ($Mode -ne 'Outline' -and $Pages) {
    if ($Pages -notmatch '^\s*(\d+)\s*(-\s*(\d+))?\s*$') { Fail 'Pages ต้องเป็นหน้าเดียวหรือช่วงเดียว เช่น 3 หรือ 3-5' }
    $first = [int]$Matches[1]; $last = if ($Matches[3]) { [int]$Matches[3] } else { $first }
    $pdfArgs += @('-f', $first, '-l', $last)
  } else { $first = 1 }
  $raw = & $tool.Source @pdfArgs $file - 2>$null
  if ($LASTEXITCODE -ne 0) { Fail 'อ่าน PDF ไม่ได้ ไฟล์อาจมีรหัสผ่านหรือเสียหาย' }
  $text = ($raw -join "`n")
  $pagesText = $text -split "`f"
  $letters = ($text -replace '\s', '').Length
  if ($Mode -eq 'Outline') {
    if (-not $pageCount) { $pageCount = @($pagesText | Where-Object { $_.Trim() }).Count }
    Write-Output "ไฟล์: $(Split-Path $file -Leaf)"
    Write-Output "หน้า: $pageCount | ตัวอักษรประมาณ: $letters"
    if ($letters -lt 20) { Write-Output 'หมายเหตุ: แทบไม่มีข้อความ อาจเป็น PDF สแกน/รูปภาพ ซึ่งตรวจไม่ได้' }
    Write-Output 'PDF เลือกได้เป็นช่วงหน้า'
    return
  }
  if ($letters -lt 20) { Fail 'แทบไม่มีข้อความในช่วงนี้ อาจเป็น PDF สแกน/รูปภาพ ซึ่งตรวจไม่ได้' }
  $p = $first
  if ($Mode -eq 'Marks') {
    # PDF has no numbered paragraphs: print each page's lines that contain a mark under its [หน้า N] label
    $script:markArrow = 0; $script:markEm = 0; $script:markEn = 0; $script:markRange = 0; $script:markUnits = 0
    foreach ($pg in $pagesText) {
      $hitLines = @(($pg -split '\r?\n') | Where-Object { Add-MarkHit (Get-MarkInfo $_) })
      if ($hitLines.Count) { Write-Output "[หน้า $p]"; foreach ($hl in $hitLines) { Write-Output $hl.Trim() }; Write-Output '' }
      $p++
    }
    Write-MarkSummary
    return
  }
  foreach ($pg in $pagesText) {
    if ($pg.Trim()) { Write-Output "[หน้า $p]"; Write-Output $pg.Trim(); Write-Output '' }
    $p++
  }
}

$full = Resolve-Input $Path
$ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()

if ($Mode -eq 'Locate') {
  if ($ext -eq '.pdf') { Fail 'โหมด Locate ใช้กับ PDF ไม่ได้ ให้ใช้เลขหน้า (-Pages)' }
  if (-not $Paragraphs) { Fail 'โหมด Locate ต้องระบุ -Paragraphs เช่น -Paragraphs "79,831"' }
}

switch ($ext) {
  '.pdf' { Run-Pdf $full; exit 0 }
  '.docx' { $items = Read-Docx $full }
  '.txt' { $items = Read-Txt $full }
  '.md' { $items = Read-Txt $full }
  default { Fail "ยังไม่รองรับไฟล์ $ext รองรับ .docx, .pdf, .txt (ไฟล์ .doc แบบเก่าให้บันทึกเป็น .docx ก่อน)" }
}

if (-not $items.Count) { Fail 'ไม่พบข้อความในไฟล์' }
if ($Mode -eq 'Outline') { Show-Outline $items (Split-Path $full -Leaf) } elseif ($Mode -eq 'Locate') { Show-Locate $items } else { Show-Text $items }
