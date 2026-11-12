<#
  Write revised paragraphs to a NEW file for check-humanity. Never modifies or overwrites the source.

  -Path       source .docx, .txt, or .pdf
  -Revisions  UTF-8 JSON file:
                .docx/.txt: [{ "paragraph": 7, "row": 2, "cell": 1, "cell_paragraph": 2,
                               "original_prefix": "first 10+ chars as extracted",
                               "text": "revised paragraph", "items": ["R1","R3"] }]
                .pdf:       [{ "page": 3, "original": "original passage", "text": "revised passage" }]
              "row"/"cell" (see extract-text.ps1 -Cells for the rRcC coordinates) target a table
              cell; "cell_paragraph" is required when that cell has more than one paragraph.
              "items" lists the review-file R-ids folded into this one paragraph/cell edit.
  -Out        optional exact output path (must not exist).
  -OutDir     optional output folder (file goes directly inside it, de-duplicated with _2, _3, ...).
              Default (neither given): <output root>\humanized\<source base name>\
              <name>_humanized_<yyyy-MM-dd>.docx|.txt (PDF: .md under <output root>\review\).
              <output root> = <MyDocuments>\HumanizerAgent\output, or the folder in env
              HUMANIZER_OUTPUT_DIR when that is set. Nothing is ever written inside the skill's
              own folder (the folder that holds SKILL.md): such a path is refused.

  .docx output keeps the whole document and marks each change as a Word tracked change.
  Paragraph numbers are the same as extract-text.ps1.

  Tracked changes are word-level (only the words/spaces that changed can be accepted or rejected
  one by one). A paragraph whose runs have different formatting (bold in part, different language
  or complex-script settings, ...) is still edited word by word: the text of all runs is compared
  as one string, unchanged and deleted parts keep the formatting of the run they came from, and an
  inserted part takes the formatting of the run just before it (the next run at the very start).
  A paragraph (or table cell paragraph) with a line break (w:br), a special hyphen (w:noBreakHyphen /
  w:softHyphen), a hyperlink (w:hyperlink) or a field (w:fldSimple, or runs from w:fldChar begin to
  end) is edited word by word too, and every break, hyphen, link and field stays exactly where it is
  (a link/field keeps its text and its XML byte for byte; bookmarks stay in place). The text of a
  line break is a space in the extracted text, so a change that leaves that space alone keeps the
  break. When a change would touch a break, a special hyphen, or the text of a link/field, the
  paragraph is skipped with the reason ("ข้าม: ...") and is never rewritten as delete-the-whole-
  paragraph then insert (that used to lose the break or the link). This rule lives in doc-lib.ps1
  (Get-ParagraphUnitPlan); extract-text.ps1 needs no label for these paragraphs.
  Paragraphs with pictures, footnotes, comments, existing tracked changes or other content
  outside a link/field are skipped (reported as "ข้าม").
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$Revisions,
  [string]$Out,
  [string]$OutDir,
  [string]$Author = 'HumanizerAgent'
)

. (Join-Path $PSScriptRoot 'doc-lib.ps1')


function Norm([string]$s) { return (($s -replace '\s+', ' ').Trim()) }

function Load-Revisions {
  if (-not (Test-Path -LiteralPath $Revisions -PathType Leaf)) { Fail "ไม่พบไฟล์ revisions: $Revisions" }
  $json = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Revisions).ProviderPath, $utf8)
  try { $parsed = ConvertFrom-Json -InputObject $json } catch { Fail 'ไฟล์ revisions ไม่ใช่ JSON ที่ถูกต้อง' }
  # Windows PowerShell 5.1 returns a JSON array as a single object; enumerate it explicitly
  $list = New-Object System.Collections.ArrayList
  foreach ($x in $parsed) { [void]$list.Add($x) }
  if (-not $list.Count) { Fail 'ไม่มีรายการแก้ใน revisions' }
  return , $list
}

# Checks the item/text/prefix common to both a body paragraph and a table-cell paragraph.
# Returns $null (setting $script:reason) or @{ Item; Node; Text; Label }.
function Test-RevisionText($rev, [string]$label, [string]$origText, $targetItem, $targetNode) {
  $text = [string]$rev.text
  if (-not $text.Trim()) { $script:reason = "$label ไม่มีข้อความฉบับปรับ"; return $null }
  $prefix = Norm ([string]$rev.original_prefix)
  $orig = Norm $origText
  $need = [Math]::Min(10, $orig.Length)
  # Ordinal on purpose: culture-aware StartsWith treats Thai vowel/tone mark clusters as units, so a prefix cut inside a
  # cluster (for example "...สร" of "...สร้างได้") would wrongly report "no match".
  if ($prefix.Length -lt $need -or -not $orig.StartsWith($prefix, [StringComparison]::Ordinal)) { $script:reason = "$label original_prefix ไม่ตรงกับต้นฉบับ (เลขย่อหน้า/แถว/เซลล์อาจผิด)"; return $null }
  if ($text.Trim() -ceq $origText) { $script:reason = "$label ฉบับปรับเหมือนต้นฉบับ"; return $null }
  return @{ Item = $targetItem; Node = $targetNode; Text = $text; Label = $label }
}

# Checks a docx/txt revision (paragraph, or table row/cell/cell_paragraph) against the numbered
# items. Returns @{ Item; Node; Text; Label } or $null with a reason in $script:reason.
# $seen keys are strings: "9" for a body paragraph, "9 r2c1p1" for a table cell paragraph.
function Match-Item($items, $rev, [System.Collections.Generic.HashSet[string]]$seen) {
  $script:reason = $null
  $num = $rev.paragraph -as [int]
  if (-not $num) { $script:reason = 'ไม่มีเลข paragraph'; return $null }
  $item = $items | Where-Object { $_.N -eq $num } | Select-Object -First 1
  if (-not $item) { $script:reason = "ไม่มี ¶$num"; return $null }

  if ($null -ne $rev.items) {
    foreach ($rid in @($rev.items)) {
      if ([string]$rid -cnotmatch '^R\d{1,4}$') { $script:reason = "¶$num items ต้องเป็นรายการรหัส เช่น [`"R1`",`"R3`"]"; return $null }
    }
  }

  $hasRow = ($null -ne $rev.row); $hasCell = ($null -ne $rev.cell)
  if ($item.Kind -eq 'table') {
    if (-not $hasRow -or -not $hasCell) { $script:reason = "¶$num เป็นตาราง ต้องระบุ row และ cell"; return $null }
  } elseif ($hasRow -or $hasCell) { $script:reason = "¶$num ไม่ใช่ตาราง ไม่ต้องระบุ row/cell"; return $null }

  if ($item.Kind -ne 'table') {
    $label = "¶$num"; $key = "$num"
    if ($seen.Contains($key)) { $script:reason = "$label ซ้ำในรายการ"; return $null }
    $r = Test-RevisionText $rev $label $item.Text $item $item.Node
    if (-not $r) { return $null }
    [void]$seen.Add($key)
    return $r
  }

  $row = $rev.row -as [int]; $cell = $rev.cell -as [int]
  if (-not $row -or -not $cell -or $row -lt 1 -or $cell -lt 1) { $script:reason = "¶$num row/cell ต้องเป็นเลขตั้งแต่ 1"; return $null }
  $cells = Get-TableCells $item.Node
  $rows = @($cells | ForEach-Object { $_.Row } | Select-Object -Unique)
  $maxRow = if ($rows.Count) { ($rows | Measure-Object -Maximum).Maximum } else { 0 }
  if ($row -gt $maxRow) { $script:reason = "¶$num ไม่มีแถว $row (มี $maxRow แถว)"; return $null }
  $rowCells = @($cells | Where-Object { $_.Row -eq $row })
  $cellNums = @($rowCells | ForEach-Object { $_.Cell } | Select-Object -Unique)
  $maxCell = if ($cellNums.Count) { ($cellNums | Measure-Object -Maximum).Maximum } else { 0 }
  if ($cell -gt $maxCell) { $script:reason = "¶$num แถว $row ไม่มีเซลล์ $cell (มี $maxCell เซลล์)"; return $null }
  $cellRecs = @($rowCells | Where-Object { $_.Cell -eq $cell })
  $rec0 = $cellRecs[0]
  $baseLabel = "¶$num r${row}c${cell}"
  if ($rec0.VCont) { $script:reason = "$baseLabel เป็นส่วนต่อของเซลล์ผสานแนวตั้ง ให้แก้ที่เซลล์แรกของกลุ่ม"; return $null }
  if ($rec0.Nested) { $script:reason = "$baseLabel มีตารางซ้อน ยังไม่รองรับการแก้ในไฟล์"; return $null }
  if ($rec0.ParaCount -eq 0) { $script:reason = "$baseLabel เป็นเซลล์ว่าง"; return $null }

  $targetRec = $null
  if ($rec0.ParaCount -le 1) {
    $targetRec = $rec0
  } else {
    if ($null -eq $rev.cell_paragraph) { $script:reason = "$baseLabel มี $($rec0.ParaCount) ย่อหน้า ต้องระบุ cell_paragraph"; return $null }
    $cp = $rev.cell_paragraph -as [int]
    $targetRec = $cellRecs | Where-Object { $_.Para -eq $cp } | Select-Object -First 1
    if (-not $targetRec) { $script:reason = "$baseLabel ไม่มีย่อหน้าที่ $($rev.cell_paragraph) (มี $($rec0.ParaCount) ย่อหน้า)"; return $null }
  }
  $paraSuffix = if ($targetRec.Para) { "p$($targetRec.Para)" } else { '' }
  $label = "$baseLabel$paraSuffix"
  $key = "$num r${row}c${cell}p" + $(if ($targetRec.Para) { $targetRec.Para } else { 1 })
  if ($seen.Contains($key)) { $script:reason = "$label ซ้ำในรายการ"; return $null }
  $r = Test-RevisionText $rev $label $targetRec.Text $item $targetRec.Node
  if (-not $r) { return $null }
  [void]$seen.Add($key)
  return $r
}

$script:nextId = 0
$script:revDate = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

function New-W([string]$name) { return $script:DocxXml.CreateElement('w', $name, $W_NS) }

function New-Run($rPr, [string]$text, [bool]$deleted) {
  $r = New-W 'r'
  if ($rPr) { [void]$r.AppendChild($rPr.CloneNode($true)) }
  $parts = $text -split "(`t)"
  foreach ($part in $parts) {
    if ($part -eq '') { continue }
    if ($part -eq "`t") { [void]$r.AppendChild((New-W 'tab')); continue }
    $t = New-W $(if ($deleted) { 'delText' } else { 't' })
    [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
    $t.InnerText = $part
    [void]$r.AppendChild($t)
  }
  return $r
}

function New-Change([string]$kind, $run) {
  $c = New-W $kind
  [void]$c.SetAttribute('id', $W_NS, [string]$script:nextId); $script:nextId++
  [void]$c.SetAttribute('author', $W_NS, $Author)
  [void]$c.SetAttribute('date', $W_NS, $script:revDate)
  [void]$c.AppendChild($run)
  return $c
}

# $script:RevAllowedChildren / $script:RevAllowedInRun, Get-RunSegments and Get-WholeParagraphReason (the
# rule for "delete the whole paragraph, insert the new text") live in doc-lib.ps1, so extract-text.ps1
# can label the same paragraphs before an edit is chosen.

# One line per paragraph that had to be shown as delete-whole-paragraph then insert (see header). Only a paragraph
# whose run texts cannot be read back as the extracted text still ends here; a line break, a special hyphen, a
# hyperlink or a field never does (Revise-ParagraphByUnits keeps it or skips the paragraph).
$script:wholeParaNotes = New-Object System.Collections.ArrayList

# The right rPr for text inserted at old position $pos in a paragraph read as units: the run text just before it
# (never a link or field, so inserted text does not look like a link), the first run text at the very start.
function Get-UnitInsertRPr($units, [int]$pos) {
  $prev = $null; $firstUnit = $null
  foreach ($u in $units) {
    if ($u.Kind -ne 'text' -and $u.Kind -ne 'br') { continue }
    if ($null -eq $firstUnit) { $firstUnit = $u }
    if ($u.S -lt $pos) { $prev = $u }
  }
  if ($prev) { return $prev.RPr }
  if ($firstUnit) { return $firstUnit.RPr }
  return $null
}

# A run holding a copy of the break / hyphen element $node, with the formatting of the run it came from.
function New-UnitSpecialRun($rPr, $node) {
  $r = New-W 'r'
  if ($rPr) { [void]$r.AppendChild($rPr.CloneNode($true)) }
  [void]$r.AppendChild($node.CloneNode($true))
  return $r
}

# Adds a unit that takes no characters (a mark, a hyphen, a link/field without text) to $nodes, in its place.
function Add-UnitZero($u, $nodes) {
  if ($u.Kind -eq 'mark') { [void]$nodes.Add($u.Node) }
  elseif ($u.Kind -eq 'hyphen') { [void]$nodes.Add((New-UnitSpecialRun $u.RPr $u.Node)) }
  else { foreach ($n in $u.Nodes) { [void]$nodes.Add($n) } }
}

# Builds, without touching the paragraph, the new child list (after w:pPr) for a paragraph plan from
# Get-ParagraphUnitPlan: unchanged text keeps the formatting of the run it came from, a deleted piece goes in w:del
# (same formatting), inserted text goes in w:ins; every br / hyphen / mark / link / field is put back in its place
# (links and fields are the original nodes, so their XML is exactly the same).
function Get-UnitNodes($plan) {
  $units = $plan.Units
  $nodes = New-Object System.Collections.ArrayList
  $ui = 0
  foreach ($op in $plan.Ops) {
    if ($op.Kind -eq 'ins') {
      while ($ui -lt $units.Count -and $units[$ui].E -eq $units[$ui].S -and $units[$ui].S -le $op.A) { Add-UnitZero $units[$ui] $nodes; $ui++ }
      [void]$nodes.Add((New-Change 'ins' (New-Run (Get-UnitInsertRPr $units $op.A) $op.Text $false)))
      continue
    }
    while ($ui -lt $units.Count) {
      $u = $units[$ui]
      if ($u.S -ge $op.B) { break }
      if ($u.E -eq $u.S) { Add-UnitZero $u $nodes; $ui++; continue }
      $from = [Math]::Max($op.A, $u.S); $to = [Math]::Min($op.B, $u.E)
      if ($to -gt $from) {
        if ($u.Kind -eq 'text') {
          $piece = $op.Text.Substring($from - $op.A, $to - $from)
          if ($op.Kind -eq 'eq') { [void]$nodes.Add((New-Run $u.RPr $piece $false)) }
          else { [void]$nodes.Add((New-Change 'del' (New-Run $u.RPr $piece $true))) }
        } elseif ($u.Kind -eq 'br') {
          [void]$nodes.Add((New-UnitSpecialRun $u.RPr $u.Node))
        } elseif ($from -eq $u.S) {
          foreach ($n in $u.Nodes) { [void]$nodes.Add($n) }
        }
      }
      if ($u.E -le $op.B) { $ui++ } else { break }
    }
  }
  while ($ui -lt $units.Count -and $units[$ui].E -eq $units[$ui].S) { Add-UnitZero $units[$ui] $nodes; $ui++ }
  return , $nodes
}

# A paragraph with a line break, a special hyphen, a hyperlink or a field (doc-lib.ps1 Test-ParagraphHasUnitContent):
# word-level tracked changes that leave every break / hyphen / link / field exactly where it is. Returns $null on
# success or the reason it was skipped (a change touches one of them, or the paragraph holds something that cannot
# be edited). Changes nothing before it has checked the new child list: the text read back (accepting every change)
# must be the new text and no break/hyphen may be lost.
function Revise-ParagraphByUnits($p, [string]$newText) {
  $ns = $script:DocxNs
  $plan = Get-ParagraphUnitPlan $p $newText
  if ($plan.Reason) { return $plan.Reason }
  $nodes = Get-UnitNodes $plan
  $accepted = New-Object System.Text.StringBuilder
  $brAfter = 0; $hyphenAfter = 0
  foreach ($n in $nodes) {
    [void]$accepted.Append((Get-NodeText $n -Raw))
    $brAfter += @($n.SelectNodes('descendant-or-self::w:br[not(ancestor::w:del)]', $ns)).Count
    $hyphenAfter += @($n.SelectNodes('descendant-or-self::w:noBreakHyphen[not(ancestor::w:del)] | descendant-or-self::w:softHyphen[not(ancestor::w:del)]', $ns)).Count
  }
  $brBefore = @($p.SelectNodes('.//w:br[not(ancestor::w:del)]', $ns)).Count
  $hyphenBefore = @($p.SelectNodes('.//w:noBreakHyphen[not(ancestor::w:del)] | .//w:softHyphen[not(ancestor::w:del)]', $ns)).Count
  if (-not [string]::Equals($accepted.ToString(), $plan.NewRaw, [StringComparison]::Ordinal) -or $brAfter -ne $brBefore -or $hyphenAfter -ne $hyphenBefore) {
    return 'ตรวจไม่ผ่านก่อนบันทึก (ข้อความหลังแก้ไม่ตรงกับที่ต้องการ หรือการขึ้นบรรทัดใหม่/ขีดกลางพิเศษไม่ครบ) จึงไม่แก้ย่อหน้านี้ ให้แก้เองจากข้อเสนอในแชต'
  }
  foreach ($c in @($p.ChildNodes)) { if ($c.LocalName -ne 'pPr') { [void]$p.RemoveChild($c) } }
  foreach ($n in $nodes) { [void]$p.AppendChild($n) }
  return $null
}

# Rewrites one paragraph as tracked changes. Returns $null on success or a reason it was skipped.
function Revise-Paragraph($p, [string]$newText, [string]$label) {
  $ns = $script:DocxNs
  if (Test-ParagraphHasUnitContent $p) { return (Revise-ParagraphByUnits $p $newText) }
  # doc-lib.ps1 holds the rules both this script and the table-cell rewrite use: Get-ParagraphEditBlocker (what
  # cannot be edited at all), Get-ParagraphEditPlan (word by word, or the whole paragraph) and
  # Get-ParagraphDiffPieces (which run formatting each piece of text gets).
  $blocker = Get-ParagraphEditBlocker $p
  if ($blocker) { return "$blocker ให้แก้เองจากข้อเสนอในแชต" }
  $plan = Get-ParagraphEditPlan $p $newText

  if ($plan.Mode -eq 'word') {
    # word-level tracked changes. Uniform formatting: every piece takes the formatting of the paragraph.
    # Runs with different formatting (bold in part, different language or complex-script settings, ...): the
    # text of all runs is compared as one string, equal/deleted parts keep their own run's rPr, inserted text
    # takes the rPr of the run before it.
    $bookStarts = @($p.SelectNodes('w:bookmarkStart', $ns))
    $bookEnds = @($p.SelectNodes('w:bookmarkEnd', $ns))
    foreach ($c in @($p.ChildNodes)) { if ($c.LocalName -ne 'pPr') { [void]$p.RemoveChild($c) } }
    foreach ($b in $bookStarts) { [void]$p.AppendChild($b) }
    foreach ($pc in $plan.Pieces) {
      if ($pc.Kind -eq 'eq') { [void]$p.AppendChild((New-Run $pc.RPr $pc.Text $false)) }
      elseif ($pc.Kind -eq 'del') { [void]$p.AppendChild((New-Change 'del' (New-Run $pc.RPr $pc.Text $true))) }
      else { [void]$p.AppendChild((New-Change 'ins' (New-Run $pc.RPr $pc.Text $false))) }
    }
    foreach ($b in $bookEnds) { [void]$p.AppendChild($b) }
    return $null
  }

  # line break or special hyphen: delete old runs as-is (keeps their formatting), insert the new text once
  $script:wholeParaNotes.Add("${label}: แสดงเป็นลบทั้งย่อหน้าแล้วแทรกใหม่ เพราะ$($plan.Why) (Accept/Reject ได้ทั้งย่อหน้าเท่านั้น)") | Out-Null
  foreach ($r in $plan.Runs) {
    foreach ($t in @($r.SelectNodes('w:t', $ns))) {
      $dt = New-W 'delText'
      [void]$dt.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
      $dt.InnerText = $t.InnerText
      [void]$r.ReplaceChild($dt, $t)
    }
    $del = New-W 'del'
    [void]$del.SetAttribute('id', $W_NS, [string]$script:nextId); $script:nextId++
    [void]$del.SetAttribute('author', $W_NS, $Author)
    [void]$del.SetAttribute('date', $W_NS, $script:revDate)
    [void]$p.InsertBefore($del, $r)
    [void]$del.AppendChild($r)
  }
  $ins = New-Change 'ins' (New-Run $plan.FirstRPr $newText.Trim() $false)
  $lastEnd = $p.SelectSingleNode('w:bookmarkEnd[last()]', $ns)
  if ($lastEnd) { [void]$p.InsertBefore($ins, $lastEnd) } else { [void]$p.AppendChild($ins) }
  return $null
}

function Write-Docx([string]$src, [string]$dest) {
  Add-Type -AssemblyName System.IO.Compression
  $in = Open-Shared $src
  try {
    $srcZip = New-Object System.IO.Compression.ZipArchive($in, [System.IO.Compression.ZipArchiveMode]::Read)
    $opened = New-FileExclusive $dest
    $script:writtenDest = $opened.Path
    $outStream = $opened.Stream
    try {
      $outZip = New-Object System.IO.Compression.ZipArchive($outStream, [System.IO.Compression.ZipArchiveMode]::Create)
      foreach ($e in $srcZip.Entries) {
        $ne = $outZip.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
        $ws = $ne.Open()
        try {
          if ($e.FullName -eq 'word/document.xml') {
            $settings = New-Object System.Xml.XmlWriterSettings
            $settings.Encoding = $utf8
            $xw = [System.Xml.XmlWriter]::Create($ws, $settings)
            try { $script:DocxXml.Save($xw) } finally { $xw.Dispose() }
          } else {
            $rs = $e.Open(); try { $rs.CopyTo($ws) } finally { $rs.Dispose() }
          }
        } finally { $ws.Dispose() }
      }
      $outZip.Dispose()
    } finally { $outStream.Dispose() }
  } finally { $in.Dispose() }
}

function Format-AppliedLabel($rev, [string]$label) {
  if ($rev.items) {
    $codes = (@($rev.items) | ForEach-Object { [string]$_ }) -join ', '
    return "$label ($codes)"
  }
  return $label
}

$full = Resolve-Input $Path
$ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()
try {
$revs = Load-Revisions
$applied = New-Object System.Collections.ArrayList
$skipped = New-Object System.Collections.ArrayList
$seen = New-Object 'System.Collections.Generic.HashSet[string]'

switch ($ext) {
  '.docx' {
    $items = Read-Docx $full
    $ids = @($script:DocxXml.SelectNodes('//@w:id', $script:DocxNs) | ForEach-Object { $_.Value -as [int] } | Where-Object { $_ -ne $null })
    $script:nextId = if ($ids.Count) { ($ids | Measure-Object -Maximum).Maximum + 1 } else { 1 }
    foreach ($rev in $revs) {
      $m = Match-Item $items $rev $seen
      if (-not $m) { [void]$skipped.Add($script:reason); continue }
      $why = Revise-Paragraph $m.Node ([string]$rev.text) $m.Label
      if ($why) { [void]$skipped.Add("$($m.Label) $why") } else { [void]$applied.Add((Format-AppliedLabel $rev $m.Label)) }
    }
    if (-not $applied.Count) { $skipped | ForEach-Object { [Console]::Error.WriteLine("ข้าม: $_") }; Fail 'ไม่มีย่อหน้าที่แก้ได้ จึงไม่ได้สร้างไฟล์' }
    $dest = Get-NewPath $full 'humanized' '.docx' $Out $OutDir 'humanized'
    Write-Docx $full $dest
    $dest = $script:writtenDest
  }
  { $_ -in '.txt', '.md' } {
    $items = Read-Txt $full
    $edits = New-Object System.Collections.ArrayList
    foreach ($rev in $revs) {
      $m = Match-Item $items $rev $seen
      if (-not $m) { [void]$skipped.Add($script:reason); continue }
      [void]$edits.Add(@{ Item = $m.Item; Text = ([string]$rev.text).Trim() }); [void]$applied.Add((Format-AppliedLabel $rev $m.Label))
    }
    if (-not $applied.Count) { $skipped | ForEach-Object { [Console]::Error.WriteLine("ข้าม: $_") }; Fail 'ไม่มีย่อหน้าที่แก้ได้ จึงไม่ได้สร้างไฟล์' }
    $text = $script:TxtText
    foreach ($e in ($edits | Sort-Object { $_.Item.Start } -Descending)) {
      $text = $text.Substring(0, $e.Item.Start) + $e.Text + $text.Substring($e.Item.Start + $e.Item.Length)
    }
    if ($script:TxtCrlf) { $text = $text -replace "`n", "`r`n" }
    $enc = switch ($script:TxtEncoding) { 'utf16' { New-Object System.Text.UnicodeEncoding($false, $true) } 'utf8' { $utf8 } default { New-Object System.Text.UTF8Encoding($true) } }
    $dest = Get-NewPath $full 'humanized' $ext $Out $OutDir 'humanized'
    $opened = New-FileExclusive $dest
    $dest = $opened.Path
    try { $sw = New-Object System.IO.StreamWriter($opened.Stream, $enc); $sw.Write($text); $sw.Dispose() } finally { $opened.Stream.Dispose() }
    if ($script:TxtEncoding -eq 'cp874') { Write-Output 'หมายเหตุ: ต้นฉบับเป็น Windows-874 ไฟล์ใหม่บันทึกเป็น UTF-8' }
  }
  '.pdf' {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# ฉบับปรับจาก $(Split-Path $full -Leaf)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('ต้นฉบับเป็น PDF จึงสร้างเป็นไฟล์ข้อความคู่ต้นฉบับ/ฉบับปรับ ให้นำไปแก้ในไฟล์ต้นทาง (เช่น Word) เอง')
    $i = 0
    foreach ($rev in $revs) {
      $i++
      $page = $rev.page -as [int]
      if (-not $page -or -not ([string]$rev.original).Trim() -or -not ([string]$rev.text).Trim()) { [void]$skipped.Add("รายการที่ $i ต้องมี page, original และ text"); continue }
      [void]$sb.AppendLine(''); [void]$sb.AppendLine("## หน้า $page (รายการที่ $i)"); [void]$sb.AppendLine('')
      [void]$sb.AppendLine('**ต้นฉบับ**'); [void]$sb.AppendLine(''); [void]$sb.AppendLine('> ' + (([string]$rev.original).Trim() -replace "`n", "`n> ")); [void]$sb.AppendLine('')
      [void]$sb.AppendLine('**ฉบับปรับ**'); [void]$sb.AppendLine(''); [void]$sb.AppendLine('> ' + (([string]$rev.text).Trim() -replace "`n", "`n> "))
      [void]$applied.Add("หน้า $page")
    }
    if (-not $applied.Count) { $skipped | ForEach-Object { [Console]::Error.WriteLine("ข้าม: $_") }; Fail 'ไม่มีรายการที่ใช้ได้ จึงไม่ได้สร้างไฟล์' }
    $dest = Get-NewPath $full 'humanized' '.md' $Out $OutDir 'review'
    $opened = New-FileExclusive $dest
    $dest = $opened.Path
    try { $sw = New-Object System.IO.StreamWriter($opened.Stream, $utf8); $sw.Write($sb.ToString()); $sw.Dispose() } finally { $opened.Stream.Dispose() }
  }
  default { Fail "ยังไม่รองรับไฟล์ $ext รองรับ .docx, .pdf, .txt" }
}

Write-Output "สร้างไฟล์ใหม่: $dest"
Write-Output ("แก้แล้ว: " + ($applied -join ', '))
foreach ($n in $script:wholeParaNotes) { Write-Output $n }
foreach ($s in $skipped) { Write-Output "ข้าม: $s" }
if ($ext -eq '.docx') { Write-Output "การแก้เป็น tracked changes ชื่อผู้แก้ '$Author' ตรวจและกด Accept/Reject ใน Word ได้" }
} finally { Remove-JsonInput $Revisions }
