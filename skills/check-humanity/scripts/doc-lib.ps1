# Shared reading and paragraph numbering for extract-text.ps1 and apply-revisions.ps1.
# Both scripts must number paragraphs identically, so the logic lives only here.

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$script:skippedTextBoxes = 0
$script:skippedToc = 0
$script:tocTitleRemoved = $false
$script:encodingNote = $null

function Fail([string]$message) {
  [Console]::Error.WriteLine($message)
  exit 2
}

# ---------------------------------------------------------------------------------------------
# Shared review/report/brief/rewrite .md file hashing and reading (review-file.ps1, write-rewrite.ps1
# -Review; docs/v2-plan.md 3.6/3.8 F1). Kept here so write-rewrite.ps1 -Review reads a preview file
# with the SAME parser review-file.ps1 -Action Read uses, instead of a second, drifting copy.
# ---------------------------------------------------------------------------------------------

# Body normalization used both to compute the file's own content hash (Get-BodyHash) and to decide
# whether a file was "edited outside the checkboxes": strips trailing whitespace per line and blanks
# out any checkbox mark (so re-ticking a box does not itself change the hash), then trims trailing
# blank lines (a stray blank line before the trailer must not look like an edit).
function Normalize-Body([string]$body) {
  $b = $body -replace "`r`n", "`n"
  $lines = $b -split "`n" | ForEach-Object {
    $l = $_ -replace '[ \t]+$', ''
    $l -replace '\[[ xX✓✔]\]', '[ ]'
  }
  return (($lines -join "`n").TrimEnd("`n"))
}

function Get-BodyHash([string]$body) {
  $norm = Normalize-Body $body
  $bytes = $utf8.GetBytes($norm)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant() }
  finally { $sha.Dispose() }
}

# Reads and parses a review/report/brief/rewrite .md file created by review-file.ps1 -Action Create
# (see that script's header comment for the trailer format). Shared by review-file.ps1's own
# -Action Read and by write-rewrite.ps1 -Review (F1), so both read the same trailer/checkbox/hash/
# source-name logic instead of two independently drifting parsers. Does not apply kind-specific
# rules (e.g. "report/brief has no checkboxes to read") - callers still decide that themselves from
# the returned .Kind. Fails (exit 2) on a missing file or a malformed/foreign trailer.
#
# Returns @{ Kind; AllIds; Ticked; Unknown; ShaAttr; LastWrite; SourceName; Edited; BodyLines; CreatedAt }:
#   Kind        'review' | 'report' | 'brief' | 'rewrite'
#   AllIds      every R<n>/S<n> id declared in the trailer's items="..." attribute, in order
#   Ticked      ids with a ticked checkbox in the body, in first-appearance order, de-duplicated
#   Unknown     ticked ids that are not in AllIds, de-duplicated
#   ShaAttr     the trailer's own recorded sha256 (lowercase hex, 64 chars)
#   LastWrite   the file's last-write time, 'yyyy-MM-dd HH:mm'
#   SourceName  the source file name recorded in the body's "ไฟล์ต้นฉบับ: <name> (...)" line, or
#               $null if that line is missing/unparseable
#   Edited      $true when the file's current body hash no longer matches the trailer's sha256
#               (edited somewhere other than a checkbox mark)
#   BodyLines   the body lines (everything before the trailer), for a caller that needs more detail
#   CreatedAt   the "- สร้างเมื่อ: yyyy-MM-dd HH:mm" value recorded in the body, or $null if that
#               line is missing/unparseable
function Read-ReviewFile([string]$path) {
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Fail "ไม่พบไฟล์: $path" }
  $full = (Resolve-Path -LiteralPath $path).ProviderPath
  $stream = Open-Shared $full
  try {
    $ms = New-Object System.IO.MemoryStream
    $stream.CopyTo($ms)
    $bytes = $ms.ToArray()
  } finally { $stream.Dispose() }
  $text = $utf8.GetString($bytes)
  if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
  $lastWrite = (Get-Item -LiteralPath $full).LastWriteTime.ToString('yyyy-MM-dd HH:mm')

  $normText = $text -replace "`r`n", "`n"
  $lines = $normText -split "`n"
  $trailerIdx = -1
  for ($ti = $lines.Count - 1; $ti -ge 0; $ti--) {
    if ($lines[$ti] -match '^<!--\s*hz-review\s') { $trailerIdx = $ti; break }
  }
  if ($trailerIdx -lt 0) { Fail 'ไม่ใช่ไฟล์รายการจาก skill นี้ หรือบรรทัดท้ายไฟล์ถูกลบ' }
  $trailerLine = $lines[$trailerIdx]
  if ($trailerLine -notmatch '^<!--\s*hz-review\s+v1\s+kind="(review|report|brief|rewrite)"\s+items="([^"]*)"\s+sha256="([0-9a-f]{64})"\s*-->\s*$') {
    Fail 'ไม่ใช่ไฟล์รายการจาก skill นี้ หรือบรรทัดท้ายไฟล์ถูกลบ'
  }
  $kind = $Matches[1]; $itemsAttr = $Matches[2]; $shaAttr = $Matches[3]
  $allIds = @($itemsAttr -split ',' | Where-Object { $_ })
  $bodyLines = if ($trailerIdx -gt 0) { $lines[0..($trailerIdx - 1)] } else { @() }
  $body = ($bodyLines -join "`n")
  $curHash = Get-BodyHash $body

  $sourceName = $null
  foreach ($l in $bodyLines) {
    if ($l -match '^-\s*ไฟล์ต้นฉบับ:\s*(.+?)\s*\(ไม่ได้แก้ไฟล์นี้\)\s*$') { $sourceName = $Matches[1]; break }
  }
  # G2 (docs/v2-plan.md 3.9): the "- สร้างเมื่อ: yyyy-MM-dd HH:mm" line every Create writes (review/
  # report/brief/rewrite alike) - used by write-rewrite.ps1 -Review to pick the most RECENTLY
  # CREATED preview containing a given section when several preview files are supplied.
  $createdAt = $null
  foreach ($l in $bodyLines) {
    if ($l -match '^-\s*สร้างเมื่อ:\s*(.+?)\s*$') { $createdAt = $Matches[1]; break }
  }

  $tickPattern = if ($kind -eq 'rewrite') { '^\s*[-*+]\s*\[([ xX✓✔])\]\s*ใช้\s*(S\d+)\b' } else { '^\s*[-*+]\s*\[([ xX✓✔])\]\s*ใช้\s*(R\d+)\b' }
  $ticked = New-Object System.Collections.ArrayList
  $unknown = New-Object System.Collections.ArrayList
  foreach ($l in $bodyLines) {
    if ($l -match $tickPattern) {
      $mark = $Matches[1]; $rid = $Matches[2]
      if ($mark -ne ' ') {
        if (-not $ticked.Contains($rid)) { [void]$ticked.Add($rid) }
        if ($allIds -notcontains $rid) { [void]$unknown.Add($rid) }
      }
    }
  }
  $edited = -not [string]::Equals($curHash, $shaAttr, [StringComparison]::OrdinalIgnoreCase)

  return @{
    Kind = $kind; AllIds = $allIds; Ticked = @($ticked); Unknown = @($unknown | Select-Object -Unique)
    ShaAttr = $shaAttr; LastWrite = $lastWrite; SourceName = $sourceName; Edited = $edited
    BodyLines = $bodyLines; CreatedAt = $createdAt
  }
}

function Parse-Ranges([string]$spec) {
  $set = New-Object 'System.Collections.Generic.HashSet[int]'
  foreach ($part in ($spec -split ',')) {
    $p = $part.Trim()
    if ($p -match '^(\d{1,9})\s*-\s*(\d{1,9})$') {
      $a = [int]$Matches[1]; $b = [int]$Matches[2]
      if ($a -gt $b) { Fail "ช่วงไม่ถูกต้อง: $p (เลขแรกต้องไม่มากกว่าเลขหลัง)" }
      if ($b - $a -gt 100000) { Fail "ช่วงกว้างเกินไป: $p" }
      for ($i = $a; $i -le $b; $i++) { [void]$set.Add($i) }
    } elseif ($p -match '^\d{1,9}$') {
      [void]$set.Add([int]$p)
    } elseif ($p) {
      Fail "ช่วงไม่ถูกต้อง: $p (ใช้รูปแบบ 5-9,12)"
    }
  }
  return , $set
}

function Open-Shared([string]$file) {
  try {
    return [System.IO.File]::Open($file, 'Open', 'Read', 'ReadWrite, Delete')
  } catch [System.IO.IOException] {
    Fail 'เปิดไฟล์ไม่ได้เพราะไฟล์ถูกใช้งานอยู่ เช่น เปิดค้างใน Word หรือโปรแกรมซิงก์ไฟล์กำลังใช้อยู่ ให้บันทึกและปิดไฟล์ใน Word แล้วลองใหม่'
  } catch [System.UnauthorizedAccessException] {
    Fail 'ไม่มีสิทธิ์อ่านไฟล์นี้'
  }
}

function Resolve-Input([string]$path) {
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Fail "ไม่พบไฟล์: $path" }
  return (Resolve-Path -LiteralPath $path).ProviderPath
}

# True if $path (already a resolved full path) is inside $dir (also resolved to a full path),
# comparing case-insensitively with a trailing separator so e.g. "C:\TempX" never matches "C:\Temp".
function Test-UnderTempDir([string]$path, [string]$dir) {
  if (-not $dir) { return $false }
  try { $d = ([System.IO.Path]::GetFullPath($dir)).TrimEnd('\', '/') + '\' } catch { return $false }
  $p = $path.TrimEnd('\', '/') + '\'
  return $p.StartsWith($d, [StringComparison]::OrdinalIgnoreCase)
}

# Best-effort cleanup of a temp JSON input file (review-file.ps1 -Items / apply-revisions.ps1
# -Revisions). Called from a `finally` block so it runs on both success and error (including a
# JSON parse error or a validation failure discovered after reading). Only ever deletes a path
# that is an existing FILE with a .json extension AND that resolves to somewhere under the
# system temp directory ([System.IO.Path]::GetTempPath() / $env:TEMP / $env:TMP) - never a file
# elsewhere, even one named "...json" - so a caller error can never delete a real project or user
# file. A missing path, a non-temp location, a non-.json path, or a deletion failure (e.g. the
# file is locked) is silently left alone - this is cleanup, not part of the script's pass/fail
# contract.
function Remove-JsonInput([string]$path) {
  if (-not $path) { return }
  try {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
    $ext = [System.IO.Path]::GetExtension($path)
    if (-not [string]::Equals($ext, '.json', [StringComparison]::OrdinalIgnoreCase)) { return }
    $full = (Resolve-Path -LiteralPath $path).ProviderPath
    $inTemp = (Test-UnderTempDir $full ([System.IO.Path]::GetTempPath())) -or
      (Test-UnderTempDir $full $env:TEMP) -or (Test-UnderTempDir $full $env:TMP)
    if (-not $inTemp) { return }
    Remove-Item -LiteralPath $full -Force -ErrorAction Stop
  } catch { }
}

$script:skipTextXPath = 'ancestor::w:txbxContent or ancestor::w:del or ancestor::w:moveFrom or ancestor::mc:Fallback'

# Visible text of a node after accepting tracked changes. -Raw keeps leading/trailing whitespace.
function Get-NodeText($node, [switch]$Raw) {
  $ns = $script:DocxNs
  $sb = New-Object System.Text.StringBuilder
  $x = ".//w:t[not($($script:skipTextXPath))] | .//w:tab[not($($script:skipTextXPath))] | .//w:br[not($($script:skipTextXPath))]"
  foreach ($t in $node.SelectNodes($x, $ns)) {
    if ($t.LocalName -eq 't') { [void]$sb.Append($t.InnerText) } elseif ($t.LocalName -eq 'tab') { [void]$sb.Append("`t") } else { [void]$sb.Append(' ') }
  }
  if ($Raw) { return $sb.ToString() }
  return $sb.ToString().Trim()
}

# Text of a paragraph as its runs/links/fields carry it: Get-NodeText of every child except w:pPr, joined. Get-NodeText $p
# itself also matches the tab-STOP definitions in w:pPr/w:tabs/w:tab and would start the text with a tab per stop,
# so every edit that compares or rewrites the paragraph text must use this instead.
function Get-ParagraphBodyText($p) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($c in $p.ChildNodes) {
    if ($c.NodeType -ne "Element" -or $c.LocalName -eq "pPr") { continue }
    [void]$sb.Append((Get-NodeText $c -Raw))
  }
  return $sb.ToString()
}

# Items: @{ N; Level; Kind; Caption; Text; Node }
# Also sets $script:DocxXml and $script:DocxNs for callers that modify the document.
function Read-Docx([string]$file) {
  Add-Type -AssemblyName System.IO.Compression
  $stream = Open-Shared $file
  $rawParts = @{}
  try {
    try {
      $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read)
      foreach ($name in 'word/document.xml', 'word/styles.xml') {
        $entry = $zip.GetEntry($name)
        if (-not $entry) { continue }
        $reader = New-Object System.IO.StreamReader($entry.Open(), $utf8)
        try { $rawParts[$name] = $reader.ReadToEnd() } finally { $reader.Dispose() }
      }
    } catch {
      Fail 'ไฟล์นี้ไม่ใช่ .docx ที่อ่านได้ หรือไฟล์เสียหาย ให้เปิดใน Word แล้วบันทึกเป็น .docx ใหม่'
    }
  } finally { $stream.Dispose() }

  $docXmlText = $rawParts['word/document.xml']
  if (-not $docXmlText) { Fail 'ไม่พบเนื้อหาเอกสาร Word (word/document.xml) ไฟล์อาจไม่ใช่ .docx หรือเสียหาย' }
  $doc = New-Object System.Xml.XmlDocument
  $doc.PreserveWhitespace = $true
  try { $doc.LoadXml($docXmlText) } catch { Fail 'อ่านเนื้อหาเอกสาร Word ไม่ได้ ไฟล์อาจเสียหาย' }

  $styles = $null
  if ($rawParts.ContainsKey('word/styles.xml')) {
    try {
      $styles = New-Object System.Xml.XmlDocument
      $styles.PreserveWhitespace = $true
      $styles.LoadXml($rawParts['word/styles.xml'])
    } catch { $styles = $null }
  }

  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('w', $W_NS)
  $ns.AddNamespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
  $script:DocxXml = $doc
  $script:DocxNs = $ns

  $headingStyles = @{}
  $tocStyles = @{}
  $tocHeadingStyles = @{}
  $captionStyles = @{}
  if ($styles) {
    $sns = New-Object System.Xml.XmlNamespaceManager($styles.NameTable)
    $sns.AddNamespace('w', $W_NS)
    foreach ($s in $styles.SelectNodes("//w:style[@w:type='paragraph']", $sns)) {
      $id = $s.GetAttribute('styleId', $W_NS)
      $nameNode = $s.SelectSingleNode('w:name', $sns)
      $name = if ($nameNode) { $nameNode.GetAttribute('val', $W_NS) } else { '' }
      $lvlNode = $s.SelectSingleNode('w:pPr/w:outlineLvl', $sns)
      if ($name -match '^(?i)caption$') { $captionStyles[$id] = $true }
      if ($name -match '^(?i)toc\s*heading$') { $tocHeadingStyles[$id] = $true; continue }
      if ($name -match '^(?i)toc\s*\d$') { $tocStyles[$id] = $true; continue }
      if ($name -match '^(?i)heading\s*(\d)$') { $headingStyles[$id] = [int]$Matches[1] }
      elseif ($name -match '^(?i)title$') { $headingStyles[$id] = 1 }
      elseif ($lvlNode) {
        $v = [int]$lvlNode.GetAttribute('val', $W_NS)
        if ($v -lt 9) { $headingStyles[$id] = $v + 1 }
      }
    }
  }

  # TOC bookkeeping belongs to one call: without this reset a second Read-Docx in the same process skips the
  # fallback rule that drops the TOC title, so its numbered paragraphs shift by one and skippedToc keeps adding up.
  $script:skippedToc = 0
  $script:tocTitleRemoved = $false
  $script:skippedTextBoxes = $doc.SelectNodes('//w:txbxContent[not(ancestor::mc:Fallback)]', $ns).Count
  $items = New-Object System.Collections.ArrayList
  $n = 0
  $blocks = $doc.SelectNodes('//w:body//w:p[not(ancestor::w:tbl) and not(ancestor::w:txbxContent)] | //w:body//w:tbl[not(ancestor::w:tbl)]', $ns)
  foreach ($b in $blocks) {
    if ($b.LocalName -eq 'tbl') {
      $rows = foreach ($tr in $b.SelectNodes('w:tr', $ns)) {
        (($tr.SelectNodes('w:tc', $ns) | ForEach-Object { Get-NodeText $_ }) -join ' | ')
      }
      $text = (($rows | Where-Object { $_ -replace '[\s|]', '' }) -join "`n")
      if (-not $text) { continue }
      $n++; [void]$items.Add(@{ N = $n; Level = 0; Kind = 'table'; Caption = $false; Text = $text; Node = $b })
      continue
    }
    $text = Get-NodeText $b
    if (-not $text) { continue }
    $level = 0
    $sid = $null
    $style = $b.SelectSingleNode('w:pPr/w:pStyle', $ns)
    if ($style) { $sid = $style.GetAttribute('val', $W_NS) }
    if ($sid -and $tocHeadingStyles.ContainsKey($sid)) {
      # this paragraph IS the TOC title itself (e.g. "TOC Heading" style => สารบัญ)
      $script:skippedToc++
      $script:tocTitleRemoved = $true
      continue
    }
    if ($sid -and $tocStyles.ContainsKey($sid)) {
      $script:skippedToc++
      # fallback heuristic only when no explicit TOC Heading style was seen: a heading right
      # before the TOC entries whose text itself reads as a TOC title (e.g. สารบัญ / Contents)
      if ($items.Count -and $items[$items.Count - 1].Level -gt 0 -and -not $script:tocTitleRemoved) {
        $prevText = $items[$items.Count - 1].Text -replace '\s+', ' '
        if ($prevText -match '^(?i)(สารบัญ|สารบาญ|table of contents|contents)(\s|$)') {
          $items.RemoveAt($items.Count - 1); $n--; $script:skippedToc++; $script:tocTitleRemoved = $true
        }
      }
      continue
    }
    if ($sid -and $headingStyles.ContainsKey($sid)) { $level = $headingStyles[$sid] }
    $lvl = $b.SelectSingleNode('w:pPr/w:outlineLvl', $ns)
    if (-not $level -and $lvl) {
      $v = [int]$lvl.GetAttribute('val', $W_NS)
      if ($v -lt 9) { $level = $v + 1 }
    }
    $caption = [bool]($level -eq 0 -and $sid -and $captionStyles.ContainsKey($sid))
    $n++; [void]$items.Add(@{ N = $n; Level = $level; Kind = 'p'; Caption = $caption; Text = $text; Node = $b })
  }
  return , $items
}

function Test-Body($it) {
  return ($it.Kind -eq 'p' -and $it.Level -eq 0 -and -not $it.Caption)
}

# An "anchor" item (table, or a paragraph carrying an image, an embedded (OLE) object, hyperlink,
# field, footnote/endnote reference, comment reference, an existing tracked change (w:ins/w:del),
# or a content control (w:sdt) anywhere in its descendants) - shared by write-rewrite.ps1
# (docs/v2-plan.md 3.6 3A-4/3.9 G1) and review-file.ps1 -Kind rewrite (3.9 G5), so both scripts agree
# on exactly the same definition of "anchor" instead of drifting copies. w:object (an embedded OLE
# object, e.g. an Excel sheet or a Visio drawing) is an anchor too: a paragraph that has text AND an
# object must be kept, otherwise a rewrite would drop the object without a word.
$script:AnchorXPath = './/w:drawing | .//w:pict | .//w:object | .//w:hyperlink | .//w:fldSimple | .//w:fldChar' +
' | .//w:footnoteReference | .//w:endnoteReference | .//w:commentReference' +
' | .//w:commentRangeStart | .//w:commentRangeEnd | .//w:ins | .//w:del | .//w:sdt'
function Test-AnchorItem($item, $ns) {
  if ($item.Kind -eq 'table') { return $true }
  $node = $item.Node
  if ($node.SelectSingleNode('ancestor::w:sdt', $ns)) { return $true }
  return [bool]$node.SelectSingleNode($script:AnchorXPath, $ns)
}

# "Gap" nodes (docx only): body-level nodes that sit BETWEEN two consecutive items but are not items
# themselves, because Read-Docx skips paragraphs/tables with no text (an image-only paragraph, an
# empty paragraph, a text-less table, a section-break paragraph, a skipped TOC line, a bare
# bookmark/comment marker ...). Read-Docx's own numbering is deliberately NOT changed by this - every
# other script uses that ¶ numbering. Consecutive gap nodes form one "group" (one group at most per
# pair of neighbouring items). A group whose content is anchor-like (image, picture, OLE object,
# field, hyperlink, comment/footnote reference, tracked change, content control, table, section
# break - the same list as Test-AnchorItem plus w:tbl/w:sectPr) is an "implicit anchor" (it
# has no ¶ and cannot be kept by number); any other group (only empty paragraphs and the like) is a
# "spacer". Shared by write-rewrite.ps1 and extract-text.ps1.
$script:GapAnchorNames = @('drawing', 'pict', 'object', 'hyperlink', 'fldSimple', 'fldChar', 'footnoteReference',
  'endnoteReference', 'commentReference', 'commentRangeStart', 'commentRangeEnd', 'ins', 'del', 'sdt', 'tbl', 'sectPr')
$script:GapAnchorXPath = $script:AnchorXPath + ' | .//w:tbl | .//w:sectPr'
$script:GapImageNames = @('drawing', 'pict', 'object')
$script:GapImageXPath = './/w:drawing | .//w:pict | .//w:object'
function Test-GapNodeAnchor($node, $ns) {
  if ($script:GapAnchorNames -contains $node.LocalName) { return $true }
  return [bool]$node.SelectSingleNode($script:GapAnchorXPath, $ns)
}
function Test-GapNodeImage($node, $ns) {
  if ($script:GapImageNames -contains $node.LocalName) { return $true }
  return [bool]$node.SelectSingleNode($script:GapImageXPath, $ns)
}
# Returns an array of @{ PrevN; NextN; Nodes = <element nodes in order>; Anchor = <bool>; HasImage = <bool> },
# one per pair of neighbouring items (PrevN, NextN) that have at least one gap node between them.
# Pairs whose nodes are not siblings of each other (e.g. one inside a content control) are skipped.
function Get-DocxGapGroups($items, $ns) {
  $groups = New-Object System.Collections.ArrayList
  for ($i = 0; $i -lt $items.Count - 1; $i++) {
    $a = $items[$i]; $b = $items[$i + 1]
    if ($null -eq $a.Node -or $null -eq $b.Node) { continue }
    if (-not [object]::ReferenceEquals($a.Node.ParentNode, $b.Node.ParentNode)) { continue }
    $nodes = New-Object System.Collections.ArrayList
    $found = $false
    $cur = $a.Node.NextSibling
    while ($null -ne $cur) {
      if ([object]::ReferenceEquals($cur, $b.Node)) { $found = $true; break }
      if ($cur.NodeType -eq [System.Xml.XmlNodeType]::Element) { [void]$nodes.Add($cur) }
      $cur = $cur.NextSibling
    }
    if (-not $found -or -not $nodes.Count) { continue }
    $anchor = $false; $image = $false
    foreach ($nd in $nodes) {
      if (Test-GapNodeAnchor $nd $ns) { $anchor = $true }
      if (Test-GapNodeImage $nd $ns) { $image = $true }
    }
    [void]$groups.Add(@{ PrevN = $a.N; NextN = $b.N; Nodes = @($nodes); Anchor = $anchor; HasImage = $image })
  }
  return , $groups.ToArray()
}

# Items: @{ N; Level; Kind = 'p'; Text; Start; Length } with offsets into $script:TxtText (LF line endings).
function Read-Txt([string]$file) {
  $s = Open-Shared $file
  try { $ms = New-Object System.IO.MemoryStream; $s.CopyTo($ms); $bytes = $ms.ToArray() } finally { $s.Dispose() }
  $text = $null
  $script:TxtEncoding = 'utf8'
  if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2); $script:TxtEncoding = 'utf16' }
  elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $text = [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2); $script:TxtEncoding = 'utf16' }
  else {
    try {
      $strict = New-Object System.Text.UTF8Encoding($false, $true)
      $text = $strict.GetString($bytes)
      if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1); $script:TxtEncoding = 'utf8bom' }
    } catch {
      $text = [System.Text.Encoding]::GetEncoding(874).GetString($bytes)
      $script:TxtEncoding = 'cp874'
      $script:encodingNote = 'อ่านด้วยรหัส Windows-874 (ไทย) เพราะไม่ใช่ UTF-8 โปรดตรวจว่าตัวอักษรถูกต้อง'
    }
  }
  $script:TxtCrlf = $text.Contains("`r`n")
  $text = $text -replace "`r`n?", "`n"
  $script:TxtText = $text
  $sep = if ($text -match "\n[ \t]*\n") { "\n[ \t]*\n" } else { "\n" }
  $bounds = New-Object System.Collections.ArrayList
  $pos = 0
  foreach ($m in [regex]::Matches($text, $sep)) { [void]$bounds.Add(@($pos, $m.Index)); $pos = $m.Index + $m.Length }
  [void]$bounds.Add(@($pos, $text.Length))
  $items = New-Object System.Collections.ArrayList
  $n = 0
  foreach ($bd in $bounds) {
    $seg = $text.Substring($bd[0], $bd[1] - $bd[0])
    $t = $seg.Trim()
    if (-not $t) { continue }
    $lead = $seg.Length - $seg.TrimStart().Length
    $level = 0
    if ($t -notmatch "\n" -and $t -match '^(#{1,6})\s+\S') { $level = $Matches[1].Length }
    $n++; [void]$items.Add(@{ N = $n; Level = $level; Kind = 'p'; Caption = $false; Text = $t; Start = $bd[0] + $lead; Length = $t.Length })
  }
  return , $items
}

function Get-SectionEnd($items, $head) {
  $end = $items[$items.Count - 1].N
  foreach ($it in $items) {
    if ($it.N -gt $head.N -and $it.Level -gt 0 -and $it.Level -le $head.Level) { $end = $it.N - 1; break }
  }
  return $end
}

# Table-cell enumeration shared by extract-text.ps1 (-Cells) and apply-revisions.ps1 (row/cell edits).
# Returns a list of @{ Row; Cell; Para; ParaCount; Text; Node; Span; VCont; Nested }, one record
# per grid cell (gridSpan counted once) unless the cell has >1 non-empty direct paragraph, in which
# case one record per non-empty paragraph is returned (Para = 1-based index among the non-empty
# ones). vMerge-continuation cells and cells containing a nested w:tbl still get exactly one record
# (Node = $null, Text = joined text of the cell's direct paragraphs, possibly empty); a cell with no
# text at all (and not vMerge/nested) also gets exactly one record with Node = $null, ParaCount = 0.
function Get-TableCells($tbl) {
  $ns = $script:DocxNs
  $result = New-Object System.Collections.ArrayList
  $rowIdx = 0
  foreach ($tr in $tbl.SelectNodes('w:tr', $ns)) {
    $rowIdx++
    $cellIdx = 0
    foreach ($tc in $tr.SelectNodes('w:tc', $ns)) {
      $cellIdx++
      $tcPr = $tc.SelectSingleNode('w:tcPr', $ns)
      $span = 1
      if ($tcPr) {
        $gridSpan = $tcPr.SelectSingleNode('w:gridSpan', $ns)
        if ($gridSpan) { $v = $gridSpan.GetAttribute('val', $W_NS); if ($v) { $span = [int]$v } }
      }
      $isVCont = $false
      if ($tcPr) {
        $vMergeNode = $tcPr.SelectSingleNode('w:vMerge', $ns)
        if ($vMergeNode) {
          $val = $vMergeNode.GetAttribute('val', $W_NS)
          if (-not $val -or $val -eq 'continue') { $isVCont = $true }
        }
      }
      $hasNestedTbl = [bool]$tc.SelectSingleNode('w:tbl', $ns)
      # Direct cell paragraphs, including ones wrapped in a content control (w:sdt/w:sdtContent/w:p).
      $directParas = @($tc.SelectNodes('w:p | w:sdt/w:sdtContent/w:p', $ns))
      $textParas = New-Object System.Collections.ArrayList
      foreach ($p in $directParas) {
        $t = Get-NodeText $p
        if ($t) { [void]$textParas.Add(@{ Node = $p; Text = $t }) }
      }
      if ($isVCont -or $hasNestedTbl) {
        # A nested table's own text is included so the reported Text reflects everything the cell
        # shows, even though the cell still cannot be edited in the file (Nested = $true below).
        $text = if ($hasNestedTbl) { Get-NodeText $tc } else { ($textParas | ForEach-Object { $_.Text }) -join ' ' }
        [void]$result.Add(@{ Row = $rowIdx; Cell = $cellIdx; Para = $null; ParaCount = $textParas.Count; Text = $text; Node = $null; Span = $span; VCont = $isVCont; Nested = $hasNestedTbl })
        continue
      }
      if (-not $textParas.Count) {
        [void]$result.Add(@{ Row = $rowIdx; Cell = $cellIdx; Para = $null; ParaCount = 0; Text = ''; Node = $null; Span = $span; VCont = $false; Nested = $false })
        continue
      }
      if ($textParas.Count -eq 1) {
        [void]$result.Add(@{ Row = $rowIdx; Cell = $cellIdx; Para = $null; ParaCount = 1; Text = $textParas[0].Text; Node = $textParas[0].Node; Span = $span; VCont = $false; Nested = $false })
      } else {
        $pk = 0
        foreach ($tp in $textParas) {
          $pk++
          [void]$result.Add(@{ Row = $rowIdx; Cell = $cellIdx; Para = $pk; ParaCount = $textParas.Count; Text = $tp.Text; Node = $tp.Node; Span = $span; VCont = $false; Nested = $false })
        }
      }
    }
  }
  return , $result
}

# The skill's own folder (the one that holds SKILL.md): the parent of scripts\. This works wherever the skill is
# installed (<home>\.agents\skills\check-humanity, <home>\.claude\skills\check-humanity, or the dev repo's
# .agents\skills\check-humanity) because it does not assume anything about the folders above it.
function Get-SkillFolder {
  return (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
}

# The root folder for every output file (review / report / brief / rewrite preview / humanized / rewritten).
# Default: <MyDocuments>\HumanizerAgent\output (MyDocuments follows a OneDrive-redirected Documents folder).
# Env HUMANIZER_OUTPUT_DIR, when set and not empty, replaces it (used by the tests so they never write into
# the real Documents folder). Fails with a Thai message pointing to -OutDir when MyDocuments cannot be found.
function Get-OutputRoot {
  $envDir = [Environment]::GetEnvironmentVariable('HUMANIZER_OUTPUT_DIR')
  if ($envDir -and $envDir.Trim()) { return $envDir.Trim().TrimEnd('\', '/') }
  $docs = $null
  try { $docs = [Environment]::GetFolderPath('MyDocuments') } catch { $docs = $null }
  if (-not $docs) { Fail 'หาโฟลเดอร์ Documents ของเครื่องนี้ไม่พบ ให้ระบุโฟลเดอร์ปลายทางเองด้วย -OutDir' }
  return (Join-Path $docs 'HumanizerAgent\output')
}

function Test-UnderDir([string]$path, [string]$dir) {
  $p = $path.TrimEnd('\', '/') + '\'
  $d = $dir.TrimEnd('\', '/') + '\'
  return $p.StartsWith($d, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($path.TrimEnd('\', '/'), $dir.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)
}

# O2 guard: nothing is ever written inside the skill's own folder (the one that holds SKILL.md).
function Test-OutputGuard([string]$dir, [string]$skillFolder) {
  if (Test-UnderDir $dir $skillFolder) {
    Fail 'ปลายทางอยู่ในโฟลเดอร์ของ skill เอง ห้ามเขียนไฟล์ผลลัพธ์ที่นี่ ให้ใช้ค่าเริ่มต้น (โฟลเดอร์ HumanizerAgent\output ใน Documents) หรือเลือกโฟลเดอร์อื่น'
  }
}

# Resolves the output path for review/report/humanized files. kindFolder is 'review' or 'humanized'.
# Default (no -out/-outDir): <Get-OutputRoot>\<kindFolder>\<base>\<base>_<suffix>_<yyyy-MM-dd><ext>
# -outDir: file goes directly inside that folder (still de-duplicated with _2, _3, ...). A missing folder is created if its parent exists.
# -out: exact path, must not already exist (no de-duplication).
function Get-NewPath([string]$src, [string]$suffix, [string]$ext, [string]$out, [string]$outDir, [string]$kindFolder) {
  if ($out -and $outDir) { Fail 'ใช้ -Out หรือ -OutDir อย่างใดอย่างหนึ่ง' }
  $lenFail = 'พาธปลายทางยาวเกินไป ให้ย่อชื่อไฟล์ต้นฉบับหรือใช้ -OutDir ที่สั้นกว่า'
  $root = Get-SkillFolder
  $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
  $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
  $safeBase = ($base -replace ("[" + [regex]::Escape($invalid) + "]"), '_')
  if ($safeBase.Length -gt 50) { $safeBase = $safeBase.Substring(0, 50).TrimEnd(' ', '.') }
  $today = Get-Date -Format 'yyyy-MM-dd'

  if ($out) {
    # Length checks come before any call that touches the filesystem with the raw (possibly
    # over-limit) path: GetFullPath/Test-Path/New-Item can throw a .NET PathTooLongException
    # (an English message) instead of returning cleanly once a literal path nears 260 chars.
    if ($out.Length -ge 260) { Fail $lenFail }
    try { $full = [System.IO.Path]::GetFullPath($out) } catch { Fail $lenFail }
    if ($full.Length -ge 260) { Fail $lenFail }
    $outDirPart = Split-Path $full -Parent
    Test-OutputGuard $outDirPart $root
    if (Test-Path -LiteralPath $full) { Fail "มีไฟล์ปลายทางอยู่แล้ว จะไม่เขียนทับ: $full" }
    if ($outDirPart -and -not (Test-Path -LiteralPath $outDirPart)) { New-Item -ItemType Directory -Path $outDirPart -Force | Out-Null }
    return $full
  }

  if ($outDir) {
    if ($outDir.Length -ge 260) { Fail $lenFail }
    if (Test-Path -LiteralPath $outDir -PathType Leaf) { Fail "-OutDir ต้องเป็นโฟลเดอร์ แต่พาธนี้เป็นไฟล์: $outDir" }
    if (Test-Path -LiteralPath $outDir -PathType Container) {
      $dir = (Resolve-Path -LiteralPath $outDir).ProviderPath
    } else {
      # A missing -OutDir is created for the caller (one level only: the parent must already exist, so a mistyped
      # path cannot scatter files into a new tree). Guard first; the actual New-Item happens below, after
      # Test-OutputGuard. Resolve to a full path first so the guard sees where it would really be created.
      try { $dir = [System.IO.Path]::GetFullPath($outDir) } catch { Fail $lenFail }
      $dir = $dir.TrimEnd('\', '/')
      if ($dir.Length -ge 260) { Fail $lenFail }
      Test-OutputGuard $dir $root
      $outParent = Split-Path $dir -Parent
      if (-not $outParent -or -not (Test-Path -LiteralPath $outParent -PathType Container)) {
        Fail "ไม่พบโฟลเดอร์แม่ของ -OutDir: $outParent (สร้างโฟลเดอร์ปลายทางนี้ให้ได้ชั้นเดียว) ให้สร้างโฟลเดอร์เองก่อน หรือไม่ใส่ -OutDir เพื่อใช้ค่าเริ่มต้น"
      }
    }
  } else {
    $dir = Join-Path (Get-OutputRoot) ("{0}\{1}" -f $kindFolder, $safeBase)
  }
  Test-OutputGuard $dir $root
  if ($dir.Length -ge 260) { Fail $lenFail }
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

  $stem = "{0}_{1}_{2}" -f $safeBase, $suffix, $today
  $candidate = Join-Path $dir ($stem + $ext)
  if ($candidate.Length -ge 260) { Fail $lenFail }
  $i = 2
  while (Test-Path -LiteralPath $candidate) {
    $candidate = Join-Path $dir ("{0}_{1}{2}" -f $stem, $i, $ext)
    if ($candidate.Length -ge 260) { Fail $lenFail }
    $i++
  }
  return $candidate
}

# Opens a brand-new file for writing at $path (CreateNew semantics: never overwrites). If another
# process/script instance raced us and created that exact name first (or an earlier de-dup check
# in Get-NewPath is stale), retries the next "<stem>_N<ext>" suffix - the same naming Get-NewPath
# itself uses - up to $maxTries times. Returns @{ Stream; Path } (Path is the name actually used).
function New-FileExclusive([string]$path, [int]$maxTries = 20) {
  $dir = Split-Path $path -Parent
  $ext = [System.IO.Path]::GetExtension($path)
  $stem = [System.IO.Path]::GetFileNameWithoutExtension($path)
  $candidate = $path
  $n = 2
  for ($try = 0; $try -le $maxTries; $try++) {
    try {
      $fs = [System.IO.File]::Open($candidate, 'CreateNew', 'Write', 'None')
      return @{ Stream = $fs; Path = $candidate }
    } catch [System.IO.IOException] {
      $candidate = Join-Path $dir ("{0}_{1}{2}" -f $stem, $n, $ext)
      $n++
    }
  }
  Fail 'สร้างไฟล์ปลายทางไม่ได้ มีไฟล์ชื่อซ้ำอยู่หลายไฟล์ ลองใหม่อีกครั้งหรือเลือกโฟลเดอร์อื่น'
}

# ---------------------------------------------------------------------------------------------
# Small text helpers shared by review-file.ps1 (preview rendering) and the rewrite code below.
# ---------------------------------------------------------------------------------------------
function Get-QuotedLines([string]$text) {
  $norm = $text -replace "`r`n", "`n" -replace "`r", "`n"
  return @($norm -split "`n" | ForEach-Object { "> $_" })
}

function Get-OneLine([string]$text) {
  return ($text -replace '\s*\r?\n\s*', ' ')
}

# F2 (docs/v2-plan.md 3.8): a short single-line excerpt for a "[คงไว้: ¶7 ...]" anchor label -
# collapsed to one line and capped so a long kept paragraph/table cell does not blow up the
# preview's Before/After section.
function Get-ShortText([string]$text, [int]$max = 60) {
  $t = (Get-OneLine $text).Trim()
  if ($t.Length -gt $max) { $t = $t.Substring(0, $max).TrimEnd() + '…' }
  return $t
}

# Wraps a possibly-null / possibly-scalar JSON property as a plain object[] (Windows PowerShell 5.1's
# ConvertFrom-Json returns a bare object, not a 1-element array, for a JSON array with one element).
# Always assign the result to a variable before enumerating it (see review-file.ps1 for the long
# explanation of the two pipeline quirks this avoids).
function Get-AsArray($x) {
  $list = New-Object System.Collections.ArrayList
  if ($null -eq $x) { return , $list.ToArray() }
  foreach ($it in @($x)) { [void]$list.Add($it) }
  return , $list.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Whole-paragraph edit rule used by Get-ParagraphEditPlan (Set-ParagraphTextPlain, the table-cell rewrite of
# write-rewrite.ps1, for a paragraph WITHOUT a line break / special hyphen / link / field). Neither apply-revisions.ps1
# nor the table-cell rewrite uses it for a paragraph with a line break or a special hyphen any more: both edit those by
# units (Get-ParagraphUnitPlan below) and keep the break/hyphen in place.
# ---------------------------------------------------------------------------------------------
$script:RevAllowedChildren = @('pPr', 'r', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'permStart', 'permEnd')
$script:RevAllowedInRun = @('rPr', 't', 'tab', 'br', 'lastRenderedPageBreak', 'noBreakHyphen', 'softHyphen')

# Text of a paragraph's runs (t = itself, tab = "`t") as one string plus, per run that has text,
# its character range [Start, End) and rPr node. Only direct child runs are read (callers have
# already rejected any other child besides the harmless ones in $script:RevAllowedChildren).
function Get-RunSegments($p) {
  $ns = $script:DocxNs
  $segs = New-Object System.Collections.ArrayList
  $sb = New-Object System.Text.StringBuilder
  foreach ($r in $p.SelectNodes('w:r', $ns)) {
    $start = $sb.Length
    foreach ($rc in $r.ChildNodes) {
      if ($rc.NodeType -ne 'Element') { continue }
      if ($rc.LocalName -eq 't') { [void]$sb.Append($rc.InnerText) }
      elseif ($rc.LocalName -eq 'tab') { [void]$sb.Append("`t") }
    }
    if ($sb.Length -gt $start) { [void]$segs.Add(@{ Start = $start; End = $sb.Length; RPr = $r.SelectSingleNode('w:rPr', $ns) }) }
  }
  return @{ Segs = $segs; Text = $sb.ToString() }
}

# Returns the Thai reason a change to paragraph node $p has to be written as "delete the whole paragraph,
# insert the new text" by Get-ParagraphEditPlan, or $null when the paragraph is edited word by word - or cannot be
# edited at all (it has a hyperlink/field/picture/footnote/comment/existing tracked change). Needs $script:DocxNs
# (set by Read-Docx). Neither apply-revisions.ps1 nor Set-ParagraphTextPlain reaches the line-break / special-hyphen
# reasons any more (they edit such a paragraph by units); only extract-text.ps1's labels may still look at them.
function Get-WholeParagraphReason($p) {
  $ns = $script:DocxNs
  foreach ($c in $p.ChildNodes) {
    if ($c.NodeType -ne 'Element') { continue }
    if ($script:RevAllowedChildren -notcontains $c.LocalName) { return $null }
    if ($c.LocalName -eq 'r') {
      foreach ($rc in $c.ChildNodes) {
        if ($rc.NodeType -eq 'Element' -and $script:RevAllowedInRun -notcontains $rc.LocalName) { return $null }
      }
    }
  }
  $runs = @($p.SelectNodes('w:r', $ns))
  $textRuns = @($runs | Where-Object { $_.SelectSingleNode('w:t', $ns) })
  $rPrXml = @($textRuns | ForEach-Object { $x = $_.SelectSingleNode('w:rPr', $ns); if ($x) { $x.OuterXml } else { '' } } | Select-Object -Unique)
  $hasBreak = [bool]$p.SelectSingleNode('w:r/w:br', $ns)
  if ($rPrXml.Count -le 1 -and -not $hasBreak) { return $null }
  if ($hasBreak) { return 'มีการขึ้นบรรทัดใหม่ในย่อหน้า' }
  if ($p.SelectSingleNode('w:r/w:noBreakHyphen | w:r/w:softHyphen', $ns)) { return 'มีขีดกลางพิเศษ (ขีดกลางไม่ตัดคำหรือขีดกลางตัดคำอ่อน) ในย่อหน้า' }
  $oldRaw = Get-ParagraphBodyText $p
  $segInfo = Get-RunSegments $p
  if (-not [string]::Equals($segInfo.Text, $oldRaw, [StringComparison]::Ordinal)) { return 'อ่านข้อความของย่อหน้าตามรูปแบบของแต่ละส่วนไม่ตรงกับข้อความที่ดึงมา' }
  return $null
}

# ---------------------------------------------------------------------------------------------
# Word-by-word paragraph edit, shared by apply-revisions.ps1 (shows the change as tracked changes) and
# the table-cell rewrite below (writes the new text as plain runs, no tracked changes). Both take the same
# decisions - which paragraphs cannot be edited at all, word-level vs whole-paragraph, which run
# formatting each piece of text gets - from the functions here, so the two never drift apart.
# ---------------------------------------------------------------------------------------------

# Word-level diff of two strings (tokens: a run of spaces, a run of ASCII letters/digits, or one other
# character with its combining marks). Compiled on first use only, so the scripts that never edit a
# paragraph do not pay for it.
function Initialize-HzDiff {
  if ('HzDiff' -as [type]) { return }
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
public static class HzDiff {
  public static List<string> Tokens(string s) {
    var list = new List<string>();
    int i = 0;
    while (i < s.Length) {
      int j = i + 1;
      char c = s[i];
      if (char.IsWhiteSpace(c) && c != '\t') { while (j < s.Length && char.IsWhiteSpace(s[j]) && s[j] != '\t') j++; }
      else if (c < 128 && char.IsLetterOrDigit(c)) { while (j < s.Length && s[j] < 128 && char.IsLetterOrDigit(s[j])) j++; }
      else { while (j < s.Length && CharUnicodeInfo.GetUnicodeCategory(s[j]) == UnicodeCategory.NonSpacingMark) j++; }
      list.Add(s.Substring(i, j - i));
      i = j;
    }
    return list;
  }
  // ops: 0 equal, -1 deleted, 1 inserted; consecutive ops merged
  public static List<KeyValuePair<int, string>> Diff(string a, string b) {
    var x = Tokens(a); var y = Tokens(b);
    int n = x.Count, m = y.Count;
    var raw = new List<KeyValuePair<int, string>>();
    if ((long)n * m > 4000000L) {
      raw.Add(new KeyValuePair<int, string>(-1, a)); raw.Add(new KeyValuePair<int, string>(1, b));
    } else {
      var d = new int[n + 1, m + 1];
      for (int i = n - 1; i >= 0; i--)
        for (int j = m - 1; j >= 0; j--)
          d[i, j] = x[i] == y[j] ? d[i + 1, j + 1] + 1 : Math.Max(d[i + 1, j], d[i, j + 1]);
      int p = 0, q = 0;
      while (p < n && q < m) {
        if (x[p] == y[q]) { raw.Add(new KeyValuePair<int, string>(0, x[p])); p++; q++; }
        else if (d[p + 1, q] >= d[p, q + 1]) { raw.Add(new KeyValuePair<int, string>(-1, x[p])); p++; }
        else { raw.Add(new KeyValuePair<int, string>(1, y[q])); q++; }
      }
      while (p < n) { raw.Add(new KeyValuePair<int, string>(-1, x[p])); p++; }
      while (q < m) { raw.Add(new KeyValuePair<int, string>(1, y[q])); q++; }
    }
    var merged = new List<KeyValuePair<int, string>>();
    foreach (var op in raw) {
      if (op.Value.Length == 0) continue;
      if (merged.Count > 0 && merged[merged.Count - 1].Key == op.Key)
        merged[merged.Count - 1] = new KeyValuePair<int, string>(op.Key, merged[merged.Count - 1].Value + op.Value);
      else merged.Add(op);
    }
    return merged;
  }
}
'@
}

# Why paragraph $p cannot be edited at all, or $null when it can: it holds something other than plain runs
# (a hyperlink, a field, an existing tracked change, a comment mark ...) or a run holding something other than
# text/tab/break (a picture, a field character, a footnote reference ...). Thai text without the closing
# advice - each caller adds its own (apply-revisions.ps1: "ให้แก้เองจากข้อเสนอในแชต").
function Get-ParagraphEditBlocker($p) {
  foreach ($c in $p.ChildNodes) {
    if ($c.NodeType -ne 'Element') { continue }
    if ($script:RevAllowedChildren -notcontains $c.LocalName) { return "มี $($c.LocalName) (ลิงก์/ฟิลด์/tracked changes/comment เดิม)" }
    if ($c.LocalName -eq 'r') {
      foreach ($rc in $c.ChildNodes) {
        if ($rc.NodeType -eq 'Element' -and $script:RevAllowedInRun -notcontains $rc.LocalName) { return "มี $($rc.LocalName) (รูป/ฟิลด์/เชิงอรรถ)" }
      }
    }
  }
  return $null
}

# ---------------------------------------------------------------------------------------------
# Paragraph edit by "units" (apply-revisions.ps1 as tracked changes, Set-ParagraphTextByUnits as plain text for the
# table-cell rewrite). A paragraph with a line break (w:br), a special hyphen, a hyperlink or a field is edited word by
# word like any other paragraph, but its break / hyphen / link / field must stay exactly where it is. The paragraph is
# read as an ordered list of units:
#   text   the text of one run (a run holding a w:br or a hyphen is cut at it)
#   br     one w:br: one character of the extracted text (a space)
#   hyphen one w:noBreakHyphen / w:softHyphen: no character in the extracted text
#   mark   a bookmark / permission range marker: no character
#   lock   a w:hyperlink, a w:fldSimple, or a field begin..end range of runs (w:fldChar): its text is part of the
#          paragraph text but its XML is never touched
# Get-ParagraphUnitPlan checks BEFORE anything is changed that every br / hyphen / lock lies inside text that does not
# change, and returns the Thai reason the paragraph is skipped when one does not. It never falls back to the old
# "delete the whole paragraph, insert the new text", which dropped the break or the link without a word.
# The shared word-level functions below (Get-ParagraphEditPlan, Set-ParagraphTextPlain for a paragraph without any of
# these) are untouched. The table-cell rewrite (write-rewrite.ps1) uses the same plan for a paragraph with a line break
# or a special hyphen, writes it as plain runs (no w:ins / w:del) and refuses, with a message, a change that touches one.
# ---------------------------------------------------------------------------------------------
$script:UnitMarkNames = @('bookmarkStart', 'bookmarkEnd', 'permStart', 'permEnd')

# True when paragraph $p holds something the unit edit is for: a hyperlink, a field, a line break or a special hyphen.
# Any other paragraph is edited by the plain functions (Get-ParagraphEditPlan ...), exactly as before.
function Test-ParagraphHasUnitContent($p) {
  return [bool]$p.SelectSingleNode('w:hyperlink | w:fldSimple | w:r/w:fldChar | w:r/w:br | w:r/w:noBreakHyphen | w:r/w:softHyphen', $script:DocxNs)
}

# Reads paragraph $p as units (see above). Returns @{ Blocker = <Thai reason, no closing advice>; Units = $null } when
# the paragraph holds something that cannot be edited at all (a picture, a footnote reference, a comment mark, an
# earlier tracked change ... outside a link/field), otherwise @{ Blocker = $null; Units; Text } where Text is the
# concatenated text of the units (it must equal Get-ParagraphBodyText $p). Every unit has Kind, S and E (character range
# [S, E) in Text; S = E for a hyphen, a mark and a lock without text).
function Get-ParagraphUnits($p) {
  $ns = $script:DocxNs
  $units = New-Object System.Collections.ArrayList
  $sb = New-Object System.Text.StringBuilder
  $depth = 0
  $group = $null
  foreach ($c in $p.ChildNodes) {
    if ($c.NodeType -ne 'Element') { continue }
    $name = $c.LocalName
    if ($name -eq 'pPr' -or $name -eq 'proofErr') { continue }
    # a field written as runs: begin ... end (also a stray begin or end: the runs from it to the paragraph's end/start)
    $isFieldRun = ($name -eq 'r' -and [bool]$c.SelectSingleNode('w:fldChar', $ns))
    if ($depth -gt 0 -or $isFieldRun) {
      if ($null -eq $group) { $group = @{ Kind = 'lock'; What = 'ฟิลด์'; Nodes = (New-Object System.Collections.ArrayList); S = $sb.Length; E = 0 } }
      [void]$group.Nodes.Add($c)
      [void]$sb.Append((Get-NodeText $c -Raw))
      if ($name -eq 'r') {
        foreach ($fc in $c.SelectNodes('w:fldChar', $ns)) {
          $type = $fc.GetAttribute('fldCharType', $W_NS)
          if ($type -eq 'begin') { $depth++ } elseif ($type -eq 'end' -and $depth -gt 0) { $depth-- }
        }
      }
      if ($depth -eq 0) { $group.E = $sb.Length; [void]$units.Add($group); $group = $null }
      continue
    }
    if ($name -eq 'hyperlink' -or $name -eq 'fldSimple') {
      $s = $sb.Length
      [void]$sb.Append((Get-NodeText $c -Raw))
      [void]$units.Add(@{ Kind = 'lock'; What = $(if ($name -eq 'hyperlink') { 'ลิงก์' } else { 'ฟิลด์' }); Nodes = @($c); S = $s; E = $sb.Length })
      continue
    }
    if ($script:UnitMarkNames -contains $name) { [void]$units.Add(@{ Kind = 'mark'; Node = $c; S = $sb.Length; E = $sb.Length }); continue }
    if ($name -ne 'r') { return @{ Blocker = "มี $name (ลิงก์/ฟิลด์/tracked changes/comment เดิม)"; Units = $null; Text = $null } }
    foreach ($rc in $c.ChildNodes) {
      if ($rc.NodeType -eq 'Element' -and $script:RevAllowedInRun -notcontains $rc.LocalName) { return @{ Blocker = "มี $($rc.LocalName) (รูป/ฟิลด์/เชิงอรรถ)"; Units = $null; Text = $null } }
    }
    $rPr = $c.SelectSingleNode('w:rPr', $ns)
    $cur = $null
    $flush = {
      if ($cur) {
        if ($cur.Sb.Length -gt 0) { $cur.Text = $cur.Sb.ToString(); $cur.E = $sb.Length; [void]$units.Add($cur) }
        $cur = $null
      }
    }
    foreach ($rc in $c.ChildNodes) {
      if ($rc.NodeType -ne 'Element') { continue }
      $rn = $rc.LocalName
      if ($rn -eq 't' -or $rn -eq 'tab') {
        $piece = if ($rn -eq 't') { $rc.InnerText } else { "`t" }
        if (-not $cur) { $cur = @{ Kind = 'text'; RPr = $rPr; S = $sb.Length; E = 0; Text = ''; Sb = (New-Object System.Text.StringBuilder) } }
        [void]$cur.Sb.Append($piece); [void]$sb.Append($piece)
      } elseif ($rn -eq 'br') {
        . $flush
        [void]$units.Add(@{ Kind = 'br'; Node = $rc; RPr = $rPr; S = $sb.Length; E = $sb.Length + 1 })
        [void]$sb.Append(' ')
      } elseif ($rn -eq 'noBreakHyphen' -or $rn -eq 'softHyphen') {
        . $flush
        [void]$units.Add(@{ Kind = 'hyphen'; Node = $rc; RPr = $rPr; S = $sb.Length; E = $sb.Length })
      }
    }
    . $flush
  }
  if ($group) { $group.E = $sb.Length; [void]$units.Add($group) }
  return @{ Blocker = $null; Units = $units; Text = $sb.ToString() }
}

# Word-level change from $a to $b as ops: @{ Kind = 'eq' | 'del' | 'ins'; A; B; Text }. [A, B) is the range in the OLD
# text ($oldBase = where $a starts in it); an insertion has A = B = the old position it goes in.
function Get-DiffOpList([string]$a, [string]$b, [int]$oldBase) {
  Initialize-HzDiff
  $ops = New-Object System.Collections.ArrayList
  $pos = $oldBase
  foreach ($d in [HzDiff]::Diff($a, $b)) {
    if ($d.Key -gt 0) { [void]$ops.Add(@{ Kind = 'ins'; A = $pos; B = $pos; Text = $d.Value }); continue }
    $end = $pos + $d.Value.Length
    [void]$ops.Add(@{ Kind = $(if ($d.Key -eq 0) { 'eq' } else { 'del' }); A = $pos; B = $end; Text = $d.Value })
    $pos = $end
  }
  return , $ops
}

# The first br / hyphen / lock unit that the ops do not leave alone, or $null. A br or a lock with text is left alone
# when its whole range lies inside one unchanged ('eq') op; a hyphen (or a lock without text) sits between two
# characters and is left alone when it lies strictly inside an unchanged op (or at the very start/end of the text next
# to one). Marks never conflict (they can go anywhere). An insertion right before or after a br / lock is fine.
function Find-UnitConflict($units, $ops, [int]$len) {
  foreach ($u in $units) {
    if ($u.Kind -eq 'text' -or $u.Kind -eq 'mark') { continue }
    $zero = ($u.E -eq $u.S)
    $ok = $false
    foreach ($op in $ops) {
      if ($op.Kind -ne 'eq') { continue }
      if ($zero) {
        if (($op.A -lt $u.S -and $u.S -lt $op.B) -or ($u.S -eq 0 -and $op.A -eq 0) -or ($u.S -eq $len -and $op.B -eq $len)) { $ok = $true; break }
      } elseif ($op.A -le $u.S -and $u.E -le $op.B) { $ok = $true; break }
    }
    if (-not $ok) { return $u }
  }
  return $null
}

# Second try when the plain diff puts a change inside a link/field (it can, when the same words stand before and in the
# link): cut the old text at each link/field with text, find that same text again in the new text (in order, first
# occurrence after the previous one) and diff only the pieces between. Returns the ops, or $null when a link/field
# text is not found in the new text (then the change really is inside it).
function Get-LockedSegmentOps([string]$oldRaw, [string]$newRaw, $units) {
  $ops = New-Object System.Collections.ArrayList
  $oldPos = 0; $newPos = 0
  foreach ($u in $units) {
    if ($u.Kind -ne 'lock' -or $u.E -le $u.S) { continue }
    $ltxt = $oldRaw.Substring($u.S, $u.E - $u.S)
    $idx = $newRaw.IndexOf($ltxt, $newPos, [StringComparison]::Ordinal)
    if ($idx -lt 0) { return $null }
    foreach ($o in (Get-DiffOpList $oldRaw.Substring($oldPos, $u.S - $oldPos) $newRaw.Substring($newPos, $idx - $newPos) $oldPos)) { [void]$ops.Add($o) }
    [void]$ops.Add(@{ Kind = 'eq'; A = $u.S; B = $u.E; Text = $ltxt })
    $oldPos = $u.E; $newPos = $idx + $ltxt.Length
  }
  foreach ($o in (Get-DiffOpList $oldRaw.Substring($oldPos) $newRaw.Substring($newPos) $oldPos)) { [void]$ops.Add($o) }
  $merged = New-Object System.Collections.ArrayList
  foreach ($o in $ops) {
    $last = if ($merged.Count) { $merged[$merged.Count - 1] } else { $null }
    if ($last -and $last.Kind -eq $o.Kind -and $last.B -eq $o.A) { $last.B = $o.B; $last.Text = $last.Text + $o.Text } else { [void]$merged.Add($o) }
  }
  return , $merged
}

function Get-UnitConflictReason($u) {
  if ($u.Kind -eq 'br') { return 'จุดแก้คร่อมหรือแตะตำแหน่งขึ้นบรรทัดใหม่ในย่อหน้า จึงคงการขึ้นบรรทัดไว้ไม่ได้ ให้แก้เองจากข้อเสนอในแชต' }
  if ($u.Kind -eq 'hyphen') { return 'จุดแก้แตะตำแหน่งขีดกลางพิเศษ (ขีดกลางไม่ตัดคำหรือขีดกลางตัดคำอ่อน) จึงคงขีดกลางไว้ไม่ได้ ให้แก้เองจากข้อเสนอในแชต' }
  return "จุดแก้อยู่ในข้อความของ$($u.What) ซึ่งต้องคงเดิม ให้แก้เองจากข้อเสนอในแชต"
}

# How paragraph $p (one with Test-ParagraphHasUnitContent) changes to $newText: @{ Reason = <Thai reason it is skipped,
# with the closing advice> } (plus Conflict = the unit the change touches, Units, OldRaw when that is the reason) or
# @{ Reason = $null; Units; Ops; OldRaw; NewRaw } (NewRaw = the new text with the old leading/trailing whitespace
# kept, like Get-ParagraphEditPlan). Changes nothing.
function Get-ParagraphUnitPlan($p, [string]$newText) {
  $info = Get-ParagraphUnits $p
  if ($info.Blocker) { return @{ Reason = "$($info.Blocker) ให้แก้เองจากข้อเสนอในแชต" } }
  $oldRaw = Get-ParagraphBodyText $p
  if (-not [string]::Equals($info.Text, $oldRaw, [StringComparison]::Ordinal)) { return @{ Reason = 'อ่านข้อความของย่อหน้าตามส่วนต่าง ๆ ไม่ตรงกับข้อความที่ดึงมา ให้แก้เองจากข้อเสนอในแชต' } }
  $lead = $oldRaw.Substring(0, $oldRaw.Length - $oldRaw.TrimStart().Length)
  $trail = $oldRaw.Substring($oldRaw.TrimEnd().Length)
  $newRaw = $lead + $newText.Trim() + $trail
  $ops = Get-DiffOpList $oldRaw $newRaw 0
  $conflict = Find-UnitConflict $info.Units $ops $oldRaw.Length
  if ($conflict -and @($info.Units | Where-Object { $_.Kind -eq 'lock' -and $_.E -gt $_.S }).Count) {
    $alt = Get-LockedSegmentOps $oldRaw $newRaw $info.Units
    if ($alt -and -not (Find-UnitConflict $info.Units $alt $oldRaw.Length)) { $ops = $alt; $conflict = $null }
  }
  if ($conflict) { return @{ Reason = (Get-UnitConflictReason $conflict); Conflict = $conflict; Units = $info.Units; OldRaw = $oldRaw } }
  return @{ Reason = $null; Units = $info.Units; Ops = $ops; OldRaw = $oldRaw; NewRaw = $newRaw }
}

# The word-level change from $oldRaw to $newRaw as an ordered list of pieces @{ Kind = 'eq' | 'del' | 'ins';
# Text; RPr }. $segs = $null: the paragraph has one formatting ($uniformRPr) for every piece. Otherwise $segs is
# the run segments of Get-RunSegments: an equal or deleted piece keeps the formatting of the run its text came
# from (a piece spanning several runs is cut at each run's edge), an inserted piece takes the formatting of the
# run just before its position (the first run at the very start).
function Get-ParagraphDiffPieces([string]$oldRaw, [string]$newRaw, $segs, $uniformRPr) {
  Initialize-HzDiff
  $pieces = New-Object System.Collections.ArrayList
  if ($null -eq $segs) {
    foreach ($op in [HzDiff]::Diff($oldRaw, $newRaw)) {
      $kind = if ($op.Key -eq 0) { 'eq' } elseif ($op.Key -lt 0) { 'del' } else { 'ins' }
      [void]$pieces.Add(@{ Kind = $kind; Text = $op.Value; RPr = $uniformRPr })
    }
    return , $pieces
  }
  $pos = 0
  foreach ($op in [HzDiff]::Diff($oldRaw, $newRaw)) {
    if ($op.Key -gt 0) {
      $prev = $null
      foreach ($s in $segs) { if ($s.Start -le [Math]::Max($pos - 1, 0) -and [Math]::Max($pos - 1, 0) -lt $s.End) { $prev = $s; break } }
      [void]$pieces.Add(@{ Kind = 'ins'; Text = $op.Value; RPr = $(if ($prev) { $prev.RPr } else { $null }) })
      continue
    }
    $end = $pos + $op.Value.Length
    foreach ($s in $segs) {
      $from = [Math]::Max($pos, $s.Start); $to = [Math]::Min($end, $s.End)
      if ($to -le $from) { continue }
      [void]$pieces.Add(@{ Kind = $(if ($op.Key -eq 0) { 'eq' } else { 'del' }); Text = $op.Value.Substring($from - $pos, $to - $from); RPr = $s.RPr })
    }
    $pos = $end
  }
  return , $pieces
}

# How paragraph $p (one that Get-ParagraphEditBlocker accepts) changes to $newText: @{ Mode = 'word' | 'whole';
# Why (the Get-WholeParagraphReason text, 'whole' only); Pieces (Get-ParagraphDiffPieces, 'word' only); Runs;
# FirstRPr; OldRaw; NewRaw (the new text with the old leading/trailing whitespace kept); NewTrim }. 'whole' is the
# paragraph that cannot be edited word by word (line break, special hyphen ...): the old text goes and the
# new text is written once, in the formatting of the first text run.
function Get-ParagraphEditPlan($p, [string]$newText) {
  $ns = $script:DocxNs
  $runs = @($p.SelectNodes('w:r', $ns))
  $textRuns = @($runs | Where-Object { $_.SelectSingleNode('w:t', $ns) })
  $firstRPr = if ($textRuns.Count) { $textRuns[0].SelectSingleNode('w:rPr', $ns) } else { $null }
  $rPrXml = @($textRuns | ForEach-Object { $x = $_.SelectSingleNode('w:rPr', $ns); if ($x) { $x.OuterXml } else { '' } } | Select-Object -Unique)
  $hasBreak = [bool]$p.SelectSingleNode('w:r/w:br', $ns)
  $oldRaw = Get-ParagraphBodyText $p
  $lead = $oldRaw.Substring(0, $oldRaw.Length - $oldRaw.TrimStart().Length)
  $trail = $oldRaw.Substring($oldRaw.TrimEnd().Length)
  $newRaw = $lead + $newText.Trim() + $trail
  $plan = @{ Mode = 'word'; Why = $null; Pieces = $null; Runs = $runs; FirstRPr = $firstRPr; OldRaw = $oldRaw; NewRaw = $newRaw; NewTrim = $newText.Trim() }
  if ($rPrXml.Count -le 1 -and -not $hasBreak) {
    $plan.Pieces = Get-ParagraphDiffPieces $oldRaw $newRaw $null $firstRPr
    return $plan
  }
  $why = Get-WholeParagraphReason $p
  if ($why) { $plan.Mode = 'whole'; $plan.Why = $why; return $plan }
  $segInfo = Get-RunSegments $p
  $plan.Pieces = Get-ParagraphDiffPieces $oldRaw $newRaw $segInfo.Segs $null
  return $plan
}

# --- plain-text twin of apply-revisions.ps1's Get-UnitNodes / Revise-ParagraphByUnits (names differ on purpose: the
# apply script defines its own copies and must not be changed) ---

# A run holding a copy of the break / hyphen element $node, with the formatting of the run it came from.
function New-PlainSpecialRun($rPr, $node) {
  $r = New-W 'r'
  if ($rPr) { [void]$r.AppendChild($rPr.CloneNode($true)) }
  [void]$r.AppendChild($node.CloneNode($true))
  return $r
}

# The rPr for text inserted at old position $pos: the run text just before it (never a link or field), the first run
# text at the very start.
function Get-PlainUnitInsertRPr($units, [int]$pos) {
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

# Adds a unit that takes no characters (a mark, a hyphen, a link/field without text) to the item list, in its place.
function Add-PlainUnitZero($u, $items) {
  if ($u.Kind -eq 'mark') { [void]$items.Add(@{ Kind = 'node'; Node = $u.Node }) }
  elseif ($u.Kind -eq 'hyphen') { [void]$items.Add(@{ Kind = 'node'; Node = (New-PlainSpecialRun $u.RPr $u.Node) }) }
  else { foreach ($n in $u.Nodes) { [void]$items.Add(@{ Kind = 'node'; Node = $n }) } }
}

# Builds, without touching the paragraph, the new child list (after w:pPr) for a plan from Get-ParagraphUnitPlan, as
# plain runs: unchanged text keeps the formatting of the run it came from, deleted text is dropped, inserted text takes
# the formatting of the run before it; every br / hyphen / mark / link / field is put back in its place. Neighbouring
# pieces of text with the same formatting (and no break/hyphen/mark between them) become one run.
function Get-PlainUnitNodes($plan) {
  $units = $plan.Units
  $items = New-Object System.Collections.ArrayList
  $ui = 0
  foreach ($op in $plan.Ops) {
    if ($op.Kind -eq 'ins') {
      # a bookmark / permission END right at the insertion point comes after the new text, so the new text stays inside it
      while ($ui -lt $units.Count -and $units[$ui].E -eq $units[$ui].S -and $units[$ui].S -le $op.A -and -not ($units[$ui].Kind -eq 'mark' -and $units[$ui].S -eq $op.A -and $units[$ui].Node.LocalName -like '*End')) { Add-PlainUnitZero $units[$ui] $items; $ui++ }
      [void]$items.Add(@{ Kind = 'text'; RPr = (Get-PlainUnitInsertRPr $units $op.A); Text = $op.Text })
      continue
    }
    while ($ui -lt $units.Count) {
      $u = $units[$ui]
      if ($u.S -ge $op.B) { break }
      if ($u.E -eq $u.S) { Add-PlainUnitZero $u $items; $ui++; continue }
      $from = [Math]::Max($op.A, $u.S); $to = [Math]::Min($op.B, $u.E)
      if ($to -gt $from) {
        if ($u.Kind -eq 'text') {
          if ($op.Kind -eq 'eq') { [void]$items.Add(@{ Kind = 'text'; RPr = $u.RPr; Text = $op.Text.Substring($from - $op.A, $to - $from) }) }
        } elseif ($u.Kind -eq 'br') {
          [void]$items.Add(@{ Kind = 'node'; Node = (New-PlainSpecialRun $u.RPr $u.Node) })
        } elseif ($from -eq $u.S) {
          foreach ($n in $u.Nodes) { [void]$items.Add(@{ Kind = 'node'; Node = $n }) }
        }
      }
      if ($u.E -le $op.B) { $ui++ } else { break }
    }
  }
  while ($ui -lt $units.Count -and $units[$ui].E -eq $units[$ui].S) { Add-PlainUnitZero $units[$ui] $items; $ui++ }
  $nodes = New-Object System.Collections.ArrayList
  $pend = $null
  foreach ($it in $items) {
    if ($it.Kind -eq 'text') {
      if ($it.Text -eq '') { continue }
      if ($pend -and [object]::ReferenceEquals($pend.RPr, $it.RPr)) { [void]$pend.Sb.Append($it.Text); continue }
      if ($pend) { [void]$nodes.Add((New-PlainRun $pend.RPr $pend.Sb.ToString())) }
      $pend = @{ RPr = $it.RPr; Sb = (New-Object System.Text.StringBuilder) }
      [void]$pend.Sb.Append($it.Text)
    } else {
      if ($pend) { [void]$nodes.Add((New-PlainRun $pend.RPr $pend.Sb.ToString())); $pend = $null }
      [void]$nodes.Add($it.Node)
    }
  }
  if ($pend) { [void]$nodes.Add((New-PlainRun $pend.RPr $pend.Sb.ToString())) }
  return , $nodes
}

# The old text of a paragraph read as units, for an error message: each line break shown as " ⏎ ", each special hyphen
# as "[ขีดพิเศษ]" (the plain text has no character for it).
function Get-UnitDisplayText($units, [string]$oldRaw) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($u in $units) {
    if ($u.Kind -eq 'text' -or $u.Kind -eq 'lock') { [void]$sb.Append($oldRaw.Substring($u.S, $u.E - $u.S)) }
    elseif ($u.Kind -eq 'br') { [void]$sb.Append(' ⏎ ') }
    elseif ($u.Kind -eq 'hyphen') { [void]$sb.Append('[ขีดพิเศษ]') }
  }
  return $sb.ToString()
}

# Number of w:br and of w:noBreakHyphen / w:softHyphen in paragraph $p (tracked deletions left out).
function Get-ParagraphBreakCounts($p) {
  $ns = $script:DocxNs
  return @{
    Breaks = @($p.SelectNodes('.//w:br[not(ancestor::w:del)]', $ns)).Count
    Hyphens = @($p.SelectNodes('.//w:noBreakHyphen[not(ancestor::w:del)] | .//w:softHyphen[not(ancestor::w:del)]', $ns)).Count
  }
}

# The table-cell rewrite of a paragraph with a line break / special hyphen (or a link / field - never reached, the caller
# refuses those): the word-level change of Get-ParagraphUnitPlan written as plain runs, every break / hyphen / bookmark
# where it was. A change that touches a break or a hyphen is refused with the cell's coordinate ($where, e.g.
# "S1: ตาราง ¶3 r11c2") and the old text with the break marked, never written as "the whole paragraph" (the break would
# be lost). Checks before the file is written that the text read back is the new text and that no break / hyphen was
# lost. Returns @{ Whole = $false; Why = $null; Breaks; Hyphens } (how many it kept).
function Set-ParagraphTextByUnits($p, [string]$newText, [string]$where) {
  $prefix = if ($where) { "${where}: " } else { '' }
  $plan = Get-ParagraphUnitPlan $p $newText
  if ($plan.Reason) {
    if ($plan.Conflict -and $plan.Conflict.Kind -eq 'br') {
      Fail ($prefix + 'จุดแก้แตะตำแหน่งขึ้นบรรทัดใหม่ในเซลล์นี้ จึงคงการขึ้นบรรทัดไว้ไม่ได้ (ข้อความเดิม โดย ⏎ คือตำแหน่งขึ้นบรรทัดใหม่: ' + (Get-UnitDisplayText $plan.Units $plan.OldRaw) + ') ให้คงคำที่อยู่ติดกับ ⏎ ไว้ตามเดิม หรือเอาเซลล์นี้ออกจาก cells แล้วแจ้งผู้ใช้แก้เอง (ไม่สร้างไฟล์)')
    }
    if ($plan.Conflict -and $plan.Conflict.Kind -eq 'hyphen') {
      Fail ($prefix + 'จุดแก้แตะตำแหน่งขีดกลางพิเศษ (ขีดกลางไม่ตัดคำหรือขีดกลางตัดคำอ่อน) ในเซลล์นี้ จึงคงขีดกลางไว้ไม่ได้ (ข้อความเดิม โดย [ขีดพิเศษ] คือตำแหน่งขีดกลางพิเศษ: ' + (Get-UnitDisplayText $plan.Units $plan.OldRaw) + ') ให้คงคำที่อยู่ติดกับ [ขีดพิเศษ] ไว้ตามเดิม หรือเอาเซลล์นี้ออกจาก cells แล้วแจ้งผู้ใช้แก้เอง (ไม่สร้างไฟล์)')
    }
    Fail ($prefix + ($plan.Reason -replace ' ให้แก้เองจากข้อเสนอในแชต$', '') + ' หรือเอาเซลล์นี้ออกจาก cells แล้วแจ้งผู้ใช้แก้เอง (ไม่สร้างไฟล์)')
  }
  $before = Get-ParagraphBreakCounts $p
  $nodes = Get-PlainUnitNodes $plan
  foreach ($c in @($p.ChildNodes)) { if ($c.LocalName -ne 'pPr') { [void]$p.RemoveChild($c) } }
  foreach ($n in $nodes) { [void]$p.AppendChild($n) }
  $got = Get-ParagraphBodyText $p
  if (-not [string]::Equals($got, $plan.NewRaw, [StringComparison]::Ordinal)) {
    Fail 'ตรวจไม่ผ่านก่อนบันทึก: ข้อความของย่อหน้าหลังแก้ไม่ตรงกับข้อความที่ผู้ใช้เห็นใน preview จึงไม่สร้างไฟล์'
  }
  $after = Get-ParagraphBreakCounts $p
  if ($after.Breaks -ne $before.Breaks -or $after.Hyphens -ne $before.Hyphens) {
    Fail 'ตรวจไม่ผ่านก่อนบันทึก: การขึ้นบรรทัดใหม่หรือขีดกลางพิเศษในย่อหน้าไม่เท่าเดิมหลังแก้ จึงไม่สร้างไฟล์'
  }
  return @{ Whole = $false; Why = $null; Breaks = $before.Breaks; Hyphens = $before.Hyphens }
}

# Writes $newText into paragraph $p as plain text - no tracked changes: the same word-level change
# apply-revisions.ps1 shows as tracked changes, accepted on the spot. Unchanged text keeps its own run's
# formatting, new text takes the formatting of the run before it (Get-ParagraphDiffPieces), neighbouring pieces
# with the same formatting become one run; the paragraph properties (w:pPr) and bookmarks stay. A paragraph with a
# line break / special hyphen (Test-ParagraphHasUnitContent) goes to Set-ParagraphTextByUnits: the break stays where it
# is, or the change is refused (Fail) when it touches the break - never rewritten as a whole paragraph. Any other
# paragraph that cannot be edited word by word is rewritten in the formatting of its first text run. The caller has
# checked Get-ParagraphEditBlocker; $where = the cell's label for the refusal message. Returns @{ Whole = <bool>;
# Why = <reason when whole>; Breaks; Hyphens (unit path only) }. Fails when the text read back is not the text asked for
# (the file is then never written).
function Set-ParagraphTextPlain($p, [string]$newText, [string]$where = '') {
  $ns = $script:DocxNs
  if (Test-ParagraphHasUnitContent $p) { return (Set-ParagraphTextByUnits $p $newText $where) }
  $plan = Get-ParagraphEditPlan $p $newText
  $bookStarts = @($p.SelectNodes('w:bookmarkStart', $ns))
  $bookEnds = @($p.SelectNodes('w:bookmarkEnd', $ns))
  foreach ($c in @($p.ChildNodes)) { if ($c.LocalName -ne 'pPr') { [void]$p.RemoveChild($c) } }
  foreach ($b in $bookStarts) { [void]$p.AppendChild($b) }
  if ($plan.Mode -eq 'word') {
    $groups = New-Object System.Collections.ArrayList
    foreach ($pc in $plan.Pieces) {
      if ($pc.Kind -eq 'del') { continue }
      $last = if ($groups.Count) { $groups[$groups.Count - 1] } else { $null }
      if ($last -and [object]::ReferenceEquals($last.RPr, $pc.RPr)) { [void]$last.Sb.Append($pc.Text) }
      else {
        $g = @{ RPr = $pc.RPr; Sb = (New-Object System.Text.StringBuilder) }
        [void]$g.Sb.Append($pc.Text)
        [void]$groups.Add($g)
      }
    }
    foreach ($g in $groups) { [void]$p.AppendChild((New-PlainRun $g.RPr $g.Sb.ToString())) }
    $expected = $plan.NewRaw
  } else {
    [void]$p.AppendChild((New-PlainRun $plan.FirstRPr $plan.NewTrim))
    $expected = $plan.NewTrim
  }
  foreach ($b in $bookEnds) { [void]$p.AppendChild($b) }
  $got = Get-ParagraphBodyText $p
  if (-not [string]::Equals($got, $expected, [StringComparison]::Ordinal)) {
    Fail 'ตรวจไม่ผ่านก่อนบันทึก: ข้อความของย่อหน้าหลังแก้ไม่ตรงกับข้อความที่ผู้ใช้เห็นใน preview จึงไม่สร้างไฟล์'
  }
  return @{ Whole = ($plan.Mode -eq 'whole'); Why = $plan.Why }
}

# ---------------------------------------------------------------------------------------------
# Table cells that a rewrite may change (only the cells the confirmed brief lists - review-file.ps1 checks that).
# Coordinates are the ones extract-text.ps1 -Cells prints: rRcC, plus pK when the cell has several paragraphs.
# ---------------------------------------------------------------------------------------------

# Looks a coordinate up in the Get-TableCells records of one table. Returns @{ Rec = <the record>; Error = $null }
# or @{ Rec = $null; Error = <Thai reason> }: a bad coordinate, a row/cell that is not there, the continuation of
# a vertically merged cell, a cell holding a nested table, an empty cell, a paragraph number that is missing or
# wrong, or a paragraph that cannot be edited (hyperlink, field, picture, footnote, comment, tracked change, or
# inside a content control).
function Resolve-TableCellTarget($cells, [string]$coord) {
  $ns = $script:DocxNs
  $cm = [regex]::Match($coord, '^r([1-9]\d*)c([1-9]\d*)(?:p([1-9]\d*))?$')
  if (-not $cm.Success) {
    return @{ Rec = $null; Error = "พิกัดเซลล์ต้องเป็นรูป r4c2 หรือ r4c2p1 (พบ `"$coord`")" }
  }
  $row = [int]$cm.Groups[1].Value; $cell = [int]$cm.Groups[2].Value
  $para = if ($cm.Groups[3].Success) { [int]$cm.Groups[3].Value } else { 0 }
  $rowCells = @($cells | Where-Object { $_.Row -eq $row })
  if (-not $rowCells.Count) {
    $maxRow = if ($cells.Count) { (@($cells | ForEach-Object { $_.Row }) | Measure-Object -Maximum).Maximum } else { 0 }
    return @{ Rec = $null; Error = "ไม่มีแถว $row (ตารางมี $maxRow แถว)" }
  }
  $cellRecs = @($rowCells | Where-Object { $_.Cell -eq $cell })
  if (-not $cellRecs.Count) {
    $maxCell = (@($rowCells | ForEach-Object { $_.Cell }) | Measure-Object -Maximum).Maximum
    return @{ Rec = $null; Error = "แถว $row ไม่มีเซลล์ $cell (มี $maxCell เซลล์)" }
  }
  $rec0 = $cellRecs[0]
  if ($rec0.VCont) { return @{ Rec = $null; Error = 'เป็นส่วนต่อของเซลล์ผสานแนวตั้ง ให้ใช้เซลล์แรกของกลุ่มที่ผสาน' } }
  if ($rec0.Nested) { return @{ Rec = $null; Error = 'มีตารางซ้อนอยู่ในเซลล์ ยังไม่รองรับการเขียนใหม่' } }
  if ($rec0.ParaCount -eq 0) { return @{ Rec = $null; Error = 'เป็นเซลล์ว่าง ไม่มีข้อความให้เขียนใหม่' } }
  if ($rec0.ParaCount -gt 1) {
    if (-not $para) { return @{ Rec = $null; Error = "มี $($rec0.ParaCount) ย่อหน้า ต้องระบุย่อหน้า เช่น ${coord}p1" } }
    $target = $cellRecs | Where-Object { $_.Para -eq $para } | Select-Object -First 1
    if (-not $target) { return @{ Rec = $null; Error = "ไม่มีย่อหน้าที่ $para (เซลล์นี้มี $($rec0.ParaCount) ย่อหน้า)" } }
  } else {
    if ($para) { return @{ Rec = $null; Error = "เซลล์นี้มีย่อหน้าเดียว ใช้พิกัด r${row}c${cell} (ไม่ต้องมี p)" } }
    $target = $rec0
  }
  $blocker = Get-ParagraphEditBlocker $target.Node
  if ($blocker) { return @{ Rec = $null; Error = "$blocker แก้ในทางเขียนใหม่ไม่ได้ ใช้ทาง `"ตรวจและแก้`" หรือแก้เอง" } }
  if ($target.Node.SelectSingleNode('ancestor::w:sdt', $ns)) { return @{ Rec = $null; Error = 'ย่อหน้าอยู่ใน content control แก้ในทางเขียนใหม่ไม่ได้' } }
  return @{ Rec = $target; Error = $null }
}

# Text of one cell (all its paragraphs, joined by a space) from the Get-TableCells records.
function Get-TableCellJoinedText($cells, [int]$row, [int]$cell) {
  $texts = @($cells | Where-Object { $_.Row -eq $row -and $_.Cell -eq $cell -and $_.Text } | ForEach-Object { [string]$_.Text })
  return ($texts -join ' ')
}

# Where a cell sits, to show its meaning next to it: ColumnHead = the text of the first row's cell above it (the
# one covering the same grid column, so merged columns line up), RowHead = the text of the first cell of its row.
# Either is '' when there is none (the cell is in row 1 / is the first cell, or that cell has no text).
function Get-TableCellContext($cells, [int]$row, [int]$cell) {
  $start = @{}
  $next = @{}
  foreach ($rec in $cells) {
    $k = "$($rec.Row)|$($rec.Cell)"
    if ($start.ContainsKey($k)) { continue }
    $pos = if ($next.ContainsKey($rec.Row)) { [int]$next[$rec.Row] } else { 0 }
    $start[$k] = @{ Start = $pos; Span = [int]$rec.Span }
    $next[$rec.Row] = $pos + [int]$rec.Span
  }
  $me = $start["$row|$cell"]
  $colHead = ''; $rowHead = ''
  if ($me -and $row -gt 1) {
    foreach ($rec in $cells) {
      if ($rec.Row -ne 1) { continue }
      $h = $start["1|$($rec.Cell)"]
      if ($me.Start -ge $h.Start -and $me.Start -lt ($h.Start + $h.Span)) { $colHead = Get-TableCellJoinedText $cells 1 $rec.Cell; break }
    }
  }
  if ($cell -gt 1) { $rowHead = Get-TableCellJoinedText $cells $row 1 }
  return @{ ColumnHead = (Get-ShortText $colHead 40); RowHead = (Get-ShortText $rowHead 40) }
}

# =============================================================================================
# Where a point sits, written for a person: the heading path (by the numbers the document itself shows) and a short
# text to paste into Ctrl+F so the reader lands on it in Word. The ¶ number (and rRcC for a table cell) stays as a
# secondary reference only: Word does not show it, and it counts after the table of contents. Shared by
# review-file.ps1 (-Kind review / report) and extract-text.ps1 (-Mode Locate).
#   Get-HeadingPathLabel  path of headings above a ¶ (and the item number written in the body, e.g. 5.2.25)
#   Get-FindCorpus        one string with every paragraph of the document as Word's Ctrl+F sees it (to count matches)
#   Get-FindQuote         the shortest text from a paragraph that is found once in the corpus
#   Get-ItemFinder        path + quote + reference of one review item / one ¶
# =============================================================================================
$script:PathSep = ' ' + [string][char]0x203A + ' '
$script:LQuote = [string][char]0x201C
$script:RQuote = [string][char]0x201D
$script:FindVSep = [string][char]0x0B

# Items of Read-Docx / Read-Txt are numbered 1..Count in order, so ¶n is normally $items[$n - 1].
function Get-ItemByNumber($items, [int]$n) {
  if ($n -ge 1 -and $n -le $items.Count -and $items[$n - 1].N -eq $n) { return $items[$n - 1] }
  foreach ($it in $items) { if ($it.N -eq $n) { return $it } }
  return $null
}

# The short label of one heading: the number the document gives it, else a short piece of its text.
function Get-HeadingLabel($item) {
  $t = [string]$item.Text
  if ($null -eq $item.Node) { $t = ($t -replace '^#{1,6}\s+', '') }
  $t = (Get-OneLine $t).Trim()
  $end = '(?=$|[\s.:)\]\-–—,])'
  $m = [regex]::Match($t, '^(บทที่|ภาคผนวก|ส่วนที่|หมวดที่|หมวด|ข้อ)\s*([0-9๐-๙A-Za-zก-ฮ]+(?:\.[0-9๐-๙]+)*)' + $end)
  if ($m.Success) { return ($m.Groups[1].Value + ' ' + $m.Groups[2].Value) }
  $m = [regex]::Match($t, '^((?i:chapter|section|part|appendix|annex))\s+([0-9]+(?:\.[0-9]+)*|[A-Z])' + $end)
  if ($m.Success) { return ($m.Groups[1].Value + ' ' + $m.Groups[2].Value) }
  $m = [regex]::Match($t, '^([0-9๐-๙]+(?:\.[0-9๐-๙]+)*)\.?(?=\s|$)')
  if ($m.Success) { return $m.Groups[1].Value }
  return (Get-ShortText $t 40)
}

# An item number written at the start of a body paragraph (5.2.25 <title>) that sits between the last heading and ¶$toN:
# the one nearest to ¶$toN. $null when there is none.
function Get-PseudoHeadingLabel($items, [int]$fromN, [int]$toN) {
  for ($k = $toN; $k -gt $fromN; $k--) {
    $it = Get-ItemByNumber $items $k
    if (-not $it -or $it.Kind -ne 'p' -or $it.Level -ne 0) { continue }
    $t = [string]$it.Text
    # plain spaces only after the number: a tab-separated line such as a cover page's "0.1<tab>10 Nov" is not an item number
    if ($t.Length -gt 120 -or $t.Contains("`t")) { continue }
    if (-not [regex]::IsMatch($t, '^\d+(\.\d+)+\.? +\S')) { continue }
    return [regex]::Match($t, '^\d+(?:\.\d+)+').Value
  }
  return $null
}

# The path of ¶$n: "บทที่ 8 › 5.2 › 5.2.25" (headings above it, itself included when it is a heading, then the item number
# written in the body). "ก่อนหัวข้อแรก" before the first heading, "ไม่มีหัวข้อ" in a document without any.
function Get-HeadingPathLabel($items, [int]$n) {
  $stack = New-Object System.Collections.ArrayList
  $anyHeading = $false
  foreach ($it in $items) {
    if ($it.Level -le 0) { continue }
    $anyHeading = $true
    if ($it.N -gt $n) { break }
    while ($stack.Count -gt 0 -and $stack[$stack.Count - 1].Level -ge $it.Level) { $stack.RemoveAt($stack.Count - 1) }
    [void]$stack.Add($it)
  }
  $parts = New-Object System.Collections.ArrayList
  foreach ($h in $stack) { [void]$parts.Add((Get-HeadingLabel $h)) }
  $lastN = if ($stack.Count) { [int]$stack[$stack.Count - 1].N } else { 0 }
  $pseudo = Get-PseudoHeadingLabel $items $lastN $n
  if ($pseudo) { [void]$parts.Add($pseudo) }
  if (-not $stack.Count) {
    if ($anyHeading) { $parts.Insert(0, 'ก่อนหัวข้อแรก') }
    elseif (-not $pseudo) { [void]$parts.Add('ไม่มีหัวข้อ') }
  }
  $out = New-Object System.Collections.ArrayList
  foreach ($p in $parts) {
    $c = [string]$p
    if ($out.Count) {
      $last = [string]$out[$out.Count - 1]
      if ($last -match '^\d+(\.\d+)*$' -and $c -match '^\d+(\.\d+)*$' -and $c.StartsWith($last + '.', [StringComparison]::Ordinal)) { $out.RemoveAt($out.Count - 1) }
    }
    [void]$out.Add($c)
  }
  return ($out -join $script:PathSep)
}

# Every paragraph of the document as Word's Ctrl+F finds text, one string: the paragraphs of the body (table of
# contents, table cells and text boxes included; each paragraph holds only its own text), tracked changes accepted
# (no w:del / w:moveFrom / mc:Fallback), a tab as "`t", a line break and a special hyphen as the separator
# "`v" (never part of a quote), a paragraph end as "`n". .txt: the text of every item, each followed by "`n".
# Built once per document; the matches of a quote are counted with IndexOf on this one string.
function Get-FindCorpus($items) {
  $all = New-Object System.Text.StringBuilder
  if (-not $items.Count -or $null -eq $items[0].Node) {
    foreach ($it in $items) { [void]$all.Append([string]$it.Text); [void]$all.Append("`n") }
    return $all.ToString()
  }
  $skip = 'ancestor::w:del or ancestor::w:moveFrom or ancestor::mc:Fallback'
  $x = "//w:body//w:t[not($skip)] | //w:body//w:tab[not(parent::w:tabs) and not($skip)] | //w:body//w:br[not($skip)]" +
       " | //w:body//w:cr[not($skip)] | //w:body//w:noBreakHyphen[not($skip)] | //w:body//w:softHyphen[not($skip)]"
  $map = New-Object 'System.Collections.Generic.Dictionary[System.Xml.XmlNode,System.Text.StringBuilder]'
  $order = New-Object 'System.Collections.Generic.List[System.Xml.XmlNode]'
  foreach ($nd in $script:DocxXml.SelectNodes($x, $script:DocxNs)) {
    $p = $nd.ParentNode
    while ($null -ne $p -and -not ($p.LocalName -eq 'p' -and $p.NamespaceURI -eq $W_NS)) { $p = $p.ParentNode }
    if ($null -eq $p) { continue }
    $sb = $null
    if (-not $map.TryGetValue($p, [ref]$sb)) { $sb = New-Object System.Text.StringBuilder; $map[$p] = $sb; $order.Add($p) }
    $ln = $nd.LocalName
    if ($ln -eq 't') { [void]$sb.Append($nd.InnerText) }
    elseif ($ln -eq 'tab') { [void]$sb.Append("`t") }
    else { [void]$sb.Append($script:FindVSep) }
  }
  foreach ($p in $order) { [void]$all.Append($map[$p].ToString()); [void]$all.Append("`n") }
  return $all.ToString()
}

# A character a person cannot type into Ctrl+F as it looks: any white space but the plain space (tab, no-break space,
# line break ...) and the invisible ones. -Light (.txt): only the tab and the line break.
function Test-FindBadChar([char]$c, [bool]$light) {
  if ($c -eq ' ') { return $false }
  if ($light) { return ($c -eq "`t" -or $c -eq "`n" -or $c -eq "`r") }
  $code = [int]$c
  if ([char]::IsWhiteSpace($c) -or [char]::IsControl($c)) { return $true }
  return ($code -eq 0xAD -or ($code -ge 0x200B -and $code -le 0x200F) -or $code -eq 0x2060 -or $code -eq 0xFEFF)
}

# The text of a paragraph / cell paragraph as extract-text.ps1 shows it (Get-NodeText, trimmed) with the places Ctrl+F
# cannot cross: Bad[i] = char i cannot be typed (tab, the space that stands for a line break, no-break space ...);
# Seam = positions where Word has a character that the extracted text does not (a no-break / soft hyphen, a w:cr).
#   -Node  a .docx paragraph        -Text  plain text (.txt item, or the joined text of a merged / nested cell)
function New-FindTarget($Node, [string]$Text, [bool]$Light = $false, [bool]$TxtHeading = $false) {
  $sb = New-Object System.Text.StringBuilder
  $bad = New-Object 'System.Collections.Generic.List[bool]'
  $seam = New-Object 'System.Collections.Generic.List[int]'
  if ($null -ne $Node) {
    $skip = 'ancestor::w:txbxContent or ancestor::w:del or ancestor::w:moveFrom or ancestor::mc:Fallback'
    $x = ".//w:t[not($skip)] | .//w:tab[not(parent::w:tabs) and not($skip)] | .//w:br[not($skip)]" +
         " | .//w:cr[not($skip)] | .//w:noBreakHyphen[not($skip)] | .//w:softHyphen[not($skip)]"
    foreach ($nd in $Node.SelectNodes($x, $script:DocxNs)) {
      $ln = $nd.LocalName
      if ($ln -eq 't') {
        $s = $nd.InnerText
        foreach ($ch in $s.ToCharArray()) { $bad.Add((Test-FindBadChar $ch $false)) }
        [void]$sb.Append($s)
      } elseif ($ln -eq 'tab') { [void]$sb.Append("`t"); $bad.Add($true) }
      elseif ($ln -eq 'br') { [void]$sb.Append(' '); $bad.Add($true) }
      else { $seam.Add($sb.Length) }
    }
  } else {
    foreach ($ch in $Text.ToCharArray()) { $bad.Add((Test-FindBadChar $ch $Light)) }
    [void]$sb.Append($Text)
    if ($TxtHeading) {
      $hm = [regex]::Match($Text, '^#{1,6}\s+')
      if ($hm.Success) { for ($i = 0; $i -lt $hm.Length; $i++) { $bad[$i] = $true } }
    }
  }
  $raw = $sb.ToString()
  $lead = $raw.Length - $raw.TrimStart().Length
  $trimmed = $raw.Trim()
  $badArr = New-Object 'bool[]' $trimmed.Length
  for ($i = 0; $i -lt $trimmed.Length; $i++) { $badArr[$i] = $bad[$i + $lead] }
  $seamSet = New-Object 'System.Collections.Generic.HashSet[int]'
  foreach ($p in $seam) { $q = $p - $lead; if ($q -gt 0 -and $q -lt $trimmed.Length) { [void]$seamSet.Add($q) } }
  return @{ Text = $trimmed; Bad = $badArr; Seam = $seamSet }
}

# The stretches of a target that can be typed into Ctrl+F: @(start, end) pairs (end exclusive), without leading / trailing spaces.
function Get-FindSegments($target) {
  $text = $target.Text
  $len = $text.Length
  $segs = New-Object System.Collections.ArrayList
  $start = -1
  for ($i = 0; $i -le $len; $i++) {
    $cut = ($i -eq $len) -or $target.Bad[$i]
    $atSeam = ($i -lt $len) -and $target.Seam.Contains($i)
    if ($start -ge 0 -and ($cut -or $atSeam)) {
      $a = $start; $b = $i
      while ($a -lt $b -and $text[$a] -eq ' ') { $a++ }
      while ($b -gt $a -and $text[$b - 1] -eq ' ') { $b-- }
      if ($b -gt $a) { [void]$segs.Add(@($a, $b)) }
      $start = -1
    }
    if ($start -lt 0 -and $i -lt $len -and -not $target.Bad[$i]) { $start = $i }
  }
  return , $segs
}

# A mark written on the character before it (Thai vowels above / below, tone marks ...): a quote never starts or ends next to one.
function Test-FindCombiningChar([char]$c) {
  # ำ (sara am) is a letter to .NET but Word keeps it with its consonant (ข้อก does not find ข้อกำหนด; checked in Word)
  if ([int]$c -eq 0x0E33) { return $true }
  $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
  return ($cat -eq [System.Globalization.UnicodeCategory]::NonSpacingMark -or $cat -eq [System.Globalization.UnicodeCategory]::SpacingCombiningMark -or $cat -eq [System.Globalization.UnicodeCategory]::EnclosingMark)
}

# An ASCII letter or digit (a character inside an English word or a number).
function Test-FindAsciiWordChar([char]$c) { return ([int]$c -lt 128 -and [char]::IsLetterOrDigit($c)) }

# How many times $needle is in $corpus (ignoring case, as Ctrl+F does by default). Stops counting at 999.
function Get-FindCount([string]$corpus, [string]$needle) {
  if (-not $needle) { return 0 }
  $n = 0; $i = 0
  while ($n -lt 999) {
    $i = $corpus.IndexOf($needle, $i, [StringComparison]::OrdinalIgnoreCase)
    if ($i -lt 0) { break }
    $n++; $i++
  }
  return $n
}

# $text with every run of white space as one space, and for each char of the result the index it has in $text.
function Get-CollapsedText([string]$text) {
  $sb = New-Object System.Text.StringBuilder
  $map = New-Object 'System.Collections.Generic.List[int]'
  $prevSpace = $false
  for ($i = 0; $i -lt $text.Length; $i++) {
    if ([char]::IsWhiteSpace($text[$i])) {
      if (-not $prevSpace) { [void]$sb.Append(' '); $map.Add($i) }
      $prevSpace = $true
    } else { [void]$sb.Append($text[$i]); $map.Add($i); $prevSpace = $false }
  }
  return @{ Text = $sb.ToString(); Map = $map }
}

# Does $text contain $needle, the second try with every run of white space as one space.
function Test-FindTextHas([string]$text, [string]$needle) {
  if (-not $needle) { return $false }
  if ($text.IndexOf($needle, [StringComparison]::Ordinal) -ge 0) { return $true }
  return ((Get-CollapsedText $text).Text.IndexOf(($needle -replace '\s+', ' '), [StringComparison]::Ordinal) -ge 0)
}

# The text to put in Ctrl+F for one paragraph / cell paragraph: @{ Quote; Count; BeforeFound }.
#   1. "before" itself (8-60 chars, in one typable stretch) when it is found once in the document
#   2. else from the start of "before" (or of the paragraph when "before" is empty / not found) 20, 30, 45, 60, 80 chars
#      (ending at the next space when one is within 10 chars, else at a text-element boundary so no Thai vowel / tone
#      mark is cut off), the first length found once
#   3. else the longest one, Count = how many times it is found
# Quote = '' when the paragraph has nothing typable.
function Get-FindQuote([string]$corpus, $target, [string]$before) {
  $text = $target.Text
  $res = @{ Quote = ''; Count = 0; BeforeFound = $true }
  $segs = Get-FindSegments $target
  $b = if ($before) { $before.Trim() } else { '' }
  $pos = -1; $bEnd = -1
  if ($b) {
    $pos = $text.IndexOf($b, [StringComparison]::Ordinal)
    if ($pos -ge 0) { $bEnd = $pos + $b.Length }
    else {
      $cm = Get-CollapsedText $text
      $bc = ($b -replace '\s+', ' ')
      $ci = $cm.Text.IndexOf($bc, [StringComparison]::Ordinal)
      if ($ci -ge 0) { $pos = $cm.Map[$ci]; $bEnd = $cm.Map[$ci + $bc.Length - 1] + 1 }
    }
    if ($pos -lt 0) { $res.BeforeFound = $false }
    else {
      # Word's Ctrl+F does not find a Thai consonant without the vowel / tone mark written on it (ส่วนที does not find
      # ส่วนที่, บัญช does not find บัญชี; checked in Word): widen a "before" that starts or ends inside such a cluster
      while ($pos -gt 0 -and (Test-FindCombiningChar $text[$pos])) { $pos-- }
      while ($bEnd -lt $text.Length -and (Test-FindCombiningChar $text[$bEnd])) { $bEnd++ }
    }
  }
  if (-not $segs.Count) { return $res }
  # 1. before itself
  if ($pos -ge 0) {
    $len = $bEnd - $pos
    if ($len -ge 8 -and $len -le 60) {
      foreach ($sg in $segs) {
        if ($sg[0] -le $pos -and $bEnd -le $sg[1]) {
          $cand = $text.Substring($pos, $len)
          $cnt = Get-FindCount $corpus $cand
          if ($cnt -le 1) { $res.Quote = $cand; $res.Count = [Math]::Max($cnt, 1); return $res }
          break
        }
      }
    }
  }
  # 2. from the start of before
  $s = if ($pos -ge 0) { $pos } else { $segs[0][0] }
  $seg = $null
  foreach ($sg in $segs) { if ($sg[1] -gt $s) { $seg = $sg; break } }
  if ($null -eq $seg) { $seg = $segs[$segs.Count - 1]; $s = $seg[0] }
  $segA = [int]$seg[0]; $segB = [int]$seg[1]
  $start0 = [Math]::Max($s, $segA)
  $bounds = [System.Globalization.StringInfo]::ParseCombiningCharacters($text)
  $lastQuote = ''; $lastCount = 0
  foreach ($L in @(20, 30, 45, 60, 80)) {
    $st = $start0
    if ($segB - $st -lt $L) {
      $st = [Math]::Max($segA, $segB - $L)
      $sb2 = $segA
      foreach ($bd in $bounds) { if ($bd -le $st -and $bd -ge $segA) { $sb2 = $bd } }
      $st = $sb2
    }
    $end = $st + $L
    if ($end -ge $segB) { $end = $segB }
    else {
      $j = -1
      for ($k = $end; $k -le [Math]::Min($end + 10, $segB - 1); $k++) { if ($text[$k] -eq ' ') { $j = $k; break } }
      if ($j -ge 0) { $end = $j }
      else {
        $nb = $segB
        foreach ($bd in $bounds) { if ($bd -ge $end) { $nb = [Math]::Min($bd, $segB); break } }
        $end = $nb
      }
      # a leading vowel (เ แ โ ไ ใ) is not left at the end of the quote
      while ($end -lt $segB -and $end -gt $st -and [int]$text[$end - 1] -ge 0x0E40 -and [int]$text[$end - 1] -le 0x0E44) {
        $nb = $segB
        foreach ($bd in $bounds) { if ($bd -gt $end) { $nb = [Math]::Min($bd, $segB); break } }
        $end = $nb
      }
      # an English word / number is not cut in two at the end (go-l -> go-live): extend to its end, up to 15 chars
      $w = $end
      while ($w -lt $segB -and ($w - $end) -lt 15 -and (Test-FindAsciiWordChar $text[$w - 1]) -and (Test-FindAsciiWordChar $text[$w])) { $w++ }
      if ($w -eq $segB -or -not (Test-FindAsciiWordChar $text[$w])) { $end = $w }
    }
    # never end before / start on a mark of the cluster (Test-FindCombiningChar: Word does not find a cut cluster)
    while ($end -lt $segB -and (Test-FindCombiningChar $text[$end])) { $end++ }
    while ($st -gt $segA -and (Test-FindCombiningChar $text[$st])) { $st-- }
    # nor at the start, when the quote starts inside one (the start of "before" is kept as the model gave it)
    if ($st -ne $pos) {
      $w = $st
      while ($w -gt $segA -and ($st - $w) -lt 15 -and (Test-FindAsciiWordChar $text[$w - 1]) -and (Test-FindAsciiWordChar $text[$w])) { $w-- }
      if ($w -eq $segA -or -not (Test-FindAsciiWordChar $text[$w - 1])) { $st = $w }
    }
    $cand = $text.Substring($st, $end - $st).Trim(' ')
    if (-not $cand) { continue }
    $cnt = Get-FindCount $corpus $cand
    $lastQuote = $cand; $lastCount = $cnt
    if ($cnt -le 1) { break }
  }
  $res.Quote = $lastQuote
  $res.Count = [Math]::Max($lastCount, 1)
  return $res
}

# The table label for a table at ¶$n: the caption next to it ("ตารางที่ 3"), else just "ตาราง". The item above is looked at first
# (captions are usually above a table), then the one below.
function Get-TableLabel($items, [int]$n) {
  foreach ($k in @(($n - 1), ($n + 1))) {
    $c = Get-ItemByNumber $items $k
    if (-not $c -or $c.Kind -ne 'p' -or $c.Level -gt 0) { continue }
    $t = (Get-OneLine ([string]$c.Text)).Trim()
    if (-not $t) { continue }
    $m = [regex]::Match($t, '^(ตารางที่|ตาราง|(?i:table))\s*([0-9๐-๙]+(?:[.\-][0-9๐-๙]+)*)')
    if ($m.Success -and ($c.Caption -or $t.Length -le 120)) {
      return ('ตาราง ' + $script:LQuote + $m.Groups[1].Value + ' ' + $m.Groups[2].Value + $script:RQuote)
    }
    if ($c.Caption -and $t -notmatch '^(รูป|ภาพ|(?i:figure|fig\b))') { return ('ตาราง ' + $script:LQuote + (Get-ShortText $t 40) + $script:RQuote) }
  }
  return 'ตาราง'
}

# Path + Ctrl+F text + reference of one point. $paragraph = ¶ number; $row / $cell / $cellParagraph only for a table cell
# (the coordinates extract-text.ps1 -Cells prints); $before = the text the item quotes ('' = the start of the paragraph).
# Returns @{ Path; Quote; Count; BeforeFound; Ref; Error }. Error is set (and the rest is empty) only for a ¶ / row / cell /
# cell paragraph that is not in the file; a cell that cannot be edited (vertical-merge continuation, nested table, empty)
# is still pointed at.
function Get-ItemFinder($items, [string]$corpus, $paragraph, $row, $cell, $cellParagraph, [string]$before) {
  $res = @{ Path = ''; Quote = ''; Count = 0; BeforeFound = $true; Ref = ''; Error = $null }
  $n = [int]$paragraph
  $it = Get-ItemByNumber $items $n
  if (-not $it) { $res.Error = "ไม่พบ ¶$n ในไฟล์ต้นฉบับ"; return $res }
  $isDocx = ($null -ne $it.Node)
  $headPath = Get-HeadingPathLabel $items $n
  $hasCoord = ($null -ne $row -and $null -ne $cell)

  if ($it.Kind -ne 'table') {
    if ($hasCoord) { $res.Error = "¶$n ไม่ใช่ตาราง (ใช้ row/cell ไม่ได้)"; return $res }
    $target = if ($isDocx) { New-FindTarget -Node $it.Node } else { New-FindTarget -Text ([string]$it.Text) -Light $true -TxtHeading ($it.Level -gt 0) }
    $q = Get-FindQuote $corpus $target $before
    $res.Path = $headPath; $res.Quote = $q.Quote; $res.Count = $q.Count; $res.BeforeFound = $q.BeforeFound; $res.Ref = "¶$n"
    return $res
  }

  $cells = Get-TableCells $it.Node
  $tblLabel = Get-TableLabel $items $n
  $rec = $null
  $rowNo = 0; $cellNo = 0
  $b = if ($before) { $before.Trim() } else { '' }
  if ($hasCoord) {
    $rowNo = [int]$row; $cellNo = [int]$cell
    $rowCells = @($cells | Where-Object { $_.Row -eq $rowNo })
    if (-not $rowCells.Count) {
      $maxRow = if ($cells.Count) { (@($cells | ForEach-Object { $_.Row }) | Measure-Object -Maximum).Maximum } else { 0 }
      $res.Error = "ไม่มีแถว $rowNo (ตารางมี $maxRow แถว)"; return $res
    }
    $cellRecs = @($rowCells | Where-Object { $_.Cell -eq $cellNo })
    if (-not $cellRecs.Count) {
      $maxCell = (@($rowCells | ForEach-Object { $_.Cell }) | Measure-Object -Maximum).Maximum
      $res.Error = "แถว $rowNo ไม่มีเซลล์ $cellNo (มี $maxCell เซลล์)"; return $res
    }
    $rec = $cellRecs[0]
    if ($cellRecs.Count -gt 1) {
      if ($null -ne $cellParagraph) {
        $rec = $cellRecs | Where-Object { $_.Para -eq [int]$cellParagraph } | Select-Object -First 1
        if (-not $rec) { $res.Error = "ไม่มีย่อหน้าที่ $([int]$cellParagraph) (เซลล์นี้มี $($cellRecs[0].ParaCount) ย่อหน้า)"; return $res }
      } else {
        $hit = $cellRecs | Where-Object { $b -and (Test-FindTextHas ([string]$_.Text) $b) } | Select-Object -First 1
        if ($hit) { $rec = $hit }
      }
    }
  } else {
    if ($b) { foreach ($c in $cells) { if ($c.Text -and (Test-FindTextHas ([string]$c.Text) $b)) { $rec = $c; break } } }
    if ($rec) { $rowNo = [int]$rec.Row; $cellNo = [int]$rec.Cell }
    else {
      if ($b) { $res.BeforeFound = $false }
      foreach ($c in $cells) { if ($c.Text) { $rec = $c; break } }
    }
  }

  $res.Path = $headPath + $script:PathSep + $tblLabel
  $res.Ref = "¶$n"
  if ($rowNo -gt 0) {
    $ctx = Get-TableCellContext $cells $rowNo $cellNo
    $lab = if ($cellNo -gt 1 -and $ctx.RowHead) { 'แถว ' + $script:LQuote + $ctx.RowHead + $script:RQuote } else { "แถว $rowNo" }
    if ($rowNo -gt 1 -and $ctx.ColumnHead) { $lab += ' · คอลัมน์ ' + $script:LQuote + $ctx.ColumnHead + $script:RQuote }
    $res.Path += $script:PathSep + $lab
    $res.Ref += " r${rowNo}c${cellNo}"
    if ($rec -and $rec.Para) { $res.Ref += "p$($rec.Para)" }
  }
  if (-not $rec) { return $res }
  $target = if ($rec.Node) { New-FindTarget -Node $rec.Node } else { New-FindTarget -Text ([string]$rec.Text) -Light $false }
  $q = Get-FindQuote $corpus $target $before
  $res.Quote = $q.Quote; $res.Count = $q.Count
  if ($hasCoord) { $res.BeforeFound = $q.BeforeFound }
  return $res
}

# A Ctrl+F text as a Markdown code span (a backtick inside it: a longer fence, with a space on each side, as CommonMark has it).
function Format-CodeSpan([string]$s) {
  $max = 0
  foreach ($m in [regex]::Matches($s, '`+')) { if ($m.Length -gt $max) { $max = $m.Length } }
  if ($max -eq 0) { return ('`' + $s + '`') }
  $f = '`' * ($max + 1)
  return ($f + ' ' + $s + ' ' + $f)
}

# =============================================================================================
# Rewrite of whole sections (write-rewrite.ps1 builds the file; review-file.ps1 -Kind rewrite runs
# the very same build as a dry run on an in-memory copy, so a preview the user is shown is always
# one that write-rewrite.ps1 can build). Everything below used to live in write-rewrite.ps1.
# =============================================================================================

# G1 (docs/v2-plan.md 3.9): every "heading"/"paragraph" block must carry "from" - an array of one
# or more original ¶ numbers this block is REWRITTEN FROM. Only format is validated here (each
# entry a plain integer >= 1); whether those ¶ actually exist in the section, are not anchors, and
# line up with the block's position relative to any anchor "keep"s in the plan is checked later,
# per section, by Get-FromSegment (needs the section's own item list first).
function Parse-FromIds($raw, [string]$label) {
  $arr = Get-AsArray $raw
  if (-not $arr.Count) { Fail "${label}: ต้องมี from (¶ ที่มาอย่างน้อยหนึ่งตัว)" }
  $ids = New-Object System.Collections.ArrayList
  foreach ($f in $arr) {
    $n = $f -as [int]
    if (-not $n -or $n -lt 1) { Fail "${label}: from ต้องเป็นเลข ¶ ตั้งแต่ 1 ทุกตัว" }
    [void]$ids.Add($n)
  }
  return , @($ids)
}

# A "table" block: the table ¶ stays where it is (it counts as a keep of that anchor) and only the listed cells get
# new text - Cells = @( @{ Cell = 'r4c2' (or 'r4c2p1'); Text } ). One coordinate = one paragraph of a cell, so the text
# is one paragraph (no line break) and not empty; a coordinate is listed once. Whether the cell exists and can be
# changed is checked against the document by Build-DocxSection (Resolve-TableCellTarget).
function New-TableBlock([int]$paragraph, $cellsRaw, [string]$label) {
  $arr = Get-AsArray $cellsRaw
  $cells = New-Object System.Collections.ArrayList
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($c in $arr) {
    $coord = [string]$c.cell
    if ($coord -cnotmatch '^r[1-9]\d*c[1-9]\d*(?:p[1-9]\d*)?$') { Fail "${label}: cell ต้องเป็นพิกัดรูป r4c2 หรือ r4c2p1 (พบ `"$coord`")" }
    if (-not $seen.Add($coord)) { Fail "${label}: cell $coord ซ้ำ" }
    $text = [string]$c.text
    if (-not $text.Trim()) { Fail "${label}: ข้อความใหม่ของ $coord ว่าง (เซลล์ที่ไม่เปลี่ยนไม่ต้องใส่ใน cells)" }
    if ($text -match '[\r\n]') { Fail "${label}: ข้อความใหม่ของ $coord ต้องเป็นย่อหน้าเดียว ห้ามขึ้นบรรทัดใหม่หรือขึ้นย่อหน้าใหม่ (หนึ่งพิกัด = หนึ่งย่อหน้าในเซลล์)" }
    [void]$cells.Add(@{ Cell = $coord; Text = $text })
  }
  if (-not $cells.Count) { Fail "${label}: table ต้องมี cells อย่างน้อยหนึ่งเซลล์" }
  return @{ Type = 'table'; Paragraph = $paragraph; Cells = @($cells) }
}

# Parses the "blocks" array of one plan section (write-rewrite.ps1 -Plan, or the plan data that
# review-file.ps1 embeds in a rewrite preview - the same shape) into block records:
#   @{ Type='keep'; Paragraph } | @{ Type='heading'; Level; Text; From } | @{ Type='paragraph'; Text; From }
#   | @{ Type='table'; Paragraph; Cells } (a kept table with new text in some cells, see New-TableBlock)
function ConvertTo-PlanBlocks($blocksRaw, [string]$sid) {
  $blocks = New-Object System.Collections.ArrayList
  $bi = 0
  foreach ($b in $blocksRaw) {
    $bi++
    $type = [string]$b.type
    $label = "${sid} blocks[$bi]"
    switch ($type) {
      'keep' {
        $paraNum = $b.paragraph -as [int]
        if (-not $paraNum -or $paraNum -lt 1) { Fail "${label}: keep ต้องมี paragraph เป็นเลขตั้งแต่ 1" }
        [void]$blocks.Add(@{ Type = 'keep'; Paragraph = $paraNum })
      }
      'heading' {
        $level = $b.level -as [int]
        if (-not $level -or $level -lt 1) { Fail "${label}: heading ต้องมี level เป็นเลขตั้งแต่ 1" }
        $text = [string]$b.text
        if (-not $text.Trim()) { Fail "${label}: heading ต้องมี text" }
        $from = Parse-FromIds $b.from $label
        [void]$blocks.Add(@{ Type = 'heading'; Level = $level; Text = $text; From = $from })
      }
      'paragraph' {
        $text = [string]$b.text
        if (-not $text.Trim()) { Fail "${label}: paragraph ต้องมี text" }
        $from = Parse-FromIds $b.from $label
        [void]$blocks.Add(@{ Type = 'paragraph'; Text = $text; From = $from })
      }
      'table' {
        $paraNum = $b.paragraph -as [int]
        if (-not $paraNum -or $paraNum -lt 1) { Fail "${label}: table ต้องมี paragraph เป็นเลขตั้งแต่ 1" }
        [void]$blocks.Add((New-TableBlock $paraNum $b.cells $label))
      }
      default { Fail "${label}: type ต้องเป็น keep, heading, paragraph หรือ table" }
    }
  }
  return , $blocks
}

# Resolves "S<n>" against the outline (same numbering as extract-text.ps1 -Mode Outline: every item
# with Level > 0, in document order, 1-based). Returns @{ Head = <item>; End = <int N> }.
function Resolve-Section($items, [string]$sid) {
  if ($sid -notmatch '^S(\d+)$') { Fail "section ไม่ถูกต้อง: $sid" }
  $idx = [int]$Matches[1]
  $heads = @($items | Where-Object { $_.Level -gt 0 })
  if ($idx -lt 1 -or $idx -gt $heads.Count) { Fail "ไม่มีหัวข้อ $sid (มี $($heads.Count) หัวข้อ)" }
  $head = $heads[$idx - 1]
  $end = Get-SectionEnd $items $head
  return @{ Head = $head; End = $end }
}

# ---------------------------------------------------------------------------------------------
# Section profile helpers shared by review-file.ps1 (-Kind brief and -Kind rewrite). Adding them here keeps
# ONE definition of "a section with no text to rewrite" and of "a section too big for one preview".
# ---------------------------------------------------------------------------------------------

# A section whose range is longer than this many ¶ (the ¶a-b span shown in the outline) is "large": when it
# can be split (Get-SplitSubsectionIds) the rewrite of the whole file makes one preview per sub-heading.
# This is the only place the number is defined; references/full-rewrite.md quotes it.
$script:LargeSectionParagraphs = 60

# True for an ordinary paragraph that a rewrite can change: not a heading, not a table, and (.docx) not an
# anchor (picture, link, field, footnote, tracked change, content control ...). A caption counts (it is
# plain text). $ns is the docx namespace manager, or $null for a .txt (which has no anchors).
function Test-RewritableTextItem($item, $ns) {
  if ($item.Kind -ne 'p' -or $item.Level -gt 0) { return $false }
  if ($ns -and (Test-AnchorItem $item $ns)) { return $false }
  return $true
}

# Facts about section $sid (resolved like Resolve-Section): Count = number of ¶ in its range, TableNs = the ¶
# numbers of the tables in it (top-level tables that carry text; sub-sections included), HasText = whether
# the range holds at least one paragraph a rewrite could change.
function Get-SectionProfile($items, $ns, [string]$sid) {
  $r = Resolve-Section $items $sid
  $inRange = @($items | Where-Object { $_.N -ge $r.Head.N -and $_.N -le $r.End })
  $tableNs = @($inRange | Where-Object { $_.Kind -eq 'table' } | ForEach-Object { $_.N })
  $hasText = [bool](@($inRange | Where-Object { Test-RewritableTextItem $_ $ns }).Count)
  return @{ Head = $r.Head; End = $r.End; Count = $inRange.Count; TableNs = $tableNs; HasText = $hasText }
}

# The ids (S<n>) of the sub-headings a large section should be split into, or an empty array when it should
# not be split: the section is not longer than $script:LargeSectionParagraphs, has no sub-heading, or has
# paragraphs of its own before its first sub-heading (splitting would leave those unrewritten).
# One level only: the sub-headings meant are the direct children (a heading with a deeper level than the
# section whose level is not deeper than a heading already seen after the section heading).
# Only a top-level section is ever advised to be split: its heading has the shallowest level of any heading
# in the document. A section that is itself a sub-heading (already the result of a split) returns an empty
# array, so the note never asks for a second layer of splitting.
function Get-SplitSubsectionIds($items, $ns, [string]$sid) {
  $r = Resolve-Section $items $sid
  if (($r.End - $r.Head.N + 1) -le $script:LargeSectionParagraphs) { return @() }
  $heads = @($items | Where-Object { $_.Level -gt 0 })
  $topLevel = [int]::MaxValue
  foreach ($hd in $heads) { if ($hd.Level -lt $topLevel) { $topLevel = $hd.Level } }
  if ($r.Head.Level -ne $topLevel) { return @() }
  $children = New-Object System.Collections.ArrayList
  $minLevel = [int]::MaxValue
  for ($hi = 0; $hi -lt $heads.Count; $hi++) {
    $h = $heads[$hi]
    if ($h.N -le $r.Head.N -or $h.N -gt $r.End) { continue }
    if ($h.Level -le $minLevel) { [void]$children.Add(@{ Id = ('S' + ($hi + 1)); N = $h.N }) }
    if ($h.Level -lt $minLevel) { $minLevel = $h.Level }
  }
  if (-not $children.Count) { return @() }
  $firstChildN = $children[0].N
  $own = @($items | Where-Object { $_.N -gt $r.Head.N -and $_.N -lt $firstChildN -and (Test-RewritableTextItem $_ $ns) })
  if ($own.Count) { return @() }
  return @($children | ForEach-Object { $_.Id })
}

# The section units of a rewrite of the WHOLE file (references/full-rewrite.md, heading "ทั้งไฟล์"): every
# top-level heading (one that sits under no other heading) is one unit with its sub-headings, except a large top-level
# section that can be split (Get-SplitSubsectionIds), whose sub-headings are the units instead (its own heading stays
# as it is and belongs to no unit). Returns @{ Id; Start; End; HasText } in document order (Start/End = the section's
# ¶ range as the outline shows it; HasText = a paragraph a rewrite could change is in it). A unit with no such
# paragraph is skipped in a whole-file rewrite (no preview, kept as it is). One definition, used by review-file.ps1.
function Get-WholeFileUnits($items, $ns) {
  $heads = @($items | Where-Object { $_.Level -gt 0 })
  $units = New-Object System.Collections.ArrayList
  $coveredUntil = 0
  for ($hi = 0; $hi -lt $heads.Count; $hi++) {
    $h = $heads[$hi]
    if ($h.N -le $coveredUntil) { continue }
    $topId = 'S' + ($hi + 1)
    $coveredUntil = Get-SectionEnd $items $h
    $splitIds = @(Get-SplitSubsectionIds $items $ns $topId)
    $unitIds = if ($splitIds.Count) { $splitIds } else { @($topId) }
    foreach ($uid in $unitIds) {
      $prof = Get-SectionProfile $items $ns $uid
      [void]$units.Add(@{ Id = $uid; Start = $prof.Head.N; End = $prof.End; HasText = $prof.HasText })
    }
  }
  return @($units)
}
function New-W([string]$name) { return $script:DocxXml.CreateElement('w', $name, $W_NS) }

function New-PlainRun($rPr, [string]$text) {
  $r = New-W 'r'
  if ($rPr) { [void]$r.AppendChild($rPr.CloneNode($true)) }
  $parts = $text -split "(`t)"
  foreach ($part in $parts) {
    if ($part -eq '') { continue }
    if ($part -eq "`t") { [void]$r.AppendChild((New-W 'tab')); continue }
    $t = New-W 't'
    [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
    $t.InnerText = $part
    [void]$r.AppendChild($t)
  }
  return $r
}

function New-PlainParagraph($pPr, $rPr, [string]$text) {
  $p = New-W 'p'
  if ($pPr) { [void]$p.AppendChild($pPr.CloneNode($true)) }
  [void]$p.AppendChild((New-PlainRun $rPr $text))
  return $p
}

# Finds ANY item anywhere in the whole document with the given outline level, to reuse its heading
# style (pPr) for a brand-new "heading" block. Returns the item or $null.
function Find-HeadingStyleItem($items, [int]$level) {
  return ($items | Where-Object { $_.Level -eq $level -and $_.Kind -eq 'p' } | Select-Object -First 1)
}

# Validates the anchor/keep contract for one section and returns the ordered list of "keep"
# paragraph numbers actually referenced (for the summary line). Fails with a clear message and the
# offending ¶ number(s) on any anchor mismatch. $itemsInSection is the ordered original items for
# [head.N .. end.N] INCLUDING the heading item; $ns is $null for .txt (no anchors possible there).
function Test-AnchorContract([string]$sid, $itemsInSection, $headN, $blocks, $ns) {
  $anchorNs = New-Object System.Collections.ArrayList
  if ($ns) {
    foreach ($it in $itemsInSection) {
      if ($it.N -eq $headN) { continue }
      if (Test-AnchorItem $it $ns) { [void]$anchorNs.Add($it.N) }
    }
  }
  $keepNs = New-Object System.Collections.ArrayList
  $seenKeep = New-Object 'System.Collections.Generic.HashSet[int]'
  foreach ($b in $blocks) {
    if ($b.Type -ne 'keep') { continue }
    if ($seenKeep.Contains($b.Paragraph)) { Fail "${sid}: keep ¶$($b.Paragraph) ซ้ำในแผน" }
    [void]$seenKeep.Add($b.Paragraph)
    $inSection = [bool]($itemsInSection | Where-Object { $_.N -eq $b.Paragraph })
    if (-not $inSection) { Fail "${sid}: keep ¶$($b.Paragraph) อยู่นอก section นี้" }
    [void]$keepNs.Add($b.Paragraph)
  }
  # keepAnchorNs: the subset of $keepNs that are anchors, in the order the plan listed them.
  $anchorSet = New-Object 'System.Collections.Generic.HashSet[int]'
  foreach ($n in $anchorNs) { [void]$anchorSet.Add($n) }
  $keepAnchorNs = @($keepNs | Where-Object { $anchorSet.Contains($_) })
  $expected = @($anchorNs)
  if ($keepAnchorNs.Count -ne $expected.Count) {
    $missing = @($expected | Where-Object { $keepAnchorNs -notcontains $_ })
    Fail ("${sid}: ต้องมี keep ครบทุกตาราง/ย่อหน้าที่มีรูป-วัตถุฝัง (OLE)-ลิงก์-ฟิลด์-เชิงอรรถ-comment-tracked changes-content control ในช่วงนี้ ขาด ¶" + ($missing -join ', ¶'))
  }
  for ($i = 0; $i -lt $expected.Count; $i++) {
    if ($keepAnchorNs[$i] -ne $expected[$i]) {
      Fail ("${sid}: ลำดับ keep ของ ¶ ที่ต้องคงไว้ไม่ตรงกับต้นฉบับ (ต้องเรียง ¶" + ($expected -join ', ¶') + ")")
    }
  }
  return $keepNs
}

# G1 (docs/v2-plan.md 3.9): splits the section's ORIGINAL non-heading items into segments by
# anchor - segment 0 is everything before the section's first anchor, segment 1 is everything
# between the 1st and 2nd anchor, and so on (the section's own heading item is a member of
# segment 0 too, so a "heading" block rewriting it can legally cite it as its own "from"). Returns
# @{ AnchorSet; SegmentOf = @{ paragraphN -> segment index }; CoverageRequired = HashSet[int] of
# every non-heading, non-anchor ¶ that MUST show up in some block's "from" or as a "keep" (rule 3 -
# no silent drop). $ns is $null for .txt (no anchors possible there, so every non-heading item is
# segment 0 and anchorSet stays empty).
#
# Text-less pictures/blank paragraphs (Get-DocxGapGroups): $gapGroups are the gap
# groups between neighbouring items inside the section. A group with anchor-like content is an
# "implicit anchor": it is one more boundary between segments (no ¶, cannot be kept by number). A
# group without any is a "spacer" that travels with the item that follows it. Also returned:
#   Boundaries = ordered list of @{ Kind = 'explicit' (a "keep" anchor, N) | 'implicit' (Group, PrevN) };
#                boundary k separates segment k from segment k+1
#   AnchorBoundary = @{ anchorN -> boundary index }
#   SpacerByNext = @{ itemN -> gap group } (spacers keyed by the ¶ that follows them)
function Get-SegmentInfo($itemsInSection, [int]$headN, $ns, $gapGroups) {
  $anchorSet = New-Object 'System.Collections.Generic.HashSet[int]'
  if ($ns) {
    foreach ($it in $itemsInSection) {
      if ($it.N -eq $headN) { continue }
      if (Test-AnchorItem $it $ns) { [void]$anchorSet.Add($it.N) }
    }
  }
  $gapByNext = @{}
  foreach ($g in $gapGroups) { if ($null -ne $g) { $gapByNext[[int]$g.NextN] = $g } }
  $segmentOf = @{}
  $coverageRequired = New-Object 'System.Collections.Generic.HashSet[int]'
  $boundaries = New-Object System.Collections.ArrayList
  $anchorBoundary = @{}
  $spacerByNext = @{}
  foreach ($it in $itemsInSection) {
    if ($it.N -eq $headN) { $segmentOf[$it.N] = 0; continue }
    if ($gapByNext.ContainsKey($it.N)) {
      $g = $gapByNext[$it.N]
      if ($g.Anchor) { [void]$boundaries.Add(@{ Kind = 'implicit'; Group = $g; PrevN = $g.PrevN }) }
      else { $spacerByNext[$it.N] = $g }
    }
    if ($anchorSet.Contains($it.N)) {
      $anchorBoundary[$it.N] = $boundaries.Count
      [void]$boundaries.Add(@{ Kind = 'explicit'; N = $it.N })
      continue
    }
    $segmentOf[$it.N] = $boundaries.Count
    [void]$coverageRequired.Add($it.N)
  }
  return @{ AnchorSet = $anchorSet; SegmentOf = $segmentOf; CoverageRequired = $coverageRequired
            Boundaries = $boundaries; AnchorBoundary = $anchorBoundary; SpacerByNext = $spacerByNext }
}

# G1: validates one "heading"/"paragraph" block's "from" list against the section's segment map -
# (1) every from-¶ must exist in the section and must NOT be an anchor (anchors can only be cited
# via "keep", never rewritten into a new block); (2) all from-¶ of the SAME block must fall in the
# SAME anchor-delimited segment (a block cannot silently splice text from opposite sides of a table/
# image/etc. together); (3) that segment must equal $anchorKeepsSoFar - the number of anchor "keep"
# blocks the Plan has already placed before this block - i.e. the block's position in the Plan must
# line up with where its source text actually sits relative to the section's anchors, so content
# cannot be moved across an anchor. Marks every visited from-¶ as covered in $coveredIds (rule 3).
function Get-FromSegment([string]$sid, $segInfo, [int[]]$from, $coveredIds) {
  $segs = New-Object 'System.Collections.Generic.HashSet[int]'
  foreach ($f in $from) {
    if ($segInfo.AnchorSet.Contains($f)) { Fail "${sid}: from ¶$f ชี้ไปที่ตาราง/anchor เขียนใหม่ไม่ได้ (ต้องใช้ keep เท่านั้น)" }
    if (-not $segInfo.SegmentOf.ContainsKey($f)) { Fail "${sid}: from ¶$f ไม่ได้อยู่ใน section นี้" }
    [void]$segs.Add($segInfo.SegmentOf[$f])
    [void]$coveredIds.Add($f)
  }
  if ($segs.Count -ne 1) {
    $lo = [int](($segs | Measure-Object -Minimum).Minimum)
    $hi = [int](($segs | Measure-Object -Maximum).Maximum)
    $onlyImplicit = $true
    for ($k = $lo; $k -lt $hi; $k++) { if ($segInfo.Boundaries[$k].Kind -ne 'implicit') { $onlyImplicit = $false } }
    if ($onlyImplicit) {
      Fail ("${sid}: รวมเนื้อหาข้ามรูปที่ไม่มีข้อความไม่ได้ (from ของบล็อกเดียวกันต้องอยู่ฝั่งเดียวกันของรูป ¶" + ($from -join ', ¶') + ")")
    }
    Fail ("${sid}: from ของบล็อกเดียวกันต้องอยู่ช่วงเดียวกันระหว่าง anchor (¶" + ($from -join ', ¶') + ")")
  }
  return [int]@($segs)[0]
}

# .txt path (no implicit anchors, position is always segment 0).
function Test-FromIds([string]$sid, $segInfo, [int[]]$from, [int]$anchorKeepsSoFar, $coveredIds) {
  $segVal = Get-FromSegment $sid $segInfo $from $coveredIds
  if ($segVal -ne $anchorKeepsSoFar) {
    Fail ("${sid}: ตำแหน่งของบล็อกในแผนไม่ตรงกับตำแหน่งเดิมของ ¶" + ($from -join ', ¶') + " เทียบกับ anchor (ย้ายเนื้อหาข้าม anchor ไม่ได้)")
  }
}

# .docx path: moves the plan position $st.Pos (= number of segment boundaries already passed) to
# segment $target before a block whose source ¶ ($refs) lies in that segment. Boundaries in between
# that are implicit anchors (text-less pictures) get their nodes inserted (cloned, original order,
# tagged "G<k>" for the final order check); a "keep" anchor boundary in between, or a target that is
# behind the current position, is refused.
function Step-ToSegment([string]$sid, $segInfo, $st, [int]$target, [int[]]$refs, $newNodes, $newTags) {
  $refLabel = '¶' + ($refs -join ', ¶')
  $anchorMsg = "${sid}: ตำแหน่งของบล็อกในแผนไม่ตรงกับตำแหน่งเดิมของ $refLabel เทียบกับ anchor (ย้ายเนื้อหาข้าม anchor ไม่ได้)"
  if ($target -lt $st.Pos) {
    for ($k = $target; $k -lt $st.Pos; $k++) { if ($segInfo.Boundaries[$k].Kind -eq 'explicit') { Fail $anchorMsg } }
    $pic = $segInfo.Boundaries[$target]
    Fail "${sid}: ย้ายเนื้อหาข้ามรูปไม่ได้ ($refLabel อยู่ก่อนรูปที่ไม่มีข้อความซึ่งอยู่ต่อจาก ¶$($pic.PrevN) แต่แผนวางไว้หลังรูปนั้น)"
  }
  for ($k = $st.Pos; $k -lt $target; $k++) {
    $bd = $segInfo.Boundaries[$k]
    if ($bd.Kind -eq 'explicit') { Fail $anchorMsg }
    foreach ($gn in $bd.Group.Nodes) {
      [void]$newNodes.Add($gn.CloneNode($true))
      [void]$newTags.Add("G$k")
    }
  }
  $st.Pos = $target
}

# Spacers (blank paragraphs etc.): re-inserted before the first new node that comes from the ¶ they
# preceded in the original ($refs = the ¶ this block keeps / is rewritten from), once each.
function Add-Spacers($segInfo, $refs, $spacerDone, $newNodes, $newTags) {
  foreach ($rn in @($refs | Sort-Object -Unique)) {
    if ($segInfo.SpacerByNext.ContainsKey([int]$rn) -and $spacerDone.Add([int]$rn)) {
      foreach ($gn in $segInfo.SpacerByNext[[int]$rn].Nodes) { [void]$newNodes.Add($gn.CloneNode($true)); [void]$newTags.Add('') }
    }
  }
}

# Whole-document counts taken before anything is changed and re-taken before the file is written.
function Get-DocxCounts {
  $ns = $script:DocxNs
  return @{
    drawing = $script:DocxXml.SelectNodes('//w:drawing', $ns).Count
    pict    = $script:DocxXml.SelectNodes('//w:pict', $ns).Count
    object  = $script:DocxXml.SelectNodes('//w:object', $ns).Count
    tbl     = $script:DocxXml.SelectNodes('//w:tbl', $ns).Count
  }
}

# Self-check BEFORE any file is written: whole-document picture/OLE object/table counts unchanged, and
# every section's heading / "keep" anchors / implicit anchors are still in their original relative
# order. $verifyList holds the .Verify records returned by Build-DocxSection.
function Test-DocxRewriteIntact($countsBefore, $verifyList) {
  $countsAfter = Get-DocxCounts
  foreach ($key in 'drawing', 'pict', 'object', 'tbl') {
    if ($countsAfter[$key] -ne $countsBefore[$key]) {
      Fail "ตรวจไม่ผ่านก่อนบันทึก: จำนวน w:$key ในเอกสารไม่เท่าเดิม (เดิม $($countsBefore[$key]) ใหม่ $($countsAfter[$key])) จึงไม่สร้างไฟล์"
    }
  }
  foreach ($v in $verifyList) {
    $actual = New-Object System.Collections.ArrayList
    foreach ($c in $v.Parent.ChildNodes) { if ($v.TagMap.ContainsKey($c)) { [void]$actual.Add($v.TagMap[$c]) } }
    if (($actual -join ',') -ne $v.Expected) {
      Fail "$($v.Section): ตรวจไม่ผ่านก่อนบันทึก: ลำดับหัวข้อ ตาราง/รูปที่คงไว้ หลังสร้างไม่ตรงกับต้นฉบับ จึงไม่สร้างไฟล์"
    }
  }
}

# The blocks of a plan as the anchor/keep rules see them: a "table" block (a table kept in place whose cells get
# new text) is a "keep" of that table ¶.
function ConvertTo-ContractBlocks($blocks) {
  $out = New-Object System.Collections.ArrayList
  foreach ($b in $blocks) {
    if ($b.Type -eq 'table') { [void]$out.Add(@{ Type = 'keep'; Paragraph = [int]$b.Paragraph }) }
    else { [void]$out.Add($b) }
  }
  return , $out.ToArray()
}

# Writes the new text of the cells of one table block into $tblNode (the copy of the table that goes into the new
# document; $tableBlock = @{ Paragraph; Cells = @( @{ Cell; Text } ) }). Each cell must exist, be one that can be
# changed (Resolve-TableCellTarget), differ from its current text, and is listed once. Only the paragraph of the
# cell named is touched (Set-ParagraphTextPlain; a line break / special hyphen in it stays where it is, or the change
# that touches it is refused with the cell's coordinate); the cell's properties (w:tcPr), the paragraph's properties
# (w:pPr), every other cell and the table's structure stay as they are. Returns @{ Keys = <"<table ordinal>|<row>|
# <cell>|<paragraph>" per changed paragraph, for Test-DocxTablesIntact>; Notes = <lines to print> }.
function Set-TableCellTexts([string]$sid, [int]$tableN, $tblNode, $tableBlock, [int]$ordinal) {
  $cells = Get-TableCells $tblNode
  $keys = New-Object System.Collections.ArrayList
  $notes = New-Object System.Collections.ArrayList
  $coords = New-Object System.Collections.ArrayList
  foreach ($e in $tableBlock.Cells) {
    $res = Resolve-TableCellTarget $cells ([string]$e.Cell)
    if ($res.Error) { Fail "${sid}: ตาราง ¶$tableN $($e.Cell): $($res.Error)" }
    $rec = $res.Rec
    if ([string]::Equals(([string]$e.Text).Trim(), [string]$rec.Text, [StringComparison]::Ordinal)) {
      Fail "${sid}: ตาราง ¶$tableN $($e.Cell): ข้อความใหม่เหมือนข้อความเดิม ให้เอาเซลล์นี้ออกจาก cells"
    }
    $r = Set-ParagraphTextPlain $rec.Node ([string]$e.Text) "${sid}: ตาราง ¶$tableN $($e.Cell)"
    [void]$keys.Add(("{0}|{1}|{2}|{3}" -f $ordinal, $rec.Row, $rec.Cell, $(if ($rec.Para) { $rec.Para } else { 1 })))
    [void]$coords.Add([string]$e.Cell)
    if ($r.Whole) { [void]$notes.Add("${sid}: ตาราง ¶$tableN $($e.Cell): แทนทั้งย่อหน้าด้วยรูปแบบของข้อความส่วนแรก เพราะ$($r.Why)") }
    elseif ([int]$r.Breaks -gt 0 -or [int]$r.Hyphens -gt 0) {
      $kept = New-Object System.Collections.ArrayList
      if ([int]$r.Breaks -gt 0) { [void]$kept.Add("การขึ้นบรรทัดใหม่ $([int]$r.Breaks) จุด") }
      if ([int]$r.Hyphens -gt 0) { [void]$kept.Add("ขีดกลางพิเศษ $([int]$r.Hyphens) จุด") }
      [void]$notes.Add("${sid}: ตาราง ¶$tableN $($e.Cell): คง$($kept -join ' และ')ไว้ตำแหน่งเดิม")
    }
  }
  [void]$notes.Insert(0, "${sid}: ตาราง ¶$tableN เขียนใหม่ $($coords.Count) เซลล์ ($($coords -join ', ')) เซลล์อื่นและโครงตารางคงเดิม")
  return @{ Keys = @($keys); Notes = @($notes) }
}

# What every table of the document looks like, before anything is changed (taken by a script that is about to
# rewrite table cells): per table in document order (nested tables included) its properties and grid, and per
# row/cell the cell's whole XML, its w:tcPr, and its paragraphs' w:pPr and text.
function Get-DocxTableSnapshot {
  $ns = $script:DocxNs
  $list = New-Object System.Collections.ArrayList
  foreach ($tbl in $script:DocxXml.SelectNodes('//w:tbl', $ns)) {
    $struct = New-Object System.Collections.ArrayList
    foreach ($nm in 'w:tblPr', 'w:tblGrid') { $nd = $tbl.SelectSingleNode($nm, $ns); [void]$struct.Add($(if ($nd) { $nd.OuterXml } else { '' })) }
    $rows = New-Object System.Collections.ArrayList
    foreach ($tr in $tbl.SelectNodes('w:tr', $ns)) {
      $trPr = $tr.SelectSingleNode('w:trPr', $ns)
      [void]$struct.Add($(if ($trPr) { $trPr.OuterXml } else { '' }))
      $cellList = New-Object System.Collections.ArrayList
      foreach ($tc in $tr.SelectNodes('w:tc', $ns)) {
        $tcPr = $tc.SelectSingleNode('w:tcPr', $ns)
        $paras = New-Object System.Collections.ArrayList
        foreach ($pp in $tc.SelectNodes('w:p | w:sdt/w:sdtContent/w:p', $ns)) {
          $ppr = $pp.SelectSingleNode('w:pPr', $ns)
          [void]$paras.Add(@{ PPr = $(if ($ppr) { $ppr.OuterXml } else { '' }); Text = (Get-NodeText $pp -Raw) })
        }
        [void]$cellList.Add(@{ Xml = $tc.OuterXml; TcPr = $(if ($tcPr) { $tcPr.OuterXml } else { '' }); Paras = $paras })
      }
      [void]$rows.Add($cellList)
    }
    [void]$list.Add(@{ Struct = ($struct -join "`n"); Rows = $rows })
  }
  return , $list
}

# Self-check BEFORE the file is written, after a rewrite that changed table cells: every table of the document is
# still there with the same structure (properties, grid, rows, cells, each cell's w:tcPr and each paragraph's
# w:pPr) and every cell is exactly as it was - except the paragraphs named in $allowedKeys (the keys Set-TableCellTexts
# returned), whose text alone may differ. $before = Get-DocxTableSnapshot taken before any change.
function Test-DocxTablesIntact($before, $allowedKeys) {
  $after = Get-DocxTableSnapshot
  $allowed = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($k in $allowedKeys) { [void]$allowed.Add([string]$k) }
  $tail = 'จึงไม่สร้างไฟล์'
  if ($after.Count -ne $before.Count) { Fail "ตรวจไม่ผ่านก่อนบันทึก: จำนวนตารางในเอกสารไม่เท่าเดิม (เดิม $($before.Count) ใหม่ $($after.Count)) $tail" }
  for ($ti = 0; $ti -lt $before.Count; $ti++) {
    $b = $before[$ti]; $a = $after[$ti]
    $tn = $ti + 1
    if (-not [string]::Equals($a.Struct, $b.Struct, [StringComparison]::Ordinal)) { Fail "ตรวจไม่ผ่านก่อนบันทึก: โครงสร้างของตารางที่ $tn (คุณสมบัติ ตารางกริด หรือคุณสมบัติแถว) เปลี่ยนไป $tail" }
    if ($a.Rows.Count -ne $b.Rows.Count) { Fail "ตรวจไม่ผ่านก่อนบันทึก: จำนวนแถวของตารางที่ $tn ไม่เท่าเดิม (เดิม $($b.Rows.Count) ใหม่ $($a.Rows.Count)) $tail" }
    for ($ri = 0; $ri -lt $b.Rows.Count; $ri++) {
      if ($a.Rows[$ri].Count -ne $b.Rows[$ri].Count) { Fail "ตรวจไม่ผ่านก่อนบันทึก: จำนวนเซลล์ของตารางที่ $tn แถว $($ri + 1) ไม่เท่าเดิม (เดิม $($b.Rows[$ri].Count) ใหม่ $($a.Rows[$ri].Count)) $tail" }
      for ($ci = 0; $ci -lt $b.Rows[$ri].Count; $ci++) {
        $bc = $b.Rows[$ri][$ci]; $ac = $a.Rows[$ri][$ci]
        if ([string]::Equals($ac.Xml, $bc.Xml, [StringComparison]::Ordinal)) { continue }
        $where = "ตารางที่ $tn แถว $($ri + 1) เซลล์ $($ci + 1)"
        if (-not [string]::Equals($ac.TcPr, $bc.TcPr, [StringComparison]::Ordinal)) { Fail "ตรวจไม่ผ่านก่อนบันทึก: คุณสมบัติของเซลล์ ($where) เปลี่ยนไป $tail" }
        if ($ac.Paras.Count -ne $bc.Paras.Count) { Fail "ตรวจไม่ผ่านก่อนบันทึก: จำนวนย่อหน้าในเซลล์ ($where) ไม่เท่าเดิม $tail" }
        $k = 0
        $edited = $false
        for ($pj = 0; $pj -lt $bc.Paras.Count; $pj++) {
          if ($bc.Paras[$pj].Text.Trim()) { $k++ }
          if (-not [string]::Equals($ac.Paras[$pj].PPr, $bc.Paras[$pj].PPr, [StringComparison]::Ordinal)) { Fail "ตรวจไม่ผ่านก่อนบันทึก: คุณสมบัติของย่อหน้าในเซลล์ ($where) เปลี่ยนไป $tail" }
          $textSame = [string]::Equals($ac.Paras[$pj].Text, $bc.Paras[$pj].Text, [StringComparison]::Ordinal)
          if ($textSame) { continue }
          if (-not ($bc.Paras[$pj].Text.Trim() -and $allowed.Contains(("{0}|{1}|{2}|{3}" -f $ti, ($ri + 1), ($ci + 1), $k)))) {
            Fail "ตรวจไม่ผ่านก่อนบันทึก: ข้อความในเซลล์ที่ไม่ได้ระบุ ($where) เปลี่ยนไป $tail"
          }
          $edited = $true
        }
        # the cell's XML differs, so something other than the named paragraph's text must not be what changed
        if (-not $edited) { Fail "ตรวจไม่ผ่านก่อนบันทึก: เซลล์ที่ไม่ได้ระบุ ($where) เปลี่ยนไป $tail" }
      }
    }
  }
}

# Builds one section's replacement into the in-memory document ($script:DocxXml) and returns
# @{ Verify; Summary }. $sec = @{ Section; Blocks; Head; End } (Head/End from Resolve-Section);
# $items = Read-Docx result, $gapAll = Get-DocxGapGroups of the ORIGINAL items. Fails (exit 2) with a
# message naming the section and the ¶ at fault on any rule violation.
function Build-DocxSection($items, $ns, $gapAll, $sec) {
  $sid = $sec.Section; $head = $sec.Head; $end = $sec.End
  # A "table" block keeps its table in place exactly like a "keep" of an anchor (every rule below sees a keep);
  # its new cell texts are applied to the copy of the table that is put back (see the 'keep' case).
  $tableEdits = @{}
  foreach ($tb in $sec.Blocks) { if ($tb.Type -eq 'table') { $tableEdits[[int]$tb.Paragraph] = $tb } }
  $blocks = ConvertTo-ContractBlocks $sec.Blocks
  $tableEditKeys = New-Object System.Collections.ArrayList
  $tableNotes = New-Object System.Collections.ArrayList
  $itemsInSection = @($items | Where-Object { $_.N -ge $head.N -and $_.N -le $end })
  foreach ($it in $itemsInSection) {
    if (-not [object]::ReferenceEquals($it.Node.ParentNode, $head.Node.ParentNode)) {
      Fail "${sid}: เนื้อหาใน section นี้อยู่คนละระดับในเอกสาร (เช่น บางส่วนอยู่ใน content control) จึงสร้างใหม่ทั้ง section ไม่ได้ (¶$($it.N))"
    }
  }

  # ---- first block: must be this section's own heading ----
  $first = $blocks[0]
  $headIsAnchor = Test-AnchorItem $head $ns
  if ($first.Type -eq 'keep') {
    if ($first.Paragraph -ne $head.N) { Fail "${sid}: บล็อกแรกต้องเป็นหัวข้อของ section นี้ (keep ¶$($head.N)) ไม่ใช่ ¶$($first.Paragraph)" }
  } elseif ($first.Type -eq 'heading') {
    if ($headIsAnchor) { Fail "${sid}: หัวข้อเดิม (¶$($head.N)) มีรูป/วัตถุฝัง (OLE)/ลิงก์/ฟิลด์/เชิงอรรถ/comment/tracked changes/content control ต้องใช้ keep เท่านั้น เขียนใหม่ไม่ได้" }
    if ($first.Level -ne $head.Level) { Fail "${sid}: บล็อกแรกต้องเป็น heading level $($head.Level) (เท่าหัวข้อเดิม) ไม่ใช่ level $($first.Level)" }
  } else {
    Fail "${sid}: บล็อกแรกต้องเป็น keep หรือ heading ของหัวข้อ section นี้ ไม่ใช่ $($first.Type)"
  }

  # ---- anchor contract for the rest of the section ----
  [void](Test-AnchorContract $sid $itemsInSection $head.N $blocks $ns)

  # G1 (docs/v2-plan.md 3.9): segment map for this section's "from" checks, and the running
  # count of anchor "keep" blocks seen so far as we walk $blocks in Plan order (the position a
  # "from" block's segment must match).
  # Gap groups (text-less pictures / blank paragraphs) between this section's items only.
  $gapGroups = @($gapAll | Where-Object { $_.PrevN -ge $head.N -and $_.NextN -le $end })
  $segInfo = Get-SegmentInfo $itemsInSection $head.N $ns $gapGroups
  $coveredIds = New-Object 'System.Collections.Generic.HashSet[int]'
  $st = @{ Pos = 0 }
  $spacerDone = New-Object 'System.Collections.Generic.HashSet[int]'

  # ---- template nodes for new heading/paragraph blocks ----
  $firstBodyItem = ($itemsInSection | Where-Object { Test-Body $_ } | Select-Object -First 1)
  $bodyPPr = $null; $bodyRPr = $null
  if ($firstBodyItem) {
    $bodyPPr = $firstBodyItem.Node.SelectSingleNode('w:pPr', $ns)
    $firstRun = $firstBodyItem.Node.SelectSingleNode('w:r[w:t]', $ns)
    if ($firstRun) { $bodyRPr = $firstRun.SelectSingleNode('w:rPr', $ns) }
  }

  # ---- build the replacement node sequence ----
  $newNodes = New-Object System.Collections.ArrayList
  $newTags = New-Object System.Collections.ArrayList   # parallel to $newNodes: 'H' heading, 'A<N>' keep anchor, 'G<k>' implicit anchor node, else ''
  $newParaCount = 0
  $blockCount = 0
  foreach ($b in $blocks) {
    $blockCount++
    switch ($b.Type) {
      'keep' {
        $refs = @([int]$b.Paragraph)
        $isAnchorKeep = $segInfo.AnchorSet.Contains($b.Paragraph)
        $targetSeg = if ($isAnchorKeep) { [int]$segInfo.AnchorBoundary[$b.Paragraph] } else { [int]$segInfo.SegmentOf[$b.Paragraph] }
        Step-ToSegment $sid $segInfo $st $targetSeg $refs $newNodes $newTags
        [void]$coveredIds.Add($b.Paragraph)
        Add-Spacers $segInfo $refs $spacerDone $newNodes $newTags
        $srcItem = $itemsInSection | Where-Object { $_.N -eq $b.Paragraph } | Select-Object -First 1
        $keptNode = $srcItem.Node.CloneNode($true)
        if ($tableEdits.ContainsKey([int]$b.Paragraph)) {
          if ($srcItem.Kind -ne 'table') { Fail "${sid}: ¶$($b.Paragraph) ไม่ใช่ตาราง เขียนข้อความในเซลล์ไม่ได้ (ใช้ keep ตามปกติ)" }
          $ordinal = [array]::IndexOf(@($script:DocxXml.SelectNodes('//w:tbl', $ns)), $srcItem.Node)
          $applied = Set-TableCellTexts $sid ([int]$b.Paragraph) $keptNode $tableEdits[[int]$b.Paragraph] $ordinal
          foreach ($k in $applied.Keys) { [void]$tableEditKeys.Add($k) }
          foreach ($nl in $applied.Notes) { [void]$tableNotes.Add($nl) }
        }
        [void]$newNodes.Add($keptNode)
        if ($isAnchorKeep) { [void]$newTags.Add("A$($b.Paragraph)"); $st.Pos = $targetSeg + 1 }
        elseif ($blockCount -eq 1) { [void]$newTags.Add('H') }
        else { [void]$newTags.Add('') }
      }
      'heading' {
        $styleItem = Find-HeadingStyleItem $items $b.Level
        if (-not $styleItem) { Fail "${sid}: ไม่มี heading level $($b.Level) อยู่ในเอกสารนี้เลย จึงไม่มีสไตล์ให้ใช้ (สร้าง heading ระดับใหม่ไม่ได้)" }
        $targetSeg = Get-FromSegment $sid $segInfo $b.From $coveredIds
        Step-ToSegment $sid $segInfo $st $targetSeg $b.From $newNodes $newTags
        Add-Spacers $segInfo $b.From $spacerDone $newNodes $newTags
        $hPPr = $styleItem.Node.SelectSingleNode('w:pPr', $ns)
        $hRunNode = $styleItem.Node.SelectSingleNode('w:r[w:t]', $ns)
        $hRPr = if ($hRunNode) { $hRunNode.SelectSingleNode('w:rPr', $ns) } else { $null }
        [void]$newNodes.Add((New-PlainParagraph $hPPr $hRPr $b.Text))
        if ($blockCount -eq 1) { [void]$newTags.Add('H') } else { [void]$newTags.Add('') }
        $newParaCount++
      }
      'paragraph' {
        $targetSeg = Get-FromSegment $sid $segInfo $b.From $coveredIds
        Step-ToSegment $sid $segInfo $st $targetSeg $b.From $newNodes $newTags
        Add-Spacers $segInfo $b.From $spacerDone $newNodes $newTags
        # L1 (dry-run round 7): the new paragraph takes its style from its FIRST source ¶
        # ("from"[0]) when that is an ordinary non-heading paragraph - so a caption, list item or
        # other specially-styled paragraph keeps its own style instead of silently becoming the
        # section's first body style. A heading ¶ used as a source falls back to the body style
        # (a "paragraph" block must never turn into a heading).
        $tpl = $itemsInSection | Where-Object { $_.N -eq $b.From[0] } | Select-Object -First 1
        if ($tpl -and $tpl.Kind -eq 'p' -and $tpl.Level -eq 0) {
          $pPr = $tpl.Node.SelectSingleNode('w:pPr', $ns)
          $tplRun = $tpl.Node.SelectSingleNode('w:r[w:t]', $ns)
          $rPr = if ($tplRun) { $tplRun.SelectSingleNode('w:rPr', $ns) } else { $null }
        } else {
          if (-not $firstBodyItem) { Fail "${sid}: ไม่มีย่อหน้าเนื้อความเดิมใน section นี้ให้ใช้เป็นต้นแบบสไตล์ จึงสร้างย่อหน้าใหม่ไม่ได้ (มีแต่หัวข้อ/ตาราง)" }
          $pPr = $bodyPPr; $rPr = $bodyRPr
        }
        [void]$newNodes.Add((New-PlainParagraph $pPr $rPr $b.Text))
        [void]$newTags.Add('')
        $newParaCount++
      }
    }
  }

  # G1 rule 3: every original non-heading, non-anchor ¶ must appear in some block's "from" or as
  # a "keep" - no silent drop.
  $missingCov = @($segInfo.CoverageRequired | Where-Object { -not $coveredIds.Contains($_) })
  if ($missingCov.Count) {
    Fail ("${sid}: ย่อหน้าเดิม ¶" + ($missingCov -join ', ¶') + " หายไปจากแผน (ต้องปรากฏใน from ของบล็อกใดบล็อกหนึ่ง หรือเป็น keep)")
  }

  # Any text-less pictures left after the last block go at the end of the section (in order).
  Step-ToSegment $sid $segInfo $st $segInfo.Boundaries.Count @() $newNodes $newTags

  # ---- splice into the DOM: remove the original section's nodes (numbered items AND the gap
  # nodes between them, all re-inserted above as spacers/implicit anchors), insert the new sequence ----
  $origNodes = New-Object System.Collections.ArrayList
  foreach ($it in $itemsInSection) { [void]$origNodes.Add($it.Node) }
  foreach ($g in $gapGroups) { foreach ($gn in $g.Nodes) { [void]$origNodes.Add($gn) } }
  $parent = $head.Node.ParentNode
  $lastNode = $itemsInSection[$itemsInSection.Count - 1].Node
  $insertBefore = $lastNode.NextSibling
  foreach ($n in $origNodes) { [void]$parent.RemoveChild($n) }
  $tagMap = New-Object 'System.Collections.Generic.Dictionary[System.Xml.XmlNode,string]'
  for ($ni = 0; $ni -lt $newNodes.Count; $ni++) {
    $n = $newNodes[$ni]
    if ($newTags[$ni]) { $tagMap[$n] = [string]$newTags[$ni] }
    if ($null -ne $insertBefore) { [void]$parent.InsertBefore($n, $insertBefore) } else { [void]$parent.AppendChild($n) }
  }
  # Expected relative order of the heading, "keep" anchors and implicit anchors (from the ORIGINAL).
  $expectedTags = New-Object System.Collections.ArrayList
  [void]$expectedTags.Add('H')
  for ($k = 0; $k -lt $segInfo.Boundaries.Count; $k++) {
    $bd = $segInfo.Boundaries[$k]
    if ($bd.Kind -eq 'explicit') { [void]$expectedTags.Add("A$($bd.N)") }
    else { foreach ($gn in $bd.Group.Nodes) { [void]$expectedTags.Add("G$k") } }
  }
  $verify = @{ Section = $sid; Parent = $parent; TagMap = $tagMap; Expected = ($expectedTags -join ',') }

  $anchorCount = @($itemsInSection | Where-Object { $_.N -ne $head.N -and (Test-AnchorItem $_ $ns) }).Count
  $implicitGroups = @($segInfo.Boundaries | Where-Object { $_.Kind -eq 'implicit' })
  $imgCount = @($implicitGroups | Where-Object { $_.Group.HasImage }).Count
  $otherImplicit = $implicitGroups.Count - $imgCount
  $extra = ''
  if ($imgCount -gt 0) { $extra += ", รูปไม่มีข้อความ $imgCount จุด" }
  if ($otherImplicit -gt 0) { $extra += ", เนื้อหาอื่นที่ไม่มีข้อความ $otherImplicit จุด" }
  return @{ Verify = $verify; Summary = "$sid`: ย่อหน้าเดิม $($itemsInSection.Count) -> ใหม่ $blockCount (คง anchor $anchorCount รายการ$extra)"
            TableKeys = @($tableEditKeys); TableNotes = @($tableNotes) }
}

# .txt counterpart of Build-DocxSection: validates one section against the same rules (a .txt has no
# tables/images/links/footnotes/comments/tracked-changes/content-controls to anchor) and returns
# @{ SectionStart; SectionEnd; NewText; Summary } - the caller splices NewText over [SectionStart,
# SectionEnd) of $text, back to front. $sep is the paragraph separator of the document.
function Build-TxtSection($items, [string]$text, [string]$sep, $sec) {
  $sid = $sec.Section; $head = $sec.Head; $end = $sec.End; $blocks = $sec.Blocks
  if (@($blocks | Where-Object { $_.Type -eq 'table' }).Count) { Fail "${sid}: ไฟล์ .txt ไม่มีตาราง เขียนข้อความในเซลล์ตารางไม่ได้" }
  $itemsInSection = @($items | Where-Object { $_.N -ge $head.N -and $_.N -le $end })

  $first = $blocks[0]
  if ($first.Type -eq 'keep') {
    if ($first.Paragraph -ne $head.N) { Fail "${sid}: บล็อกแรกต้องเป็นหัวข้อของ section นี้ (keep ¶$($head.N)) ไม่ใช่ ¶$($first.Paragraph)" }
  } elseif ($first.Type -eq 'heading') {
    if ($first.Level -ne $head.Level) { Fail "${sid}: บล็อกแรกต้องเป็น heading level $($head.Level) (เท่าหัวข้อเดิม) ไม่ใช่ level $($first.Level)" }
  } else {
    Fail "${sid}: บล็อกแรกต้องเป็น keep หรือ heading ของหัวข้อ section นี้ ไม่ใช่ $($first.Type)"
  }
  [void](Test-AnchorContract $sid $itemsInSection $head.N $blocks $null)

  # G1: .txt has no anchors, so the whole section is segment 0 - but the "from"/no-silent-drop
  # checks still apply (rule 3: every original ¶ must be covered by a from or a keep).
  $segInfo = Get-SegmentInfo $itemsInSection $head.N $null @()
  $coveredIds = New-Object 'System.Collections.Generic.HashSet[int]'
  $anchorKeepsSoFar = 0

  $newTexts = New-Object System.Collections.ArrayList
  foreach ($b in $blocks) {
    switch ($b.Type) {
      'keep' {
        [void]$coveredIds.Add($b.Paragraph)
        $srcItem = $itemsInSection | Where-Object { $_.N -eq $b.Paragraph } | Select-Object -First 1
        [void]$newTexts.Add($text.Substring($srcItem.Start, $srcItem.Length))
      }
      'heading' {
        $styleItem = Find-HeadingStyleItem $items $b.Level
        if (-not $styleItem) { Fail "${sid}: ไม่มี heading level $($b.Level) อยู่ในเอกสารนี้เลย จึงไม่มีสไตล์ให้ใช้ (สร้าง heading ระดับใหม่ไม่ได้)" }
        Test-FromIds $sid $segInfo $b.From $anchorKeepsSoFar $coveredIds
        [void]$newTexts.Add((('#' * $b.Level) + ' ' + $b.Text))
      }
      'paragraph' {
        $hasBody = [bool]($itemsInSection | Where-Object { Test-Body $_ })
        if (-not $hasBody) { Fail "${sid}: ไม่มีย่อหน้าเนื้อความเดิมใน section นี้ จึงสร้างย่อหน้าใหม่ไม่ได้ (มีแต่หัวข้อ)" }
        Test-FromIds $sid $segInfo $b.From $anchorKeepsSoFar $coveredIds
        [void]$newTexts.Add($b.Text)
      }
    }
  }
  $missingCov = @($segInfo.CoverageRequired | Where-Object { -not $coveredIds.Contains($_) })
  if ($missingCov.Count) {
    Fail ("${sid}: ย่อหน้าเดิม ¶" + ($missingCov -join ', ¶') + " หายไปจากแผน (ต้องปรากฏใน from ของบล็อกใดบล็อกหนึ่ง หรือเป็น keep)")
  }
  $lastItem = $itemsInSection[$itemsInSection.Count - 1]
  return @{
    SectionStart = $head.Start; SectionEnd = ($lastItem.Start + $lastItem.Length)
    NewText = ($newTexts -join $sep); Summary = "$sid`: ย่อหน้าเดิม $($itemsInSection.Count) -> ใหม่ $($newTexts.Count)"
  }
}

# =============================================================================================
# The rewrite preview is the single source of the text that ends up in the new file: review-file.ps1
# -Kind rewrite writes, per section, (a) the human-readable Before/After and (b) one machine-readable
# line holding the section's blocks (the same shape as a write-rewrite.ps1 -Plan section). Both are
# inside the part of the file the trailer's sha256 covers. write-rewrite.ps1 builds from (b) and
# refuses a section whose visible "หลัง" text no longer matches (b) - a user edit in the file.
# =============================================================================================

# "ย่อหน้าใหม่ N:" (or "ย่อหน้าใหม่ N (หัวข้อ ระดับ L):" for a rewritten heading) + the quoted text.
function Get-NewBlockPreviewLines([int]$index, $block) {
  $head = if ($block.Type -eq 'heading') { "ย่อหน้าใหม่ $index (หัวข้อ ระดับ $($block.Level)):" } else { "ย่อหน้าใหม่ ${index}:" }
  $lines = New-Object System.Collections.ArrayList
  [void]$lines.Add($head)
  foreach ($l in (Get-QuotedLines $block.Text)) { [void]$lines.Add($l) }
  return , @($lines)
}

# A kept item as shown under both ก่อน and หลัง: an anchor (table, image, link, ...) or a heading is
# "[คงไว้: ¶7 <short excerpt>]"; an ordinary paragraph is "[คงเดิม: ¶8]" followed by its full text
# in a quote block, so the user can read what stays exactly as it is.
function Get-KeepPreviewLines([int]$n, [string]$text, [bool]$isAnchorOrHeading) {
  if ($isAnchorOrHeading) { return , @("[คงไว้: ¶$n $(Get-ShortText $text)]") }
  $lines = New-Object System.Collections.ArrayList
  [void]$lines.Add("[คงเดิม: ¶$n]")
  foreach ($l in (Get-QuotedLines $text)) { [void]$lines.Add($l) }
  return , @($lines)
}

# What the หลัง side of a preview shows under the "[คงไว้: ¶7 ...]" line of a table whose cells are rewritten:
#   [ตาราง ¶7: เขียนใหม่ 1 เซลล์ เซลล์อื่นคงเดิม]
#   เซลล์ r4c2 (หัวคอลัมน์: ... | หัวแถว: ...):
#   ก่อน:
#   > <the cell's text now>
#   หลัง:
#   > <the new text>
# $infos = per cell, in the block's order: @{ Cell; Before; ColumnHead; RowHead } (read from the source document).
function Get-TableBlockPreviewLines($block, $infos) {
  $lines = New-Object System.Collections.ArrayList
  [void]$lines.Add("[ตาราง ¶$($block.Paragraph): เขียนใหม่ $(@($block.Cells).Count) เซลล์ เซลล์อื่นคงเดิม]")
  foreach ($c in $block.Cells) {
    $info = $infos | Where-Object { $_.Cell -ceq $c.Cell } | Select-Object -First 1
    $ctx = New-Object System.Collections.ArrayList
    if ($info -and $info.ColumnHead) { [void]$ctx.Add("หัวคอลัมน์: $($info.ColumnHead)") }
    if ($info -and $info.RowHead) { [void]$ctx.Add("หัวแถว: $($info.RowHead)") }
    $label = "เซลล์ $($c.Cell)"
    if ($ctx.Count) { $label += ' (' + ($ctx -join ' | ') + ')' }
    [void]$lines.Add($label + ':')
    [void]$lines.Add('ก่อน:')
    foreach ($l in (Get-QuotedLines $(if ($info) { [string]$info.Before } else { '' }))) { [void]$lines.Add($l) }
    [void]$lines.Add('หลัง:')
    foreach ($l in (Get-QuotedLines ([string]$c.Text))) { [void]$lines.Add($l) }
  }
  return , @($lines)
}

# One line per section: <!-- hz-plan v2 section="S3" data="<base64 of gzip of UTF-8 JSON>" -->
# (v1, still readable by Get-PreviewPlanData, was the base64 of the UTF-8 JSON without compression.) A section
# that rewrites table cells (a "table" block) is written as v3, same encoding as v2, so that a reader that does
# not know table blocks refuses the line instead of dropping the cells; a section without any stays v2.
function ConvertTo-PlanDataLine([string]$sid, $blocks) {
  $list = New-Object System.Collections.ArrayList
  $version = 2
  foreach ($b in $blocks) {
    switch ($b.Type) {
      'keep' { [void]$list.Add([ordered]@{ type = 'keep'; paragraph = [int]$b.Paragraph }) }
      'table' {
        $version = 3
        [void]$list.Add([ordered]@{ type = 'table'; paragraph = [int]$b.Paragraph; cells = @($b.Cells | ForEach-Object { [ordered]@{ cell = [string]$_.Cell; text = [string]$_.Text } }) })
      }
      'heading' { [void]$list.Add([ordered]@{ type = 'heading'; level = [int]$b.Level; text = [string]$b.Text; from = @($b.From | ForEach-Object { [int]$_ }) }) }
      'paragraph' { [void]$list.Add([ordered]@{ type = 'paragraph'; text = [string]$b.Text; from = @($b.From | ForEach-Object { [int]$_ }) }) }
    }
  }
  $json = ConvertTo-Json -InputObject ([ordered]@{ section = $sid; blocks = @($list) }) -Depth 6 -Compress
  $jsonBytes = $utf8.GetBytes($json)
  $packed = New-Object System.IO.MemoryStream
  $gzOut = New-Object System.IO.Compression.GZipStream($packed, [System.IO.Compression.CompressionMode]::Compress, $true)
  try { $gzOut.Write($jsonBytes, 0, $jsonBytes.Length) } finally { $gzOut.Dispose() }
  $b64 = [Convert]::ToBase64String($packed.ToArray())
  $packed.Dispose()
  return "<!-- hz-plan v$version section=`"$sid`" data=`"$b64`" -->"
}

# The UTF-8 JSON text of an hz-plan line's data: v1 = base64 of the JSON, v2 and v3 = base64 of its gzip.
# The unpacked size is capped (8 MB) so a damaged or hostile line cannot exhaust memory; over the cap throws.
function ConvertFrom-PlanDataText([string]$version, [string]$b64) {
  $raw = [Convert]::FromBase64String($b64)
  if ($version -eq '1') { return $utf8.GetString($raw) }
  $cap = 8MB
  $src = New-Object System.IO.MemoryStream(, $raw)
  $gzIn = New-Object System.IO.Compression.GZipStream($src, [System.IO.Compression.CompressionMode]::Decompress)
  $dst = New-Object System.IO.MemoryStream
  try {
    $buf = New-Object byte[] 8192
    while (($n = $gzIn.Read($buf, 0, $buf.Length)) -gt 0) {
      $dst.Write($buf, 0, $n)
      if ($dst.Length -gt $cap) { throw 'plan data too large' }
    }
    return $utf8.GetString($dst.ToArray())
  } finally { $gzIn.Dispose(); $src.Dispose(); $dst.Dispose() }
}

# Reads the hz-plan lines (v1, v2 and v3) of a preview's body lines. Returns @{ 'S3' = <block list> | $null }; $null
# means the line is there but cannot be decoded. A section without a line is simply absent.
function Get-PreviewPlanData($bodyLines) {
  $map = @{}
  foreach ($l in $bodyLines) {
    if ($l -notmatch '^<!--\s*hz-plan\s+v([123])\s+section="(S\d+)"\s+data="([A-Za-z0-9+/=]*)"\s*-->\s*$') { continue }
    $ver = $Matches[1]; $sid = $Matches[2]; $b64 = $Matches[3]
    if ($map.ContainsKey($sid)) { continue }
    $blocks = $null
    try {
      $json = ConvertFrom-PlanDataText $ver $b64
      $obj = ConvertFrom-Json -InputObject $json
      $raw = Get-AsArray $obj.blocks
      if ($raw.Count -and [string]$obj.section -ceq $sid) { $blocks = $raw }
    } catch { $blocks = $null }
    $map[$sid] = if ($null -ne $blocks) { ConvertTo-PlanBlocks $blocks $sid } else { $null }
  }
  return $map
}

# The lines under "**หลัง**" of section $sid in a preview (trailing blank lines removed), or $null when
# the section or its หลัง block is not found.
function Get-PreviewAfterLines($bodyLines, [string]$sid) {
  $i = 0; $n = $bodyLines.Count
  $start = -1
  for (; $i -lt $n; $i++) { if ($bodyLines[$i] -match "^###\s+$sid(\s|$)") { $start = $i; break } }
  if ($start -lt 0) { return $null }
  $j = $start + 1
  $after = -1
  for (; $j -lt $n; $j++) {
    if ($bodyLines[$j] -match '^### ') { break }
    if ($bodyLines[$j].TrimEnd() -eq '**หลัง**') { $after = $j; break }
  }
  if ($after -lt 0) { return $null }
  $out = New-Object System.Collections.ArrayList
  for ($k = $after + 1; $k -lt $n; $k++) {
    $l = $bodyLines[$k]
    if ($l -match '^### ' -or $l -match '^\*\*') { break }
    [void]$out.Add($l.TrimEnd())
  }
  while ($out.Count -and -not $out[$out.Count - 1]) { $out.RemoveAt($out.Count - 1) }
  return , @($out)
}

# $true when the visible หลัง lines still say exactly what $blocks say: same order, same kept ¶,
# "ย่อหน้าใหม่ N" numbered 1.. over the non-keep blocks, same heading level and the same text, line by
# line (trailing spaces ignored). The excerpt in a "[คงไว้: ...]" label and the quoted text after a
# "[คงเดิม: ¶N]" are not compared - they come from the source file and are only there to be read.
function Test-PreviewAfterMatchesBlocks($afterLines, $blocks) {
  if ($null -eq $afterLines) { return $false }
  $entries = New-Object System.Collections.ArrayList
  $cur = $null
  $curCell = $null
  $curMode = $null
  foreach ($raw in $afterLines) {
    $l = ([string]$raw).TrimEnd()
    if (-not $l) { continue }
    if ($l -match '^ย่อหน้าใหม่ (\d+)(?: \(หัวข้อ ระดับ (\d+)\))?:$') {
      $cur = @{ Kind = 'new'; N = [int]$Matches[1]; Level = $(if ($Matches[2]) { [int]$Matches[2] } else { 0 }); Lines = (New-Object System.Collections.ArrayList) }
      $curCell = $null
      [void]$entries.Add($cur); continue
    }
    if ($l -match '^\[คงไว้: ¶(\d+)(?: .*)?\]$') { [void]$entries.Add(@{ Kind = 'keep'; Id = [int]$Matches[1] }); $cur = $null; $curCell = $null; continue }
    if ($l -match '^\[คงเดิม: ¶(\d+)\]$') { $cur = @{ Kind = 'keep'; Id = [int]$Matches[1]; Plain = $true }; $curCell = $null; [void]$entries.Add($cur); continue }
    # a table whose cells are rewritten: "[ตาราง ¶7: ...]" right after its "[คงไว้: ¶7 ...]" line, then per cell the
    # label, "ก่อน:" + quote, "หลัง:" + quote (only the หลัง text is compared - ก่อน is read from the source file)
    if ($l -match '^\[ตาราง ¶(\d+): เขียนใหม่ (\d+) เซลล์ เซลล์อื่นคงเดิม\]$') {
      $lastEntry = if ($entries.Count) { $entries[$entries.Count - 1] } else { $null }
      if (-not $lastEntry -or $lastEntry.Kind -ne 'keep' -or $lastEntry.Plain -or $lastEntry.Table -or $lastEntry.Id -ne [int]$Matches[1]) { return $false }
      $lastEntry.Table = @{ Declared = [int]$Matches[2]; Cells = (New-Object System.Collections.ArrayList) }
      $cur = $null; $curCell = $null; continue
    }
    if ($l -match '^เซลล์ (r\d+c\d+(?:p\d+)?)(?: \(.*\))?:$') {
      $lastEntry = if ($entries.Count) { $entries[$entries.Count - 1] } else { $null }
      if (-not $lastEntry -or -not $lastEntry.Table) { return $false }
      $curCell = @{ Cell = $Matches[1]; After = (New-Object System.Collections.ArrayList) }
      [void]$lastEntry.Table.Cells.Add($curCell)
      $curMode = $null; continue
    }
    if ($l -ceq 'ก่อน:' -or $l -ceq 'หลัง:') {
      if ($null -eq $curCell) { return $false }
      $curMode = if ($l -ceq 'ก่อน:') { 'before' } else { 'after' }
      continue
    }
    if ($l -match '^>(?: (.*))?$') {
      if ($curCell) {
        if ($curMode -eq 'after') { [void]$curCell.After.Add($(if ($null -ne $Matches[1]) { [string]$Matches[1] } else { '' })) }
        continue
      }
      if ($null -eq $cur) { return $false }
      if ($cur.Kind -eq 'new') { [void]$cur.Lines.Add($(if ($null -ne $Matches[1]) { [string]$Matches[1] } else { '' })) }
      continue
    }
    return $false
  }
  if ($entries.Count -ne $blocks.Count) { return $false }
  $newIdx = 0
  for ($i = 0; $i -lt $blocks.Count; $i++) {
    $b = $blocks[$i]; $e = $entries[$i]
    if ($b.Type -eq 'keep') {
      if ($e.Kind -ne 'keep' -or $e.Id -ne $b.Paragraph -or $e.Table) { return $false }
      continue
    }
    if ($b.Type -eq 'table') {
      if ($e.Kind -ne 'keep' -or $e.Plain -or $e.Id -ne $b.Paragraph -or -not $e.Table) { return $false }
      $wantCells = @($b.Cells)
      if ($e.Table.Declared -ne $wantCells.Count -or $e.Table.Cells.Count -ne $wantCells.Count) { return $false }
      for ($k = 0; $k -lt $wantCells.Count; $k++) {
        $ec = $e.Table.Cells[$k]
        if (-not [string]::Equals($ec.Cell, [string]$wantCells[$k].Cell, [StringComparison]::Ordinal)) { return $false }
        $want = @((([string]$wantCells[$k].Text) -replace "`r`n", "`n" -replace "`r", "`n") -split "`n" | ForEach-Object { $_.TrimEnd() })
        $have = @($ec.After | ForEach-Object { ([string]$_).TrimEnd() })
        if ($want.Count -ne $have.Count) { return $false }
        for ($m = 0; $m -lt $want.Count; $m++) { if (-not [string]::Equals($want[$m], $have[$m], [StringComparison]::Ordinal)) { return $false } }
      }
      continue
    }
    $newIdx++
    if ($e.Kind -ne 'new' -or $e.N -ne $newIdx) { return $false }
    $wantLevel = if ($b.Type -eq 'heading') { [int]$b.Level } else { 0 }
    if ($e.Level -ne $wantLevel) { return $false }
    $want = @((([string]$b.Text) -replace "`r`n", "`n" -replace "`r", "`n") -split "`n" | ForEach-Object { $_.TrimEnd() })
    $have = @($e.Lines | ForEach-Object { ([string]$_).TrimEnd() })
    if ($want.Count -ne $have.Count) { return $false }
    for ($k = 0; $k -lt $want.Count; $k++) { if (-not [string]::Equals($want[$k], $have[$k], [StringComparison]::Ordinal)) { return $false } }
  }
  return $true
}
