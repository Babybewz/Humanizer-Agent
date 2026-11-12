<#
  Build a NEW file with one or more sections rewritten wholesale (section/whole-file rewrite scope,
  docs/v2-plan.md Phase 3.6 item 3A-4). Never modifies or overwrites the source. No tracked changes
  in the output (D-O1): unlike apply-revisions.ps1, this replaces whole sections, so the paragraph-
  level Accept/Reject workflow does not apply here - the user compares old vs new with the review
  preview .md (review-file.ps1 -Kind rewrite) or Word's own Compare feature.

  The rewrite preview is the ONE source of the new text. review-file.ps1 -Kind rewrite stores, next to
  the readable Before/After, one machine-readable line per section (the section's blocks, inside the
  part of the file that the trailer's sha256 covers). This script builds the file from those blocks
  for every section the user ticked - the agent does not send the text a second time.

  -Path <source>   .docx or .txt (.pdf is rejected - no per-paragraph structure to anchor against;
                    "แปลงเป็น .docx ก่อน" is the same guidance apply-revisions.ps1 gives elsewhere).
  -Review <path[]> MANDATORY (F1, docs/v2-plan.md 3.8; G2, 3.9): one or more rewrite preview .md
                    files created by review-file.ps1 -Action Create -Kind rewrite for this SAME
                    source file, after the user has reviewed/ticked them - an array, a single
                    comma-separated string, or a mix (each element is itself split on ","). Each
                    file is read with the exact same parser doc-lib.ps1's Read-ReviewFile uses for
                    review-file.ps1 -Action Read (no second, drifting copy). Checked before
                    anything else runs, for EVERY supplied file:
                      - the file must parse as a review/report/brief/rewrite .md (Read-ReviewFile
                        itself refuses a foreign/malformed file);
                      - its Kind must be "rewrite" (a review/report/brief file is refused);
                      - its recorded source name (the "ไฟล์ต้นฉบับ: <name>" line in the body) must
                        match -Path's own file name, or the whole run is refused - this is the
                        code-side guarantee that a preview from one document can never be used to
                        build a DIFFERENT document.
                    Then, per section (G2): among the supplied files that contain that section at
                    all, the one most recently CREATED (its own "- สร้างเมื่อ:" line; a tie favors
                    whichever file appears LATER in -Review) is the one used for that section.
                    Without -Plan the sections built are the ones TICKED in the preview chosen for them
                    (a section that is not ticked there is left as it is and listed as skipped; if
                    nothing is ticked the run is refused). With -Plan every section named in the plan
                    must be ticked in its chosen preview, or the whole run is refused, naming the
                    unticked section(s) - a plan can never silently include a section the user did not
                    tick off (dry-run finding F1). The summary output prints, per section built, which
                    preview file (its short sha256 id and name) was used for it.
                    For every section built, the visible "หลัง" text of the chosen preview must still
                    say what the embedded blocks say (same kept ¶, same "ย่อหน้าใหม่ N", same heading
                    level, same text): if the user edited that text in the file, the section is refused
                    ("ข้อความใน preview ของ S3 ถูกแก้ในไฟล์ ...") - a change of wording goes through the
                    chat and a new preview. A preview without embedded blocks (made by an older version)
                    is refused too. If a preview used was edited elsewhere (outside its checkboxes) this
                    prints the same warning -Action Read prints - it does NOT refuse the run for that.
  -Plan <json>     OPTIONAL. UTF-8 JSON file, in the system temp folder (deleted after the run - success
                    or error - same as apply-revisions.ps1's -Revisions). It is kept for callers that
                    still build the blocks themselves; the text of EVERY block must equal the preview's
                    (the run is refused, naming the section and the ¶, when it does not), so it can never
                    produce text the user did not see. Shape:
                      { "sections": [ { "section": "S3",
                          "blocks": [
                            { "type": "keep", "paragraph": 8 },
                            { "type": "heading", "level": 2, "text": "...", "from": [6] },
                            { "type": "paragraph", "text": "...", "from": [9, 10] },
                            { "type": "table", "paragraph": 4, "cells": [ { "cell": "r4c2", "text": "..." } ] }
                          ] } ] }
                    A "table" block keeps table ¶4 where it is (it counts as the keep of that anchor) and writes
                    new text into the listed cells only - one paragraph of a cell per coordinate (rRcC, or rRcCpK
                    when the cell has several paragraphs; the coordinates extract-text.ps1 -Cells prints).
                    "S3" etc. use the same section numbering as extract-text.ps1's outline (-Mode
                    Outline), and "paragraph" in a "keep" block is the same ¶ numbering as
                    extract-text.ps1/apply-revisions.ps1 (via doc-lib.ps1 Read-Docx/Read-Txt). A
                    "heading"/"paragraph" block also requires "from" (docs/v2-plan.md 3.9 G1): an
                    array of one or more original ¶ numbers this new block is rewritten from - see
                    the "from" section below. The blocks embedded in a preview have exactly this shape.
  -Out / -OutDir   same as apply-revisions.ps1 (default: <output root>\humanized\<name>\<name>_rewritten_
                    <yyyy-MM-dd>.docx|.txt; <output root> = <MyDocuments>\HumanizerAgent\output or env HUMANIZER_OUTPUT_DIR; -Out is an exact path; -OutDir is a folder). Never
                    overwrites (Get-NewPath / New-FileExclusive de-duplicate with _2, _3, ...).

  The rules below are checked by doc-lib.ps1 Build-DocxSection / Build-TxtSection - the same code
  review-file.ps1 -Kind rewrite runs (as a dry run) when it creates the preview, so a preview the user
  is shown is one this script can build.

  Anchors (safety net against silently losing content): within each requested section, EVERY table
  and every paragraph that carries an image, an embedded (OLE) object, hyperlink, field, footnote/
  endnote reference, comment reference, an existing tracked change (w:ins/w:del), or a content
  control (w:sdt) - anywhere in its descendants - is an "anchor". Every anchor in a section (other
  than the section's own heading paragraph, which is governed by the first-block rule below) MUST
  appear as exactly one "keep" block for that same paragraph number, in the same order as the anchors
  occur in the original document. Missing, duplicated, reordered, or out-of-section "keep" blocks
  reject the WHOLE file (not just that section) and name the paragraph(s) at fault - this is the
  code-side guarantee that a rewrite cannot silently drop a table/image/embedded object/link/
  footnote/comment/existing-tracked-change/content control. A "keep" of an ordinary body paragraph
  (not an anchor) is also allowed, unchanged.

  First block of every section's "blocks" must represent that section's own heading: either
  "keep" of the original heading paragraph, or a "heading" block at the SAME level as the original
  heading (only its text may change; the paragraph style is copied from the original heading, not
  from the request). If the section's own heading paragraph is itself an anchor (contains an image/
  embedded object/link/field/footnote/comment/tracked change/content control), this script requires
  "keep" for it - rewriting it via a "heading" block would silently drop that content, which is
  exactly the unsafe case this script refuses rather than improvising past (see docs/v2-plan.md
  Phase 3.6: "ถ้าสเปค... ไม่ปลอดภัย ... ปฏิเสธกรณีที่ไม่ปลอดภัยพร้อมข้อความชัดเจน").

  A later "heading" block (e.g. rewriting an existing sub-heading inside the section) reuses the
  paragraph style of SOME existing paragraph at that same outline level anywhere in the document;
  if the document has no heading at that level anywhere, the whole file is refused (the spec does
  not want a heading level invented from nothing). A "paragraph" block reuses the pPr and first-run
  rPr of its first "from" paragraph (L1: a caption/list/other styled paragraph keeps its own style);
  only when that source is a heading does it fall back to the section's own first ORIGINAL body
  paragraph, and a section with no body paragraph at all then refuses the block with a clear
  message - this script does not guess formatting from nothing.

  Two or more requested sections whose ¶ ranges overlap (e.g. a heading and one of its own
  sub-headings, both given as separate "section" entries) are refused - the ranges are ambiguous to
  replace independently. Outside the requested sections the content and its order are unchanged: in a
  .docx only word/document.xml is rewritten (it is re-serialized when saved, so the paragraphs outside
  the sections keep their content but the XML bytes are not guaranteed identical) and every other zip
  entry is copied through unchanged, like apply-revisions.ps1 does; in a .txt the surrounding text is
  kept as it is.

  "from" (docs/v2-plan.md 3.9 G1): every "heading"/"paragraph" block's "from" array names the
  original ¶ number(s) that block's new text was rewritten from (>= 1 entry each; the same ¶ MAY
  appear in more than one block's "from", for a paragraph split into two new blocks). The section's
  original ¶ range is split into segments by its own anchors (segment 0 = before the section's
  first anchor, segment 1 = between the 1st and 2nd anchor, etc.; the section's own heading is a
  member of segment 0). The whole file is refused, naming the ¶ at fault, unless: (1) all "from" ¶
  of the SAME block fall in the SAME segment; (2) that segment equals the number of anchor "keep"
  blocks already placed earlier in the SAME section's "blocks" list (so a block's position in the
  Plan must match where its source text actually sits relative to the section's anchors - content
  cannot silently move across a table/image/etc.); (3) every original ¶ of the section that is
  neither the heading nor an anchor appears in some block's "from" OR as a "keep" - no ¶ can
  silently vanish from the Plan; (4) every "from" ¶ exists in the section and is NOT itself an
  anchor (anchors can only be cited via "keep").

  Table cells (.docx only): a "table" block (the preview of review-file.ps1 -Kind rewrite embeds it as v3 plan data)
  changes the text of the cells it lists and nothing else. The cell's w:tcPr and the paragraph's w:pPr stay; the text is
  changed word by word like apply-revisions.ps1 does (doc-lib.ps1 Set-ParagraphTextPlain: unchanged words keep their own
  run's formatting, new words take the formatting of the run before them) and the change is accepted on the spot - there
  is no w:ins/w:del. A line break (w:br) or special hyphen in the cell paragraph stays exactly where it is and the script
  prints a line saying so; a change that touches the break or the hyphen is refused with the cell's coordinate and the old
  text with the break marked (the paragraph is never rewritten as a whole, which would lose the break). Any other
  paragraph that cannot be edited word by word is rewritten in the formatting of its first text run and the script prints
  a line saying so. A cell that cannot be changed (continuation of
  a vertically merged cell, nested table, empty, a link / field / picture / footnote / comment / tracked change in the
  paragraph, content control) refuses the whole run. Before writing, the script re-counts every table and compares, per
  table, the properties, grid, rows, cells, each cell's w:tcPr and each paragraph's w:pPr and the text of every cell
  that was not named (doc-lib.ps1 Test-DocxTablesIntact); any difference stops the run without writing a file. As with
  every paragraph, the text of a cell must still be the one in the ticked preview (a cell text edited in the preview file
  refuses the section).

  Images with no text, and blank paragraphs (.docx only): extract-text.ps1 / Read-Docx skip a
  paragraph or table that has no text, so an image-only paragraph or an empty paragraph has NO ¶
  number (extract-text.ps1 -Mode Text prints "[รูป ไม่มีข้อความ]" where such an image sits, without a
  ¶ number). Inside a requested section (from its heading to its last numbered ¶) this script keeps
  them where they are, automatically - the plan does not mention them and cannot "keep" them:
    - a run of such nodes that contains an image/picture/object/field/link/comment/tracked
      change/content control/table/section break (the same list as an anchor) is an IMPLICIT
      ANCHOR. It splits the section into segments exactly like a "keep" anchor does, but it has no ¶
      and no "keep" block. Blocks are placed segment by segment: a block whose "from" lies in a
      LATER segment than the position reached so far gets the picture(s) inserted in front of it,
      automatically, in original order (so a figure stays right before its caption); a block whose
      "from" lies in an EARLIER segment than the position reached is refused ("ย้ายเนื้อหาข้ามรูป
      ไม่ได้", naming the ¶) - new text can never move across a picture. Any pictures left after the
      last block are appended at the end of the section. If a normal "keep" anchor lies between the
      current position and the block's segment, the anchor message above applies instead.
    - a run made only of empty paragraphs is a SPACER: it is re-inserted right before the first new
      node that comes from the original ¶ that followed it (the first block that lists that ¶ in
      "from", or "keep"s it). It is never dropped silently.
    - the same segment rule applies to a plain "keep" of an ordinary paragraph.
  Before writing the file this script re-counts w:drawing, w:pict, w:object and w:tbl over the whole
  document and re-checks the order of each section's heading, "keep" anchors and implicit anchors; any
  difference stops the run without writing a file. The per-section summary line shows the number
  of implicit anchors kept, e.g. "(คง anchor 5 รายการ, รูปไม่มีข้อความ 2 จุด)".
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$Plan,
  [string]$Out,
  [string]$OutDir,
  [string[]]$Review
)

. (Join-Path $PSScriptRoot 'doc-lib.ps1')

# Reads -Plan (optional; see the header) into @( @{ Section; Blocks } ). The per-block parsing lives in
# doc-lib.ps1 ConvertTo-PlanBlocks, shared with the blocks embedded in a preview.
function Load-Plan {
  if (-not (Test-Path -LiteralPath $Plan -PathType Leaf)) { Fail "ไม่พบไฟล์: $Plan" }
  $json = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Plan).ProviderPath, $utf8)
  try { $parsed = ConvertFrom-Json -InputObject $json } catch { Fail 'ไฟล์แผนไม่ใช่ JSON ที่ถูกต้อง' }
  if (-not $parsed) { Fail 'ไฟล์แผนไม่ใช่ JSON ที่ถูกต้อง' }
  $secsRaw = Get-AsArray $parsed.sections
  if (-not $secsRaw.Count) { Fail 'sections ต้องมีอย่างน้อยหนึ่งรายการ' }
  $result = New-Object System.Collections.ArrayList
  $seenSec = New-Object 'System.Collections.Generic.HashSet[string]'
  $si = 0
  foreach ($s in $secsRaw) {
    $si++
    $sid = [string]$s.section
    if ($sid -cnotmatch '^S\d+$') { Fail "sections รายการที่ ${si}: section ต้องเป็นรูปแบบ S3" }
    if ($seenSec.Contains($sid)) { Fail "sections: section $sid ซ้ำ" }
    [void]$seenSec.Add($sid)
    $blocksRaw = Get-AsArray $s.blocks
    if (-not $blocksRaw.Count) { Fail "${sid}: blocks ต้องมีอย่างน้อยหนึ่งรายการ" }
    $blocks = ConvertTo-PlanBlocks $blocksRaw $sid
    [void]$result.Add(@{ Section = $sid; Blocks = $blocks })
  }
  return , $result
}

# "ย่อหน้าใหม่ 2 (จาก ¶9, ¶10)" / "keep ¶6" - names block $idx of a block list for a refusal message.
function Get-BlockLabel($blocks, [int]$idx) {
  $b = $blocks[$idx]
  if ($b.Type -eq 'keep') { return "keep ¶$($b.Paragraph)" }
  if ($b.Type -eq 'table') { return "ตาราง ¶$($b.Paragraph)" }
  $n = 0
  for ($i = 0; $i -le $idx; $i++) { if ($blocks[$i].Type -ne 'keep') { $n++ } }
  $kind = if ($b.Type -eq 'heading') { "ย่อหน้าใหม่ $n (หัวข้อ ระดับ $($b.Level))" } else { "ย่อหน้าใหม่ $n" }
  return "$kind (จาก ¶" + ($b.From -join ', ¶') + ')'
}

# -Plan given: every block must equal the block embedded in the preview the user ticked. Names the
# section and the block (with its ¶) at the first difference.
function Test-PlanMatchesPreview([string]$sid, $planBlocks, $previewBlocks) {
  $tail = 'ให้ใช้ข้อความจาก preview เท่านั้น (ไม่ต้องส่ง -Plan) หรือบอกในแชตเพื่อสร้าง preview ใหม่'
  if ($planBlocks.Count -ne $previewBlocks.Count) {
    Fail "แผนใน -Plan ของ $sid มี $($planBlocks.Count) บล็อก ไม่ตรงกับ preview ที่ผู้ใช้ติ๊ก ($($previewBlocks.Count) บล็อก) $tail"
  }
  for ($i = 0; $i -lt $previewBlocks.Count; $i++) {
    $p = $planBlocks[$i]; $v = $previewBlocks[$i]
    $label = Get-BlockLabel $previewBlocks $i
    $sameShape = ($p.Type -eq $v.Type)
    if ($sameShape -and ($v.Type -eq 'keep' -or $v.Type -eq 'table')) { $sameShape = ($p.Paragraph -eq $v.Paragraph) }
    if ($sameShape -and $v.Type -eq 'table') {
      $sameShape = (@($p.Cells).Count -eq @($v.Cells).Count)
      for ($ci = 0; $sameShape -and $ci -lt @($v.Cells).Count; $ci++) {
        if (-not [string]::Equals([string]$p.Cells[$ci].Cell, [string]$v.Cells[$ci].Cell, [StringComparison]::Ordinal)) { $sameShape = $false }
        elseif (-not [string]::Equals([string]$p.Cells[$ci].Text, [string]$v.Cells[$ci].Text, [StringComparison]::Ordinal)) {
          Fail "แผนใน -Plan ของ $sid ข้อความของ $label เซลล์ $($v.Cells[$ci].Cell) ไม่ตรงกับ preview ที่ผู้ใช้ติ๊ก $tail"
        }
      }
    }
    if ($sameShape -and $v.Type -ne 'keep' -and $v.Type -ne 'table') {
      $sameShape = (($p.From -join ',') -eq ($v.From -join ','))
      if ($sameShape -and $v.Type -eq 'heading') { $sameShape = ($p.Level -eq $v.Level) }
    }
    if (-not $sameShape) { Fail "แผนใน -Plan ของ $sid บล็อกที่ $($i + 1) ($label) ชนิดหรือ ¶ ที่มาไม่ตรงกับ preview ที่ผู้ใช้ติ๊ก $tail" }
    if ($v.Type -ne 'keep' -and $v.Type -ne 'table' -and -not [string]::Equals([string]$p.Text, [string]$v.Text, [StringComparison]::Ordinal)) {
      Fail "แผนใน -Plan ของ $sid ข้อความของ $label ไม่ตรงกับ preview ที่ผู้ใช้ติ๊ก $tail"
    }
  }
}

$full = Resolve-Input $Path
$ext =[System.IO.Path]::GetExtension($full).ToLowerInvariant()
if ($ext -eq '.pdf') { Fail 'ยังไม่รองรับ .pdf สำหรับการ rewrite ทั้ง section ให้แปลงเป็น .docx ก่อน (เปิดใน Word แล้วบันทึกเป็น .docx)' }
if ($ext -ne '.docx' -and $ext -ne '.txt') { Fail "ยังไม่รองรับไฟล์ $ext รองรับ .docx และ .txt" }

# F1 (docs/v2-plan.md 3.8): -Review is mandatory for .docx/.txt (the only extensions reachable
# here - .pdf was already refused above). Read with the SAME parser review-file.ps1 -Action Read
# uses (doc-lib.ps1's Read-ReviewFile), not a second copy.
# G2 (docs/v2-plan.md 3.9): -Review accepts more than one preview file of the SAME source (an
# array, or one comma-separated string, or a mix of both) - each requested section resolves
# independently to whichever supplied preview most recently CREATED a preview containing that
# section (by the "- สร้างเมื่อ:" line every preview records; a tie goes to the file that appears
# LATER in -Review), and that section must be ticked in the file chosen for it.
function Get-ReviewPaths([string[]]$raw) {
  $out = New-Object System.Collections.ArrayList
  foreach ($r in $raw) {
    if (-not $r) { continue }
    foreach ($part in ($r -split ',')) {
      $p = $part.Trim()
      if ($p) { [void]$out.Add($p) }
    }
  }
  return , @($out)
}
if (-not $Review -or $Review.Count -eq 0) { Fail 'ต้องระบุ -Review (ไฟล์ตัวอย่างก่อนสร้างไฟล์ rewrite ที่ผู้ใช้ตรวจและติ๊กแล้ว จาก review-file.ps1 -Action Create -Kind rewrite)' }
$reviewPaths = Get-ReviewPaths $Review
if (-not $reviewPaths.Count) { Fail 'ต้องระบุ -Review (ไฟล์ตัวอย่างก่อนสร้างไฟล์ rewrite ที่ผู้ใช้ตรวจและติ๊กแล้ว จาก review-file.ps1 -Action Create -Kind rewrite)' }
$srcLeaf = Split-Path $full -Leaf
$reviews = New-Object System.Collections.ArrayList
$rIdx = 0
foreach ($rp in $reviewPaths) {
  $rIdx++
  $rv1 = Read-ReviewFile $rp
  if ($rv1.Kind -ne 'rewrite') { Fail 'ไฟล์ -Review ต้องเป็นไฟล์ตัวอย่างก่อนสร้างไฟล์ (rewrite) จาก review-file.ps1 -Kind rewrite เท่านั้น' }
  if (-not $rv1.SourceName -or -not [string]::Equals($rv1.SourceName, $srcLeaf, [StringComparison]::OrdinalIgnoreCase)) {
    Fail "ไฟล์ -Review เป็นของไฟล์ต้นฉบับอื่น ($($rv1.SourceName)) ไม่ตรงกับ -Path ($srcLeaf)"
  }
  $createdAt1 = $null
  if ($rv1.CreatedAt) { try { $createdAt1 = [DateTime]::ParseExact($rv1.CreatedAt, 'yyyy-MM-dd HH:mm', $null) } catch { $createdAt1 = $null } }
  [void]$reviews.Add(@{ Index = $rIdx; Path = $rp; Rv = $rv1; CreatedAt = $createdAt1 })
}

try {
$planFromFile = $null
if ($Plan) { $planFromFile = Load-Plan }

# G2: per section, the preview that decides it = the most recently CREATED file that lists the section
# (a tie favors the file appearing later in -Review).
function Get-ChosenPreview([string]$sid) {
  $best = $null
  foreach ($c in @($reviews | Where-Object { $_.Rv.AllIds -contains $sid })) {
    if (-not $best) { $best = $c; continue }
    $better = $false
    if ($null -ne $c.CreatedAt -and $null -ne $best.CreatedAt) {
      if ($c.CreatedAt -ge $best.CreatedAt) { $better = $true }
    } elseif ($c.Index -ge $best.Index) { $better = $true }
    if ($better) { $best = $c }
  }
  return $best
}

$chosenBySection = @{}
$skippedSections = New-Object System.Collections.ArrayList
if ($planFromFile) {
  # F1/G2 with -Plan: every section requested in -Plan must resolve to SOME supplied preview that
  # contains it, and that preview must have it ticked, or the whole run is refused.
  $requestedIds = @($planFromFile | ForEach-Object { $_.Section })
  $unresolved = New-Object System.Collections.ArrayList
  foreach ($sid in $requestedIds) {
    $best = Get-ChosenPreview $sid
    if (-not $best) { [void]$unresolved.Add($sid); continue }
    $chosenBySection[$sid] = $best
  }
  if ($unresolved.Count) {
    Fail ("ไม่พบ section ต่อไปนี้ในไฟล์ -Review ที่ระบุเลย: " + ($unresolved -join ', '))
  }
  $unticked = @($requestedIds | Where-Object { $chosenBySection[$_].Rv.Ticked -notcontains $_ })
  if ($unticked.Count) {
    Fail ("ไฟล์ -Review ยังไม่ได้ติ๊ก section ต่อไปนี้: " + ($unticked -join ', ') + " ให้ผู้ใช้ตรวจและติ๊กในไฟล์ preview ก่อน แล้วลองใหม่")
  }
} else {
  # No -Plan: build every section that is ticked in the preview chosen for it.
  $allIds = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($c in $reviews) { foreach ($i in $c.Rv.AllIds) { [void]$allIds.Add($i) } }
  $orderedIds = @($allIds | Sort-Object { [int]($_ -replace '^S', '') })
  $requestedIds = New-Object System.Collections.ArrayList
  foreach ($sid in $orderedIds) {
    $best = Get-ChosenPreview $sid
    if ($best.Rv.Ticked -contains $sid) { $chosenBySection[$sid] = $best; [void]$requestedIds.Add($sid) }
    else { [void]$skippedSections.Add($sid) }
  }
  if (-not $requestedIds.Count) {
    Fail ("ไฟล์ -Review ยังไม่ได้ติ๊ก section ใดเลย (ยังไม่ติ๊ก: " + ($skippedSections -join ', ') + ") ให้ผู้ใช้ตรวจและติ๊กในไฟล์ preview ก่อน แล้วลองใหม่")
  }
  $requestedIds = @($requestedIds)
}

# The blocks of every section come from the blocks embedded in its (ticked) preview; the visible
# "หลัง" text of that preview must still match them, or the section is refused.
$planData = @{}
$previewBlocksBySection = @{}
foreach ($sid in $requestedIds) {
  $c = $chosenBySection[$sid]
  if (-not $planData.ContainsKey($c.Index)) { $planData[$c.Index] = Get-PreviewPlanData $c.Rv.BodyLines }
  $dataMap = $planData[$c.Index]
  if (-not $dataMap.ContainsKey($sid)) {
    Fail "preview ของ $sid ไม่มีข้อมูลแผนที่ใช้สร้างไฟล์ (สร้างจากรุ่นเก่า หรือบรรทัดข้อมูลท้ายไฟล์ถูกลบ) ให้บอกในแชตเพื่อสร้าง preview ใหม่"
  }
  if ($null -eq $dataMap[$sid]) {
    Fail "ข้อมูลแผนใน preview ของ $sid อ่านไม่ได้ ให้บอกในแชตเพื่อสร้าง preview ใหม่"
  }
  $blocksHere = $dataMap[$sid]
  if (-not (Test-PreviewAfterMatchesBlocks (Get-PreviewAfterLines $c.Rv.BodyLines $sid) $blocksHere)) {
    Fail "ข้อความใน preview ของ $sid ถูกแก้ในไฟล์ ถ้าต้องการเปลี่ยนข้อความ ให้บอกในแชตเพื่อสร้าง preview ใหม่"
  }
  $previewBlocksBySection[$sid] = $blocksHere
}

# The sections to build: with -Plan its own blocks (checked against the preview at the end), without it
# the preview's blocks.
$planSections = New-Object System.Collections.ArrayList
if ($planFromFile) { foreach ($ps in $planFromFile) { [void]$planSections.Add($ps) } }
else { foreach ($sid in $requestedIds) { [void]$planSections.Add(@{ Section = $sid; Blocks = $previewBlocksBySection[$sid] }) } }

# G2: print, per requested section, which preview (file id) resolved that section.
$previewSummaryLines = New-Object System.Collections.ArrayList
foreach ($sid in $requestedIds) {
  $c = $chosenBySection[$sid]
  [void]$previewSummaryLines.Add(("section {0}: ใช้ preview {1} ({2})" -f $sid, $c.Rv.ShaAttr.Substring(0, 8), (Split-Path $c.Path -Leaf)))
}
# Backward-compatible single line (when every section resolved to the SAME one preview file).
$distinctReviewPaths = @($chosenBySection.Values | ForEach-Object { $_.Path } | Select-Object -Unique)
$legacyPreviewLine = $null
if ($distinctReviewPaths.Count -eq 1) {
  $onlyRv = ($chosenBySection.Values | Select-Object -First 1).Rv
  $legacyPreviewLine = "รหัสไฟล์ preview: $($onlyRv.ShaAttr.Substring(0, 8))"
}
$editedNoteNeeded = [bool]($chosenBySection.Values | Where-Object { $_.Rv.Edited } | Select-Object -First 1)
$summaryLines = New-Object System.Collections.ArrayList

if ($ext -eq '.docx') {
  $items = Read-Docx $full
  $ns = $script:DocxNs
  # Text-less pictures / blank paragraphs between the numbered items (doc-lib.ps1 Get-DocxGapGroups),
  # and whole-document counts taken BEFORE anything is changed, for the final self-check.
  $gapAll = Get-DocxGapGroups $items $ns
  $countsBefore = Get-DocxCounts
  $verifyList = New-Object System.Collections.ArrayList
  # A plan that rewrites table cells ("table" blocks): every table is snapshotted before anything is built, and the
  # snapshot is compared with the built document before the file is written (Test-DocxTablesIntact).
  $wantTableSnap = [bool](@($planSections | ForEach-Object { $_.Blocks } | Where-Object { $_.Type -eq 'table' }).Count)
  $tablesBefore = if ($wantTableSnap) { Get-DocxTableSnapshot } else { $null }
  $tableKeys = New-Object System.Collections.ArrayList
  $tableNoteLines = New-Object System.Collections.ArrayList

  # Resolve every requested section up front (against the ORIGINAL, unmutated document) and check
  # for overlap before mutating anything.
  $resolved = New-Object System.Collections.ArrayList
  foreach ($sec in $planSections) {
    $r = Resolve-Section $items $sec.Section
    [void]$resolved.Add(@{ Section = $sec.Section; Blocks = $sec.Blocks; Head = $r.Head; End = $r.End })
  }
  $byHeadN = @($resolved | Sort-Object { $_.Head.N })
  for ($i = 1; $i -lt $byHeadN.Count; $i++) {
    if ($byHeadN[$i].Head.N -le $byHeadN[$i - 1].End) {
      Fail ("ช่วง section ที่เลือกซ้อนกัน: $($byHeadN[$i - 1].Section) กับ $($byHeadN[$i].Section) ให้เลือกช่วงที่ไม่ทับกัน")
    }
  }

  foreach ($sec in $resolved) {
    $built = Build-DocxSection $items $ns $gapAll $sec
    [void]$verifyList.Add($built.Verify)
    [void]$summaryLines.Add($built.Summary)
    foreach ($tk in $built.TableKeys) { [void]$tableKeys.Add($tk) }
    foreach ($tl in $built.TableNotes) { [void]$tableNoteLines.Add($tl) }
  }

  # Self-check BEFORE any file is written (picture/object/table counts, order of anchors, and - when table cells were
  # rewritten - every table's structure and every cell that was not named).
  Test-DocxRewriteIntact $countsBefore $verifyList
  if ($tablesBefore) { Test-DocxTablesIntact $tablesBefore $tableKeys }
  if ($planFromFile) { foreach ($ps in $planFromFile) { Test-PlanMatchesPreview $ps.Section $ps.Blocks $previewBlocksBySection[$ps.Section] } }

  $dest = Get-NewPath $full 'rewritten' '.docx' $Out $OutDir 'humanized'
  Add-Type -AssemblyName System.IO.Compression
  $in = Open-Shared $full
  try {
    $srcZip = New-Object System.IO.Compression.ZipArchive($in, [System.IO.Compression.ZipArchiveMode]::Read)
    $opened = New-FileExclusive $dest
    $dest = $opened.Path
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

  Write-Output "สร้างไฟล์ใหม่: $dest"
  foreach ($l in $summaryLines) { Write-Output $l }
  foreach ($l in $tableNoteLines) { Write-Output $l }
  foreach ($l in $previewSummaryLines) { Write-Output $l }
  if ($legacyPreviewLine) { Write-Output $legacyPreviewLine }
  if ($skippedSections.Count) { Write-Output ("ไม่ได้สร้าง section ที่ไม่ได้ติ๊ก: " + ($skippedSections -join ', ')) }
  if ($editedNoteNeeded) { Write-Output 'หมายเหตุ: ไฟล์ถูกแก้นอกจากช่องติ๊ก จะใช้ข้อความจากแชตเท่านั้น' }
  Write-Output 'ไฟล์นี้ไม่มี tracked changes เปรียบเทียบกับต้นฉบับได้ด้วย Word Review > Compare หรือ preview .md'
}
else {
  # ---- .txt ----
  $items = Read-Txt $full
  $resolved = New-Object System.Collections.ArrayList
  foreach ($sec in $planSections) {
    $r = Resolve-Section $items $sec.Section
    [void]$resolved.Add(@{ Section = $sec.Section; Blocks = $sec.Blocks; Head = $r.Head; End = $r.End })
  }
  $byHeadN = @($resolved | Sort-Object { $_.Head.N })
  for ($i = 1; $i -lt $byHeadN.Count; $i++) {
    if ($byHeadN[$i].Head.N -le $byHeadN[$i - 1].End) {
      Fail ("ช่วง section ที่เลือกซ้อนกัน: $($byHeadN[$i - 1].Section) กับ $($byHeadN[$i].Section) ให้เลือกช่วงที่ไม่ทับกัน")
    }
  }
  # Blank-line-separated vs single-newline: recompute the same heuristic Read-Txt itself used
  # (doc-lib.ps1 does not expose it), so newly-joined blocks use the document's own convention.
  $sep = if ($script:TxtText -match "\n[ \t]*\n") { "`n`n" } else { "`n" }

  # Process sections back-to-front by text offset so earlier offsets stay valid while later ones
  # are being spliced in (unlike the docx DOM, string offsets shift once a replacement changes
  # length).
  $text = $script:TxtText
  $byHeadNDesc = @($resolved | Sort-Object { $_.Head.N } -Descending)
  foreach ($sec in $byHeadNDesc) {
    $built = Build-TxtSection $items $text $sep $sec
    $text = $text.Substring(0, $built.SectionStart) + $built.NewText + $text.Substring($built.SectionEnd)
    [void]$summaryLines.Add($built.Summary)
  }
  if ($planFromFile) { foreach ($ps in $planFromFile) { Test-PlanMatchesPreview $ps.Section $ps.Blocks $previewBlocksBySection[$ps.Section] } }

  if ($script:TxtCrlf) { $text = $text -replace "`n", "`r`n" }
  $enc = switch ($script:TxtEncoding) { 'utf16' { New-Object System.Text.UnicodeEncoding($false, $true) } 'utf8' { $utf8 } default { New-Object System.Text.UTF8Encoding($true) } }
  $dest = Get-NewPath $full 'rewritten' '.txt' $Out $OutDir 'humanized'
  $opened = New-FileExclusive $dest
  $dest = $opened.Path
  try { $sw = New-Object System.IO.StreamWriter($opened.Stream, $enc); $sw.Write($text); $sw.Dispose() } finally { $opened.Stream.Dispose() }

  Write-Output "สร้างไฟล์ใหม่: $dest"
  foreach ($l in ($summaryLines | Sort-Object)) { Write-Output $l }
  foreach ($l in $previewSummaryLines) { Write-Output $l }
  if ($legacyPreviewLine) { Write-Output $legacyPreviewLine }
  if ($skippedSections.Count) { Write-Output ("ไม่ได้สร้าง section ที่ไม่ได้ติ๊ก: " + ($skippedSections -join ', ')) }
  if ($editedNoteNeeded) { Write-Output 'หมายเหตุ: ไฟล์ถูกแก้นอกจากช่องติ๊ก จะใช้ข้อความจากแชตเท่านั้น' }
  if ($script:TxtEncoding -eq 'cp874') { Write-Output 'หมายเหตุ: ต้นฉบับเป็น Windows-874 ไฟล์ใหม่บันทึกเป็น UTF-8' }
}
} finally { Remove-JsonInput $Plan }
