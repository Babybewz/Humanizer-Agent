<#
  Print-only preservation checker for check-humanity rewrites (see docs/v2-plan.md Phase 1.2 and
  docs/eval/metrics.md). Never creates a file. Compares "before" text with "after" text and
  reports numbers/units/%, Latin tokens (names/acronyms/English terms, case-sensitive), and
  obligation/negation word counts that changed. Ships in the skill from V2 onward but is not yet
  called from SKILL.md/file-input.md - that wiring is Phase 3 (rewrite preview). Usable standalone
  today for eval (docs/eval/metrics.md).

  Mode A: -Source <file> -Revised <file>   (two UTF-8 text files, compared as a single pair)
  Mode B: -Pairs <json file>                (UTF-8 JSON: [{"id":"...","before":"...","after":"..."}])
  Exactly one of the two modes must be used.

  Input files are deleted after the run (success or error) ONLY when they resolve to somewhere
  under the system temp folder - same rule as doc-lib.ps1's Remove-JsonInput, reused here rather
  than reimplemented; a -Source/-Revised file outside temp (e.g. a real project file) is left
  alone even though it is never modified.

  Findings are printed to the console (UTF-8) only. Overall status per pair is "ไม่พบความต่าง" or
  "ต้องให้คนตรวจ" - never "ผิด": Thai keyword counting is a substring heuristic (e.g. "ถูกต้อง"
  contains "ต้อง") so a flagged pair is a signal for a human to look, not a verdict.

  Exit code: 0 whenever the run completed, regardless of findings. Non-zero (2, via doc-lib.ps1's
  Fail) only for usage/argument errors or unreadable/invalid input.

  The token-extraction/diff/report logic lives in preservation-lib.ps1 (dot-sourced below, no side
  effects) so review-file.ps1 -Kind rewrite (Phase 3) can compute the same code-checked result
  without shelling out to this script. This script's own CLI behaviour is unchanged.
#>
param(
  [string]$Source,
  [string]$Revised,
  [string]$Pairs
)

. (Join-Path $PSScriptRoot 'doc-lib.ps1')
. (Join-Path $PSScriptRoot 'preservation-lib.ps1')

# Same temp-folder restriction as doc-lib.ps1's Remove-JsonInput (Test-UnderTempDir, reused as-is)
# but not limited to a .json extension, since -Source/-Revised are plain text files.
function Remove-TempInput([string]$path) {
  if (-not $path) { return }
  try {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
    $full = (Resolve-Path -LiteralPath $path).ProviderPath
    $inTemp = (Test-UnderTempDir $full ([System.IO.Path]::GetTempPath())) -or
      (Test-UnderTempDir $full $env:TEMP) -or (Test-UnderTempDir $full $env:TMP)
    if (-not $inTemp) { return }
    Remove-Item -LiteralPath $full -Force -ErrorAction Stop
  } catch { }
}

function Read-Utf8Text([string]$path) {
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { Fail "ไม่พบไฟล์: $path" }
  $full = (Resolve-Path -LiteralPath $path).ProviderPath
  $bytes = [System.IO.File]::ReadAllBytes($full)
  $text = $utf8.GetString($bytes)
  if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
  return $text
}

$script:PairCount = 0
$script:IssueCount = 0

# Thin wrapper: delegates the actual token/diff computation to preservation-lib.ps1's
# Get-PreservationResult and the console line formatting to its Format-PreservationLines (see that
# file for the wording/heuristics), and keeps this script's own running totals for the summary.
function Test-PairPreservation([string]$id, [string]$before, [string]$after) {
  $script:PairCount++
  $result = Get-PreservationResult $before $after
  if ($result.HasIssue) { $script:IssueCount++ }
  foreach ($line in (Format-PreservationLines $id $result)) { Write-Output $line }
}

$usesSR = [bool]$Source -or [bool]$Revised
$usesPairs = [bool]$Pairs
if ($usesSR -and $usesPairs) { Fail 'ใช้ -Source กับ -Revised หรือ -Pairs อย่างใดอย่างหนึ่งเท่านั้น ไม่ใช้พร้อมกัน' }
if (-not $usesSR -and -not $usesPairs) { Fail 'ต้องระบุ -Source กับ -Revised หรือ -Pairs' }
if ($usesSR -and (-not $Source -or -not $Revised)) { Fail 'โหมด -Source/-Revised ต้องระบุทั้งสองพารามิเตอร์' }

if ($usesPairs) {
  try {
    if (-not (Test-Path -LiteralPath $Pairs -PathType Leaf)) { Fail "ไม่พบไฟล์: $Pairs" }
    try {
      $jsonText = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Pairs).ProviderPath, $utf8)
      $parsed = ConvertFrom-Json -InputObject $jsonText
    } catch { Fail 'ไฟล์ pairs ไม่ใช่ JSON ที่ถูกต้อง' }
    if ($null -eq $parsed) { Fail 'ไฟล์ pairs ไม่ใช่ JSON ที่ถูกต้อง' }
    $list = New-Object System.Collections.ArrayList
    foreach ($x in @($parsed)) { [void]$list.Add($x) }
    if (-not $list.Count) { Fail 'ไม่มีรายการใน pairs' }
    $i = 0
    foreach ($it in $list) {
      $i++
      $id = if ($it.id) { [string]$it.id } else { "#$i" }
      if ($null -eq $it.before -or $null -eq $it.after) { Fail "${id}: ต้องมี before และ after" }
      Test-PairPreservation $id ([string]$it.before) ([string]$it.after)
    }
  } finally { Remove-JsonInput $Pairs }
} else {
  try {
    $beforeText = Read-Utf8Text $Source
    $afterText = Read-Utf8Text $Revised
    Test-PairPreservation 'ทั้งไฟล์' $beforeText $afterText
  } finally { Remove-TempInput $Source; Remove-TempInput $Revised }
}

Write-Output ''
Write-Output ("สรุป: ตรวจ {0} รายการ, ต้องให้คนตรวจ {1} รายการ, ไม่พบความต่าง {2} รายการ" -f $script:PairCount, $script:IssueCount, ($script:PairCount - $script:IssueCount))
Write-Output 'หมายเหตุ: การนับคำไทย (ต้อง/ควร/อาจ/ห้าม/ไม่/สามารถ) เป็นการนับ substring โดยประมาณ อาจนับผิดจากคำที่มีคำเหล่านี้ซ้อนอยู่ (เช่น "ถูกต้อง" มีคำว่า "ต้อง") ผลลัพธ์นี้เป็นสัญญาณให้คนตรวจซ้ำ ไม่ใช่การตัดสินว่าผิดหรือถูก'
exit 0
