# Shared preservation-checking logic for check-humanity (see docs/v2-plan.md Phase 1.2 and 3.6
# item 3A-1). Dot-sourced by check-preservation.ps1 and by review-file.ps1 -Kind rewrite (Phase 3).
# No side effects: defines word lists and functions only, does not read/write files, does not call
# Fail/exit, and does not depend on doc-lib.ps1 having been dot-sourced first (callers dot-source
# doc-lib.ps1 themselves for $utf8/Fail/etc.).

$ThaiWords = @('ต้อง', 'ควร', 'อาจ', 'ห้าม', 'ไม่', 'สามารถ')
$EnWords = @('must', 'shall', 'should', 'may', 'might', 'can', 'cannot', 'not', 'never', 'required')
# Case-insensitive lookup of $EnWords, used to keep an obligation/negation word out of the Latin-
# token check so each such change is reported exactly once (under the obligation/negation section
# below), not twice.
$script:EnWordSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($w in $EnWords) { [void]$script:EnWordSet.Add($w) }

# Contraction stems that are real English words on their own (modal/auxiliary/negation stems in
# "isn't", "won't", "don't", "doesn't", "can't", "couldn't", "shouldn't", "wouldn't", "aren't",
# "weren't", "hasn't", "haven't", "didn't"). Without this list, Get-LatinTokens would report the
# stem itself (e.g. "isn", "won", "don") as a Latin token that appeared/disappeared whenever a
# negative contraction is added or removed, on top of (or instead of) the intended "not"/negation
# accounting done via the "\w+n't" match in Test-PairPreservation - i.e. a spurious, meaningless
# token flagged only because a word happens to end right before "n't". Matched case-insensitively.
$script:ContractionStems = @('isn', 'won', 'don', 'doesn', 'can', 'couldn', 'shouldn', 'wouldn', 'aren', 'weren', 'hasn', 'haven', 'didn')
$script:ContractionStemSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($w in $script:ContractionStems) { [void]$script:ContractionStemSet.Add($w) }

# A closed list of common Thai unit words. Deliberately NOT an open Thai-letter character class:
# Thai script has no reliable space between words within a clause, so "180 วันของโครงการ" would
# otherwise let the unit group swallow the rest of the sentence. A word not on this list (or not
# Latin) is simply not attached as a unit - the number itself is still extracted and compared.
$script:ThaiUnitWords = @('วัน', 'ปี', 'เดือน', 'บาท', 'ชั่วโมง', 'นาที', 'ครั้ง', 'คน', 'ข้อ',
  'กิโลกรัม', 'กิโลเมตร', 'เมตร', 'ตัว', 'หน้า', 'ชิ้น', 'รายการ', 'สัปดาห์', 'ดอลลาร์')
$script:UnitAlt = ($script:ThaiUnitWords | ForEach-Object { [regex]::Escape($_) }) -join '|'

# Number tokens: digit run (Arabic 0-9 or Thai ๐-๙, allowing embedded , or . as thousand/decimal
# separators), an optional %, and an optional following unit word (a known Thai unit word from
# $script:ThaiUnitWords or a run of Latin letters) - e.g. "1,250,000 บาท", "15%", "180 วัน",
# "๑๕ วัน", "30 days". Thai and Arabic digit forms are compared as literal text, not converted to a
# common value: a rewrite that swaps ๑๕ for 15 is reported as one token missing and a different one
# added, which is intentional (the policy is not to silently equate/convert numeral scripts - see
# docs/quality-plan.md). A digit-grouping change such as "1,500" -> "1500" is likewise still
# flagged (the comma is part of the number's literal text) - that is a deliberate choice, not a bug
# (see docs/eval/metrics.md limitations): it is a formatting change worth a human glance.
#
# Whitespace BETWEEN the number and its attached % or unit is normalized away before building the
# token (any amount of space/tab, or none, between number/% and unit; any amount of space, or
# none, before %): "99.5%" and "99.5 %" both produce the token "99.5%". Only the digits themselves
# (and any embedded , or .) are compared verbatim.
function Get-NumberTokens([string]$text) {
  $re = [regex]("(?<num>[0-9๐-๙]+(?:[.,][0-9๐-๙]+)*)(?:[ \t]*(?<pct>%))?(?:[ \t]+(?<unit>$($script:UnitAlt)|[A-Za-z]+))?")
  $tokens = New-Object System.Collections.ArrayList
  foreach ($m in $re.Matches($text)) {
    $t = $m.Groups['num'].Value
    if ($m.Groups['pct'].Success) { $t += '%' }
    if ($m.Groups['unit'].Success) { $t += ' ' + $m.Groups['unit'].Value }
    [void]$tokens.Add($t)
  }
  return $tokens
}

# Latin tokens: runs of Latin letters/digits starting with a letter (names, acronyms, English
# terms). Compared case-sensitively (Ordinal), per docs/v2-plan.md Phase 1.2 - so a bare case
# change (e.g. "ERP" -> "erp") still shows up as one token missing + one different token added.
# Three kinds of matches are excluded here so each real change is reported exactly once, under the
# obligation/negation section instead of also appearing as a Latin-token change:
#   - an entry from $EnWords (must/shall/should/may/might/can/cannot/not/never/required),
#     matched case-insensitively via $script:EnWordSet;
#   - a contraction suffix fragment: the run of letters right after an apostrophe (e.g. the "t" in
#     "can't", the "s" in "it's") is never a meaningful standalone token on its own;
#   - a contraction STEM immediately followed by "'t" (e.g. "isn" in "isn't", "won" in "won't",
#     "don" in "don't") when the stem is one of $script:ContractionStemSet - see the comment there.
#   - (H2) with -ExcludeSentenceStart (only meaningful together with -Restricted): a token kept
#     solely by the sentence-initial branch of Test-RestrictedLatinToken (see Test-SoftSentenceStart).
function Get-LatinTokens([string]$text, [switch]$Restricted, [switch]$ExcludeSentenceStart) {
  $re = [regex]'[A-Za-z][A-Za-z0-9]*'
  $tokens = New-Object System.Collections.ArrayList
  foreach ($m in $re.Matches($text)) {
    if ($m.Index -gt 0 -and $text[$m.Index - 1] -eq "'") { continue }
    if ($script:EnWordSet.Contains($m.Value)) { continue }
    $followedByNt = ($m.Index + $m.Length + 1 -lt $text.Length) -and
      ($text[$m.Index + $m.Length] -eq "'") -and
      ([char]::ToLowerInvariant($text[$m.Index + $m.Length + 1]) -eq 't') -and
      $script:ContractionStemSet.Contains($m.Value)
    if ($followedByNt) { continue }
    if ($Restricted -and -not (Test-RestrictedLatinToken $m.Value $text $m.Index)) { continue }
    if ($Restricted -and $ExcludeSentenceStart -and (Test-SoftSentenceStart $m.Value $text $m.Index)) { continue }
    [void]$tokens.Add($m.Value)
  }
  return $tokens
}

# F3 (docs/v2-plan.md 3.8): true when $text has more Latin letters than Thai letters - the
# preservation checker then restricts Get-LatinTokens (see Test-RestrictedLatinToken below) instead
# of treating every ordinary English word as a Latin token to preserve verbatim. Only the BEFORE
# side of a pair/section decides this (Get-PreservationResult below), applied to both sides.
function Test-LatinDominant([string]$text) {
  $latin = ([regex]::Matches($text, '[A-Za-z]')).Count
  $thai = ([regex]::Matches($text, '[฀-๾]')).Count
  return ($latin -gt $thai)
}

# True when the character at $text[$idx] begins a new sentence: either it is the first non-
# whitespace/quote/bracket character in the whole text, or the nearest earlier non-whitespace/
# quote/bracket character is a sentence terminator (. ! ?). Used only by the restricted Latin-token
# filter below to tell a sentence-initial capitalized word (e.g. "The"/"It") from a capitalized
# word appearing mid-sentence (e.g. a name).
function Test-SentenceStart([string]$text, [int]$idx) {
  $i = $idx - 1
  while ($i -ge 0 -and $text[$i] -match '[\s"''`\)\]]') { $i-- }
  if ($i -lt 0) { return $true }
  return [bool]($text[$i] -match '[.!?]')
}

# G4 (docs/v2-plan.md 3.9): a short, closed stoplist of common English function/starter words that
# a sentence-initial capitalized word is EXEMPT from tracking when it matches (case-sensitively -
# these are always the canonical capitalized forms, since they only ever reach this branch at a
# sentence start). Anything NOT on this list - most notably a name like "Jane" - is now tracked
# even at the start of a sentence (see Test-RestrictedLatinToken below): before G4, EVERY sentence-
# initial capitalized word was dropped regardless, which silently let a dropped/added name at the
# very start of a sentence go unflagged.
$script:SentenceStartStopWords = @(
  'The', 'This', 'That', 'These', 'Those', 'It', 'Its', 'If', 'In', 'On', 'At', 'For', 'To',
  'We', 'Our', 'You', 'Your', 'They', 'Their', 'A', 'An', 'And', 'But', 'Or', 'So', 'As',
  'When', 'While', 'After', 'Before', 'All', 'Each', 'Please', 'There', 'Here',
  'Should', 'Must', 'May', 'Can'
)
$script:SentenceStartStopSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
foreach ($w in $script:SentenceStartStopWords) { [void]$script:SentenceStartStopSet.Add($w) }

# F3 restricted Latin-token filter, used only when the before-text of a pair/section is Latin-
# dominant (Test-LatinDominant). A token is kept (still checked for preservation) only if it is:
#   - an all-caps acronym, two letters or more (e.g. "QA", "ERP"),
#   - a token containing a digit (e.g. an ID or model number),
#   - a word with an internal capital that is not its first character (e.g. "iPhone", "PowerShell"),
#   - a word whose first character is a capital AND that is NOT the first word of its sentence
#     (e.g. a name like "Carter" appearing mid-sentence),
#   - (G4) a word whose first character is a capital AND that IS the first word of its sentence,
#     UNLESS it is one of the common function/starter words in $script:SentenceStartStopSet (e.g.
#     "The", "It", "This") - so a sentence-initial NAME (e.g. "Jane will approve.") is tracked just
#     like a mid-sentence one, and the result does not depend on the punctuation immediately before
#     the word (". Jane" and "; Jane" are treated identically - Test-SentenceStart decides only
#     whether ". "/"! "/"? " precedes it, and either way a non-stoplist capitalized word is kept).
# Everything else - the bulk of ordinary lowercase English wording, and a sentence-initial STOPLIST
# word like "The"/"It" - is dropped, so a same-meaning reword using common words no longer
# manufactures a Latin-token diff on every English rewrite (dry-run finding F3). Numbers and
# obligation/negation word counts are computed separately in Get-PreservationResult and are never
# affected by this restriction.
function Test-RestrictedLatinToken([string]$token, [string]$text, [int]$idx) {
  if ($token -cmatch '^[A-Z]{2,}$') { return $true }
  if ($token -match '[0-9]') { return $true }
  if ($token.Length -gt 1 -and ($token.Substring(1) -cmatch '[A-Z]')) { return $true }
  if ($token -cmatch '^[A-Z]') {
    if (Test-SentenceStart $text $idx) { return -not $script:SentenceStartStopSet.Contains($token) }
    return $true
  }
  return $false
}

# H2 (dry-run round 3): true when a restricted-mode token is kept ONLY because it is a capitalized,
# non-stoplist word at a sentence start (e.g. "Installation", "Approval") - not an acronym, not a
# token with a digit, not a word with an internal capital. A rewrite that starts a new sentence
# naturally capitalizes an ordinary word there, so such a token on the AFTER side is not reported
# as "added" (Get-PreservationResult below). It is still counted normally on both sides for the
# "missing" direction, so a sentence-initial name that is dropped (G4, e.g. "Jane will approve.")
# stays flagged. Accepted trade-off: a brand-new name added at a sentence start is not flagged.
function Test-SoftSentenceStart([string]$token, [string]$text, [int]$idx) {
  if ($token -cmatch '^[A-Z]{2,}$') { return $false }
  if ($token -match '[0-9]') { return $false }
  if ($token.Length -gt 1 -and ($token.Substring(1) -cmatch '[A-Z]')) { return $false }
  if ($token -cnotmatch '^[A-Z]') { return $false }
  return (Test-SentenceStart $text $idx)
}

function New-CountMap([string[]]$items) {
  $map = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
  foreach ($it in $items) {
    if ($map.ContainsKey($it)) { $map[$it] = $map[$it] + 1 } else { $map[$it] = 1 }
  }
  return $map
}

# Multiset diff: how many more times each token appears in one side than the other.
function Get-CountDiff($beforeMap, $afterMap) {
  $missing = New-Object System.Collections.ArrayList
  $added = New-Object System.Collections.ArrayList
  foreach ($k in $beforeMap.Keys) {
    $b = $beforeMap[$k]; $a = 0; [void]$afterMap.TryGetValue($k, [ref]$a)
    for ($i = 0; $i -lt ($b - $a); $i++) { [void]$missing.Add($k) }
  }
  foreach ($k in $afterMap.Keys) {
    $a = $afterMap[$k]; $b = 0; [void]$beforeMap.TryGetValue($k, [ref]$b)
    for ($i = 0; $i -lt ($a - $b); $i++) { [void]$added.Add($k) }
  }
  return @{ Missing = $missing; Added = $added }
}

# Punctuation whose loss can change meaning: "/" (and/or, e.g. create/view/edit) and paired parentheses
# (an exception, an example, an abbreviation or an explanation). Only the COUNT is compared; a long dash
# or a colon that merely separates sentences is deliberately not tracked. Text that is not a slash in
# the "or/and" sense is left out of the slash count so it never raises a false warning when it stays
# put: a web address (http://..., www....) and digit-slash-digit runs (dates such as 30/09/2026, fractions).
function Get-SlashCount([string]$text) {
  $t = [regex]::Replace($text, '(?i)\b(?:https?|ftp)://[A-Za-z0-9\-._~:/?#\[\]@!$&''()*+,;=%]+', '')
  $t = [regex]::Replace($t, '(?i)\bwww\.[A-Za-z0-9\-._~:/?#\[\]@!$&''()*+,;=%]+', '')
  $t = [regex]::Replace($t, '[0-9๐-๙]+(?:/[0-9๐-๙]+)+', '')
  return ([regex]::Matches($t, '/')).Count
}

# Number of matched "( ... )" pairs (full-width brackets counted as the ASCII ones). A ")" with no "("
# before it, such as a list marker "1)", is not a pair; nesting is handled.
function Get-ParenPairCount([string]$text) {
  $depth = 0; $pairs = 0
  foreach ($ch in $text.ToCharArray()) {
    if ($ch -eq '(' -or $ch -eq [char]0xFF08) { $depth++ }
    elseif (($ch -eq ')' -or $ch -eq [char]0xFF09) -and $depth -gt 0) { $depth--; $pairs++ }
  }
  return $pairs
}

# Thai inter-word spacing: one "gap" is a run of one or more spaces / tabs / no-break spaces (U+00A0)
# whose character before AND after are both Thai (U+0E00-U+0E7F). A run counts once. A space next to a
# Latin letter, a digit or a symbol is not a gap, and neither is a line break.
function Get-ThaiGapCount([string]$text) {
  if (-not $text) { return 0 }
  return ([regex]::Matches($text, '(?<=[฀-๿])[ \t ]+(?=[฀-๿])')).Count
}

# Compares Thai gaps per NEW paragraph (rewrite preview only), never per section: a section-wide total
# would let a removed gap in one paragraph hide behind an added gap in another. $sourceByN maps a source
# paragraph number to its text; $blocks are the non-keep "after" entries in order (hashtables with
# Text and From = paragraph numbers) - the position of an entry in this list is its "ย่อหน้าใหม่ N".
# Entries that share the same "from" list (one paragraph split into several) are compared as ONE group,
# once. Before = sum of the gaps of each source paragraph in "from" (counted one by one, never after
# joining the texts); after = the group's gaps. When "from" has k paragraphs (a merge) the after count may
# exceed the before count by up to k-1 without a warning (the seam between the old paragraphs); when the
# group has m new paragraphs (a split) the after count may be lower by up to m-1 (a space that became the
# paragraph break); any other difference, up or down, is reported. Returns the warning lines (wrap the call in @() to always get an
# array; it is empty when nothing differs).
function Get-ThaiGapWarnings($sourceByN, $blocks) {
  $groups = New-Object System.Collections.ArrayList
  $index = @{}
  $no = 0
  foreach ($b in $blocks) {
    $no++
    $fromNs = @($b.From | ForEach-Object { [int]$_ })
    $key = ($fromNs -join ',')
    if (-not $index.ContainsKey($key)) {
      $index[$key] = @{ From = $fromNs; Nos = (New-Object System.Collections.ArrayList); After = 0 }
      [void]$groups.Add($index[$key])
    }
    [void]$index[$key].Nos.Add($no)
    $index[$key].After += (Get-ThaiGapCount ([string]$b.Text))
  }
  $out = New-Object System.Collections.ArrayList
  foreach ($g in $groups) {
    $before = 0
    foreach ($n in $g.From) { if ($sourceByN.ContainsKey($n)) { $before += (Get-ThaiGapCount ([string]$sourceByN[$n])) } }
    $seams = [Math]::Max(0, $g.From.Count - 1)
    $splits = [Math]::Max(0, $g.Nos.Count - 1)
    if ($g.After -lt ($before - $splits) -or $g.After -gt ($before + $seams)) {
      $newLabel = 'ย่อหน้าใหม่ ' + (($g.Nos | ForEach-Object { [string]$_ }) -join ', ')
      $fromLabel = (($g.From | ForEach-Object { "¶$_" }) -join ', ')
      [void]$out.Add("ช่องว่างระหว่างคำไทยเปลี่ยน: $newLabel (จาก $fromLabel): $before -> $($g.After) (บอกเหตุผลว่าอ่านดีขึ้นอย่างไร ถ้าไม่มีเหตุผลให้คงช่องว่างเดิม)")
    }
  }
  return @($out)
}

# Computes the full preservation result for one before/after pair, without printing anything.
# Returns @{ NumDiff; LatDiff; Changed = [string[]] (one line per obligation/negation count that
# changed, same wording as check-preservation.ps1's console output); Punct = [string[]] (one line per
# punctuation count that changed - only filled with -CheckPunctuation); Notes = [string[]] (the caller's
# -InfoNotes); HasIssue = [bool] }.
# -CheckPunctuation (used by review-file.ps1 -Kind rewrite): also compare the count of "/" and of
# "( )" pairs between before and after. A change is a signal for a person to look, never a refusal.
# -InfoNotes: ready-made note lines computed by the caller (review-file.ps1 -Kind rewrite passes the
# Thai-gap lines of Get-ThaiGapWarnings). They are shown in the report as notes but are NOT a difference
# that needs a person: they never make HasIssue true.
function Get-PreservationResult([string]$before, [string]$after, [switch]$CheckPunctuation, [string[]]$InfoNotes = @()) {
  $numDiff = Get-CountDiff (New-CountMap (Get-NumberTokens $before)) (New-CountMap (Get-NumberTokens $after))
  # F3: the BEFORE side alone decides whether this pair/section is Latin-dominant; the same
  # restriction (or lack of it) is then applied to both sides so the multiset diff stays meaningful.
  $restricted = Test-LatinDominant $before
  $latBefore = New-CountMap (Get-LatinTokens $before -Restricted:$restricted)
  $latDiff = Get-CountDiff $latBefore (New-CountMap (Get-LatinTokens $after -Restricted:$restricted))
  if ($restricted) {
    # H2: "added" counts only after-side occurrences that are not soft sentence-initial words
    # (Test-SoftSentenceStart), so it can never exceed the plain diff above; "missing" is unchanged.
    $afterFirm = New-CountMap (Get-LatinTokens $after -Restricted -ExcludeSentenceStart)
    $added = New-Object System.Collections.ArrayList
    foreach ($k in $afterFirm.Keys) {
      $b = 0; [void]$latBefore.TryGetValue($k, [ref]$b)
      for ($i = 0; $i -lt ($afterFirm[$k] - $b); $i++) { [void]$added.Add($k) }
    }
    $latDiff.Added = $added
  }

  $changed = New-Object System.Collections.ArrayList
  foreach ($w in $ThaiWords) {
    $b = ([regex]::Matches($before, [regex]::Escape($w))).Count
    $a = ([regex]::Matches($after, [regex]::Escape($w))).Count
    if ($a -ne $b) { [void]$changed.Add("$w (ไทย): $b -> $a") }
  }
  foreach ($w in $EnWords) {
    $opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    # (?!') excludes a modal word that is really the start of a contraction (e.g. "can" inside
    # "can't"), which is counted separately below as a negation contraction.
    $pattern = "\b$w\b(?!')"
    $b = ([regex]::Matches($before, $pattern, $opt)).Count
    $a = ([regex]::Matches($after, $pattern, $opt)).Count
    if ($a -ne $b) { [void]$changed.Add("$w (EN): $b -> $a") }
  }
  $opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
  $ctB = ([regex]::Matches($before, "\b\w+n't\b", $opt)).Count
  $ctA = ([regex]::Matches($after, "\b\w+n't\b", $opt)).Count
  if ($ctA -ne $ctB) { [void]$changed.Add("contraction ปฏิเสธ (n't): $ctB -> $ctA") }

  $punct = New-Object System.Collections.ArrayList
  if ($CheckPunctuation) {
    $slB = Get-SlashCount $before; $slA = Get-SlashCount $after
    if ($slA -ne $slB) { [void]$punct.Add("เครื่องหมาย / เปลี่ยนจำนวน: $slB -> $slA (ตรวจว่าความหมายเท่าเดิมและแจ้งผู้ใช้)") }
    $pnB = Get-ParenPairCount $before; $pnA = Get-ParenPairCount $after
    if ($pnA -ne $pnB) { [void]$punct.Add("วงเล็บคู่ ( ) เปลี่ยนจำนวน: $pnB -> $pnA (ตรวจว่าความหมายเท่าเดิมและแจ้งผู้ใช้)") }
  }

  $notes = New-Object System.Collections.ArrayList
  foreach ($note in @($InfoNotes)) { if ($note) { [void]$notes.Add($note) } }

  $hasIssue = [bool]($numDiff.Missing.Count -or $numDiff.Added.Count -or $latDiff.Missing.Count -or $latDiff.Added.Count -or $changed.Count -or $punct.Count)
  return @{ NumDiff = $numDiff; LatDiff = $latDiff; Changed = @($changed); Punct = @($punct); Notes = @($notes); HasIssue = $hasIssue }
}

# Formats an already-computed Get-PreservationResult as the console report lines this checker has
# always used (check-preservation.ps1's -Source/-Revised and -Pairs modes; review-file.ps1
# -Kind rewrite's per-section preservation block). Returns a [string[]] rather than calling
# Write-Output itself: a caller that does `$x = Format-PreservationLines ...` must still get only
# the report text back, not also swallow it as PowerShell function "output" the way it would if
# this also printed AND returned a value.
function Format-PreservationLines([string]$id, $result) {
  $lines = New-Object System.Collections.ArrayList
  $numDiff = $result.NumDiff; $latDiff = $result.LatDiff
  [void]$lines.Add("=== $id ===")
  if ($numDiff.Missing.Count) { [void]$lines.Add("  ตัวเลข/หน่วย/% หายไป: " + ($numDiff.Missing -join ', ')) }
  if ($numDiff.Added.Count) { [void]$lines.Add("  ตัวเลข/หน่วย/% เพิ่มมา: " + ($numDiff.Added -join ', ')) }
  if ($latDiff.Missing.Count) { [void]$lines.Add("  คำละติน/ชื่อ/ตัวย่อ หายไป: " + ($latDiff.Missing -join ', ')) }
  if ($latDiff.Added.Count) { [void]$lines.Add("  คำละติน/ชื่อ/ตัวย่อ เพิ่มมา: " + ($latDiff.Added -join ', ')) }
  foreach ($c in $result.Changed) { [void]$lines.Add("  คำผูกมัด/ปฏิเสธเปลี่ยนจำนวน: $c") }
  foreach ($p in @($result.Punct)) { [void]$lines.Add("  $p") }
  # notes (e.g. a changed number of Thai inter-word spaces) are shown but do not change the status below
  foreach ($n in @($result.Notes)) { [void]$lines.Add("  หมายเหตุ (ไม่นับเป็นความต่างที่ต้องให้คนตรวจ): $n") }
  $status = if ($result.HasIssue) { 'ต้องให้คนตรวจ' } else { 'ไม่พบความต่าง' }
  [void]$lines.Add("  สถานะ: $status")
  return , @($lines)
}
