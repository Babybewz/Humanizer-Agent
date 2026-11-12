<#
  Create or read a review/report/brief/rewrite .md file for check-humanity (see references/file-input.md
  §6 and docs/v2-plan.md Phase 3.6 items 3A-2/3A-3). Never modifies the source document; Create only
  checks that it exists (and, for -Kind review/report, that it is not opened).

  -Action Create -Path <source> -Items <json file> -Kind review|report [-OutDir <dir>]
    -Items is a UTF-8 JSON file:
      { "range": "...", "mode": "...",
        "items": [ { "id":"R1", "paragraph":9, "row":2, "cell":1, "cell_paragraph":2,
                     "heading":"...", "category":"must|may|confirm", "before":"...",
                     "after":"...", "reason":"..." } ] }
    Use "paragraph" or "page" (one of them); "row"/"cell"/"cell_paragraph" only for a table cell
    (coordinates match extract-text.ps1 -Cells). "after" may be "" only for category "confirm",
    and only in -Kind review (a report allows "" for any category).
    Where each item sits is written for a person, not as a bare ¶ number (Word does not show ¶ numbers, which count after
    the table of contents). For a .docx / .txt (-Path is read once, never modified) the item heading is the heading path
    by the numbers the document shows ("### R1 — บทที่ 1 › 1.3.3", a table: "... › ตาราง “ตารางที่ 3” › แถว “...” ·
    คอลัมน์ “...”"), followed by "- ค้นใน Word (Ctrl+F): `text`" (.txt: "ค้นในไฟล์"): a short text of the paragraph that is
    found once in the whole document (table of contents included; "(พบ N ที่ ...)" is added when it cannot be made
    unique) and "- อ้างอิง: ¶42" (table cell: "¶824 r4c4"), the ¶ / rRcC reference that extract-text.ps1 prints.
    The path and the text are computed from the file itself (doc-lib.ps1 Get-ItemFinder); "heading" in the JSON is not
    used for a .docx / .txt (it is still accepted). A "paragraph" / row / cell that is not in the file stops the run.
    When an item's "before" is not in the paragraph it names, the start of the paragraph is used as the text and stdout
    says so ("หมายเหตุ: ข้อความ “ก่อน” ของ R5 ไม่พบ ...") and so does the item's Ctrl+F line ("(ข้อความต้นย่อหน้า ไม่มีข้อความ “ก่อน” ของรายการนี้)"). A .pdf keeps "หน้า N · <heading>" as the item heading and gets
    only "- ค้นใน PDF (Ctrl+F): `text`" (the first words of "before"); no อ้างอิง line.
    The file's header gets one line "- หาจุดในไฟล์: ..." (.docx / .txt) saying how to use the two.
    Optional groups (review and report): top level "groups":[{"id":"A","title":"...","note":"..."}] (id = 1-3 chars,
    ^[A-Z][A-Z0-9]{0,2}$, unique; title required; note optional) and "group":"A" on every item. With groups every item
    needs a known group, every group needs an item, and "group" without "groups" is refused; "groups":[] = no groups.
    The script writes "## สรุปตามกลุ่ม" (a table in the groups order, part of the hashed body) before "## รายการ", puts
    the items under "### กลุ่ม A — <title>" headings (item headings become "#### R1 — ...") and orders the trailer's
    items="..." as they appear in the file. -Action Read then prints one "กลุ่ม A (<title>): R1, R2" line per group.
    Without groups the file and stdout are exactly as before.

  -Action Create -Path <source> -Items <json file> -Kind brief [-OutDir <dir>]
    -Items is a UTF-8 JSON file (see docs/v2-plan.md 3.6 3A-2 and 3.8 F5):
      { "range":"S3", "mode":"...",
        "brief": { "purpose":"...", "audience":"...", "language":"...", "relations":["..."],
                   "sections":[ { "section":"S3", "heading":"...", "points":["..."] } ],
                   "preserve":["..."], "structure_issues":["..."], "style_notes":["..."],
                   "confirm":["...", {"section":"S3","text":"..."}] } }
    Writes <name>_brief_<date>.md. No checkboxes. "style_notes" (array of strings) is OPTIONAL: the
    linguist role's observations (passages that read as AI-like, with paragraph numbers and
    reasons); it renders under "## ข้อสังเกตด้านสำนวน" right after "ปัญหาระดับโครงสร้าง", and a
    missing key or an empty array renders "ไม่มี" like the other optional fields. "language" (string) and "relations" (array) are
    REQUIRED keys (F5) - a missing key is refused; every text field/array may still be "" or empty,
    which renders as "ไม่มี" under its own heading ("ภาษาของเอกสาร" / "ความเชื่อมโยงระหว่าง section")
    - the script never invents content for a blank field, but it also never lets the model skip
    stating language/relations outright.
    brief.tables is OPTIONAL (.docx only): the table cells the author proposes to rewrite -
        "tables":[ { "section":"S2", "table":"¶7", "cells":["r4c2","r5c2p1"], "reason":"..." } ]
    Every entry is checked against the source: "section" must be one of brief.sections, "table" must be a table in
    that section, and every coordinate (the ones extract-text.ps1 -Cells prints; pK when a cell has several
    paragraphs) must be a cell that can be changed - it is refused, with the reason, when it is the continuation of a
    vertically merged cell, holds a nested table, is empty, or its paragraph has a link / field / picture / footnote /
    comment / tracked change or sits in a content control. The brief then shows "## ข้อความในตารางที่จะเขียนใหม่": per
    table its ¶ and section, the reason, and per cell the coordinate, the column head (row 1) and the row head (first
    cell of the row) when there are any, and the current text. A section that has only headings/tables is NOT skipped
    when it holds a listed table. The list is also written as one machine-readable line
        <!-- hz-tables v1 data="<base64 of the UTF-8 JSON of the list>" -->
    inside the part the trailer's sha256 covers; the preview step (-Kind rewrite) accepts a table cell only if the
    latest brief of the same file in the same folder lists it.
    brief.confirm (the points the author must confirm) takes each point either as a plain string (the point names no
    section) or as {"section":"S3","text":"..."} (the point belongs to that section, which must be one of
    brief.sections; an empty text, a missing section or a section that is not in the brief is refused with the reason).
    The brief shows a point that names its section as "- S3: <text>" and the others as "- <text>".
    Every brief also carries one machine-readable line
        <!-- hz-brief v2 data="<base64 of the UTF-8 JSON {"sections":["S1",...],"confirm":[{"section":"S3","text":"..."}]}>" -->
    (also inside the sha256): the section ids it covers and its confirm points ("section":"" = the point names none;
    blank ones dropped). The preview step (-Kind rewrite) reads it back (version 1, written by earlier versions, held the
    points as plain strings and is still read: every point of it names no section); a brief without the line or edited in
    the file enforces nothing there.
    stdout, after the summary lines, ends with the text the user must be shown in the chat, between
        ----- ข้อความ brief สำหรับแสดงในแชต (แสดงตามนี้ครบทุกบรรทัด ห้ามย่อหรือสรุปใหม่) -----
        ----- จบข้อความ brief -----
    That block is the brief as the file has it (every heading in the file's order, the lines under them, the table
    lines, the skip labels, the "ข้อความในตารางที่จะเขียนใหม่" part) without the data lines (hz-...), the sha256 trailer
    and the file-storage note. The script itself adds to the block: when brief.tables lists cells, a line saying that
    any cell can be dropped from the list before the user confirms the brief (answer "แก้ brief"); and the closing
    question "ถูกต้อง / แก้ brief / ยกเลิก". The author copies the block as it is (references/full-rewrite.md step 2).
    -Action Read on a brief file is refused with a clear message (brief has no checkboxes to read).
    When -Path is a .docx or .txt, every brief.sections[].section is resolved against the source (a section
    that is not there is refused) and the script adds, by itself, from the source's own structure:
      - for a section that holds no paragraph a rewrite could change (only headings / tables / pictures): the
        line "ข้าม: ไม่มีข้อความให้เขียนใหม่ (มีแต่หัวข้อ/ตาราง/รูป) คงไว้ตามเดิม" under it, and the section is listed
        on stdout ("ข้าม (ไม่มีข้อความให้เขียนใหม่): S1, S11")
      - for a .docx section with tables: the line "ตารางใน section นี้ N ตาราง ไม่ถูกเขียนใหม่ ..." under it (N =
        the tables of the section that brief.tables does not name; for the ones it names the line says the table is
        rewritten only in the cells listed), and one summary line in the header when two or more listed sections have
        tables
      - for a .docx section whose ¶ range (sub-sections included) holds paragraphs that are anchors (a link, picture,
        embedded (OLE) object, field, footnote, comment, earlier tracked change or content control - the same
        Test-AnchorItem definition the preview and the file build use; the section's own heading counts, a table does
        not): the line "ย่อหน้าที่เขียนทับไม่ได้ (มีลิงก์/รูป/ฟิลด์ ฯลฯ) จะคงไว้ตามเดิม: ¶15 (ลิงก์), ¶18 (ลิงก์)" under it,
        in the file and in the chat block. No such paragraph (and a .txt) = no line. The model never writes it.
      - on stdout, for a top-level .docx section (its heading has the shallowest level of any heading in the
        document) longer than doc-lib.ps1 $script:LargeSectionParagraphs (60) ¶ that can be split by its
        sub-headings (it has none of its own text before the first one): a note naming them. A section that
        is itself a sub-heading never gets the note (the split is one layer only)

  -Action Create -Path <source> -Items <json file> -Kind rewrite [-OutDir <dir>]
    The preview is the ONE source of the new text: write-rewrite.ps1 builds the file from the blocks this
    script stores in the preview (no second copy of the text is sent) and refuses a section whose visible
    text was edited in the file. -Path must be a .docx or .txt (it is read for the ¶ numbers, the texts and
    the anchors - never modified).
    -Items is a UTF-8 JSON file (see docs/v2-plan.md 3.6 3A-3 and 3.8 F2/F6):
      { "range":"S3,S4", "mode":"...",
        "sections":[ { "section":"S3", "heading":"...",
                       "before":[{"id":"¶6","keep":true}, {"id":"¶7"}, {"id":"¶8","keep":true}, {"id":"¶9"}],
                       "after":[ {"keep":"¶6"},
                                 {"text":"ย่อหน้าใหม่ ...","from":["¶7"]},
                                 {"keep":"¶8"},
                                 {"heading":2,"text":"หัวข้อที่แก้ข้อความ","from":["¶9"]} ],
                       "model_check":["..."], "confirm":["..."], "from_chat":["..."] } ] }
    "before" lists EVERY ¶ of the section once, in the source's order ("id":"¶N"); "text" is not needed - if
    given it is ignored, the text always comes from -Path's own document (read with the same extractor as
    extract-text.ps1; a table renders as its rows joined by newline, cells within a row joined by " | ").
    "keep":true marks a ¶ that stays exactly as it is: an anchor (table, image, embedded (OLE) object, link,
    field, footnote/endnote/comment reference, tracked change, content control), an unchanged heading, or
    an ordinary paragraph left alone. Every "after" entry is an object, one of
      {"text":"...","from":["¶9","¶10"]}              a new paragraph rewritten from those ¶
      {"heading":2,"text":"...","from":["¶6"]}       a heading whose text is rewritten (level = the level it has now)
      {"keep":"¶7"}                                   the ¶ kept (must be a before item marked keep:true)
      {"table":"¶7","cells":[{"cell":"r4c2","text":"..."}]}   a table kept in place whose listed cells get new text
    A "table" entry counts as the keep of that table in the keep/order rules (same place, never missing or moved);
    "cells" empty = the table is just kept. Each cell text is ONE paragraph (no line break), not empty, different from
    the current one, and the cell must be one that can be changed (same rules as brief.tables). A cell is accepted
    ONLY if the latest brief of this file (<name>_brief_*.md, same folder as the preview) lists it in its hz-tables line
    and that brief file is not edited (its sha256 still matches): otherwise the whole run is refused and no file is
    created. The preview shows, under the table's "[คงไว้: ¶7 ...]" line, "[ตาราง ¶7: เขียนใหม่ N เซลล์ เซลล์อื่นคงเดิม]"
    and per cell its coordinate, context (column head / row head), the text now ("ก่อน:") and the new text ("หลัง:").
    The preservation check compares each cell alone (numbers, Latin words, obligation/negation words, "/", "( )" and
    the Thai spaces of that cell) and lists the result per cell; a section with rewritten cells is never pre-ticked.
    The plan data line is v3 when the section has a "table" entry (v2 otherwise).
    A plain string entry (the old format) is refused with a message that shows these three shapes. Every
    before item marked keep:true must appear in "after" as {"keep":id} EXACTLY once, in the same relative
    order - otherwise the whole file is refused, naming the mismatch. A ¶ marked keep cannot also be a
    "from" of a new paragraph. Blank "text" entries are dropped.
    The same rules write-rewrite.ps1 enforces are checked here, while the preview is created, with the same
    code (doc-lib.ps1 Build-DocxSection / Build-TxtSection, run as a dry run on an in-memory copy that is
    never saved): the first entry is the section's own heading, every anchor is kept in order, every "from"
    is in the section, is not an anchor and sits on the right side of any anchor or text-less picture, no ¶
    is left out. A violation is refused with the same message write-rewrite.ps1 would give, so the user is
    never shown a preview that cannot be turned into a file.
    "map" is IGNORED if sent (one note line is printed). The script builds the map "¶ เดิม → ย่อหน้าใหม่ N"
    itself from the "from" lists; N counts every entry that is not a keep, exactly as the "ย่อหน้าใหม่ N:"
    headings do (a heading entry is shown as "ย่อหน้าใหม่ N (หัวข้อ ระดับ L):").
    Both Before and After render a kept anchor/heading as "[คงไว้: ¶id <short excerpt>]" and a kept
    ordinary paragraph as "[คงเดิม: ¶id]" followed by its full text in a quote block, so the user can review
    what stays. Kept items are EXCLUDED from the code-computed preservation comparison (unchanged by
    definition). Whenever a section has a kept anchor (table, image, ...; not a kept heading or an ordinary
    paragraph), the script itself adds the line "ตาราง/ส่วนที่คงไว้ (anchor) อยู่ตำแหน่งเดิม
    เนื้อหาย้ายข้ามส่วนนี้ไม่ได้ในรุ่นนี้" under that section (F6) - the model never has to remember to say it.
    The script computes the preservation result itself, per section, from the section's own
    source text and the new text via preservation-lib.ps1 (Get-PreservationResult) - any "preservation" /
    "code_check" field supplied in the JSON is ignored entirely, because the code-computed result
    must never be overwritten by the model (code-computed results are authoritative). Besides numbers, Latin terms and
    obligation/negation words, the result also warns when the count of "/" or of "( )" pairs differs
    between before and after (a changed count means the section is not pre-ticked; it never refuses the
    file, and long dashes / colons are not counted). "/" inside a web address or between digits (dates)
    is not counted. It also warns, per NEW paragraph, when the number of spaces between Thai words
    (a run of spaces / tabs / no-break spaces with a Thai letter on both sides; a space next to a Latin
    letter, a digit or a symbol is not counted) differs from the source paragraph(s) it was written from:
    "ช่องว่างระหว่างคำไทยเปลี่ยน: ย่อหน้าใหม่ 2 (จาก ¶7): 3 -> 1 (...)". New paragraphs written from the
    same "from" list (one paragraph split in two) are compared together as one group; a merge of k
    paragraphs may gain up to k-1 spaces (the seams) and a split into m may lose up to m-1 (the space that
    became a break) without a warning; any other change, up or down, is
    reported. Unlike the other warnings it is only a NOTE: it is printed under the result as
    "หมายเหตุ (ไม่นับเป็นความต่างที่ต้องให้คนตรวจ): ช่องว่างระหว่างคำไทยเปลี่ยน: ..." but it does not make the
    section "ต้องให้คนตรวจ", is not counted in the issue count and does not stop the pre-tick; the file is
    never refused (the language rules allow removing or adding a space when it really reads better and the
    reason is given). The same holds for the Thai spaces of a rewritten table cell. "model_check" is printed under its own
    "ทวนโดยโมเดล (ไม่ใช่ผลตรวจด้วยโค้ด)" heading, clearly separate from that code-computed result.
    Each section gets exactly one checkbox ("ใช้ <section>"), pre-ticked (D-R2) only when the
    code-computed result found no difference AND that section's "confirm" list is empty; otherwise
    left unticked. I1: also never pre-ticked when the section has a non-empty optional "from_chat"
    list (facts the user gave in chat that are not in the source; rendered under their own
    heading) or when an earlier rewrite preview of the same source in the same folder listed that
    section with an open "จุดต้องยืนยัน" item - such a section is printed on a
    "ไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจและติ๊กเอง: S2 (<reason>)" console line (the reason
    distinguishes a chat answer from an earlier preview with an open confirm point).
    A section that holds no paragraph a rewrite could change (only headings / tables / pictures; see
    doc-lib.ps1 Get-SectionProfile) and whose "after" only keeps things as they are is refused, naming it,
    with the message that it can be skipped: there is no text to review, so no preview is made. (A plan that
    really rewords a heading of such a section is not refused: that is a change to review.)
    Embedded plan data: after the last section, one line per section
      <!-- hz-plan v2 section="S3" data="<base64 of the gzip of the section's blocks as UTF-8 JSON>" -->
    holds the blocks (the same shape as a write-rewrite.ps1 -Plan section: keep / heading / paragraph with
    "from"). The lines are part of the file body, so the trailer's sha256 covers them like everything else.
    A v1 line (the same base64 without gzip, written by earlier versions) is still read by write-rewrite.ps1.
    -Action Read on a rewrite file prints the file id, which sections are ticked, and a note if the
    file was edited outside the checkboxes (same contract as -Kind review).
    The preview step also looks at the latest brief of the same source file in the same folder
    (Get-LatestBriefInfo; the same lookup as for table cells) and acts on it:
      - no brief at all: the whole run is refused, nothing is created ("ไม่พบ brief ของไฟล์ต้นฉบับนี้ในโฟลเดอร์เดียวกับ
        preview ... สร้าง brief ก่อนด้วย -Kind brief"). The brief is a mandatory step; a user who says to skip it
        only skips WAITING for the user's confirmation of it, never making it. When -OutDir is used, the brief and the
        preview must be written to the same folder.
      - the brief spans the WHOLE file and the preview holds more than one section: the whole run is refused, nothing is
        created ("ขอบเขตทั้งไฟล์ต้องสร้าง preview ทีละ section (ส่งมา 3 section: S2, S4, S5) ให้สร้าง preview ของ S2 ก่อน ...").
        "Whole file" is computed from the source, never from the free-text "range": every unit of a whole-file rewrite
        that has text to rewrite (doc-lib.ps1 Get-WholeFileUnits: the top-level headings, a large one that can be split
        replaced by its sub-headings) lies inside the ¶ range of a section listed in the brief. Several sections that are
        not the whole file still make one preview; a brief without the hz-brief line enforces nothing.
      - the brief has confirm points that the preview did not carry: a point that names its section (hz-brief v2) is judged
        per section - a section of the preview that the point concerns (the same section, or one whose ¶ range overlaps it,
        e.g. a sub-heading of a large section) is not pre-ticked while its own "confirm" is empty, and the
        other sections of the preview are not affected; a point without a section (a plain string, or any point of a
        version 1 brief) is judged for the whole preview: when no section of it has a "confirm" entry,
        no section is pre-ticked. "from_chat" never settles a point (the script cannot tell whether an entry answers it;
        an answer such as "correct" to the brief itself does not): a section that has "from_chat" but an empty
        "confirm" is still listed, the header and the stdout then say to check whether each point got a real answer
        in the chat or is still open, and the reason line names both causes. A section left unticked this way is printed on stdout ("ไม่ติ๊กไว้ล่วงหน้า ...: S2 (brief
        มีจุดต้องยืนยันที่ preview ไม่ได้ยกมา)") and the points are listed in the preview under "## จุดต้องยืนยันจาก brief
        (preview นี้ไม่ได้ยกมา)". The points printed on stdout and listed in the preview are the ones that concern a section of
        the preview (points of its sections + points without a section); whenever there are any they are printed on stdout.
        Keeping the source's own wording never settles a point; a point stays in "confirm" until the user has answered its
        content in chat, and only then does the answer go into "from_chat".
        The script cannot tell which section a point without a section belongs to, so those are compared at the level of
        the whole preview only.

  -Action Read -Review <review.md|rewrite.md>
    -Kind review: prints the R-ids, which are ticked, the file's last-write time, and any notes
    (unknown ids, or the file edited outside the checkbox marks).
    -Kind rewrite: prints the same, but ids are section ids (e.g. S3) instead of R-ids.
    -Kind report / brief: refused with a clear message (neither has checkboxes to read).

  Output files are UTF-8 with BOM, CRLF line endings, and never overwrite an existing file
  (CreateNew). See doc-lib.ps1 Get-NewPath for the default/-OutDir/O2-guard rules.
#>
param(
  [Parameter(Mandatory = $true)][ValidateSet('Create', 'Read')][string]$Action,
  [string]$Path,
  [string]$Items,
  [ValidateSet('review', 'report', 'brief', 'rewrite')][string]$Kind = 'review',
  [string]$OutDir,
  [string]$Review
)

. (Join-Path $PSScriptRoot 'doc-lib.ps1')
. (Join-Path $PSScriptRoot 'preservation-lib.ps1')

# Get-QuotedLines / Get-OneLine / Get-ShortText / Get-AsArray now live in doc-lib.ps1 (shared with the
# rewrite preview code there). Windows PowerShell 5.1's ConvertFrom-Json returns a bare scalar/object
# (not a 1-element array) for a JSON array with a single element, so every place below that reads a
# JSON array goes through Get-AsArray and assigns the result to a variable before enumerating it.

function Get-NonEmpty([string]$text) { return [bool]($text -and $text.Trim()) }

# -Kind rewrite reads the SOURCE document (-Path) with the same extractor extract-text.ps1 uses
# (doc-lib.ps1 Read-Docx/Read-Txt): the text of every "before" item, the text of a kept item and
# whether it is an anchor all come from the source, never from the JSON (G5, docs/v2-plan.md 3.9), and
# the structure of every section's new blocks is checked with the very code write-rewrite.ps1 builds
# the file with (doc-lib.ps1 Build-DocxSection / Build-TxtSection), run on an in-memory copy that is
# never saved - so a preview the user is shown is one that can be built.
#
# State = the items of one Read-Docx/Read-Txt call. Building a section mutates the in-memory copy, so a
# later section whose ¶ range overlaps an earlier one gets a fresh copy (Reset-RewriteSourceState).
#
# When the plan rewrites table cells ($script:WantTableSnap), the state also holds a snapshot of every table taken
# before anything is built (TableSnap) and the keys of the cell paragraphs built so far (TableKeys), for the
# structure self-check of the dry run (doc-lib.ps1 Test-DocxTablesIntact).
function New-RewriteSourceState {
  if ($script:SrcExt -eq '.docx') {
    $its = Read-Docx $srcFull
    $ns = $script:DocxNs
    $snap = if ($script:WantTableSnap) { Get-DocxTableSnapshot } else { $null }
    return @{ Items = $its; Ns = $ns; Gap = (Get-DocxGapGroups $its $ns); Counts = (Get-DocxCounts); Sep = $null; Ranges = (New-Object System.Collections.ArrayList)
              TableSnap = $snap; TableKeys = (New-Object System.Collections.ArrayList) }
  }
  $its = Read-Txt $srcFull
  $sepTxt = if ($script:TxtText -match "\n[ \t]*\n") { "`n`n" } else { "`n" }
  return @{ Items = $its; Ns = $null; Gap = @(); Counts = $null; Sep = $sepTxt; Ranges = (New-Object System.Collections.ArrayList); TableSnap = $null; TableKeys = (New-Object System.Collections.ArrayList) }
}

# The source item for ¶ number $n as @{ Text; IsAnchor; IsHeading } (never the JSON's text), or Fails
# naming the id when the source has no such ¶.
function Get-SourceItemInfo($state, [string]$sid, [string]$bid, [int]$n) {
  $srcItem = $state.Items | Where-Object { $_.N -eq $n } | Select-Object -First 1
  if (-not $srcItem) { Fail "${sid}: before ${bid}: ไม่พบ ¶$n ในไฟล์ต้นฉบับ" }
  $isAnchor = if ($state.Ns) { Test-AnchorItem $srcItem $state.Ns } else { $false }
  $isHeading = ($srcItem.Kind -eq 'p' -and $srcItem.Level -gt 0)
  return @{ Text = $srcItem.Text; IsAnchor = $isAnchor; IsHeading = $isHeading }
}

# ¶ id in a "from" list: "¶9" or a plain 9. Returns the number or $null.
function ConvertTo-ParagraphNumber($v) {
  if ($v -is [string]) { if ($v -match '^¶(\d+)$') { return [int]$Matches[1] }; return $null }
  $n = $v -as [int]
  if ($n -and $n -ge 1) { return $n }
  return $null
}


# Renders a text field or a list of strings for brief/rewrite output: "ไม่มี" when empty, per the
# instruction that an empty field means "none" and is never guessed/invented.
function Format-TextOrNone([string]$text) {
  $t = Get-OneLine $text
  if (-not $t.Trim()) { return 'ไม่มี' }
  return $t
}
function Format-ListOrNone($items) {
  $arr = Get-AsArray $items
  $vals = @($arr | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
  if (-not $vals.Count) { return @('ไม่มี') }
  return @($vals | ForEach-Object { "- $_" })
}

# brief.confirm items: a plain string (the point belongs to no named section) or an object {"section":"S2","text":"..."}
# (the point belongs to that section, which must be one of brief.sections). Returns a list of @{ Section = '' | 'S2'; Text }.
# Blank plain strings are dropped as before; an object with an empty text, a missing/empty section or a section that is not
# in the brief is refused with the reason.
function Get-BriefConfirmItems($raw, $sectionIds) {
  $out = New-Object System.Collections.ArrayList
  $n = 0
  foreach ($c in (Get-AsArray $raw)) {
    $n++
    if ($null -eq $c) { continue }
    if ($c -isnot [System.Management.Automation.PSCustomObject]) {
      $t = (Get-OneLine ([string]$c)).Trim()
      if ($t) { [void]$out.Add(@{ Section = ''; Text = $t }) }
      continue
    }
    $sec = ([string]$c.section).Trim()
    if (-not $sec) { Fail "brief.confirm รายการที่ ${n}: ต้องมี section (ถ้าจุดนี้ไม่ได้เป็นของ section ใด ให้ใส่เป็นข้อความล้วนแทน)" }
    if (-not $sectionIds.Contains($sec)) { Fail "brief.confirm รายการที่ ${n}: section `"$sec`" ไม่อยู่ใน brief.sections" }
    $t = (Get-OneLine ([string]$c.text)).Trim()
    if (-not $t) { Fail "brief.confirm รายการที่ ${n}: text ว่าง" }
    [void]$out.Add(@{ Section = $sec; Text = $t })
  }
  return , $out
}
# "S2: text" for a point that names its section, the text alone otherwise (the form shown in the brief and on stdout)
function Format-ConfirmPoint($point) {
  if ($point.Section) { return "$($point.Section): $($point.Text)" }
  return [string]$point.Text
}
# Does a brief point that names section $BriefSid concern the preview section $PreviewSid? True for the same section, and
# when the ¶ ranges of the two overlap (a large section and a sub-heading of it). $HeadRanges maps "S<n>" -> @{ Start; End }.
function Test-SectionRangesTouch($HeadRanges, [string]$BriefSid, [string]$PreviewSid) {
  if ($BriefSid -ceq $PreviewSid) { return $true }
  if (-not $HeadRanges.ContainsKey($BriefSid) -or -not $HeadRanges.ContainsKey($PreviewSid)) { return $false }
  return ($HeadRanges[$BriefSid].Start -le $HeadRanges[$PreviewSid].End -and $HeadRanges[$PreviewSid].Start -le $HeadRanges[$BriefSid].End)
}

# I1 (dry-run round 4): section ids that an EARLIER rewrite preview of the same source, in the same
# folder, listed with at least one open "จุดต้องยืนยัน" item. A section regenerated after such a
# preview is never pre-ticked: the user answered (or dropped) a confirm point, so the new text must
# be checked and ticked by the user, even when the code finds no difference and "confirm" is now
# empty. Reads the preview markdown as written by this script (### <id> headings, the
# **จุดต้องยืนยัน** block that is either "ไม่มี" or "- " lines). A file whose trailer is not
# kind="rewrite" or whose source line names another file is ignored.
function Get-PriorOpenConfirmSections([string]$dir, [string]$srcName) {
  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) { return , $set }
  $srcLine = "- ไฟล์ต้นฉบับ: $srcName (ไม่ได้แก้ไฟล์นี้)"
  foreach ($f in (Get-ChildItem -LiteralPath $dir -Filter '*_rewrite_*.md' -File -ErrorAction SilentlyContinue)) {
    try { $text = [System.IO.File]::ReadAllText($f.FullName) } catch { continue }
    if ($text -notmatch 'hz-review v1 kind="rewrite"') { continue }
    $fileLines = $text -split "`r?`n"
    if ($fileLines -notcontains $srcLine) { continue }
    $cur = $null; $inConfirm = $false
    foreach ($l in $fileLines) {
      if ($l -match '^### (\S+)') { $cur = $Matches[1]; $inConfirm = $false; continue }
      if ($l -eq '**จุดต้องยืนยัน**') { $inConfirm = $true; continue }
      if ($inConfirm) {
        if (-not $l.Trim()) { $inConfirm = $false; continue }
        if ($cur -and $l.StartsWith('- ')) { [void]$set.Add($cur) }
      }
    }
  }
  return , $set
}

# The latest brief of a file: among the brief files (<name>_brief_*.md, kind "brief") of the source named $srcName in
# folder $dir, the one with the newest own "- สร้างเมื่อ:" time (then the file's write time). The preview step
# (-Kind rewrite) reads three things from it: which table cells the user confirmed (hz-tables), which sections it covers
# and which confirm points it still carries (hz-brief).
# Returns $null when there is none, else @{ Path; Name; Edited = <bool>; Allowed = HashSet of "<table ¶>|<cell>";
# Info = $null | @{ Sections = <ids>; Confirm = <points> } }.
# Edited is $true when the file no longer matches its own sha256 trailer or one of its data lines cannot be decoded -
# the lists in it are then not trusted (Allowed stays empty, Info is $null). The brief carries the lists as machine-
# readable lines "<!-- hz-tables v1 data="<base64 of the UTF-8 JSON>" -->" (only when it lists table cells) and
# "<!-- hz-brief v2 data="<base64 of the UTF-8 JSON>" -->" inside the part the trailer's sha256 covers (v1, from earlier
# versions, is still read: its confirm points are plain texts, i.e. name no section). A brief written before the hz-brief
# line existed has Info = $null: nothing that depends on it is enforced for such a brief.
# Info.Confirm is a list of @{ Section = '' | 'S2'; Text }.
function Get-LatestBriefInfo([string]$dir, [string]$srcName) {
  if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) { return $null }
  $srcLine = "- ไฟล์ต้นฉบับ: $srcName (ไม่ได้แก้ไฟล์นี้)"
  $best = $null
  foreach ($f in (Get-ChildItem -LiteralPath $dir -Filter '*_brief_*.md' -File -ErrorAction SilentlyContinue)) {
    try { $text = [System.IO.File]::ReadAllText($f.FullName) } catch { continue }
    if ($text -notmatch 'hz-review v1 kind="brief"') { continue }
    $normText = ($text -replace "`r`n", "`n")
    if ($normText.Length -gt 0 -and $normText[0] -eq [char]0xFEFF) { $normText = $normText.Substring(1) }
    $fileLines = $normText -split "`n"
    if ($fileLines -notcontains $srcLine) { continue }
    $created = $null
    foreach ($l in $fileLines) {
      if ($l -match '^-\s*สร้างเมื่อ:\s*(.+?)\s*$') { try { $created = [DateTime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm', $null) } catch { $created = $null }; break }
    }
    $cand = @{ File = $f; Lines = $fileLines; Created = $created; Write = $f.LastWriteTimeUtc }
    $better = $false
    if (-not $best) { $better = $true }
    else {
      $a = if ($null -ne $cand.Created) { $cand.Created } else { [DateTime]::MinValue }
      $b = if ($null -ne $best.Created) { $best.Created } else { [DateTime]::MinValue }
      if ($a -gt $b -or ($a -eq $b -and $cand.Write -ge $best.Write)) { $better = $true }
    }
    if ($better) { $best = $cand }
  }
  if (-not $best) { return $null }
  $allowed = New-Object 'System.Collections.Generic.HashSet[string]'
  $info = $null
  $edited = $false
  $lines = $best.Lines
  $tIdx = -1
  for ($ti = $lines.Count - 1; $ti -ge 0; $ti--) { if ($lines[$ti] -match '^<!--\s*hz-review\s') { $tIdx = $ti; break } }
  $trailerMatch = if ($tIdx -ge 0) { [regex]::Match($lines[$tIdx], '^<!--\s*hz-review\s+v1\s+kind="brief"\s+items="[^"]*"\s+sha256="([0-9a-f]{64})"\s*-->\s*$') } else { $null }
  if (-not $trailerMatch -or -not $trailerMatch.Success) { $edited = $true }
  else {
    $shaAttr = $trailerMatch.Groups[1].Value
    $bodyLines = if ($tIdx -gt 0) { $lines[0..($tIdx - 1)] } else { @() }
    if (-not [string]::Equals((Get-BodyHash ($bodyLines -join "`n")), $shaAttr, [StringComparison]::OrdinalIgnoreCase)) { $edited = $true }
    else {
      $seenTablesLine = $false; $seenBriefLine = $false
      foreach ($l in $bodyLines) {
        if (-not $seenTablesLine -and $l -match '^<!--\s*hz-tables\s+v1\s+data="([A-Za-z0-9+/=]*)"\s*-->\s*$') {
          $seenTablesLine = $true
          try {
            $obj = ConvertFrom-Json -InputObject ($utf8.GetString([Convert]::FromBase64String($Matches[1])))
            foreach ($t in (Get-AsArray $obj.tables)) {
              foreach ($c in (Get-AsArray $t.cells)) { [void]$allowed.Add("$([int]$t.table)|$([string]$c)") }
            }
          } catch { $edited = $true; $allowed.Clear() }
        } elseif (-not $seenBriefLine -and $l -match '^<!--\s*hz-brief\s+v([12])\s+data="([A-Za-z0-9+/=]*)"\s*-->\s*$') {
          $seenBriefLine = $true
          $briefDataVer = [int]$Matches[1]
          try {
            $obj = ConvertFrom-Json -InputObject ($utf8.GetString([Convert]::FromBase64String($Matches[2])))
            # v1: confirm is a list of texts (no section). v2: a list of {"section":"S2"|"","text":"..."}; a bare text in v2 has no section.
            $confirmPoints = New-Object System.Collections.ArrayList
            foreach ($cp in (Get-AsArray $obj.confirm)) {
              if ($briefDataVer -ge 2 -and $cp -is [System.Management.Automation.PSCustomObject]) { $cpSec = ([string]$cp.section).Trim(); $cpText = (Get-OneLine ([string]$cp.text)).Trim() }
              else { $cpSec = ''; $cpText = (Get-OneLine ([string]$cp)).Trim() }
              if ($cpText) { [void]$confirmPoints.Add(@{ Section = $cpSec; Text = $cpText }) }
            }
            $info = @{
              Sections = @((Get-AsArray $obj.sections) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
              Confirm  = @($confirmPoints)
            }
          } catch { $edited = $true; $info = $null }
        }
      }
    }
  }
  if ($edited) { $info = $null }
  return @{ Path = $best.File.FullName; Name = $best.File.Name; Edited = $edited; Allowed = $allowed; Info = $info }
}

# The Thai labels of what makes a .docx paragraph an anchor, read from the very nodes Test-AnchorItem looks for
# ($script:AnchorXPath, plus a content control around the paragraph), in a fixed order. Only used to word the line that
# tells the user which paragraphs a rewrite cannot overwrite; whether a paragraph IS an anchor is always Test-AnchorItem.
function Get-AnchorKindLabels($item, $ns) {
  $found = New-Object 'System.Collections.Generic.HashSet[string]'
  $node = $item.Node
  if ($node.SelectSingleNode('ancestor::w:sdt', $ns)) { [void]$found.Add('content control') }
  foreach ($n in $node.SelectNodes($script:AnchorXPath, $ns)) {
    switch ($n.LocalName) {
      'hyperlink' { [void]$found.Add('ลิงก์') }
      { $_ -in @('drawing', 'pict') } { [void]$found.Add('รูป') }
      'object' { [void]$found.Add('วัตถุฝัง (OLE)') }
      { $_ -in @('fldSimple', 'fldChar') } { [void]$found.Add('ฟิลด์') }
      { $_ -in @('footnoteReference', 'endnoteReference') } { [void]$found.Add('เชิงอรรถ') }
      { $_ -in @('commentReference', 'commentRangeStart', 'commentRangeEnd') } { [void]$found.Add('comment') }
      { $_ -in @('ins', 'del') } { [void]$found.Add('tracked changes เดิม') }
      'sdt' { [void]$found.Add('content control') }
    }
  }
  $order = @('ลิงก์', 'รูป', 'วัตถุฝัง (OLE)', 'ฟิลด์', 'เชิงอรรถ', 'comment', 'tracked changes เดิม', 'content control')
  return @($order | Where-Object { $found.Contains($_) })
}

# Brief (.docx): the line that names, for one section, the paragraphs in its ¶ range (the heading and the sub-sections
# included) that a rewrite cannot overwrite because they are anchors - the same definition the preview and the file build
# use (Test-AnchorItem) - so the user knows before confirming the brief which paragraphs stay as they are. Tables are not
# listed (the brief has its own table lines). $null when the section has none or the source is a .txt ($ns = $null).
function Get-BriefAnchorLine($items, $ns, [string]$sid) {
  if (-not $ns) { return $null }
  $r = Resolve-Section $items $sid
  $parts = New-Object System.Collections.ArrayList
  foreach ($it in $items) {
    if ($it.N -lt $r.Head.N -or $it.N -gt $r.End) { continue }
    if ($it.Kind -ne 'p') { continue }
    if (-not (Test-AnchorItem $it $ns)) { continue }
    $labels = @(Get-AnchorKindLabels $it $ns)
    if ($labels.Count) { [void]$parts.Add("¶$($it.N) ($($labels -join ', '))") } else { [void]$parts.Add("¶$($it.N)") }
  }
  if (-not $parts.Count) { return $null }
  return ('ย่อหน้าที่เขียนทับไม่ได้ (มีลิงก์/รูป/ฟิลด์ ฯลฯ) จะคงไว้ตามเดิม: ' + ($parts -join ', '))
}

function Get-ItemLocation($it) {
  if ($null -ne $it.paragraph) {
    $loc = "¶$($it.paragraph)"
    if ($null -ne $it.row -and $null -ne $it.cell) {
      $loc += " ตาราง แถว $($it.row) เซลล์ $($it.cell)"
      if ($null -ne $it.cell_paragraph) { $loc += " ย่อหน้าในเซลล์ $($it.cell_paragraph)" }
    }
    return $loc
  } elseif ($null -ne $it.page) {
    return "หน้า $($it.page)"
  }
  return $null
}

# .pdf: the Ctrl+F text is the first words of "before" (about 40 characters, cut at a space or at a text-element boundary so no
# Thai vowel / tone mark is cut off); the matches are not counted (a PDF is not read here).
function Get-PdfQuote([string]$before) {
  $t = (Get-OneLine $before).Trim()
  $t = [regex]::Match($t, '^[^\t\u00A0\u000B]*').Value.Trim()
  if ($t.Length -le 40) { return $t }
  $sp = $t.LastIndexOf(' ', 40)
  if ($sp -ge 15) { return $t.Substring(0, $sp).Trim() }
  $cut = 0
  foreach ($bd in [System.Globalization.StringInfo]::ParseCombiningCharacters($t)) { if ($bd -le 40) { $cut = $bd } }
  while ($cut -gt 1 -and [int]$t[$cut - 1] -ge 0x0E40 -and [int]$t[$cut - 1] -le 0x0E44) { $cut-- }
  if ($cut -lt 1) { $cut = 40 }
  while ($cut -lt $t.Length -and (Test-FindCombiningChar $t[$cut])) { $cut++ }
  return $t.Substring(0, $cut).Trim()
}

$catThai = [ordered]@{ must = 'ควรแก้ตามหลักภาษา'; may = 'แก้ได้เพื่ออ่านง่าย'; confirm = 'ต้องยืนยัน' }

# The sentence about the chat's AI service processing the document text (same wording as references/file-input.md section 6).
# Written right after the file-storage note of a brief and of a review/report file, and printed on stdout as its own line, so the
# model repeats it from the script's output instead of remembering to add it. Not used by the rewrite preview.
$AiPrivacySentence = 'สคริปต์ของ skill สร้างไฟล์ไว้ในโฟลเดอร์นี้เท่านั้น ไม่อัปโหลดไฟล์ไปที่ใด แต่ข้อความในเอกสารที่ให้ตรวจถูกประมวลผลโดยบริการ AI ของแชตที่ใช้อยู่ จึงใช้กับเอกสารที่องค์กรอนุญาตให้ใช้กับบริการนั้นเท่านั้น'

# ---------------------------------------------------------------------------------------------
# Create
# ---------------------------------------------------------------------------------------------
if ($Action -eq 'Create') {
  if (-not $Path) { Fail 'ต้องระบุ -Path' }
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { Fail "ไม่พบไฟล์: $Path" }
  $srcFull = (Resolve-Path -LiteralPath $Path).ProviderPath
  $srcName = Split-Path $srcFull -Leaf

  if (-not $Items) { Fail 'ต้องระบุ -Items' }
  try {
    if (-not (Test-Path -LiteralPath $Items -PathType Leaf)) { Fail "ไม่พบไฟล์: $Items" }
    try {
      $jsonText = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Items).ProviderPath, $utf8)
      $parsed = ConvertFrom-Json -InputObject $jsonText
    } catch { Fail 'ไฟล์รายการไม่ใช่ JSON ที่ถูกต้อง' }
    if (-not $parsed) { Fail 'ไฟล์รายการไม่ใช่ JSON ที่ถูกต้อง' }

    if ($Kind -eq 'brief') {
      # ---------------- brief ----------------
      if (-not $parsed.brief) { Fail 'ต้องระบุ brief' }
      $brief = $parsed.brief
      $sections = Get-AsArray $brief.sections
      if (-not $sections.Count) { Fail 'brief.sections ต้องมีอย่างน้อยหนึ่งรายการ' }
      $seenSec = New-Object 'System.Collections.Generic.HashSet[string]'
      $secValidated = New-Object System.Collections.ArrayList
      $i = 0
      foreach ($s in $sections) {
        $i++
        $sid = [string]$s.section
        if (-not $sid.Trim()) { Fail "brief.sections รายการที่ ${i}: ต้องมี section" }
        if ($seenSec.Contains($sid)) { Fail "brief.sections: section $sid ซ้ำ" }
        [void]$seenSec.Add($sid)
        [void]$secValidated.Add(@{ Section = $sid; Heading = [string]$s.heading; Points = (Get-AsArray $s.points) })
      }
      # F5 (docs/v2-plan.md 3.8): "language" and "relations" must both be present keys on brief -
      # checked by NAME (PSObject.Properties), not by truthiness, so an explicit "" / [] still
      # counts as present (and renders "ไม่มี") but an outright MISSING key is refused. This is the
      # code-side guarantee that the reader role cannot silently skip stating the document's
      # language or the cross-section relations it was asked to note.
      $briefPropNames = @($brief.PSObject.Properties.Name)
      if ($briefPropNames -notcontains 'language') { Fail 'brief ต้องมี language' }
      if ($briefPropNames -notcontains 'relations') { Fail 'brief ต้องมี relations' }
      # confirm points: plain texts (no section) and/or {"section","text"} objects (a section of brief.sections)
      $confirmPoints = Get-BriefConfirmItems $brief.confirm $seenSec

      # The source's own structure, for a .docx/.txt (a .pdf has no ¶ numbering to resolve against, so
      # nothing below applies to it): each listed section must exist; a section with no paragraph a rewrite
      # could change is labelled "ข้าม"; the tables in a section are counted (tables are never rewritten);
      # a large section that can be split by its sub-headings is reported on stdout only.
      $briefExt = [System.IO.Path]::GetExtension($srcFull).ToLowerInvariant()
      $tableNsAll = New-Object 'System.Collections.Generic.HashSet[int]'
      $tableNsUnlisted = New-Object 'System.Collections.Generic.HashSet[int]'
      $tableNsListed = New-Object 'System.Collections.Generic.HashSet[int]'
      $skippedSecs = New-Object System.Collections.ArrayList
      $splitNotes = New-Object System.Collections.ArrayList
      $tableStdout = New-Object System.Collections.ArrayList
      $listedStdout = New-Object System.Collections.ArrayList
      # brief.tables (optional): the table cells the author proposes to rewrite. Each entry is checked against the
      # source itself: the ¶ must be a table in that section and every coordinate must be a cell that can be changed.
      # Only the cells listed here (and confirmed by the user with the brief) can be rewritten in the preview step.
      $tablesRaw = Get-AsArray $brief.tables
      $briefTables = New-Object System.Collections.ArrayList
      $listedTableNs = New-Object 'System.Collections.Generic.HashSet[int]'
      if ($tablesRaw.Count -and $briefExt -ne '.docx') { Fail 'brief.tables: ใช้ได้เฉพาะไฟล์ .docx (ไฟล์ .txt และ .pdf ไม่มีตารางให้เขียนใหม่)' }
      if ($briefExt -eq '.docx' -or $briefExt -eq '.txt') {
        if ($briefExt -eq '.docx') { $bItems = Read-Docx $srcFull; $bNs = $script:DocxNs } else { $bItems = Read-Txt $srcFull; $bNs = $null }
        $bHeadCount = @($bItems | Where-Object { $_.Level -gt 0 }).Count
        $ti = 0
        $seenCells = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($t in $tablesRaw) {
          $ti++
          $tsid = [string]$t.section
          if (-not $seenSec.Contains($tsid)) { Fail "brief.tables รายการที่ ${ti}: section `"$tsid`" ไม่อยู่ใน brief.sections" }
          $tn = ConvertTo-ParagraphNumber $t.table
          if (-not $tn) { Fail "brief.tables รายการที่ ${ti}: table ต้องเป็นรูปแบบ ¶7 (เลข ¶ ของตาราง)" }
          $tprof = Get-SectionProfile $bItems $bNs $tsid
          if ($tprof.TableNs -notcontains $tn) {
            $tableListText = if ($tprof.TableNs.Count) { (($tprof.TableNs | ForEach-Object { "¶$_" }) -join ', ') } else { 'ไม่มี' }
            Fail "brief.tables: ¶$tn ไม่ใช่ตารางใน $tsid (ตารางใน $tsid คือ: $tableListText)"
          }
          $tItem = $bItems | Where-Object { $_.N -eq $tn } | Select-Object -First 1
          $tCells = Get-TableCells $tItem.Node
          $coordList = @((Get-AsArray $t.cells) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
          if (-not $coordList.Count) { Fail "brief.tables: ¶$tn ต้องมี cells อย่างน้อยหนึ่งเซลล์" }
          $cellInfos = New-Object System.Collections.ArrayList
          foreach ($coord in $coordList) {
            $res = Resolve-TableCellTarget $tCells $coord
            if ($res.Error) { Fail "brief.tables: ¶$tn ${coord}: $($res.Error)" }
            if (-not $seenCells.Add("$tn|$coord")) { Fail "brief.tables: ¶$tn $coord ซ้ำ" }
            $ctx = Get-TableCellContext $tCells $res.Rec.Row $res.Rec.Cell
            [void]$cellInfos.Add(@{ Cell = $coord; Before = [string]$res.Rec.Text; ColumnHead = $ctx.ColumnHead; RowHead = $ctx.RowHead })
          }
          [void]$listedTableNs.Add($tn)
          [void]$briefTables.Add(@{ Section = $tsid; Table = $tn; Cells = $coordList; Reason = [string]$t.reason; Infos = $cellInfos })
        }
        foreach ($s in $secValidated) {
          $sid = $s.Section
          $sidMatch = [regex]::Match($sid, '^S(\d+)$')
          if (-not $sidMatch.Success) { Fail "brief.sections: section ต้องเป็นรูปแบบ S3 (พบ `"$sid`")" }
          $sidNo = 0
          if (-not [int]::TryParse($sidMatch.Groups[1].Value, [ref]$sidNo)) { $sidNo = 0 }
          if ($sidNo -lt 1 -or $sidNo -gt $bHeadCount) { Fail "brief.sections: ไม่มีหัวข้อ $sid ในไฟล์ต้นฉบับ (ต้นฉบับมี $bHeadCount หัวข้อ)" }
          $prof = Get-SectionProfile $bItems $bNs $sid
          $s.TableCount = $prof.TableNs.Count
          # a table is "listed" when brief.tables names it (by ¶, so a section and its sub-section agree)
          $s.ListedTableNs = @($prof.TableNs | Where-Object { $listedTableNs.Contains([int]$_) })
          $s.UnlistedCount = $s.TableCount - $s.ListedTableNs.Count
          # the paragraphs of the section that a rewrite cannot overwrite (anchors), named for the user by the script itself
          $s.AnchorLine = Get-BriefAnchorLine $bItems $bNs $sid
          # a section with no paragraph to rewrite is skipped - unless a table cell in it is to be rewritten
          $s.Skip = (-not $prof.HasText) -and (-not $s.ListedTableNs.Count)
          if ($s.Skip) { [void]$skippedSecs.Add($sid) }
          if ($s.TableCount) {
            foreach ($tn in $prof.TableNs) {
              [void]$tableNsAll.Add([int]$tn)
              if ($listedTableNs.Contains([int]$tn)) { [void]$tableNsListed.Add([int]$tn) } else { [void]$tableNsUnlisted.Add([int]$tn) }
            }
            if ($s.UnlistedCount) { [void]$tableStdout.Add("${sid}: $($s.UnlistedCount)") }
            if ($s.ListedTableNs.Count) { [void]$listedStdout.Add("${sid}: $($s.ListedTableNs.Count)") }
          }
          if ($briefExt -eq '.docx') {
            $splitIds = @(Get-SplitSubsectionIds $bItems $bNs $sid)
            if ($splitIds.Count) {
              [void]$splitNotes.Add("หมายเหตุ: $sid ใหญ่ ($($prof.Count) ¶ เกิน $($script:LargeSectionParagraphs) ¶) และไม่มีเนื้อความของตัวเองก่อนหัวข้อย่อย ควรแบ่ง preview ตามหัวข้อย่อย: $($splitIds -join ', ') (หัวข้อ $sid คงไว้ตามเดิม)")
            }
          }
        }
      }
      $skipLabel = 'ข้าม: ไม่มีข้อความให้เขียนใหม่ (มีแต่หัวข้อ/ตาราง/รูป) คงไว้ตามเดิม'
      $tableSummaryLine = $null
      if (@($secValidated | Where-Object { $_.TableCount }).Count -ge 2) {
        if (-not $tableNsListed.Count) {
          $tableSummaryLine = "- ตารางในช่วงที่สรุป: รวม $($tableNsAll.Count) ตาราง ไม่ถูกเขียนใหม่ (คงไว้ตามเดิม) ถ้าต้องการปรับข้อความในตาราง ใช้ทาง `"ตรวจและแก้`" กับช่วงนั้น"
        } else {
          $tableSummaryLine = "- ตารางในช่วงที่สรุป: รวม $($tableNsAll.Count) ตาราง เขียนใหม่เฉพาะเซลล์ที่ระบุ $($tableNsListed.Count) ตาราง (ดูหัว ข้อความในตารางที่จะเขียนใหม่) ไม่ถูกเขียนใหม่ $($tableNsUnlisted.Count) ตาราง (คงไว้ตามเดิม) ถ้าต้องการปรับข้อความในตารางอื่น ใช้ทาง `"ตรวจและแก้`" กับช่วงนั้น"
        }
      }

      $now = Get-Date -Format 'yyyy-MM-dd HH:mm'
      $range = Get-OneLine ([string]$parsed.range)
      $mode = Get-OneLine ([string]$parsed.mode)

      $lines = New-Object System.Collections.ArrayList
      [void]$lines.Add("# สรุปความเข้าใจ (brief): $srcName")
      [void]$lines.Add('')
      [void]$lines.Add("- ไฟล์ต้นฉบับ: $srcName (ไม่ได้แก้ไฟล์นี้)")
      [void]$lines.Add("- ช่วงที่สรุป: $range | โหมด: $mode")
      [void]$lines.Add("- สร้างเมื่อ: $now")
      if ($tableSummaryLine) { [void]$lines.Add($tableSummaryLine) }
      [void]$lines.Add('')
      [void]$lines.Add('## วัตถุประสงค์')
      [void]$lines.Add((Format-TextOrNone ([string]$brief.purpose)))
      [void]$lines.Add('')
      [void]$lines.Add('## ผู้อ่านของเอกสาร')
      [void]$lines.Add((Format-TextOrNone ([string]$brief.audience)))
      [void]$lines.Add('')
      [void]$lines.Add('## ภาษาของเอกสาร')
      [void]$lines.Add((Format-TextOrNone ([string]$brief.language)))
      [void]$lines.Add('')
      [void]$lines.Add('## ประเด็นหลักต่อ section')
      foreach ($s in $secValidated) {
        $head = Get-OneLine $s.Heading
        $header = "### $($s.Section)"
        if ($head.Trim()) { $header += " — $head" }
        [void]$lines.Add($header)
        if ($s.Skip) {
          # No paragraph to rewrite: the script says so itself. Points the reader role gave are still
          # shown; a blank list is not padded with "ไม่มี" because the label already states the case.
          [void]$lines.Add($skipLabel)
          $skipPoints = @((Get-AsArray $s.Points) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
          if ($skipPoints.Count) { foreach ($l in (Format-ListOrNone $skipPoints)) { [void]$lines.Add($l) } }
        } else {
          foreach ($l in (Format-ListOrNone $s.Points)) { [void]$lines.Add($l) }
        }
        if ($s.AnchorLine) {
          [void]$lines.Add('')
          [void]$lines.Add($s.AnchorLine)
        }
        if ($s.TableCount) {
          [void]$lines.Add('')
          if ($s.UnlistedCount) {
            [void]$lines.Add("ตารางใน section นี้ $($s.UnlistedCount) ตาราง ไม่ถูกเขียนใหม่ (คงไว้ตามเดิม) ถ้าต้องการปรับข้อความในตาราง ใช้ทาง `"ตรวจและแก้`" กับช่วงนี้")
          }
          if ($s.ListedTableNs.Count) {
            [void]$lines.Add("ตารางใน section นี้ $($s.ListedTableNs.Count) ตาราง ($(($s.ListedTableNs | ForEach-Object { "¶$_" }) -join ', ')) เขียนใหม่เฉพาะเซลล์ที่ระบุ (ดูหัว ข้อความในตารางที่จะเขียนใหม่) เซลล์อื่นและโครงตารางคงเดิม")
          }
        }
        [void]$lines.Add('')
      }
      if ($briefTables.Count) {
        [void]$lines.Add('## ข้อความในตารางที่จะเขียนใหม่')
        [void]$lines.Add('เฉพาะเซลล์ที่ระบุข้างล่างเท่านั้นที่ถูกเขียนใหม่ได้ในขั้นถัดไป เซลล์อื่นและโครงตาราง (แถว เซลล์ การผสาน สไตล์) คงเดิม ข้อความใหม่ต้องยังตอบหัวคอลัมน์/หัวแถวเดิม และสอดคล้องกับเซลล์อื่นในแถวและเนื้อความของ section ถ้าไม่ต้องการเซลล์ใด ให้บอกตอนยืนยัน brief')
        [void]$lines.Add('')
        foreach ($bt in $briefTables) {
          [void]$lines.Add("### ตาราง ¶$($bt.Table) ($($bt.Section))")
          [void]$lines.Add("- เหตุผล: $(Format-TextOrNone $bt.Reason)")
          foreach ($ci in $bt.Infos) {
            $ctxParts = New-Object System.Collections.ArrayList
            if ($ci.ColumnHead) { [void]$ctxParts.Add("หัวคอลัมน์: $($ci.ColumnHead)") }
            if ($ci.RowHead) { [void]$ctxParts.Add("หัวแถว: $($ci.RowHead)") }
            $ctxText = if ($ctxParts.Count) { ' (' + ($ctxParts -join ' | ') + ')' } else { '' }
            [void]$lines.Add("- เซลล์ $($ci.Cell)$ctxText")
            [void]$lines.Add("  - ข้อความเดิม: $(Get-ShortText $ci.Before 120)")
          }
          [void]$lines.Add('')
        }
      }
      [void]$lines.Add('## ความเชื่อมโยงระหว่าง section')
      foreach ($l in (Format-ListOrNone (Get-AsArray $brief.relations))) { [void]$lines.Add($l) }
      [void]$lines.Add('')
      [void]$lines.Add('## รายการที่ต้องคงไว้ (ชื่อ/ตัวย่อ/ตัวเลข/หน่วย/วันที่/คำผูกมัด-ปฏิเสธ/การอ้างอิง/ศัพท์เฉพาะ)')
      foreach ($l in (Format-ListOrNone (Get-AsArray $brief.preserve))) { [void]$lines.Add($l) }
      [void]$lines.Add('')
      [void]$lines.Add('## ปัญหาระดับโครงสร้าง')
      foreach ($l in (Format-ListOrNone (Get-AsArray $brief.structure_issues))) { [void]$lines.Add($l) }
      [void]$lines.Add('')
      # style_notes (optional key): the linguist role's observations on passages that read as
      # AI-like, with paragraph numbers and reasons. Absent or empty renders "ไม่มี" like any other
      # optional field; it is never required and never invented by the script.
      [void]$lines.Add('## ข้อสังเกตด้านสำนวน')
      foreach ($l in (Format-ListOrNone (Get-AsArray $brief.style_notes))) { [void]$lines.Add($l) }
      [void]$lines.Add('')
      [void]$lines.Add('## จุดกำกวมที่ต้องให้ผู้เขียนยืนยัน')
      # a point that names its section is shown as "- S2: <text>", one that does not as "- <text>"
      foreach ($l in (Format-ListOrNone @($confirmPoints | ForEach-Object { Format-ConfirmPoint $_ }))) { [void]$lines.Add($l) }
      [void]$lines.Add('')
      # What the user must be shown in the chat: the brief exactly as the file has it (every heading, in the file's order),
      # from the title to the confirm points, plus the fixed lines the script adds itself (the table-cell trim line when
      # brief.tables lists cells, and the closing question). Printed on stdout after the summary lines below, between an
      # opening and a closing line, so the author copies it as it is instead of re-writing it. It leaves out the data lines
      # (hz-...), the sha256 trailer and the file-storage note: those are about the file, not for the user to read.
      $chatLines = New-Object System.Collections.ArrayList
      foreach ($cl in $lines) { [void]$chatLines.Add($cl) }
      if ($briefTables.Count) {
        [void]$chatLines.Add('ตัดเซลล์ใดออกจากรายการเขียนใหม่ได้ก่อนยืนยัน brief (ตอบ "แก้ brief" แล้วบอกเซลล์ที่ไม่ต้องการ)')
        [void]$chatLines.Add('')
      }
      [void]$chatLines.Add('ถูกต้อง / แก้ brief / ยกเลิก (ไม่ตอบ = ไม่ไปต่อ)')
      [void]$lines.Add('ตอบในแชตว่า "ถูกต้อง" เพื่อเขียนต่อจาก brief นี้, "แก้ brief" เพื่อขอแก้ไข, หรือ "ยกเลิก" (ไม่ตอบ = ไม่ไปต่อ)')
      [void]$lines.Add('')
      if ($OutDir) {
        [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ ถ้าโฟลเดอร์ที่เก็บไฟล์นี้แชร์กับผู้อื่นหรือถูกส่งขึ้น repo ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
      } else {
        [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ เก็บอยู่ในโฟลเดอร์ HumanizerAgent\output\review ใน Documents ของคุณ ถ้า Documents ซิงก์กับ OneDrive ไฟล์นี้จะถูกซิงก์ขึ้นไปด้วย ถ้าโฟลเดอร์นี้แชร์กับผู้อื่น ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
      }
      [void]$lines.Add($AiPrivacySentence)

      # Machine-readable lines inside the part the trailer's sha256 covers, read back by the preview step (-Kind rewrite)
      # through Get-LatestBriefInfo:
      #   hz-tables (only when brief.tables lists cells): the cells the preview may rewrite;
      #   hz-brief (v2): the section ids this brief covers (to know whether it spans the whole file) and its confirm points,
      #   each with the section it names ("" = none) - the preview must carry each one in its own "confirm" until the user
      #   has answered it in chat. Version 1 (written by earlier versions) held the points as plain texts.
      [void]$lines.Add('')
      if ($briefTables.Count) {
        $tablesJson = ConvertTo-Json -InputObject ([ordered]@{ tables = @($briefTables | ForEach-Object { [ordered]@{ section = [string]$_.Section; table = [int]$_.Table; cells = @($_.Cells | ForEach-Object { [string]$_ }) } }) }) -Depth 6 -Compress
        [void]$lines.Add("<!-- hz-tables v1 data=`"$([Convert]::ToBase64String($utf8.GetBytes($tablesJson)))`" -->")
      }
      $briefDataJson = ConvertTo-Json -InputObject ([ordered]@{
          sections = @($secValidated | ForEach-Object { [string]$_.Section })
          confirm  = @($confirmPoints | ForEach-Object { [ordered]@{ section = [string]$_.Section; text = [string]$_.Text } })
        }) -Depth 4 -Compress
      [void]$lines.Add("<!-- hz-brief v2 data=`"$([Convert]::ToBase64String($utf8.GetBytes($briefDataJson)))`" -->")

      $body = ($lines -join "`n").TrimEnd("`n")
      $hash = Get-BodyHash $body
      $trailer = "<!-- hz-review v1 kind=`"brief`" items=`"`" sha256=`"$hash`" -->"
      $fullContent = (($body + "`n`n" + $trailer + "`n") -replace "`n", "`r`n")

      $dest = Get-NewPath $srcFull 'brief' '.md' $null $OutDir 'review'
      $opened = New-FileExclusive $dest
      $dest = $opened.Path
      try {
        $bom = New-Object System.Text.UTF8Encoding($true)
        $sw = New-Object System.IO.StreamWriter($opened.Stream, $bom)
        $sw.Write($fullContent)
        $sw.Dispose()
      } finally { $opened.Stream.Dispose() }

      Write-Output "สร้างไฟล์สรุปความเข้าใจ: $dest"
      Write-Output ("section: " + (($secValidated | ForEach-Object { $_.Section }) -join ', '))
      Write-Output "รหัสไฟล์: $($hash.Substring(0, 8))"
      if ($skippedSecs.Count) { Write-Output ("ข้าม (ไม่มีข้อความให้เขียนใหม่): " + ($skippedSecs -join ', ')) }
      if ($tableStdout.Count) { Write-Output ("ตารางที่ไม่ถูกเขียนใหม่: รวม {0} ตาราง ({1})" -f $tableNsUnlisted.Count, ($tableStdout -join ', ')) }
      if ($listedStdout.Count) {
        Write-Output ("ตารางที่เขียนใหม่เฉพาะเซลล์ที่ระบุ: รวม {0} ตาราง {1} เซลล์ ({2})" -f $tableNsListed.Count, (@($briefTables | ForEach-Object { $_.Cells.Count } | Measure-Object -Sum).Sum), ($listedStdout -join ', '))
      }
      foreach ($n in $splitNotes) { Write-Output $n }
      Write-Output '----- ข้อความ brief สำหรับแสดงในแชต (แสดงตามนี้ครบทุกบรรทัด ห้ามย่อหรือสรุปใหม่) -----'
      foreach ($cl in $chatLines) { Write-Output $cl }
      Write-Output '----- จบข้อความ brief -----'
      Write-Output $AiPrivacySentence
      exit 0
    }

    if ($Kind -eq 'rewrite') {
      # ---------------- rewrite ----------------
      $sections = Get-AsArray $parsed.sections
      if (-not $sections.Count) { Fail 'sections ต้องมีอย่างน้อยหนึ่งรายการ' }
      # Ids first, for the whole file: a bad or repeated section id is named before any section is checked.
      $seenSec = New-Object 'System.Collections.Generic.HashSet[string]'
      $si = 0
      foreach ($s in $sections) {
        $si++
        $sid = [string]$s.section
        if ($sid -cnotmatch '^S\d+$') { Fail "sections รายการที่ ${si}: section ต้องเป็นรูปแบบ S3" }
        if ($seenSec.Contains($sid)) { Fail "sections: section $sid ซ้ำ" }
        [void]$seenSec.Add($sid)
      }
      $script:SrcExt = [System.IO.Path]::GetExtension($srcFull).ToLowerInvariant()
      if ($script:SrcExt -ne '.docx' -and $script:SrcExt -ne '.txt') { Fail "ตัวอย่างก่อนสร้างไฟล์ (rewrite) ทำได้เฉพาะไฟล์ .docx และ .txt เท่านั้น (ไฟล์นี้เป็น $($script:SrcExt))" }
      $mapIgnored = $false
      # Does any "after" entry rewrite table cells ({"table":"¶7","cells":[...]})? Then a snapshot of the tables is
      # kept for the structure self-check, and the cells must be authorized by the latest brief (checked below).
      $script:WantTableSnap = $false
      foreach ($s in $sections) {
        foreach ($a in (Get-AsArray $s.after)) {
          if ($null -ne $a -and $a -isnot [string] -and (@($a.PSObject.Properties.Name) -contains 'table') -and @(Get-AsArray $a.cells).Count) { $script:WantTableSnap = $true }
        }
      }
      $state = New-RewriteSourceState
      # The units of a whole-file rewrite and the ¶ range of every heading, taken from the untouched source (the dry-run
      # builds below change the in-memory copy). Used by the whole-file check once the brief has been found.
      $wholeUnits = @(Get-WholeFileUnits $state.Items $state.Ns)
      $headRanges = @{}
      $headItems = @($state.Items | Where-Object { $_.Level -gt 0 })
      for ($hx = 0; $hx -lt $headItems.Count; $hx++) { $headRanges["S$($hx + 1)"] = @{ Start = $headItems[$hx].N; End = (Get-SectionEnd $state.Items $headItems[$hx]) } }

      $secValidated = New-Object System.Collections.ArrayList
      foreach ($s in $sections) {
        $sid = [string]$s.section
        if ($s.PSObject.Properties.Name -contains 'map') { $mapIgnored = $true }

        $beforeList = Get-AsArray $s.before
        if (-not $beforeList.Count) { Fail "${sid}: ต้องมี before อย่างน้อยหนึ่งย่อหน้า" }

        # The section's ¶ range in the source. A section that overlaps one already dry-run gets a
        # fresh in-memory copy of the document (the earlier build changed the previous copy).
        $range = Resolve-Section $state.Items $sid
        $rangeStart = $range.Head.N; $rangeEnd = $range.End
        if ($script:SrcExt -eq '.docx') {
          $overlap = [bool](@($state.Ranges | Where-Object { $rangeStart -le $_.End -and $rangeEnd -ge $_.Start }).Count)
          if ($overlap) { $state = New-RewriteSourceState; $range = Resolve-Section $state.Items $sid; $rangeStart = $range.Head.N; $rangeEnd = $range.End }
        }

        # ---- before: one entry per ¶ of the section. "keep":true = stays exactly as it is (anchor,
        # unchanged heading, or an ordinary paragraph left alone). The text is always the source's. ----
        $beforeParts = New-Object System.Collections.ArrayList
        $beforeIdSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($b in $beforeList) {
          $bid = [string]$b.id
          if (-not $bid.Trim()) { Fail "${sid}: before แต่ละรายการต้องมี id" }
          if ($bid -notmatch '^¶(\d+)$') { Fail "${sid}: before ${bid}: id ต้องเป็นรูปแบบ ¶ตัวเลข ตรงกับ ¶ ในต้นฉบับ" }
          $pnum = [int]$Matches[1]
          $bkeep = [bool]$b.keep
          $info = Get-SourceItemInfo $state $sid $bid $pnum
          if (-not $beforeIdSet.Add($bid)) { Fail "${sid}: before $bid ซ้ำ" }
          [void]$beforeParts.Add(@{ Id = $bid; N = $pnum; Text = $info.Text; Keep = $bkeep; Anchor = $info.IsAnchor; Marked = ($info.IsAnchor -or $info.IsHeading) })
        }

        # ---- after: every entry is an object - {"text","from"} (a new paragraph), {"heading":level,
        # "text","from"} (a heading whose text is rewritten) or {"keep":"¶id"}. A plain string is the old
        # format and is refused. Blank text entries are dropped (they leave their ¶ uncovered, which the
        # build check below names). ----
        $afterRaw = Get-AsArray $s.after
        $blocks = New-Object System.Collections.ArrayList
        $ai0 = 0
        $formatMsg = '{"text":"...","from":["¶9"]} (ย่อหน้าใหม่), {"heading":2,"text":"...","from":["¶6"]} (หัวข้อที่แก้ข้อความ) หรือ {"keep":"¶7"} (คงไว้) ข้อความล้วนแบบเดิมใช้ไม่ได้แล้ว'
        foreach ($a in $afterRaw) {
          $ai0++
          if ($null -eq $a -or $a -is [string]) { Fail "${sid}: after รายการที่ ${ai0}: ต้องเป็น $formatMsg" }
          $names = @($a.PSObject.Properties.Name)
          if ($names -contains 'keep') {
            $kid = [string]$a.keep
            if (-not $kid.Trim()) { Fail "${sid}: after รายการที่ ${ai0}: keep ต้องระบุ id" }
            $kn = ConvertTo-ParagraphNumber $kid
            if (-not $kn) { Fail "${sid}: after รายการที่ ${ai0}: keep ${kid}: id ต้องเป็นรูปแบบ ¶ตัวเลข" }
            [void]$blocks.Add(@{ Type = 'keep'; Paragraph = $kn })
          } elseif ($names -contains 'table') {
            # {"table":"¶7","cells":[{"cell":"r4c2","text":"..."}]}: the table stays at its place (it counts as the
            # keep of that anchor) and only the listed cells get new text. No cells = the table is just kept.
            $tid = [string]$a.table
            $tn = ConvertTo-ParagraphNumber $tid
            if (-not $tn) { Fail "${sid}: after รายการที่ ${ai0}: table ${tid}: id ต้องเป็นรูปแบบ ¶ตัวเลข" }
            $cellsArr = Get-AsArray $a.cells
            if (-not $cellsArr.Count) { [void]$blocks.Add(@{ Type = 'keep'; Paragraph = $tn }) }
            else {
              if ($script:SrcExt -ne '.docx') { Fail "${sid}: after รายการที่ ${ai0}: ไฟล์ .txt ไม่มีตาราง เขียนข้อความในเซลล์ตารางไม่ได้" }
              [void]$blocks.Add((New-TableBlock $tn $cellsArr "${sid}: after รายการที่ ${ai0} (table $tid)"))
            }
          } elseif ($names -contains 'heading' -or $names -contains 'text') {
            $isHeading = ($names -contains 'heading')
            $level = 0
            if ($isHeading) {
              $level = $a.heading -as [int]
              if (-not $level -or $level -lt 1) { Fail "${sid}: after รายการที่ ${ai0}: heading ต้องเป็นเลขระดับหัวข้อตั้งแต่ 1" }
            }
            $fromRaw = Get-AsArray $a.from
            $from = New-Object System.Collections.ArrayList
            foreach ($f in $fromRaw) {
              $fn = ConvertTo-ParagraphNumber $f
              if (-not $fn) { Fail "${sid}: after รายการที่ ${ai0}: from ต้องเป็น ¶ ตัวเลข เช่น [""¶9""]" }
              [void]$from.Add($fn)
            }
            if (-not $from.Count) { Fail "${sid}: after รายการที่ ${ai0}: ต้องมี from (¶ ที่มาอย่างน้อยหนึ่งตัว) เช่น [""¶9""]" }
            $atext = [string]$a.text
            if (-not $atext.Trim()) { continue }
            if ($isHeading) { [void]$blocks.Add(@{ Type = 'heading'; Level = $level; Text = $atext; From = @($from) }) }
            else { [void]$blocks.Add(@{ Type = 'paragraph'; Text = $atext; From = @($from) }) }
          } else {
            Fail "${sid}: after รายการที่ ${ai0}: ต้องเป็น $formatMsg"
          }
        }

        # F2 keep contract: every before item marked keep must appear in "after" as {"keep"} EXACTLY once,
        # in the same relative order; an after keep whose id is not one of before's own ids is refused.
        # Checked before the "after has at least one entry" rule so a keep-contract violation is named
        # clearly rather than masked by the more generic message.
        $expectedKeepIds = @($beforeParts | Where-Object { $_.Keep } | ForEach-Object { $_.Id })
        $actualKeepIds = @($blocks | Where-Object { $_.Type -eq 'keep' -or $_.Type -eq 'table' } | ForEach-Object { "¶$($_.Paragraph)" })
        $seenAfterKeep = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($kid in $actualKeepIds) {
          if (-not $beforeIdSet.Contains($kid)) { Fail "${sid}: after keep $kid ไม่มีใน before ของ section นี้" }
          if (-not $seenAfterKeep.Add($kid)) { Fail "${sid}: after keep $kid ซ้ำ" }
        }
        if ($actualKeepIds.Count -ne $expectedKeepIds.Count) {
          $missing = @($expectedKeepIds | Where-Object { $actualKeepIds -notcontains $_ })
          if ($missing.Count) { Fail ("${sid}: after ต้องมี keep ครบทุกรายการที่ before ทำเครื่องหมาย keep ไว้ ขาด " + ($missing -join ', ')) }
          $extra = @($actualKeepIds | Where-Object { $expectedKeepIds -notcontains $_ })
          Fail ("${sid}: after มี keep ที่ before ไม่ได้ทำเครื่องหมายไว้ (" + ($extra -join ', ') + ")")
        }
        for ($ki = 0; $ki -lt $expectedKeepIds.Count; $ki++) {
          if ($actualKeepIds[$ki] -ne $expectedKeepIds[$ki]) {
            Fail ("${sid}: ลำดับ keep ใน after ไม่ตรงกับ before (ต้องเรียง " + ($expectedKeepIds -join ', ') + ")")
          }
        }
        if (-not $blocks.Count) { Fail "${sid}: after ต้องมีข้อความอย่างน้อยหนึ่งย่อหน้า" }

        # A section with nothing a rewrite could change (only headings / tables / pictures) has no preview:
        # the user would only be asked to confirm text that cannot differ. The file build keeps it as it is.
        # The exception is a plan that really rewords a heading of such a section (a "heading" entry in
        # "after"): that is a change the user can review, so it is not refused.
        $rewriteProfile = Get-SectionProfile $state.Items $state.Ns $sid
        $changesNothing = -not (@($blocks | Where-Object { $_.Type -ne 'keep' }).Count)
        if ((-not $rewriteProfile.HasText) -and $changesNothing) { Fail "${sid}: section นี้ไม่มีข้อความให้เขียนใหม่ (มีแต่หัวข้อ/ตาราง/รูป) ข้ามได้ ไม่ต้องสร้าง preview ของ section นี้ คงไว้ตามเดิม" }

        # before must list every ¶ of the section, once, in the source's order.
        $expectedNs = @($state.Items | Where-Object { $_.N -ge $rangeStart -and $_.N -le $rangeEnd } | ForEach-Object { $_.N })
        $rangeLabel = "¶$rangeStart-¶$rangeEnd"
        foreach ($bp in $beforeParts) {
          if ($expectedNs -notcontains $bp.N) { Fail "${sid}: before $($bp.Id) อยู่นอก section นี้ (section นี้คือ $rangeLabel)" }
        }
        $missingNs = @($expectedNs | Where-Object { @($beforeParts | ForEach-Object { $_.N }) -notcontains $_ })
        if ($missingNs.Count) { Fail ("${sid}: before ต้องมีทุก ¶ ของ section นี้ ($rangeLabel) ขาด ¶" + ($missingNs -join ', ¶')) }
        for ($bi = 1; $bi -lt $beforeParts.Count; $bi++) {
          if ($beforeParts[$bi].N -lt $beforeParts[$bi - 1].N) { Fail "${sid}: before ต้องเรียงตามลำดับ ¶ ในต้นฉบับ" }
        }
        # A ¶ that stays as it is cannot also be a source of a new paragraph.
        $keptNs = @($beforeParts | Where-Object { $_.Keep } | ForEach-Object { $_.N })
        foreach ($b in $blocks) {
          if ($b.Type -eq 'keep' -or $b.Type -eq 'table') { continue }
          foreach ($fn in $b.From) {
            if ($keptNs -contains $fn) { Fail "${sid}: ¶$fn ถูกทำเครื่องหมายคงไว้ (keep) แต่ยังอยู่ใน from ของย่อหน้าใหม่ ให้เลือกอย่างใดอย่างหนึ่ง" }
          }
        }

        # Dry run of the file build (the same code write-rewrite.ps1 uses): first block = the section's
        # heading, every anchor kept, every "from" in the right place, no ¶ dropped, nothing moved across
        # an anchor or a picture. A failure here is the message the user would otherwise meet only at the end.
        # Table cells to rewrite: each must be a cell that can be changed (the file build checks the same), and what the
        # user sees next to it - the text now and where the cell sits - is read from the source document here. The
        # preservation check of a cell compares that cell alone (numbers, Latin words, obligation/negation words, "/",
        # "( )" pairs and the Thai inter-word spacing), never the cell against the rest of the section.
        $tableInfos = @{}
        $cellResults = New-Object System.Collections.ArrayList
        foreach ($tb in @($blocks | Where-Object { $_.Type -eq 'table' })) {
          $tItem = $state.Items | Where-Object { $_.N -eq $tb.Paragraph } | Select-Object -First 1
          if (-not $tItem -or $tItem.Kind -ne 'table') { Fail "${sid}: ¶$($tb.Paragraph) ไม่ใช่ตาราง เขียนข้อความในเซลล์ไม่ได้ (ใช้ keep ตามปกติ)" }
          $tCells = Get-TableCells $tItem.Node
          $infos = New-Object System.Collections.ArrayList
          foreach ($tc in $tb.Cells) {
            $res = Resolve-TableCellTarget $tCells ([string]$tc.Cell)
            if ($res.Error) { Fail "${sid}: ตาราง ¶$($tb.Paragraph) $($tc.Cell): $($res.Error)" }
            $cellBefore = [string]$res.Rec.Text
            $cellAfter = ([string]$tc.Text).Trim()
            if ([string]::Equals($cellAfter, $cellBefore, [StringComparison]::Ordinal)) { Fail "${sid}: ตาราง ¶$($tb.Paragraph) $($tc.Cell): ข้อความใหม่เหมือนข้อความเดิม ให้เอาเซลล์นี้ออกจาก cells" }
            $ctx = Get-TableCellContext $tCells $res.Rec.Row $res.Rec.Cell
            [void]$infos.Add(@{ Cell = [string]$tc.Cell; Before = $cellBefore; ColumnHead = $ctx.ColumnHead; RowHead = $ctx.RowHead })
            $gapB = Get-ThaiGapCount $cellBefore; $gapA = Get-ThaiGapCount $cellAfter
            $cellNotes = @()
            if ($gapB -ne $gapA) { $cellNotes = @("ช่องว่างระหว่างคำไทยเปลี่ยน: เซลล์ $($tc.Cell): $gapB -> $gapA (บอกเหตุผลว่าอ่านดีขึ้นอย่างไร ถ้าไม่มีเหตุผลให้คงช่องว่างเดิม)") }
            [void]$cellResults.Add(@{ Table = [int]$tb.Paragraph; Cell = [string]$tc.Cell; Result = (Get-PreservationResult $cellBefore $cellAfter -CheckPunctuation -InfoNotes $cellNotes) })
          }
          $tableInfos[[int]$tb.Paragraph] = @($infos)
        }

        $sectionPlan = @{ Section = $sid; Blocks = $blocks; Head = $range.Head; End = $rangeEnd }
        if ($script:SrcExt -eq '.docx') {
          $built = Build-DocxSection $state.Items $state.Ns $state.Gap $sectionPlan
          Test-DocxRewriteIntact $state.Counts @($built.Verify)
          foreach ($tk in $built.TableKeys) { [void]$state.TableKeys.Add($tk) }
          if ($state.TableSnap) { Test-DocxTablesIntact $state.TableSnap $state.TableKeys }
          [void]$state.Ranges.Add(@{ Start = $rangeStart; End = $rangeEnd })
        } else {
          [void](Build-TxtSection $state.Items $script:TxtText $state.Sep $sectionPlan)
        }

        # The anchor-limitation line (F6) is about real anchors only; a kept heading (H3) is not one.
        $hasKeep = [bool](@($beforeParts | Where-Object { $_.Keep -and $_.Anchor }).Count)

        $modelCheckRaw = Get-AsArray $s.model_check
        $modelCheck = @($modelCheckRaw | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
        $confirmRaw = Get-AsArray $s.confirm
        $confirm = @($confirmRaw | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })
        $fromChatRaw = Get-AsArray $s.from_chat
        $fromChat = @($fromChatRaw | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() })

        # Code-computed preservation result for this section - never taken from the JSON, even if
        # the JSON happens to include its own "preservation"/"code_check" field (that field, if
        # present, is simply never read below). F2: kept items are excluded on BOTH sides - they are
        # unchanged by definition, so comparing them would only add noise.
        $newBlocks = @($blocks | Where-Object { $_.Type -ne 'keep' -and $_.Type -ne 'table' })
        $beforeAll = (($beforeParts | Where-Object { -not $_.Keep } | ForEach-Object { $_.Text }) -join "`n")
        $afterAll = (($newBlocks | ForEach-Object { $_.Text }) -join "`n")
        # Thai inter-word spacing is compared per NEW paragraph (not per section) from the source paragraph(s)
        # it was written from; a difference is a note line in the same result (not a difference that needs a person:
        # it never makes the section "ต้องให้คนตรวจ" or blocks the pre-tick), never a refusal.
        $sourceByN = @{}
        foreach ($bp in $beforeParts) { if (-not $bp.Keep) { $sourceByN[[int]$bp.N] = [string]$bp.Text } }
        $gapWarnings = @(Get-ThaiGapWarnings $sourceByN $newBlocks)
        $result = Get-PreservationResult $beforeAll $afterAll -CheckPunctuation -InfoNotes $gapWarnings

        # The map "¶ เดิม → ย่อหน้าใหม่ N" is built here from the blocks' own "from" - N counts the
        # non-keep entries exactly as the "ย่อหน้าใหม่ N:" headings below do. Same ¶ set -> one line.
        $mapLines = New-Object System.Collections.ArrayList
        $mapIndex = @{}
        $newNo = 0
        foreach ($b in $blocks) {
          if ($b.Type -eq 'keep' -or $b.Type -eq 'table') { continue }
          $newNo++
          $key = ($b.From -join ',')
          if (-not $mapIndex.ContainsKey($key)) {
            $mapIndex[$key] = @{ From = (($b.From | ForEach-Object { "¶$_" }) -join ', '); To = (New-Object System.Collections.ArrayList) }
            [void]$mapLines.Add($mapIndex[$key])
          }
          [void]$mapIndex[$key].To.Add("ย่อหน้าใหม่ $newNo")
        }

        [void]$secValidated.Add(@{
            Section = $sid; Heading = [string]$s.heading; Before = $beforeParts; Blocks = $blocks
            HasKeep = $hasKeep; MapLines = $mapLines; ModelCheck = $modelCheck; Confirm = $confirm; Result = $result
            FromChat = $fromChat; TableInfos = $tableInfos; CellResults = $cellResults
          })
      }

      $dest = Get-NewPath $srcFull 'rewrite' '.md' $null $OutDir 'review'
      $priorOpen = Get-PriorOpenConfirmSections (Split-Path $dest -Parent) $srcName
      $notPreTicked = New-Object System.Collections.ArrayList

      # The latest brief of this file in the same folder. The brief is a mandatory step of a rewrite: without one the
      # run is refused and nothing is created (the user may say to skip WAITING for the confirmation of the brief, never
      # the making of it). It decides:
      #  - whether the preview may hold several sections: a brief that covers the WHOLE file means one preview per section
      #    (references/full-rewrite.md "ทั้งไฟล์"), refused here otherwise - nothing is created;
      #  - whether the brief's confirm points were carried (below);
      #  - which table cells are authorized (next block).
      $briefInfo = Get-LatestBriefInfo (Split-Path $dest -Parent) $srcName
      if (-not $briefInfo) {
        Fail 'ไม่พบ brief (สรุปภาพรวม) ของไฟล์ต้นฉบับนี้ในโฟลเดอร์เดียวกับ preview ต้องสร้าง brief ก่อนด้วย -Kind brief แล้วจึงสร้าง preview (brief เป็นขั้นบังคับ ผู้ใช้สั่งข้ามได้เฉพาะการรอยืนยัน brief ไม่ใช่การสร้าง brief) ถ้าใช้ -OutDir ให้ใช้โฟลเดอร์เดียวกับตอนสร้าง brief (ไม่สร้างไฟล์)'
      }
      if ($briefInfo -and $briefInfo.Info -and $secValidated.Count -gt 1) {
        $briefRanges = @()
        foreach ($bsid in @($briefInfo.Info.Sections)) { if ($headRanges.ContainsKey($bsid)) { $briefRanges += , $headRanges[$bsid] } }
        # whole file = every unit that has text to rewrite lies inside a section of the brief (units with no text are skipped
        # by a whole-file rewrite and need not be listed)
        $needUnits = @($wholeUnits | Where-Object { $_.HasText })
        $uncovered = @($needUnits | Where-Object { $u = $_; -not (@($briefRanges | Where-Object { $_.Start -le $u.Start -and $_.End -ge $u.End }).Count) })
        if ($needUnits.Count -and -not $uncovered.Count) {
          $sentIds = @($secValidated | ForEach-Object { [string]$_.Section } | Sort-Object { [int]($_ -replace '^S', '') })
          Fail ("ขอบเขตทั้งไฟล์ต้องสร้าง preview ทีละ section (ส่งมา $($sentIds.Count) section: $($sentIds -join ', ')) ให้สร้าง preview ของ $($sentIds[0]) ก่อน แล้วทำ section ถัดไปหลังผู้ใช้ยืนยัน")
        }
      }
      # The brief's confirm points: the preview has to carry each one in "confirm" until the user has answered its content in
      # chat. Keeping the original wording of the source does NOT settle a point, and neither does a "from_chat" entry: the
      # script cannot tell whether an entry answers the point (an answer such as "correct" to the brief itself is not one), so
      # the gap is judged on "confirm" alone and a section that has "from_chat" is still listed (and shown the brief's points)
      # so that the user sees which points are still open. The script cannot tell which section a point belongs to, so it
      # only notices the plain gap: the brief has points, and no section of this preview has a "confirm" entry. Then no
      # section is pre-ticked.
      # A point that names its section (brief.confirm {"section":"S2","text":"..."}) is judged per section instead: a
      # section of the preview that the point concerns is not pre-ticked while its own "confirm" is empty (whether or not it
      # has "from_chat"); the other sections of the preview are not affected by it. A point concerns a preview section when the ¶ ranges
      # of the two sections overlap (the same section, or a large section and a sub-heading of it), so a point named on a
      # large section is not lost when its sub-headings are previewed one at a time. The points listed in the preview and on
      # stdout are the ones that concern a section of this preview (points of the sections in it + points without a section).
      $briefConfirmAll = @(if ($briefInfo -and $briefInfo.Info) { $briefInfo.Info.Confirm })
      $untaggedPoints = @($briefConfirmAll | Where-Object { -not $_.Section })
      $secHasTagged = @{}
      $relevantPoints = New-Object System.Collections.ArrayList
      foreach ($bp in $briefConfirmAll) {
        $rel = -not $bp.Section
        if (-not $rel) {
          foreach ($sv in $secValidated) {
            if (Test-SectionRangesTouch $headRanges ([string]$bp.Section) ([string]$sv.Section)) { $secHasTagged[[string]$sv.Section] = $true; $rel = $true }
          }
        }
        if ($rel) { [void]$relevantPoints.Add($bp) }
      }
      $previewCarriesConfirm = [bool](@($secValidated | Where-Object { $_.Confirm.Count }).Count)
      $untaggedGap = ($untaggedPoints.Count -gt 0) -and (-not $previewCarriesConfirm)
      $gapSections = New-Object System.Collections.ArrayList
      $gapChatSections = New-Object System.Collections.ArrayList
      foreach ($sv in $secValidated) {
        $sv.BriefGap = $untaggedGap -or ([bool]$secHasTagged[[string]$sv.Section] -and -not $sv.Confirm.Count)
        if ($sv.BriefGap) {
          [void]$gapSections.Add([string]$sv.Section)
          if ($sv.FromChat.Count) { [void]$gapChatSections.Add([string]$sv.Section) }
        }
      }
      $briefGap = $gapSections.Count -gt 0
      $gapHasChat = $gapChatSections.Count -gt 0

      # Table cells: the user confirmed WHICH cells may be rewritten when they confirmed the brief, so every cell rewritten
      # here must be in the cell list of the latest brief of this file (brief.tables, kept in that brief's hz-tables line)
      # in the same folder. Nothing is created otherwise.
      $needAuth = @($secValidated | Where-Object { $_.CellResults.Count })
      if ($needAuth.Count) {
        $auth = $briefInfo
        foreach ($nsec in $needAuth) {
          foreach ($cr in $nsec.CellResults) {
            $cellName = "ตาราง ¶$($cr.Table) เซลล์ $($cr.Cell)"
            if ($auth.Edited) { Fail "$($nsec.Section): $cellName จะเขียนใหม่ แต่ brief ล่าสุด ($($auth.Name)) ถูกแก้ในไฟล์หรืออ่านรายการเซลล์ไม่ได้ จึงใช้ไม่ได้ ให้สร้าง brief ใหม่ที่ใส่เซลล์นี้ใน brief.tables และให้ผู้ใช้ยืนยันก่อน (ไม่สร้างไฟล์)" }
            if (-not $auth.Allowed.Contains("$($cr.Table)|$($cr.Cell)")) { Fail "$($nsec.Section): $cellName ไม่อยู่ใน brief ล่าสุด ($($auth.Name)) ที่ผู้ใช้ยืนยัน ต้องใส่เซลล์นี้ใน brief (brief.tables) และให้ผู้ใช้ยืนยัน brief ใหม่ก่อน แล้วจึงสร้าง preview (ไม่สร้างไฟล์)" }
          }
        }
      }

      $now = Get-Date -Format 'yyyy-MM-dd HH:mm'
      $range = Get-OneLine ([string]$parsed.range)
      $mode = Get-OneLine ([string]$parsed.mode)

      $lines = New-Object System.Collections.ArrayList
      [void]$lines.Add("# ตัวอย่างก่อนสร้างไฟล์ (rewrite): $srcName")
      [void]$lines.Add('')
      [void]$lines.Add("- ไฟล์ต้นฉบับ: $srcName (ไม่ได้แก้ไฟล์นี้)")
      [void]$lines.Add("- ช่วงที่ rewrite: $range | โหมด: $mode")
      [void]$lines.Add("- สร้างเมื่อ: $now")
      [void]$lines.Add('')
      if ($briefGap) {
        [void]$lines.Add('## จุดต้องยืนยันจาก brief (preview นี้ไม่ได้ยกมา)')
        if ($untaggedGap -and $gapHasChat) {
          [void]$lines.Add("brief ของไฟล์นี้ ($($briefInfo.Name)) มีจุดต้องยืนยันต่อไปนี้ แต่ preview นี้ไม่มีจุดต้องยืนยัน (confirm) ใน section ใดเลย (section $($gapChatSections -join ', ') มีข้อมูลจากคำตอบในแชต (from_chat) แต่ข้อมูลนั้นอาจไม่ใช่คำตอบของจุดเหล่านี้) ทุก section จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจว่าแต่ละข้อได้คำตอบจริงในแชตแล้วหรือยังค้าง (การตอบ `"ถูกต้อง`" ต่อ brief เป็นการยืนยัน brief ไม่ใช่คำตอบของจุดต้องยืนยัน และการคงถ้อยคำเดิมของต้นฉบับก็ไม่นับเป็นการคลายจุดต้องยืนยัน)")
        } elseif ($untaggedGap) {
          [void]$lines.Add("brief ของไฟล์นี้ ($($briefInfo.Name)) มีจุดต้องยืนยันต่อไปนี้ แต่ preview นี้ไม่มีจุดต้องยืนยัน (confirm) และไม่มีข้อมูลจากคำตอบในแชต (from_chat) เลย ทุก section จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจว่าแต่ละข้อมีคำตอบแล้วหรือยัง (การคงถ้อยคำเดิมของต้นฉบับไม่นับเป็นการคลายจุดต้องยืนยัน)")
        } elseif ($gapHasChat) {
          [void]$lines.Add("brief ของไฟล์นี้ ($($briefInfo.Name)) มีจุดต้องยืนยันต่อไปนี้ที่เกี่ยวกับ section ใน preview นี้ (ข้อที่ขึ้นต้นด้วยรหัส section เป็นของ section นั้น) แต่ section ที่เกี่ยวข้อง ($($gapSections -join ', ')) ไม่มีจุดต้องยืนยัน (confirm) (section $($gapChatSections -join ', ') มีข้อมูลจากคำตอบในแชต (from_chat) แต่ข้อมูลนั้นอาจไม่ใช่คำตอบของจุดเหล่านี้) section เหล่านี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจว่าแต่ละข้อได้คำตอบจริงในแชตแล้วหรือยังค้าง (การตอบ `"ถูกต้อง`" ต่อ brief เป็นการยืนยัน brief ไม่ใช่คำตอบของจุดต้องยืนยัน และการคงถ้อยคำเดิมของต้นฉบับก็ไม่นับเป็นการคลายจุดต้องยืนยัน)")
        } else {
          [void]$lines.Add("brief ของไฟล์นี้ ($($briefInfo.Name)) มีจุดต้องยืนยันต่อไปนี้ที่เกี่ยวกับ section ใน preview นี้ (ข้อที่ขึ้นต้นด้วยรหัส section เป็นของ section นั้น) แต่ section ที่เกี่ยวข้อง ($($gapSections -join ', ')) ไม่มีจุดต้องยืนยัน (confirm) และไม่มีข้อมูลจากคำตอบในแชต (from_chat) เลย section เหล่านี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจว่าแต่ละข้อมีคำตอบแล้วหรือยัง (การคงถ้อยคำเดิมของต้นฉบับไม่นับเป็นการคลายจุดต้องยืนยัน)")
        }
        foreach ($bc in $relevantPoints) { [void]$lines.Add("- $(Format-ConfirmPoint $bc)") }
        [void]$lines.Add('')
      }
      [void]$lines.Add('## วิธีใช้')
      [void]$lines.Add('1. เปลี่ยน `[ ]` เป็น `[x]` หน้า section ที่ตรวจแล้วว่าโอเค แล้วบันทึกไฟล์')
      [void]$lines.Add('2. กลับไปที่แชตแล้วพิมพ์ "ยืนยัน" (ทุก section ที่ติ๊ก) หรือ "แก้" หรือ "ไม่สร้าง" (ไม่ตอบ = ไม่สร้าง)')
      [void]$lines.Add('3. ไฟล์ที่สร้างจะใช้ข้อความในช่อง "หลัง" ของไฟล์นี้เท่านั้น ห้ามแก้ข้อความในไฟล์ (ถ้าแก้ ระบบจะไม่สร้างไฟล์จาก section นั้น) ถ้าต้องการเปลี่ยนข้อความให้บอกในแชตเพื่อสร้าง preview ใหม่')
      [void]$lines.Add('')

      foreach ($s in $secValidated) {
        $head = Get-OneLine $s.Heading
        $header = "### $($s.Section)"
        if ($head.Trim()) { $header += " — $head" }
        [void]$lines.Add($header)
        $hasCellEdits = [bool]$s.CellResults.Count
        $cellIssue = [bool](@($s.CellResults | Where-Object { $_.Result.HasIssue }).Count)
        # a section with rewritten table cells is never pre-ticked: whether the new cell text still fits the column and
        # row it sits in, the other cells of the row and the section's text is for the user to judge
        $userMustTick = [bool]$s.FromChat.Count -or $priorOpen.Contains($s.Section) -or $hasCellEdits -or [bool]$s.BriefGap
        $checked = (-not $s.Result.HasIssue) -and (-not $cellIssue) -and (-not $s.Confirm.Count) -and (-not $userMustTick)
        # State the real reason - a chat answer (from_chat) vs. an earlier preview of this
        # section that still had an open confirm point (which may be the model's own discarded
        # draft, not a user answer).
        $mustTickReason = if ($s.FromChat.Count -and $s.BriefGap) { 'มีข้อมูลจากคำตอบในแชต และ brief มีจุดต้องยืนยันที่ preview ไม่ได้ยกมา' } elseif ($s.FromChat.Count) { 'มีข้อมูลจากคำตอบในแชต' } elseif ($hasCellEdits) { 'มีเซลล์ตารางที่เขียนใหม่ ต้องตรวจบริบทของตาราง' } elseif ($priorOpen.Contains($s.Section)) { 'preview ก่อนหน้ามีจุดต้องยืนยันค้าง' } else { 'brief มีจุดต้องยืนยันที่ preview ไม่ได้ยกมา' }
        if ($userMustTick -and -not $checked) { [void]$notPreTicked.Add("$($s.Section) ($mustTickReason)") }
        $mark = if ($checked) { 'x' } else { ' ' }
        [void]$lines.Add("- [$mark] ใช้ $($s.Section)")
        [void]$lines.Add('')
        if ($userMustTick) {
          if ($s.FromChat.Count) {
            [void]$lines.Add('section นี้มีข้อมูลจากคำตอบของผู้ใช้ในแชต (ดูหัวข้อ "ข้อมูลจากคำตอบในแชต") ช่องนี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจข้อความแล้วติ๊กเอง')
            [void]$lines.Add('')
          } elseif ($priorOpen.Contains($s.Section)) {
            [void]$lines.Add('preview ก่อนหน้าของ section นี้มีจุดต้องยืนยันค้างอยู่ ช่องนี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจข้อความแล้วติ๊กเอง')
            [void]$lines.Add('')
          }
          if ($s.BriefGap -and $s.FromChat.Count) {
            [void]$lines.Add('brief ของไฟล์นี้มีจุดต้องยืนยันที่ preview นี้ไม่ได้ยกมาใน confirm แม้ section นี้มีข้อมูลจากคำตอบในแชต (ดูหัวข้อ "จุดต้องยืนยันจาก brief" ต้นไฟล์ ให้ผู้ใช้ตรวจว่าแต่ละข้อได้คำตอบจริงหรือยังค้าง) ช่องนี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจข้อความแล้วติ๊กเอง')
            [void]$lines.Add('')
          } elseif ($s.BriefGap) {
            [void]$lines.Add('brief ของไฟล์นี้มีจุดต้องยืนยันที่ preview นี้ไม่ได้ยกมา (ดูหัวข้อ "จุดต้องยืนยันจาก brief" ต้นไฟล์) ช่องนี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจข้อความแล้วติ๊กเอง')
            [void]$lines.Add('')
          }
          if ($hasCellEdits) {
            [void]$lines.Add('section นี้มีเซลล์ในตารางที่เขียนใหม่ (ดู "[ตาราง ¶N: ...]" ใต้ "หลัง") ต้องตรวจว่าข้อความใหม่ยังตอบหัวคอลัมน์และหัวแถวเดิม สอดคล้องกับเซลล์อื่นในแถวและเนื้อความของ section ช่องนี้จึงไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจแล้วติ๊กเอง')
            [void]$lines.Add('')
          }
        }
        # F6 (docs/v2-plan.md 3.8): the script itself states the anchor limitation whenever this
        # section has any kept anchor - not left to the model to remember to say.
        if ($s.HasKeep) {
          [void]$lines.Add('ตาราง/ส่วนที่คงไว้ (anchor) อยู่ตำแหน่งเดิม เนื้อหาย้ายข้ามส่วนนี้ไม่ได้ในรุ่นนี้')
          [void]$lines.Add('')
        }
        [void]$lines.Add('**ก่อน**')
        $textById = @{}
        $markedById = @{}
        foreach ($b in $s.Before) {
          $textById[$b.N] = $b.Text
          $markedById[$b.N] = $b.Marked
          if ($b.Keep) {
            foreach ($l in (Get-KeepPreviewLines $b.N $b.Text $b.Marked)) { [void]$lines.Add($l) }
            continue
          }
          [void]$lines.Add("$($b.Id):")
          foreach ($l in (Get-QuotedLines $b.Text)) { [void]$lines.Add($l) }
        }
        [void]$lines.Add('')
        [void]$lines.Add('**หลัง**')
        $ai = 0
        foreach ($a in $s.Blocks) {
          if ($a.Type -eq 'keep' -or $a.Type -eq 'table') {
            foreach ($l in (Get-KeepPreviewLines $a.Paragraph $textById[$a.Paragraph] $markedById[$a.Paragraph])) { [void]$lines.Add($l) }
            if ($a.Type -eq 'table') { foreach ($l in (Get-TableBlockPreviewLines $a $s.TableInfos[[int]$a.Paragraph])) { [void]$lines.Add($l) } }
            continue
          }
          $ai++
          foreach ($l in (Get-NewBlockPreviewLines $ai $a)) { [void]$lines.Add($l) }
        }
        [void]$lines.Add('')
        [void]$lines.Add('**แผนผังโครงสร้าง (¶ เดิม → ย่อหน้าใหม่)**')
        foreach ($m in $s.MapLines) { [void]$lines.Add("- $($m.From) → $($m.To -join ', ')") }
        if (-not $s.MapLines.Count -and $hasCellEdits) { [void]$lines.Add('- ไม่มีย่อหน้าใหม่ (เปลี่ยนเฉพาะข้อความในเซลล์ตาราง)') }
        [void]$lines.Add('')
        [void]$lines.Add('**ผลตรวจการคงสาระ (โค้ด คำนวณเอง ไม่รับผลจากโมเดล)**')
        foreach ($l in (Format-PreservationLines $s.Section $s.Result)) { [void]$lines.Add($l) }
        foreach ($cr in $s.CellResults) {
          foreach ($l in (Format-PreservationLines "$($s.Section) ตาราง ¶$($cr.Table) เซลล์ $($cr.Cell)" $cr.Result)) { [void]$lines.Add($l) }
        }
        [void]$lines.Add('')
        [void]$lines.Add('**ทวนโดยโมเดล (ไม่ใช่ผลตรวจด้วยโค้ด)**')
        foreach ($l in (Format-ListOrNone $s.ModelCheck)) { [void]$lines.Add($l) }
        [void]$lines.Add('')
        if ($s.FromChat.Count) {
          [void]$lines.Add('**ข้อมูลจากคำตอบในแชต (ไม่มีในต้นฉบับ)**')
          foreach ($l in (Format-ListOrNone $s.FromChat)) { [void]$lines.Add($l) }
          [void]$lines.Add('')
        }
        [void]$lines.Add('**จุดต้องยืนยัน**')
        foreach ($l in (Format-ListOrNone $s.Confirm)) { [void]$lines.Add($l) }
        [void]$lines.Add('')
      }

      if ($OutDir) {
        [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ ถ้าโฟลเดอร์ที่เก็บไฟล์นี้แชร์กับผู้อื่นหรือถูกส่งขึ้น repo ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
      } else {
        [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ เก็บอยู่ในโฟลเดอร์ HumanizerAgent\output\review ใน Documents ของคุณ ถ้า Documents ซิงก์กับ OneDrive ไฟล์นี้จะถูกซิงก์ขึ้นไปด้วย ถ้าโฟลเดอร์นี้แชร์กับผู้อื่น ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
      }
      # One machine-readable line per section (the blocks write-rewrite.ps1 builds the file from). It sits
      # inside the body, so the trailer's sha256 covers it like the rest of the file.
      [void]$lines.Add('')
      foreach ($s in $secValidated) { [void]$lines.Add((ConvertTo-PlanDataLine $s.Section $s.Blocks)) }

      $body = ($lines -join "`n").TrimEnd("`n")
      $hash = Get-BodyHash $body
      $idsJoined = ($secValidated | ForEach-Object { $_.Section }) -join ','
      $trailer = "<!-- hz-review v1 kind=`"rewrite`" items=`"$idsJoined`" sha256=`"$hash`" -->"
      $fullContent = (($body + "`n`n" + $trailer + "`n") -replace "`n", "`r`n")

      $opened = New-FileExclusive $dest
      $dest = $opened.Path
      try {
        $bom = New-Object System.Text.UTF8Encoding($true)
        $sw = New-Object System.IO.StreamWriter($opened.Stream, $bom)
        $sw.Write($fullContent)
        $sw.Dispose()
      } finally { $opened.Stream.Dispose() }

      $issueCount = @($secValidated | Where-Object { $_.Result.HasIssue -or @($_.CellResults | Where-Object { $_.Result.HasIssue }).Count }).Count
      Write-Output "สร้างไฟล์ตัวอย่างก่อนสร้างไฟล์: $dest"
      Write-Output ("section: {0} (ต้องให้คนตรวจ {1}, ไม่พบความต่าง {2})" -f $idsJoined, $issueCount, ($secValidated.Count - $issueCount))
      Write-Output "รหัสไฟล์: $($hash.Substring(0, 8))"
      if ($notPreTicked.Count) { Write-Output ("ไม่ติ๊กไว้ล่วงหน้า ให้ผู้ใช้ตรวจและติ๊กเอง: {0}" -f ($notPreTicked -join ', ')) }
      if ($relevantPoints.Count) {
        Write-Output ("จุดต้องยืนยันจาก brief ({0}) {1} ข้อ: ให้ตรวจว่าแต่ละข้อยังอยู่ใน confirm ของ preview หรือผู้ใช้ตอบในแชตแล้ว (การคงถ้อยคำเดิมของต้นฉบับไม่นับเป็นการคลายจุดนี้):" -f $briefInfo.Name, $relevantPoints.Count)
        foreach ($bc in $relevantPoints) { Write-Output "- $(Format-ConfirmPoint $bc)" }
      }
      if ($gapHasChat) {
        Write-Output ("ตรวจ: {0} มี from_chat แต่ไม่มี confirm ทั้งที่ brief มีจุดต้องยืนยันของ section นี้ การตอบ `"ถูกต้อง`" ต่อ brief ไม่ใช่คำตอบของจุดต้องยืนยัน ถ้าผู้ใช้ยังไม่ตอบเนื้อหาของจุดใด ให้ย้ายจุดนั้นกลับไปไว้ใน confirm และเอาออกจาก from_chat แล้วสร้าง preview ใหม่" -f ($gapChatSections -join ', '))
      }
      if ($mapIgnored) { Write-Output 'หมายเหตุ: ไม่ใช้ map ที่ส่งมา แผนผังโครงสร้างสร้างจาก from ให้เอง' }
      exit 0
    }

    # ---------------- review / report ----------------
    $itemsList = New-Object System.Collections.ArrayList
    if ($parsed.items) { foreach ($x in $parsed.items) { [void]$itemsList.Add($x) } }
    if (-not $itemsList.Count) { Fail 'ไม่มีรายการใน items' }

    # How the points are located: a .docx / .txt is read once (path + Ctrl+F text computed from the file itself), a .pdf (or
    # anything else) keeps the page and the heading from the JSON.
    $srcExt = [System.IO.Path]::GetExtension($srcFull).ToLowerInvariant()
    $locKind = if ($srcExt -eq '.docx') { 'docx' } elseif ($srcExt -eq '.txt' -or $srcExt -eq '.md') { 'txt' } else { 'other' }
    # Optional groups: "groups":[{"id":"A","title":"...","note":"..."}] and "group":"A" on each item. The script writes the
    # per-group summary table and the group headings itself. No groups (missing or empty) = the file is written as before.
    $groupDefs = New-Object System.Collections.ArrayList
    $groupIdSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $groupInput = Get-AsArray $parsed.groups
    if ($null -ne $parsed.groups -and $parsed.groups -isnot [System.Array] -and $parsed.groups -isnot [System.Management.Automation.PSCustomObject]) { Fail 'groups ต้องเป็นรายการของกลุ่ม' }
    $gi = 0
    foreach ($g in $groupInput) {
      $gi++
      if ($g -isnot [System.Management.Automation.PSCustomObject]) { Fail "กลุ่มที่ ${gi}: ต้องเป็นข้อมูลที่มี id และ title" }
      $gid = [string]$g.id
      if ($gid -cnotmatch '^[A-Z][A-Z0-9]{0,2}$') { Fail "กลุ่มที่ ${gi}: id ต้องเป็นตัวพิมพ์ใหญ่ 1-3 ตัว เช่น A หรือ G1" }
      if (-not $groupIdSet.Add($gid)) { Fail "กลุ่ม $gid ซ้ำ" }
      $gtitle = (Get-OneLine ([string]$g.title)).Trim()
      if (-not $gtitle) { Fail "กลุ่ม ${gid}: title ว่าง" }
      $gnote = (Get-OneLine ([string]$g.note)).Trim()
      [void]$groupDefs.Add(@{ Id = $gid; Title = $gtitle; Note = $gnote })
    }
    $hasGroups = ($groupDefs.Count -gt 0)

    $seenIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $counts = @{ must = 0; may = 0; confirm = 0 }
    $validated = New-Object System.Collections.ArrayList
    $itemGroup = @{}
    $i = 0
    foreach ($it in $itemsList) {
      $i++
      $id = [string]$it.id
      if ($id -cnotmatch '^R\d+$') { Fail "รายการที่ ${i}: id ต้องเป็นรูปแบบ R1" }
      if ($seenIds.Contains($id)) { Fail "รหัส $id ซ้ำ" }
      [void]$seenIds.Add($id)
      $cat = [string]$it.category
      if ($cat -notin @('must', 'may', 'confirm')) { Fail "${id}: category ต้องเป็น must, may หรือ confirm" }
      $itemGroupId = if ($null -ne $it.group) { [string]$it.group } else { '' }
      if ($hasGroups) {
        if (-not $itemGroupId) { Fail "${id}: ไม่มี group" }
        if (-not $groupIdSet.Contains($itemGroupId)) { Fail ("${id}: ไม่รู้จักกลุ่ม $itemGroupId (กลุ่มที่มี: " + (($groupDefs | ForEach-Object { $_.Id }) -join ', ') + ')') }
        $itemGroup[$id] = $itemGroupId
      } elseif ($itemGroupId) {
        Fail "${id}: มี group แต่ไม่มี groups ระดับบนสุด"
      }
      if ($null -eq $it.paragraph -and $null -eq $it.page) { Fail "${id}: ต้องมี paragraph หรือ page" }
      if ($locKind -ne 'other') {
        if ($null -eq $it.paragraph) { Fail "${id}: ไฟล์ .docx/.txt ต้องมี paragraph (เลข ¶ จาก extract-text.ps1)" }
        $pn = $it.paragraph -as [int]
        if ($null -eq $pn -or $pn -lt 1) { Fail "${id}: paragraph ต้องเป็นเลข ¶ ตั้งแต่ 1" }
      }
      $before = [string]$it.before
      if (-not $before.Trim()) { Fail "${id}: before ว่าง" }
      $after = [string]$it.after
      if ($Kind -eq 'review' -and -not $after.Trim() -and $cat -ne 'confirm') { Fail "${id}: after ว่าง (ว่างได้เฉพาะ confirm)" }
      $counts[$cat]++
      [void]$validated.Add($it)
    }

    # File order: with groups, group by group (groups order), items inside a group in JSON order. -Action Read and the trailer
    # follow this order.
    $ordered = New-Object System.Collections.ArrayList
    $groupItems = @{}
    if ($hasGroups) {
      foreach ($gd in $groupDefs) {
        $members = @($validated | Where-Object { $itemGroup[[string]$_.id] -ceq $gd.Id })
        if (-not $members.Count) { Fail "กลุ่ม $($gd.Id) ไม่มีรายการ" }
        $groupItems[$gd.Id] = $members
        foreach ($m in $members) { [void]$ordered.Add($m) }
      }
    } else {
      foreach ($m in $validated) { [void]$ordered.Add($m) }
    }

    $finds = @{}
    if ($locKind -ne 'other') {
      $locItems = if ($locKind -eq 'docx') { Read-Docx $srcFull } else { Read-Txt $srcFull }
      $locCorpus = Get-FindCorpus $locItems
      foreach ($it in $validated) {
        $f = Get-ItemFinder $locItems $locCorpus $it.paragraph $it.row $it.cell $it.cell_paragraph ([string]$it.before)
        if ($f.Error) { Fail "$([string]$it.id): $($f.Error)" }
        $finds[[string]$it.id] = $f
      }
    }

    $now = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $range = Get-OneLine ([string]$parsed.range)
    $mode = Get-OneLine ([string]$parsed.mode)
    $titleLine = if ($Kind -eq 'report') { "# รายงานผลตรวจ: $srcName" } else { "# รายการจุดแก้: $srcName" }
    $summary = "สรุป: ควรแก้ตามหลักภาษา $($counts.must) | แก้ได้เพื่ออ่านง่าย $($counts.may) | ต้องยืนยัน $($counts.confirm)"

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add($titleLine)
    [void]$lines.Add('')
    [void]$lines.Add("- ไฟล์ต้นฉบับ: $srcName (ไม่ได้แก้ไฟล์นี้)")
    [void]$lines.Add("- ช่วงที่ตรวจ: $range | โหมด: $mode")
    [void]$lines.Add("- สร้างเมื่อ: $now")
    [void]$lines.Add("- $summary")
    if ($locKind -ne 'other') {
      $whereOpen = if ($locKind -eq 'docx') { 'ใน Word' } else { 'ในโปรแกรมที่เปิดไฟล์' }
      [void]$lines.Add('- หาจุดในไฟล์: แต่ละรายการบอกหัวข้อและข้อความสำหรับค้นด้วย Ctrl+F ' + $whereOpen + ' ข้อความค้นตัดมาจากข้อความ “ก่อน” ของรายการนั้น (ถ้าใช้ต้นย่อหน้าแทน บรรทัดค้นจะบอกไว้) ส่วน ¶ คือเลขอ้างอิงของ skill (นับหลังสารบัญ Word ไม่แสดงเลขนี้)')
    }
    [void]$lines.Add('')
    if ($Kind -eq 'review') {
      [void]$lines.Add('## วิธีใช้')
      [void]$lines.Add('1. เปลี่ยน `[ ]` เป็น `[x]` หน้ารายการที่ต้องการ แล้วบันทึกไฟล์')
      [void]$lines.Add('2. กลับไปที่แชตแล้วพิมพ์ “ใช้ตามไฟล์” หรือพิมพ์เลขรายการ เช่น “R1, R3”')
      [void]$lines.Add('3. ข้อความที่แก้ในไฟล์นี้จะไม่ถูกนำไปใช้ ถ้าต้องการเปลี่ยนข้อความให้บอกในแชต')
      if ($hasGroups) { [void]$lines.Add('4. ตัดสินทีละกลุ่มได้: พิมพ์ในแชต เช่น “ใช้กลุ่ม A, B” หรือ “กลุ่ม A ยกเว้น R38” แล้วยืนยันตามปกติ') }
      [void]$lines.Add('')
    } else {
      [void]$lines.Add('ถ้าต้องการรายการแก้พร้อมช่องเลือก พิมพ์ “ตรวจและแก้” หรือ “ทำรายการแก้” ในแชต')
      if ($hasGroups) { [void]$lines.Add('รายการจัดเป็นกลุ่ม ขอไฟล์รายการแก้ทีละกลุ่มได้ เช่น “ทำรายการแก้กลุ่ม A”') }
      [void]$lines.Add('')
    }
    if ($OutDir) {
      [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ ถ้าโฟลเดอร์ที่เก็บไฟล์นี้แชร์กับผู้อื่นหรือถูกส่งขึ้น repo ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
    } else {
      [void]$lines.Add('ไฟล์นี้มีข้อความจากเอกสารของคุณ เก็บอยู่ในโฟลเดอร์ HumanizerAgent\output\review ใน Documents ของคุณ ถ้า Documents ซิงก์กับ OneDrive ไฟล์นี้จะถูกซิงก์ขึ้นไปด้วย ถ้าโฟลเดอร์นี้แชร์กับผู้อื่น ผู้ที่มีสิทธิ์เข้าถึงจะเห็นได้ ลบได้เมื่อไม่ใช้แล้ว')
    }
    [void]$lines.Add($AiPrivacySentence)
    [void]$lines.Add('')
    if ($hasGroups) {
      # One row per group in the groups order (the table is part of the hashed body). "มีช่องเลือก" = items with a checkbox
      # (review: a non-empty "after"); a report has no checkboxes, so it has no such column.
      $esc = { param([string]$s) $s.Replace('|', '\|') }
      $showBox = ($Kind -eq 'review')
      $sumHead = '| กลุ่ม | เรื่อง | จำนวน | ควรแก้ | แก้ได้ | ต้องยืนยัน |' + $(if ($showBox) { ' มีช่องเลือก |' } else { '' }) + ' รายการ |'
      $sumSep = '| --- | --- | ---: | ---: | ---: | ---: |' + $(if ($showBox) { ' ---: |' } else { '' }) + ' --- |'
      [void]$lines.Add('## สรุปตามกลุ่ม')
      [void]$lines.Add('')
      [void]$lines.Add($sumHead)
      [void]$lines.Add($sumSep)
      $totBox = 0
      foreach ($gd in $groupDefs) {
        $mem = @($groupItems[$gd.Id])
        $nMust = @($mem | Where-Object { [string]$_.category -eq 'must' }).Count
        $nMay = @($mem | Where-Object { [string]$_.category -eq 'may' }).Count
        $nConfirm = @($mem | Where-Object { [string]$_.category -eq 'confirm' }).Count
        $nBox = @($mem | Where-Object { Get-NonEmpty ([string]$_.after) }).Count
        $totBox += $nBox
        $topic = $gd.Title
        if ($gd.Note) { $topic += ' — ' + $gd.Note }
        $row = '| {0} | {1} | {2} | {3} | {4} | {5} |' -f $gd.Id, (& $esc $topic), $mem.Count, $nMust, $nMay, $nConfirm
        if ($showBox) { $row += " $nBox |" }
        $row += ' ' + (($mem | ForEach-Object { [string]$_.id }) -join ', ') + ' |'
        [void]$lines.Add($row)
      }
      $totRow = '| รวม | | {0} | {1} | {2} | {3} |' -f $validated.Count, $counts.must, $counts.may, $counts.confirm
      if ($showBox) { $totRow += " $totBox |" }
      $totRow += ' |'
      [void]$lines.Add($totRow)
      [void]$lines.Add('')
    }
    [void]$lines.Add('## รายการ')
    [void]$lines.Add('')
    $itemHead = if ($hasGroups) { '####' } else { '###' }
    $curGroup = $null
    foreach ($it in $ordered) {
      $id = [string]$it.id
      $f = $finds[$id]
      if ($hasGroups -and $curGroup -cne $itemGroup[$id]) {
        $curGroup = $itemGroup[$id]
        $gdef = @($groupDefs | Where-Object { $_.Id -ceq $curGroup })[0]
        [void]$lines.Add("### กลุ่ม $curGroup — $($gdef.Title)")
        [void]$lines.Add('')
      }
      if ($f) {
        $header = "$itemHead $id — $($f.Path)"
      } else {
        $loc = Get-ItemLocation $it
        $heading = Get-OneLine ([string]$it.heading)
        $header = "$itemHead $id — $loc"
        if ($heading.Trim()) { $header += " · $heading" }
      }
      [void]$lines.Add($header)
      $cat = [string]$it.category
      $after = [string]$it.after
      $hasCheckbox = ($Kind -eq 'review') -and $after.Trim()
      if ($hasCheckbox) { [void]$lines.Add("- [ ] ใช้ $id") }
      if ($f) {
        if ($f.Quote) {
          $findLabel = if ($locKind -eq 'docx') { 'ค้นใน Word (Ctrl+F)' } else { 'ค้นในไฟล์ (Ctrl+F)' }
          $findLine = "- ${findLabel}: $(Format-CodeSpan $f.Quote)"
          if ($f.Count -gt 1) { $findLine += " (พบ $($f.Count) ที่ในเอกสาร ดูหัวข้อด้านบนประกอบ)" }
          if (-not $f.BeforeFound) { $findLine += ' (ข้อความต้นย่อหน้า ไม่มีข้อความ “ก่อน” ของรายการนี้)' }
          [void]$lines.Add($findLine)
        }
        [void]$lines.Add("- อ้างอิง: $($f.Ref)")
      } else {
        $pdfQuote = Get-PdfQuote ([string]$it.before)
        if ($pdfQuote) {
          $findLabel = if ($srcExt -eq '.pdf') { 'ค้นใน PDF (Ctrl+F)' } else { 'ค้นในไฟล์ (Ctrl+F)' }
          [void]$lines.Add("- ${findLabel}: $(Format-CodeSpan $pdfQuote)")
        }
      }
      [void]$lines.Add("- ประเภท: $($catThai[$cat])")
      [void]$lines.Add('- ก่อน:')
      foreach ($ln in (Get-QuotedLines ([string]$it.before))) { [void]$lines.Add($ln) }
      if ($after.Trim()) {
        $afterLabel = if ($Kind -eq 'report') { 'ข้อเสนอ' } else { 'หลัง' }
        [void]$lines.Add("- ${afterLabel}:")
        foreach ($ln in (Get-QuotedLines $after)) { [void]$lines.Add($ln) }
      } elseif ($Kind -eq 'review') {
        [void]$lines.Add('- ไม่มีข้อเสนอ ต้องให้ผู้เขียนระบุ (ไม่มีช่องเลือก)')
      }
      $reason = Get-OneLine ([string]$it.reason)
      if ($reason.Trim()) { [void]$lines.Add("- เหตุผล: $reason") }
      [void]$lines.Add('')
    }
    $body = ($lines -join "`n").TrimEnd("`n")
    $hash = Get-BodyHash $body
    $idsJoined = ($ordered | ForEach-Object { [string]$_.id }) -join ','
    $trailer = "<!-- hz-review v1 kind=`"$Kind`" items=`"$idsJoined`" sha256=`"$hash`" -->"
    # A blank line before the trailer for readability. -Action Read hashes "lines before the trailer
    # line" through Normalize-Body, which trims trailing blank lines, so this blank line (or any extra
    # blank lines an editor adds) does not count as editing outside the checkboxes.
    $fullContent = (($body + "`n`n" + $trailer + "`n") -replace "`n", "`r`n")

    $suffix = if ($Kind -eq 'report') { 'report' } else { 'review' }
    $dest = Get-NewPath $srcFull $suffix '.md' $null $OutDir 'review'
    $opened = New-FileExclusive $dest
    $dest = $opened.Path
    try {
      $bom = New-Object System.Text.UTF8Encoding($true)
      $sw = New-Object System.IO.StreamWriter($opened.Stream, $bom)
      $sw.Write($fullContent)
      $sw.Dispose()
    } finally { $opened.Stream.Dispose() }

    $verb = if ($Kind -eq 'report') { 'สร้างไฟล์รายงาน' } else { 'สร้างไฟล์รายการ' }
    Write-Output "${verb}: $dest"
    Write-Output ("รายการ: {0} ({1} {2}, {3} {4}, {5} {6})" -f $validated.Count, $catThai.must, $counts.must, $catThai.may, $counts.may, $catThai.confirm, $counts.confirm)
    if ($hasGroups) {
      Write-Output ('กลุ่ม: ' + (($groupDefs | ForEach-Object { '{0} {1} รายการ' -f $_.Id, @($groupItems[$_.Id]).Count }) -join ', '))
    }
    Write-Output "รหัสไฟล์: $($hash.Substring(0, 8))"
    $notFoundIds = @($ordered | Where-Object { $finds[[string]$_.id] -and $finds[[string]$_.id].BeforeFound -eq $false } | ForEach-Object { [string]$_.id })
    if ($notFoundIds.Count) {
      Write-Output ('หมายเหตุ: ข้อความ “ก่อน” ของ {0} ไม่พบในย่อหน้าที่ระบุ ใช้ต้นย่อหน้าเป็นข้อความค้น ตรวจ before อีกครั้ง' -f ($notFoundIds -join ', '))
    }
    Write-Output $AiPrivacySentence
    exit 0
  } finally { Remove-JsonInput $Items }
}

# ---------------------------------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------------------------------
if (-not $Review) { Fail 'ต้องระบุ -Review' }
$rv = Read-ReviewFile $Review
if ($rv.Kind -eq 'report') { Fail 'ไฟล์นี้เป็นรายงาน ไม่มีช่องเลือก' }
if ($rv.Kind -eq 'brief') { Fail 'ไฟล์นี้เป็นสรุปความเข้าใจ (brief) ไม่มีช่องเลือก ให้อ่านเนื้อหาแล้วตอบในแชตว่า "ถูกต้อง" "แก้ brief" หรือ "ยกเลิก"' }
$fileLabel = if ($rv.Kind -eq 'rewrite') { 'ไฟล์ rewrite' } else { 'ไฟล์รายการ' }
$reviewFull = (Resolve-Path -LiteralPath $Review).ProviderPath

Write-Output "${fileLabel}: $(Split-Path $reviewFull -Leaf)"
Write-Output "รหัสไฟล์: $($rv.ShaAttr.Substring(0, 8))"
Write-Output ("รายการทั้งหมด: " + ($rv.AllIds -join ', '))
if ($rv.Kind -eq 'review') {
  # A file with groups ("### กลุ่ม A — <title>" headings, items as "#### R<n> — ..." under them): one line per group, so a
  # request like "ใช้กลุ่ม A" is turned into R-ids from the script's output instead of by guessing.
  $rdGroupId = $null; $rdGroupMap = [ordered]@{}
  foreach ($rdLine in $rv.BodyLines) {
    if ($rdLine -cmatch '^### กลุ่ม ([A-Z][A-Z0-9]{0,2}) — (.*?)\s*$') {
      $rdGroupId = $Matches[1]
      $rdGroupMap[$rdGroupId] = @{ Title = $Matches[2]; Ids = (New-Object System.Collections.ArrayList) }
    } elseif ($rdLine -cmatch '^### ') {
      $rdGroupId = $null
    } elseif ($rdGroupId -and $rdLine -cmatch '^#### (R\d+) — ') {
      [void]$rdGroupMap[$rdGroupId].Ids.Add($Matches[1])
    }
  }
  foreach ($rdKey in $rdGroupMap.Keys) {
    Write-Output ('กลุ่ม {0} ({1}): {2}' -f $rdKey, $rdGroupMap[$rdKey].Title, ($rdGroupMap[$rdKey].Ids -join ', '))
  }
}
if ($rv.Ticked.Count) { Write-Output ("เลือก: " + ($rv.Ticked -join ', ')) } else { Write-Output 'เลือก: ไม่มี' }
Write-Output "แก้ไขไฟล์ล่าสุด: $($rv.LastWrite)"
if ($rv.Unknown.Count) { Write-Output ("หมายเหตุ: ไม่รู้จัก " + ($rv.Unknown -join ', ')) }
if ($rv.Edited) { Write-Output 'หมายเหตุ: ไฟล์ถูกแก้นอกจากช่องติ๊ก จะใช้ข้อความจากแชตเท่านั้น' }
exit 0
