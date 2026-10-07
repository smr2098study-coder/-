param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputDocx,
    [Parameter(Mandatory=$true)][string]$ReportJson
)

$ErrorActionPreference = 'Stop'
$taskInput = (Resolve-Path -LiteralPath $InputDocx).Path
$taskOutput = [System.IO.Path]::GetFullPath($OutputDocx)
$taskReport = [System.IO.Path]::GetFullPath($ReportJson)
if ($taskInput -eq $taskOutput) { throw 'InputDocx and OutputDocx must be different paths.' }
if (Test-Path -LiteralPath $taskOutput) { throw "Output already exists: $taskOutput" }
if (Test-Path -LiteralPath $taskReport) { throw "Report already exists: $taskReport" }
if ([System.IO.Path]::GetExtension($taskInput).ToLowerInvariant() -ne '.docx') { throw 'Input must be a DOCX file.' }

$taskOutputParent = Split-Path -Parent $taskOutput
$taskReportParent = Split-Path -Parent $taskReport
New-Item -ItemType Directory -Path $taskOutputParent -Force | Out-Null
New-Item -ItemType Directory -Path $taskReportParent -Force | Out-Null
Copy-Item -LiteralPath $taskInput -Destination $taskOutput
$taskDynamicCopy = Join-Path ([System.IO.Path]::GetTempPath()) ("shimen-mode-b-" + [guid]::NewGuid().ToString('N') + '.docx')

function Update-AllWordFields($taskDocument) {
    if ($taskDocument.Fields.Count -gt 0) { [void]$taskDocument.Fields.Update() }
    foreach ($taskStory in $taskDocument.StoryRanges) {
        $taskRange = $taskStory
        while ($null -ne $taskRange) {
            if ($taskRange.Fields.Count -gt 0) { [void]$taskRange.Fields.Update() }
            try {
                foreach ($taskShape in $taskRange.ShapeRange) {
                    if ($taskShape.TextFrame.HasText -ne 0 -and $taskShape.TextFrame.TextRange.Fields.Count -gt 0) {
                        [void]$taskShape.TextFrame.TextRange.Fields.Update()
                    }
                }
            } catch {}
            try { $taskRange = $taskRange.NextStoryRange } catch { $taskRange = $null }
        }
    }
    foreach ($taskSection in $taskDocument.Sections) {
        foreach ($taskHeader in $taskSection.Headers) {
            if ($taskHeader.Exists -and $taskHeader.Range.Fields.Count -gt 0) { [void]$taskHeader.Range.Fields.Update() }
            try {
                foreach ($taskShape in $taskHeader.Shapes) {
                    if ($taskShape.TextFrame.HasText -ne 0 -and $taskShape.TextFrame.TextRange.Fields.Count -gt 0) {
                        [void]$taskShape.TextFrame.TextRange.Fields.Update()
                    }
                }
            } catch {}
        }
        foreach ($taskFooter in $taskSection.Footers) {
            if ($taskFooter.Exists -and $taskFooter.Range.Fields.Count -gt 0) { [void]$taskFooter.Range.Fields.Update() }
            try {
                foreach ($taskShape in $taskFooter.Shapes) {
                    if ($taskShape.TextFrame.HasText -ne 0 -and $taskShape.TextFrame.TextRange.Fields.Count -gt 0) {
                        [void]$taskShape.TextFrame.TextRange.Fields.Update()
                    }
                }
            } catch {}
        }
    }
}

function Get-NoterefState($taskDocument) {
    $taskRows = @()
    foreach ($taskStory in $taskDocument.StoryRanges) {
        $taskRange = $taskStory
        while ($null -ne $taskRange) {
            foreach ($taskField in $taskRange.Fields) {
                $taskCode = [string]$taskField.Code.Text
                if ($taskCode -match '(?i)\bNOTEREF\s+"?([^\s"\\]+)') {
                    $taskRows += [pscustomobject]@{
                        target = $matches[1]
                        code = $taskCode.Trim()
                        result = ([string]$taskField.Result.Text).Trim()
                    }
                }
            }
            try { $taskRange = $taskRange.NextStoryRange } catch { $taskRange = $null }
        }
    }
    return @($taskRows)
}

function Get-EndnoteState($taskDocument) {
    $taskRows = @()
    for ($taskIndex = 1; $taskIndex -le $taskDocument.Endnotes.Count; $taskIndex++) {
        $taskNote = $taskDocument.Endnotes.Item($taskIndex)
        $taskRows += [pscustomobject]@{
            body = ([string]$taskNote.Range.Text).Trim()
            reference = ([string]$taskNote.Reference.Text).Trim()
        }
    }
    return @($taskRows)
}

function Get-FirstInteger([string]$taskText) {
    if ($taskText -match '\d+') { return [int]$matches[0] }
    return $null
}

$taskWord = $null
$taskDocument = $null
$taskResult = [ordered]@{
    word_field_update_validation = 'failed'
    save_close_reopen_passed = $false
    dynamic_insert_test_passed = $false
    dynamic_delete_test_passed = $false
    navigation_ok = $false
    noteref_count = 0
    broken_noteref_count = 0
    error = $null
}

try {
    $taskWord = New-Object -ComObject Word.Application
    $taskWord.Visible = $false
    $taskWord.DisplayAlerts = 0
    $taskDocument = $taskWord.Documents.Open($taskOutput, $false, $false)
    Update-AllWordFields $taskDocument
    $taskBeforeSave = Get-NoterefState $taskDocument
    $taskResult.noteref_count = $taskBeforeSave.Count
    $taskResult.broken_noteref_count = @($taskBeforeSave | Where-Object {
        $_.result -match '(?i)Error!|错误！|未找到引用源|Reference source not found'
    }).Count
    $taskResult.navigation_ok = @($taskBeforeSave | Where-Object { $_.code -notmatch '(?i)(?:^|\s)\\h(?:\s|$)' }).Count -eq 0
    $taskDocument.Save()
    $taskDocument.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDocument) | Out-Null
    $taskDocument = $null

    $taskDocument = $taskWord.Documents.Open($taskOutput, $false, $false)
    $taskAfterReopen = Get-NoterefState $taskDocument
    $taskResult.save_close_reopen_passed = (
        ($taskAfterReopen | ConvertTo-Json -Compress -Depth 5) -eq
        ($taskBeforeSave | ConvertTo-Json -Compress -Depth 5)
    ) -and (@($taskAfterReopen | Where-Object {
        $_.result -match '(?i)Error!|错误！|未找到引用源|Reference source not found'
    }).Count -eq 0)
    $taskDocument.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDocument) | Out-Null
    $taskDocument = $null

    Copy-Item -LiteralPath $taskOutput -Destination $taskDynamicCopy
    $taskDocument = $taskWord.Documents.Open($taskDynamicCopy, $false, $false)
    Update-AllWordFields $taskDocument
    $taskNotesBefore = Get-EndnoteState $taskDocument
    $taskFieldsBefore = Get-NoterefState $taskDocument
    if ($taskNotesBefore.Count -lt 1) { throw 'Dynamic test requires at least one true Endnote.' }
    $taskInsert = $taskDocument.Content.Duplicate
    $taskInsert.Collapse(1)
    $taskMissing = [Type]::Missing
    [void]$taskDocument.Endnotes.Add($taskInsert, $taskMissing, '__SHIMEN_DYNAMIC_RENUMBER_TEST__')
    Update-AllWordFields $taskDocument
    $taskNotesAfter = Get-EndnoteState $taskDocument
    $taskFieldsAfter = Get-NoterefState $taskDocument

    $taskNotesShifted = $true
    foreach ($taskOldNote in $taskNotesBefore) {
        $taskNewNote = @($taskNotesAfter | Where-Object { $_.body -eq $taskOldNote.body } | Select-Object -First 1)
        if ($taskNewNote.Count -ne 1) { $taskNotesShifted = $false; break }
        $taskOldNumber = Get-FirstInteger $taskOldNote.reference
        $taskNewNumber = Get-FirstInteger $taskNewNote[0].reference
        if ($null -eq $taskOldNumber -or $null -eq $taskNewNumber -or $taskNewNumber -ne ($taskOldNumber + 1)) {
            $taskNotesShifted = $false; break
        }
    }
    $taskFieldsShifted = $true
    if ($taskFieldsBefore.Count -ne $taskFieldsAfter.Count) { $taskFieldsShifted = $false }
    for ($taskIndex = 0; $taskIndex -lt [Math]::Min($taskFieldsBefore.Count, $taskFieldsAfter.Count); $taskIndex++) {
        $taskOldNumber = Get-FirstInteger $taskFieldsBefore[$taskIndex].result
        $taskNewNumber = Get-FirstInteger $taskFieldsAfter[$taskIndex].result
        if ($null -eq $taskOldNumber -or $null -eq $taskNewNumber -or $taskNewNumber -ne ($taskOldNumber + 1)) {
            $taskFieldsShifted = $false; break
        }
    }
    $taskResult.dynamic_insert_test_passed = $taskNotesShifted -and $taskFieldsShifted

    $taskTestEndnote = $null
    for ($taskIndex = 1; $taskIndex -le $taskDocument.Endnotes.Count; $taskIndex++) {
        $taskCandidate = $taskDocument.Endnotes.Item($taskIndex)
        if (([string]$taskCandidate.Range.Text) -match '__SHIMEN_DYNAMIC_RENUMBER_TEST__') {
            $taskTestEndnote = $taskCandidate
            break
        }
    }
    if ($null -eq $taskTestEndnote) { throw 'Could not locate the temporary Endnote for deletion regression.' }
    $taskTestEndnote.Delete()
    Update-AllWordFields $taskDocument
    $taskNotesAfterDelete = Get-EndnoteState $taskDocument
    $taskFieldsAfterDelete = Get-NoterefState $taskDocument
    $taskResult.dynamic_delete_test_passed = (
        ($taskNotesAfterDelete | ConvertTo-Json -Compress -Depth 5) -eq
        ($taskNotesBefore | ConvertTo-Json -Compress -Depth 5)
    ) -and (
        ($taskFieldsAfterDelete | ConvertTo-Json -Compress -Depth 5) -eq
        ($taskFieldsBefore | ConvertTo-Json -Compress -Depth 5)
    )

    $taskDocument.Save()
    $taskDocument.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDocument) | Out-Null
    $taskDocument = $null
    $taskDocument = $taskWord.Documents.Open($taskDynamicCopy, $false, $false)
    Update-AllWordFields $taskDocument
    $taskDynamicReopen = Get-NoterefState $taskDocument
    $taskDynamicReopenOk = (($taskDynamicReopen | ConvertTo-Json -Compress -Depth 5) -eq
                            ($taskFieldsAfterDelete | ConvertTo-Json -Compress -Depth 5))
    $taskDocument.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDocument) | Out-Null
    $taskDocument = $null

    $taskResult.dynamic_delete_test_passed = $taskResult.dynamic_delete_test_passed -and $taskDynamicReopenOk
    if ($taskResult.save_close_reopen_passed -and $taskResult.dynamic_insert_test_passed -and
        $taskResult.dynamic_delete_test_passed -and
        $taskResult.navigation_ok -and $taskResult.broken_noteref_count -eq 0) {
        $taskResult.word_field_update_validation = 'passed'
    }
} catch {
    $taskResult.error = $_.Exception.Message
} finally {
    if ($null -ne $taskDocument) {
        try { $taskDocument.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskDocument) | Out-Null
    }
    if ($null -ne $taskWord) {
        try { $taskWord.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($taskWord) | Out-Null
    }
    if (Test-Path -LiteralPath $taskDynamicCopy) { Remove-Item -LiteralPath $taskDynamicCopy -Force }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}

$taskResult | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $taskReport -Encoding UTF8
$taskResult | ConvertTo-Json -Depth 8
if ($taskResult.word_field_update_validation -ne 'passed') { exit 1 }
