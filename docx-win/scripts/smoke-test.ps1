param(
    [string]$WorkingDirectory = (Join-Path $env:TEMP 'docx-win-smoke-test')
)

. "$PSScriptRoot\word-common.ps1"

$word = $null
$doc = $null
$docxPath = Resolve-FullPath -Path (Join-Path $WorkingDirectory 'smoke-test.docx') -AllowMissing
$pdfPath = Resolve-FullPath -Path (Join-Path $WorkingDirectory 'smoke-test.pdf') -AllowMissing

Ensure-ParentDirectory -Path $docxPath
if (-not (Test-Path -LiteralPath $WorkingDirectory)) {
    New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
}

try {
    $word = New-WordApplication
    $doc = $word.Documents.Add()

    $title = $doc.Range(0, 0)
    $title.Text = "DOCX Win Smoke Test`r`n"
    $title.Style = 'Title'

    $body = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $body.InsertAfter("This file validates Word COM automation, tracked changes, comments, and PDF export.`r`n")
    $doc.Paragraphs.Item(2).Range.Style = 'Normal'

    $body = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $table = $doc.Tables.Add($body, 2, 2)
    $table.Cell(1,1).Range.Text = 'Check'
    $table.Cell(1,2).Range.Text = 'Status'
    $table.Cell(2,1).Range.Text = 'Word COM'
    $table.Cell(2,2).Range.Text = 'OK'
    $table.Rows.Item(1).Range.Bold = 1

    $afterTable = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $afterTable.InsertBreak(7)
    $afterTable = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $afterTable.InsertAfter("Second page content for pagination verification.`r`n")
    $doc.Paragraphs.Last.Range.Style = 'Heading 1'

    $doc.TrackRevisions = $true
    $editRange = $doc.Paragraphs.Item(2).Range
    $editRange.Find.ClearFormatting()
    $editRange.Find.Replacement.ClearFormatting()
    $wdFindContinue = 1
    $wdReplaceAll = 2
    $null = $editRange.Find.Execute('validates', $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, 'confirms', $wdReplaceAll)

    $commentRange = $doc.Paragraphs.Item(2).Range
    $comment = $doc.Comments.Add($commentRange, 'Smoke test comment')
    try {
        $comment.Author = 'Codex'
        $comment.Initial = 'CX'
    }
    catch {
    }

    $doc.Revisions.AcceptAll()
    Update-DocumentFields -Document $doc
    $pageCount = $doc.ComputeStatistics(2)
    if ($pageCount -lt 2) {
        throw "Smoke test expected at least 2 pages, found $pageCount"
    }

    $null = Save-WordDocument -Document $doc -OutputPath $docxPath -Format 16
    $null = Export-WordPdf -Document $doc -OutputPath $pdfPath

    if (-not (Test-Path -LiteralPath $docxPath)) {
        throw "Smoke test DOCX was not created: $docxPath"
    }
    if (-not (Test-Path -LiteralPath $pdfPath)) {
        throw "Smoke test PDF was not created: $pdfPath"
    }

    Write-Output "Smoke test passed. DOCX: $docxPath"
    Write-Output "Smoke test passed. PDF:  $pdfPath"
    Write-Output "Smoke test page count: $pageCount"
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
