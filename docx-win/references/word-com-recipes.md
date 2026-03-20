# Word COM recipes for Windows

Use these patterns when writing or patching PowerShell that automates Microsoft Word through COM.

## Startup and cleanup

```powershell
. "$PSScriptRoot\..\scripts\word-common.ps1"
$word = $null
$doc = $null
try {
    $word = New-WordApplication
    $doc = $word.Documents.Add()
    # work here
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
```

Always release COM objects. Do not leave `WINWORD.EXE` running.

## Create a new document with title and headings

```powershell
$doc = $word.Documents.Add()

$range = $doc.Range(0, 0)
$range.Text = "Quarterly Operating Review`r`n"
$range.Style = "Title"

$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
$range.InsertAfter("Executive Summary`r`n")
$range.Paragraphs.Last.Range.Style = "Heading 1"

$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
$range.InsertAfter("This quarter improved margins and reduced cycle time.`r`n")
$range.Paragraphs.Last.Range.Style = "Normal"
```

## Insert a page break

```powershell
$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
$range.InsertBreak(7)  # wdPageBreak
```

## Add a simple table

```powershell
$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
$table = $doc.Tables.Add($range, 3, 2)
$table.Cell(1,1).Range.Text = "Metric"
$table.Cell(1,2).Range.Text = "Value"
$table.Cell(2,1).Range.Text = "Revenue"
$table.Cell(2,2).Range.Text = "$12.4M"
$table.Cell(3,1).Range.Text = "EBITDA"
$table.Cell(3,2).Range.Text = "$3.1M"
$table.Rows.Item(1).Range.Bold = 1
```

## Add a table of contents

```powershell
$range = $doc.Range(0, 0)
$null = $doc.TablesOfContents.Add($range, $true, 1, 3)
```

After changing headings, call `Update-DocumentFields`.

## Headers, footers, and page numbers

```powershell
$section = $doc.Sections.Item(1)
$header = $section.Headers.Item(1)
$header.Range.Text = "Confidential"

$footer = $section.Footers.Item(1)
$footer.PageNumbers.Add() | Out-Null
```

## Find and replace

```powershell
$find = $doc.Content.Find
$find.ClearFormatting()
$find.Replacement.ClearFormatting()
$wdFindContinue = 1
$wdReplaceAll = 2
$null = $find.Execute("Old Name", $false, $false, $false, $false, $false, $true, $wdFindContinue, $false, "New Name", $wdReplaceAll)
```

## Tracked changes

```powershell
$doc.TrackRevisions = $true
$doc.Range(0,0).InsertAfter("Inserted with track changes on.`r`n")
```

## Comments

```powershell
$range = $doc.Paragraphs.Item(1).Range
$null = $doc.Comments.Add($range, "verify this paragraph")
```

## Accept all revisions

```powershell
$doc.Revisions.AcceptAll()
```

## Export to PDF

```powershell
$doc.ExportAsFixedFormat("C:\temp\output.pdf", 17)
```

## Update fields and repaginate

```powershell
Update-DocumentFields -Document $doc
$pages = $doc.ComputeStatistics(2)  # wdStatisticPages
```

Use this before final save on layout-sensitive deliverables.

## Images

```powershell
$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
$shape = $doc.InlineShapes.AddPicture("C:\temp\chart.png", $false, $true, $range)
$shape.Width = 320
$shape.Height = 180
```

## Convert legacy .doc to .docx

```powershell
$doc = $word.Documents.Open("C:\temp\legacy.doc")
$doc.SaveAs2("C:\temp\converted.docx", 16)
```

## Reliability notes

- Use `.docx` as the working format.
- Prefer Word built-in styles over manual formatting.
- Save to a new output path when preserving source material matters.
- Export PDF for visual QA after meaningful layout changes.
- If automation leaves a stuck `WINWORD.EXE`, close the document, quit Word, release COM objects, and retry.
