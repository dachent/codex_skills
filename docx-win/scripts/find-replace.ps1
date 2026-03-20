param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$FindText,

    [Parameter(Mandatory = $true)]
    [string]$ReplaceText,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [bool]$MatchCase = $false
)

. "$PSScriptRoot\word-common.ps1"

$word = $null
$doc = $null
try {
    $word = New-WordApplication
    $doc = Open-WordDocument -Word $word -InputPath $InputPath

    $find = $doc.Content.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $wdFindContinue = 1
    $wdReplaceAll = 2
    $null = $find.Execute($FindText, $MatchCase, $false, $false, $false, $false, $true, $wdFindContinue, $false, $ReplaceText, $wdReplaceAll)

    Update-DocumentFields -Document $doc
    $saved = Save-WordDocument -Document $doc -OutputPath $OutputPath -Format 16
    Write-Output "Applied find and replace and saved: $saved"
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
