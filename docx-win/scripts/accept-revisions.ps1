param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

. "$PSScriptRoot\word-common.ps1"

$word = $null
$doc = $null
try {
    $word = New-WordApplication
    $doc = Open-WordDocument -Word $word -InputPath $InputPath
    $doc.Revisions.AcceptAll()
    Update-DocumentFields -Document $doc
    $saved = Save-WordDocument -Document $doc -OutputPath $OutputPath -Format 16
    Write-Output "Accepted revisions and saved: $saved"
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
