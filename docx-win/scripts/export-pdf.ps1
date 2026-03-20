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
    $doc = Open-WordDocument -Word $word -InputPath $InputPath -ReadOnly $true
    Update-DocumentFields -Document $doc
    $pdf = Export-WordPdf -Document $doc -OutputPath $OutputPath
    Write-Output "Exported PDF: $pdf"
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
