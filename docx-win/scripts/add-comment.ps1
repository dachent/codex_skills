param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [int]$Start,

    [Parameter(Mandatory = $true)]
    [int]$End,

    [Parameter(Mandatory = $true)]
    [string]$CommentText,

    [string]$Author = 'Codex',

    [string]$Initials = 'CX',

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

. "$PSScriptRoot\word-common.ps1"

$word = $null
$doc = $null
try {
    $word = New-WordApplication
    $doc = Open-WordDocument -Word $word -InputPath $InputPath

    if ($End -lt $Start) {
        throw 'End must be greater than or equal to Start.'
    }

    $range = $doc.Range($Start, $End)
    $comment = $doc.Comments.Add($range, $CommentText)
    try {
        if ($Author) { $comment.Author = $Author }
        if ($Initials) { $comment.Initial = $Initials }
    }
    catch {
    }

    Update-DocumentFields -Document $doc
    $saved = Save-WordDocument -Document $doc -OutputPath $OutputPath -Format 16
    Write-Output "Added comment and saved: $saved"
}
finally {
    Close-WordAutomation -Word $word -Document $doc
}
