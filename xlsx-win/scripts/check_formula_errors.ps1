param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonScript = Join-Path $scriptDir "check_formula_errors.py"

function Write-JsonError {
    param(
        [string]$Message,
        [string]$ErrorKind,
        [string[]]$SearchedCandidates = @()
    )

    $payload = [ordered]@{
        status = "error"
        workbook = (Resolve-Path -LiteralPath $WorkbookPath -ErrorAction SilentlyContinue | ForEach-Object { $_.Path })
        message = $Message
        error_kind = $ErrorKind
        searched_candidates = $SearchedCandidates
    }

    $payload | ConvertTo-Json -Depth 4
}

function Test-PythonCandidate {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$Label
    )

    try {
        $null = & $Executable @Arguments -c "import openpyxl, sys; print(sys.executable)" 2>&1
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{
                Executable = $Executable
                Arguments = $Arguments
                Label = $Label
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-PythonCandidate {
    $candidates = @()

    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pythonCommand) {
        $candidates += [pscustomobject]@{
            Executable = $pythonCommand.Source
            Arguments = @()
            Label = "python"
        }
    }

    $pyCommand = Get-Command py -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pyCommand) {
        $candidates += [pscustomobject]@{
            Executable = $pyCommand.Source
            Arguments = @("-3")
            Label = "py -3"
        }
    }

    foreach ($root in @("C:\ProgramData\.pyenv\pyenv-win\versions", "$env:USERPROFILE\.pyenv\pyenv-win\versions")) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        Get-ChildItem -LiteralPath $root -Directory | Sort-Object Name -Descending | ForEach-Object {
            $pythonExe = Join-Path $_.FullName "python.exe"
            if (Test-Path -LiteralPath $pythonExe) {
                $candidates += [pscustomobject]@{
                    Executable = $pythonExe
                    Arguments = @()
                    Label = $pythonExe
                }
            }
        }
    }

    $searched = @()
    foreach ($candidate in $candidates) {
        $searched += $candidate.Label
        $result = Test-PythonCandidate -Executable $candidate.Executable -Arguments $candidate.Arguments -Label $candidate.Label
        if ($result) {
            return [pscustomobject]@{
                Candidate = $result
                Searched = $searched
            }
        }
    }

    return [pscustomobject]@{
        Candidate = $null
        Searched = $searched
    }
}

if (-not (Test-Path -LiteralPath $pythonScript)) {
    Write-JsonError -Message "check_formula_errors.py was not found next to the wrapper script." -ErrorKind "missing_python_script"
    exit 1
}

$candidateResult = Get-PythonCandidate
if ($null -eq $candidateResult.Candidate) {
    Write-JsonError -Message "No usable Python interpreter with openpyxl was found. Tried python, py -3, and installed pyenv interpreters." -ErrorKind "python_not_found" -SearchedCandidates $candidateResult.Searched
    exit 1
}

& $candidateResult.Candidate.Executable @($candidateResult.Candidate.Arguments + $pythonScript + $WorkbookPath)
exit $LASTEXITCODE
