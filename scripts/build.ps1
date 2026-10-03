param([string]$Output = "artifacts\TerminPlanner.xlsm")

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$absoluteOutput = [IO.Path]::GetFullPath((Join-Path $projectRoot $Output))

if (-not $absoluteOutput.StartsWith([IO.Path]::GetFullPath($projectRoot), [StringComparison]::OrdinalIgnoreCase)) {
    throw "Output must stay inside the repository."
}

Push-Location $projectRoot
try {
    dotnet run --project tools\WorkbookBuilder\WorkbookBuilder.csproj --configuration Release -- $absoluteOutput
    if ($LASTEXITCODE -ne 0) {
        throw "Workbook builder failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}
