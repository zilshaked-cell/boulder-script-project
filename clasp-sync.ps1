param(
    [switch]$StatusOnly
)

Push-Location $PSScriptRoot
try {
    $statusResult = & clasp status
    $statusCode = $LASTEXITCODE
    $statusResult | ForEach-Object { Write-Host $_ }
    if ($statusCode -ne 0) { exit $statusCode }
    if ($StatusOnly) { exit 0 }

    $pushResult = & clasp push
    $pushCode = $LASTEXITCODE
    $pushResult | ForEach-Object { Write-Host $_ }
    exit $pushCode
} finally {
    Pop-Location
}
