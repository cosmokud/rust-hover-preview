$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

if (-not (Get-Command cargo-packager -ErrorAction SilentlyContinue)) {
    Write-Host "cargo-packager not found. Installing with cargo..."
    cargo install cargo-packager --locked
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$localAppData = [System.Environment]::GetFolderPath("LocalApplicationData")
$packagerToolsRoot = Join-Path $localAppData ".cargo-packager"
try {
    New-Item -ItemType Directory -Force -Path $packagerToolsRoot | Out-Null
}
catch {
    Write-Error @"
cargo-packager needs write access to its Windows tool cache:
$packagerToolsRoot

Windows denied access before installer generation could start. Run this script from
a normal, non-sandboxed PowerShell session for your user, or grant your current
user write access to Local AppData.

Original error:
$($_.Exception.Message)
"@
    exit 1
}

Write-Host "Using cargo-packager tool cache:"
Write-Host $packagerToolsRoot

Write-Host "Building NSIS installer..."
cargo packager --release --formats nsis
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$outputDir = Join-Path $repoRoot "target\packager"
$resolvedOutputDir = [System.IO.Path]::GetFullPath($outputDir)

Write-Host ""
Write-Host "Installers generated in:"
Write-Host $resolvedOutputDir

if (-not (Test-Path -LiteralPath $resolvedOutputDir)) {
    Write-Warning "Output directory was not created. No installers were found."
    exit 0
}

# Look specifically for the NSIS executable
$installers = Get-ChildItem -LiteralPath $resolvedOutputDir -File |
    Where-Object { $_.Extension -eq ".exe" }

if (-not $installers) {
    Write-Warning "No .exe installers were found in the output directory."
    exit 0
}

$installers | ForEach-Object { Write-Host (" - " + $_.FullName) }