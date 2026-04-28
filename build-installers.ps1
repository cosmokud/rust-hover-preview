$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

if (-not (Get-Command cargo-packager -ErrorAction SilentlyContinue)) {
    Write-Host "cargo-packager not found. Installing with cargo..."
    cargo install cargo-packager --locked
}

Write-Host "Building NSIS (.exe) and WiX (.msi) installers..."
cargo packager --release

$outputDir = Join-Path $repoRoot "target\packager"
$resolvedOutputDir = [System.IO.Path]::GetFullPath($outputDir)

Write-Host ""
Write-Host "Installers generated in:"
Write-Host $resolvedOutputDir

Get-ChildItem -Path $resolvedOutputDir -File |
    Where-Object { $_.Extension -in ".exe", ".msi" } |
    ForEach-Object { Write-Host (" - " + $_.FullName) }
