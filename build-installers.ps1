$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

if (-not (Get-Command cargo-packager -ErrorAction SilentlyContinue)) {
    Write-Host "cargo-packager not found. Installing with cargo..."
    cargo install cargo-packager --locked --version 0.11.8
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

# The `github` profile is what ships: fat LTO across the whole dependency graph rather
# than the thin LTO the local `cargo build --release` uses, and it puts its binary in
# `target/github` — which is where the packaging metadata looks for it.
Write-Host "Building NSIS installer..."

# The packaging pipeline stays with this checkout: `binaries-dir` and `out-dir` in Cargo.toml
# are relative to it, and the CI workflow uploads from those same places. cargo-packager
# 0.11.8 reads them from Cargo.toml and ignores --binaries-dir/--out-dir, so a machine-local
# shared target dir is opted out of here rather than followed. CI is unaffected: there the
# default target dir is this same directory.
$env:CARGO_TARGET_DIR = Join-Path $repoRoot "target"
cargo packager --profile github --formats nsis
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$outputDir = Join-Path $repoRoot "target\packager"
$resolvedOutputDir = [System.IO.Path]::GetFullPath($outputDir)

# The portable build: the exe and the hook's DLL together, because the DLL has to sit beside the
# executable to be loaded (`ui::preview_window::custom_preview`). One file of it, or the binary
# alone, is a preview that silently falls back to Invisible Layered.
$version = (& cargo metadata --format-version 1 --no-deps | ConvertFrom-Json).packages |
    Where-Object { $_.name -eq "rust-hover-preview" } |
    Select-Object -ExpandProperty version
if ($LASTEXITCODE -ne 0 -or -not $version) {
    Write-Error "cargo metadata could not say what version this tree is"
    exit 1
}

$portable = Join-Path $resolvedOutputDir "rust-hover-preview_${version}_x64-portable.zip"
if (Test-Path -LiteralPath $portable) {
    Remove-Item -LiteralPath $portable
}
$portableFiles = @("rust-hover-preview.exe", "custom_preview_hook.dll") | ForEach-Object {
    Join-Path $repoRoot "target\github\$_"
}
foreach ($file in $portableFiles) {
    if (-not (Test-Path -LiteralPath $file)) {
        Write-Error "$file was not built, so the portable zip would not be one"
        exit 1
    }
}
Compress-Archive -LiteralPath $portableFiles -DestinationPath $portable -CompressionLevel Optimal
# The marker is what makes the unzipped copy keep `config.ini`, `theme`, `engines`, `webview` and
# `libreoffice` beside itself rather than scattering them through the user's profile (see
# `paths::data_root`). An empty file says all of it, and it has to be in the archive from the
# start: a user who unzips last build's zip gets no marker and therefore no portable copy.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::Open($portable, 'Update')
try {
    $zipEntry = $archive.CreateEntry('.portable')
    $zipEntry.LastWriteTime = [System.DateTimeOffset]::Now
}
finally {
    $archive.Dispose()
}
Write-Host ""
Write-Host "Portable zip:"
Write-Host (" - " + $portable)

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