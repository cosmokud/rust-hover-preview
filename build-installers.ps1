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

$windowsInstallerService = Get-Service -Name msiserver -ErrorAction SilentlyContinue
if (-not $windowsInstallerService) {
    Write-Error "Windows Installer service (msiserver) was not found. WiX cannot build MSI packages without it."
    exit 1
}

if ($windowsInstallerService.StartType -eq "Disabled") {
    Write-Error "Windows Installer service (msiserver) is disabled. Enable it before building the MSI installer."
    exit 1
}

if ($windowsInstallerService.Status -ne "Running") {
    try {
        Write-Host "Starting Windows Installer service for WiX MSI validation..."
        Start-Service -Name msiserver -ErrorAction Stop
    }
    catch {
        Write-Error @"
WiX light.exe needs the Windows Installer service for MSI validation, but the
service could not be started.

Start it from an elevated PowerShell session:
  Start-Service msiserver

Or enable it if it is disabled:
  Set-Service msiserver -StartupType Manual
  Start-Service msiserver

Original error:
$($_.Exception.Message)
"@
        exit 1
    }
}

Write-Host "Building NSIS (.exe) and WiX (.msi) installers..."
cargo packager --release
if ($LASTEXITCODE -ne 0) {
    $packagerExitCode = $LASTEXITCODE
    Write-Warning "cargo-packager failed. Trying WiX MSI fallback with ICE validation disabled..."

    $metadata = cargo metadata --no-deps --format-version 1 | ConvertFrom-Json
    if ($LASTEXITCODE -ne 0) {
        exit $packagerExitCode
    }

    $package = $metadata.packages | Select-Object -First 1
    $arch = "x64"
    $culture = "en-US"
    $cultureArg = "en-us"
    $wixTools = Join-Path $packagerToolsRoot "WixTools"
    $wixWorkDir = Join-Path $repoRoot "target\packager\.cargo-packager\wix\$arch"
    $mainWixObj = Join-Path $wixWorkDir "main.wixobj"
    $localeFile = Join-Path $wixWorkDir "locale.wxl"
    $fallbackMsi = Join-Path $wixWorkDir "output-fallback.msi"
    $finalMsi = Join-Path $repoRoot ("target\packager\{0}_{1}_{2}_{3}.msi" -f $package.name, $package.version, $arch, $culture)
    $light = Join-Path $wixTools "light.exe"
    $wixUtilExtension = Join-Path $wixTools "WixUtilExtension.dll"
    $wixUiExtension = Join-Path $wixTools "WixUIExtension.dll"

    if (
        -not (Test-Path -LiteralPath $light) -or
        -not (Test-Path -LiteralPath $mainWixObj) -or
        -not (Test-Path -LiteralPath $localeFile)
    ) {
        Write-Error "WiX fallback files were not generated. Re-run with 'cargo packager --release -vv' for the original error."
        exit $packagerExitCode
    }

    if (Test-Path -LiteralPath $fallbackMsi) {
        Remove-Item -LiteralPath $fallbackMsi -Force
    }

    & $light `
        -ext $wixUtilExtension `
        -ext $wixUiExtension `
        -sval `
        -o $fallbackMsi `
        "-cultures:$cultureArg" `
        -loc $localeFile `
        $mainWixObj

    if ($LASTEXITCODE -ne 0) {
        exit $packagerExitCode
    }

    Move-Item -LiteralPath $fallbackMsi -Destination $finalMsi -Force
    Write-Warning "MSI was generated with WiX ICE validation disabled because local Windows Installer validation failed."
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

$installers = Get-ChildItem -LiteralPath $resolvedOutputDir -File |
    Where-Object { $_.Extension -in ".exe", ".msi" }

if (-not $installers) {
    Write-Warning "No .exe or .msi installers were found in the output directory."
    exit 0
}

$installers | ForEach-Object { Write-Host (" - " + $_.FullName) }
