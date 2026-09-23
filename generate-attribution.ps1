$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

# LICENSES/ and THIRD-PARTY.md are generated from the dependency tree rather than kept
# by hand, so this is what refreshes them: run it after anything changes Cargo.lock and
# commit what it writes.
if (-not (Get-Command cargo-tribute -ErrorAction SilentlyContinue)) {
    Write-Host "cargo-tribute not found. Installing with cargo..."
    cargo install cargo-tribute --locked
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

Write-Host "Generating third-party attribution..."
cargo tribute
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$changed = git status --short -- LICENSES THIRD-PARTY.md

if ($changed) {
    Write-Host ""
    Write-Host "Changed, and to be committed:"
    $changed | ForEach-Object { Write-Host (" " + $_) }
}
else {
    Write-Host "Already current; nothing to commit."
}
