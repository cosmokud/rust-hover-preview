# The release build that guarantees the exe at the shared target dir is *this* checkout's.
#
# With `git config rhp.targetDir` set, every checkout on the machine links into one directory, and
# cargo keys this package's unit by the checkout's own source path: each checkout keeps its own
# fingerprint of it, stays fresh against its own last build, and cargo never relinks. So the one
# `release/rust-hover-preview.exe` on that path holds whichever checkout linked last, which is why
# launching it after a `cargo build --release` in a different tree runs that tree's code.
#
# `build.rs` stamps the commit into the executable's version block, so the file can be asked what
# it is: `(Get-Item <exe>).VersionInfo.Comments`. A build whose stamp is not this checkout's HEAD
# touches `src/main.rs`, which dirties this checkout's own fingerprint through the
# `cargo:rerun-if-changed=src` the build script prints, and compiles. After the build the stamp is
# read again and a mismatch is an error rather than a binary to launch: an exe that cannot say it
# is this checkout's is the bug this script exists for.
#
# Nothing here is interactive: `cargo metadata` and `git` answer into variables, cargo's own output
# is the console's, and `-Run` starts the app in place of a shell one-liner.
[CmdletBinding()]
param(
    # Start the exe once it is this checkout's. The replacement for `start` in a shell one-liner.
    [switch]$Run,

    # Everything after the switches goes to `cargo build --release` unchanged.
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Rest
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

if (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue) {
    $PSStyle.OutputRendering = "PlainText"
}
$env:CARGO_TERM_COLOR = "never"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

# Seconds with one decimal and a full stop, whatever the machine's culture formats with.
function Format-Elapsed([TimeSpan]$Span) {
    $seconds = $Span.TotalSeconds.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
    return "$seconds" + "s"
}

# The same answer `build.rs` gives, byte for byte: `unknown` where git cannot say, so a source
# tarball and a repository still agree about what a stamp that names no commit looks like.
function Get-Commit {
    $commit = $null
    try {
        $commit = & git rev-parse --short HEAD 2>$null
    }
    catch {
        return "unknown"
    }
    if ($LASTEXITCODE -ne 0 -or -not $commit) {
        return "unknown"
    }
    return "$commit".Trim()
}

# The stamp the file carries, or nothing where there is no file yet: a first build is not a foreign
# one, but it has to compile either way.
function Get-Stamp([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }
    return "" + (Get-Item -LiteralPath $Path).VersionInfo.Comments
}

$metadata = & cargo metadata --format-version 1 --no-deps | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
$exe = Join-Path (Join-Path $metadata.target_directory "release") "rust-hover-preview.exe"

$head = Get-Commit
$stamp = Get-Stamp $exe

if ($stamp -ne $head) {
    Write-Host ("build-release: the exe was '{0}' and this checkout is '{1}', so it is rebuilt" -f $stamp, $head)
    # Dirt this checkout's own fingerprint, which is what makes cargo compile again: the artifact's
    # mtime does not count, and neither does another checkout's recorded inputs.
    #
    # Two files, because the stamp is the *build script's* output and the crate is only what carries
    # it: another checkout's build may already have recorded this `build.rs` as unchanged, in which
    # case recompiling the crate alone relinks the resource of that older script — an exe with no
    # stamp in it, which is the mismatch this block is here to clear.
    (Get-Item -LiteralPath (Join-Path $repoRoot "src\main.rs")).LastWriteTime = Get-Date
    (Get-Item -LiteralPath (Join-Path $repoRoot "build.rs")).LastWriteTime = Get-Date
}

$clock = [System.Diagnostics.Stopwatch]::StartNew()
& cargo build --release @Rest | Out-Host
$code = $LASTEXITCODE
$clock.Stop()

if ($code -ne 0) {
    Write-Host ("build-release: cargo build --release failed after {0}" -f (Format-Elapsed $clock.Elapsed))
    exit $code
}

$stamp = Get-Stamp $exe
if ($stamp -ne $head) {
    throw ("the build left an exe stamped '{0}' where this checkout is '{1}': it is not this tree's binary, so it is not launched" -f $stamp, $head)
}

Write-Host ("build-release: {0} is '{1}', built in {2}" -f $exe, $stamp, (Format-Elapsed $clock.Elapsed))
if ($Run) {
    Start-Process -FilePath $exe
}
exit 0
