# The gates this repository keeps, in the order CI runs them, behind one command.
#
# Each of them is a line somebody has to remember -- a formatting check whose `--check` is the
# half that reports rather than rewrites, a clippy invocation that names `--all-targets` so the
# `src/**/tests` modules are reached at all, a test run whose filters are the only way to look
# at one module of a suite of 1,450 -- and a line somebody has to remember is one that gets
# reconstructed slightly wrong. This is the reconstruction done once: `pwsh -File
# scripts/verify.ps1` is the whole gate, and each switch below is one of its halves.
#
# Nothing here is interactive: no prompt, no window, no pager. Every stage is a cargo
# invocation whose output is printed as it happens, and the first failure ends the run with
# that command's own exit code rather than continuing to the next stage.
[CmdletBinding(DefaultParameterSetName = "Full")]
param(
    # `cargo check --all-targets` and the three targets under `tests/`, and nothing else: about
    # a second warm, which is the loop for a change that is not done yet. It does not format or
    # lint -- those are cheap but they fail on a tree that does not compile, and this order
    # says so first.
    [Parameter(ParameterSetName = "Fast")]
    [switch]$Fast,

    # The part of the suite that needs no interactive desktop: the run to ask for from a
    # service, a container, or a session with no window station. Formatting and lints still
    # run, because neither needs a desktop.
    [Parameter(ParameterSetName = "Full")]
    [switch]$Pure,

    # One module's tests instead of the suite, by the same filter `cargo test <filter>`
    # matches: a substring of the test's path, so `pin_zoom` is 49 of them.
    [Parameter(ParameterSetName = "Filter")]
    [string]$Filter,

    # Formatting and lints only. For a change that is a document, a workflow or a hook, where
    # the suite has nothing to say and a compile is the whole cost.
    [Parameter(ParameterSetName = "Full")]
    [switch]$NoTest
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# What this prints is read in a log as often as in a terminal, and an escape sequence in a log
# is noise. `$PSStyle` is PowerShell 7's and 5.1 has no colour to turn off; `CARGO_TERM_COLOR`
# is for the children, which cargo would otherwise colour wherever a terminal is attached.
if (Get-Variable -Name PSStyle -ErrorAction SilentlyContinue) {
    $PSStyle.OutputRendering = "PlainText"
}
$env:CARGO_TERM_COLOR = "never"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

# Seconds with one decimal and a full stop, whatever the machine's culture formats with: a log
# is read beside other logs, and `1,5s` beside `1.5s` is a difference with no meaning.
function Format-Elapsed([TimeSpan]$Span) {
    $seconds = $Span.TotalSeconds.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
    return "$seconds" + "s"
}

# One stage, one line. Cargo's own output is left on the console above it -- a cold build is
# minutes, and a stage that says nothing until it is done cannot be told from one that hung.
function Invoke-Stage {
    param([int]$Number, [int]$Total, [string]$Label, [string[]]$CargoArgs)

    $prefix = "[{0}/{1}] {2,-32}" -f $Number, $Total, $Label
    $clock = [System.Diagnostics.Stopwatch]::StartNew()
    # `Out-Host` rather than the pipeline: a native command's standard output is otherwise the
    # function's output, and the caller would take cargo's own log lines as the exit code it
    # asked for. Stderr is already the console's, and both stream as they are written.
    & cargo @CargoArgs | Out-Host
    $code = $LASTEXITCODE
    $clock.Stop()

    if ($code -eq 0) {
        Write-Host ("{0} ok {1,8}" -f $prefix, (Format-Elapsed $clock.Elapsed))
    }
    else {
        Write-Host ("{0} FAILED {1,8}" -f $prefix, (Format-Elapsed $clock.Elapsed))
        Write-Host ("         cargo {0}" -f ($CargoArgs -join " "))
    }

    return $code
}

# The three targets under `tests/`: a path guard over every tracked file and every commit
# message, the icon and manifest block of the built executable, and the source scan that keeps
# the suite's windows off the desktop. They are the whole of -Fast's test half, and they need
# no desktop and no engine.
$integration = @(
    "test",
    "--test", "no_path_leaks",
    "--test", "executable_icon",
    "--test", "no_window_reaches_the_desktop"
)

$stages = @()

if ($Fast) {
    $stages += [pscustomobject]@{ Label = "check all targets"; Args = @("check", "--all-targets") }
    $stages += [pscustomobject]@{ Label = "integration targets"; Args = $integration }
}
else {
    $stages += [pscustomobject]@{ Label = "formatting"; Args = @("fmt", "--all", "--", "--check") }
    $stages += [pscustomobject]@{ Label = "clippy all targets"; Args = @("clippy", "--all-targets", "--", "-D", "warnings") }

    if (-not $NoTest) {
        if ($Filter) {
            $stages += [pscustomobject]@{ Label = "tests matching '$Filter'"; Args = @("test", $Filter) }
        }
        elseif ($Pure) {
            # There is no library target here: every in-crate `#[cfg(test)]` module -- `ui`,
            # `shell`, `text`, `formats`, `readers`, `config`, `engines`, `app` -- compiles into
            # the one test binary of the *bin* target, which is why this selects
            # `--bin rust-hover-preview` and not the `--lib` there is no target for. So the split
            # is a runtime filter rather than a compile-time one: cargo passes `--skip` to the
            # test harness, which reads it as a substring of the test's own path, and the
            # window-making half of the suite is exactly the two roots below (1005 tests). What
            # is left is what never shows a window: 475 tests, 16 ignored.
            #
            # What it cannot reach: `ui::` and `shell::` also hold tests that would pass without
            # a desktop -- a layout sum, a key parsed out of a chord -- and they are skipped with
            # the rest, so a change to one of those is checked here only by `-Filter`. In the
            # other direction, a module outside those two that reaches for a window through
            # `ui::preview_window::tests::a_window_of_the_suite` would still be run; nothing
            # outside them does today.
            $stages += [pscustomobject]@{
                Label = "suite without a desktop"
                Args  = @("test", "--bin", "rust-hover-preview", "--", "--skip", "ui::", "--skip", "shell::")
            }
            $stages += [pscustomobject]@{ Label = "integration targets"; Args = $integration }
        }
        else {
            $stages += [pscustomobject]@{ Label = "tests"; Args = @("test") }
        }
    }
}

$total = $stages.Count
$number = 0
$failed = 0
$stoppedAt = ""
$clock = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($stage in $stages) {
    $number++
    $failed = Invoke-Stage -Number $number -Total $total -Label $stage.Label -CargoArgs $stage.Args
    if ($failed -ne 0) {
        $stoppedAt = $stage.Label
        break
    }
}

$clock.Stop()

if ($failed -ne 0) {
    $elapsed = Format-Elapsed $clock.Elapsed
    Write-Host ("verify: FAILED at [{0}/{1}] {2} after {3}" -f $number, $total, $stoppedAt, $elapsed)
    exit $failed
}

Write-Host ("verify: ok, {0} stages in {1}" -f $total, (Format-Elapsed $clock.Elapsed))
exit 0
