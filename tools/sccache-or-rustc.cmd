@echo off
rem The wrapper `.cargo/config.toml` hands every rustc call to, named for the choice it
rem makes: the compiler goes through sccache, or straight to rustc.
rem
rem A machine that installed sccache has it on PATH — winget, cargo install, a manual copy
rem — and every compile is served from the cache, which is what makes this worth checking
rem in: one line of config and the whole tree builds cached, no variable to remember.
rem
rem A machine that never installed it still builds. `where` fails there, the errorlevel is
rem not zero, and the compiler runs unwrapped exactly as it would have with no wrapper
rem named at all.
where sccache >nul 2>nul
if %errorlevel% == 0 (
    sccache %*
) else (
    %*
)
