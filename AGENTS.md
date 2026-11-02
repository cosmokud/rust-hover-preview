# AGENTS.md

## What this is

Windows 11 tray app in Rust: hovering a file in Explorer previews it in-process or through an external
engine; middle-click or Space pins that preview to a window of its own.

- **Stack:** Rust 2021 (`windows` crate, raw Win32 — no GUI framework, no async runtime), GDI/D3D11 into
  layered windows, WebView2 for SVG/font/HTML pages, and external tools (mpv, ffplay, LibreOffice,
  ImageMagick, PeaZip, Calibre) as child processes. `README.md` is the user's side, `ARCHITECTURE.md`
  the engineering one.

## Ground truth

[ARCHITECTURE.md](ARCHITECTURE.md) says why, and in what order, one section per subsystem; read only the section your change touches.
`codegraph` says where a symbol is and who calls it.
Neither substitutes for reading the code before changing behaviour.