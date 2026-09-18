# Privacy Policy for Rust Hover Preview

**In short:** this app works fully offline. It collects nothing, sends nothing
anywhere, has no accounts, no telemetry, no crash reporting, no ads, and no
update checks. Everything it reads stays on your PC, and everything it keeps is
listed below.

## What the app touches, and why

To show a preview, the app must read the file you hover or focus in Explorer.
That read happens locally, only for the hovered item, and only for preview
types you leave enabled:

- **Images:** decoded on your PC.
- **Videos:** measured and played with the FFmpeg tools you installed
  (`ffprobe`, `ffmpeg`, `ffplay`). The file path is passed to those local
  programs as a command-line argument.
- **PDFs:** first page rendered with the PDF engine built into Windows.
- **Text and code:** the start of the file is read (capped at 2 MB / 2,000
  lines), decoded, and highlighted on your PC.
- **Archives:** only the archive's own table of contents is read (capped at
  20,000 entries, 64 MB of `.tar.gz` stream, 300 characters per name).
  Nothing is unpacked. Password-locked tables are shown as locked, never
  cracked.
- **Office documents:** the first page is rendered by your installed Microsoft
  Office, driven locally through COM automation. Documents open read-only,
  never added to recent files, with macros force-disabled and pop-up dialogs
  suppressed (Office's own settings are restored afterwards). A document that
  refuses — password, repair dialog, Protected View — is left alone and not
  retried for 2 minutes.

Files stored only online (OneDrive, SharePoint, Dropbox placeholders) are
**skipped on purpose**: the app checks file attributes, not content, so a
hover never starts a download or costs metered data.

Password-protected PDFs are skipped. Encrypted archives are listed as
encrypted. Nothing bypasses a password, and no password is ever stored.

## What the app does not do

- No network connections of any kind. The only URLs in the project are the
  crate registry (build time) and the release workflow (GitHub servers) — the
  app binary itself never contacts the internet.
- No telemetry, analytics, crash reports, or update checks.
- No keylogging. The app polls the pressed-or-not state of a few specific
  keys (your trigger key, arrow keys, mouse buttons, Ctrl+C while text is
  selected) to drive previews. Keystrokes are never recorded, stored, or sent.
- No screen scraping of other apps. The preview is the app's own window,
  painted by itself.
- No code injection into Explorer. It asks Explorer which item is under the
  cursor through public accessibility and Shell APIs.
- No admin rights, service, or driver. It runs as your user with a per-user
  install.

## What is stored on your PC

| Location | What it holds | How to clear it |
|---|---|---|
| `%APPDATA%\rust-hover-preview\config.ini` | Your preferences only (toggles, delays, scales, theme name, extension lists). No file contents, no history. | Edit or delete it; a fresh one is recreated. |
| `%APPDATA%\rust-hover-preview\theme\` | `.tmTheme` files you drop in yourself. | Delete the files. |
| `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` value `RustHoverPreview` | Your exe path, only when **Run at Startup** is on. | Turn **Run at Startup** off. |
| `%TEMP%\rust-hover-preview\` | Transient Office renders: one scratch copy of a long-path or downloaded document, one page file being read back. Each is deleted the moment it is read; leftovers are deleted on next launch. | Delete the folder; exit the app first. |
| `%TEMP%\rust-hover-preview-video.log` | Append-only debug log, one line per video preview: file path, preview position/size, measured dimensions, crop, and filter. **This is the only place hovered file paths are written to disk.** It never leaves your PC. | Delete the file. |
| RAM only (never written to disk) | Decoded image frames (default 32 MB), rendered Office pages (default 64 MB), rendered PDF pages (default 32 MB), painted text frames (default off), archive listings, failure latches. All keyed by path plus file version, evicted when full, gone on exit. | Set a cache to `0 MB` to keep nothing between hovers; quit to drop everything. |

Error messages (hook install failure, mutex failure) go to stderr only and are
not saved to any file. Test-only diagnostics print to the console and do not
run during normal use.

## OS access, in plain terms

- **Explorer state:** cursor position, window under the cursor, and the focused
  item, read through UI Automation and Shell COM to resolve the hovered file.
  Folder paths are held in memory as short-lived lookup caches only.
- **Mouse wheel:** a system-wide low-level wheel hook counts wheel ticks so a
  scroll under a parked cursor refreshes the preview. It records counts, not
  positions or applications. Wheel motion over a scrollable text preview is
  given to the preview instead of Explorer.
- **Clipboard, two cases:** (1) copying from a text preview writes to the
  clipboard only when you choose Copy / Select All / Ctrl+C with a selection.
  (2) On printerless machines, an Excel sheet preview copies a corner of the
  sheet via `CopyPicture`, reads that bitmap, then clears the clipboard so
  Excel can quit cleanly. That means hovering a spreadsheet can replace
  whatever you had copied — expected side effect, local only.
- **External processes:** only `ffplay`/`ffprobe`/`ffmpeg` (your install) and
  your Office apps (`WINWORD`/`EXCEL`/`POWERPNT`), the latter started hidden
  unless you already had that app open — your open instance is never hidden,
  quit, or killed. A process the app started that outlives its use is ended by
  verified PID plus executable-name check.
- **Downloaded files:** files carrying a `Zone.Identifier` stream are opened
  in Office as a temporary copy without the mark, so Office's Protected View
  path works and your original file is never modified.

## Third parties on your machine

Your file content is handed to these local programs only, never over a
network: your Microsoft Office (Office previews), your FFmpeg binaries (video
previews), and the Windows PDF engine (PDF previews). Their own vendor privacy
statements apply to them; this app adds no reporting on top.

## If you report a bug

Issue reports, screenshots, and sample files you post on GitHub are public and
voluntary. File paths and document contents can appear in them, so remove
private data first. Security vulnerabilities: see `SECURITY.md` — do not file
them as public issues.

## Contact

Questions about this policy: open an issue at
https://github.com/cosmokud/rust-hover-preview/issues.
