# Privacy Policy for Rust Hover Preview

**In short:** this app collects nothing and sends nothing about you. It has no
accounts, no telemetry, no crash reporting, and no ads, and the one thing it ever
asks the network is whether a newer release of itself has been published — which
it does when it starts and when you open its tray menu, and at most once an hour.
Everything else it reads stays on your PC, and everything it keeps is listed
below.

## What the app touches, and why

To show a preview, the app must read the file you hover or focus in Explorer.
That read happens locally, only for the hovered item, and only for preview
types you leave enabled:

- **Images:** decoded on your PC.
- **Videos:** measured and played on your PC. Where you have installed FFmpeg,
  that is the FFmpeg tools (`ffprobe`, `ffmpeg`, `ffplay`), and the file path is
  passed to those local programs as a command-line argument. Where you have not,
  it is the media engine built into Windows, which decodes the video inside this
  app's own process — nothing is started and no path leaves the app.
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

- No network connections except the update check below. Nothing else in the app
  opens a socket, and nothing is ever sent anywhere: the check asks this
  project's own GitHub releases for two files, and it carries nothing about you,
  your machine, or anything you have previewed beyond what any HTTPS request
  carries. No telemetry, no analytics, no crash reports.
- No keylogging. The app polls the pressed-or-not state of a few specific
  keys (your trigger key, the navigation keys, the letters and digits a name is
  typed with in Explorer, mouse buttons, and Ctrl+C while text is selected) to
  drive previews. Keystrokes are never recorded, stored, or sent.
- No screen scraping of other apps. The preview is the app's own window,
  painted by itself.
- No code injection into Explorer. It asks Explorer which item is under the
  cursor through public accessibility and Shell APIs.
- No admin rights, service, or driver. It runs as your user with a per-user
  install.

## Updates

The one question this app asks the network is whether a newer release than the
one you are running has been published, and it asks it when it starts and when
you open its tray menu. Nothing is checked while you work, and however often you
open the menu, a check is made at most once an hour — the hour is counted in
memory, so nothing about it is written to disk.

What it asks for is two files in this project's own GitHub releases: `version.txt`
at the newest release's stable address, and — where that names a version newer than
this one, and only once you have clicked the row it puts in the tray menu and
answered **Auto** to the dialog that asks — that release's installer. Both are
fetched over HTTPS, verified by the certificate checks Windows already performs,
and sent through the proxy Windows is configured with, if any. The request
identifies itself as this app and its version, and says nothing else.

Nothing is downloaded until you click the row above **Run at Startup** and answer
the dialog that follows: the check itself is a request for one small file. That
click asks first, and the dialog has three answers: **Auto** downloads the
installer and then runs it — it replaces the app silently, then starts it again —
**Manual** opens the release page in your browser, with nothing downloaded or run
here, and **Cancel** does nothing at all. A download that does not arrive, or that
is not a whole program, is deleted and reported; nothing is installed without
**Auto**.

A run makes one check as it starts, and at most one more an hour after that, so a
session you leave alone asks nothing beyond that first check. There is no
separate switch for the check yet.

## What is stored on your PC

| Location | What it holds | How to clear it |
|---|---|---|
| `%APPDATA%\rust-hover-preview\config.ini` | Your preferences only (toggles, delays, scales, theme name, extension lists). No file contents, no history. | Edit or delete it; a fresh one is recreated. |
| `%APPDATA%\rust-hover-preview\theme\` | `.tmTheme` files you drop in yourself. | Delete the files. |
| `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` value `RustHoverPreview` | Your exe path, only when **Run at Startup** is on. | Turn **Run at Startup** off. |
| `%TEMP%\rust-hover-preview\` | Transient renders — one scratch file an Office export is written to before it is read back, one copy of a long-path or downloaded document, the folder a conversion is staged in — and the installer of an update you clicked, while it installs. Each is deleted the moment it is read, and leftovers are deleted on the next launch. | Delete the folder; exit the app first. |
| `%TEMP%\rust-hover-preview\document\` | The page an engine has drawn for each document you hovered that an engine had to draw: the PDF Office exported, a slide's image, a picture of a workbook, or the PDF a conversion produced. Each is named for the document, the version of it and the engine — so it holds a page of that document's own **content**, and is worth treating the way you would the file. Kept so the document is not drawn again, given up oldest-read first past the **Cache → Document** size, and never removed at startup. | Set **Cache → Document** to `0 MB`, which keeps nothing between hovers; or delete the folder. Windows may clear the temp folder itself at any time. |
| `%TEMP%\rust-hover-preview\image\` | The picture an installed image converter (ImageMagick) developed for each raw or picture you hovered that it had to develop, at the preview size it was developed for. Each is named for the file, the version of it and the engine — so it holds a picture of that file's own **content**, and is worth treating the way you would the file. Kept so the file is not developed again, given up oldest-read first past the **Cache → Image (Disk)** size, and never removed at startup. | Set **Cache → Image (Disk)** to `0 MB`, which keeps nothing between hovers; or delete the folder. Windows may clear the temp folder itself at any time. |
| `%LOCALAPPDATA%\rust-hover-preview\` | Aside from the pages above, everything the engines need to run: the browser profile a run draws font and SVG previews in — which includes the small HTML page the preview is drawn in, and, for a face lifted out of a font collection, the one face written out for the browser to read — the records of the engine processes a run started, and the LibreOffice profile and one-paragraph stub document the app's own engine runs under. The profile of a run that is over is deleted by the next launch. | Delete the folder; exit the app first. |
| RAM only (never written to disk) | Decoded image frames (default 64 MB), archive listings, failure latches, the size a drawn page is placed by, the parsed text of a document and its styled lines. All keyed by path plus file version, evicted when full, gone on exit. | Set **Cache → Image (RAM)** to `0 MB` to keep nothing between hovers; quit to drop everything. |

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
- **External processes:** only `ffplay`/`ffprobe`/`ffmpeg` (your install, and
  only when it is installed), your Office apps (`WINWORD`/`EXCEL`/`POWERPNT`),
  and the WebView2 browser Windows ships with (`msedgewebview2.exe`) for SVG
  previews. Office is started hidden unless you already had that app open — your
  open instance is never hidden, quit, or killed, and neither is a browser
  another application owns.
  Everything this app starts is put in a Windows job object and written to a
  small file under `%LOCALAPPDATA%\rust-hover-preview\engines`, so that a crash,
  a kill from Task Manager, or a logoff cannot leave a process running: the job
  ends them along with the app, and the next launch ends whatever the job could
  not take. A recorded process is acted on only when its id still carries both
  the executable name and the start time it was recorded with, so a recycled id
  cannot hit something else.
- **Downloaded files:** files carrying a `Zone.Identifier` stream are opened
  in Office as a temporary copy without the mark, so Office's Protected View
  path works and your original file is never modified.

## Third parties on your machine

Your file content is handed to these local programs only, never over a
network: your Microsoft Office (Office previews), your FFmpeg binaries (video
previews, if you have them), the Windows media engine and the Windows PDF engine
(video and PDF previews), and the WebView2 runtime (SVG previews, which it is
given with all network access denied). Their
own vendor privacy statements apply to them; this app adds no reporting on top.

GitHub is the one service this app talks to at all, and only for the update
check: it is asked whether a newer release exists, and for the installer when you
ask to install one. The release page a **Manual** answer opens is handed to your
own browser rather than fetched here, and whatever it does with it is the
browser's business. No file you have previewed, and nothing about them, is part
of that request.

## If you report a bug

Issue reports, screenshots, and sample files you post on GitHub are public and
voluntary. File paths and document contents can appear in them, so remove
private data first. Security vulnerabilities: see `SECURITY.md` — do not file
them as public issues.

## Contact

Questions about this policy: open an issue at
https://github.com/cosmokud/rust-hover-preview/issues.
