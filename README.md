# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.98.1+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar. Hover a file in File Explorer — or move with the arrow keys — and a preview appears beside it. It works with mouse and keyboard, and you configure it from the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer.
- Previews appear beside the cursor or focused item and are kept on screen.
- Supports previews for images, design documents, camera raw, vector drawings, fonts, videos, PDFs, ebooks, comics, text/code, archives, Office documents, and more.
- A file’s preview depends on what’s really inside it—not its name—so renamed files still show correctly, unreadable content gets no preview, and ambiguous types are resolved by extension or content.
- Scaling from 25% to 400%, or fit-to-screen. Separate scaling for images and videos, and for vector drawings, PDFs, documents, fonts, and design documents.
- Tray menu and hand-editable `config.ini`.
- DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.

## Supported Formats

You can add or remove formats in `config.ini`. Unsupported formats show no preview, except text, which the app will try to force-read.

Almost everything is previewed by **Windows 11 and this app alone** — a codec Windows ships, WebView2, the drawing layer that plays metafiles, the Windows PDF engine, or a reader written into the app. Six things are not, and they are the app’s largest optional dependencies:

### Runs on Windows 11 alone

| Kind | Extensions |
| --- | --- |
| Pictures | `jpg` `jpeg` `jfif` `jpe` `png` `apng` `gif` `bmp` `ico` `tif` `tiff` `tga` `hdr` `exr` `ff` `qoi` `pnm` `pam` `pbm` `pgm` `ppm` `dds` `webp` |
| Pictures through a Windows codec extension | `heic` `heif` `avif` `avci` `jxl` — the extension packages are listed below; without one, no preview |
| Videos Windows 11 plays itself | `mp4` `m4v` `mov` `qt` `avi` `wmv` `asf` `dvr-ms` `mkv` `webm` `ts` `m2ts` `mts` `m2t` `mpg` `mpeg` `mpe` `m2v` `m1v` `vob` `3gp` `3g2` `3gpp` |
| Drawings | `svg` `svgz` (WebView2) · `wmf` `emf` (the drawing layer) · `eps` `epsi` `epsf` `epi` `ept` `ept2` `ept3` (the preview the file carries) |
| Design documents | `psd` `psb` `ai` `kra` `ora` `procreate` `sketch` `fig` `xd` |
| Fonts | `ttf` `otf` `ttc` `woff` `woff2` (WebView2) |
| Ebook | `pdf` `pdfa` `epdf` (the Windows PDF engine) · `cbz` `cbr` `cbc` |
| Text and code | `adb` `adoc` `ads` `asciidoc` `asm` `asp` `aspx` `astro` `awk` `bash` `bat` `bib` `bzl` `c` `cc` `cfg` `cg` `cjs` `clj` `cljc` `cljs` `cmake` `cmd` `comp` `conf` `cpp` `cs` `csh` `cshtml` `css` `csv` `csx` `cts` `cxx` `d` `dart` `diff` `diz` `edn` `ejs` `el` `elm` `env` `erb` `erl` `ex` `exs` `f` `f03` `f77` `f90` `f95` `fish` `for` `frag` `fs` `fsi` `fsx` `ftn` `fx` `geom` `glsl` `go` `gql` `gradle` `graphql` `groovy` `h` `haml` `hbs` `hcl` `hh` `hlsl` `hpp` `hrl` `hs` `htm` `html` `hxx` `inc` `ini` `ipynb` `java` `jl` `js` `json` `json5` `jsonc` `jsonl` `jsp` `jsx` `ksh` `kt` `kts` `latex` `less` `lhs` `liquid` `lisp` `ll` `lock` `log` `lsp` `lua` `m` `mak` `man` `markdown` `md` `mdown` `metal` `mjs` `mk` `mkd` `ml` `mli` `mm` `mts` `mustache` `nasm` `nfo` `nim` `ninja` `nix` `njk` `org` `pas` `patch` `php` `phtml` `pl` `plist` `pm` `properties` `proto` `ps1` `psd1` `psm1` `py` `pyi` `pyw` `r` `rake` `rb` `rkt` `rmd` `rs` `rst` `rtf` `s` `sass` `scala` `scm` `scss` `sh` `slim` `sol` `sql` `srt` `ss` `styl` `sv` `svelte` `svh` `swift` `tcl` `tex` `text` `tf` `tfvars` `toml` `ts` `tsv` `tsx` `twig` `txt` `v` `vbs` `vert` `vhd` `vhdl` `vtt` `vue` `wat` `wgsl` `xhtml` `xml` `xsd` `xsl` `xslt` `yaml` `yml` `zig` `zsh` — and these extensionless names: `authors` `.babelrc` `brewfile` `caddyfile` `changelog` `changes` `.clang-format` `.clang-tidy` `cmakelists.txt` `code_of_conduct` `containerfile` `contributing` `contributors` `copying` `copyright` `dockerfile` `.dockerignore` `.editorconfig` `.env` `.env.example` `.env.local` `.eslintignore` `.eslintrc` `gemfile` `.gitattributes` `.gitconfig` `.gitignore` `.gitkeep` `.gitmodules` `gnumakefile` `.golangci.yml` `history` `.htaccess` `install` `jenkinsfile` `justfile` `licence` `license` `.mailmap` `makefile` `makefile.am` `makefile.in` `notice` `.npmignore` `.prettierignore` `.prettierrc` `procfile` `rakefile` `readme` `.rustfmt.toml` `security` `.stylelintrc` `unlicense` `vagrantfile` |
| Archives | `7z` `apk` `jar` `rar` `tar` `tar.gz` `tgz` `xpi` `zip` `zipx` |

### Notes on the table

Animated GIF, APNG and WebP play. `hdr` and `exr` are tone-mapped for preview. `dds` previews BC1–BC7, both BC6H variants and uncompressed textures — including packed HDR and depth — at the first face and the nearest-size mip.

Comics are read by the app itself, so they need nothing installed: a `.cbz` is a zip of pages, a `.cbr` a rar of them, and a `.cbc` Calibre’s own container — a zip of several comics’ pages. What a hover shows is the **first page**, at the book kind’s scale and under the same **`Ebook`** switch a PDF answers to. Nothing is unpacked to disk: the archive’s headers are walked for the plate’s name and only that one plate is read, so a hundred-megabyte comic costs a page rather than an unpacking. The plate is chosen the way comic readers choose it — the first by name with the digits counted, so `2.jpg` comes before `10.jpg` — and a two-page spread is shown whole rather than cut in half.

### Needs FFmpeg, or the media engine Windows has

Videos the table does not already cover need FFmpeg, and these are all of them: `264` `265` `266` `apv` `av1` `avc` `avs` `avs2` `avs3` `bik` `bk2` `c93` `cavs` `cdg` `cdxl` `cin` `cpk` `dav` `dif` `divx` `drc` `dv` `evc` `f4v` `flm` `flv` `gxf` `h261` `h263` `h264` `h265` `h266` `h26l` `hevc` `ifv` `imx` `ismv` `ivf` `ivr` `kux` `m2p` `mj2` `mjpeg` `mjpg` `mk3d` `moflex` `mpv` `mve` `mvi` `mxf` `mxg` `nsv` `nut` `obu` `ogm` `ogv` `pmp` `psp` `rcv` `rm` `rmvb` `roq` `rsd` `smk` `str` `swf` `thp` `tod` `tp` `tr` `ty` `ty+` `usm` `vc1` `vc2` `viv` `vro` `vvc` `vw` `wtv` `xl` `xmv` `y4m` `yop`.

The containers and codecs the table above lists are the ones Windows plays on its own, and they play better with FFmpeg installed — more codecs inside the same container, and seeking rather than a still frame. The tray's **Codecs** menu shows what this machine can play.

### Needs Microsoft Office, or LibreOffice

Office documents — `doc` `docm` `docx` `dot` `dotm` `dotx` `xls` `xlsb` `xlsm` `xlsx` `xlt` `xltm` `xltx` `ppt` `pptm` `pptx` `pps` `ppsm` `ppsx` `pot` `potm` `potx` — are drawn in the background by an installed Office so previews appear quickly after the first hover. Excel needs a print queue to export a page — **Microsoft Print to PDF** is enough, and the Print Spooler service must be enabled; without one, Excel falls back to the sheet’s top-left corner. Where no Office is installed, an installed LibreOffice draws them instead. To have LibreOffice draw them all — with Microsoft Office installed as well — use **Engine → Select Engine → Office** and pick **LibreOffice**.

### Needs LibreOffice

The documents this app has no reader of its own for, drawn by an installed LibreOffice and sharp at any size, all of them: `123` `602` `abw` `cdr` `cgm` `cmx` `cwk` `dbf` `dif` `dxf` `fodg` `fodp` `fodt` `gnm` `gnumeric` `hwp` `key` `lwp` `mcw` `met` `mw` `numbers` `odb` `odc` `odf` `odg` `odm` `odp` `ods` `odt` `oth` `otg` `otm` `otp` `ots` `ott` `pages` `pcd` `pct` `pcx` `pdb` `pm6` `pmd` `psw` `pub` `ras` `sda` `sdc` `sdd` `sdw` `slk` `stc` `std` `sti` `stw` `svm` `sxd` `sxg` `sxi` `sxm` `sxw` `vdx` `vsd` `vsdm` `vsdx` `vstx` `wb2` `wk1` `wk3` `wk4` `wks` `wpg` `wq1` `wq2` `wpd` `wps` `wri` `xlw` `zabw` `zmf`.

### Needs ImageMagick

ImageMagick previews otherwise-unopenable files—especially camera raws—at display size, rotated, decoded from memory. Raws/own-coder: `3fr` `arw` `cr2` `cr3` `crw` `cur` `dcm` `dcr` `dcx` `dng` `dpx` `erf` `fff` `fit` `fits` `fts` `iiq` `j2c` `j2k` `jng` `jp2` `jpc` `jpm` `jpt` `k25` `kdc` `mdc` `mef` `miff` `mng` `mos` `mrw` `nef` `nrw` `orf` `pef` `pfm` `raf` `raw` `rmf` `rw2` `rwl` `sgi` `sr2` `srf` `srw` `vicar` `wbmp` `x3f` `xbm` `xcf` `xpm`. Alternates/other engine names: `pict` (`pct`), `sun` (`ras`), `pcds` (`pcd`), `dxt1`/`dxt5` (`dds`), `icb` `vda` `vst`, `picon`, `group4`;
Raw sample dumps: `rgb` `rgba` `gray` `cmyk` `ycbcr` `mono` `group4` and the rest; dimensions inferred by trying aspect ratios against file length—first exact divisor wins, else none; transpose ambiguity can make portrait landscape.

### Needs PeaZip

The archives this app has no reader of its own for, listed by an installed PeaZip and shown as the same page of contents a `.zip` is shown as, all of them: `001` `apfs` `ar` `arc` `arj` `bcm` `br` `bz2` `bzip2` `cab` `chm` `cpio` `cramfs` `deb` `dmg` `esd` `gz` `gzip` `hfs` `hfsx` `hxs` `iso` `lha` `lit` `lpaq8` `lzh` `lzma` `msi` `msp` `pkg` `ppkg` `qcow` `qcow2` `rpm` `squashfs` `swm` `taz` `tbz` `tbz2` `tpz` `txz` `tzst` `udf` `udeb` `vdi` `vhd` `vhdx` `vmdk` `wim` `xar` `xip` `xz` `z` `zpaq` `zst`. A `.tar.gz` and a `.tgz` are the archive list’s above and stay this app’s own.

### Needs Calibre

The ebooks this app has no reader of its own for, converted by an installed Calibre and drawn as a page of the PDF it wrote — the same kind, scale and backdrop a PDF this app reads itself gets, all of them: `azw` `azw3` `azw4` `djvu` `epub` `fb2` `htmlz` `lit` `lrf` `mobi` `pml` `prc` `snb` `tcr`.

A book with DRM — an `.azw`, `.azw3` or `.azw4` bought from Amazon — cannot be converted by anything here and shows nothing. See `TODO.md` for the formats the engine reads that are not in the list, and why.

### Themes

Text, code, and archive listings use Atom One Light by default, One Dark Pro, or any `.tmTheme` file placed in:

```text
%APPDATA%\rust-hover-preview\theme
```

## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

Steps:

1. Open [Releases](../../releases).
2. Download your preferred asset.
3. Run the installer, or place the portable binary wherever you like.
4. Launch Rust Hover Preview.

No Rust toolchain is needed. If upgrading from an earlier version, the installer cleans up the old `%LOCALAPPDATA%\Rust Hover Preview` folder automatically.

### Optional: Enable Video Preview with FFmpeg

Videos preview without FFmpeg through the media engine Windows ships with. FFmpeg adds far more formats and codecs. Install it if you want formats Windows does not decode. `ffplay` and `ffprobe` need to be in your `PATH`.

**Option A: winget**

```powershell
winget install --id Gyan.FFmpeg -e
```

**Option B: manual**

1. Download a Windows build from https://ffmpeg.org/download.html
2. Extract it, for example to `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to your user `PATH`

Verify either way:

```powershell
ffplay -version
ffprobe -version
```

### Optional: Enable More Video Codecs (Windows Codecs) — They are not needed while FFmpeg is installed.

The media engine decodes H.264, MPEG-4, and WMV out of the box. Each codec below is a separate free extension from the Microsoft Store:

| Codec                                  | Needs                                                                    |
| -------------------------------------- | ------------------------------------------------------------------------ |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)  |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)   |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)    |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4) |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)   |

All are free. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray’s **Codecs** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

`heic`, `heif`, `avif`, `jxl`, and still `webp` are decoded by a codec Windows provides rather than one shipped with the app. Each needs its extension installed once from the Microsoft Store:

| Format         | Needs                                                                                                                                            |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |

All are free. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error, and a multi-image file — a HEIC burst, an animated AVIF, an animated JPEG XL — shows its first frame. A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

### Optional: Enable CorelDRAW and Other Documents (LibreOffice)

`cdr` and everything else in the `[libre]` list is drawn by **LibreOffice**, which this app drives where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id TheDocumentFoundation.LibreOffice
```

Or download the official installer from https://www.libreoffice.org/download/ and run the suggested file (`LibreOffice_*_Win_x86-64.msi`; LibreOffice ships an `.msi`, not an `.exe`).

### Optional: Enable Camera Raw and More Pictures (ImageMagick)

`nef` and everything else in the `[magick]` list is developed by **ImageMagick**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id ImageMagick.ImageMagick
```

Or download the official installer from https://imagemagick.org/download/ and run the suggested file (`ImageMagick-*-Q16-HDRI-x64-dll.exe`).

### Optional: Enable Niche Archives (PeaZip)

`cab` and everything else in the `[peazip]` list is listed by **PeaZip**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id Giorgiotani.Peazip
```

Or download the official installer from https://peazip.github.io/peazip-64bit.html and run the suggested file (`peazip-*.WIN64.exe`).

### Optional: Enable Ebooks (Calibre)

`mobi` and everything else in the `[calibre]` list is converted by **Calibre**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id calibre.calibre
```

Or download the official installer from https://calibre-ebook.com/download_windows and run the suggested file (`calibre-*-64bit.msi`).

> [!WARNING]
> If `winget` reports an error, the package sources are usually why: run `winget source reset --force`.
>
> If that is refused as well, open **Terminal as administrator** (right-click the Start button → _Terminal (Admin)_) and run the same command from there.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys or Tab — to preview the focused item.
4. Right-click the tray icon to configure behavior.

## System Tray Menu

The item a setting starts at carries `(Default)` after its name, so a menu says both what is set now — the check or radio mark — and what a setting nobody has touched would be.

- **Enable Preview** — turn previews on or off.
- **Preview Types** — Images, Videos, Text, Ebook, Archives, Document, Vector, Fonts, Design: gate a kind without touching its file list. One gate covers each pair of a kind the app reads and the kind an engine draws for it, so a camera raw is gated by **Images**, a document LibreOffice draws by **Document**, an archive PeaZip lists by **Archives**, and a book Calibre converts by **Ebook** — which is also the switch over the pages the app draws itself: the PDF and the comic.
- **Text Preview**
  - **Full Mode** — adds scrolling, selection, and copy; off by default.
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in the theme folder.
  - **Font Size** — 400% at the top down to 70% at the bottom.
  - **Markdown** — Rendered or Source.
- **Timing**
  - **Trigger Key (Alt)** — the key is named in the item itself.
    - **Enable Trigger Key** — whether the key is watched at all.
    - **Hold to Disable Preview** / **Hold to Enable Preview** — what holding it does.
  - **Delay** — how long the pointer rests on a file before its preview opens: 0 ms at the top down to 1000 ms at the bottom. 0 ms by default.
  - **Rehover Delay** — the same steps, before the same file can preview again; 200 ms by default.
  - **Settling Delay** — the same steps, how long the pointer must have been still before it will preview what it is on: 0 ms, the default, is a new file previewed while the hand is still moving to it. The keyboard's own previews are not gated by it.
- **Placement**
  - **Position** — Follow Cursor or Best Position; Best Position by default.
  - **Avoid** — Avoid Nothing, Avoid Filename (default), Avoid Filename Column, or Avoid Details. This keeps a preview off the item it is about. Keyboard previews have no cursor to be anchored to, so **Avoid Nothing** places them the way **Avoid Filename** does.
- **Scaling**
  - **Image Scaling** — Fit to Screen or 25%–400%, of the image's own size.
  - **Video Scaling** — the same shares for a video, 100% (default).
  - **Animated Scaling** — the same shares for an animated GIF, WebP, or PNG; a still GIF or PNG keeps Image Scaling. 100% (default).
  - **Vector Scaling** — Fit to Screen (default), or 75%, 50%, 25%, 10% of the display.
  - **Ebook Scaling** — Fit to Screen (default), or the same shares of the display for a PDF page, a comic’s first page, and a book an installed Calibre converted into one.
  - **Document Scaling** — the same shares of the display for a document drawn as a page, whichever drew it: an Office document’s own application, or LibreOffice for the formats beside it (CorelDRAW and the older word processors, spreadsheets and presentations). A workbook’s fallback bitmap is never enlarged.
  - **Font Scaling** — the same shares for a font specimen, 50% by default.
  - **Design Scaling** — the same shares of the display for a design document (Photoshop, Illustrator, Krita, OpenRaster, Procreate); Fit to Screen by default.
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard; Checkerboard by default.
  - **Vector Background** — the same backdrops for a drawing: an SVG document, or a metafile; Checkerboard by default.
  - **Font Background** — the same backdrops for a font specimen; White by default.
  - **DDS Background** — Black or White for `.dds` textures, whose alpha channel is as often a mask or an unused channel as it is transparency; White by default. The two backdrops that show what stands behind a texture are not offered for one.
  - **Design Background** — the same backdrops as a picture's, for a design document; Checkerboard by default.
- **Volume** — Max, High, Medium, Low, Very Low, Mute: 100% down to 0%.
- **Performance**
  - **Cache** — what a preview may cost between hovers, 2 GB down to 0 MB: **`Image (RAM)`**, the decoded frames held in memory, and **`Document (Disk)`**, the pages an engine drew, kept as files under the temp folder so a document drawn once is not drawn again.
  - **Decode Budget** — 16 GB down to 512 MB, 1 GB default: a file past it gets no preview.
  - **Tick** — how often the app looks at Explorer while a folder window is in focus: 15 ms (default), 31, 47, 63 or 78 ms, one to five Windows timer ticks. Lower answers a move sooner; higher is lighter on the CPU and on Explorer.
- **Engine**
  - **AFK Timer** — how long Explorer may be out of reach before an engine that is not marked **Persistent** is let go: 1 hour, 30 minutes, 10 minutes, 5 minutes, 1 minute (default), 30 seconds, 15 seconds. It counts time with no Explorer window reachable — every one of them minimized, or every one of them behind a maximized or fullscreen app — on any monitor, so a second Explorer window on another screen keeps every engine warm. Each **`… TTL`** submenu has a **Persistent** toggle at the top for an engine you want kept regardless.
  - **Select Engine → Office** — which engine an Office document’s page is asked of: **Microsoft Office** (default) draws it with the application that owns the format and keeps LibreOffice as the fallback for a family this machine has no application for, while **LibreOffice** draws every Office document whether Microsoft Office is installed or not. The LibreOffice row is greyed out where it is not installed.
  - **Microsoft Office TTL** — **Persistent** on, how long a family’s Office app is kept warm whatever you are doing: Indefinitely, 1 hour, 30 minutes, 10 minutes (default), 5 minutes, 1 minute, 0 seconds. Off, it is kept while Explorer is reachable and let go by **AFK Timer** instead.
  - **LibreOffice TTL** — the same for the engine that draws CorelDRAW and the documents beside it: how long it is kept after the last document it converted. A kept engine converts the next document without starting up again — measured on one document, 1.2 s cold against 0.2 s — and while it is kept it is a LibreOffice with a small document of this app's own open, which is a few hundred megabytes. `0 seconds` keeps none, which is an engine per document. Both are the answer for an engine marked **Persistent**; one that is not is let go by **AFK Timer** instead. Greyed out where LibreOffice is not installed.
  - **WebView2 TTL** — **Persistent** on, how long the browser that draws SVG documents is kept warm; off, it is let go by **AFK Timer**. Greyed out where WebView2 is missing.
  - There is no `ImageMagick TTL`, no `PeaZip TTL` and no `Calibre TTL`, because none of those three is a process the app can keep: each is handed a file, writes what it read and exits. What a second hover of the same file costs is a hit in the cache its answer was kept in, so there is nothing an idle time could bound.
- **Codecs** — what this machine has: Videos, Images, and Engines. Missing ones are greyed out.
- **Run at Startup** — add or remove the Windows startup entry.
- **Config.ini** — open the configuration file; the item is named for the running version.
- **Exit** — close the app.

## Configuration

Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

The file is watched, and changes apply without a restart.

Example, trimmed:

```ini
[settings]
; General
preview_enabled=true
run_at_startup=true

; Preview Types
archive_preview_enabled=true
design_preview_enabled=true
document_preview_enabled=true
ebook_preview_enabled=true
font_preview_enabled=true
image_preview_enabled=true
text_preview_enabled=true
vector_preview_enabled=true
video_preview_enabled=true

; Text Preview
markdown_mode=rendered
text_font_scale=125
text_preview_full_mode=false
theme=light

; Timing
hover_delay_ms=0
same_file_rehover_delay_ms=200
settling_delay_ms=0
trigger_key=alt
trigger_key_enabled=true
trigger_key_mode=disable

; Placement
avoid_mode=filename
follow_cursor=false

; Scaling
animated_scale=100
design_scale=fit
document_scale=fit
ebook_scale=fit
font_scale=50
preview_scale=100
vector_scale=fit
video_scale=100

; Background
dds_background=white
design_background=checkerboard
font_background=white
image_background=checkerboard
vector_background=checkerboard

; Volume
video_volume=0

; Performance
decode_budget_gb=1
document_cache_mb=128
image_cache_mb=32
tick_ms=15

; Engine
afk_timer_seconds=60
libreoffice_idle=600
libreoffice_persistent=false
office_engine=microsoft_office
office_engine_idle=600
office_engine_persistent=false
webview_idle=600
webview_persistent=false

; Advanced
hdr_exposure=0
hdr_tone_map=reinhard
spinner_delay_ms=250
```

Key settings, in plain terms:

- `theme` — `light`, `dark`, or `custom:<name>`.
- `markdown_mode` — `rendered` or `source`.
- `*_preview_enabled` — whether that kind of preview may show at all; the file lists stay as they are.
- `text_preview_full_mode` — `true` adds scrolling, selection, and copy.
- `text_font_scale` — a percentage from 1 to 1000, default `125`; archive listings follow it too.
- `extensions` / `names` — the text-preview gates. Extensions are written without dots. Names match extensionless files.
- `image_extensions`, `video_extensions`, `archive_extensions`, `office_extensions`, `font_extensions`, `design_extensions`, `vector_extensions` — per-type preview gates, written without dots. An entry with a dot in it, like `tar.gz`, is matched against the end of the file name. `ebook_extensions`, `libre_extensions`, `magick_extensions`, `peazip_extensions` and `calibre_extensions`, in their own sections, are the pages this app draws itself (a PDF, and a comic’s first plate), the documents LibreOffice draws, the pictures ImageMagick develops, the archives PeaZip lists, and the books Calibre converts.
- `image_cache_mb` — memory for decoded image frames: default `32`, max `2048`; `0` holds nothing.
- `document_cache_mb` — disk for the pages an engine drew — an Office export, a slide’s image, a workbook’s picture, a converted PDF: default `128`, max `2048`; `0` keeps nothing between hovers but still draws for the current one. The pages live under `%TEMP%\rust-hover-preview\document`, named for the document, the version of it and the engine that drew it, and the page that has not been read for longest is the one given up first.
- `decode_budget_gb` — the most memory one hover may decode or read for: default `1`, smallest `0.25`, largest `64`. A file past it shows no preview.
- `hdr_tone_map` — how HDR/EXR light values become screen values: `reinhard` (default), `aces` (filmic), `srgb` (clips), or `off` (bare clamp). Pictures already in screen values, like PNG or JPEG, are never affected.
- `hdr_exposure` — how many stops those pictures are shifted before that curve: default `0`, clamped to `-10`–`10`.
- `spinner_delay_ms` — how long a hover’s load may run before the waiting spinner is put up, in milliseconds: default `250`, and `0` puts it up with the load. One delay answers every kind of preview — a decode, a page Office is rendering, a browser that has to start — and there is no tray entry for it.
- `office_engine` — which engine draws an Office document’s page: `microsoft_office` (default) asks the application that owns the format and falls back to LibreOffice for a family this machine has no application for, while `libreoffice` asks the render engine for every Office document whether Microsoft Office is installed or not. With no LibreOffice installed the second falls back to the first, and the tray row is greyed out.
- `libreoffice_idle` — seconds the LibreOffice engine is kept after the last page it drew, or `indefinitely`; default `600`. `0` keeps no engine at all, which is a launch per document; while one is kept it is a LibreOffice with a small document of this app's own open, and it is ended by the app when the time is up.
- `office_engine_idle` — seconds an Office engine is kept after its last page, or `indefinitely`; default `600`. `0` lets it go as soon as it has drawn a page. It bounds an engine the same way `libreoffice_idle` does, and like it, only while that engine is marked persistent.
- `afk_timer_seconds` — seconds with no Explorer window reachable before an engine that is not marked persistent is let go; default `60`, clamped to a day. `0` lets an engine go the moment Explorer goes out of reach.
- `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent` — `true` keeps that engine whatever you are doing, with its `…_idle` time as the only bound; `false` (default) keeps it while Explorer is reachable and lets `afk_timer_seconds` bound it.
- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled` — the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched at all; `true` by default.
- `follow_cursor` — `true` for Follow Cursor, `false` for Best Position.
- `avoid_mode` — `filename` (default), `filename_column`, `details`, or `off`: what a preview is kept off.
- `preview_scale` — percentage or `fit`, read against the picture's own size.
- `video_scale` — the same for a video, `100` by default.
- `animated_scale` — the same for an animated GIF, WebP, or PNG, `100` by default; a still GIF or PNG follows `preview_scale`.
- `vector_scale` — percentage or `fit`, read against the screen. `fit` is default, is all of the room, and `100` or more reads as it. It covers every drawing the Vector kind holds. The name it used to be written under (`svg_scale`) is not read: a line like that is removed the next time the file is written, and the setting goes back to its default.
- `magick_preview_enabled` and `peazip_preview_enabled` are not read: a picture the ImageMagick engine develops is gated by `image_preview_enabled` and an archive the PeaZip engine lists by `archive_preview_enabled`, since what either is previewed as is a picture or an archive. `libre_preview_enabled` and `libre_scale` are gone the same way — what draws those documents is `document_preview_enabled` and `document_scale`. Lines left under the old names are removed the next time the file is written. `office_cache_mb` and `libre_cache_mb` are gone the same way: one `document_cache_mb` replaces both, and a file that still holds either is read once for the larger of the two and written without them.
- There is no TTL for ImageMagick, and no `magick_idle` key: it is the one engine here that is a converter rather than a process this app can hold open, so what a file costs is a conversion or a hit in the picture cache (`image_cache_mb`) and never a file of the app's own.
- There is no TTL for PeaZip either, and no `peazip_idle` key, for the same reason: what this app runs of it are the tools PeaZip carries, each of which prints an archive’s table of contents and exits — and for the three single-stream names whose tools have no listing to print, `br`, `bcm` and `lpaq8`, nothing is run at all. What a file costs is a listing or a hit in the listing cache, and a second hover of the same archive starts nothing at all.
- And there is no TTL for Calibre, and no `calibre_idle` key, for the engine’s own answer a third time: `ebook-convert` reads a book, writes a PDF of it and exits, booting a whole Python application to do it. There is no instance to hold open between books and nothing an idle time could bound, however much a launch costs; what a second hover of the same book costs is a read of the page it converted, kept in the page cache (`document_cache_mb`) beside every other engine’s page.
- `ebook_scale` — percentage or `fit`, read against the screen, for a PDF page; `fit` by default. `100` or more reads as `fit`.
- `document_scale` — the same for a document drawn as a page, `fit` by default, whichever drew it: an Office document’s own application, or LibreOffice for the formats beside it. A workbook’s fallback bitmap follows its own size and is never enlarged.
- `font_scale` — percentage or `fit`, read against the screen. `50` is default, `fit` is all of it, and `100` or more reads as `fit`. A font has no size of its own, so the share is of the display.
- `design_scale` — the same for a design document, `fit` by default. A design preview is the picture the file keeps of the whole document, so the share is of the screen the way a page’s is rather than of the document’s own size.
- `ttc_face` — which face of a `.ttc` collection is drawn: `1` is the first face, and the highest setting is `10`. The heading says which face came out.
- `image_background` — `checkerboard` (default), `black`, `white`, or `transparent`: what a picture is drawn over, and with it a PDF page, a painted text frame and a page a document was drawn as. The one name every kind used to share (`transparent_background`) is not read: lines like that are removed the next time the file is written.
- `font_background` — `white` (default), `black`, `checkerboard`, or `transparent`.
- `dds_background` — `white` (default) or `black`, and only those two: a texture's alpha channel is as often a mask, a height or a roughness as it is transparency, so the backdrops that show what stands behind a preview are not offered for one, and a file that names one of them is read as `white`.
- `vector_background` — `checkerboard` (default), `white`, `black`, or `transparent`: what an SVG document's page, or a metafile drawing, is drawn over. The name it used to be written under (`svg_background`) is not read: a line like that is removed the next time the file is written.
- `design_background` — `checkerboard` (default), `white`, `black`, or `transparent`.
- A deleted `extensions=` line or whole section comes back with built-in entries. An `extensions=` line left empty stays empty.

## Build from Source

Requirements: Windows 11, Rust 1.98.1+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

```bash
cargo build            # debug
cargo build --release  # release
```

The release binary is written to `target/release/rust-hover-preview.exe`. A release build ends a running copy first, since Windows will not let the linker replace an open binary. Debug builds are left alone.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item. GDI paints the preview into a topmost layered window. Text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives’ own tables of contents — or from the listing an installed PeaZip produces for the formats no reader here has — Office documents are drawn from a page Office renders in the background, WebView2 draws SVG documents and font specimens, camera raw and the pictures beside it are developed by ImageMagick, and video is played by FFmpeg where installed and by Windows’ own media engine where it is not.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview is local-first, previews work without an internet connection, and the only network request is an update check, which runs only when you open the tray menu and at most once an hour. There is no telemetry, analytics, ads, accounts, or crash reporting. It reads only the item you hover or focus in Explorer, locally and only for enabled preview types. Cloud-only placeholders are skipped on purpose; password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use locally installed FFmpeg, LibreOffice, or ImageMagick when available, plus Microsoft Office, Windows' own media engine, and the Windows PDF engine. Caches are bounded by `config.ini`: decoded images stay in memory, while the page an engine drew for a document — an Office export, a converted PDF — is kept as a file under the temp folder, where Windows is free to clear it. See `PRIVACY.md` for full details.

## License

MIT. See [LICENSE](LICENSE).

The app is built out of other people’s code as much as its own — the Windows bindings, the image decoders, the syntax highlighter, the archive readers, the browser bindings — and each of those carries its own licence, with the notices MIT and BSD ask to be reproduced. [THIRD-PARTY.md](THIRD-PARTY.md) lists every dependency grouped by licence, with the copyright holders beside it, and the full texts are in [`LICENSES/`](LICENSES). Both are generated from the dependency tree rather than kept by hand; `generate-attribution.ps1` refreshes them.
