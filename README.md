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
| Pictures | jpg jpeg jfif jpe png apng gif bmp ico tif tiff tga hdr exr ff qoi pnm pam pbm pgm ppm dds webp |
| Pictures through a Windows codec extension | heic heif avif avci jxl |
| Videos Windows 11 plays itself | mp4 m4v mov qt avi wmv asf dvr-ms mkv webm ts m2ts mts m2t mpg mpeg mpe m2v m1v vob 3gp 3g2 3gpp |
| Sounds Windows 11 plays itself | mp3 wav wave m4a m4b aac wma aif aifc aiff amr awb flac dsf |
| Drawings | svg svgz (WebView2) · wmf emf (the drawing layer) · eps epsi epsf epi ept ept2 ept3 (the preview the file carries) |
| Design documents | psd psb ai kra ora procreate sketch fig xd |
| Fonts | ttf otf ttc woff woff2 (WebView2) |
| Ebook | pdf pdfa epdf (the Windows PDF engine) · cbz cbr cbc |
| Text and code | adb adoc ads asciidoc asm asp aspx astro awk bash bat bib bzl c cc cfg cg cjs clj cljc cljs cmake cmd comp conf cpp cs csh cshtml css csv csx cts cxx d dart diff diz edn ejs el elm env erb erl ex exs f f03 f77 f90 f95 fish for frag fs fsi fsx ftn fx geom glsl go gql gradle graphql groovy h haml hbs hcl hh hlsl hpp hrl hs htm html hxx inc ini ipynb java jl js json json5 jsonc jsonl jsp jsx ksh kt kts latex less lhs liquid lisp ll lock log lsp lua m mak man markdown md mdown metal mjs mk mkd ml mli mm mts mustache nasm nfo nim ninja nix njk org pas patch php phtml pl plist pm properties proto ps1 psd1 psm1 py pyi pyw r rake rb rkt rmd rs rst rtf s sass scala scm scss sh slim sol sql srt ss styl sv svelte svh swift tcl tex text tf tfvars toml ts tsv tsx twig txt v vbs vert vhd vhdl vtt vue wat wgsl xhtml xml xsd xsl xslt yaml yml zig zshExtensionless / Specific Files:authors .babelrc brewfile caddyfile changelog changes .clang-format .clang-tidy cmakelists.txt code_of_conduct containerfile contributing contributors copying copyright dockerfile .dockerignore .editorconfig .env .env.example .env.local .eslintignore .eslintrc gemfile .gitattributes .gitconfig .gitignore .gitkeep .gitmodules gnumakefile .golangci.yml history .htaccess install jenkinsfile justfile licence license .mailmap makefile makefile.am makefile.in notice .npmignore .prettierignore .prettierrc procfile rakefile readme .rustfmt.toml security .stylelintrc unlicense vagrantfile |
| Archives | 7z apk jar rar tar tar.gz tgz xpi zip zipx |


### Notes on the table

Animated GIF, APNG and WebP play. `hdr` and `exr` are tone-mapped for preview. `dds` previews BC1–BC7, both BC6H variants and uncompressed textures — including packed HDR and depth — at the first face and the nearest-size mip.

### Needs FFmpeg

Videos the table does not already cover need FFmpeg, and these are all of them: `264` `265` `266` `apv` `av1` `avc` `avs` `avs2` `avs3` `bik` `bk2` `c93` `cavs` `cdg` `cdxl` `cin` `cpk` `dav` `dif` `divx` `drc` `dv` `evc` `f4v` `flm` `flv` `gxf` `h261` `h263` `h264` `h265` `h266` `h26l` `hevc` `ifv` `imx` `ismv` `ivf` `ivr` `kux` `m2p` `mj2` `mjpeg` `mjpg` `mk3d` `moflex` `mpv` `mve` `mvi` `mxf` `mxg` `nsv` `nut` `obu` `ogm` `ogv` `pmp` `psp` `rcv` `rm` `rmvb` `roq` `rsd` `smk` `str` `swf` `thp` `tod` `tp` `tr` `ty` `ty+` `usm` `vc1` `vc2` `viv` `vro` `vvc` `vw` `wtv` `xl` `xmv` `y4m` `yop`.

The containers and codecs the table above lists are the ones Windows plays on its own, and they play better with FFmpeg installed — more codecs inside the same container, and seeking rather than a still frame. The tray's **Codecs** menu shows what this machine can play.

Sounds Windows does not decode need FFmpeg as well, and these are all of them: `ac3` `ape` `au` `caf` `dff` `dts` `dtshd` `eac3` `mka` `mp2` `mpa` `mpc` `oga` `ogg` `ofr` `ofs` `opus` `ra` `shn` `snd` `spx` `tak` `tta` `voc` `wv`

**Normalize** needs it too: what measures a sound's peak and what applies it are FFmpeg's, so on a machine without it the row is greyed and every sound plays as the file holds it.

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

1. Download a Windows build from [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html)
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

Or download the official installer from [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/) and run the suggested file (`LibreOffice_*_Win_x86-64.msi`; LibreOffice ships an `.msi`, not an `.exe`).

### Optional: Enable Camera Raw and More Pictures (ImageMagick)

`nef` and everything else in the `[magick]` list is developed by **ImageMagick**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id ImageMagick.ImageMagick
```

Or download the official installer from [https://imagemagick.org/download/](https://imagemagick.org/download/) and run the suggested file (`ImageMagick-*-Q16-HDRI-x64-dll.exe`).

### Optional: Enable Niche Archives (PeaZip)

`cab` and everything else in the `[peazip]` list is listed by **PeaZip**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id Giorgiotani.Peazip
```

Or download the official installer from [https://peazip.github.io/peazip-64bit.html](https://peazip.github.io/peazip-64bit.html) and run the suggested file (`peazip-*.WIN64.exe`).

### Optional: Enable Ebooks (Calibre)

`mobi` and everything else in the `[calibre]` list is converted by **Calibre**, which this app runs where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id calibre.calibre
```

Or download the official installer from [https://calibre-ebook.com/download\_windows](https://calibre-ebook.com/download_windows) and run the suggested file (`calibre-*-64bit.msi`).

> \[!WARNING\]
> If `winget` reports an error, the package sources are usually why: run `winget source reset --force`.
>
> If that is refused as well, open **Terminal as administrator** (right-click the Start button → *Terminal (Admin)*) and run the same command from there.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys, Tab, or a file's name typed to jump to it — to preview the focused item.
4. Right-click the tray icon to configure behavior.

Here is a concise, plain-language version:

## System Tray Menu

A setting marked `(Default)` is what an untouched setting would be. The check or radio mark shows what is set now.

- **Enable Preview** — Turn previews on or off.
- **Preview Types** — Choose which file kinds can preview: Images, Videos, Audio, Text, Ebook, Archives, Document, Vector, Fonts, Design. One switch covers both the original file and the engine-drawn preview. Examples: camera raw uses **Images**; LibreOffice document uses **Document**; PeaZip archive uses **Archives**; Calibre book uses **Ebook** — also the app-drawn PDF and comic pages.
- **Text Preview**
  - **Full Mode** — Adds scrolling, selection, and copy. Off by default.
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in the theme folder.
  - **Font Size** — `400%` at the top down to `70%` at the bottom.
  - **Markdown** — Rendered or Source.
- **Timing**
  - **Prioritize Keyboard** — On by default. The file under a pointer that has not been moved does not preview of its own while the keyboard is driving Explorer, so pressing a key onto a file with no preview of its own behaves like pressing one onto a file that has a preview. The pointer takes the screen back when it is moved, when the wheel is turned, or when a folder change hands it over. Off, the pointer's own hover always wins.
  - **Trigger Key (Alt)** — The key is named in the item.
    - **Enable Trigger Key** — Whether the key is watched.
    - **Hold to Disable Preview** / **Hold to Enable Preview** — What holding the key does.
  - **Delay** — How long the pointer rests before a preview opens: `0 ms` at the top down to `1000 ms` at the bottom. Default `0 ms`.
  - **Rehover Delay** — Wait before the same file can preview again. Default `200 ms`.
  - **Settling Delay** — How still the pointer must be before previewing what it is on. Default `0 ms` means a new file can preview while the hand is still moving. Keyboard previews ignore this.
- **Placement**
  - **Position** — Follow Cursor or Best Position. Default **Best Position**.
  - **Avoid** — Avoid Nothing, Avoid Filename (`default`), Avoid Filename Column, or Avoid Details. Keeps a preview off the item it is about. Keyboard previews have no cursor, so **Avoid Nothing** acts like **Avoid Filename**.
- **Scaling**
  - **Image Scaling** — Fit to Screen or `25%`–`400%` of the image’s own size.
  - **Video Scaling** — Same shares for a video. Default `100%`.
  - **Animated Scaling** — Same shares for an animated GIF, WebP, or PNG. A still GIF or PNG uses **Image Scaling**. Default `100%`.
  - **Vector Scaling** — Fit to Screen (`default`), or `75%`, `50%`, `25%`, `10%` of the display.
  - **Ebook Scaling** — Fit to Screen (`default`), or the same display shares for a PDF page, a comic’s first page, and a Calibre-converted book.
  - **Document Scaling** — Same display shares for a document drawn as a page, whether by its own Office app or by LibreOffice. A workbook’s fallback bitmap is never enlarged.
  - **Font Scaling** — Same shares for a font specimen. Default `50%`.
  - **Design Scaling** — Same display shares for design documents: Photoshop, Illustrator, Krita, OpenRaster, Procreate. Default **Fit to Screen**.
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard. Default **Checkerboard**.
  - **Vector Background** — Same backdrops for an SVG document or metafile. Default **Checkerboard**.
  - **Font Background** — Same backdrops for a font specimen. Default **White**.
  - **DDS Background** — Black or White for `.dds` textures. Default **White**. The two see-through backdrops are not offered.
  - **Design Background** — Same as picture backdrops for a design document. Default **Checkerboard**.
- **Volume** — **Video** and **Audio**, each offering `100%`, `80%`, `65%`, `50%`, `35%`, `20%`, `10%`, `5%`, `1%` and `0%`, loudest first. A video's soundtrack starts at `0%` — silent, so a hover never makes a sound the pointer did not ask for — and a sound file at `10%`: a video is looked at and a song is listened to, so the two are settings of their own. A sound at `0%` still shows its card, silently.
  - **Normalize** — The first row of each half, above the levels: a file's loudest sample is measured and brought to full scale before it plays, so a folder of sounds — or a set of films — is heard at one level rather than at each file's own. On by default for **Audio** and off for **Video**, whose soundtrack is heard beside a picture that was asked for and whose measurement is a decode of the film. Both are greyed out unless FFmpeg is installed — FFmpeg is what measures the peak and what applies it. The peak is measured once per file and kept, so only a file's first hover waits for it.
  - **Audio Seek** — Where in a file a hovered sound starts playing: **Remember** (`default`) picks it up where the last hover left it, **From the Start** always begins at the beginning, **From the Middle** drops it half way in, and **Random** anywhere at all. The remembered positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so they survive a restart; nothing is remembered while another mode is chosen. Whatever a sound is started at, it goes back to the beginning of the file when it reaches the end of it and loops from there for as long as the hover lasts. A video is always played from its beginning.
- **Performance**
  - **Cache** — What a preview may cost between hovers: `2 GB` down to `0 MB`. **`Image (RAM)`** = decoded frames kept in memory. **`Document (Disk)`** = engine-drawn pages kept as temp files. **`Image (Disk)`** = the pictures ImageMagick developed, kept as temp files.
  - **Decode Budget** — `16 GB` down to `512 MB`; default `1 GB`. A file past it gets no preview.
  - **Tick** — How often the app checks Explorer while a folder window is focused: `15 ms` (`default`), `31`, `47`, `63`, or `78 ms`. Lower answers a move sooner; higher is lighter on CPU and Explorer.
- **Engine**
  - **AFK Timer** — How long Explorer may be unreachable before a non-**Persistent** engine is let go: `1 hour`, `30 minutes`, `10 minutes`, `5 minutes`, `1 minute` (`default`), `30 seconds`, `15 seconds`. Counts time when no Explorer window is reachable on any monitor. A second Explorer window keeps engines warm. Each **`… TTL`** submenu has a **Persistent** toggle at the top.
  - **Select Engine → Office** — Which engine draws Office documents: **Microsoft Office** (`default`) uses the format’s own app and falls back to LibreOffice; **LibreOffice** draws every Office document. LibreOffice is greyed out if not installed.
  - **Microsoft Office TTL** — With **Persistent** on: how long a family’s Office app stays warm: Indefinitely, `1 hour`, `30 minutes`, `10 minutes` (`default`), `5 minutes`, `1 minute`, `0 seconds`. Off: kept while Explorer is reachable, then let go by **AFK Timer**.
  - **LibreOffice TTL** — Same for the engine that draws CorelDRAW and nearby formats. A kept engine converts the next document faster: `1.2 s` cold vs `0.2 s`, but uses a few hundred MB. `0 seconds` means one engine per document. Both TTLs apply to **Persistent** engines; non-persistent ones use **AFK Timer**. Greyed out if LibreOffice is not installed.
  - **WebView2 TTL** — With **Persistent** on: how long the SVG browser stays warm. Off: let go by **AFK Timer**. Greyed out if WebView2 is missing.
  - No `ImageMagick TTL`, `PeaZip TTL`, or `Calibre TTL`: those tools run once and exit, so idle time cannot bound them. A second hover is a cache hit.
- **Codecs** — What this machine has: Videos, Audio, Images, Engines. A missing one carries a cross, and where the README names a page for it, picking the row offers to open that page — nothing is installed or downloaded by the app itself.
- **Run at Startup** — Add or remove the Windows startup entry.
- **Config.ini** — Open the configuration file; named for the running version.
- **Exit** — Close the app.

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
prioritize_keyboard=true
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
audio_seek=remember
audio_volume=10
normalize_video_volume=false
normalize_volume=true
video_volume=0

; Performance
decode_budget_gb=1
document_cache_mb=256
image_cache_mb=64
image_disk_cache_mb=512
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

Here’s what this .INI does:

**Look and text**

- `theme`: `light`, `dark`, or `custom:<name>`.
- `markdown_mode`: show Markdown as `rendered` or `source`.
- `text_preview_full_mode`: `true` adds scrolling, selection, and copy.
- `text_font_scale`: 1–1000%, default `125`; archive lists follow it too.

**What can preview**

- `*_preview_enabled`: turn each preview type on or off. File lists stay normal.
- `extensions` / `names`: which files get text previews. Extensions have no dots. `names` matches files with no extension.
- `image_extensions`, `video_extensions`, `archive_extensions`, `office_extensions`, `font_extensions`, `design_extensions`, `vector_extensions`: per-type preview filters. No dots. An entry with a dot, like `tar.gz`, matches the end of the file name.
- `audio_extensions`: which files are previewed as sounds. One list rather than two, because which engine plays a file is the machine's answer and not a setting: Windows' own decoders are asked first and an installed FFmpeg second.
- `ebook_extensions`, `libre_extensions`, `magick_extensions`, `peazip_extensions`, `calibre_extensions`: pages this app draws itself — PDFs, comic covers, LibreOffice documents, ImageMagick pictures, PeaZip archives, and Calibre books.

**Memory and cache**

- `image_cache_mb`: memory for decoded image frames. Default `64`, max `2048`; `0` holds nothing.
- `document_cache_mb`: disk cache for rendered document pages. Default `256`, max `2048`; `0` keeps nothing between hovers but still draws the current one. Pages live under `%TEMP%\rust-hover-preview\document`; the least recently read page is removed first.
- `image_disk_cache_mb`: disk cache for the pictures ImageMagick developed. Default `512`, max `2048`; `0` keeps nothing between hovers but still develops the current one. Pictures live under `%TEMP%\rust-hover-preview\image`; the least recently read one is removed first, and they survive a restart.
- `decode_budget_gb`: most memory one hover may use. Default `1`, range `0.25`–`64`. A file over the limit shows no preview.

**HDR**

- `hdr_tone_map`: how HDR/EXR light becomes screen values: `reinhard` default, `aces`, `srgb`, or `off`. PNG/JPEG are not affected.
- `hdr_exposure`: stops shifted before that curve. Default `0`, range `-10` to `10`.

**Waiting and engines**

- `spinner_delay_ms`: wait before showing the loading spinner. Default `250`; `0` shows it immediately. One delay covers all preview types.
- `office_engine`: `microsoft_office` default, falls back to LibreOffice when needed; or `libreoffice`, which always uses LibreOffice. If LibreOffice is missing, it falls back and the tray row is greyed out.
- `libreoffice_idle`: seconds LibreOffice is kept after its last page, or `indefinitely`. Default `600`; `0` launches it per document.
- `office_engine_idle`: same idea for the Office engine. Default `600`; `0` lets it go after drawing a page.
- `afk_timer_seconds`: seconds with no Explorer window before a non-persistent engine is released. Default `60`, max one day. `0` releases immediately.
- `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent`: `true` keeps that engine always, bounded only by its `…_idle` time. `false` default keeps it while Explorer is reachable, bounded by `afk_timer_seconds`.

**Trigger and position**

- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled`: the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched. Default `true`.
- `follow_cursor`: `true` = Follow Cursor; `false` = Best Position.
- `avoid_mode`: `filename` default, `filename_column`, `details`, or `off` — what the preview avoids.

**Scaling**

- `preview_scale`: percentage or `fit`, based on the picture’s own size.
- `video_scale`: same for video. Default `100`.
- `animated_scale`: same for animated GIF, WebP, or PNG. Default `100`. Still GIF/PNG follow `preview_scale`.
- `vector_scale`: percentage or `fit`, based on the screen. Default `fit`; `100` or more reads as `fit`. Covers all Vector drawings. Old `svg_scale` is ignored and removed.
- `ebook_scale`: PDF page scale against the screen. Default `fit`; `100` or more reads as `fit`.
- `document_scale`: document page scale against the screen. Default `fit`. A workbook’s fallback bitmap keeps its own size and is never enlarged.
- `font_scale`: font preview scale against the screen. Default `50`; `fit` means all of it; `100` or more reads as `fit`.
- `design_scale`: design document scale against the screen. Default `fit`.

**Fonts**

- `ttc_face`: which face of a `.ttc` is drawn. `1` is first, max `10`. The heading shows which face came out.

**Backgrounds**

- `image_background`: `checkerboard` default, `black`, `white`, or `transparent`. Also used for PDF pages, painted text frames, and document pages. Old `transparent_background` is ignored and removed.
- `font_background`: `white` default, `black`, `checkerboard`, or `transparent`.
- `dds_background`: `white` default or `black`, and only those two. Any other value reads as `white`.
- `vector_background`: `checkerboard` default, `white`, `black`, or `transparent`. Used for SVG pages and metafiles. Old `svg_background` is ignored and removed.
- `design_background`: `checkerboard` default, `white`, `black`, or `transparent`.

**Old or removed names**

- `magick_preview_enabled` and `peazip_preview_enabled`: not read. ImageMagick pictures use `image_preview_enabled`; PeaZip archives use `archive_preview_enabled`.
- `libre_preview_enabled` and `libre_scale`: gone. Use `document_preview_enabled` and `document_scale`.
- `office_cache_mb` and `libre_cache_mb`: gone. Use `document_cache_mb`. If both old ones exist, the larger is read once, then both are removed.
- No `magick_idle`, `peazip_idle`, or `calibre_idle`. ImageMagick, PeaZip, and Calibre are converters, not engines kept open. A second hover usually costs only a cache hit.

**Extension list behavior**

- A deleted `extensions=` line or whole section comes back with built-in entries.
- An `extensions=` line left empty stays empty.

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