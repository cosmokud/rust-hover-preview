# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.98.1+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

Rust Hover Preview is a Windows 11 tray app inspired by QTTabBar and QuickLook. Hover over a file in File Explorer for an instant preview, or middle-click (or press `Space`) to open a pinned preview window. Works with keyboard navigation, and customizes easily via the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer, beside the cursor or the focused item and kept on screen.
- Press the middle mouse button or `Space` (or your own `pin_key`) and a preview becomes a window of its own — captioned, movable, resizable, and there until you close it, always on top unless `Pin Mode → Always On Top` is turned off. 
- Supports images, videos, audios, ebooks, Office documents, design documents, text/code, archives, camera raw, vector drawings, fonts, and more.
- Granular display scaling: Flexible sizing (fit-to-screen, percentage-based, or 25%–400% native scale) with independent scaling profiles for media vs. text, vectors, and documents.
- Tray menu and hand-editable `config.ini`; DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.
- Plugs into what you have: built-in decoders and Windows codecs cover the basics; ImageMagick, FFmpeg, mpv, LibreOffice, Office, WebView2, PeaZip, Calibre, and the Windows PDF engine extend it where installed — nothing bundled.

## Supported Formats

### Images


| Engine                                                                   | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| --- | --- | --- |
| This app's own decoders and the codecs Windows ships                     | `[image]`    | apng, avci, avif, bmp, dds, exr, ff, gif, hdr, heic, heif, ico, jfif, jpe, jpeg, jpg, jxl, pam, pbm, pgm, png, pnm, ppm, qoi, tga, tif, tiff, webp                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| [ImageMagick](#optional-enable-camera-raw-and-more-pictures-imagemagick) | `[magick]`   | 3fr, aai, art, arw, ase, aseprite, bayer, bayera, bgr, bgra, bgro, c2pa, cal, cals, cmyk, cmyka, cr2, cr3, crw, cube, cur, cut, dcm, dcr, dcx, dng, dpx, dxt1, dxt5, erf, fax, fff, fit, fits, fl32, fts, ftxt, g3, g4, gray, graya, group4, hrz, icb, iiq, ipl, j2c, j2k, jng, jnx, jp2, jpc, jpm, jpt, k25, kdc, mac, map, mat, mdc, mef, miff, mng, mono, mos, mpc, mrw, mtv, nef, nrw, orf, otb, pal, palm, pango, pcds, pef, pes, pfm, pgx, phm, picon, pict, pix, pwp, raf, raw, rgb, rgb565, rgba, rgbo, rgf, rla, rle, rmf, rw2, rwl, scr, sct, sf3, sfw, sgi, six, sixel, sr2, srf, srw, stegano, sun, tim, tm2, uyvy, vda, vicar, viff, vips, vst, wbinfo, wbmp, x3f, xbm, xcf, xpm, xv, ycbcr, ycbcra, yuv |


`heic`, `heif`, `avif`, `avci`, `jxl` and a still `webp` are decoded by a codec Windows provides rather than one this app carries, so each of those six stills needs its [Store extension](#optional-enable-heic-avif-jpeg-xl-and-webp-preview-windows-codecs) installed once; every other entry of `[image]` needs nothing installed.

### Videos

For video, **mpv is recommended** for performance, hardware acceleration, reliability,  and subtitles (embedded tracks and sidecar `.srt`) that just work.

| Engine                                                                   | `config.ini`           | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                               |
| --- | --- | --- |
| The media engine Windows 11 ships                                        | `[video]`              | 3g2, 3gp, 3gpp, asf, avi, dvr-ms, m1v, m2t, m2ts, m2v, m4v, mkv, mov, mp4, mpe, mpeg, mpg, mts, qt, ts, vob, webm, wmv                                                                                                                                                                                                                                                                                                                           |
| [FFmpeg](#optional-enable-video-preview-and-advanced-audio-processing-with-ffmpeg)'s player (`ffplay`) | `[ffmpeg]`             | 264, 265, 266, apv, av1, avc, avs, avs2, avs3, bik, bk2, c93, cavs, cdg, cdxl, cin, cpk, dav, dif, divx, drc, dv, evc, f4v, flm, flv, gxf, h261, h263, h264, h265, h266, h26l, hevc, ifv, imx, ismv, ivf, ivr, kux, m2p, mj2, mjpeg, mjpg, mk3d, moflex, mpv, mve, mvi, mxf, mxg, nsv, nut, obu, ogm, ogv, pmp, psp, rcv, rm, rmvb, roq, rsd, smk, str, swf, thp, tod, tp, tr, ty, ty+, usm, vc1, vc2, viv, vro, vvc, vw, wtv, xl, xmv, y4m, yop |
| [mpv](#optional-enable-video-preview-with-mpv)                           | — (no list of its own) | union of `[video]` + `[ffmpeg]` above (23 + 84 = 107 names)                                                                                                                                                                                                                                                                                                                                                                                      |


### Audio


| Engine                                                                                                                                   | `config.ini` | Default extensions                                                                                                                                                                                     |
| --- | --- | --- |
| The media engine Windows ships, then [FFmpeg](#optional-enable-video-preview-and-advanced-audio-processing-with-ffmpeg)'s player — the machine's answer, not a setting | `[audio]`    | aac, ac3, aif, aifc, aiff, amr, ape, au, awb, caf, dff, dsf, dts, dtshd, eac3, flac, m4a, m4b, mka, mp2, mp3, mpa, mpc, oga, ogg, ofr, ofs, opus, ra, shn, snd, spx, tak, tta, voc, wav, wave, wma, wv |


### Text


| Engine                       | `config.ini`          | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| --- | --- | --- |
| This app, syntax-highlighted | `[text]` `extensions` | adb, adoc, ads, asciidoc, asm, asp, aspx, astro, awk, bash, bat, bib, bzl, c, cc, cfg, cg, cjs, clj, cljc, cljs, cmake, cmd, comp, conf, cpp, cs, csh, cshtml, css, csv, csx, cts, cxx, d, dart, diff, diz, edn, ejs, el, elm, env, erb, erl, ex, exs, f, f03, f77, f90, f95, fish, for, frag, fs, fsi, fsx, ftn, fx, geom, glsl, go, gql, gradle, graphql, groovy, h, haml, hbs, hcl, hh, hlsl, hpp, hrl, hs, htm, html, hxx, inc, ini, ipynb, java, jl, js, json, json5, jsonc, jsonl, jsp, jsx, ksh, kt, kts, latex, less, lhs, liquid, lisp, ll, lock, log, lsp, lua, m, mak, man, markdown, md, mdown, metal, mjs, mk, mkd, ml, mli, mm, mts, mustache, nasm, nfo, nim, ninja, nix, njk, org, pas, patch, php, phtml, pl, plist, pm, properties, proto, ps1, psd1, psm1, py, pyi, pyw, r, rake, rb, rkt, rmd, rs, rst, rtf, s, sass, scala, scm, scss, sh, slim, sol, sql, srt, ss, styl, sv, svelte, svh, swift, tcl, tex, text, tf, tfvars, toml, ts, tsv, tsx, twig, txt, v, vbs, vert, vhd, vhdl, vtt, vue, wat, wgsl, xhtml, xml, xsd, xsl, xslt, yaml, yml, zig, zsh |
| This app, as plain text      | `[text]` `names`      | authors, .babelrc, brewfile, caddyfile, changelog, changes, .clang-format, .clang-tidy, cmakelists.txt, code\_of\_conduct, containerfile, contributing, contributors, copying, copyright, dockerfile, .dockerignore, .editorconfig, .env, .env.example, .env.local, .eslintignore, .eslintrc, gemfile, .gitattributes, .gitconfig, .gitignore, .gitkeep, .gitmodules, gnumakefile, .golangci.yml, history, .htaccess, install, jenkinsfile, justfile, licence, license, .mailmap, makefile, makefile.am, makefile.in, notice, .npmignore, .prettierignore, .prettierrc, procfile, rakefile, readme, .rustfmt.toml, security, .stylelintrc, unlicense, vagrantfile                                                                                                                                                                                                                                                                                                                                                                                                               |


### Ebook


| Engine                                                                                                      | `config.ini` | Default extensions                                                          |
| --- | --- | --- |
| The Windows PDF engine for the PDF's three spellings; this app's own readers for the three comic containers | `[ebook]`    | cbc, cbr, cbz, epdf, pdf, pdfa                                              |
| [Calibre](#optional-enable-ebooks-calibre), which converts the book to a PDF first                          | `[calibre]`  | azw, azw3, azw4, djvu, epub, fb2, htmlz, lit, lrf, mobi, pml, prc, snb, tcr |


### Archives


| Engine                                              | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                |
| --- | --- | --- |
| This app, from each archive's own table of contents | `[archive]`  | 7z, apk, jar, rar, tar, tar.gz, tgz, xpi, zip, zipx                                                                                                                                                                                                                                               |
| [PeaZip](#optional-enable-niche-archives-peazip)    | `[peazip]`   | 001, apfs, ar, arc, arj, bcm, br, bz2, bzip2, cab, chm, cpio, cramfs, deb, dmg, esd, gz, gzip, hfs, hfsx, hxs, iso, lha, lpaq8, lzh, lzma, msi, msp, pkg, ppkg, qcow, qcow2, rpm, squashfs, swm, taz, tbz, tbz2, tpz, txz, tzst, udf, udeb, vdi, vhd, vhdx, vmdk, wim, xar, xip, xz, z, zpaq, zst |


### Document


| Engine                                                                                                                             | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                              |
| --- | --- | --- |
| Microsoft Office, falling back to [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice) where Office is absent | `[office]`   | doc, docm, docx, dot, dotm, dotx, pot, potm, potx, pps, ppsm, ppsx, ppt, pptm, pptx, xls, xlsb, xlsm, xlsx, xlt, xltm, xltx                                                                                                                                                                                                                                                                                                     |
| [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice)                                                          | `[libre]`    | 123, 602, abw, cdr, cgm, cmx, cwk, dbf, dif, dxf, fodg, fodp, fodt, gnm, gnumeric, hwp, key, lwp, mcw, met, mw, numbers, odb, odc, odf, odg, odm, odp, ods, odt, oth, otg, otm, otp, ots, ott, pages, pcd, pct, pcx, pdb, pm6, pmd, psw, pub, ras, sda, sdc, sdd, sdw, slk, stc, std, sti, stw, svm, sxd, sxg, sxi, sxm, sxw, vdx, vsd, vsdm, vsdx, vstx, wb2, wk1, wk3, wk4, wks, wpg, wq1, wq2, wpd, wps, wri, xlw, zabw, zmf |


### Vector


| Engine                                                                                                    | `config.ini` | Default extensions                                         |
| --- | --- | --- |
| WebView2 for `svg` and `svgz`; the drawing layer for the metafiles and the preview picture an EPS carries | `[vector]`   | emf, epi, eps, epsf, epsi, ept, ept2, ept3, svg, svgz, wmf |


### Fonts


| Engine                                         | `config.ini` | Default extensions         |
| --- | --- | --- |
| WebView2 — a specimen page with the font in it | `[font]`     | otf, ttc, ttf, woff, woff2 |


### Design


| Engine                                                                       | `config.ini` | Default extensions                                 |
| --- | --- | --- |
| This app, from the flattened picture each format keeps of the whole document | `[design]`   | ai, fig, kra, ora, procreate, psb, psd, sketch, xd |


## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

1. Open [Releases](../../releases).
2. Download your preferred asset.
3. Run the installer, or place the portable binary wherever you like.
4. Launch Rust Hover Preview.

No Rust toolchain is needed.

### Optional: Enable Video Preview with mpv

For smoother playback and a better picture, install mpv: videos then play through it automatically, using your graphics card, with subtitles just working — including any `.srt` file beside the video. 

The app finds `mpv.exe` beside itself, in your `PATH`, in `MPV Player` under Program Files, and in SVP 4's `mpv64`; where more than one is found, **Codecs → Videos → MPV** lets you pick which one plays.

mpv ships no installer of its own, so grab a Windows build from [mpv's own installation page](https://mpv.io/installation/).

**Option A: winget**

```powershell
winget install --id shinchiro.mpv -e
```

**Option B: manual**

1. Download a Windows build from [https://mpv.io/installation/](https://mpv.io/installation/)
2. Extract it, for example to `C:\mpv`
3. Add `C:\mpv` to your user `PATH`

Verify either way:

```powershell
mpv --version
```

### Optional: Enable Video Preview and Advanced Audio Processing with FFmpeg

FFmpeg adds the video formats and the audio ones Windows does not decode — the `[ffmpeg]` list, and the `[audio]` names Windows' own decoders do not reach. It also plays any `[video]` name the media engine turns down, or every video there is, where **Engine → Video Engine** names FFmpeg. `ffplay` and `ffprobe` need to be in your `PATH`.

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

### Optional: Enable More Video Codecs (Windows Codecs)

The media engine decodes H.264, MPEG-4, and WMV out of the box; each codec below is a separate free extension from the Microsoft Store, and none of them is needed while FFmpeg is installed.


| Codec                                  | Needs                                                                    |
| --- | --- |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)  |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)   |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)    |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4) |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)   |


All are free, official from Microsoft. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray's **Codecs** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

A **still** `heic`, `heif`, `avif`, `jxl`, and still `webp` is decoded by a codec Windows provides rather than one shipped with the app. Each still needs its extension installed once from the Microsoft Store:


| Still format   | Needs                                                                                                                                            |
| --- | --- |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |


All are free, official from Microsoft. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error.

Which of these a moving file needs is under [Supported Formats](#supported-formats) above.

A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

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

Or download the official installer from [https://calibre-ebook.com/download_windows](https://calibre-ebook.com/download_windows) and run the suggested file (`calibre-*-64bit.msi`).

> [!WARNING]
> If `winget` reports an error, the package sources are usually why: run `winget source reset --force`.
>
> If that is refused as well, open **Terminal as administrator** (right-click the Start button → *Terminal (Admin)*) and run the same command from there.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys, Tab, or a file's name typed to jump to it — to preview the focused item.
4. Use the middle mouse button or `Space` (by default) to pin the preview.
5. Right-click the tray icon to configure behavior.

Pin mode works like Quick Look: pressing the middle mouse button or `Space` (`pin_key`, watched while `pin_enabled`)
turns the settled preview into a window of its own — always on top unless **Pin Mode → Always
On Top** is off, in which case it is an ordinary window that also wears this app's own taskbar
button while it is up. `Space` in a focused pin pauses and resumes a sound; a minimized pin
sits in a draggable bubble — or in that taskbar button — that a middle click or `Space` in Explorer brings back.
Arrow keys walk siblings, `Escape` closes. A right-click on a pinned window opens its own menu,
and a pinned page of text takes the keyboard: the arrows and page keys scroll it, `F3` finds
the next place the selection appears, and `Ctrl` and a wheel sizes it.

## System Tray Menu

A setting marked `(Default)` is what an untouched setting would be. The check or radio mark shows what is set now.

- **Enable Preview** — Turn previews on or off. The row names the shortcut that throws the same switch while File Explorer is in front — `Enable Preview (F7)` unless `toggle_key` names another key.
- **Pin Mode** — Everything about pinning, in one place.
  - **Enable (Mouse3, Space)** — Turn pinning on or off.
    - Pressing the key while a preview is up turns it into a window of its own: captioned, movable, resizable, always on top unless **Always On Top** is off, and still there when the pointer leaves.
    - The caption carries **Previous**, **Next**, **Open With**, and **Open With...**.
      - **Previous** / **Next** — Step the pin through the sibling files in the folder it was taken up in, in the order the listing is showing them; step over any one that will not open rather than stopping there.
      - **Open With** — Opens the file in whatever program Windows has registered for it.
      - **Open With...** — Asks you to pick one from the list Windows keeps of the programs that could.
      - The two hand-off buttons name themselves while the pointer rests on them; **Open With** names the program it would use.
    - A **right-click** on a pinned window opens the pin's own menu — the panel a collapsed pin's bubble opens: the hold, **Next File**, **Previous File**, **Pause on Minimize** where the kind plays something, the two picks of where a minimize sends the pin, **To Bubble** / **To Taskbar** to put it away at once, then **Close**. A pinned page carries **Select All**, **Copy** and **Find Next** above them, and a sound's card keeps its own menu. A film's title bar is the way to it, since the player's own band takes the clicks in the picture.
    - A caption the hand has narrowed past the point of saying the file's name leaves the name out, and one narrowed past ten pixels leaves the strip out altogether — a row of window buttons over the top of somebody's picture is not what the strip is for. Both come back as the window is widened.
    - Its edges resize it; what that does depends on what is inside.
      - A picture, video, or rendered page keeps its own shape.
      - A document or listing is laid out to whatever box it is given.
      - A sound's card keeps the size **Audio** names, and a video `ffplay` has is not resized at all.
    - A picture's caption and bar are drawn over the media rather than in bands around it, and fade in as the pointer arrives; only the strip under the pointer is shown.
    - A wheel over a pinned picture zooms it: the window grows about the picture's middle to the size its scale's fit names, then the picture grows inside the window that has stopped, about the cursor. A wheel back out takes it down the way it came, and a picture the hand has carried off its window comes home over the last of the zoom with a small swing past its place. A film is not zoomed — its picture belongs to its player.
    - A pinned video carries a transport bar of its own: play/pause, a draggable seek bar, and the two clocks.
      - Where FFmpeg's player is the one playing it, a read-out bar appears instead, since that player can be told nothing.
    - A pinned sound's card carries four buttons of its own: **Previous**, **Play/Pause**, **Next**, and **Volume**.
      - `Space` holds a sound where it stood and sets it going again from there.
      - The mark beside the name is a dot the ripples of a beat travel out from: the bass under the sound sends them, and a pause brings them home. Which band they follow is `beat_hz` in `config.ini` — `off`, or a top of band between `20 Hz` and `16 kHz` (`250 Hz` by default). The dot is centred on the name line's own ink, and `audio_mark_offset_y` shifts it up or down from there in pixels (`0` by default, `-100` to `100`).
      - A click on a sound's bar takes the file to that second.
      - A right-click on the card opens the card's own menu in a panel of its own over the screen: the whole of what a pin does — the two ways out, the file's own playback, then **Shuffle**, **End of File** and **Seek**.
    - A pinned text preview comes up in full mode, so it scrolls and its text can be selected and copied.
      - `Up`, `Down`, `Page Up` and `Page Down` move a screenful at a time, and `Home` and `End` take the page to its top and its last screenful.
      - `Ctrl+A` selects all of it; `Ctrl+C` copies what is selected; `F3` takes the page to the next place what is selected appears.
      - `Ctrl` and a wheel steps the page's own text size through the sizes **Text → Font Size** offers; the size rides across a resize and the next file of a walk, and closing the pin gives it up.
      - A press begins a selection, `Shift` grows it, and a second press takes the word under the pointer; a drag carried past the page's edge scrolls it three lines a tick, a wrapped line copies whole, and **Select All** means the page rather than the screenful.
      - A page of code keeps a column of line numbers down its left.
    - A pinned video has a transport bar with a volume of its own, set on the window; it does not change **Volume → Video**.
    - **Maximize** fits the window to the screen.
    - A **double-click** puts the window back to the size its kind's scaling setting names and in the middle of the screen it stands on (both halves are switches in `config.ini`: `pin_double_click_reset_size` and `pin_double_click_center`, both on).
    - **Minimize** puts the window away, either into a round bubble on the desktop that can be dragged anywhere and clicked to bring the window back (a film's bubble keeps showing the film, and holds the frame it stopped on when it is paused; a right-click on the bubble opens its own menu: the same playback rows, then **Pause on Minimize** where the kind behind it plays something, then the ways out; **Play** there gives back a park the collapse made, without waiting for the window), or into the app's own taskbar button — see **Minimize** under **Pin Mode** below.
    - Clicking the window gives it the keyboard; then the arrow keys step it and `Escape` closes it.
    - The pin key itself only ever puts a pin up and brings back a bubble; it never hides one.
    - The key is named in the item and can be changed in `config.ini` (`pin_key`).
  - **Always On Top** — On by default: a pin stays above every other window. Off, the pin is an ordinary window that other programs can cover — it keeps the place your own clicks gave it, and it wears this app's own taskbar button while it is up. The mark left of the window buttons is the same switch, and the wheel over such a pin is the pin's only while the pin is the window you are in.
  - **Update Preview** — Whether a pin that is up is shown the file you pick next.
    - **Follow Selection** — On by default.
      - A file you click, or one the keyboard selects, is shown where the window already stands, rather than as a second preview beside it.
      - A file named this way that cannot be shown leaves the window on the file it already had.
      - Stepping over a file that will not open belongs to the caption's own **Previous** and **Next**, which you pressed.
      - A sound's card is drawn at the size its **Audio** setting names, in the middle of that box; it follows **System → Theme**, and no longer **Font Size**.
      - The keyboard half follows Explorer, so it pauses while a pinned window holds the keyboard; click back into the listing and it carries on.
    - **On Hover** — Off by default.
      - On, the pointer's own hover is one of the ways a pin is told about a file, the way a Quick Look window follows a listing.
      - Greyed while **Follow Selection** is off.
  - **Minimize** — Where a window put away by its **Minimize** button goes, and what it does with what it was playing while it is away.
    - **To Bubble** (Default) — The window collapses into the round bubble on the desktop, and this app never appears in the taskbar.
    - **To Taskbar** — The window leaves the screen and the app appears in the taskbar as its own button, named for the pinned file and wearing this app's icon; its hover holds the pin — the live picture while the window is up, the frame it went down on once it is away, and **Previous**, **Play/Pause** and **Next** for a film or a sound. Clicking the button brings the pinned window back where it stood, and a right-click closes the window. A pin that is not **Always On Top** wears this button while it is up as well.
    - A window is away either way, so previews stay held back until it is brought back, whichever place the setting names.
    - What a window put away holds where it is is **Pause on Minimize** at the top of each kind's own **Pin Mode** — one under **Preview Settings → Video**, one under **Preview Settings → Audio** — because what should go quiet while a pin is away is a question about one kind of file.
  - **Navigation** — What **Previous** and **Next** on the caption step through: three anchors — **Video**, **Audio**, **Others** — each with **All Files** / **Same Category**. The walk is anchored on the first file pinned, so a video walks videos while a document walks end to end.
    - The walk is the pin's own: one playlist spawned when the pin is taken up and stepped along, so a **Previous** is the file the walk came from and the **Next** after it is the file it was interrupted on. A file picked in the listing inside that folder carries the playlist on; one anywhere else starts fresh.
    - Video and Audio start at **Same Category**, Others at **All Files**.
    - Both work with **Update Preview** off; they are a thing you pressed, and no setting asks for them.
    - The arrow keys step the same way, and are asked for by nothing at all — but only once you have clicked the pinned window, which gives it the keyboard. Until you do, the arrows belong to whatever is in front.
    - `Space` in a pinned sound pauses and resumes it, on the same terms.
    - `Escape` in a pinned window closes it.
    - **All Files** — Every file this build can preview, in the folder the pin was taken up in and not any subfolder of it, in the order the Explorer listing is showing them. A video sits beside a sound.
    - **Same Category** — Only the files of the walk's anchor category: video, audio, or the anchor's own kind for Others.
      - A camera raw is a picture and a book is a document, because that is what you call them.
      - A kind switched off under **Preview Types**, or one with no engine installed on the machine, is not a step either way.
      - A file this build *could* preview but that will not open is a step the window steps over rather than one it stops on.
- **Preview Types** — Choose which file kinds can preview: Image, Video, Audio, Text, Ebook, Archive, Document, Vector, Font, Design.
  - A switch is a switch over behaviour, not over files: the lists deciding which files are previewed are untouched, so switching a kind off and back on restores what was configured.
  - One switch covers both the original file and the engine-drawn preview.
    - Camera raw uses **Images**.
    - A LibreOffice document uses **Document**.
    - A PeaZip archive uses **Archives**.
    - A Calibre book uses **Ebook**, as do the app-drawn PDF and comic pages.
- **Preview Settings** — What each kind's previews are like, below **Preview Types**: one half per kind that has settings beyond whether it previews at all, named for the kind and nothing else. What a film and a sound are heard *at* is the **Volume** submenu below **Background**, since the level is the one question both kinds answer the same way.
  - **Video** — What a film's previews are like: whether a hovered one goes round, and what a *pinned* one does.
    - **Subtitles** — Whether a film's own subtitle tracks are probed for and copied out at all, which is what **General (Disk)** holds between hovers; `0` there draws a film without them.
      - Where the engine that will play is mpv the row reads **Subtitles (MPV)** and is not about a copy at all: mpv finds a film's own tracks and the `.srt` beside it and draws them itself, so nothing is extracted and the pinned `T` key does nothing on that route.
    - **Loop** — On by default: a hovered film goes round at the end of its file. Off, it stops there.
    - **Pin Mode** — What a *pinned* film does.
      - **Pause on Minimize** — On by default: a film put away by the **Minimize** button is held where it stood and started again at the second it stopped at when the window comes back, rather than playing on behind a bubble nobody can see.
      - **Shuffle** — Off by default: whether a forward step of the walk takes a random film of the anchor's kind rather than the one beside it. A step back is the file the walk came from.
      - **End of File** — What a pinned film does at the end of its file: **Loop** (default), **Pause** (the file left where it ended, the bar and the play button standing there), or **Jump** — the walk's next file, which the **Shuffle** row decides between the file beside it and a random one of the same kind.
  - **Audio** — The same questions for a sound, and the two places a sound starts: where a hovered one does and where a *pinned* one does.
  - **Loop** — On by default: a hovered sound goes round at the end of its file. Off, it stops there.
  - **Seek** — Where in a file a hovered sound starts playing.
    - **Remember** (`default`) — Picks it up where the last hover left it.
    - **From the Start** — Always begins at the beginning.
    - **From the Middle** — Drops it half way in.
    - **Random** — Anywhere at all.
    - The remembered positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so they survive a restart; nothing is remembered while another mode is chosen.
    - A video is always played from its beginning.
  - **Pin Mode** — What a *pinned* sound does.
    - **Pause on Minimize** — On by default: the sound a pin put away was playing is held where it stood and started again at the second it stopped at when the window comes back.
    - **Shuffle** and **End of File** — the same two rows as a film's, asked of a sound: a forward step takes a random sound rather than the one beside it, and a file that ends goes round, holds where it stopped, or steps on to the next.
    - **Seek** — Where in a file a *pinned* sound starts playing: **From the Start** (default), **Remember**, **From the Middle**, or **Random**.
  - **Text** — What a text preview is painted with: the font size, how Markdown is shown, and whether a page of HTML runs.
    - **Font Size** — `400%` at the top down to `70%` at the bottom.
    - **Markdown** — Rendered or Source.
    - **Render HTML** — Off by default.
      - On, a `.htm` or `.html` file is previewed as the page it holds rather than as its markup, run by the browser engine.
      - A page is the one thing previewed as a page rather than as a picture, so it is the one thing that runs, and a page that draws itself with script has nothing to show without a run.
      - What it may still do is bounded: the frame it is shown in cannot open forms or popups or navigate, and nothing a page links to is fetched.
      - SVG documents and font specimens, which a browser is handed as an image, are still drawn rather than run.
      - A running page can be pointed at, clicked into and typed into; a document, a specimen and every other preview still cannot.
    - Without the engine on the machine it stays a text preview.
  - **Image** — **Pin Mode**, and the whole of it: **Shuffle**, whether a *pinned* picture's walk takes the next file at random, and **Zoom Background**, the backdrop the band a zoom leaves behind the picture is filled with.
    - A picture has no level, no end of file and nowhere in it to start, and a hover has no next file, so the shuffle is the one thing a picture answers.
    - **Pin Mode → Zoom Background** — **Black** (default), **White** or **Checkerboard**: the backdrops that fill the band a hand carries a zoomed picture off its window leaves behind it. The see-through one is not offered, since the band is the window's rather than the file's. Written as `zoom_background`. A picture with a transparent pixel in it keeps what **Background → Image** asks for instead — see-through included — so one meant to show the desktop through it keeps doing so at any size, zoom or not.
- **Timing**
  - **Prioritize Keyboard** — On by default.
    - The file under a pointer that has not been moved does not preview of its own while the keyboard is driving Explorer, so pressing a key onto a file with no preview of its own behaves like pressing one onto a file that has a preview.
    - The pointer takes the screen back when it is moved, when the wheel is turned, or when a folder change hands it over.
    - Off, the pointer's own hover always wins.
  - **Trigger Key (Alt)** — The key is named in the item.
    - **Enable Trigger Key** — Whether the key is watched.
    - **Affect Pin Mode** — Off by default, and off at the start.
      - A pinned window is one you put there, so the key that holds back hovers is not read while one is up.
      - On, holding it brings the pin down along with the previews it stops.
      - It speaks for **Hold to Disable Preview**.
    - **Hold to Disable Preview** / **Hold to Enable Preview** — What holding the key does.
  - **Hover Delay** — How long the pointer rests before a preview opens: `0 ms` at the top down to `1000 ms` at the bottom. Default `0 ms`.
  - **Rehover Delay** — Wait before the same file can preview again. Default `200 ms`.
  - **Settling Delay** — How still the pointer must be before previewing what it is on.
    - Default `0 ms` means a new file can preview while the hand is still moving.
    - Keyboard previews ignore this.
- **Placement**
  - **Position** — Follow Cursor or Best Position. Default **Best Position**.
  - **Avoid** — Nothing, Filename (`default`), Filename Column, or Details.
    - Keeps a preview off the item it is about.
    - Keyboard previews have no cursor, so **Nothing** acts like **Filename**.
- **Scaling**
  - **Image** — **Fit to Screen**, a share of the display's fitted size (`75%`, `50%`, `25%`, `10%`), or a share of the image's own size (`400%`–`25%`). The two groups are listed under **By Screen** and **By Own Size**. Default `100%`.
  - **Video** — Same shares for a video, except that **Fit to Screen** is the display's room: a film smaller than the screen fills it, as a pinned window and a maximize already do. Default `100%`.
  - **Animated** — Same shares for an animated GIF, WebP, or PNG — and for an animated JPEG XL, which plays through the same machinery.
    - A still GIF or PNG uses **Image**.
    - Default `100%`.
  - **Audio** — `25%`, `20%`, `15%`, `10%` (`default`), `7%`, or `5%` of the display, for a sound's card — the whole card scales with the share: font, height and width, the way Windows scaling sizes a window.
    - The card at any share is the `10%` card uniformly smaller or larger (~188×50 at `5%` to ~936×250 at `25%` on a 3440×1440 display), its seek bar stretching with the width, its height what its content measures at the share's font.
    - Built at the share's own font (the default text size, `125%`, at the `10%` anchor), so **Font Size** does not resize it; a pinned sound's card is the same size as the hover's.
  - **Vector** — Fit to Screen (`default`), or `75%`, `50%`, `25%`, `10%` of the display.
  - **Text** — Same display shares for a text preview — a plain text file, code, or Markdown.
    - It caps how much of the screen the preview box may take when it opens; a short file still gets the small box its own text needs.
    - Default **Fit to Screen**.
  - **Ebook** — Fit to Screen (`default`), or the same display shares for a PDF page, a comic's first page, and a Calibre-converted book.
  - **Document** — Same display shares for a document drawn as a page, whether by its own Office app or by LibreOffice.
    - A workbook's fallback bitmap is never enlarged.
  - **Font** — Same shares for a font specimen. Default `50%`.
  - **Design** — Same display shares for design documents: Photoshop, Illustrator, Krita, OpenRaster, Procreate. Default **Fit to Screen**.
- **Background**
  - **Image** — Transparent, Black, White, or Checkerboard. Default **Checkerboard**.
  - **Vector** — Same backdrops for an SVG document or metafile. Default **Checkerboard**.
  - **HTML** — White, Black, or Checkerboard for a page of HTML. Default **White**.
    - The see-through backdrop is not offered: a page is drawn on a page, whether it runs or not, and the backdrop is behind the page rather than behind what it paints.
  - **Font** — Same backdrops for a font specimen. Default **White**.
  - **DDS** — Black or White for `.dds` textures. Default **White**.
    - The two see-through backdrops are not offered.
  - **Design** — Same as picture backdrops for a design document. Default **Checkerboard**.
- **Volume** — What a film and a sound are heard at: one menu with a **Video** half and an **Audio** half, below **Background** and above **Performance**. Each half carries **Normalize (Expensive)** (loudness measured and brought to `-14 LUFS`; greyed without FFmpeg) and **Remember Level** above its levels.
  - **Video** — `100%` down to `0%`, loudest first. A film's soundtrack starts at `0%`, so a hover never makes a sound the pointer did not ask for.
  - **Audio** — The same levels; a sound file starts at `10%`, and a sound at `0%` still shows its card, silently.
- **Performance**
  - **Cache** — What a preview may cost between hovers: `Image (RAM)` offers `16 GB` down to `0 MB`, the three disk caches `128 GB` down to `0 MB`, and a size that is a whole gigabyte is named in it (`4 GB`, `128 GB`).
    - **`Image (RAM)`** = decoded frames kept in memory.
    - **`Document (Disk)`** = engine-drawn pages kept as temp files.
    - **`Image (Disk)`** = the pictures ImageMagick developed, kept as temp files.
    - **`General (Disk)`** = the subtitle tracks a film's embedded ones are copied into, kept as temp files; `0` draws a film without its subtitles.
  - **Decode Budget** — `16 GB` down to `512 MB`; default `1 GB`. A file past it gets no preview.
  - **Explorer Poll** — How often the app checks Explorer while a folder window is focused: `15 ms` (`default`), `31`, `47`, `63`, or `78 ms`.
    - Lower answers a move sooner; higher is lighter on CPU and Explorer.
  - **Hardware Acceleration (MPV Only)** — checked by default: a film mpv plays is decoded on the graphics card through mpv's own `--hwdec=auto`, whose safe list and software fallback are that player's business.
    - It is mpv's switch and nobody else's: FFmpeg's player is given no decode device, and the bubble's own frame reader decodes on a core.
    - A machine with no `mpv.exe` has no player for the switch to act on, so the row is greyed and unchecked there and a `config.ini` that says on is read as off.
- **Engine**
  - **Away Timer** — How long Explorer may be unreachable before an engine set to **Until away** is let go: `1 hour`, `30 minutes`, `10 minutes`, `5 minutes`, `1 minute` (`default`), `30 seconds`, `15 seconds`.
    - Counts time when no Explorer window is reachable on any monitor.
    - A second Explorer window keeps engines warm.
  - **Video Engine** — Which engine plays a video.
    - **Best** (`default`) — The machine's own answer: mpv for every video where it is installed, and otherwise the media engine for a film at or under 3.2 megapixels and FFmpeg's player for a larger one or one it could not measure.
    - **MPV** — mpv alone.
    - **FFmpeg** — `ffplay` alone.
    - **Native** — The media engine alone.
    - **Native (FFmpeg above 3.2MP)** — That size rule stated plainly.
    - A **Fallback** switch at the top (`default` on) lets a chosen engine that cannot play a given file fall through to the engines it may: **MPV** reaches both of the others, while **Native** and **FFmpeg** still only reach each other. Off, the engine you chose stands alone.
    - **MPV** is greyed out where `mpv` is not installed; **FFmpeg** and **Native (FFmpeg above 3.2MP)** where `ffplay` is not.
    - A choice is read where a preview begins, so a film already playing or a document already drawn keeps the engine it was built under; a change reaches the next hover, or a pin's next file.
  - **Office Engine** — Which engine draws Office documents.
    - **Microsoft Office** (`default`) — Uses the format's own app and falls back to LibreOffice.
    - **LibreOffice** — Draws every Office document.
    - LibreOffice is greyed out if not installed.
  - **MPV TTL** — How long the player is kept between films: **Until away** (`default`), `30 seconds`, `1 minute`, `5 minutes`, `10 minutes`, `30 minutes`, `1 hour`, `Indefinitely`, `Never (0 seconds)`.
    - The setting is stored and marked today; keeping a player warm between films is not wired to it yet.
    - Greyed out where `mpv` is not installed.
  - **Microsoft Office TTL** — How long a family's Office app stays warm: **Until away** (`default`), `30 seconds`, `1 minute`, `5 minutes`, `10 minutes`, `30 minutes`, `1 hour`, `Indefinitely`, `Never (0 seconds)`.
    - **Until away** is kept while Explorer is reachable, then let go by **Away Timer**; a time is that long after the family's last page.
  - **LibreOffice TTL** — Same for the engine that draws CorelDRAW and nearby formats.
    - A kept engine converts the next document faster: `1.2 s` cold vs `0.2 s`, but uses a few hundred MB.
    - `Never (0 seconds)` means one engine per document.
    - Greyed out if LibreOffice is not installed.
  - **WebView2 TTL** — How long the SVG browser stays warm: the same nine answers, **Until away** at the top.
    - Greyed out if WebView2 is missing.
  - No `ImageMagick TTL`, `PeaZip TTL`, or `Calibre TTL`: those tools run once and exit, so idle time cannot bound them. A second hover is a cache hit.
- **Codecs** — What this machine has: Videos, Audio, Images, Engines.
  - A missing one carries a cross, and where the README names a page for it, picking the row offers to open that page — nothing is installed or downloaded by the app itself.
  - **Videos → MPV** opens a submenu of every mpv this machine has where there is more than one — the standalone install and SVP's own, say — each named by its path, with the one in use marked. Picking one is remembered as `mpv_path`.
- **System (vX.Y.Z)**
  - **Check for Updates** — Whether the app asks GitHub for a newer release on its own, at startup and when the menu is opened. On by default; at most once every six hours. Where one is found, the **Update Available** row appears above this submenu on the menu's next opening; where none is, nothing happens. Turning it on asks at once.
  - **Run at Startup** — Add or remove the Windows startup entry.
    - On every start, an entry that names another copy of the app — a portable copy, an older version, a path that has moved — is pointed back at the one you are running.
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in `%APPDATA%\rust-hover-preview\theme`. The look is the app's own rather than a text preview's, so it sits beside the rows about the app.
  - **Edit Config.ini** — Open the configuration file in its default editor.
- **Exit** — Close the app.

## Configuration

Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

The file is watched, and changes apply without a restart. A read takes what the file says and puts right in the app what it cannot read — the file itself is left alone until the next start, so a hand edit is never written over while it is being made. A save writes the whole file beside it, flushed to the disk, and moves it into place, so a save cut short leaves the file that was there; a file holding no setting of its own is not read as a configuration, and is left exactly where it is.

Example, trimmed:

```ini
[settings]
; General
check_for_updates=true
pin_always_on_top=true
pin_double_click_center=true
pin_double_click_reset_size=true
pin_enabled=true
pin_key=mouse3,space
pin_nav_audio=category
pin_nav_others=all
pin_nav_video=category
pin_update_enabled=true
pin_update_on_hover=false
preview_enabled=true
run_at_startup=true
theme=light
toggle_key=f7

; Minimize
pin_minimize_to=bubble
pin_pause_audio=true
pin_pause_video=true

; Preview Types
archive_preview_enabled=true
audio_preview_enabled=true
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
render_html=false
text_font_scale=125

; Timing
hover_delay_ms=0
prioritize_keyboard=true
same_file_rehover_delay_ms=200
settling_delay_ms=0
trigger_key=alt
trigger_key_affect_pin_mode=false
trigger_key_enabled=true
trigger_key_mode=disable

; Placement
avoid_mode=filename
follow_cursor=false

; Scaling
animated_scale=100
audio_scale=10
design_scale=fit
document_scale=fit
ebook_scale=fit
font_scale=50
preview_scale=100
text_scale=fit
vector_scale=fit
video_scale=100

; Background
dds_background=white
design_background=checkerboard
font_background=white
html_background=white
image_background=checkerboard
vector_background=checkerboard

; Volume
audio_loop=true
audio_seek=remember
audio_volume=10
normalize_video_volume=false
normalize_volume=false
pin_mode_audio_end_of_file=loop
pin_mode_audio_seek=start
pin_mode_audio_shuffle=false
pin_mode_image_shuffle=false
pin_mode_video_end_of_file=loop
pin_mode_video_shuffle=false
remember_audio_volume=true
remember_video_volume=true
video_loop=true
video_volume=0

; Performance
decode_budget_gb=1
document_cache_mb=256
general_disk_cache_mb=128
image_cache_mb=128
image_disk_cache_mb=512
tick_ms=15
video_hw_accel=true

; Engine
afk_timer_seconds=60
libreoffice_idle=600
libreoffice_persistent=false
office_engine=microsoft_office
office_engine_idle=600
office_engine_persistent=false
video_engine=best
video_engine_fallback=true
webview_idle=600
webview_persistent=false

; Advanced
hdr_exposure=0
hdr_tone_map=reinhard
spinner_delay_ms=250
```

## Here's what this .INI does:

### Look and text

- `theme`: `light`, `dark`, or `custom:<name>`.
- `markdown_mode`: show Markdown as `rendered` or `source`.
- `text_font_scale`: 1–1000%, default `125`; archive lists follow it too.
- `render_html`: `false` by default.
  - On, `.htm` and `.html` files are run as the pages they hold by the browser engine, at the share of the screen **Document Scaling** names.
  - That is the one kind of preview a pointer and a keyboard reach: the page's own box holds the pointer, and a click into it gives the keys.
  - Without that engine, they are text like before.

### What can preview

- `*_preview_enabled`: turn each preview type on or off. File lists stay normal.
- `extensions` under each kind's section: which files that kind previews.
  - Kinds: `image`, `video`, `ffmpeg`, `audio`, `text` (which also carries a `names` key for files with no extension), `archive`, `office`, `font`, `design`, `vector`, `ebook`, `libre`, `magick`, `peazip`, `calibre`.
  - No dots.
  - An entry with a dot, like `tar.gz`, matches the end of the file name.
  - A name two sections hold belongs to the earlier kind, so the kind decides which switch applies to it.
- Which engine plays a sound is the machine's answer rather than a setting: the media engine Windows has is asked first and an installed FFmpeg second.

### Video and audio

- `video_loop` / `audio_loop`: whether a hovered film or sound goes round at the end of its file. Both default `true`; off, it stops at the end.
- `pin_mode_video_end_of_file` / `pin_mode_audio_end_of_file`: what a *pinned* film or sound does at the end of its file — `loop` (default), `pause` (the file left where it ended, the card holding it with a play button), or `continue` (the walk's next file).
- `pin_mode_video_shuffle` / `pin_mode_audio_shuffle`: whether a forward step of a *pinned* walk takes a random file of the anchor's kind rather than the one beside it. Both default `false`; a step back is the file the walk came from.
- `pin_mode_image_shuffle`: the same question asked of a *pinned* picture's walk, which has no end of file to jump from — its shuffle is the caption's buttons and the keyboard's alone. Default `false`.
- `pin_mode_audio_seek`: where in a file a *pinned* sound starts playing — `start` (default), `remember`, `middle`, or `random`.
- `beat_hz`: the band a sound's card's mark beats with — `off`, or a top of band in hertz between `20` (the fixed floor) and `16000`. Default `250`. A setting of `config.ini` alone; the tray has no row for it.
- `audio_mark_offset_y`: a vertical shift in pixels for the mark beside the name — `0` default, `-100` to `100`. A file naming it `audio_mark_vertical_offset`, `audio_mark_y_offset` or `audio_mark_vertical_alignment` is read the same way.
- `audio_seek`: where in a file a hovered sound starts playing.
  - `remember` (default) picks it up where the last hover left it, `start` always begins at the beginning, `middle` drops it half way in, and `random` anywhere at all.
  - The remembered positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so they survive a restart; nothing is remembered while another mode is chosen.
  - A video is always played from its beginning.
- A video's engine is the `video_engine` setting; the two lists decide which engine is *asked* about a name — the media engine takes only `[video]` names, and an `[ffmpeg]` name is FFmpeg's alone.

### Memory and cache

- `image_cache_mb`: memory for decoded image frames — one pool shared with pinned windows, so a hover sweep does not evict the frame a pin is drawn from. Default `128`, max `16384`; `0` holds nothing.
- `document_cache_mb`: disk cache for rendered document pages. Default `256`, max `131072`; `0` keeps nothing between hovers but still draws the current one. Pages live under `%TEMP%\rust-hover-preview\document`; the least recently read page is removed first.
- `image_disk_cache_mb`: disk cache for the pictures ImageMagick developed. Default `512`, max `131072`; `0` keeps nothing between hovers but still develops the current one. Pictures live under `%TEMP%\rust-hover-preview\image`; the least recently read one is removed first, and they survive a restart.
- `general_disk_cache_mb`: disk cache for the subtitle tracks a film's embedded ones are copied into. Default `128`, max `131072`; `0` extracts nothing, and a film whose subtitles are embedded is drawn without them. Files live under `%TEMP%\rust-hover-preview\general`, least recently used first.
- `decode_budget_gb`: most memory one hover may use. Default `1`, range `0.25`–`64`. A file over the limit shows no preview.

### Performance

- `video_hw_accel`: whether a film mpv plays is decoded on the graphics card — the tray's **Performance → Hardware Acceleration (MPV Only)**.
  - Default `true`.
  - It is mpv's switch alone: what it reaches is mpv's own `--hwdec=auto`, and FFmpeg's player is given no decode device at all. A machine with no `mpv.exe` reads it as off, whatever the file says.

### HDR

- `hdr_tone_map`: how HDR/EXR light becomes screen values: `reinhard` default, `aces`, `srgb`, or `off`. PNG/JPEG are not affected.
- `hdr_exposure`: stops shifted before that curve. Default `0`, range `-10` to `10`.

### Waiting and engines

- `spinner_delay_ms`: wait before showing the loading spinner. Default `250`; `0` shows it immediately.
  - One delay covers all preview types, and a pinned window being shown another file with it.
  - A pin that has to wait shows a spinner of its own in the middle of itself, drawn over the file it is still showing rather than in place of it, and the window is never hidden while it waits.
  - A hover whose wait ends with nothing to show ends as the cross a pinned window shows for the same failure, in the box the wait had, and comes down three seconds later.
- `office_engine`: `microsoft_office` default, falls back to LibreOffice when needed; or `libreoffice`, which always uses LibreOffice. If LibreOffice is missing, it falls back and the tray row is greyed out. Read where a preview begins, the way `video_engine` is.
- `video_engine`: which engine plays a video: `best` (default), `mpv`, `ffmpeg`, `native`, or `hybrid`.
  - **Best** is the machine's own answer — mpv where it is installed, and otherwise `hybrid` where FFmpeg is installed (the media engine draws a film at or under 3.2 megapixels, FFmpeg's player takes a larger one or one it could not measure) and `native` where neither is.
  - `mpv` names mpv alone, `native` names the media engine alone, `ffmpeg` names `ffplay` alone, and `hybrid` states that size rule plainly.
  - A choice this machine cannot supply is read as `best`.
  - Read where a preview begins, so a film already playing keeps the player it was begun under; a change reaches the next hover, or a pin's next file.
- `mpv_path`: which `mpv.exe` plays, written by **Codecs → Videos → MPV** where the machine has more than one. Empty, the default, means the first one found — beside the app, then the `PATH`, then `MPV Player` under Program Files, then SVP's `mpv64`. The only setting that holds a path.
- `video_engine_fallback`: `true` (default) lets a chosen engine that cannot play a given file fall through to the engines it may — a chosen `mpv` to both of the others, `native` and `ffmpeg` to each other — and `false` makes the engine you chose stand alone. Read only where `video_engine` names an engine; **Best** is its own walk either way.
- `libreoffice_idle` / `office_engine_idle`: seconds the engine is kept after its last page, or `indefinitely`. Default `600`; `0` launches it per document.
- `afk_timer_seconds`: seconds with no Explorer window before a non-persistent engine is released. Default `60`, max one day. `0` releases immediately.
- `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent`: `true` keeps that engine always, bounded only by its `…_idle` time. `false` default keeps it while Explorer is reachable, bounded by `afk_timer_seconds`.

### Trigger and position

- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled`: the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched. Default `true`.
- `toggle_key`: the key that turns previews on or off while File Explorer is the foreground window. Default `f7`; it takes the same spellings and alternatives as `pin_key` (`ctrl+f7`, `f7,f8`, …).
  - The key is not a global hotkey: it is only acted on when Explorer is in front, and it is never swallowed, so it still reaches Explorer and other programs as usual.
  - Holding it down is one toggle, not one per keyboard repeat. The tray's **Enable Preview** row names the configured key.
- `trigger_key_affect_pin_mode`: whether the trigger key reaches a pinned preview. Default `false` — the key is not read while a preview is pinned, up or collapsed into its bubble, since a pin is a window you put there rather than a hover for the key to hold back. `true` lets holding the key bring the pin down with the previews it stops. It speaks for `disable`; the `enable` mode is left as it is.
- `pin_key` / `pin_enabled`: the key that pins the preview on screen (`mouse3,space` by default — the same spellings the trigger key takes: `alt`, `f8`, `a`, `space`, …), and whether it is watched. See **Pin Mode** above for what a pin does.
  - A binding may name alternatives and chords: `,` separates alternatives (either one matches) and `+` joins keys that must all be held. The last key of a chord is the one whose press counts.
    - `pin_key=space,page_down` — either key pins.
    - `pin_key=ctrl+space` — Space pins while Ctrl is held; `ctrl+alt+space` needs both.
    - `pin_key=space,mouse3` — a mouse button can pin too (`mouse3`, `mouse4`, `mouse5`; `mouse1` and `mouse2` are not allowed).
  - The same syntax applies to `trigger_key`, where an alternative counts as held only while all of its keys are.
  - Names are case-insensitive, spaces around `,` and `+` are ignored, and underscore spellings such as `page_down`, `page_up` and `caps_lock` work.
- `pin_always_on_top`: whether a pin stays above every other window. `true` (default) keeps it topmost; `false` makes it an ordinary window that other programs can cover and that wears this app's own taskbar button while it is up. The caption's own mark toggles the same setting, a film mpv plays is put in the same band, and the wheel over such a pin is the pin's only while the pin is the window you are in.
- `pin_double_click_center` / `pin_double_click_reset_size`: what a double-click on a pinned window does. Both default `true`: the window is put back to the size its kind's scaling setting asks for, and in the middle of the screen it stands on. Off, either half is left out — with the size reset off, a hand's own size stands and the double-click is the move alone.
- `pin_update_enabled`: whether a pin that is up is shown the file picked next, by click or by keyboard, without moving the window. Default `true`.
  - Off, the pin keeps the file it was taken up on until it is closed — a pin is also a window to read in, and one that swapped its file out from under the hand on every keystroke would be unreadable.
  - The keyboard half follows Explorer, so it is inert while a pinned window holds the keyboard — click back into the listing and it resumes.
- `pin_update_on_hover`: whether that following includes the pointer's own hover, or only what a click or a key asks for. Default `false`; greyed in the tray while `pin_update_enabled` is off.
- `pin_nav_video` / `pin_nav_audio` / `pin_nav_others`: what the caption's **Previous** and **Next** buttons step through, by the walk's anchor — the first file pinned. Video and Audio default to `category`, Others to `all`.
  - `all` is every file this build can preview, so a video steps to the sound beside it.
  - `category` narrows the walk to the anchor's category — video, audio, or the anchor's own kind for Others, where a camera raw is a picture and a book is a document.
  - The walk is the folder the pin was taken up in and no subfolder of it, is in the order the Explorer listing is showing, wraps at both ends, and is held per folder, read on the first press rather than when the pin is taken up.
  - A file it reaches that the window cannot show is stepped over, and the walk is bounded by the folder: each other file is offered at most once, so a folder of nothing this build can read ends the walk instead of going round for ever.
  - A file holding only the old `pin_nav_file_types` seeds all three with it; a file written before any of them existed walks as the fresh defaults do.
- `pin_zoom_bounce_ms` / `pin_zoom_bounce_gate` / `pin_zoom_bounce_speed_ms` / `pin_zoom_bounce_power`: the landing a wheel zoom out ends with — how long the image takes to swing back to its place (`100` default), whether the swing past its place is on (`true` default), how fast it snaps back (`0` default), and how far past it drifts as a share of the offset it is coming home from (`0.04` default). A hand-edited value past the ceiling is reduced to it.
- `pin_minimize_to`: where a pin put away by the window's own **Minimize** button goes. `bubble` (default) leaves the round bubble on the desktop and keeps this app out of the taskbar; `taskbar` takes the window off the screen and shows this app's own taskbar button — named for the pinned file and wearing this app's icon, its hover holding the pin with **Previous**, **Play/Pause** and **Next** for a film or a sound — which brings the pin back where it stood when clicked and closes it on a right-click. Read when the button is pressed, so a change reaches the next minimize; a file written before this setting existed is read as `bubble`.
- `pin_pause_video` / `pin_pause_audio`: whether a pin put away into its bubble or the taskbar holds what it was playing where it stands, and starts it again at the second it stopped at when the window comes back up. Both default `true`. They are two switches because a video and a sound are two different things to want quiet.
- `follow_cursor`: `true` = Follow Cursor; `false` = Best Position.
- `avoid_mode`: `filename` default, `filename_column`, `details`, or `off` — what the preview avoids.

### Scaling

- `preview_scale`: `fit`, a share of the display's fitted size (`screen 75`, `screen 50`, `screen 25` or `screen 10` — the display fitted to the media, reduced to that share of it), or a share of the picture's own size (`25`–`400`). Default `100`.
- `video_scale`: same for video. Default `100`.
- `audio_scale`: the share of the display a sound's card is laid out over — `25`, `20`, `15`, `10`, `7` or `5`. Default `10`. The whole card scales with the share (font, height, width), so it's the `10%` card uniformly smaller or larger, and `Text Size` no longer resizes it. A hand-edited value is sanitized rather than fatal: `100` or more reads as `100`, `0` reads as `10`, an unrecognized value is ignored.
- `animated_scale`: same for animated GIF, WebP, or PNG — and for an animated JPEG XL, which is played as a picture rather than as a video. Default `100`. Still GIF/PNG follow `preview_scale`.
- `vector_scale`: percentage or `fit`, based on the screen. Default `fit`; `100` or more reads as `fit`. Covers all Vector drawings. Old `svg_scale` is ignored and removed.
- `text_scale`: how much of the screen a text preview box may take when it opens, for plain text, code, and Markdown. Default `fit`; `100` or more reads as `fit`. `Text Size` still scales the text inside that box.
- `ebook_scale`: PDF page scale against the screen. Default `fit`; `100` or more reads as `fit`.
- `document_scale`: document page scale against the screen. Default `fit`. A workbook's fallback bitmap keeps its own size and is never enlarged.
- `font_scale`: font preview scale against the screen. Default `50`; `fit` means all of it; `100` or more reads as `fit`.
- `design_scale`: design document scale against the screen. Default `fit`.

### Fonts

- `ttc_face`: which face of a `.ttc` is drawn. `1` is first, max `10`. The heading shows which face came out.

### Backgrounds

- `image_background`: `checkerboard` default, `black`, `white`, or `transparent`. Also used for PDF pages, painted text frames, and document pages. Old `transparent_background` is ignored and removed.
- `font_background`: `white` default, `black`, `checkerboard`, or `transparent`.
- `dds_background`: `white` default or `black`, and only those two. Any other value reads as `white`.
- `vector_background`: `checkerboard` default, `white`, `black`, or `transparent`. Used for SVG pages and metafiles. Old `svg_background` is ignored and removed.
- `html_background`: `white` default, `black`, or `checkerboard`, and only those three — a page is drawn on a page, so transparency is not one of the backdrops it is offered. Used for a page of HTML previewed by `render_html`, run or not; any other value reads as `white`.
- `design_background`: `checkerboard` default, `white`, `black`, or `transparent`.
- `zoom_background`: `black` default, `white`, or `checkerboard`, and only those three — the band a hand carries a zoomed picture off its window is the window's to fill rather than a hole to look through, since a layered window is hit-tested by the shape of its pixels. Asked of a zoomed picture with no transparency in it: a picture with a transparent pixel keeps its own kind's backdrop (`image_background` — **Background → Image**), see-through included. Any other value reads as `black`.

### Old or removed names

- `magick_preview_enabled` and `peazip_preview_enabled`: not read. ImageMagick pictures use `image_preview_enabled`; PeaZip archives use `archive_preview_enabled`.
- `libre_preview_enabled` and `libre_scale`: gone. Use `document_preview_enabled` and `document_scale`.
- `office_cache_mb` and `libre_cache_mb`: gone. Use `document_cache_mb`. If both old ones exist, the larger is read once, then both are removed.
- `text_preview_full_mode`: gone. A pinned text preview comes up in full mode instead, so it scrolls and its text can be selected and copied.
- `pin_nav_file_types`: gone. Use `pin_nav_video`, `pin_nav_audio` and `pin_nav_others`; a file holding only the old key seeds all three with it.
- `video_jump_enabled` / `audio_jump_enabled`: gone. Use `video_loop` / `audio_loop`.
- `pin_mode_video_jump_enabled`, `pin_mode_audio_jump_enabled`, `pin_mode_video_jump`, `pin_mode_audio_jump` and `pin_mode_audio_loop`: gone, replaced by `pin_mode_video_end_of_file`/`pin_mode_audio_end_of_file` beside `pin_mode_video_shuffle`/`pin_mode_audio_shuffle`. A file holding one is read once into the settings that replaced it — `…_jump_enabled = false` as `pause`, `…_jump = shuffle` as `continue` with shuffle on, an old `…_loop` as `continue` — and the retired keys are dropped on the next save.
- No `magick_idle`, `peazip_idle`, or `calibre_idle`. ImageMagick, PeaZip, and Calibre are converters, not engines kept open. A second hover usually costs only a cache hit.

### Extension list behavior

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

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item, and GDI paints the preview into a topmost layered window. Text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives' own tables of contents — or from the listing an installed PeaZip produces for the formats no reader here has — Office documents are drawn from a page Office renders in the background, WebView2 draws SVG documents and font specimens, and runs HTML pages where Render HTML is switched on, camera raw and the pictures beside it are developed by ImageMagick, and video is played by mpv where installed, and otherwise by Windows' own media engine for small films and FFmpeg's player for large ones. A middle click or `Space` pins the preview into a captioned window of its own — always on top unless you turn that off, in which case it wears a taskbar button of its own.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview is local-first: previews work without an internet connection, and the only network request is an update check, which runs only at startup and when you open the tray menu, at most once every six hours, and never while `Check for Updates` is switched off. There is no telemetry, analytics, ads, accounts, or crash reporting; it reads only the item you hover or focus in Explorer, locally and only for enabled preview types, and cloud-only placeholders are skipped on purpose while password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use locally installed mpv, FFmpeg, LibreOffice, or ImageMagick when available, plus Microsoft Office, Windows' own media engine, and the Windows PDF engine; caches are bounded by `config.ini` — decoded images stay in memory, while the page an engine drew for a document is kept as a file under the temp folder, where Windows is free to clear it. See [PRIVACY.md](PRIVACY.md) for full details.

## License

MIT. See [LICENSE](LICENSE).

The app is built out of other people's code as much as its own — the Windows bindings, the image decoders, the syntax highlighter, the archive readers, the browser bindings — and each of those carries its own licence, with the notices MIT and BSD ask to be reproduced. [THIRD-PARTY.md](THIRD-PARTY.md) lists every dependency grouped by licence, with the copyright holders beside it, and the full texts are in [`LICENSES/`](LICENSES). Both are generated from the dependency tree rather than kept by hand; `generate-attribution.ps1` refreshes them.

RAR archives are handled through the official UnRAR sources compiled in by `unrar_sys`: that code is free to use for reading RAR archives, but it may not be used to build a RAR-compatible archiver or to recreate the RAR compression algorithm. See `THIRD-PARTY.md` (`LicenseRef-UnRAR`).
