# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.98.1+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar. Hover a file in File Explorer — or move with the arrow keys — and a preview appears beside it. It works with mouse and keyboard, and you configure it from the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer, beside the cursor or the focused item and kept on screen.
- Press a key and a preview becomes a window of its own — captioned, movable, resizable, always on top, where it stays until you close it.
- Supports images, design documents, camera raw, vector drawings, fonts, videos, PDFs, ebooks, comics, text/code, archives, Office documents, and more.
- A file's preview depends on what's really inside it, not its name — so renamed files still show correctly, unreadable content gets no preview, and ambiguous types are resolved by extension or content.
- Scaling from 25% to 400%, or fit-to-screen, with separate scaling for images and videos and for vector drawings, PDFs, documents, fonts, design documents, and sounds.
- Tray menu and hand-editable `config.ini`; DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.

## Supported Formats

### Images

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| This app's own decoders and the codecs Windows ships | `[image]` | apng, avci, avif, bmp, dds, exr, ff, gif, hdr, heic, heif, ico, jfif, jpe, jpeg, jpg, jxl, pam, pbm, pgm, png, pnm, ppm, qoi, tga, tif, tiff, webp |
| [ImageMagick](#optional-enable-camera-raw-and-more-pictures-imagemagick) | `[magick]` | 3fr, aai, art, arw, ase, aseprite, bayer, bayera, bgr, bgra, bgro, c2pa, cal, cals, cmyk, cmyka, cr2, cr3, crw, cube, cur, cut, dcm, dcr, dcx, dng, dpx, dxt1, dxt5, erf, fax, fff, fit, fits, fl32, fts, ftxt, g3, g4, gray, graya, group4, hrz, icb, iiq, ipl, j2c, j2k, jng, jnx, jp2, jpc, jpm, jpt, k25, kdc, mac, map, mat, mdc, mef, miff, mng, mono, mos, mpc, mrw, mtv, nef, nrw, orf, otb, pal, palm, pango, pcds, pef, pes, pfm, pgx, phm, picon, pict, pix, pwp, raf, raw, rgb, rgb565, rgba, rgbo, rgf, rla, rle, rmf, rw2, rwl, scr, sct, sf3, sfw, sgi, six, sixel, sr2, srf, srw, stegano, sun, tim, tm2, uyvy, vda, vicar, viff, vips, vst, wbinfo, wbmp, x3f, xbm, xcf, xpm, xv, ycbcr, ycbcra, yuv |

`heic`, `heif`, `avif`, `avci`, `jxl` and a still `webp` are decoded by a codec Windows provides rather than one this app carries, so each of those six stills needs its [Store extension](#optional-enable-heic-avif-jpeg-xl-and-webp-preview-windows-codecs) installed once; every other entry of `[image]` needs nothing installed.

### Videos

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| The media engine Windows 11 ships | `[video]` | 3g2, 3gp, 3gpp, asf, avi, dvr-ms, m1v, m2t, m2ts, m2v, m4v, mkv, mov, mp4, mpe, mpeg, mpg, mts, qt, ts, vob, webm, wmv |
| [FFmpeg](#optional-enable-video-preview-with-ffmpeg)'s player (`ffplay`) | `[ffmpeg]` | 264, 265, 266, apv, av1, avc, avs, avs2, avs3, bik, bk2, c93, cavs, cdg, cdxl, cin, cpk, dav, dif, divx, drc, dv, evc, f4v, flm, flv, gxf, h261, h263, h264, h265, h266, h26l, hevc, ifv, imx, ismv, ivf, ivr, kux, m2p, mj2, mjpeg, mjpg, mk3d, moflex, mpv, mve, mvi, mxf, mxg, nsv, nut, obu, ogm, ogv, pmp, psp, rcv, rm, rmvb, roq, rsd, smk, str, swf, thp, tod, tp, tr, ty, ty+, usm, vc1, vc2, viv, vro, vvc, vw, wtv, xl, xmv, y4m, yop |

### Audio

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| The media engine Windows ships, then [FFmpeg](#optional-enable-video-preview-with-ffmpeg)'s player — the machine's answer, not a setting | `[audio]` | aac, ac3, aif, aifc, aiff, amr, ape, au, awb, caf, dff, dsf, dts, dtshd, eac3, flac, m4a, m4b, mka, mp2, mp3, mpa, mpc, oga, ogg, ofr, ofs, opus, ra, shn, snd, spx, tak, tta, voc, wav, wave, wma, wv |

### Text

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| This app, syntax-highlighted | `[text]` `extensions` | adb, adoc, ads, asciidoc, asm, asp, aspx, astro, awk, bash, bat, bib, bzl, c, cc, cfg, cg, cjs, clj, cljc, cljs, cmake, cmd, comp, conf, cpp, cs, csh, cshtml, css, csv, csx, cts, cxx, d, dart, diff, diz, edn, ejs, el, elm, env, erb, erl, ex, exs, f, f03, f77, f90, f95, fish, for, frag, fs, fsi, fsx, ftn, fx, geom, glsl, go, gql, gradle, graphql, groovy, h, haml, hbs, hcl, hh, hlsl, hpp, hrl, hs, htm, html, hxx, inc, ini, ipynb, java, jl, js, json, json5, jsonc, jsonl, jsp, jsx, ksh, kt, kts, latex, less, lhs, liquid, lisp, ll, lock, log, lsp, lua, m, mak, man, markdown, md, mdown, metal, mjs, mk, mkd, ml, mli, mm, mts, mustache, nasm, nfo, nim, ninja, nix, njk, org, pas, patch, php, phtml, pl, plist, pm, properties, proto, ps1, psd1, psm1, py, pyi, pyw, r, rake, rb, rkt, rmd, rs, rst, rtf, s, sass, scala, scm, scss, sh, slim, sol, sql, srt, ss, styl, sv, svelte, svh, swift, tcl, tex, text, tf, tfvars, toml, ts, tsv, tsx, twig, txt, v, vbs, vert, vhd, vhdl, vtt, vue, wat, wgsl, xhtml, xml, xsd, xsl, xslt, yaml, yml, zig, zsh |
| This app, as plain text | `[text]` `names` | authors, .babelrc, brewfile, caddyfile, changelog, changes, .clang-format, .clang-tidy, cmakelists.txt, code_of_conduct, containerfile, contributing, contributors, copying, copyright, dockerfile, .dockerignore, .editorconfig, .env, .env.example, .env.local, .eslintignore, .eslintrc, gemfile, .gitattributes, .gitconfig, .gitignore, .gitkeep, .gitmodules, gnumakefile, .golangci.yml, history, .htaccess, install, jenkinsfile, justfile, licence, license, .mailmap, makefile, makefile.am, makefile.in, notice, .npmignore, .prettierignore, .prettierrc, procfile, rakefile, readme, .rustfmt.toml, security, .stylelintrc, unlicense, vagrantfile |

### Ebook

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| The Windows PDF engine for the PDF's three spellings; this app's own readers for the three comic containers | `[ebook]` | cbc, cbr, cbz, epdf, pdf, pdfa |
| [Calibre](#optional-enable-ebooks-calibre), which converts the book to a PDF first | `[calibre]` | azw, azw3, azw4, djvu, epub, fb2, htmlz, lit, lrf, mobi, pml, prc, snb, tcr |

### Archives

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| This app, from each archive's own table of contents | `[archive]` | 7z, apk, jar, rar, tar, tar.gz, tgz, xpi, zip, zipx |
| [PeaZip](#optional-enable-niche-archives-peazip) | `[peazip]` | 001, apfs, ar, arc, arj, bcm, br, bz2, bzip2, cab, chm, cpio, cramfs, deb, dmg, esd, gz, gzip, hfs, hfsx, hxs, iso, lha, lpaq8, lzh, lzma, msi, msp, pkg, ppkg, qcow, qcow2, rpm, squashfs, swm, taz, tbz, tbz2, tpz, txz, tzst, udf, udeb, vdi, vhd, vhdx, vmdk, wim, xar, xip, xz, z, zpaq, zst |

### Document

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| Microsoft Office, falling back to [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice) where Office is absent | `[office]` | doc, docm, docx, dot, dotm, dotx, pot, potm, potx, pps, ppsm, ppsx, ppt, pptm, pptx, xls, xlsb, xlsm, xlsx, xlt, xltm, xltx |
| [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice) | `[libre]` | 123, 602, abw, cdr, cgm, cmx, cwk, dbf, dif, dxf, fodg, fodp, fodt, gnm, gnumeric, hwp, key, lwp, mcw, met, mw, numbers, odb, odc, odf, odg, odm, odp, ods, odt, oth, otg, otm, otp, ots, ott, pages, pcd, pct, pcx, pdb, pm6, pmd, psw, pub, ras, sda, sdc, sdd, sdw, slk, stc, std, sti, stw, svm, sxd, sxg, sxi, sxm, sxw, vdx, vsd, vsdm, vsdx, vstx, wb2, wk1, wk3, wk4, wks, wpg, wq1, wq2, wpd, wps, wri, xlw, zabw, zmf |

### Vector

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| WebView2 for `svg` and `svgz`; the drawing layer for the metafiles and the preview picture an EPS carries | `[vector]` | emf, epi, eps, epsf, epsi, ept, ept2, ept3, svg, svgz, wmf |

### Fonts

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| WebView2 — a specimen page with the font in it | `[font]` | otf, ttc, ttf, woff, woff2 |

### Design

| Engine | `config.ini` | Default extensions |
| --- | --- | --- |
| This app, from the flattened picture each format keeps of the whole document | `[design]` | ai, fig, kra, ora, procreate, psb, psd, sketch, xd |

## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

1. Open [Releases](../../releases).
2. Download your preferred asset.
3. Run the installer, or place the portable binary wherever you like.
4. Launch Rust Hover Preview.

No Rust toolchain is needed.

### Optional: Enable Video Preview with FFmpeg

FFmpeg adds the video formats and the audio ones Windows does not decode — the `[ffmpeg]` list, and the `[audio]` names Windows' own decoders do not reach. It also plays any `[video]` name the media engine turns down, or every video there is, where **Engine → Select Engine → Video** names FFmpeg. `ffplay` and `ffprobe` need to be in your `PATH`.

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
| -------------------------------------- | ------------------------------------------------------------------------ |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)  |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)   |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)    |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4) |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)   |

All are free, official from Microsoft. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray's **Codecs** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

A **still** `heic`, `heif`, `avif`, `jxl`, and still `webp` is decoded by a codec Windows provides rather than one shipped with the app. Each still needs its extension installed once from the Microsoft Store:

| Still format | Needs                                                                                                                                           |
| ------------ | ----------------------------------------------------------------------------------------------------------------------------------------------- |
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

Or download the official installer from [https://calibre-ebook.com/download\_windows](https://calibre-ebook.com/download_windows) and run the suggested file (`calibre-*-64bit.msi`).

> [!WARNING]
> If `winget` reports an error, the package sources are usually why: run `winget source reset --force`.
>
> If that is refused as well, open **Terminal as administrator** (right-click the Start button → *Terminal (Admin)*) and run the same command from there.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys, Tab, or a file's name typed to jump to it — to preview the focused item.
4. Right-click the tray icon to configure behavior.

## System Tray Menu

A setting marked `(Default)` is what an untouched setting would be. The check or radio mark shows what is set now.

- **Enable Preview** — Turn previews on or off.
- **Pin Mode** — Everything about pinning, in one place.
  - **Enable (Space)** — Turn pinning on or off.
    - Pressing the key while a preview is up turns it into a window of its own: captioned, movable, resizable, always on top, and still there when the pointer leaves.
    - The caption carries **Previous**, **Next**, **Open With**, and **Open With...**.
      - **Previous** / **Next** — Step the pin through the sibling files in the folder it was taken up in, in the order the listing is showing them; step over any one that will not open rather than stopping there.
      - **Open With** — Opens the file in whatever program Windows has registered for it.
      - **Open With...** — Asks you to pick one from the list Windows keeps of the programs that could.
      - The two hand-off buttons name themselves while the pointer rests on them; **Open With** names the program it would use.
    - Its edges resize it; what that does depends on what is inside.
      - A picture, video, or rendered page keeps its own shape.
      - A document or listing is laid out to whatever box it is given.
      - A sound's card keeps the size **Audio Scaling** names, and a video `ffplay` has is not resized at all.
    - A picture's caption and bar are drawn over the media rather than in bands around it, and fade in as the pointer arrives; only the strip under the pointer is shown.
    - A pinned video carries a transport bar of its own: play/pause, a draggable seek bar, and the two clocks.
      - Where FFmpeg's player is the one playing it, a read-out bar appears instead, since that player can be told nothing.
    - A pinned sound's card carries four buttons of its own: **Previous**, **Play/Pause**, **Next**, and **Volume**.
      - `Space` holds a sound where it stood and sets it going again from there.
      - A click on a sound's bar takes the file to that second.
    - A pinned text preview comes up in full mode, so it scrolls and its text can be selected and copied.
      - `Ctrl+A` selects all of it; `Ctrl+C` copies what is selected.
    - A pinned video has a transport bar with a volume of its own, set on the window; it does not change **Volume → Video**.
    - **Maximize** fits the window to the screen.
    - **Minimize** collapses it into a round bubble that can be dragged anywhere and clicked to bring the window back; a right-click on the bubble closes it.
    - Clicking the window gives it the keyboard; then the arrow keys step it and `Escape` closes it.
    - The pin key itself only ever puts a pin up and brings back a bubble; it never hides one.
    - The key is named in the item and can be changed in `config.ini` (`pin_key`).
  - **Update Preview** — Whether a pin that is up is shown the file you pick next.
    - **Enabled** — On by default.
      - A file you click, or one the keyboard selects, is shown where the window already stands, rather than as a second preview beside it.
      - A file named this way that cannot be shown leaves the window on the file it already had.
      - Stepping over a file that will not open belongs to the caption's own **Previous** and **Next**, which you pressed.
      - A sound's card is drawn at the size its **Audio Scaling** setting names, in the middle of that box; it follows the **Theme** setting, and no longer **Font Size**.
      - The keyboard half follows Explorer, so it pauses while a pinned window holds the keyboard; click back into the listing and it carries on.
    - **On Hover** — Off by default.
      - On, the pointer's own hover is one of the ways a pin is told about a file, the way a Quick Look window follows a listing.
      - Greyed while **Enabled** is off.
  - **Pause Preview** — What a window collapsed into its bubble does with what it was playing.
    - Both on by default: the media engine is paused where it stood and started again at the second it stopped at when the window comes back, rather than playing on behind a bubble nobody can see.
    - **Audio**
    - **Video**
  - **Nav File Types** — What **Previous** and **Next** on the caption step through.
    - Both work with **Update Preview** off; they are a thing you pressed, and no setting asks for them.
    - The arrow keys step the same way, and are asked for by nothing at all — but only once you have clicked the pinned window, which gives it the keyboard. Until you do, the arrows belong to whatever is in front.
    - `Space` in a pinned sound pauses and resumes it, on the same terms.
    - `Escape` in a pinned window closes it.
    - **All** (`default`) — Every file this build can preview, in the folder the pin was taken up in and not any subfolder of it, in the order the Explorer listing is showing them. A video sits beside a sound.
    - **Category** — Only the files of the pinned file's own kind of thing: pictures, video, audio, documents, archives, text, fonts, or design.
      - A camera raw is a picture and a book is a document, because that is what you call them.
      - A kind switched off under **Preview Types**, or one with no engine installed on the machine, is not a step either way.
      - A file this build *could* preview but that will not open is a step the window steps over rather than one it stops on.
- **Preview Types** — Choose which file kinds can preview: Images, Videos, Audio, Text, Ebook, Archives, Document, Vector, Fonts, Design.
  - A switch is a switch over behaviour, not over files: the lists deciding which files are previewed are untouched, so switching a kind off and back on restores what was configured.
  - One switch covers both the original file and the engine-drawn preview.
    - Camera raw uses **Images**.
    - A LibreOffice document uses **Document**.
    - A PeaZip archive uses **Archives**.
    - A Calibre book uses **Ebook**, as do the app-drawn PDF and comic pages.
- **Text Preview**
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in `%APPDATA%\rust-hover-preview\theme`.
  - **Font Size** — `400%` at the top down to `70%` at the bottom.
  - **Markdown** — Rendered or Source.
  - **Render HTML** — Off by default.
    - On, a `.htm` or `.html` file is previewed as the page it holds rather than as its markup, run by the browser engine.
    - A page is the one thing previewed as a page rather than as a picture, so it is the one thing that runs, and a page that draws itself with script has nothing to show without a run.
    - What it may still do is bounded: the frame it is shown in cannot open forms or popups or navigate, and nothing a page links to is fetched.
    - SVG documents and font specimens, which a browser is handed as an image, are still drawn rather than run.
    - A running page can be pointed at, clicked into and typed into; a document, a specimen and every other preview still cannot.
    - Without the engine on the machine it stays a text preview.
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
  - **Delay** — How long the pointer rests before a preview opens: `0 ms` at the top down to `1000 ms` at the bottom. Default `0 ms`.
  - **Rehover Delay** — Wait before the same file can preview again. Default `200 ms`.
  - **Settling Delay** — How still the pointer must be before previewing what it is on.
    - Default `0 ms` means a new file can preview while the hand is still moving.
    - Keyboard previews ignore this.
- **Placement**
  - **Position** — Follow Cursor or Best Position. Default **Best Position**.
  - **Avoid** — Avoid Nothing, Avoid Filename (`default`), Avoid Filename Column, or Avoid Details.
    - Keeps a preview off the item it is about.
    - Keyboard previews have no cursor, so **Avoid Nothing** acts like **Avoid Filename**.
- **Scaling**
  - **Image Scaling** — Fit to Screen or `25%`–`400%` of the image's own size.
  - **Video Scaling** — Same shares for a video. Default `100%`.
  - **Audio Scaling** — `25%`, `20%`, `15%`, `10%` (`default`), or `5%` of the display, for a sound's card — the whole card scales with the share: the font it is set in, its height, and its width together, the way Windows scaling sizes a window.
    - The card at any share is the `10%` card scaled by the share's fraction of `10%` — ~188×50 at `5%` to ~936×250 at `25%` on a 3440×1440 display — its seek bar stretching with the width and its height what its content measures at the share's font.
    - The card is built at the share's own font (the default text size, `125%`, at the `10%` anchor), so **Font Size** does not resize it; a pinned sound's card is the same size as the hover's.
  - **Animated Scaling** — Same shares for an animated GIF, WebP, or PNG — and for an animated JPEG XL, which plays through the same machinery.
    - A still GIF or PNG uses **Image Scaling**.
    - Default `100%`.
  - **Vector Scaling** — Fit to Screen (`default`), or `75%`, `50%`, `25%`, `10%` of the display.
  - **Text Scaling** — Same display shares for a text preview — a plain text file, code, or Markdown.
    - It caps how much of the screen the preview box may take when it opens; a short file still gets the small box its own text needs.
    - Default **Fit to Screen**.
  - **Ebook Scaling** — Fit to Screen (`default`), or the same display shares for a PDF page, a comic's first page, and a Calibre-converted book.
  - **Document Scaling** — Same display shares for a document drawn as a page, whether by its own Office app or by LibreOffice.
    - A workbook's fallback bitmap is never enlarged.
  - **Font Scaling** — Same shares for a font specimen. Default `50%`.
  - **Design Scaling** — Same display shares for design documents: Photoshop, Illustrator, Krita, OpenRaster, Procreate. Default **Fit to Screen**.
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard. Default **Checkerboard**.
  - **Vector Background** — Same backdrops for an SVG document or metafile. Default **Checkerboard**.
  - **HTML Background** — White, Black, or Checkerboard for a page of HTML. Default **White**.
    - The see-through backdrop is not offered: a page is drawn on a page, whether it runs or not, and the backdrop is behind the page rather than behind what it paints.
  - **Font Background** — Same backdrops for a font specimen. Default **White**.
  - **DDS Background** — Black or White for `.dds` textures. Default **White**.
    - The two see-through backdrops are not offered.
  - **Design Background** — Same as picture backdrops for a design document. Default **Checkerboard**.
- **Volume** — **Video** and **Audio**, each offering `100%`, `80%`, `65%`, `50%`, `35%`, `20%`, `10%`, `5%`, `1%`, and `0%`, loudest first.
  - A video's soundtrack starts at `0%` — silent, so a hover never makes a sound the pointer did not ask for — and a sound file at `10%`.
  - A video is looked at and a song is listened to, so the two are settings of their own.
  - A sound at `0%` still shows its card, silently.
  - **Normalize** — The first row of each half, above the levels.
    - A file's integrated loudness is measured (ITU-R BS.1770 LUFS, by FFmpeg's `ebur128`) and brought to `-14 LUFS` before it plays, so a folder of sounds — or a set of films — is heard at one level rather than at each file's own.
    - A file is not lifted past its own true peak, so a quiet one is corrected as far as its headroom allows rather than clipped.
    - On by default for **Audio** and off for **Video**, whose soundtrack is heard beside a picture that was asked for and whose measurement is a decode of the film.
    - Both are greyed out unless FFmpeg is installed — FFmpeg is what measures the loudness and what applies the gain.
    - The loudness is measured once per file and kept, so only a file's first hover waits for it.
  - **Remember** — The row under **Normalize**, above the levels.
    - Whether a level moved with a pinned window's own volume knob is the level the next preview is played at.
    - Off by default for both halves, so a knob belongs to the window it was turned on and the level in the list above it is what every preview starts from.
    - With it on, letting go of the knob writes the level to `config.ini` and the next hover — or the next film — is played at it; nothing is rebuilt on screen and a preview already playing is left where it is.
    - A pinned window also keeps that level across the other kind: a sound turned up to 100% and then stepped onto a film and back is played at 100% again, where a knob left on a film is kept for films rather than for the sounds either side of it.
  - **Audio Seek** — Where in a file a hovered sound starts playing.
    - **Remember** (`default`) — Picks it up where the last hover left it.
    - **From the Start** — Always begins at the beginning.
    - **From the Middle** — Drops it half way in.
    - **Random** — Anywhere at all.
    - The remembered positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so they survive a restart; nothing is remembered while another mode is chosen.
    - Whatever a sound is started at, it goes back to the beginning of the file when it reaches the end of it and loops from there for as long as the hover lasts.
    - A video is always played from its beginning.
- **Performance**
  - **Cache** — What a preview may cost between hovers: `2 GB` down to `0 MB`.
    - **`Image (RAM)`** = decoded frames kept in memory.
    - **`Document (Disk)`** = engine-drawn pages kept as temp files.
    - **`Image (Disk)`** = the pictures ImageMagick developed, kept as temp files.
    - **`General (Disk)`** = the subtitle tracks a film's embedded ones are copied into, kept as temp files; `0` draws a film without its subtitles.
  - **Decode Budget** — `16 GB` down to `512 MB`; default `1 GB`. A file past it gets no preview.
  - **Tick** — How often the app checks Explorer while a folder window is focused: `15 ms` (`default`), `31`, `47`, `63`, or `78 ms`.
    - Lower answers a move sooner; higher is lighter on CPU and Explorer.
  - **Hardware Acceleration** — **Video** (checked by default): a video FFmpeg's player has is decoded on the graphics card rather than on a core.
    - It is a setting about this build's FFmpeg rather than about a kind of file — a file the media engine plays is decoded by Windows either way.
    - A machine with no device FFmpeg can decode on is the one the setting is for, where FFmpeg falls back by itself.
- **Engine**
  - **AFK Timer** — How long Explorer may be unreachable before a non-**Persistent** engine is let go: `1 hour`, `30 minutes`, `10 minutes`, `5 minutes`, `1 minute` (`default`), `30 seconds`, `15 seconds`.
    - Counts time when no Explorer window is reachable on any monitor.
    - A second Explorer window keeps engines warm.
    - Each **`… TTL`** submenu has a **Persistent** toggle at the top.
  - **Select Engine → Office** — Which engine draws Office documents.
    - **Microsoft Office** (`default`) — Uses the format's own app and falls back to LibreOffice.
    - **LibreOffice** — Draws every Office document.
    - LibreOffice is greyed out if not installed.
  - **Select Engine → Video** — Which engine plays a video.
    - **Best** (`default`) — The machine's own answer: the media engine for a film at or under 3.2 megapixels, FFmpeg's player for a larger one or one it could not measure.
    - **Native** — The media engine alone.
    - **FFmpeg** — `ffplay` alone.
    - **Native (FFmpeg above 3.2MP)** — That size rule stated plainly.
    - A **Fallback** switch at the top (`default` on) lets a chosen engine that cannot play a given file fall through to the other; off, the engine you chose stands alone.
    - **FFmpeg** and **Native (FFmpeg above 3.2MP)** are greyed out where `ffplay` is not installed.
  - **Microsoft Office TTL** — With **Persistent** on: how long a family's Office app stays warm: Indefinitely, `1 hour`, `30 minutes`, `10 minutes` (`default`), `5 minutes`, `1 minute`, `0 seconds`.
    - Off: kept while Explorer is reachable, then let go by **AFK Timer**.
  - **LibreOffice TTL** — Same for the engine that draws CorelDRAW and nearby formats.
    - A kept engine converts the next document faster: `1.2 s` cold vs `0.2 s`, but uses a few hundred MB.
    - `0 seconds` means one engine per document.
    - Both TTLs apply to **Persistent** engines; non-persistent ones use **AFK Timer**.
    - Greyed out if LibreOffice is not installed.
  - **WebView2 TTL** — With **Persistent** on: how long the SVG browser stays warm.
    - Off: let go by **AFK Timer**.
    - Greyed out if WebView2 is missing.
  - No `ImageMagick TTL`, `PeaZip TTL`, or `Calibre TTL`: those tools run once and exit, so idle time cannot bound them. A second hover is a cache hit.
- **Codecs** — What this machine has: Videos, Audio, Images, Engines.
  - A missing one carries a cross, and where the README names a page for it, picking the row offers to open that page — nothing is installed or downloaded by the app itself.
- **Run at Startup** — Add or remove the Windows startup entry.
  - On every start, an entry that names another copy of the app — a portable copy, an older version, a path that has moved — is pointed back at the one you are running.
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
pin_enabled=true
pin_key=space
pin_nav_file_types=all
pin_pause_audio=true
pin_pause_video=true
pin_update_enabled=true
pin_update_on_hover=false
preview_enabled=true
run_at_startup=true

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
theme=light

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
audio_seek=remember
audio_volume=10
normalize_video_volume=false
normalize_volume=true
remember_audio_volume=false
remember_video_volume=false
video_volume=0

; Performance
decode_budget_gb=1
document_cache_mb=256
general_disk_cache_mb=128
image_cache_mb=64
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
- A video's engine is the `video_engine` setting; the two lists decide which engine is *asked* about a name — the media engine takes only `[video]` names, and an `[ffmpeg]` name is FFmpeg's alone.

### Memory and cache

- `image_cache_mb`: memory for decoded image frames. Default `64`, max `2048`; `0` holds nothing.
- `document_cache_mb`: disk cache for rendered document pages. Default `256`, max `2048`; `0` keeps nothing between hovers but still draws the current one. Pages live under `%TEMP%\rust-hover-preview\document`; the least recently read page is removed first.
- `image_disk_cache_mb`: disk cache for the pictures ImageMagick developed. Default `512`, max `2048`; `0` keeps nothing between hovers but still develops the current one. Pictures live under `%TEMP%\rust-hover-preview\image`; the least recently read one is removed first, and they survive a restart.
- `general_disk_cache_mb`: disk cache for the subtitle tracks a film's embedded ones are copied into. Default `128`, max `2048`; `0` extracts nothing, and a film whose subtitles are embedded is drawn without them. Files live under `%TEMP%\rust-hover-preview\general`, least recently used first.
- `decode_budget_gb`: most memory one hover may use. Default `1`, range `0.25`–`64`. A file over the limit shows no preview.

### Performance

- `video_hw_accel`: whether a video FFmpeg's player has is decoded on the graphics card rather than on a core — the tray's **Performance → Hardware Acceleration → Video**.
  - Default `true`.
  - It is a setting about this build's FFmpeg rather than about a kind of file: a file the media engine plays is decoded by Windows either way, and a machine with no device FFmpeg can decode on is the one the setting is for, where FFmpeg falls back by itself.

### HDR

- `hdr_tone_map`: how HDR/EXR light becomes screen values: `reinhard` default, `aces`, `srgb`, or `off`. PNG/JPEG are not affected.
- `hdr_exposure`: stops shifted before that curve. Default `0`, range `-10` to `10`.

### Waiting and engines

- `spinner_delay_ms`: wait before showing the loading spinner. Default `250`; `0` shows it immediately.
  - One delay covers all preview types, and a pinned window being shown another file with it.
  - A pin that has to wait shows a spinner of its own in the middle of itself, drawn over the file it is still showing rather than in place of it, and the window is never hidden while it waits.
- `office_engine`: `microsoft_office` default, falls back to LibreOffice when needed; or `libreoffice`, which always uses LibreOffice. If LibreOffice is missing, it falls back and the tray row is greyed out.
- `video_engine`: which engine plays a video: `best` (default), `native`, `ffmpeg`, or `hybrid`.
  - **Best** is the machine's own answer — `hybrid` where FFmpeg is installed (the media engine draws a film at or under 3.2 megapixels, FFmpeg's player takes a larger one or one it could not measure) and `native` where it is not.
  - `native` names the media engine alone, `ffmpeg` names `ffplay` alone, and `hybrid` states that size rule plainly.
  - A choice this machine cannot supply is read as `best`.
- `video_engine_fallback`: `true` (default) lets a chosen engine that cannot play a given file fall through to the other; `false` makes the engine you chose stand alone. Read only where `video_engine` names an engine — **Best** walks both either way.
- `libreoffice_idle` / `office_engine_idle`: seconds the engine is kept after its last page, or `indefinitely`. Default `600`; `0` launches it per document.
- `afk_timer_seconds`: seconds with no Explorer window before a non-persistent engine is released. Default `60`, max one day. `0` releases immediately.
- `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent`: `true` keeps that engine always, bounded only by its `…_idle` time. `false` default keeps it while Explorer is reachable, bounded by `afk_timer_seconds`.

### Trigger and position

- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled`: the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched. Default `true`.
- `trigger_key_affect_pin_mode`: whether the trigger key reaches a pinned preview. Default `false` — the key is not read while a preview is pinned, up or collapsed into its bubble, since a pin is a window you put there rather than a hover for the key to hold back. `true` lets holding the key bring the pin down with the previews it stops. It speaks for `disable`; the `enable` mode is left as it is.
- `pin_key` / `pin_enabled`: the key that pins the preview on screen (`space` by default — the same spellings the trigger key takes: `alt`, `f8`, `a`, `space`, …), and whether it is watched. See **Pin Mode** above for what a pin does.
- `pin_update_enabled`: whether a pin that is up is shown the file picked next, by click or by keyboard, without moving the window. Default `true`.
  - Off, the pin keeps the file it was taken up on until it is closed — a pin is also a window to read in, and one that swapped its file out from under the hand on every keystroke would be unreadable.
  - The keyboard half follows Explorer, so it is inert while a pinned window holds the keyboard — click back into the listing and it resumes.
- `pin_update_on_hover`: whether that following includes the pointer's own hover, or only what a click or a key asks for. Default `false`; greyed in the tray while `pin_update_enabled` is off.
- `pin_nav_file_types`: what the caption's **Previous** and **Next** buttons step through.
  - `all` (default) is every file this build can preview, so a video steps to the sound beside it.
  - `category` narrows the walk to the pinned file's own kind of thing — pictures, video, audio, documents, archives, text, fonts, or design, where a camera raw is a picture and a book is a document.
  - The walk is the folder the pin was taken up in and no subfolder of it, is in the order the Explorer listing is showing, wraps at both ends, and is held per folder, read on the first press rather than when the pin is taken up.
  - A file it reaches that the window cannot show is stepped over, and the walk is bounded by the folder: each other file is offered at most once, so a folder of nothing this build can read ends the walk instead of going round for ever.
  - A file written before this setting existed is read as `all`.
- `pin_pause_video` / `pin_pause_audio`: whether a pin collapsed into its bubble holds what it was playing where it stands, and starts it again at the second it stopped at when the window comes back up. Both default `true`. They are two switches because a video and a sound are two different things to want quiet.
- `follow_cursor`: `true` = Follow Cursor; `false` = Best Position.
- `avoid_mode`: `filename` default, `filename_column`, `details`, or `off` — what the preview avoids.

### Scaling

- `preview_scale`: percentage or `fit`, based on the picture's own size.
- `video_scale`: same for video. Default `100`.
- `audio_scale`: the share of the display a sound's card is laid out over — `25`, `20`, `15`, `10` or `5`. Default `10`; the whole card scales with the share — the font it is set in (the default text size at the `10%` anchor, scaled by the share's fraction of it), its height at that font, and its width, which fills the share — so the card at any share is the `10%` card uniformly smaller or larger, and `Text Size` no longer resizes it. A hand-edited value is sanitized rather than fatal: `100` or more reads as `100`, `0` reads as `10`, and an unrecognized value is ignored.
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

### Old or removed names

- `magick_preview_enabled` and `peazip_preview_enabled`: not read. ImageMagick pictures use `image_preview_enabled`; PeaZip archives use `archive_preview_enabled`.
- `libre_preview_enabled` and `libre_scale`: gone. Use `document_preview_enabled` and `document_scale`.
- `office_cache_mb` and `libre_cache_mb`: gone. Use `document_cache_mb`. If both old ones exist, the larger is read once, then both are removed.
- `text_preview_full_mode`: gone. A pinned text preview comes up in full mode instead, so it scrolls and its text can be selected and copied.
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

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item, and GDI paints the preview into a topmost layered window. Text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives' own tables of contents — or from the listing an installed PeaZip produces for the formats no reader here has — Office documents are drawn from a page Office renders in the background, WebView2 draws SVG documents and font specimens, camera raw and the pictures beside it are developed by ImageMagick, and video is played by FFmpeg where installed and by Windows' own media engine where it is not.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview is local-first: previews work without an internet connection, and the only network request is an update check, which runs only when you open the tray menu and at most once an hour. There is no telemetry, analytics, ads, accounts, or crash reporting; it reads only the item you hover or focus in Explorer, locally and only for enabled preview types, and cloud-only placeholders are skipped on purpose while password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use locally installed FFmpeg, LibreOffice, or ImageMagick when available, plus Microsoft Office, Windows' own media engine, and the Windows PDF engine; caches are bounded by `config.ini` — decoded images stay in memory, while the page an engine drew for a document is kept as a file under the temp folder, where Windows is free to clear it. See `PRIVACY.md` for full details.

## License

MIT. See [LICENSE](LICENSE).

The app is built out of other people's code as much as its own — the Windows bindings, the image decoders, the syntax highlighter, the archive readers, the browser bindings — and each of those carries its own licence, with the notices MIT and BSD ask to be reproduced. [THIRD-PARTY.md](THIRD-PARTY.md) lists every dependency grouped by licence, with the copyright holders beside it, and the full texts are in [`LICENSES/`](LICENSES). Both are generated from the dependency tree rather than kept by hand; `generate-attribution.ps1` refreshes them.
