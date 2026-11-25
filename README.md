<p align="center">
  <img src="assets/icon.png" width="112" height="112" alt="Rust Hover Preview Windows 11 system tray icon" />
</p>

<h1 align="center">Rust Hover Preview — Instant File Hover &amp; Quick Preview for Windows 11 Explorer</h1>

<p align="center">
  <b>Hover over any file in File Explorer to preview it instantly. Press Space or middle-click to pin it.</b>
</p>

<p align="center">
  A free, open-source, lightweight Windows 11 system tray utility for instant file previews in File Explorer —<br/>
  inspired by QTTabBar and QuickLook, built in native Rust &amp; Win32 with support for images, videos, audio, PDFs, Office docs, code, archives, fonts, design files, ebooks, and 3D models.
</p>

<p align="center">
  <a href="https://github.com/cosmokud/rust-hover-preview/releases/latest"><img alt="Latest release" src="https://img.shields.io/github/v/release/cosmokud/rust-hover-preview?style=for-the-badge&color=024bd1" /></a>
  <img alt="Rust" src="https://img.shields.io/badge/Rust-1.98.1%2B-orange?style=for-the-badge&logo=rust" />
  <img alt="Windows 11" src="https://img.shields.io/badge/Platform-Windows%2011-blue?style=for-the-badge&logo=windows" />
  <a href="LICENSE"><img alt="MIT license" src="https://img.shields.io/badge/license-MIT-111111?style=for-the-badge" /></a>
</p>

<p align="center">
  <a href="#overview">Overview</a> ·
  <a href="#features">Features</a> ·
  <a href="#supported-formats">Formats</a> ·
  <a href="#installation">Install</a> ·
  <a href="#usage">Usage</a> ·
  <a href="#configuration">Configuration</a> ·
  <a href="#faq">FAQ</a> ·
  <a href="#build-from-source">Build</a> ·
  <a href="#architecture">Architecture</a> ·
  <a href="#privacy">Privacy</a> ·
  <a href="#security">Security</a> ·
  <a href="#license">License</a> ·
  <a href="CONTRIBUTING.md">Contribute</a>
</p>

---

## Overview

Rust Hover Preview is a free, open-source Windows 11 File Explorer preview utility inspired by **QTTabBar** and **QuickLook** — and a lightweight native alternative to **PowerToys Peek**. Hover over any file in **Windows 11 File Explorer**, **Open/Save Browse Dialogs**, or **voidtools Everything** search results (or navigate with the keyboard) for an instant preview without opening a separate app, or press `Space` or middle-click to pin it into its own movable, resizable window.

Built in native Rust with no Electron overhead and no code injection into File Explorer, it stays light on idle CPU, runs per-user with no admin rights required, and configures cleanly from the system tray menu or a plain `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Features

- **Instant mouse-hover & keyboard previews:** Works in Windows 11 File Explorer, Open/Save **Browse Dialogs**, and **voidtools Everything** search results — hover with the mouse or move with arrow keys, `Tab`, or type-ahead search, with optional **Trigger Key (`Alt`)** gating and hover delays.
- **Pin previews with Space or middle-click:** Turn any preview into a movable, resizable window of its own — always on top by default, with **Minimize to Bubble** or **Minimize to Taskbar** (live preview + media buttons), **Detach Window** to keep multiple previews open, folder navigation, and interactive PDF/text/audio controls.
- **Broad format coverage out of the box:** Previews images, videos, audio, ebooks, PDFs, Office documents, design files (PSD, AI, Figma, Sketch, XD), syntax-highlighted text and code, archives, camera RAW, vector drawings (SVG, EPS, WMF/EMF), fonts, and 3D models.
- **Flexible scaling & placement:** Independent sizing (fit-to-screen, percentage-based, or 25%–400% native scale) across all 11 file categories, **Best Position** vs **Follow Cursor** with filename avoidance, and **Bypass Placement** to size against the whole screen.
- **Lightweight native tray app & plain `config.ini`:** System tray controls alongside a hand-editable `config.ini`; DPI-aware, single-instance, sleep/resume resilient, and light on idle CPU.
- **Plugs into what you already have:** Built-in decoders and Windows codecs cover everyday formats; ImageMagick, FFmpeg, mpv, LibreOffice, Microsoft Office, WebView2, PeaZip, Calibre, f3d, and the Windows PDF engine extend support where installed — nothing bundled.
- **Custom preview with your own programs:** Open `Tray → Custom Preview → Edit` to map any file extension to a program of your choice — with your own arguments and an optional forced placement — preview its window live on hover with its own scaling and volume, and open or promote it cleanly when pinned.
- **Privacy-first & safe by default:** Local-first with zero telemetry, skips OneDrive/SharePoint/Dropbox online-only placeholders without downloading, and blocks `.env`, SSH keys, and other secret files by default — see [Privacy](#privacy).

## Supported Formats

Rust Hover Preview covers everyday and specialized formats across 11 categories — including PNG, JPG, WebP, GIF, HEIC, AVIF, JXL, Camera RAW, MP4, MKV, WebM, MP3, FLAC, OGG, Opus, PDF, CBZ/CBR, EPUB, MOBI, DOCX, XLSX, PPTX, Markdown, source code, ZIP, 7z, RAR, TAR, SVG, TTF/OTF/WOFF2, PSD, AI, Figma, and 3D CAD/mesh files (STL, GLB, OBJ, FBX, STEP).

### Images

| Engine                                                                                                                                          | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ----------------------------------------------------------------------------------------------------------------------------------------------- | ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This app's own decoders — nothing to install                                                                                                    | `[image]`    | apng, bmp, cur, dds, dib, exr, exif, ff, gif, hdr, ico, icon, jfif, jpe, jpeg, jpg, pam, pbm, pgm, png, pnm, ppm, qoi, rle, tga, tif, tiff                                                                                                                                                                                                                                                                                                                                                                                                                           |
| The codecs the [Store's image extensions](#optional-enable-heic-avif-jpeg-xl-webp-camera-raws-preview-windows-codecs) add — each installed once | `[image]`    | 3fr, arw, avci, avcs, avif, avifs, cr2, cr3, crw, dcr, dng, erf, fff, heic, heics, heif, heifs, hif, iiq, jxl, jxr, k25, kdc, mef, mos, mrw, nef, nrw, orf, pef, raf, raw, rw2, rwl, sr2, srf, srw, wdp, webp, x3f                                                                                                                                                                                                                                                                                                                                                   |
| [ImageMagick](#optional-enable-camera-raw-and-more-pictures-imagemagick)                                                                        | `[magick]`   | aai, art, ase, aseprite, bayer, bayera, bgr, bgra, bgro, c2pa, cal, cals, cmyk, cmyka, cube, cut, dcm, dcx, dpx, dxt1, dxt5, fax, fit, fits, fl32, fts, ftxt, g3, g4, gray, graya, group4, hrz, icb, ipl, j2c, j2k, jng, jnx, jp2, jpc, jpm, jpt, mac, map, mat, mdc, miff, mng, mono, mpc, mtv, otb, pal, palm, pango, pcds, pes, pfm, pgx, phm, picon, pict, pix, pwp, rgb, rgb565, rgba, rgbo, rgf, rla, rmf, scr, sct, sf3, sfw, sgi, six, sixel, stegano, sun, tim, tm2, uyvy, vda, vicar, viff, vips, vst, wbinfo, wbmp, xbm, xcf, xpm, xv, ycbcr, ycbcra, yuv |

### Videos

For video, **mpv is recommended** for performance, hardware acceleration, reliability, and subtitles (embedded tracks and sidecar `.srt`) that just work.

| Engine                                                                                                                                                           | `config.ini`           | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                               |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| The media engine Windows 11 ships — H.264, MPEG-4 and WMV are decoded out of the box                                                                             | `[video]`              | 3g2, 3gp, 3gpp, asf, avi, dvr-ms, m2ts, m4v, mkv, mov, mp4, mts, qt, ts, wmv                                                                                                                                                                                                                                                                                                                                                                     |
| The codecs the [Store's video extensions](#optional-enable-more-video-codecs-windows-codecs) add — HEVC, VP9, AV1, MPEG-1/MPEG-2 and Theora, each installed once | `[video]`              | m1v, m2t, m2v, mpe, mpeg, mpg, vob, webm                                                                                                                                                                                                                                                                                                                                                                                                         |
| [FFmpeg](#optional-enable-video-preview-and-advanced-audio-processing-with-ffmpeg)'s player (`ffplay`)                                                           | `[ffmpeg]`             | 264, 265, 266, apv, av1, avc, avs, avs2, avs3, bik, bk2, c93, cavs, cdg, cdxl, cin, cpk, dav, dif, divx, drc, dv, evc, f4v, flm, flv, gxf, h261, h263, h264, h265, h266, h26l, hevc, ifv, imx, ismv, ivf, ivr, kux, m2p, mj2, mjpeg, mjpg, mk3d, moflex, mpv, mve, mvi, mxf, mxg, nsv, nut, obu, ogm, ogv, pmp, psp, rcv, rm, rmvb, roq, rsd, smk, str, swf, thp, tod, tp, tr, ty, ty+, usm, vc1, vc2, viv, vro, vvc, vw, wtv, xl, xmv, y4m, yop |
| [mpv](#optional-enable-video-preview-with-mpv)                                                                                                                   | — (no list of its own) | union of `[video]` + `[ffmpeg]` above (23 + 84 = 107 names)                                                                                                                                                                                                                                                                                                                                                                                      |

### Audio

| Engine                                                                                                                                                                                        | `config.ini` | Default extensions                                                                                           |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------ |
| The media engine Windows 11 ships — FLAC, ALAC, AAC, MP3 and MPEG audio, AMR-NB, WMA, AIFF and WAV are decoded by codecs Windows has                                                          | `[audio]`    | aac, aif, aifc, aiff, amr, flac, m4a, m4b, mp2, mp3, mpa, wav, wave, wma                                     |
| The media engine with the [Store's media extension](#optional-enable-ogg-vorbis-and-opus-preview-windows-media-extension) — the Ogg demuxer and Vorbis, which Windows 11 does not always ship | `[audio]`    | oga, ogg, opus                                                                                               |
| [FFmpeg](#optional-enable-video-preview-and-advanced-audio-processing-with-ffmpeg)'s player — nothing Windows ships decodes these                                                             | `[audio]`    | ac3, ape, au, awb, caf, dff, dsf, dts, dtshd, eac3, mka, mpc, ofr, ofs, ra, shn, snd, spx, tak, tta, voc, wv |

### Text

| Engine                       | `config.ini`          | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| ---------------------------- | --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This app, syntax-highlighted | `[text]` `extensions` | adb, adoc, ads, asciidoc, asm, asp, aspx, astro, awk, bash, bat, bib, bzl, c, cc, cfg, cg, cjs, clj, cljc, cljs, cmake, cmd, comp, conf, cpp, cs, csh, cshtml, css, csv, csx, cts, cxx, d, dart, diff, diz, edn, ejs, el, elm, env, erb, erl, ex, exs, f, f03, f77, f90, f95, fish, for, frag, fs, fsi, fsx, ftn, fx, geom, glsl, go, gql, gradle, graphql, groovy, h, haml, hbs, hcl, hh, hlsl, hpp, hrl, hs, htm, html, hxx, inc, ini, ipynb, java, jl, js, json, json5, jsonc, jsonl, jsp, jsx, ksh, kt, kts, latex, less, lhs, liquid, lisp, ll, lock, log, lsp, lua, m, mak, man, markdown, md, mdown, metal, mjs, mk, mkd, ml, mli, mm, mts, mustache, nasm, nfo, nim, ninja, nix, njk, org, pas, patch, php, phtml, pl, plist, pm, properties, proto, ps1, psd1, psm1, py, pyi, pyw, r, rake, rb, rkt, rmd, rs, rst, rtf, s, sass, scala, scm, scss, sh, slim, sol, sql, srt, ss, styl, sv, svelte, svh, swift, tcl, tex, text, tf, tfvars, toml, ts, tsv, tsx, twig, txt, v, vbs, vert, vhd, vhdl, vtt, vue, wat, wgsl, xhtml, xml, xsd, xsl, xslt, yaml, yml, zig, zsh |
| This app, as plain text      | `[text]` `names`      | authors, .babelrc, brewfile, caddyfile, changelog, changes, .clang-format, .clang-tidy, cmakelists.txt, code\_of\_conduct, containerfile, contributing, contributors, copying, copyright, dockerfile, .dockerignore, .editorconfig, .env, .env.example, .env.local, .eslintignore, .eslintrc, gemfile, .gitattributes, .gitconfig, .gitignore, .gitkeep, .gitmodules, gnumakefile, .golangci.yml, history, .htaccess, install, jenkinsfile, justfile, licence, license, .mailmap, makefile, makefile.am, makefile.in, notice, .npmignore, .prettierignore, .prettierrc, procfile, rakefile, readme, .rustfmt.toml, security, .stylelintrc, unlicense, vagrantfile                                                                                                                                                                                                                                                                                                                                                                                                               |

A secret or credential file is never previewed while **Don't Preview Secret Files** is on, which it is by default: `.env` and its variants, `id_rsa` and the like, key and keystore extensions, and anything under `~/.ssh`, `~/.aws` or a coding agent's credential folder. A file with no extension at all is previewed as text only under **Preview All Extensionless**, which is off by default.

### Ebook

| Engine                                                                                                                           | `config.ini` | Default extensions                                                          |
| -------------------------------------------------------------------------------------------------------------------------------- | ------------ | --------------------------------------------------------------------------- |
| The Windows PDF engine for hover previews, WebView2 for interactive pinned documents; this app's own readers for comic containers | `[ebook]`    | cbc, cbr, cbz, epdf, pdf, pdfa                                              |
| [Calibre](#optional-enable-ebooks-calibre), which converts the book to a PDF first                                               | `[calibre]`  | azw, azw3, azw4, djvu, epub, fb2, htmlz, lit, lrf, mobi, pml, prc, snb, tcr |

### Archives

| Engine                                              | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                |
| --------------------------------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| This app, from each archive's own table of contents | `[archive]`  | 7z, apk, jar, rar, tar, tar.gz, tgz, xpi, zip, zipx                                                                                                                                                                                                                                               |
| [PeaZip](#optional-enable-niche-archives-peazip)    | `[peazip]`   | 001, apfs, ar, arc, arj, bcm, br, bz2, bzip2, cab, chm, cpio, cramfs, deb, dmg, esd, gz, gzip, hfs, hfsx, hxs, iso, lha, lpaq8, lzh, lzma, msi, msp, pkg, ppkg, qcow, qcow2, rpm, squashfs, swm, taz, tbz, tbz2, tpz, txz, tzst, udf, udeb, vdi, vhd, vhdx, vmdk, wim, xar, xip, xz, z, zpaq, zst |

### Document

| Engine                                                                                                                             | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                              |
| ---------------------------------------------------------------------------------------------------------------------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Microsoft Office, falling back to [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice) where Office is absent | `[office]`   | doc, docm, docx, dot, dotm, dotx, pot, potm, potx, pps, ppsm, ppsx, ppt, pptm, pptx, xls, xlsb, xlsm, xlsx, xlt, xltm, xltx                                                                                                                                                                                                                                                                                                     |
| [LibreOffice](#optional-enable-coreldraw-and-other-documents-libreoffice)                                                          | `[libre]`    | 123, 602, abw, cdr, cgm, cmx, cwk, dbf, dif, dxf, fodg, fodp, fodt, gnm, gnumeric, hwp, key, lwp, mcw, met, mw, numbers, odb, odc, odf, odg, odm, odp, ods, odt, oth, otg, otm, otp, ots, ott, pages, pcd, pct, pcx, pdb, pm6, pmd, psw, pub, ras, sda, sdc, sdd, sdw, slk, stc, std, sti, stw, svm, sxd, sxg, sxi, sxm, sxw, vdx, vsd, vsdm, vsdx, vstx, wb2, wk1, wk3, wk4, wks, wpg, wq1, wq2, wpd, wps, wri, xlw, zabw, zmf |

### Vector

| Engine                                                                                                    | `config.ini` | Default extensions                                         |
| --------------------------------------------------------------------------------------------------------- | ------------ | ---------------------------------------------------------- |
| WebView2 for `svg` and `svgz`; the drawing layer for the metafiles and the preview picture an EPS carries | `[vector]`   | emf, epi, eps, epsf, epsi, ept, ept2, ept3, svg, svgz, wmf |

### Fonts

| Engine                                         | `config.ini` | Default extensions         |
| ---------------------------------------------- | ------------ | -------------------------- |
| WebView2 — a specimen page with the font in it | `[font]`     | otf, ttc, ttf, woff, woff2 |

### Design

| Engine                                                                       | `config.ini` | Default extensions                                 |
| ---------------------------------------------------------------------------- | ------------ | -------------------------------------------------- |
| This app, from the flattened picture each format keeps of the whole document | `[design]`   | ai, fig, kra, ora, procreate, psb, psd, sketch, xd |

### 3D Models

| Engine                                                                              | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                            |
| ----------------------------------------------------------------------------------- | ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [f3d](#optional-enable-3d-models-f3d), which draws a scene into a window of its own | `[f3d]`      | 210, 3ds, 3mf, abc, amf, bin, bpf, brep, cdf, csd, dae, dxf, ex2, exo, fbx, fcstd, glb, gltf, gml, ifc, iges, igs, las, laz, mdl, mha, mhd, nc, ncdf, nhdr, nrrd, obj, off, p21, pic, ply, pts, ptx, qi, sbet, slpk, splat, spz, step, stl, stp, stpnc, usd, usda, usdc, usdz, vdb, vrml, vti, vtk, vtkhdf, vtm, vtp, vtr, vts, vtu, wrl, xbf |

## Installation

Download the latest Windows 11 build from the repository's Releases page:

- [Download the latest release from GitHub](https://github.com/cosmokud/rust-hover-preview/releases/latest)

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry. **No administrator rights required.**
- `rust-hover-preview_<version>_x64-portable.zip` — Portable build: `rust-hover-preview.exe`, the `custom_preview_hook.dll` it loads, and the `.portable` marker that keeps everything beside itself. Unzip it into any folder.

1. Open [Releases](https://github.com/cosmokud/rust-hover-preview/releases/latest).
2. Download your preferred asset.
3. Run the installer, or unzip the portable build wherever you like — keeping all three files together, since the DLL is what `In-Process Hook` loads and `.portable` is what keeps `config.ini`, `theme`, `engines`, `webview` and `libreoffice` in that folder instead of in your profile. Only the caches stay in `%TEMP%`, and `Start with Windows` still writes the one entry in the registry.
4. Launch Rust Hover Preview.

No Rust toolchain is needed.

Runs on any processor Windows 11 itself supports: the build asks for SSE4.2 and POPCNT, not AVX2, so x64 emulation on an ARM64 device runs it too.

### Optional: Enable Video Preview with mpv

For smoother playback and a better picture, install mpv: videos then play through it automatically, using your graphics card, with subtitles just working. The app finds `mpv.exe` beside itself, in your `PATH`, in `MPV Player` under Program Files, and in SVP 4's `mpv64`; where more than one is found, **Components → Videos → MPV** lets you pick which one plays.

**Option A: winget**

```powershell
winget install --id shinchiro.mpv -e
```

**Option B: manual**

1. Download a Windows build from [https://mpv.io/installation/](https://mpv.io/installation/)
2. Extract it, for example to `%LOCALAPPDATA%\Programs\mpv`
3. Add `%LOCALAPPDATA%\Programs\mpv` to your user `PATH`

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
2. Extract it, for example to `%LOCALAPPDATA%\Programs\ffmpeg`
3. Add `%LOCALAPPDATA%\Programs\ffmpeg\bin` to your user `PATH`

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

All are free, official from Microsoft. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray's **Components** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL, WebP, Camera Raws Preview (Windows Codecs)

A **still** `heic`, `heif`, `avif`, `jxl`, and still `webp` is decoded by a codec Windows provides rather than one shipped with the app. Each still needs its extension installed once from the Microsoft Store:

| Still format                | Needs                                                                                                                                            |
| --------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `heic`, `heif`              | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`                      | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`                       | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`                      | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |
| camera raws (`3fr` … `x3f`) | [Raw Image Extension](https://apps.microsoft.com/detail/9NCTDW2W1BH8) — optional: an installed ImageMagick develops one without it               |

All are free, official from Microsoft. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error.

Which of these a moving file needs is under [Supported Formats](#supported-formats) above.

A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

### Optional: Enable Ogg, Vorbis and Opus Preview (Windows Media Extension)

Windows 11 decodes Opus, but not always the Ogg container it travels in. Without the extension below, an `ogg`, `oga` or `opus` file is handed to FFmpeg instead, and shows no preview where FFmpeg is not installed.

| Audio format                          | Needs                                                                  |
| ------------------------------------- | ---------------------------------------------------------------------- |
| `ogg`, `oga` (Vorbis, Theora), `opus` | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS) |

It is free and official from Microsoft. Windows 11 does not always ship it. Installing it takes effect the next time the tray's **Components** menu is opened — no restart, nothing to configure.

Which of these a sound needs is under [Supported Formats](#supported-formats) above.

### Optional: Enable CorelDRAW and Other Documents (LibreOffice)

`cdr` and everything else in the `[libre]` list is drawn by **LibreOffice**, which this app drives where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id TheDocumentFoundation.LibreOffice
```

Or download the official installer from [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/) and run the suggested file (`LibreOffice_*_Win_x86-64.msi`; LibreOffice ships an `.msi`, not an `.exe`).

`dxf` is in this list and in the `[f3d]` one, because the two engines draw different halves of the format: LibreOffice imports a drafting sheet whole, and f3d reads only the 3D geometry it holds. Which of them a file is, is decided by the file rather than by the name — see [3D Models](#optional-enable-3d-models-f3d).

### Optional: Enable Camera Raw and More Pictures (ImageMagick)

The camera raws and the pictures beside them in `[image]` are decoded by the codecs Windows has where those are installed — the **Raw Image Extension** for a `.nef`, a `.cr3`, a `.dng` or any other raw, and the HEIF, JPEG XL or WebP extension for the still beside them. Where one of those codecs is missing, and for everything in the `[magick]` list, **ImageMagick** develops the picture — which this app runs where it finds it, with nothing bundled and nothing to configure. Install it once:

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

### Optional: Enable 3D Models (F3D)

`stl` and everything else in the `[f3d]` list — meshes, CAD exchange files, point clouds and scientific datasets — is drawn by **f3d**, which this app runs where it finds it. Nothing is bundled with the app, and there is nothing to configure: without f3d a model shows no preview at all rather than a broken one. Install it once:

```text
winget install -e --id f3d-app.f3d
```

Or download the official Windows build from [https://github.com/f3d-app/f3d/releases](https://github.com/f3d-app/f3d/releases).

> [!WARNING]
> `winget` keeps its sources and its state in your own profile — `%LOCALAPPDATA%\Microsoft\WinGet` — so every step below is run from your own Terminal, and none of them as an administrator: an elevated window belongs to whichever administrator account it was opened from, so it repairs that account's sources rather than yours, or finds none at all.
>
> 1. Read what actually failed: `winget source list`.
> 2. Close every other `winget` window — an install or an update holding the sources is enough to fail the next one — and try again.
> 3. Rebuild the sources this downloads from: `winget source remove --name msstore`, then `winget source update`. If the default source is the one failing, `winget source remove --name winget`, then `winget source add --name winget https://cdn.winget.microsoft.com/cache`, then `winget source update`.
> 4. Still failing: close every `winget` window and rename `%LOCALAPPDATA%\Microsoft\WinGet\State` to `State.bak`. The next run rebuilds it. Your installed programs are untouched; what goes is the cached index and `winget`'s own ledger of what it installed, which it rebuilds from the registry.
>
> `winget source reset --force` is not the way back from this: it re-adds and re-downloads the whole index, so it fails for every network and certificate reason as well, and — since the sources live in your profile — running it elevated changes nothing at all.

The tray does the install for you where it can: right-click the tray icon and pick the engine under `Components → Install` — mpv, FFmpeg, LibreOffice, ImageMagick, PeaZip, Calibre and f3d — which opens a console window showing the same `winget` command run without administrator rights and re-checks the machine when it is done.

## Usage

1. Start the app — a tray icon appears.
2. Hover files in File Explorer to preview them.
3. Or navigate with the keyboard — arrow keys, `Tab`, or a file's name typed to jump to it — to preview the focused item.
4. Use the middle mouse button or `Space` (by default) to pin the preview.
5. Right-click the tray icon to configure behavior.

## System Tray Menu

See [docs/tray.md](docs/tray.md) for every row in the tray menu and what each one does.

## Configuration

See [docs/configuration.md](docs/configuration.md) for every `config.ini` setting, the file's default contents, and the names that are gone.

## FAQ

### How is this different from QuickLook, PowerToys Peek, or QTTabBar?

QuickLook and PowerToys Peek require selecting a file and pressing a shortcut to preview it, while QTTabBar injects a shell add-on into File Explorer. Rust Hover Preview gives you instant mouse-hover and keyboard previews across File Explorer, Open/Save dialogs, and Everything with no Explorer code injection, plus pinning, Bubble minimize, and multi-window Detach in one lightweight native app.

| Capability                         | Rust Hover Preview          | QuickLook             | PowerToys Peek             | QTTabBar           |
| :--------------------------------- | :-------------------------- | :-------------------- | :------------------------- | :----------------- |
| **Mouse-hover preview**            | Yes (instant, no click)     | No (select + `Space`) | No (select + `Ctrl+Space`) | Yes                |
| **Open/Save dialogs & Everything** | Yes (on by default)         | File dialogs (varies) | Explorer only              | Explorer only      |
| **Pin / Bubble minimize / Detach** | Yes (`Space`/middle-click)  | Single preview window | Single preview window      | Hover tooltip only |
| **Custom preview per extension**   | Yes (`Tray → Custom Preview`) | Via plugins         | Built-in set               | Built-in set       |
| **No Explorer code injection**     | Yes                         | Yes                   | Yes                        | No (shell add-on)  |

### How do I preview files on hover in Windows 11 File Explorer without opening them?

Run Rust Hover Preview in the system tray and hover over any file—or move through files with arrow keys—in Windows 11 File Explorer. An instant preview appears for images, videos, audio, PDFs, Office docs, archives, code, fonts, ebooks, and 3D models, and pressing `Space` or middle-clicking pins it into its own window.

### Can I preview files outside File Explorer, like in Open/Save dialogs or third-party apps?

Yes. Rust Hover Preview works in **voidtools Everything** search results and standard Windows **Open / Save As** dialogs out of the box with both mouse hover and keyboard navigation, and either can be toggled in the tray menu.

### How do I preview unsupported file formats in Windows File Explorer?

Open **Tray → Custom Preview → Edit** and map any file extension to a Windows program you already have installed. Rust Hover Preview renders that app's window live on hover (using a lightweight helper hook on the target app to prevent window flashing, with non-hook fallback modes available) and opens or promotes it cleanly when pinned.

### Can I pin a file preview on top or view multiple previews side by side?

Yes. Press `Space` or middle-click to pin a preview always on top, use **Detach Window** to keep multiple files open side by side, or minimize a pinned preview to a floating desktop **Bubble** or the taskbar with live playback controls.

### Can I prevent accidental hover previews or pause them temporarily?

Press `F7` in File Explorer anytime to pause or resume previews, or enable **Trigger Key (`Alt`)** under **Timing** so previews only appear while holding `Alt` (or stay hidden while holding it). You can also add a **Hover Delay** or set **Placement** to avoid covering filenames.

### Does it require administrator privileges or download OneDrive cloud files?

No. Rust Hover Preview installs per-user with no admin rights, never injects code into File Explorer, skips online-only OneDrive, SharePoint, and Dropbox files without downloading them, and blocks `.env` files and SSH keys by default.

Starting it as an administrator works, with two things worth knowing. **The startup entry is written into the account at the desktop**, even where the elevation prompt was answered by a *different* administrator's credentials: that account's own `Run` key under `HKEY_USERS` is used instead of the elevated token's, so the next logon of the user at the keyboard still starts the app — an elevated run writing into its own hive would be an entry nobody logs on to read. And a `Run` entry always starts **unelevated**, because Windows never raises a UAC prompt for one, so the copy the logon starts is the ordinary one even where the copy it was registered from was elevated.

## Build from Source

Requirements: Windows 11, Rust 1.98.1+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK. The build targets the x86-64-v2 level (`-C target-cpu=x86-64-v2` in `.cargo/config.toml`), so the machine that compiles it has to be one that can run the result.

```bash
git clone https://github.com/cosmokud/rust-hover-preview.git
cd rust-hover-preview
cargo build            # debug
cargo build --release  # release
```

The release binary is written to `target/release/rust-hover-preview.exe`. A release build ends a running copy first, since Windows will not let the linker replace an open binary. Debug builds are left alone.

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, test commands, and hooks.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item, and GDI paints the preview into a topmost layered window. Text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives' own tables of contents — or from the listing an installed PeaZip produces for the formats no reader here has — Office documents are drawn from a page Office renders in the background, WebView2 draws SVG documents and font specimens, hosts the interactive pinned PDF viewer, and runs HTML pages where Render HTML is switched on, camera raw and the pictures beside it are developed by ImageMagick, and video is played by mpv where installed, and otherwise by Windows' own media engine for small films and FFmpeg's player for large ones. A middle click or `Space` pins the preview into a captioned window of its own — always on top unless you turn that off, in which case it wears a taskbar button of its own.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview is local-first: previews work without an internet connection, and the only network request is an update check, which runs at startup, on every opening of the tray menu and once per `update_check_hours` hours (six by default, the one period both update switches share), never while `Check for Updates` is switched off — or, where you have turned `Silent Auto-Update` on, the installer of the release a check finds, run without asking but only while the app is idle: nothing pinned, no preview on screen and no menu or dialog open, so nothing of yours is ever closed by it. There is no telemetry, analytics, ads, accounts, or crash reporting; it reads only the item you hover or focus in Explorer, locally and only for enabled preview types, and cloud-only placeholders are skipped on purpose while password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use locally installed mpv, FFmpeg, LibreOffice, or ImageMagick when available, plus Microsoft Office, Windows' own media engine, and the Windows PDF engine; caches are bounded by `config.ini` — decoded images stay in memory, while the page an engine drew for a document is kept as a file under the temp folder, where Windows is free to clear it. See [PRIVACY.md](PRIVACY.md) for full details.

## Security

Please report vulnerabilities privately, not in a public issue. See [SECURITY.md](SECURITY.md).

## License

Licensed under the MIT License. See [LICENSE](LICENSE).

The app is built out of other people's code as much as its own — the Windows bindings, the image decoders, the syntax highlighter, the archive readers, the browser bindings — and each of those carries its own licence, with the notices MIT and BSD ask to be reproduced. [THIRD-PARTY.md](THIRD-PARTY.md) lists every dependency grouped by licence, with the copyright holders beside it, and the full texts are in [`LICENSES/`](LICENSES). Both are generated from the dependency tree rather than kept by hand; `generate-attribution.ps1` refreshes them.

---

<h2 align="center">Part of Rust Windows Series</h2>

<div align="center">
<table>
  <tr>
    <td align="center" width="130">
      <a href="https://github.com/cosmokud/rust-bg-muter"><img src="assets/series/rust-bg-muter.svg" width="56" height="56" alt="Rust Background Muter" /></a><br/>
      <sub><b>Background Muter</b></sub>
    </td>
    <td align="center" width="130">
      <a href="https://github.com/cosmokud/rust-image-viewer"><img src="assets/series/rust-image-viewer.svg" width="56" height="56" alt="Rust Image Viewer" /></a><br/>
      <sub><b>Image Viewer</b></sub>
    </td>
    <td align="center" width="130">
      <a href="https://github.com/cosmokud/rust-hover-preview"><img src="assets/series/rust-hover-preview.svg" width="56" height="56" alt="Rust Hover Preview" /></a><br/>
      <sub><b>Hover Preview</b></sub>
    </td>
    <td align="center" width="130">
      <a href="https://github.com/cosmokud/rust-win-cleaner"><img src="assets/series/rust-win-cleaner.svg" width="56" height="56" alt="Rust Win Cleaner" /></a><br/>
      <sub><b>Win Cleaner</b></sub>
    </td>
  </tr>
</table>
</div>
