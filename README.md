# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.98.1+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

Rust Hover Preview is a Windows 11 tray app inspired by QTTabBar and QuickLook. Hover over a file in File Explorer for an instant preview, or press`Space`to open a pinned preview window. Works with keyboard navigation, and customizes easily via the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer, beside the cursor or the focused item and kept on screen.
- Press the middle mouse button or `Space` (or your own `pin_key`) and a preview becomes a window of its own — captioned, movable, resizable, and there until you close it, always on top unless `Pin Mode → Always On Top` is turned off. 
- Supports images, videos, audios, ebooks, Office documents, design documents, text/code, archives, camera raw, vector drawings, fonts, and more.
- Granular display scaling: Flexible sizing (fit-to-screen, percentage-based, or 25%–400% native scale) with independent scaling profiles for media vs. text, vectors, and documents.
- Tray menu and hand-editable `config.ini`; DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.
- Plugs into what you have: built-in decoders and Windows codecs cover the basics; ImageMagick, FFmpeg, mpv, LibreOffice, Office, WebView2, PeaZip, Calibre, f3d, and the Windows PDF engine extend it where installed — nothing bundled.

## Supported Formats

### Images


| Engine                                                                   | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| --- | --- | --- |
| This app's own decoders and the codecs Windows ships                     | `[image]`    | 3fr, apng, arw, avci, avcs, avif, avifs, bmp, cr2, cr3, crw, cur, dcr, dds, dib, dng, erf, exr, exif, ff, fff, gif, hdr, heic, heics, heif, heifs, hif, ico, icon, iiq, jfif, jpe, jpeg, jpg, jxl, jxr, k25, kdc, mef, mos, mrw, nef, nrw, orf, pam, pbm, pef, pgm, png, pnm, ppm, qoi, raf, raw, rle, rw2, rwl, sr2, srf, srw, tga, tif, tiff, wdp, webp, x3f                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| [ImageMagick](#optional-enable-camera-raw-and-more-pictures-imagemagick) | `[magick]`   | aai, art, ase, aseprite, bayer, bayera, bgr, bgra, bgro, c2pa, cal, cals, cmyk, cmyka, cube, cut, dcm, dcx, dpx, dxt1, dxt5, fax, fit, fits, fl32, fts, ftxt, g3, g4, gray, graya, group4, hrz, icb, ipl, j2c, j2k, jng, jnx, jp2, jpc, jpm, jpt, mac, map, mat, mdc, miff, mng, mono, mpc, mtv, otb, pal, palm, pango, pcds, pes, pfm, pgx, phm, picon, pict, pix, pwp, rgb, rgb565, rgba, rgbo, rgf, rla, rmf, scr, sct, sf3, sfw, sgi, six, sixel, stegano, sun, tim, tm2, uyvy, vda, vicar, viff, vips, vst, wbinfo, wbmp, xbm, xcf, xpm, xv, ycbcr, ycbcra, yuv |


`[image]` is the master list for pictures, and it holds everything Windows decodes that *is* a
picture: the ones this app decodes itself and the ones a codec Windows ships decodes — `heic`,
`heif`, `avif`, `avci`, `jxl`, `jxr`, `wdp`, a still `webp`, the icon and cursor resources
`ico`/`icon`/`cur`, the bitmap shapes `dib`/`rle`/`exif`, and every camera raw from `3fr` to
`x3f`.  Everything an extension provides has to be installed once — the ones a codec Windows ships (`heic`, `heif`, `avif`, `avci`, `jxl`, `jxr`, `wdp`, `webp`) come from the [Store
extensions](#optional-enable-heic-avif-jpeg-xl-and-webp-preview-windows-codecs), and the camera
raws from the **Raw Image Extension** beside them — and every other entry of `[image]` needs
nothing installed. Where a codec is missing, a [ImageMagick](#optional-enable-camera-raw-and-more-pictures-imagemagick)
install develops the picture instead of Windows.

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

A secret or credential file is never previewed while **Don't Preview Secret Files** is on, which it is by default: `.env` and its variants, `id_rsa` and the like, key and keystore extensions, and anything under `~/.ssh`, `~/.aws` or a coding agent's credential folder. A file with no extension at all is previewed as text only under **Preview All Extensionless**, which is off by default.


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


### 3D Models


| Engine                                                     | `config.ini` | Default extensions                                                                                                                                                                                                                                                                                                         |
| --- | --- | --- |
| [f3d](#optional-enable-3d-models-f3d), which draws a scene into a window of its own | `[f3d]`      | 210, 3ds, 3mf, abc, amf, bin, bpf, brep, cdf, csd, dae, dxf, ex2, exo, fbx, fcstd, glb, gltf, gml, ifc, iges, igs, las, laz, mdl, mha, mhd, nc, ncdf, nhdr, nrrd, obj, off, p21, pic, ply, pts, ptx, qi, sbet, slpk, splat, spz, step, stl, stp, stpnc, usd, usda, usdc, usdz, vdb, vrml, vti, vtk, vtkhdf, vtm, vtp, vtr, vts, vtu, wrl, xbf |


## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

1. Open [Releases](../../releases).
2. Download your preferred asset.
3. Run the installer, or place the portable binary wherever you like.
4. Launch Rust Hover Preview.

No Rust toolchain is needed.

Runs on any processor Windows 11 itself supports: the build asks for SSE4.2 and POPCNT, not AVX2, so an x64 emulation on an ARM64 device runs it too.

### Optional: Enable Video Preview with mpv

For smoother playback and a better picture, install mpv: videos then play through it automatically, using your graphics card, with subtitles just working — including any `.srt` file beside the video.

The app finds `mpv.exe` beside itself, in your `PATH`, in `MPV Player` under Program Files, and in SVP 4's `mpv64`; where more than one is found, **Components → Videos → MPV** lets you pick which one plays.

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


All are free, official from Microsoft. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray's **Components** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL, WebP, Camera Raws Preview (Windows Codecs)

A **still** `heic`, `heif`, `avif`, `jxl`, and still `webp` is decoded by a codec Windows provides rather than one shipped with the app. Each still needs its extension installed once from the Microsoft Store:


| Still format   | Needs                                                                                                                                            |
| --- | --- |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |
| camera raws (`3fr` … `x3f`) | [Raw Image Extension](https://apps.microsoft.com/detail/9NCTDW2W1BH8) — optional: an installed ImageMagick develops one without it |


All are free, official from Microsoft. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error.

Which of these a moving file needs is under [Supported Formats](#supported-formats) above.

A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

### Optional: Enable CorelDRAW and Other Documents (LibreOffice)

`cdr` and everything else in the `[libre]` list is drawn by **LibreOffice**, which this app drives where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id TheDocumentFoundation.LibreOffice
```

Or download the official installer from [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/) and run the suggested file (`LibreOffice_*_Win_x86-64.msi`; LibreOffice ships an `.msi`, not an `.exe`).

`dxf` is in this list and in the `[f3d]` one, because the two engines draw different halves of the format: LibreOffice imports a drafting sheet whole, and f3d reads only the 3D geometry it holds. Which of them a file is, is decided by the file rather than by the name — see [3D Models](#optional-enable-3d-models-f3d).

### Optional: Enable Camera Raw and More Pictures (ImageMagick)

The camera raws and the pictures beside them in `[image]` are decoded by the codecs Windows has
where those are installed — the **Raw Image Extension** for a `.nef`, a `.cr3`, a `.dng` or any
other raw, and the HEIF, JPEG XL or WebP extension for the still beside them. Where one of those
codecs is missing, and for everything in the `[magick]` list, **ImageMagick** develops the picture
— which this app runs where it finds it, with nothing bundled and nothing to configure. Install it
once:

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
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys, Tab, or a file's name typed to jump to it — to preview the focused item.
4. Use the middle mouse button or `Space` (by default) to pin the preview.
5. Right-click the tray icon to configure behavior.

Pin mode works like QuickLook: pressing the middle mouse button or `Space` (`pin_key`, watched while `pin_enabled`) turns the settled preview into a window of its own — always on top unless **Pin Mode → Always On Top** is off, in which case it is an ordinary window that also wears this app's own taskbar button while it is up. `Space` in a focused pin pauses and resumes a sound; a minimized pin sits in a draggable bubble — or in that taskbar button — that a middle click or `Space` in Explorer brings back. Arrow keys walk siblings, `Escape` closes. A right-click on a pinned window opens its own menu, and a press inside a pinned page of text gives it the caret: the arrows then move it a line at a time and the page keys a screenful, `Ctrl+F` opens a find bar, `F3` finds the next place the selection appears, and `Ctrl` and a wheel changes that page's text size. Until that press the arrows walk the folder, and a press on the page's caption gives them back.

## System Tray Menu

See [docs/tray.md](docs/tray.md) for every row in the tray menu and what each one does.

## Detach Mode

A pinned window can be handed to a process of its own: **Detach Window** in the pin's right-click menu, in its own menu where the file is a sound's card, or on the button between **Always On Top** and **Minimize**. The window that comes up shows the same file at the same place and the same size, and it is completely independent from then on — a click, a hover, a key in Explorer, or `Tray → Pin Mode → Update Preview` never reaches it, because no Explorer watch of this app's is running in the process that draws it. Everything a pin does it does too: the caption and its buttons, the transport bar, previous/next through the folder, the bubble, the taskbar button, `Escape`, and its own **Detach Window** row, which hands it to yet another process.

You can have as many detached windows as you like, and they all live under the tray icon: closing the app closes every one of them, while closing one leaves the others and the tray app alone. Each is a process, so its film, its document engine and its browser are its own — the tray's `Engine` timers never touch them — and the taskbar shows them all under this app's single button, with a thumbnail per window.

Settings reach a detached window the way they reach a pin: what is written down (`config.ini`) is read by the window itself, so a level, a scale or a theme changes live, while what a window only reads when it opens a file — the engine a film plays through, the size its box was measured at — arrives at its next start, which is the answer a pin has always given.

## Configuration

See [docs/configuration.md](docs/configuration.md) for every `config.ini` setting, the file's default contents, and the names that are gone.

## Build from Source

Requirements: Windows 11, Rust 1.98.1+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK. The build targets the x86-64-v2 level (`-C target-cpu=x86-64-v2` in `.cargo/config.toml`), so the machine that compiles it has to be one that can run the result.

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
