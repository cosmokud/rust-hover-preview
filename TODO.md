# TODO

## Features

- **File Format Support:**
  - More design & project files (`.xcf`, ...)
  - CAD files (`.dwg`, `.step`, `.stl`, ...)
  - 3D files (`.obj`, `.fbx`, `.gltf`, `.glb`, ...)
- Add preview integration for voidtools Everything search.

## Unsupported File Format

Two gaps. The first is a file that is shown only when its name is right, because nothing
in the file itself says what it is. The second is a file the app cannot show at all.
Beside each name is the one thing that would change that.

**Nothing in the file confirms what it is**

- **A container with no label inside it.** A zip (`.apk` `.cbz` `.jar` `.xpi` `.zipx`, and
  the `.fig` `.procreate` `.sketch` `.xd` design projects), an Office package (`.docm`
  `.dotm` `.dotx` `.potm` `.potx` `.ppsm` `.ppsx` `.pptm` `.xlsm` `.xlsb` `.xltm` `.xltx`),
  an iWork document (`.key` `.numbers` `.pages`), an older Office, Visio or Publisher file
  (`.dot` `.pot` `.pps` `.xlt` `.vsd` `.pub`) or a newer Visio one (`.vsdm` `.vsdx` `.vstx`),
  a StarOffice 5 document (`.sda` `.sdc` `.sdd` `.sdw`), a Pocket Word `.psw`, a Zoner
  `.zmf`, or a gzip stream (`.tar.gz` `.tgz` `.gnumeric` `.gnm` `.abw` `.zabw`). The bytes
  are the same whatever the box holds, so a renamed file cannot be sorted out.
  OpenDocuments, StarOffice XML documents and Krita projects are the exception: each writes
  its own type inside itself, and that is what they are read by.
- **A format whose marker sits at the end of the file, or past what is read.** `.tga` (its
  marker is the last 18 bytes), `.flm` (36 bytes from the end), a TiVo `.ty` or `.ty+` (its
  markers repeat every 128 KB). Only the first 4 KB of a file is read.
- **A format whose signature is its own text.** `.svg` `.svgz`, a flat OpenDocument
  (`.fodt` `.fodg` `.fodp`), a Visio `.vdx`. Text is left to the text preview: guessing a
  format from text would misread ordinary documents.
- **A format with no marker at all, or none FFmpeg reads.** `.mvi` `.mxg` `.psp` `.vw`, and
  `.cin`, whose FFmpeg demuxer is gone. These preview only when the name is right.
- **A marker too weak to trust.** `.mw` (MacWrite): its header is two bytes that any file
  could have, so a guess from them would mislabel ordinary files as documents.

**A file was offered to an engine that could not read it**

- `qxp` (older QuarkXPress), `pm3` `pm4` `pm5` (PageMaker before 6), `vssm` `vst` `vstm`
  `vtx` `vsx` (Visio stencils and templates), `epub` — LibreOffice was asked, answered
  nothing, and each cost a launch and showed nothing, so they are out of the `[libre]` list.
  What its filters do read are the neighbouring names `qxd` `qxt`, `pm` `p65` `pm6` `pmd`,
  and `vdx` `vsd` `vsdm` `vsdx` `vstx`.
- `agd` `fhd` `jtd` `jtt` `plt` `pxl` `rl` `sdp` `sgf` `sgl` `uof` `uop` `uos` `uot` `vor` —
  no LibreOffice filter claims them at all. A reader for one is a piece of work of its own
  (Ichitaro, the uniform office formats), with nothing here to test it against.
- `swf` — a Flash animation. The engine did not fail on it, it froze, so the name was taken
  out a step earlier and is played as a video instead.
- Any of them comes back with a reader, or an engine that has the filter: add the name to
  `[libre]` in `config.ini` on a machine whose LibreOffice reads it.
- The engine also reads names the list has never held — `qxd` `qxt` (QuarkXPress) and `p65`
  (PageMaker) — which could be added just as the list stands; what is missing is a file of
  one to settle that it converts.

## Unsupported ImageMagick File Format

Every name an installed engine's own registry carries — `magick -list format`, the same list the
built-in `[magick]` extensions were read out of one name at a time — and this app does not preview
through it: 163 of them, measured against ImageMagick 7.1.2 Q16-HDRI for Windows, which is the
build whose coders the `[magick]` list was settled on. Another build carries what its own coders
are built with, so `magick -list format` on the machine in hand is the question to ask there. None
of the names below is missing by accident: the groups are the reasons, and what each would take is
beside it. Adding a name to `[magick]` in `config.ini` is the whole change wherever the engine here
really reads the format; the rest is a reader, a delegate or a build feature. What the list *is*
made of is the camera raw formats and the pictures beside them, and the README says what a preview
of one of those is worth.

**A second name for a format this app already previews** — the engine is never asked for one of
these, and a file carrying one shows nothing although its bytes are a format this app reads:
`bmp2` `bmp3`, `gif87`, `jps` `pjpeg` `mpo`, `png8` `png00` `png24` `png32` `png48` `png64`, `tiff64`
`ptif` `group4`, `dxt1` `dxt5`, `icb` `vda` `vst`, `icn` `icon`, `pict`, `pcds`, `picon`, `sun`,
`farbfeld`. Most are the engine's own output selectors rather than names a file arrives under; the
ones a real file could carry are each one entry in the list that already holds the format — `mpo`
(the multi-picture JPEG a phone writes) in `[image]` beside `jpg`, `farbfeld` beside `ff`, and
`sun`/`pict` in `[libre]` beside `ras`/`pct`.

**A name that needs no engine at all** — `epsf` `epi` and `ept` `ept2` `ept3` are the other
spellings of the encapsulated PostScript our own reader already takes (each carries the metafile or
TIFF preview `eps_image` reads), so they are entries in `[vector]` beside `eps` and `epsi`; `avci`
is an AVC still in a HEIF container, which the Windows codec claims to read — one entry in
`[image]` beside `heic`, and one in `wic_image`'s own list of codec names; `pdfa` and `epdf` are
PDF spellings, and the PDF gate is the one name written in code rather than in a list, so they are
a line there.

**A font the specimen cannot draw** — `pfa` `pfb` (PostScript Type 1, ASCII and binary) and `dfont`
(a Macintosh suitcase). ImageMagick rasterizes all three through FreeType, while a specimen of ours
is text the browser engine draws in a font it is pointed at, and that engine reads none of them. A
preview of one would be the engine's own rendering of a face rather than our lines about it —
the shape of work a raw takes, not a list entry.

**A raw sample with no head to read** — `rgb` `rgba` `rgbo` `bgr` `bgra` `bgro` `gray` `graya`
`mono` `cmyk` `cmyka` `ycbcr` `ycbcra` `uyvy` `yuv` `pal` `bayer` `bayera` `rgb565` `map`. The
engine reads these only when it is told the size and the depth, which a hover has no way to know,
so a file of one of these names is one neither side can make a picture of.

**A document rather than a picture** — `ps` `pcl` `xps` (all drawn through Ghostscript), `djvu`
(through DjVuLibre), `gv` (through Graphviz — `dot`, its other spelling, is a name our office list
already claims as Word's template), and the engine's own page languages `mvg` `msl` `pocketmod`.
No delegate is bundled and this build declares only the `ps` one, so each of these is a
Ghostscript, a DjVuLibre or a Graphviz that has to be on the machine as well, and a page is what
the PDF path and the render engine are for.

**The engine's own notation** — names of images to *make*, and of things that are not files: `xc`
`canvas` `caption` `label` `gradient` `fractal` `plasma` `hald` `histogram` `pattern` `tile` `null`
`strimg` `vid` `inline` `data` `clipboard` `file` `http` `https` `ftp` `screenshot` `thumbnail`
`mask` `clip` `msvg` `rsvg` `dcraw`. A hover on a file called `https` is a hover on nothing.

**Read here, and never asked for** — a reader this build really has, on a format no picture preview
was ever wanted for. Each is one entry in `[magick]` away; what is missing is a reason and a file
to settle it against: `aai` (AAI Dune), `art` (PFS clip art), `ase` `aseprite` (Aseprite sprites —
the one of these a person may actually have), `cal` `cals` (CALS), `cut` (Dr Halo), `fax` `g3` `g4`
(fax bitstreams), `fl32` (FilmLight frames), `ftxt` (an image of formatted text), `hrz` (slow-scan
television), `ipl` (an IPL sequence), `jnx` (Garmin map tiles), `mac` (MacPaint), `mat` (MATLAB),
`mpc` (the engine's own pixel cache, meaningless without the `.cache` file beside it), `mtv`, `otb`
(the on-the-air bitmap), `palm`, `pes` (an embroidery machine's pattern), `pgx`, `phm` (the half
float variant of the PNM family, beside the `pfm` we do read), `pix` `rla` `rle` (Alias and Utah
RLE), `pwp` `sfw` (Seattle Film Works), `rgf` (a LEGO EV3 icon), `scr` (a ZX Spectrum screen), `sct`
(Scitex), `sf3`, `six` `sixel` (DEC terminal graphics), `stegano` (a picture with a message hidden
in it), `tim` `tm2` (PlayStation textures), `viff` `xv` (Khoros), `vips`, `wbinfo` (an Amiga icon),
`c2pa` (provenance metadata), `cube` (a colour lookup table), `pango` (a markup language).

**A name this build cannot read at all** — the registry carries it and the mode says no read, so
nothing can be asked of it until a delegate is installed or a build feature is turned on: `bie`
`jbg` `jbig` (JBIG), `flif`, `fpx` (FlashPix), `dmr`, `dps` (Display PostScript), `uhdr` (Ultra
HDR), and the write-only rest, which are files the engine *makes* rather than opens — `ashlar`
`brf` `cip` `eps2` `eps3` `info` `isobrl` `isobrl6` `kernel` `matte` `ps2` `ps3` `shtml` `ubrl`
`ubrl6` `uil`.

**And one the registry does not carry, because this build has no coder for it** — `xwd`, the X11
window dump. ImageMagick has the coder (registered as `_XWD`) and the Windows installers leave its
module out, so asking one of those builds for a `.xwd` fails looking for a module that is not
there, and `magick -list format` — which names the formats whose coders are installed — does not
name it at all. A build that ships it declares it, and then the name is one entry away like the
rest.

## Configuration

- The `*_BEFORE_*` lists `repair_older_lists` reads are there for files written before this build, and can go when a file that old can no longer be in use — a release or two away. Nothing else carries the history of a list, and the names this app once wrote are not read at all any more: a file that holds one loses the line and nothing else has to know about it.
