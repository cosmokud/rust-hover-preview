# TODO

## Features

- Add preview integration for voidtools Everything search.
- **File Format Support:**
  - More design & project files (`.xcf`, ...)
  - CAD files (`.dwg`, `.step`, `.stl`, ...)
  - 3D files (`.obj`, `.fbx`, `.gltf`, `.glb`, ...)
  - Audio files (`.mp3`, `.opus`, ...)

## Unsupported LibreOffice Format

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
through it: 80 of them, measured against ImageMagick 7.1.2 Q16-HDRI for Windows, which is the
build whose coders the `[magick]` list was settled on. Another build carries what its own coders
are built with, so `magick -list format` on the machine in hand is the question to ask there. None
of the names below is missing by accident: the groups are the reasons, and what each would take is
beside it. Adding a name to `[magick]` in `config.ini` is the whole change wherever the engine here
really reads the format; the rest is a reader, a delegate or a build feature.

What the list is made of is the camera raw formats and, beside them, the names the engine will read
when it is forced and the raw sample dumps. Every name the engine draws when forced is an entry
there — the second spelling of a format another list holds (`pict`, `sun`, `pcds`, `dxt1`, `dxt5`,
`group4`, `icb`, `vda`, `vst`, `picon`) as much as the formats no preview was ever asked for (an
`.hrz`, a `.fax`, a `.mac`, a `.wbinfo`) — and the dumps (`rgb`, `rgba`, `gray`, `cmyk` and their
kin) are read at a shape worked out from the file's own length, since a dump has no header to read
one out of. What is left below is what neither of those reaches.

**A second name for a format this app already previews** — the engine is never asked for one of
these, and a file carrying one shows nothing although its bytes are a format this app reads:
`bmp2` `bmp3`, `gif87`, `jps` `pjpeg` `mpo`, `png8` `png00` `png24` `png32` `png48` `png64`, `tiff64`
`ptif`, `icn` `icon`, `farbfeld`. Most are the engine's own output selectors rather than names a
file arrives under; the ones a real file could carry are each one entry in the list that already
holds the format — `mpo` (the multi-picture JPEG a phone writes) in `[image]` beside `jpg`, and
`farbfeld` beside `ff`.

**A font the specimen cannot draw** — `pfa` `pfb` (PostScript Type 1, ASCII and binary) and `dfont`
(a Macintosh suitcase). Both routes are closed, and each is closed by the thing the format would
have to be drawn with. A specimen of ours is text the browser engine draws in a font it is pointed
at, and every font a page asks for goes through that engine's own sanitizer — the OpenType
Sanitizer, which parses OpenType in its two shapes and the two WebFont containers and turns
everything else down — so a Type 1 face and a suitcase are files the specimen can be given and
cannot be drawn with: measured on the runtime this app uses, a `.ttf` loads and an 11 KB Windows
bitmap font (`.fon`) beside it is refused. ImageMagick is the other route and reads all three
through FreeType, but a font is the font preview's kind: what it would show is a rendering of a
face rather than our lines about it, which is the shape of work a raw takes rather than a list
entry. `.ttf`, `.otf`, `.ttc`, `.woff` and `.woff2` stay the whole of the font list for that
reason.

**A document rather than a picture** — `ps` `pcl` `xps` (all drawn through Ghostscript), `djvu`
(through DjVuLibre), `gv` (through Graphviz — `dot`, its other spelling, is a name our office list
already claims as Word's template), and the engine's own page languages `mvg` `msl` `pocketmod`.
The engine draws the first three through a delegate rather than itself, and none of the three
delegates is on the machine this was measured on — no `gswin64c`, no `dot`, no `ddjvu` — so each of
these answers nothing here today, before a list entry is even reached. A page is also what the PDF
path and the render engine are for, so a name here is a third way to draw one. (`mvg` is the one of
them that needs no delegate: forced it works, but only where the coder is named — `mvg:file.mvg` —
because a plain path is not recognized as its own format.)

**The engine's own notation** — names of images to _make_, and of things that are not files: `xc`
`canvas` `caption` `label` `gradient` `fractal` `plasma` `hald` `histogram` `pattern` `tile` `null`
`strimg` `vid` `inline` `data` `clipboard` `file` `http` `https` `ftp` `screenshot` `thumbnail`
`mask` `clip` `msvg` `rsvg` `dcraw`. A hover on a file called `https` is a hover on nothing.

**A name this build cannot read at all** — the registry carries it and the mode says no read, so
nothing can be asked of it until a delegate is installed or a build feature is turned on: `bie`
`jbg` `jbig` (JBIG), `flif`, `fpx` (FlashPix), `dmr`, `dps` (Display PostScript), `uhdr` (Ultra
HDR), and the write-only rest, which are files the engine _makes_ rather than opens — `ashlar`
`brf` `cip` `eps2` `eps3` `info` `isobrl` `isobrl6` `kernel` `matte` `ps2` `ps3` `shtml` `ubrl`
`ubrl6` `uil`. `xwd` belongs here too although no registry entry of this build lists it: the format
is documented, the coder for it is not among the modules the installer put down, so the name would
be asked for and would answer nothing.

**And one the registry does not carry, because this build has no coder for it** — `xwd`, the X11
window dump. ImageMagick has the coder (registered as `_XWD`) and the Windows installers leave its
module out, so asking one of those builds for a `.xwd` fails looking for a module that is not
there, and `magick -list format` — which names the formats whose coders are installed — does not
name it at all. A build that ships it declares it, and then the name is one entry away like the
rest.

## Unsupported PeaZip File Format

The archives PeaZip opens that nothing here shows. They are of two kinds, and neither is a name
missing from `[peazip]` by accident: a name put in that list is a name the engine is *asked*
about, and one it turns down costs a launch before the answer is remembered, so a name the
engine cannot read is left out on purpose and written down here instead.

Beside each name is the one thing that would change it.

**A format PeaZip handles with another of its own tools** — the engine this app drives is the
console archiver PeaZip carries, and these are the formats its other backends handle: `pea`
(PeaZip's own format), `arc` (FreeArc), `zpaq` and `paq`/`lpaq`, `upx`, and the codecs its build
carries no format for — `br` (Brotli), `lz4`, `lz5`, `lizard`, `flzma2`. Asking the engine for
one of them is answered with `Cannot open the file as archive`, measured against PeaZip 10.9.0
for a `.arc`, a `.zpaq` and a `.br`. What would show them:

- `arc`, `zpaq`: the backend tools PeaZip ships list their own archives — `Arc.exe l` and
  `zpaq.exe l` — each with a syntax and an output of its own, so each is a second and a third
  listing parser beside the one `archive_listing` reads. What they are asked for would then be
  settled per format, since a file is not one of theirs by its bytes alone.
- `pea`: there is no list command to drive. The PEA documentation says so — "there is no
  separate list/test command" — and what the format holds is written in object headers that its
  extractor walks. A reader here would be a walk of those headers (they are not compressed), or
  a preview built by extracting a copy to a temp folder, which is an extraction this app makes
  for nothing else.
- `br`, `lz4`, `lz5`, `lizard`, `flzma2`: single-stream codecs. PeaZip's `brotli.exe` and
  `zstd.exe` compress and decompress one file at a time, so a preview of a `.br` is a preview of
  one member rather than of a container — the same shape the `.gz`, `.bz2`, `.xz` and `.zst`
  entries have, and a second engine rather than a list entry. (`zstd` is in the list already:
  PeaZip's build declares it as a *format* as well as a codec, which is the difference.)

**And a format whose marker a hover cannot reach.** An `.iso` and a `.udf` say what they are
thirty-two kilobytes into the file — `CD001` at 32769, a UDF descriptor at 32768 — and what a
hover reads of any file is its first four kilobytes, so neither can be confirmed by its own
bytes. Both preview by name, through the engine, and a renamed one is a file nothing recognizes
rather than a file shown wrongly. `lzma` and a split archive's `001` are here for the plainer
version of the same thing: neither has a marker at all, so both are the name's business, and
that is the whole of what can be done with them.

What is *not* here is the names the engine reads and this app leaves out on purpose — the
programs it lists as resources (`exe`, `dll`, `sys`, `obj`, `elf`, `macho`), the names that are
words rather than formats (`img`, `ext`, `fat`, `mbr`, `gpt`), and the index files beside a
compiled help file. Each of those *could* be shown, and each is left out for a reason that is
written where the list is (`peazip_formats`), because a preview nobody asked for is worse than
none.

## Configuration

- The `*_BEFORE_*` lists `repair_older_lists` reads are there for files written before this build, and can go when a file that old can no longer be in use — a release or two away. Nothing else carries the history of a list, and the names this app once wrote are not read at all any more: a file that holds one loses the line and nothing else has to know about it.
