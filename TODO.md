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

## Configuration

- The `*_BEFORE_*` lists `repair_older_lists` reads are there for files written before this build, and can go when a file that old can no longer be in use — a release or two away. Nothing else carries the history of a list, and the names this app once wrote are not read at all any more: a file that holds one loses the line and nothing else has to know about it.
