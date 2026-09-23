# TODO

## Features

- **File Format Support:**
  - More Design & project files (`.xcf`, ...)
  - CAD files (`.dxf`, `.dwg`, `.step`, `.stl`, ...)
  - 3D files (`.obj`, `.fbx`, `.gltf`, `.glb`, ...)
  - The `[libre]` names no filter of an installed LibreOffice declares, which the engine is
    no longer asked about: each was a launch that answered nothing when a file of that name
    was hovered. They were read out of the engine's own registry (`share/registry/*.xcd`),
    one name at a time, and each is kept here with what it was mistaken for:
    - `qxp` — older QuarkXPress. The filter reads `qxd` and `qxt`.
    - `pm3` `pm4` `pm5` — PageMaker before 6. The filter reads `pm`, `p65`, `pm6`, `pmd`.
    - `vssm` `vst` `vstm` `vtx` `vsx` — Visio stencils and templates. The filter reads `vdx`,
      `vsd`, `vsdm`, `vsdx`, `vstx`.
    - `epub` — the engine writes EPUB and does not read one; a preview needs an ebook reader
      of this app's own.
    - `agd` `fhd` `jtd` `jtt` `plt` `pxl` `rl` `sdp` `sgf` `sgl` `uof` `uop` `uos` `uot`
      `vor` — declared by no filter at all. A reader for one of them is a piece of work of
      its own (Ichitaro, the uniform office formats) with nothing here to test it against.
    - `swf` came out of the same list a step earlier, for the worst of the reasons: the
      engine does not fail on a Flash file, it spins — which is what the give-up in
      `libreoffice_render` now ends — and the name is FFmpeg's to play, in the video list.
    - What brings any of them back is a reader, or an engine that has the filter: a machine
      whose LibreOffice reads one can have the name back by adding it to the list.
  - And the other half of that pass, for completeness: the engine reads names the list has
    never held, because the list was written from the formats the engine *says* it supports
    rather than from its filters. QuarkXPress `qxd`/`qxt` and PageMaker `p65` are three of
    them, no other kind claims any of the three, and each could be added to `[libre]` as the
    list stands — what is missing is a file of that format to settle that it converts.
- Add preview integration for voidtools Everything search.

## Configuration

- The `*_BEFORE_*` lists `repair_older_lists` reads are there for files written before this build, and can go when a file that old can no longer be in use — a release or two away. Nothing else carries the history of a list, and the names this app once wrote are not read at all any more: a file that holds one loses the line and nothing else has to know about it.
