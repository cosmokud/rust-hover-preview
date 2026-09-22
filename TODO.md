# TODO

## Features

- **File Format Support:**
  - Design & project files (`.psd`, `.xcf`, `.ai`)
- Add Office documents fallback using LibreOffice / soffice.
- Add preview integration for voidtools Everything search.
- Implement automatic updates.

## Configuration

- Tell a value the app wrote apart from one the user chose, so a default that moves can reach the installations that never touched it: every key is written today, so a changed default reaches fresh installations only. What it takes is not writing the values nobody chose. The files that exist already hold all of them, so the change also has to decide once what to do with those, and to remember which files it has decided for.
- The older names and lists the repair knows about — `off_trigger_key`, `svg_background`, `svg_scale`, `svg_preview_enabled`, `avoid_filename`, `transparent_background`, and the `*_BEFORE_*` lists — are there for files written before this build. They can go when a file that old can no longer be in use, which is a release or two away; nothing else in the app reads an older name, so it goes in one piece.
