# TODO

## Features

- **File Format Support:**
  - Design & project files (`.psd`, `.xcf`, `.ai`)
- Add Office documents fallback using LibreOffice / soffice.
- Add preview integration for voidtools Everything search.
- Implement automatic updates.

## Configuration

- Retire what the stamp has made redundant, a release or two after it: the reads `apply_ini` does for a key that has been renamed (`off_trigger_key`, `svg_scale`, the backdrops' older names), and the first step of `migrate` together with the `*_BEFORE_*` lists it reads — no file that old can be in use by then, and a step is only worth keeping while one can.
- Tell a value the app wrote apart from one the user chose, so a default that moves can reach the installations that never touched it: every key is written today, so a changed default reaches fresh installations only. The stamp is what makes that possible — a file this build wrote can be told from one an older build wrote — but it is not what does it: what does it is not writing the values nobody chose.
