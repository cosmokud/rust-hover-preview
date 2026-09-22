# TODO

## Features

- **File Format Support:**
  - Design & project files (`.psd`, `.xcf`, `.ai`)
- Add Office documents fallback using LibreOffice / soffice.
- Add preview integration for voidtools Everything search.
- Implement automatic updates.

## Configuration

- Stamp `config.ini` with a schema version and migrate an older file through ordered steps, so the legacy key and list fallbacks can be retired a release or two after the rename they cover — and so a file written by a newer version is left alone rather than rewritten without whatever it added.
- Tell a value the app wrote apart from one the user chose, so a default that moves reaches the installations that never touched it: every key is written today, so a changed default reaches fresh installations only — which is what the `(Default)` mark a menu derives says nothing about.
