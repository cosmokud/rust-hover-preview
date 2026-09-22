# TODO

## Features

- **File Format Support:**
  - Design & project files (`.psd`, `.xcf`, `.ai`)
- Add Office documents fallback using LibreOffice / soffice.
- Add preview integration for voidtools Everything search.
- Implement automatic updates.

## Configuration

- Tell a value the app wrote apart from one the user chose, so a default that moves can reach the installations that never touched it: every key is written today, so a changed default reaches fresh installations only. What it takes is not writing the values nobody chose. The files that exist already hold all of them, so the change also has to decide once what to do with those, and to remember which files it has decided for.
- The `*_BEFORE_*` lists `repair_older_lists` reads are there for files written before this build, and can go when a file that old can no longer be in use — a release or two away. Nothing else carries the history of a list, and the names this app once wrote are not read at all any more: a file that holds one loses the line and nothing else has to know about it.
