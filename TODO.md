# TODO

## Features

- **File Format Support:**
  - Design & project files (`.psd`, `.xcf`, `.ai`)
- Add Office documents fallback using LibreOffice / soffice.
- Add preview integration for voidtools Everything search.
- Implement automatic updates.

## Configuration

- Delete the first step of `migrate` — `adopt_older_names` and `adopt_older_lists` with the `*_BEFORE_*` lists — a release or two after the stamp, when no file written before it can still be in use. It is the only thing in the app that knows a name or a list this app no longer writes, so it goes in one piece: nothing else reads an older name, and the lists it carries are otherwise dead weight.
- Tell a value the app wrote apart from one the user chose, so a default that moves can reach the installations that never touched it: every key is written today, so a changed default reaches fresh installations only. The stamp is what makes that possible — a file this build wrote can be told from one an older build wrote — but it is not what does it: what does it is not writing the values nobody chose.
