# Changelog

## [0.3.5] - 2026-09-26

### Added

- **`Cache → Image (Disk)`**: a budget for the pictures the image converter develops. Each is kept as a file under `%TEMP%\rust-hover-preview\image` (`image_disk_cache_mb`, `512` by default), so hovering a camera raw a second time — or again after a restart — is a read rather than another conversion. `0` keeps nothing between hovers.
- **`Reset to Recommended Settings`** and **`Reset Extension Lists`**: two rows inside the tray's `Config.ini` item, each asking first and each offered only while there is something to put back. The first returns every setting to what this build recommends and leaves the extension lists alone; the second returns the lists and leaves every other setting alone.
- The installer asks the same two questions on a page of its own, with both boxes clear and neither one shown to a silent install — so an update taken with `Auto` changes nothing. What a box asks for is applied the next time the app starts.


### Fixed

- **A hover no longer stalls on the files a hover is worst at.** A box measured by a read — a PDF's first page, an archive's table of contents, a comic's first plate, an SVG, a metafile, a font — was measured on the thread that draws the hover, so a large one held the pointer, the spinner and every tick of the loop for as long as the read took, and a `.tar.gz` inflated to reach its listing or a book of a thousand pages was felt at the hand. Those measures run on a thread of their own now: the hover is laid out as the spinner and replayed the moment the answer lands, which is the arrangement a video's geometry probe has had for its own two processes, and the answer is held for the file, its version and the room it was measured in, so a second hover of a file waits for nothing. An answer that the reader has nothing for the file is the one answer a wait comes down on rather than a spinner left standing over nothing.
- **Repainting a video or an animation at the size of a display no longer costs a pass over every one of its pixels.** A frame whose own pixels are opaque everywhere is copied into the layered window's surface rather than composited pixel by pixel: where the alpha is 255 both the blend over a backdrop and the premultiply a transparent one asks for are the pixel's own bytes, in every backdrop kind, so what a frame of a video costs is a copy rather than eight million divisions per channel. Whether a frame is opaque is asked where the pixels are made — off the thread that draws them, once per frame rather than once per repaint — and the corner spinner a streaming preview draws is drawn into the box it sits in rather than into a copy of the whole frame, so an overlay repaint copies a few kilobytes instead of thirty megabytes.
- **A document's page is no longer stamped, and the page cache is no longer walked.** The order pages are given up in is kept beside them rather than written onto each one as it is read, and a trim works from that table instead of listing the folder and reading a timestamp per file: what a hover on a document used to cost was a write to the volume for every read of a page, a walk of the whole cache every time a hover ended, and three lookups of the page by name where the table answers with one.
- **A hover asks the volume about a file once instead of five or six times.** Whether a file is there, that it is a file, what version it is at and whether its content is on this machine are all in one directory entry, and the gate read that entry once per question — for existence, for the key the content answer is held under, for the cloud check, and again on either side of the path being normalized and canonicalized. One reading of the entry now answers all of them, which is what a pointer swept down a OneDrive or network folder is paid for: one read of a file's entry rather than six, with the file itself opened once as before.
- **The folder behind a view is read from the view the pointer is in, rather than from every open window and tab.** What the hook compares to notice a place has changed — the folder a view has open and the URL it was opened with — was asked of the whole Shell window collection on every folder probe: up to 64 registrations, six or seven crossings into the shell apiece, and the window it settled on was whichever tab of a frame answered with the frame's window first, which for a window holding tabs is not the tab the pointer is in. It now asks the frame's own set of views, kept per window, and the one view the point is drawn inside — the tab that is showing — so a probe costs a description of one view rather than a walk of the desktop however many tabs are open, and the folder it reads is the one the pointer is actually over. A pointer over the navigation pane, the toolbar or the details pane is answered by the view it was last inside of that frame, the tab the hand was last working in, which is a guess that at least does not move while the pointer does not. The fresh search-view check the miss path makes asks the same view, so the two cannot disagree about which view is being asked.
- **The keyboard is polled in one pass instead of two.** The navigation keys and the shortcut keys overlap — the arrows are navigation keys and half of an Alt+arrow — and a key read twice in a tick is not the same answer twice, because the press bit it reports is consumed by whoever reads it first. The two passes are now one and the modifier state is read once for both, which is three fewer key queries per tick — twenty-two down to nineteen; the shortcut keys that are *not* navigation keys are still read every tick, because a Backspace navigates with no modifier at all and a `T` typed anywhere sets a press bit that a read gated on Ctrl would leave standing until the next Ctrl was pressed. The trigger key's name is resolved when it is spelled differently rather than once per tick, which takes an allocation and a lookup off every tick.
- **The pointer is read once per tick.** Where it is, the scale of the display it is on and the window under it are one reading, and the item under the cursor, the place, the move threshold, the wheel's own question and the preview's hold are all answered from it rather than each reading the cursor again; the display's scale is kept per display, so a pointer that has not crossed to another one is not asking for a number that cannot have moved. The one read that stays fresh on purpose is the preview loop's own, taken a step before a hover is laid out, because that one is about where the preview goes.
- **The tray wakes for its own messages rather than a hundred times a second.** Its loop polled the message queue every ten milliseconds and slept in between; a window's messages are queued whether or not anyone looks, so the loop now blocks on `GetMessageW` — every wake it has is a message: the icon, the menu, Explorer restarting, the system coming back from sleep, and the quit the exit paths post. And the thread that keeps a player's window from taking focus now waits on an event between videos instead of waking every eighty milliseconds to re-assert a window no player is holding: the hand-over a new player makes signals it, and the timeout behind that wait is the cadence it keeps the window in step at while one is playing.
- **A hover on a video can no longer end in a spinner that never goes away.** Measuring a video is two external processes — `ffprobe` for the file's shape, and an `ffmpeg` pass over forty-eight decoded frames for its letterboxing — and neither of them was given a deadline: a corrupt, truncated or merely enormous file could hold the wait for as long as it liked, because the hover waiting on a probe is not waiting on an engine and the cap a page is waited for under was not standing behind it, so the spinner sat at the pointer until it was moved. Each of the two is now given ten seconds and is *killed* when it outruns that, which is what stops a detector that cannot keep up with a file from reading it to the end for a hover that has given up. What a file slower than the deadline is answered with is a video placed without its crop rather than a video that does not appear — the shape is the probe's and the crop is the extra — and a file neither process says anything about is remembered as unmeasurable, so it costs one wait and not a wait per hover. The wait is also answered **whatever became of the probe**: a probe that panicked used to unwind past the message that ends the hover, which left the spinner standing over a file nothing was still reading, and the answer is now sent through a guard that catches that.
- **The file that answers for a run's engines is written when something starts, not when something ends.** `%LOCALAPPDATA%\rust-hover-preview\engines\<pid>.state` names the engines, the browser and the video player a run has started, so that a run which is killed or crashes has what it left behind ended by the next one — and it was rewritten whole on every engine started *and* every engine stopped, which made a hover on a video two writes where only the first one buys anything. It is now written when a process is taken on and not when one is let go of: a line naming a process that has ended is inert by construction, because the reaper matches the image and the moment the process started before it touches anything, and the next start rewrites the file whole. What is left is one small write at the moment an engine or a player is actually started, and none when one is stopped.
- **`Avoid Filename` keeps a `Details` row's name off to the end of what is drawn — a name the column cut included.** The region a preview was moved clear of was cut to the width of the name the item *goes by*, measured with the shell's icon font, while a name too long for its column is *drawn* cut to that column: the two are not the same width, and the box came out shorter than the text on screen, so the preview — placed a gap past the box — landed in the middle of the name it was moved beside. The label the piece the name is drawn in reports for itself is read beside the boxes the walk already reads, one of the same batched properties and so no call of its own, and the wider of the two names is what the box is cut to, from the same font at the same display's scale. A name with room to spare measures the same either way and keeps the region it always had; a name that fills its column now comes out at the room itself — the region **`Avoid Filename Column`** keeps, and there is no tail of that room left to place a preview in anyway. Keyboard previews are unchanged, since what a keyboard placement is read from is the item alone, with no pointer on it to say which part of it the hand is at.
- **The launch no longer waits for the housekeeping.** The start cleared away everything earlier versions had left behind — the folders they cached their pages in, the log one of them wrote for every video hover, the installer an older update check had fetched, and the browser profiles of runs that are gone — before it created the tray window, so what those sweeps cost grew with what those versions had left in the temp folder and came straight off the time between the launch and the icon. They now run on a thread of their own, started beside the preview, hook, wheel and config threads rather than in front of them, and the two sweeps that *end processes* stay where they were: those are what stop a leftover engine being joined by a duplicate, so they still run before the hook does. What the move can cost is a page rendered again in the first moments — the temp folder being swept is the tree the page cache writes in, and the cache drops an entry whose file is gone — and what it buys is a tray icon that is a window and an icon and nothing else. A start that is being measured can say so: `RHP_STARTUP_TRACE` writes every step of it, the window and the icon included, to `%TEMP%\rhp-startup-trace.log`.

### Changed

- Version bumped to `0.3.5` in `Cargo.toml` and `Cargo.lock`.
- `image_cache_mb` starts at `64` rather than `32`, so a picture shown at the size of a large display is held at all; `document_cache_mb` starts at `256` rather than `128`, so one converted drawing or scanned book no longer fills the folder on its own. An installation that already has a `config.ini` keeps the value it wrote — only a new file gets the new defaults.
- A workbook's fallback page — the picture Excel is asked for where no printer can export a page — is written as a PNG rather than a BMP: the same picture, a few times smaller, and cheaper to read back.
- The render engine is asked for as soon as the pointer settles on a document it draws, instead of when the page is asked for, so the first hover of a session overlaps the launch with its own wait. Office is not warmed — its tier can be asked for a document and not for an application — and the ebook engine keeps no instance to warm.

## [0.3.4] - 2026-09-25

### Added

- **`Calibre` previews** for the ebooks nothing here opens: the Kindle and Mobipocket families (`azw`, `azw3`, `azw4`, `mobi`, `prc`), the open `epub`, the `fb2`, the scanned `djvu`, the Sony `lrf`, the Microsoft Reader book (`lit`), and the `pml`, `snb` and `tcr` of the dedicated readers. An installed Calibre converts the book and the app draws the first page of the PDF it wrote, exactly as it draws a PDF: the same `Ebook` kind, the same `Ebook Scaling` and backdrop, the same reader. Nothing is bundled, the Calibre window is never opened, a conversion leaves nothing behind in your folders, and a second hover of the same book starts no conversion at all.
- **`Calibre`** appears in **`Codecs → Engines`** to show if it is installed, and the conversion runs without a console window.
- `[calibre] extensions` in `config.ini` controls which formats are asked about; an installation that already exists is given the section and the built-in list on its next run, and a name the engine cannot read is remembered so it costs one conversion and never another.
- Books are recognized by their own bytes as well as by their names, so a renamed one still previews: a Mobipocket file — an `.azw`, `.azw3`, `.azw4`, `.mobi` or `.prc` — by the two identifiers every one of them carries, an `.epub` by the type it declares inside itself, a `.fb2` by its root element, a `.djvu` by the chunk its format opens with, and a `.lrf` by the letters it writes with a zero byte between them. A file whose bytes are another kind is still shown as what it is.
- **`Comic` previews**: hovering a `.cbz`, `.cbr` or `.cbc` now shows the comic's **first page** instead of a page of contents. The plate is read out of the container by the app itself — the archive's headers are walked for the plate's name and only that one plate is inflated, decoded and drawn — so a hundred-megabyte comic costs a page rather than an unpacking, and nothing needs to be installed. The order is the one comic readers use: the first plate by **name, counting the digits** (`2.jpg` before `10.jpg`), skipping anything that is not a picture (`ComicInfo.xml`, a `readme.txt`) and anything a Macintosh left behind (`__MACOSX/…`). A two-page spread is shown whole rather than cut in half.
- **A new `[ebook] extensions` section** in `config.ini`, which is the list of what the app draws as a book: `pdf` and the two spellings the PDF world writes beside it (`pdfa`, `epdf`), and the three comic containers. A name you add is a comic if the file it names is a container of pictures, and the PDF reader is asked about the PDF's own spellings — so the one list answers one question: what is a book.
- Comics are previewed at the **book** kind's scale and over the book kind's backdrop, under the same **`Ebook`** switch a PDF answers to, and `ebook_scale` sizes them: what a hover on a comic shows is a page, the same as a PDF.

### Changed

- Version bumped to `0.3.4` in `Cargo.toml` and `Cargo.lock`.
- **`cbz` moved out of `[archive]`**, and with it the page of contents a comic used to be shown as: a `.cbz` is the third of the comic containers and is now previewed like the other two. `lit` moved out of `[peazip]` and into `[calibre]` for the same reason — one name, one answer. `chm` was moved the same way and put back: the engine draws a real page for a help file and takes two to three seconds to do it, which is not a wait a file a pointer crosses on its way somewhere else should cost, so it is the archiver's name again and a hover on one is the listing that is there immediately. An existing `config.ini` is brought up to all of that on the next run, whichever of the lists it happens to hold. A name you want in another list is a line you add, and the lists are asked in the order they were always asked in.
- There is **no `Calibre TTL`** in the `Engine` submenu, and that is the engine's own answer rather than an omission. `ebook-convert.exe` is a converter — handed a book it writes a PDF of it and exits, booting a whole Python application to do it — so there is no instance to hold open between books and nothing an idle time could bound. It is the engine a TTL would suit best on paper, since a launch costs seconds, and the one engine with nothing to hold: what a second hover of the same book costs is a read of the page it already converted.
- A converted book's page is kept in the same page cache every other engine's page is kept in — `document_cache_mb`, `128` by default — named for the book, the version of it and the engine that wrote it, so a page one engine drew is never handed back as another's work. What is read back is bounded by the decode budget, so a scan converted into a PDF of gigabytes is answered with no preview rather than with the read.
- The page cache now keys a page by the name of the engine that drew it rather than by one of the two Office choices, which is what lets an engine that is neither of them keep its pages beside theirs.
- `TODO.md` gains two sections. **`Unsupported Calibre File Format`**: every ebook format the engine reads or refuses that is not in `[calibre]`, grouped by reason, with what each one would take — the names another kind already answers (`docx`, `odt`, `pdb`, `html`, `rtf`, `txt`, `pdf`), the comics this app reads itself, the formats the engine does not read at all (`tpz`, `kfx`, `recipe`), the spelling its own plugin does not declare (`lrx`), and the one thing no list can fix: a DRM-protected book is a book no engine here can convert. And **`Unsupported Comic Formats`**: the plates written in a format only the codec Windows has decodes (`webp`, `avif` — a plate is decoded from bytes out of a member and those decoders are handed a path), the comic containers this reader does not open (`cb7`, `cbt`), the `.cbc` whose comics are boxes inside it rather than pages, and a first plate that is a colour.
- **`Confirm File Type` is gone, and checking a file's own bytes against its name is not optional any more.** The tray item under `Performance` and the `confirm_file_type` key in `config.ini` are removed, so every hover asks a file's own bytes what it is before its name is asked — which is what the setting was already doing for a default installation. An installation that had turned it off starts confirming every file; one that had it on sees no change in what is previewed.
- **A still picture and one that moves are told apart by the file's own bytes rather than by its name.** A `.gif`, a `.webp` and a `.png` each cover both, and what says which is the file's own structure — its frame blocks, its `ANIM` chunk, its `acTL` chunk — which is read out of the head of the file without decoding anything and with no setting behind it. A still `.gif` is no longer handed to the animated reader first and then decoded a second time by the still path, which is a decoded frame it used to spend on every hover; a `.webp` that does not animate is no longer read whole into libwebp to find that out; and a file whose name says it is not an animation at all — a `.docx` holding an animated GIF — is played as the animation it is. The same read tells an animated `.avif`, `.heic` or `.jxl` from a still one, so a sequence is placed at the scale animations are given rather than at the picture's.
- The head of a file is read once and shared. `head` owns the probe — the front of a file first, sixteen bytes, and the whole four-kilobyte window only where that front settles nothing — and holds what the probe answered with, so the kind of a file, the reader chosen for it and the box it is placed at are all answered from one read rather than each opening the file for itself.

### Fixed

- **A preview no longer stays on screen when the pointer is swept over it at speed.** A preview that is still loading stands in as the spinner, and the spinner is placed at the hand and follows it — so a pointer crossing off the file mid-wait was read as a pointer still waiting for it. The load that would have answered the wait was then dropped as belonging to a hover the pointer had left, and nothing took the spinner down with it: the window stayed at the hand, the hold it had published kept the pointer from dismissing it, and it sat there until the pointer happened to move off it. A wait now holds the pointer through the *item* it is waiting for and not through the box its spinner occupies, so a pointer that has moved on takes the wait down as it goes; and a load whose hover has gone takes its spinner, its hold and anything it had asked for down with it.
- **A mouse hover is placed for the hand that actually asked for it.** The point a hover laid its box out from was the one the Explorer hook sampled at the top of its own tick — ahead of the walk through the shell that resolves the file and of the look for the `Avoid` region the layout is placed by — so the better part of a frame had passed by the time anything was laid out from it, which at a fast hand is dozens of pixels. The one thing the layout does with the point is keep the box clear of it, and that is worth something only while the point is still the hand's: a mouse hover is now anchored to a fresh read of the pointer, taken one step before the layout, and a read that fails keeps the point the message carried.
- **A preview no longer arrives under the pointer.** The box a frame was painted in was the one its layout came out at when the hover was asked for, and the hand had the whole load to travel into it — so the preview appeared under the cursor and the touch rule took it down again at the next tick, which is a spawn and a dismissal the eye reads as one event. The frame is placed once more from the pointer as it is, one read before the paint, which is the placement every tick of a wait already makes (see `PendingLoad::follow_pointer`), so what lands is the preview of the file under the hand rather than one under the hand itself. A keyboard hover carries no placement of its own and is left where the item put it.
- **The waiting spinner is kept off the pointer.** The arc is drawn in the preview window — the window the pointer's own messages land on — and it was placed a single pixel off the cursor, so a hand waiting on a file had the wait sitting under it and had to click and probe through it. What a wait is owed is the hand's own corner and not the hand itself: the arc is now a pointer gap off the cursor, the same gap every other preview keeps, in whichever of the display's four corners has room for it, and the box a document waiting on a page is placed in moves with it. A wait that has a page or a probe on the way is placed by the same rule, so the pointer is never inside it to begin with.
- **A book whose first page says nothing is previewed from the page behind it.** A cover is what an EPub and a Kindle file both put first, and a cover is very often one flat colour — the quick start guide Calibre ships inside its own installation carries a cover that is a single pixel, stretched over a whole page — so hovering one of those books showed a rectangle of that colour, which is what the book's first page *is* and says nothing whatever about the book. What a converted book is previewed from is now the first of its opening pages that holds more than one colour, which is the cover where there is one and the title page or the first page of text where there is not. A page with anything at all on it — a line, a photograph, a page of text — is never skipped, and a book whose opening pages are all one colour is still shown from its first page.
- A `.pdb` that is a **Mobipocket book** is no longer handed to the render engine, which cannot read one. The two identifiers the format is defined by are read off the file's own header, and a book of that kind is converted by the engine that reads books. A `.pdb` that is the other thing the name means — an AportisDoc, or a compiler's program database — is answered exactly as it was.

## [0.3.3] - 2026-09-24

### Added

- **`Engine → AFK Timer`**: an engine that is not marked `Persistent` is let go once no Explorer window has been reachable for the time you pick — 1 hour down to 15 seconds, 1 minute by default.
- A **`Persistent`** toggle at the top of each **`… TTL`** submenu: on, the engine is kept for its TTL whatever you are doing, which is how it worked before; off, which is how they all start, the AFK timer bounds it.
- New `config.ini` keys: `afk_timer_seconds`, `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent`.
- **`PeaZip` previews** for the archives nothing here reads — `cab`, `iso`, `udf`, `wim`, `msi`, `deb`, `rpm`, `arj`, `lzh`, `hfs`, `vhd`, `dmg` and the rest of the `[peazip]` list. An installed PeaZip lists them and the app shows the same page of contents a `.zip` is shown as; nothing is bundled, the PeaZip window is never opened, and a second hover of the same archive starts nothing. A single-stream `gz`, `bz2`, `xz`, `zst` or `z` is shown as the one member it holds.
- A **`Peazip`** switch in the tray **`Preview Types`** submenu, below **`Magick`**; `[peazip] extensions` in `config.ini` controls which formats are asked about, and a name the engine cannot list is remembered so it costs one launch and never another.
- **`PeaZip`** appears in **`Codecs → Engines`** to show if it is installed.
- Archives are recognized by content as well as by name, so a renamed one still previews: a `.cab` renamed to `.dat` is listed all the same. The formats PeaZip opens through its other tools — `pea`, `arc`, `zpaq`, `br` and the codecs its build carries no format for — are written down with their reasons in `TODO.md`.
- **PeaZip previews now drive the rest of the tools PeaZip ships**, each for the format only it reads: FreeArc’s archiver for `arc`, zpaq for `zpaq`, and Zstandard’s own tool for `zst`, which reports how large the stream was before it was compressed where the 7-Zip console leaves that blank. `arc`, `bcm`, `br`, `lpaq8` and `zpaq` join `[peazip] extensions`, and an existing `config.ini` is brought up to the list on the next run.
- `br`, `bcm` and `lpaq8` — single-stream compressors whose tools print nothing about what is inside — preview as the one member an extraction would write, named after the file, with nothing started at all.

### Changed

- Version bumped to `0.3.3` in `Cargo.toml` and `Cargo.lock`.
- The update installer is downloaded only after you click the update row and confirm, not as soon as a newer release is found; a check now costs one small request.
- The update check now runs at every startup as well as when the tray menu is opened, and the hour between checks is counted in memory rather than written to `%LOCALAPPDATA%`.
- The update prompt now asks with three answers: **`Auto`** installs the update as before, **`Manual`** opens the release page in your browser, and **`Cancel`** does nothing.
- **The `Cache` submenu has two entries** where it had five: **`Image (RAM)`**, the decoded frames held between hovers, and **`Document (Disk)`**, the pages an engine drew. **Text** and **Ebook** previews are no longer cached at all — a painted text frame and a rasterized PDF page are the cheapest things here to make again and the dearest to hold, so a text hover keeps the parsed document and its styled lines, and a PDF hover keeps only the page's size. What either costs is a layout, a raster and an encode rather than the memory a screenful of pixels takes.
- **The pages both document engines draw are one cache, on disk.** `document_cache_mb` — `128` by default, `0`–`2048` — replaces `office_cache_mb` and `libre_cache_mb`, and a `config.ini` holding either older key is read once for the larger of the two and written without them. The pages live under `%TEMP%\rust-hover-preview\document`, named for the document, the version of it and the engine that drew it, so a page outlives the run that drew it and the engine it was drawn by — and so a page one engine drew is never handed back as the other's work. What is given up first is the page that has not been _read_ for longest, not the one converted longest ago, and the page a hover is waiting for is never given up: at `0` a page is kept from the moment it is drawn until the hover that asked for it ends.
- The LibreOffice engine's own profile and the stub document it holds open moved out of `%APPDATA%\rust-hover-preview\rendered` — a folder that is deleted at startup now, along with everything else an earlier version cached — to `%LOCALAPPDATA%\rust-hover-preview\libreoffice`, beside the app's other engine state.

### Fixed

- A restart no longer hides an available update until the hour is up.
- The update row appears whether or not the installer could be fetched, and a download that fails says so instead of leaving no row at all.
- An installation made from a pre-release is offered the stable release its version names, instead of never being offered an update again.
- A preview of an archive PeaZip listed was drawn at the size of the screen’s free room instead of the size the layout planned, so a `.cab` or an `.iso` came up as a page stretched to the display rather than as the page a `.zip` of the same kind comes up as.
- Hovering a file whose preview an engine makes — an Office document, a `cdr` the render engine draws, a camera raw, an archive PeaZip lists — no longer leaves the spinner up for good when the file is left and taken up again while that engine is still working. The answer the engine had already produced was read as belonging to a hover that had gone, because it named the hover before the one waiting: the wait under the spinner was cleared with it, and the page, picture or listing in hand was not shown until the file was hovered again. An answer for the file a hover is waiting on is now that wait’s own, and a hover waiting on an engine is bounded by the same timeout whether or not the request it made is the one being watched for.
- The `Cache → Libre` budget is evicted by when a page was last read rather than when it was converted, which is what "the oldest first" was always meant to be.
- A document an engine would not draw is no longer refused for good: the mark ages out after two minutes, so a document that was locked, half-copied, or read while a filter was still being installed is asked about again.
- The tray's `Cache` submenu no longer says everything it sizes is held in memory and nowhere else, which stopped being true when converted pages landed on disk.

## [0.3.2] - 2026-09-24

### Added

- **`ImageMagick` previews** for camera raws (`nef`, `cr2`, `cr3`, `arw`, `dng`, `raf`, etc.) and niche images (`xcf`, `sgi`, `jp2`, `dpx`, `fits`, `dcm`, etc.). Uses your installed copy or a portable copy beside `config.ini`; nothing is bundled.
- A **`Magick`** switch in the tray **`Preview Types`** submenu, below **`Libre`**; `[magick] extensions` in `config.ini` controls which formats are asked about and remembers unsupported names.
- **`ImageMagick`** appears in **`Codecs → Engines`** to show if it is installed.
- Raw files are recognized by content and name, so renamed raws still preview. Scaling, background, and rotation work; conversions leave no files; hover frames use `image_cache_mb`; no `ImageMagick TTL`; no console window.
- More formats/spellings work: `pict`/`pct`, `sun`/`ras`, `pcds`/`pcd`, `dxt1`/`dxt5`/`dds`, raw sample dumps (`rgb`, `rgba`, `bgr`, `gray`, `cmyk`, etc.), `[vector]` (`epsf`, `epi`, `ept`, `ept2`, `ept3`), `[image]` (`avci`), and PDF gate (`pdfa`/`epdf`). Exclusions are explained in the module and `TODO.md`.
- **`Performance → Tick`** / `tick_ms` now defaults to **15 ms** (up to **78 ms**), changed from **30 ms**, so Explorer changes are noticed sooner. Larger values are lighter on CPU.

### Changed

- Version bumped to `0.3.2` in `Cargo.toml` and `Cargo.lock`.

### Fixed

- `[libre]` in `config.ini` is now read, and **`Preview Types → Libre`** works correctly. New built-in format lists also reach existing installations.
- Previews no longer blink, vanish, reappear, stay after the pointer leaves, or cover the pointer. This works at any `Timing → Delay`, including `0 ms`.
- Fewer false “pointer left” events; empty Explorer reads and different place names no longer cause blinking. Document/specimen previews now swap cleanly.
- Stuck spinners are cleared. Waits for documents, specimens, browsers, and the engine are bounded, so a stopped browser or preview loop cannot block or hold the app open.
- A document or specimen closes when the pointer touches it, even inside child windows. Held previews are released when their window is hidden, so old holds cannot block closing.
- A preview now closes when the pointer moves to an item with no preview, such as an application, folder, or unknown name. A failed read at the same item keeps the preview and asks again.
- A preview no longer stays on screen when the pointer leaves a **clipped item** — the last row of a scrolled `Extra large icons` or `Large icons` view — for the toolbar above it or past the bottom edge of the window. An item the view cuts off is reported with the whole box its content is drawn in, which carries on behind the search bar and past the window’s edge, and that box was what said the pointer was still on the item. The box is now kept to the part of it the view shows, so the preview goes as soon as the pointer is off the item — including over the toolbar and below the window.
- **Camera raws no longer preview at a fraction of their size.** **`ImageMagick`** now gets the display’s available room, so raws fill the preview like normal pictures.
- **`Avoid Nothing`** now places keyboard previews like **`Avoid Filename`**, not like **`Avoid Details`**; mouse previews are unchanged.
- The engine path is no longer polled; newer requests wake it immediately, and replaced navigations are dropped. An unresponsive Explorer no longer keeps the app open.

## [0.3.1] - 2026-09-24

### Added

- Third-party attribution. [`THIRD-PARTY.md`](THIRD-PARTY.md) lists every crate the app links, grouped by licence, with each one's copyright holders and a link to the text that applies; [`LICENSES/`](LICENSES) holds those texts. Both are generated from the dependency tree by `cargo tribute` and refreshed by `generate-attribution.ps1`. Nothing about the app itself changes.
- `tribute.toml`, which is what that generator reads: the licences this project accepts, and the entries for the two dependencies whose own declarations do not describe what they ship — `onig_sys`, which declares MIT for the Oniguruma sources it vendors and compiles in, and `unrar_sys`, which declares MIT for the UnRAR sources README already carries the terms of.

### Changed

- The README's License section points at `THIRD-PARTY.md` instead of carrying the UnRAR terms alone.
- Bump version to 0.3.1 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A preview no longer blinks on hover — closing and coming straight back. A frame that landed just after the pointer left is dropped instead of putting the preview up again.

## [0.3.0] - 2026-09-23

### Added

- Opening the tray menu can check for updates at most once an hour. If a newer version exists, the installer downloads and a row appears above Run at Startup. Clicking it asks whether to install; yes installs quietly and restarts the app. No check happens unless the menu is opened.
- `Timing → Settling Delay`: how long the pointer must be still before a preview may open for anything, 0 ms at the top down to 1000 ms. 0 ms by default, which is the requirement switched off. Written to `config.ini` as `settling_delay_ms`.
- `Engine → Select Engine → Office`: which engine an Office document's page is asked of. **Microsoft Office** is the default and asks the application that owns the format, keeping LibreOffice as the fallback for a family this machine has no application for. **LibreOffice** asks the render engine for every Office document, whether Microsoft Office is installed or not — useful where the application draws a page badly, or where every document should come out of one engine. Written to `config.ini` as `office_engine`; the LibreOffice row is greyed out where no LibreOffice is installed, and the app keeps drawing with Office until one is.
- `Engine → LibreOffice TTL`: how long the LibreOffice engine is kept after the last document it converted — Indefinitely, 1 hour, 30 minutes, 10 minutes, 5 minutes, 1 minute, 0 seconds, and 10 minutes by default. The next document is handed to an engine that is already up rather than paying for another start: measured on one document, 1.2 s against 0.2 s. Written to `config.ini` as `libreoffice_idle`; the row is greyed out where no LibreOffice is installed, and `0 seconds` keeps no engine at all. While one is kept it is a LibreOffice with a small document of this app's own open, which costs a few hundred megabytes, and the app ends it when the time is up.

### Changed

- `Performance → Office Engine TTL` is now `Microsoft Office TTL`, and `SVG Engine TTL` is now `WebView2 TTL`: the same settings under the names of the engines they are about. No `config.ini` key changed and no value moved.
- `Select Engine` and the three engine TTLs sit in a new `Engine` section below `Performance`, which keeps `Confirm File Type`, `Cache` and `Decode Budget`.
- `config.ini` is grouped the same way, under an `Engine` heading below `Performance`. A file that is not grouped the way this build groups it — one written before the `Engine` heading existed, or one whose headings are in the order the menus were in before — is written again the next time the app reads it, and a file already written this way is left as it is. A heading is a comment: no key changes and no value moves.
- New Libre preview type for documents LibreOffice can read but this app cannot, such as CorelDRAW `cdr` and older office formats. Previews are drawn by an installed LibreOffice, stay sharp when resized, and use new scaling and a 32 MB cache stored beside `config.ini`.
- CorelDRAW `cdr` previews no longer use the small embedded image. Without LibreOffice installed, `cdr` files show no preview. `cdr` moved to the `[libre]` list; `config.ini` updates on the next run.
- Office documents use LibreOffice when Microsoft Office is not installed, under the same Office type.
- LibreOffice conversions run in the background. Hovering no longer freezes the preview, tray, or pointer. A spinner appears while waiting, and the preview appears when ready. Late results are kept for the next hover.
- A stuck LibreOffice conversion is ended instead of blocking everything. One bad file no longer breaks later previews; that document is remembered as unsupported.
- Cleaned up the `[libre]` list: removed 25 file types LibreOffice does not support, including EPUB, older QuarkXPress, early PageMaker, and some Visio stencils and templates. `config.ini` updates on the next run; supported names can be added back by hand.
- `.swf` is now video-only and removed from the `[libre]` list. Previously it could launch LibreOffice first. `config.ini` updates on the next run; custom lists are left alone.
- Previews follow a file's content rather than its extension: a video named as a document is played, a document named as a picture is drawn, and a file whose content no reader here answers for shows nothing instead of starting an engine for it.
- File types whose bytes carry no signature are routed to the engine that reads them by extension, from the formats FFmpeg and LibreOffice are known to read. `.pdb` is one name for two formats: a Palm OS ebook is drawn by LibreOffice, and a compiler's program database is left alone.
- Formats the common signature table does not carry are now recognized by their own bytes as well: the pictures this app decodes itself (`dds`, `exr`, `hdr`, `ff`, `qoi`, the Netpbm family, `pcx`, `ras`, `pcd`, `pct`), the drawings Windows replays (`emf`, `wmf`), the documents LibreOffice imports (`wpd`, `wpg`, `wk1`/`wk3`/`wk4`/`123`, `wb2`, `wq1`/`wq2`, `wks`, `dbf`, `dxf`, `hwp`, `lwp`, `cwk`, `mcw`, `wri`, `slk`, `602`, `pm6`/`pmd`, and the type an OpenDocument, a StarOffice XML document or a Krita project declares inside itself), and FFmpeg's containers and raw streams (`rm`, `wtv`, `nsv`, `smk`, `thp`, `roq`, `dv`, `mxg`'s neighbours such as `c93`, `cdg`, `cdxl`, `moflex`, `rcv`, `viv`, `yop`, `xmv`, `dav`, `m2t`/`tp`/`tr`/`tod`, `vro`, and a raw H.265, H.266, AV1, AVS, VC-1, VC-2, Dirac or EVC stream). A file whose bytes are one of these is previewed as what it is whatever it is called. What is still routed by name is what the bytes cannot settle: a zip, an OLE compound file or a gzip stream that names nothing, a format whose signature is its own text, and a name no demuxer reads.
- `Confirm File Type` is on by default; an installation that already exists keeps the value its `config.ini` holds.
- Building from source now needs Rust 1.98.1 or newer.
- Source files are organized into folders by role. No behavior change.
- A new file under the cursor gets its preview while the mouse is still moving, instead of waiting for the pointer to stop first. Every delay is still a delay from the last movement, so the preview appears as the cursor reaches the file rather than after it settles. `Timing → Settling Delay` asks for the old behavior, at any length.
- `Timing → Delay` and `Timing → Rehover Delay` offer more steps — 0, 25, 50, 100, 150, 200, 250, 300, 400, 500, 600, 700, 800, 900 and 1000 ms — in place of Instant, Fast, Medium, Relaxed and Slow. The 750 ms step is gone: a `config.ini` holding it keeps the value, and the menu shows nothing marked until another step is picked.

### Fixed

- A preview of an Office document no longer shows a smaller version of itself first on a machine that also has LibreOffice installed. The render engine was asked to draw a page for every document, at the size and place of the waiting spinner, and the page the document's own application drew replaced it a moment later. An Office document is asked of the render engine only where its own application is not installed, and a page that engine draws is measured from the page itself and shown at the Office kind's scale, like any other page of the kind.
- Flash animations (`.swf`) no longer hang previews. They were in both the video and LibreOffice lists, so hovering could start LibreOffice, which cannot render Flash and froze the app. Now `.swf` previews as a video.
- File types are checked in the same order everywhere, so videos are no longer mistaken for documents.
- A file renamed to another type's extension is now drawn by the whole of the type its content belongs to, size rules included: a picture named `.mp4` or `.docx` follows the picture scaling rather than the video's or the document's, and one named `.txt` shows at all instead of being read as a page of text that it is not.
- A file whose content is not a document no longer starts Office to look for a page, and a video whose name is not in the video list is played rather than left on its first frame.
- Files preview on one monitor while a maximized or fullscreen window is in front on another. A maximized window anywhere was read as hiding Explorer, so the app slept without asking where the pointer was, and a file on the second monitor previewed only after a click brought Explorer to the front. A window in front now hides Explorer only where it actually covers it.
- A video player left behind by a crash, a forced close, or a Task Manager kill is ended on the next run instead of staying on screen.

## [0.2.14] - 2026-09-23

### Changed

- New defaults, for new installations: previews are placed at their best position, appear instantly, and wait 200 ms before the same file previews again. Transparency is shown over a checkerboard rather than black — for pictures, drawings and design documents (font specimens stay on a white page) — and `DDS Background` offers only Black and White, starting at white. An installation that already exists keeps the values its `config.ini` holds, so if you see black where you expected a checkerboard, that is your file and not a bug.
- `config.ini` is grouped under the same headings as the tray menu — `General`, `Preview Types`, `Text Preview`, `Timing`, `Placement`, `Scaling`, `Background`, `Volume`, `Performance`, `Advanced` — so a setting is easy to find. Every key the app knows is written there, and the tray marks the item each setting starts at with `(Default)`.
- `config.ini` is checked and tidied every time the app reads it, including right after you edit it: a missing setting comes back, a value the app cannot read is replaced with the value it is actually using, and any line that is not one the app writes is removed. A list you edited yourself is left exactly as it is.
- Settings that older versions stored under a different name are no longer carried over — `svg_scale`, `svg_background`, `svg_preview_enabled`, `off_trigger_key`, `avoid_filename`, `transparent_background`. Those lines are removed and the settings go back to their defaults, so if you are updating from 0.2.12 or older it is worth a look at the **Placement**, **Scaling** and **Background** menus afterwards. Updating from 0.2.13, nothing changes.
- The installer no longer writes into `config.ini` and no longer adds the startup entry: the app does both on its first run, at the path it is running from.
- Bump version to 0.2.14 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A CorelDRAW document shows its drawing instead of a blank white page, and the shape that keeps its pictures under `previews/` shows one at all: the picture the program wrote for a file manager is read ahead of a page rendered on its own, which can come out empty.
- Editing `config.ini` by hand now takes effect as soon as you save it, and a value the app cannot read — a misspelled tone map, a delay past its limit — is written back as the value in use instead of staying in the file.

## [0.2.13] - 2026-09-22

### Added

- Preview for design documents: Photoshop `psd` and `psb`, Krita `kra`, OpenRaster `ora`, and the project containers other drawing tools save (`sketch`, `fig`, `xd`). A layered document is shown from the finished picture its format keeps inside it, so it arrives as it was saved.
- Preview for Illustrator `ai` files, drawn as a PDF the same way a PDF is.
- `Design` under the tray's **Preview Types**, with `design_preview_enabled` and a `[design]` extension list in `config.ini`.
- `Design Scaling` under the tray's **Scaling** menu, with `design_scale` in `config.ini`: how much of the screen a design preview covers — Fit to Screen (default), or 75%, 50%, 25%, 10%.
- `Design Background` under the tray's **Background** menu, with `design_background` in `config.ini`: what a design preview is drawn over, black by default.
- Vector previews: Windows metafiles (`wmf`, `emf`) and Illustrator files saved as encapsulated PostScript (`eps`, `epsi`), drawn by Windows at the preview's size so they stay sharp. An `.eps` shows the picture its writer saved inside it; one without such a picture shows nothing.
- `Vector` under **Preview Types**, with `Vector Scaling` and `Vector Background` in the tray, and `vector_preview_enabled`, `vector_scale`, `vector_background` and a `[vector]` extension list in `config.ini`. SVG documents are covered by these — the `svg_background`, `svg_scale` and `svg_preview_enabled` keys are still read.
- `ai` files that are not PDFs — Illustrator saved as PostScript — show the preview saved inside them, under the **Design** kind.

### Changed

- `Vector Scaling` now starts at Fit to Screen rather than half the display.
- `svg` and `svgz` moved from the `[image]` list to the `[vector]` list, where the other drawings are; existing `config.ini` files are updated on the next run.
- Bump version to 0.2.13 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A name in the text list _and_ in a drawing or document list is previewed as that kind again, not as text — the two halves of the app now agree about a file whose name is written down twice.
- SVG previews are back: a document is drawn by the engine again rather than being handed to the readers that replay the other half of the Vector kind, which turned it down.
- A Photoshop document smaller than the screen shows a preview again: it was refused whenever the preview was enlarged past the document's own size.
- EPS files whose preview is a palette picture — the shape Photoshop writes into an EPS — now preview. That shape is read by the app itself, since the picture decoder turns it down.

## [0.2.12] - 2026-09-22

### Changed

- The scaling options now live in a `Scaling` menu of their own in the tray, right below `Placement`.
- `Font Background` now defaults to white; `font_background` in `config.ini` still sets it.
- Tray labels renamed for consistency: `Image Scaling`, `Video Scaling`, `Avoid Nothing`, and `Videos` in `Codecs`.
- Preview hover checks now put far less load on Windows Explorer, so the file list stays responsive while previews are running.

### Fixed

- Keyboard previews in `Details` and `Content` views now follow `Avoid Filename` and `Avoid Filename Column` instead of behaving like `Avoid Details`: a row of a view is recognized by the columns it draws beside the name rather than by its box being at least half the display across, which read the row of a narrower window as a box item and placed its preview past the whole of its columns at every way of avoiding.
- A failed startup of the Explorer lookup could leave hover previews off for the rest of the run; it is now retried until it works.
- The internal timeout that stops a stalled Explorer from blocking the app is now verified instead of assumed, and refreshes itself if a newer one is ever needed.

## [0.2.11] - 2026-09-21

### Added

- Video previews without FFmpeg: where `ffplay` is not installed, videos are decoded by the media engine Windows already has — `mp4`, `mov`, `m4v`, `mkv`, `webm`, `avi`, `wmv`, `asf`, `ts`, `m2ts`, `mts`, `3gp`, and more with the codec extensions below — played in the same preview window as everything else, with sound and looping. FFmpeg is now optional rather than required.
- `Codecs` in the tray menu: what this machine has of every engine and codec a preview can lean on, grouped into `Video`, `Images` and `Engines`, with a check or a cross per row and the missing ones greyed. The rows are for reading only — nothing in the menu does anything, and there is no setting behind it.
- `windows-core` as a declared dependency, for the media engine's event callback. Nothing is added to the build graph by it.
- Sample lines for Arabic, Hebrew, Thai and Devanagari, so a font of one of those scripts is drawn as its own script rather than by the first characters its map happens to hold: the Arabic and Hebrew pangrams, and the openings of the Thai pangram written by the Computer Association of Thailand and of the Devanagari one the script is sampled with. A line is still drawn whole or not at all, so a script a font holds every character of is a line and one it does not is the font's own characters, the same rule as before.
- `Font Face` under the tray's **Placement** menu, with `ttc_face` in `config.ini`: which face of a `.ttc` collection a specimen is drawn from — `First Face` (the default) down to `Tenth Face`. A page has no syntax for naming a face inside a collection, so the face is written out as a font of its own, and this is which one; a collection with fewer faces than the setting names is drawn from the last one it has, and the specimen's heading says which face came out — `(2 of 4)` — so what was asked for and what was drawn cannot be taken for each other.
- DDS texture previews now support almost all DDS formats, including compressed, uncompressed, cubemaps, texture arrays, and mip chains; only the first face and first level are shown.
- Signed DDS textures now display correctly: zero is in the middle, and missing color channels show as neutral gray.
- Signed HDR DDS textures now get a preview and use the same HDR tone mapping as other light-based images; negative light is shown as black.
- More DDS formats now preview: packed HDR, depth buffers, old numbered formats, and signed BC4/BC5.
- New `DDS Background` setting lets you choose what DDS textures are previewed over.
- New `hdr_tone_map` and `hdr_exposure` settings control how HDR images are shown.
- New `spinner_delay_ms` setting controls how long to wait before showing the loading spinner.
- New `Videos Scaling` under the tray's **Placement** menu, with `video_scale` in `config.ini`: how large a video is shown, 100% by default.
- New `Animated Scaling` beside it, with `animated_scale` in `config.ini`: how large an animated GIF, WebP or PNG is shown, 100% by default. A check of the file itself tells an animation apart from a still — a GIF or PNG with a single frame follows `Images Scaling` like any other picture.

### Changed

- `README.md` describes what FFmpeg adds rather than requiring it, and lists the free Microsoft Store codec extensions for HEVC, VP9, AV1, MPEG-2 and Ogg.
- Bump version to 0.2.11 in `Cargo.toml` and `Cargo.lock`.
- A specimen's lines carry their own direction, so the Arabic and Hebrew lines are laid out from the right, their full stop ending them where the script ends it rather than where a left-to-right page would; every other line is drawn as it was.
- A font that covers more of the sample lines than the specimen's box was shaped for — a pan-script one, which covers most of them — is drawn at smaller type rather than past the bottom of the box.
- The page a specimen is drawn in is named for the face as well as the backdrop, so switching faces is a page the browser has not seen rather than the one before it answered out of its cache.
- Video files are now checked in the background with a spinner, and the result is remembered so they are not checked again every time.
- Video previews show the spinner until the video player window appears.
- Video players that fail to open or hang are closed automatically.
- The loading spinner is now the same small pointer spinner for every preview type, not a preview-sized box.
- Spinner delay is now unified at 250 ms by default, including Office documents and browser starts.
- EXR and HDR previews are now tone mapped instead of clipped, so bright areas fade naturally instead of burning out.
- DDS files marked as opaque are now shown opaque instead of using an unused alpha channel.
- DDS previews now use the best detail level for the displayed size, reducing work without losing visible detail.
- Truncated or too-short DDS files are handled safely using the real file length.
- DDS is now in the built-in image list, and older config files are updated automatically.
- The tray's `Scaling` is now `Images Scaling`, and videos have a `Videos Scaling` of their own; both stay hand-editable in `config.ini`.
- Animated GIF, WebP and PNG previews follow their own `Animated Scaling`, separate from still images; a `config.ini` written before this gets the animation scale from `preview_scale`.
- `Font Face` is out of the tray menu; `ttc_face` in `config.ini` still picks the face of a `.ttc`.

### Fixed

- Animations no longer stop at the end of their first play. An animation whose frames were partly given back to keep memory in check was left holding its last frame for good, because the file was only decoded again while a decoder was still running and that decoder had already finished; the two are now settled together, so either the animation is still being decoded or the whole of it is in hand and plays from beginning to end, forever. This affected GIF, animated WebP and animated PNG previews.
- Animations with a long frame in them no longer freeze on that frame. Any frame held for a whole second or more — which a GIF is free to ask for, and one of these asks for 1.2 seconds on its first frame — was unreachable: the playhead was treated as having fallen behind once a second had passed, and its clock was reset every tick, so the wait could never end. The frame's own delay is now part of what "behind" means, and a long hold is simply a long hold.
- An animation that fits the memory it is kept in stays whole instead of being taken apart frame by frame: it is decoded once, plays through, and wraps back into the frame it started on rather than being read from disk again on every pass.

## [0.2.10] - 2026-09-21

### Added

- Font previews for `ttf`, `otf`, `ttc`, `woff` and `woff2`, drawn by the WebView2 engine the SVG previews already use: the name the font calls itself, the pangram _The quick brown fox jumps over the lazy dog._, and a line each for Japanese, Chinese, Korean, Cyrillic and Greek that the font's own character map covers. A line is drawn only where every character of it is in the font — a browser falls back per glyph and says nothing about it — so nothing is ever shown in a system font and passed off as the font; a font of a script there is no line for is shown by the characters its own map holds instead.
- `Font Scaling` under the tray's **Placement** menu, with `font_scale` in `config.ini`: how much of the screen a specimen is drawn over — `Fit to Screen`, or 75%, 50% (the default), 25%, 10%.
- `Fonts` under **Preview Types**, and a `[font]` extension list in `config.ini`; both are written on first run.
- `Font Background` under the tray's **Background** menu, with `font_background` in `config.ini`: a specimen is a page of text, so its ink follows the backdrop — light on black, dark on the light ones, and light with a shadow where there is none to be read against.
- A `.ttc` previews its first face, which is written out as a font of its own beside the browser's profile folder: a page has no syntax for naming a face inside a collection.
- `brotli-decompressor`, for a WOFF2's tables: a specimen is described by two of them — the character map and the name — parsed out of an sfnt, a WOFF, a WOFF2 or a collection's first face on this side, with nothing rasterized.

### Changed

- Build deploy artifacts with a `github` profile (fat LTO, one codegen unit), which with the three changes below takes the released exe from 7,544,320 to 6,676,992 bytes (867 KB, 11.5%); `cargo build --release` keeps thin LTO for the local loop.
- Decode still WebP with the Windows codec at the layout box, removing `image-webp`, and with the bundled libwebp (`webp_image`) where that codec is not installed — which is what a Windows 10 machine usually is; animated WebP is unchanged.
- Remove `rayon` from the build graph, making `.exr` previews decode single-threaded.
- Unify GIF on `gif 0.14` to match `image`.
- Draw SVG previews with WebView2 instead of in-app rasterization, removing `resvg`, `usvg`, `tiny-skia`, and related crates (exe 10.6 MB → 7.2 MB).
- Remove the app’s own animated-document reader, leaving all documents to the engine with no movement check.
- Drop SVG preview on machines without WebView2, since no fallback reader remains.
- Show the spinner immediately on cold-engine document hovers, and not at all when warm.
- Move the engine window with the pointer wait so documents draw where the wait ended.
- Stand down failed engines for one minute instead of five, costing only that hover and clearing its spinner.
- Paint `Checkerboard` for `svg_background` in the engine page so SVG previews keep the standard backdrop.
- Bump version to 0.2.10 in `Cargo.toml` and `Cargo.lock`.

## [0.2.9] - 2026-09-21

### Added

- HEIC (`.heic`, `.heif`), AVIF (`.avif`) and JPEG XL (`.jxl`) previews, decoded by the codec Windows has rather than by one shipped with the app. The codec extensions they need are listed in the README; a machine without one shows no preview for those files, and a multi-image file (a burst, an animated AVIF) shows its first frame.
- A picture of one of those formats is decoded at the size of the preview rather than the size of the file, so a 48-megapixel photograph costs what its preview costs, and it arrives at the codec as the frame is composed rather than through a resample and two conversions.
- `avif`, `heic`, `heif` and `jxl` in the built-in image list; a `config.ini` written by an earlier version is brought up to it rather than left without them.
- `PDF Scaling` and `Office Scaling` under the tray's **Placement** menu, with `pdf_scale` and `office_scale` in `config.ini`: how much of the screen a page is drawn over — `Fit to Screen` (the default), or 75%, 50%, 25%, 10%.
- `110%` in the tray's **Text Preview → Font Size**.

### Changed

- The `image` crate is compiled with its decoders named instead of its defaults, which takes its AV1 encoder out of the build graph: this path decodes pictures and never writes one.
- PDF and Office pages no longer follow `preview_scale`: each page kind has a share of the display of its own, so a picture's scale leaves a page alone.
- Tray menu regrouped: `Text Preview` sits below `Preview Types`, `Trigger Key` moved into `Timing` above `Delay`, `Confirm File Type` moved into `Performance`, and the two engine idle-time submenus are named for what they set — `Office Engine TTL` and `SVG Engine TTL`.
- Bumped version to 0.2.9 in Cargo.toml and Cargo.lock.

## [0.2.8] - 2026-09-20

### Added

- New default “Avoid Filename” keeps previews from covering the file name itself.

### Changed

- Tray menu gets an **Avoid** submenu, and `avoid_filename` is replaced by `avoid_mode`.
- Slide previews now use your screen width (640–1920) instead of fixed 1280, while Word/Excel pages stay the same.

### Fixed

- `Avoid Filename` now measures names using the item’s own display zoom, fixing half-covered names on scaled screens.
- Preview spacing and mouse-move distances now scale correctly with screen zoom at 100%, 150%, and 200%.
- Previews now use the screen your active window is on, fixing multi-monitor fullscreen and missing-preview problems.
- If a preview can’t attach to its screen, it now goes to the main screen instead of spanning two monitors.
- Screen/zoom changes or wake-from-sleep now close and reopen browser document previews at the mouse, not on a missing screen.
- The app now notices monitor rescaling or primary-monitor changes, so preview positions update correctly.
- PDF page size is remembered per file version, so a re-exported PDF is measured fresh.
- Screen zoom is now read from the window under the point, with the monitor as backup.
- PDF and Office-exported pages now fit inside the preview box instead of running off-screen.

## [0.2.7] - 2026-09-19

### Added

- SVG previews (`svg`, `svgz`), drawn at the size the preview is shown at and sized like a picture: `100%` is the size the document asks for, `50%` that at half.
- Animated SVGs are played by the WebView2 runtime Windows 11 ships with, and by this app's own reader where it is not installed. The engine is pointed at a page of this app's own that draws the document as an image, so it fills the preview box at every scale.
- `webview_idle` and a `Performance → Keep Animated SVG Engine` entry, ten minutes by default: the browser is kept warm between hovers rather than started for each one.
- The engine keeps its state in a folder per run, cleared at startup; one that fails is retried, then stood down for five minutes so this app's reader plays the document, and the hover that was up is laid out again. `RHP_WEBVIEW_TRACE` and `RHP_WEBVIEW_PROFILE` are the diagnostics for it.
- `trigger_key_enabled` in `config.ini` and a check in the tray `Trigger Key` submenu, to switch the trigger key off without changing its mode.
- `svg_background`, the backdrop an SVG document is drawn over, in `config.ini` and as `Background → SVG Background` in the tray: Transparent, Black, White or Checkerboard, the same four a picture is offered.
- `svg_scale`, how much of the screen an SVG document is drawn over, in `config.ini` and as `Placement → SVG Scaling` in the tray: `Fit to Screen`, or `75`, `50` (default), `25` and `10` percent of the display. A vector is drawn at whatever size it is asked for, so a document's percentage is of the room the display has rather than of the size the file asks for — half the screen at `50%`, where a picture at `50%` is half of its own size. Both readers follow it: the still frame this app draws, and the engine that plays an animated document.
- `svg_preview_enabled` and an `SVG` entry under `Preview Types`, below `Office`: a document is its own kind even though `svg` and `svgz` are entries of the image list, so the two are switched independently — `Images` off leaves a document previewing, and `SVG` off leaves pictures alone.
- `decode_budget_gb`, what one hover may decode or read for in gigabytes (default `1`), in `config.ini` and as `Performance → Decode Budget` in the tray: a file past it gets no preview instead of memory the app may not get.

### Changed

- Installing over an older version no longer asks: the previous version is uninstalled first and a running copy of the app is closed instead of prompted for.
- The NSIS setup is built from `packaging/nsis/installer.nsi` on a pinned `cargo-packager` version, and the dead `wix` format is gone.
- `cargo build --release` closes a running copy of the app before linking; debug builds are untouched.
- Video previews no longer write `%TEMP%\rust-hover-preview-video.log`, and one left by an earlier version is deleted at startup.
- `cargo fmt --check` and `cargo clippy --all-targets -- -D warnings` pass again.
- The tray's `Background` submenu is above `Volume` and holds an `Image Background` and an `SVG Background` half, each listing Transparent, Black, White and Checkerboard; `transparent_background` is read as both `image_background` and `svg_background`, and written out under the two names.
- An SVG is no longer sized by `preview_scale`: it is drawn at `svg_scale`, half the screen by default, and a picture's scale leaves it alone.
- Disabling Office/SVG now kills its engine immediately (Office at once, browser within a tick), whether toggled in UI or config.ini. Renders refused because the tier was off no longer count against the document, so re-enabling Office previews it without waiting out backoff.

### Fixed

- An animated SVG stopped moving once the engine had been let go for idle — ten minutes after the last one, on the default `webview_idle`: the thread that played it ended with its browser while the app still held its handle, so every document after that was left on its still frame for the rest of the run. The thread now outlives the engine and begins another one for the next document.
- The WebView2 browser no longer outlives the app when the app is killed, and one left behind by an earlier run is ended by the next launch instead of being left holding its profile folder.
- Every process the app starts — an Office engine, the WebView2 browser, an `ffplay` — is put in a Windows job object, so a crash, a kill from Task Manager or a logoff ends them with the app, and is recorded under `%LOCALAPPDATA%\rust-hover-preview\engines` so that the next launch ends whatever the job could not take. Nothing is ever acted on by id alone: a record is used only when the process still carries the image _and_ the start time it was recorded with.
- One engine per Office family is enforced rather than assumed: a family's engine is started only when nothing this app began for it is still running, and a replacement waits for the process it replaces to be gone. Exiting while a document is mid-render no longer leaves the engine behind either.
- Pictures over 40 megapixels preview again: the pixel cap is gone, and every reader — pictures, animated GIF/APNG/WebP, SVG documents, Office exports, themes — now asks the decode budget above before it allocates.

## [0.2.6] - 2026-09-17

### Added

- Office previews for Word, Excel and PowerPoint, rendered by the installed Office.
- `office_preview_enabled`, `office_cache_mb`, an editable `[office] extensions` list, and an `Office` entry under `Preview Types`.
- Editable `[image]` and `[video]` extension lists in `config.ini`, so which pictures and videos preview is a file edit like the text, archive and office lists.
- A deleted `extensions=` key comes back with its built-in list, in the file as well as in memory.
- `pdf_cache_mb` and `text_cache_mb`, with `PDF` and `Text` entries in the new tray `Cache` submenu; memory-only, capped at `2048`, default `32` for PDF and `0` for text.
- Tray `Cache` submenu for image, text, PDF and Office caches (`0 MB`–`2 GB`), plus `90%`, `80%` and `70%` font sizes.
- A `Performance` submenu holding the caches and a new `Keep Office Engine` setting: `0 seconds`, `1 minute`, `5 minutes`, `10 minutes` (default), `30 minutes`, `1 hour` or `Indefinitely`, with `office_engine_idle` in `config.ini` behind it.

### Changed

- The built-in extension and file-name lists are written in alphabetical order, and a list still holding them is rewritten in that order on upgrade.
- Office previews render pages instead of using saved thumbnails; rendered pages stay in memory (`office_cache_mb`, default `64 MB`).
- A document's page is asked for as soon as it is hovered rather than after a two-second rest, so a preview waits for the render instead of for a timer in front of it.
- Each family keeps its own Office engine, so a folder holding a document, a workbook and a deck starts each application once rather than quitting one and starting another every time the pointer crosses between them; an engine is let go ten minutes after its family was last asked for a page.
- The settings a render needs of an Office instance — dialogs suppressed, macros switched off — are taken for that render and put back when it is over, instead of being held for as long as the engine lives. Nothing of the instance's own is therefore held reconfigured while an engine sits warm between documents, which is what makes `Indefinitely` safe for an instance that is the user's own Word or Excel.
- PDF and Office pages use fit-to-screen sizing, but preview scales below `100%` now reduce it; `100%`+ and image previews are unchanged.
- The waiting spinner is now a transparent haloed arc, appears immediately, follows the pointer, and is placed flush at the pointer's corner.
- A page arriving over an existing preview replaces it without taking the preview down first.
- Worksheet images are copied only as far as needed and fast-shrunk; refused renders are retried, and unreadable pages are re-rendered.
- A document that refuses a page is ignored for two minutes instead of ten.
- `image_cache_mb` defaults to `32` rather than `64`, and `pdf_cache_mb` to `32`, so the decodes and rasters a folder is swept back over are already done; the text cache still starts at `0`.
- Removed the `Office Preview` submenu; `office_render_enabled` is no longer read.
- A workbook's picture — the no-printer answer — is asked for on a machine that has a printer only when an attempt is already a retry: an export that came to nothing there is the instance declining rather than the document failing, and a copy would be refused by the same instance, after the retries the clipboard needs.
- An Office engine let go because it refused a page is ended where it stands rather than asked to quit and pumped for the seconds an unanswered quit takes, so the retry reaches the fresh instance, which is what the retry is for, without the wait in front of it.

### Fixed

- PowerPoint decks and Excel workbooks on printerless machines now preview.
- Office engines no longer survive quit, and failed or unresponsive instances are replaced.
- Workbook clipboard images are released.
- A preview that crossed a display boundary was discarded and stayed gone until the pointer moved; the hover it came from is now replayed, so it is laid out again at the new display's scale and put back.
- An Office page that landed while another hover message was in hand left the preview waiting with nothing to end the wait; the cap now still applies, so the spinner comes down.
- The Office render worker makes itself known before initializing its apartment, so the first request after startup can no longer start a second worker beside it.
- Non-document files no longer start an Office engine.
- The document spinner is correctly sized.
- Replaced videos are re-measured and re-cropped using file version as well as path.
- `config.ini` is written in a fixed order — `[settings]` first, then the sections and their keys alphabetically — instead of being reshuffled on every save.
- A page landing over a waiting spinner no longer flashes the spinner stretched across the page's box: the window is given the page's size and place by the paint itself, in one call, rather than being resized ahead of it.
- A pointer that drifts onto a waiting spinner no longer dismisses the hover it belongs to — the spinner publishes the box it occupies as one that holds the pointer — so the page being waited for is not thrown away with the hover, and the first hover of a document after a folder change is no longer answered with nothing and then with its page on the second.
- Explorer crashing or being ended no longer leaves previews dead until the app is restarted.
- Office's own "Publishing…" progress window no longer flashes over a hover: a render watches for that window in the process it started and hides it, which no setting the engine holds — alerts, screen updating, the automation security mode — was able to do.

## [0.2.5] - 2026-09-17

### Added

- Archive previews: hovering an archive shows a summary line and a capped tree of what it holds, read from its own table of contents without unpacking anything.
- `archive_preview_enabled` and an `[archive] extensions` list in `config.ini`, plus an `Archives` entry in the tray's `Toggle Preview Types` submenu.
- Archive listings are painted with the text preview's themes, fonts and margins, and follow the `Text Preview Font Size` setting.
- `Avoid Filename` in the tray's `Preview Position` submenu, or `avoid_filename` in `config.ini`: a preview is moved clear of the name of the file it is about, and resized where the display leaves no room for it beside the name. On by default, and applies to both positions.

### Changed

- The tray menu is grouped into categories — `Text Preview`, `Timing`, `Placement` and `Volume` among them — with shorter labels, and `Config.ini` carries the running version.
- The minimum supported Rust version is now 1.88, the `zip` crate's floor.
- The GDI painting a text preview and an archive listing share moved into `text_paint.rs`.
- Archive listings keep room between an entry's icon and its name.
- The tray's `Preview Position` entries carry radio marks, and `Avoid Filename` a checkmark of its own.
- RAR archives are read through RARLAB's UnRAR sources compiled in by the `unrar` crate. See the licence note in README.

## [0.2.4] - 2026-09-17

### Added

- `image_cache_mb` in `config.ini` (default `64`) sets how much memory decoded image previews may be kept in, so hovering back over a folder no longer decodes the same pictures again. `0` turns the cache off, and the value is capped at `2048`.

### Changed

- The preview thread no longer wakes on a timer while nothing is on screen: it waits on the hover channel instead, so an idle app stops waking sixty times a second, and a hover is answered as it arrives rather than on the next tick.
- Animated GIF, APNG and WebP previews open on their first frames instead of waiting for a startup buffer.
- Video previews read a file's dimensions and detect its letterboxing at the same time instead of one after the other.
- A PDF's page size is remembered from the render that read it, so a file is not opened twice for one preview.
- Bumped version to 0.2.4 in Cargo.toml and Cargo.lock

### Fixed

- A PDF is no longer read into memory in full before its first page is rendered, so a large one costs memory in proportion to the page rather than to the file.
- A OneDrive or SharePoint file that has not been downloaded no longer starts downloading when the pointer rests on it.
- A very large image can no longer take the app down with it.
- Hovering no longer stalls indefinitely when Explorer stops answering.

## [0.2.3] - 2026-09-17

### Changed

- Hover previews in a search results view are resolved from the item under the pointer instead of its file name, so a search of any size previews as fast as a folder does.
- Keyboard previews resolve the same way, from the item the focus is on.
- Removed the name-based fallbacks and the search-result indexing behind them: no directory walks, no background scans, no index caches, no MSAA.
- A file nothing can vouch for is left without a preview rather than guessed at by name.
- Documentation updated: ARCHITECTURE.md describes the resolution, and the search-results known issue is gone from TODO.md.
- Text previews wrap long lines instead of cutting them off at the right edge.

### Fixed

- A text preview of a file with long lines is now tall enough to show the rows the line wraps into, instead of cutting them off below the first one.
- Search results that share a file name each preview their own file.
- A search result that lives in another folder previews from the keyboard too, not only under the pointer.
- A Details or Content row previews from anywhere on it, not only from its text.
- A keyboard preview of a Content view row is placed in the empty tail past the row's last column — the edge the row's own text stops at, read from the view — instead of in the middle of the display, and the space to the right of the columns is what sizes it.
- A search result previews in any tab of an Explorer window, and follows the tab when it changes.
- Hovering a folder, an executable or any other item this app does not preview no longer shows another tab's file.
- Bumped version to 0.2.3 in Cargo.toml and Cargo.lock

## [0.2.2] - 2026-09-16

### Added

- New tray menu **Toggle Preview Types** to turn image, video, text, and PDF previews on or off separately.
- **Select All** in the text preview right-click menu.
- Better detection for `.ts` and `.mts` files (video vs. text).

### Changed

- Removed the old **Enable Text Preview** tray item; it's now under **Toggle Preview Types**.
- Documentation updated.
- Version bumped to 0.2.2.

### Fixed

- Search results with the same file name are now told apart correctly.
- Keyboard previews respond faster, especially in search results.
- Text previews no longer accidentally respond to Ctrl+A.
- Previews no longer get stuck or fight between mouse and keyboard.
- Many other small fixes for folders, search, and keyboard navigation.

## [0.2.1] - 2026-09-16

### Added

- Text previews: see the first screen of code, Markdown, logs, `.nfo`, `.rtf`, and more.
- Syntax highlighting and rendered Markdown.
- Light and dark themes, plus custom themes from a `theme` folder.
- Font size options (100%–400%).
- Full mode: scroll, select, and copy text.
- Toggle text previews on/off.
- Better encoding detection.

### Changed

- Text previews size themselves to fit content.
- Previews are faster after first load.
- Documentation updated.
- Version bumped to 0.2.1.

### Fixed

- Scrolling and Markdown preview fixes.

## [0.2.0] - 2026-09-16

### Added

- PDF previews: see the first page of a PDF using Windows’ built-in PDF reader.

### Changed

- PDF previews always use the best available size for sharp text.
- Documentation updated.
- Version bumped to 0.2.0.

### Fixed

- PDF previews now work in normal folder views.

## [0.1.14] - 2026-09-16

### Added

- More image formats: APNG, TGA, HDR, EXR, QOI, and more.
- Animated PNG support.
- Many more video formats.
- **Preview Scaling** menu: fit to screen or 25%–400%.
- Only one copy of the app can run at a time.
- Better sleep/resume behavior.

### Changed

- Previews stay on the correct monitor.
- Smoother, flicker-free previews.
- Keyboard previews take priority over the mouse.
- Wheel scrolling refreshes the preview.
- Documentation updated.
- Version bumped to 0.1.14.

### Fixed

- Many fixes for animated images, video cleanup, keyboard/mouse switching, and folder changes.

## [0.1.14-rc.10] - 2026-09-16

### Added

- Animated PNGs recognized inside `.png` files.

### Fixed

- Animated WebP, APNG, and GIF no longer stop part-way.

### Changed

- Smoother preview window and video playback.
- Explorer hook uses less memory.

## [0.1.14-rc.9] - 2026-09-15

### Added

- More image formats (APNG, TGA, HDR, EXR, QOI, farbfeld).
- Many more video formats.
- Better `.ts`/`.mts` detection.

### Changed

- Video format list moved to one place.

### Fixed

- No stray video processes left behind.
- Folder previews work without moving the mouse.
- First arrow key press shows keyboard preview after folder change.

## [0.1.14-rc.8] - 2026-09-15

### Added

- System-wide mouse wheel detection.

### Fixed

- Wheel scrolling no longer leaves the preview stuck.
- Wheel scroll takes over from keyboard preview.
- Scroll probes only act on stable items.

## [0.1.14-rc.7] - 2026-09-14

### Added

- Single-instance enforcement (only one copy runs).

### Fixed

- Keyboard previews take priority over the mouse.
- Cursor jitter no longer cancels keyboard previews.
- Mouse takes over cleanly from keyboard.

## [0.1.14-rc.6] - 2026-09-14

### Added

- **Preview Scaling** tray menu (Fit to Screen, 25%–400%).
- `preview_scale` config key with hand-editable values.

### Changed

- Previews scale to fit screen edges.
- Smoother upscaling for GIF/WebP.
- Version bumped to 0.1.14-rc.6.

## [0.1.14-rc.5] - 2026-09-14

### Changed

- Documentation for display-aware placement.
- Installer asset name corrected.

### Fixed

- Previews bounded to nearest display.
- No flicker when changing previews.
- Reposition before installing new frame.

## [0.1.14-rc.3] - 2026-07-03

### Added

- Sleep/resume resilience for preview and tray windows.
- Cleanup on suspend/standby.

### Changed

- Installation instructions clarified.
- Config example updated.
- Version bumped to 0.1.14-rc.3.

## [0.1.14-rc.2] - 2026-07-01

### Changed

- Faster Explorer detection (less COM usage).
- Input-aware probing.
- Stationary hover probe runs only once per hover.
- Version bumped to 0.1.14-rc.2.

## [0.1.14-rc.1] - 2026-06-24

### Changed

- Better multi-monitor and display change handling.

## [0.1.13] - 2026-06-03

### Changed

- Release 0.1.13: version bump and changelog update.

## [0.1.13-rc.2] - 2026-05-13

### Added

- Features section in TODO.md for document support.

### Changed

- Increased cache limits for performance.
- Simplified deploy workflow.
- Version bumped to 0.1.13-rc.2.

## [0.1.13-rc.1] - 2026-05-05

### Added

- Per-monitor DPI awareness.

### Changed

- Removed MSI installer references.
- Switched to rust-lld.
- Documentation updates.
- Added ARCHITECTURE.md.

### Fixed

- Fixed high-DPI scaling artifacts.

## [0.1.12] - 2026-04-30

### Added

- Shell view index sync limits and cache clearing.

### Changed

- Refactored media path resolution.
- Enhanced hover preview logic for large result sets.

### Fixed

- Fixed hover preview state management.

## [0.1.11] - 2026-04-30

### Added

- WebP playback FPS configuration.

### Changed

- Refactored file handling in Explorer hook.
- Improved Windows resource description.
- Better search view metadata handling.

## [0.1.10] - 2026-04-29

### Changed

- Restored default NSIS older-version prompt.
- Limited automatic startup to first install/run.
- Removed custom installer taskkill actions.
- Updated release metadata.

## [0.1.9] - 2026-04-29

### Added

- Transparent background modes for PNG/WebP.
- Configurable same-file rehover delay.
- Installer migration cleanup.

### Changed

- Switched to Google's libwebp for animated WebP.
- Moved installer output to `%LOCALAPPDATA%`.
- Config path uses `directories` crate.
- Updated tray labels and tooltip.

### Fixed

- Installers stop running app before upgrade.
- Config default completion.
- Stuck video previews hidden more aggressively.
- Explorer slow-probe backoff.
- Repeated same-file hover timing.
- Transparent PNG/WebP rendering.
- Removed dead code.

## [0.1.8] - 2026-04-29

### Added

- Optional off-trigger key.
- Active Explorer Shell view detection.
- Media indexing from search roots.

### Changed

- Improved path normalization and URL handling.
- Cached Explorer folder lookups.
- Updated installer build orchestration.

### Fixed

- Removed branch-triggered release deployment.

## [0.1.7] - 2026-04-28

### Added

- Windows installer packaging via `cargo packager`.
- `build-installers.ps1` helper.

### Changed

- Updated deploy workflow to publish installers.
- Refactored preview window and media loading.
- Improved installer build orchestration.

### Fixed

- Added Windows Installer availability checks.
- Improved installer error handling.

## [0.1.6] - 2026-04-06

### Changed

- Optimized animated GIF/WebP streaming.
- Limited animation to 300 frames and 30 FPS.
- Added **Confirm File Type** tray option (off by default).

### Fixed

- Reduced CPU spikes for animated WebP.
- Improved rapid-hover behavior.
- Fixed `.jpeg` hover preview.
- Mislabeled images load with correct decoder.
- Prevented previous image flash.

## [0.1.5] - 2026-04-01

### Changed

- Improved Explorer folder caching.
- Optimized media file resolution.

## [0.1.4] - 2026-03-03

### Changed

- Refined folder-navigation input gate.

### Fixed

- Eliminated residual previews after folder open.

## [0.1.3] - 2026-03-03

### Added

- Post-folder-navigation input gate.

### Changed

- Mouse hover target validates against cursor.
- Prioritized media resolution in folder under cursor.

### Fixed

- Prevented auto-preview of first item.
- Keyboard preview starts only after real input.
- Removed dead-code warnings.

## [0.1.2] - 2026-03-03

### Added

- Separate cursor-over detection for image and video previews.
- Startup grace period for video previews.

### Changed

- Refined dismissal logic.
- Improved ffplay window discovery.
- Reasserted topmost styles during video playback.
- Updated ffprobe/ffplay flags.

### Fixed

- Prevented premature video dismissal.
- Reduced false preview closures.

## [0.1.1] - 2026-03-03

### Added

- Keyboard preview support.
- Animated WebP and improved GIF/WebP streaming.
- Loading spinner overlay.
- Tray options: follow cursor and video volume.
- GitHub Actions workflows.

### Changed

- Refactored preview window and media streaming.
- Enhanced animated media playback.
- Optimized CPU polling rates.
- Config uses INI format with auto-save.

### Fixed

- Preview window hides correctly during keyboard hover.
- Corrected preview delay menu wording.
- Improved frame decoding reliability.

### Miscellaneous

- Updated README.
- Initial config generation.

## [0.1.0] - 2026-02-03

### Added

- Image preview on hover: JPG, JPEG, PNG, GIF, BMP, ICO, TIFF, WebP.
- Animated GIF and WebP playback.
- Video preview: MP4, WebM, MKV, AVI, MOV, WMV, FLV, M4V.
- System tray controls.
- INI config file in `%APPDATA%\rust-hover-preview\config.ini`.
