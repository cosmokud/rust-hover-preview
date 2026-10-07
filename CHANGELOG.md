# Changelog

## [Unreleased]

### Added

- **`Pin Mode → Enable (Space)`** — Press `Space` while a preview is up to turn it into a captioned, movable, always-on-top window that stays when the pointer leaves; key set by `pin_key` (`space` default), feature toggled by `pin_enabled`.
- **Pinned caption controls** — Minimize collapses to a draggable round bubble (click to restore, right-click to close), Maximize fits the media to the screen centered while keeping its shape, and Close ends the pin.
- **Move by picture** — A pinned window can be dragged by the picture as well as the caption, except over a text preview's own text/scrollbar and a video's FFmpeg player.
- **Pinned video transport bar** — Play/pause, draggable seek bar and two clocks for media-engine video; FFmpeg video shows a read-out bar instead because that player can be told nothing.
- **Previews held back while a pin is up** — Nothing is raised or dismissed until the pin closes.
- **Caption Previous / Next / Open With** — Step through sibling files in the pin's own folder, in Explorer listing order, only onto previewable files, wrapping at both ends and stepping over broken/unopenable files; Open With uses the Windows-registered app.
- **`Open With...` button** — Opens the Windows "How do you want to open this?" dialog, steps the pin out from on top while the list stands, and restores it after.
- **Hand-off buttons name themselves** — Hovering Open With shows the program it would use (read once when pinned), Open With... says its name, and the name hangs below the caption over the picture.
- **Left/Right and Up/Down arrows walk pinned files** — Same as caption Previous/Next, works whether Update Preview is on or off, but only once the pinned window has been clicked for keyboard focus.
- **Pinned loading spinner** — A pinned swap reads and decodes on a separate thread, keeps showing the old file, and draws the spinner over it after `spinner_delay_ms`, so the window never freezes and is never hidden while waiting.
- **`Pin Mode → Update Preview`** — A pinned window can be shown the file you click or keyboard-select while it is up, in place; Enabled is default, On Hover adds pointer hover.
- **`Pin Mode → Pause Preview`** — A collapsed bubble pauses audio/video and resumes at the stopped second when restored; Audio and Video are separate switches, both on by default.
- **`Pin Mode → Nav File Types`** — All steps through every previewable file, Category narrows to the pinned file's own kind; written as `pin_nav_file_types`, old files read as All.
- **Pin key restores minimized pin** — Pressing the pin key while a bubble is up and Explorer is in front restores the window on the file picked since, in the same place and size; with nothing picked it does nothing, and right-click closes the bubble.
- **`Timing → Trigger Key → Affect Pin Mode`** — Off by default, the trigger key is not read while a preview is pinned; on, holding it brings the pin down with the previews it stops.
- **Pinned sound card controls** — Previous, Play/Pause and Next on the left, Volume on the right opening the same compact vertical level panel as pinned video; drawn on the card, with the bar row as tall as a button and the bar centered.
- **Pinned sound volume belongs to the window** — A pinned sound uses `Volume → Audio` rather than Video, the knob is heard as it moves, and FFmpeg takes the settled release level; hover cards remain buttonless and full-width.
- **Pinned sound bar easier to click** — The clickable band now reaches eight scaled pixels below the bar as well as above, clamped at the card bottom; a lengthless file only refuses seek, not its buttons.
- **Pinned sound Space pause/resume** — Click the window, press `Space` to hold and resume from the stopped second; only a keyboard-focused pin answers, pictures/pages swallow it, 0% audio still plays silently and pauses, and a video with failed geometry probe may also answer.
- **Pinned sound Normalize seek fixed** — A sound sent to FFmpeg by Normalize gain can now be seeked/paused because the actual player, not the probe, decides which engine answers.
- **Pinned sound bar click-to-seek** — Clicking the bar seeks to that second whether playing or paused; the clickable row includes the gap above, the margin below still drags the window, and lengthless streams only drag.
- **`Engine → Select Engine → Video`** — Best (default) uses media engine at/under 3.2MP and FFmpeg above or unmeasured; Native, FFmpeg, and Native (FFmpeg above 3.2MP) are explicit; Fallback on by default; written as `video_engine` and `video_engine_fallback`.
- **Video format lists split** — `[video]` is what the media engine is asked to play, `[ffmpeg]` is FFmpeg-only; existing config is divided on next run, edited lists left alone.
- **Letterbox cropped** — A video's letterboxed picture is cropped to its content whichever engine plays it, so black-barred files fill their box.
- **Video scaled up to fill** — A video shown larger than its picture is scaled up to fill the box instead of drawn native-size in the middle.
- **Animated JPEG XL plays** — Decoded by a built-in JPEG XL decoder, using the same machinery as GIF/APNG/animated WebP, so Animated Scaling and Images gate apply; animated AVIF and HEIF sequences still show first frame.
- **`Scaling → Text Scaling`** — Default Fit to Screen caps the text preview box size for plain text, code and Markdown, while short files still get a small box; Text Size still scales inside it.
- **`Text Preview → Render HTML`** — Default Off; on, `.htm`/`.html` is run by the browser engine at Document Scaling, sandboxed with no forms/popups/top-frame navigation and no fetched links; SVG and font specimens stay drawn, not run.
- **`Background → HTML Background`** — Default White; offers White, Black, Checkerboard only, because a page is not drawn over transparency; old `transparent` reads as White, while Vector and Font backgrounds still decide for their kinds.
- **Pinned picture caption/bar on top** — The window is exactly picture-sized, caption and bar fade in/out with the pointer, and only the strip under the pointer shows.
- **Pinned video volume own** — Speaker button opens a knob over the video, level belongs to that window alone and is independent of hover `Volume → Video`.
- **`Volume → Video → Remember` and `Volume → Audio → Remember`** — Off by default; on, the pin knob level is written to `config.ini` (`remember_audio_volume`, `remember_video_volume`) and the next hover uses it, including across a sound-to-film-and-back walk.
- **`Scaling → Audio Scaling`** — A sound's card is sized as a share of the display, like every other preview kind: `25%`, `20%`, `15%`, `10%` (default) or `5%`, in a new submenu between **Video Scaling** and **Animated Scaling**. The whole card scales with the share — font, height and width — so the card at any share is the `10%` card uniformly smaller or larger (~188×50 at `5%` to ~936×250 at `25%` on a 3440×1440 display). The card is built at the share's own font (the default text size, `125%`, at the `10%` anchor), so **Font Size** no longer resizes it; the Theme setting still applies. Saved to `config.ini` as `audio_scale`, read back on startup; a hand-edited value is sanitized rather than fatal.

### Changed

- **Pinned sound has no title bar** — The window is only the card: name at top, controls on its own row, no caption band or gap; Space, arrows and Escape still work from the pin's own window.
- **Maximized pin no longer maximizes text, archive listing or page** — These are already measured against display room, so maximize leaves them laid out as they are, deliberately unlike shaped files.
- **HTML page can be pointed at and typed into** — The page rectangle holds the hand, drags starting there survive leaving it, clicking gives the page keyboard, and pictures/videos still dismiss on hover; page still cannot have sound, side files, fullscreen, right-click menu, tools or find bar.
- **Dragging a maximized pin gives up maximize** — Any drag that actually moves the box ends maximize, the caption reverts to a maximize glyph, and a zero-move press leaves it standing.
- **Pin key uses low-level keyboard hook** — Pinning now uses `WH_KEYBOARD_LL` to catch press/release rather than sample state; disabling Enable Pin unbinds it; this is the second low-level hook, and `PRIVACY.md` says keystrokes are never recorded, stored or sent.
- **Version bumped** — `0.4.0` in `Cargo.toml` and `Cargo.lock`.
- **`Text Preview → Full Mode` removed** — `text_preview_full_mode` is gone because pinned text now always comes up in full mode for scroll/select/copy.
- **Paused pinned video play triangle removed** — A paused pin now shows the last frame with nothing over it.
- **New config keys** — `pin_enabled`, `pin_key`, `pin_update_enabled`, `pin_update_on_hover`, `pin_pause_audio`, `pin_pause_video` and `trigger_key_affect_pin_mode` are new; old files keep old behavior.
- **`video_engine_probe` diagnostic** — `RHP_VIDEO_PROBE` now plays the file as well as probing it and can report a seek.
- **Docs updated** — `README.md`, `ARCHITECTURE.md` and `PRIVACY.md` describe pinning, Pin Mode submenus, the two video lists and the media engine's role.

### Fixed

- **Pinned window no longer freezes or leaves pointer stuck** — Slow reads, re-decodes, folder walks and Shell queries run on their own threads; stuck presses end quickly, stale walks are dropped, lost pointer captures are released, and a pin whose thread stops is taken down as a last resort.
- **Pinned window no longer closes on unplayable video/sound** — A file nothing could draw is stepped over; if nothing is left, the window stays on the failed file and shows a large cross on a plain panel while the caption still names it, unlike a no-preview file.
- **Step over unshowable files** — Previous/Next and arrows keep walking until they find a showable file, bounded by the folder, instead of stopping dead; a file you picked in Explorer is still not stepped over.
- **Pinned window can walk onto another video** — A second media-engine video is held for the first frame, so the old picture stays until the engine draws.
- **First click after double-click folder open works** — A gesture is now two presses inside Explorer's double-click time, so the click on a file in a newly opened folder updates the pin.
- **First keyboard selection after folder change works** — The baseline is re-established when the shell cannot answer for the place, so the first arrow-key pick updates the pin instead of being lost.
- **First click in Explorer listing updates pin** — Clicks are read through this app's own windows and Explorer's, falling back to the focused item in the click's own place, so a pin standing over the listing no longer eats the first click.
- **Explorer View/Sort popups no longer change pin** — A press is a pick only if it landed on Explorer's or this app's own window and that window is still there; presses taken by another window are dropped, not retried.
- **Arrow keys only walk when keyboard is in pin** — A pin must be clicked before its keys answer, and Explorer no longer moves its own selection while the pin is on screen; Update Preview no longer follows keyboard while the pin holds it.
- **Clicking pin takes keyboard; left-click bubble restores** — Taking keyboard commits any inline rename in Explorer behind it; the pin stays out of Alt+Tab; only right-click closes the bubble.
- **Pin key never hides a pinned preview** — It does nothing to a pin wherever the keyboard is, still restores a minimized pin with Explorer in front, and Alt+F4 cannot destroy a pin.
- **Pin key works over browser-engine files; HTML stops when closed** — The engine is asked as well as app media, so pinning works over HTML/SVG/font specimens, and the engine is told to stop when its window is off screen to stop animated pages burning GPU.
- **Pinned window keeps Explorer answered as in use** — A reachable Explorer with an open pin is treated as active whichever has keyboard, so tick cadence and idle timer stay correct.
- **Pinned sound no longer closes after every pass** — A sound player not running is treated as the moment between passes, not a dead window; video/document still come down if their player/browser is gone.
- **Pinning video or sound no longer freezes app** — Fixed the freeze when pinning media.
- **Half-written subtitle copies are repaired** — Cache reads only byte-holding `sub<i>.<ext>` files from a folder with the pass mark; failed zero-byte folders are ignored, deleted and recopied next time.
- **Film with uncopyable subtitles still previews** — Failed subtitle copies no longer name the film's own embedded track or use PGS `.sup`; only small filter-drawable files are named, failed passes are removed, and the film plays without subtitles.
- **Only one film's subtitles read at a time** — One extraction slot, requests drop the previous, copies stop when hover ends/pin closes/pin swaps, request moved from probe to launch, and a pinned film restarts its player with subtitles when ready.
- **Pinned video seek bar works again** — Seek works at any playback point, far-left restarts, paused seeks follow instead of springing back, and the thumb lands under the press wherever on the bar.
- **Pinned video no longer flashes sheared picture or blank backdrop** — Resize release no longer shears, and media-engine hover no longer flashes the previous preview or blank backdrop at start.
- **Dragging pinned video faster and FFmpeg holds last frame** — Moving no longer repaints everything, resizing no longer resamples per pixel, FFmpeg pins hold the last played frame while dragged, and playing film continues across FFmpeg reload.
- **Pinned FFmpeg video no longer shows player chrome** — The app's caption/bar are drawn over the player band, the player window is held below and never takes keyboard, so all pinned videos look and behave the same.
- **Preview loop routes video like loader** — Every video no longer uses FFmpeg's own window on FFmpeg machines, so pinned videos can resize, maximize and seek.
- **Media-engine play check fixed** — Ordinary `.mp4`, `.mkv`, `.webm`, `.avi` now get the right answer, and a file accepted but not drawable falls back to FFmpeg instead of previewing as an empty box.
- **Pinned sound card cut off fixed** — The pin box is measured with controls on, so the button row and bar are on the card and the bar click works.
- **Previous button drawn correctly** — It is now the exact mirror of Next and the same size, instead of a wrong-facing triangle.
- **0% audio plays silently and stays seekable** — A zero level still uses a running player as silence instead of no player, so the card keeps its clock and seek; truly unplayable sound still has no clock.
- **Pinned sound stops when pin shows a film** — A pin swap now ends the standing sound player while still holding the media-engine video frame.
- **Pinned sound no longer killed by leftover-player sweep** — The once-a-second sweep leaves a pin's own player alone, so pinned sound can be sought and walked.
- **Pinned media scales with resize** — Media now scales as edges are dragged; all edges/corners resize, smaller as well as larger; pointer shape matches the edge.
- **Top-pinned caption no longer drawn off screen** — Caption and buttons stay visible when pinned at screen top.
- **Pinned window restores centered** — Coming out of the bubble centers where it is; maximized restore centers on the screen it is on.
- **Bubble placement fixed** — Bubble lands on the minimize button, repeated collapses do not wander, old position is forgotten, dragging is smooth and stays under the grab point.
- **Pinned swap no longer shrinks** — New file fits the window's longest side, so portrait-to-widescreen does not step down sizes, and the window stays where the hand put it.
- **Text/archive/HTML laid out for themselves** — These now use their kind's Scaling and are centered in the existing window box instead of inheriting the previous file's shape; sound cards already worked this way.
- **Maximized pin honours Scaling on walk** — A maximized pin now lays out the next file at its kind's scale rather than always fit-to-screen; text/listing/sound cards are still measured against display room.
- **HTML page uses Document Scaling in pinned window** — It is no longer fitted into whatever box the pin was standing in.
- **Text preview measured before pinned load** — A waiting text measure shows the pin spinner instead of a placeholder square.
- **Maximized window no longer shrinks or stretches on walk/restore** — Each file is fitted to the display afresh, and restore measures the file actually on screen at the chosen size.
- **Maximized pin stays maximized across a sound card** — The card itself is not maximized, but the window's maximize state carries across so files after the sound are fitted to the room and restore returns the pre-maximize box.
- **Maximized pin draws every walked file as large as display** — The walk uses the same fit as maximize, so small files are enlarged while the restore glyph is showing.
- **Sound card drawn own size centered** — A sound card no longer fills the previous file's box; showing a card while maximized gives up the maximize.
- **Pinned SVG/font specimen draggable again** — Press is read from the Explorer hook count and drag end from the same level reading, with movement-threshold drag start and no browser image drag stealing the pointer.
- **Dragging pinned animated SVG no longer wedges app** — Browser moves now use a single `SetWindowPos` instead of re-putting the document, bounds only reset on true resize, and boxes are coalesced.
- **Caption buttons redrawn correctly** — Previous/Next are same size, centered and point opposite ways; Open With is a proper open-corner box with arrow, not an upload glyph.
- **Still AVIF with sequence brand recognized; single-image HEIC safe** — The probe reads all ISO brands, so `avif`+`avis` is recognized as still and drawn as first frame, `mif1` HEIC is read as still, and JXL works as naked codestream or container.
- **Run at Startup points at running copy** — Every start rewrites the entry if it names a different copy, compares paths honestly for case/`\\?\`/short names/junctions, leaves arguments and missing entries alone, and never touches `config.ini`.
- **Normalize measures loudness not peak** — It now uses FFmpeg `ebur128` integrated loudness to `-14 LUFS` with true-peak limiting at `-1 dBTP`, so two files of the same loudness are heard the same; already-target files still take the filter, and Video Normalize uses the same measurement.

## [0.3.6] - 2026-09-27

### Added

- **`Audio` previews**: hovering a sound file now shows a card of what it holds — its name, format, sample rate, channels and bitrate — with the sound itself played behind it and a clock and bar under the title. The card follows **Theme** and **Text Size** like the rest of the painted previews, and nothing of the player is on screen.
- Sounds are played by **the media engine Windows has** where its decoders reach the format, and by **FFmpeg's `ffplay`** where they do not. A machine with neither still previews the card.
- **Which engine plays a file is asked of the machine rather than assumed**, and the answer is held per file, so a second hover costs no check. A container that really holds a sound and no picture — an `.mp4` that is really a song — is now answered as the sound it is.
- **`Volume → Video`** and **`Volume → Audio`**: the one volume setting becomes two. Both offer the same levels, and both are marked with their own default: a video starts at `0%` (silent, as before) and a sound at `10%`. A sound at `0%` still shows its card, with the clock still and the bar empty.
- **`Audio`** joins **Preview Types**, and **`Codecs`** gains an **Audio** group. A row that is missing and has a page asks before opening it, as the video and image rows do.
- **`[audio] extensions`** in `config.ini` is the list of what is previewed as a sound, written on first run and read back after.
- Sounds are recognized by their own bytes as well as by their names, so a renamed sound is still previewed as the sound it is.
- **A sound's name scrolls across its card when it has no room for it**, rather than being cut short with an ellipsis.
- `TODO.md` gains **`Unsupported Audio Formats`**: MIDI, the tracker modules, the protected files, the playlists and `.cda` — each with what it would take.
- **`Volume → Audio Seek`**: where in a file a hovered sound starts playing — **Remember** (`default`), **From the Start**, **From the Middle** or **Random**. It is read when a player is started, so a click reaches the next hover rather than moving the sound on screen; a video always plays from its beginning.
- **Remember**ed positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so a file hovered again after a restart is picked up where it was left. Nothing is remembered while the tray is on any of the other three answers.
- A sound's own length is now read on the machine's side as well, so a card on a machine without FFmpeg has a clock from its first frame.
- **`Timing → Prioritize Keyboard`**: with it on, a pointer parked on a file no longer previews of its own while the keyboard is driving Explorer, until the pointer moves, the wheel turns, or a folder change hands the screen over. On by default.
- **`Volume → Video → Normalize`** and **`Volume → Audio → Normalize`**: a file's loudest sample is measured and brought to full scale before it plays, so a folder is heard at one level rather than at each file's own. The sound's row is on by default and the video's is off, and both are greyed out where FFmpeg is not installed.

### Fixed

- A sound no longer goes on playing after the pointer leaves its file.
- A sound card at volume `0%` no longer draws a clock that runs.
- **Sounds played by Windows' own media engine now play at all** — every one of them was silent, and a video it played on a machine without FFmpeg was refused the same way.
- **A sound started part-way through a file now goes round to the beginning** on its next pass, for every format either engine plays.
- A file's name typed in Explorer now previews it, as an arrow key does.
- A preview no longer outlives the folder it belongs to.
- A key pressed onto a file with no preview of its own no longer takes the pointer's preview down and puts it back.
- A sound no longer plays on when the keyboard takes the screen from a mouse hover.
- A preview no longer sticks when Show Desktop is pressed (Win+D).

### Changed

- The `Volume` lists now run loudest first: `100%` at the top, `0%` at the bottom.
- `README.md`, `ARCHITECTURE.md` and `PRIVACY.md` describe the sound kind and what typing a name previews.

## [0.3.5] - 2026-09-26

### Added

- **`Cache → Image (Disk)`**: a budget for developed pictures, each kept as a file under `%TEMP%\rust-hover-preview\image` (`image_disk_cache_mb`, `512` by default), so a camera raw hovered again — or again after a restart — is a read rather than another conversion. `0` keeps nothing between hovers.
- **`Reset to Recommended Settings`** and **`Reset Extension Lists`**: two rows inside the tray's `Config.ini` item, each asking first and each offered only while there is something to put back. The first returns every setting to what this build recommends and leaves the extension lists alone; the second returns the lists and leaves every other setting alone.
- The installer asks the same two questions on a page of its own, with both boxes clear and neither one shown to a silent install — so an update taken with `Auto` changes nothing. What a box asks for is applied the next time the app starts.

### Fixed

- **A hover no longer stalls on the files it is worst at** — a PDF, an archive, a comic, an SVG, a metafile or a font. Those measurements now run on a thread of their own, and the answer is held for the file, so a second hover waits for nothing.
- Repainting a video or an animation at the size of a display no longer costs a pass over every one of its pixels.
- A document's page is no longer stamped on the volume, and the page cache is no longer walked on every hover.
- A hover asks the volume about a file once instead of five or six times.
- The folder behind a view is now read from the view the pointer is in, rather than from every open window and tab.
- The keyboard is polled in one pass instead of two, and the pointer is read once per tick.
- The tray wakes for its own messages rather than a hundred times a second.
- **A hover on a video can no longer end in a spinner that never goes away** — measuring a video is now given a ten-second deadline and killed when it outruns that.
- The file that answers for a run's engines is written when something starts, not when something ends.
- A text preview with long wrapped lines no longer covers the filename it was moved clear of.
- **A folder that was just entered no longer stops previewing until something else happens.**
- A read that came apart no longer leaves its file waiting behind a spinner that never goes away.
- **The launch no longer waits for the housekeeping**, and a hover works from the first moment after it.

### Changed

- Version bumped to `0.3.5` in `Cargo.toml` and `Cargo.lock`.
- `image_cache_mb` starts at `64` rather than `32` and `document_cache_mb` at `256` rather than `128`, so a picture shown at the size of a large display is held at all and one converted drawing no longer fills the folder on its own. An installation that already has a `config.ini` keeps the value it wrote.
- A workbook's fallback page is written as a PNG rather than a BMP: the same picture, a few times smaller.
- The render engine is asked for as soon as the pointer settles on a document it draws, so the first hover of a session overlaps the launch with its own wait.

## [0.3.4] - 2026-09-25

### Added

- **`Calibre` previews** for the ebooks nothing here opens: the Kindle and Mobipocket families, `epub`, `fb2`, `djvu`, `lrf`, `lit`, and the `pml`/`snb`/`tcr` of the dedicated readers. An installed Calibre converts the book and the app draws the first page of the PDF it wrote, exactly as it draws a PDF. Nothing is bundled, the Calibre window is never opened, and a second hover starts no conversion at all.
- **`Calibre`** appears in **`Codecs → Engines`** to show if it is installed, and the conversion runs without a console window.
- `[calibre] extensions` in `config.ini` controls which formats are asked about; an existing installation is given the section and the built-in list on its next run, and a name the engine cannot read is remembered so it costs one conversion and never another.
- Books are recognized by their own bytes as well as by their names, so a renamed one still previews.
- **`Comic` previews**: hovering a `.cbz`, `.cbr` or `.cbc` now shows the comic's **first page** instead of a page of contents. Only that one plate is read out of the container, so a hundred-megabyte comic costs a page rather than an unpacking, and nothing needs to be installed. Plates are ordered by name counting the digits (`2.jpg` before `10.jpg`), and a two-page spread is shown whole.
- **A new `[ebook] extensions` section** in `config.ini`: `pdf`, `pdfa`, `epub` and the three comic containers — one list answering one question, what is a book.
- Comics are previewed at the **book** kind's scale and backdrop, under the same **`Ebook`** switch a PDF answers to.

### Changed

- Version bumped to `0.3.4` in `Cargo.toml` and `Cargo.lock`.
- **`cbz` moved out of `[archive]`** and is now previewed like a comic, and `lit` moved out of `[peazip]` and into `[calibre]` — one name, one answer. An existing `config.ini` is brought up to that on the next run. `chm` stays with the archiver, because drawing a help file takes two to three seconds and a hover should not pay for it.
- There is **no `Calibre TTL`**: the engine is a converter that exits once it has written its PDF, so there is no instance to hold open between books.
- A converted book's page is kept in the same page cache every other engine's page is kept in, named for the book, its version and the engine that drew it.
- `TODO.md` gains `Unsupported Calibre File Format` and `Unsupported Comic Formats`, with what each would take.
- **`Confirm File Type` is gone, and checking a file's own bytes against its name is not optional any more** — every hover asks the bytes what a file is before its name is asked. An installation that had turned it off starts confirming every file; one that had it on sees no change.
- **A still picture and one that moves are told apart by the file's own bytes rather than by its name**, so a renamed file or a `.docx` holding an animated GIF is shown as what it is.
- The head of a file is read once and shared, so the kind of a file, the reader chosen for it and the box it is placed at all come from one read.

### Fixed

- **A preview no longer stays on screen when the pointer is swept over it at speed.**
- **A mouse hover is placed for the hand that actually asked for it**, rather than for a point sampled at the top of the tick.
- **A preview no longer arrives under the pointer** and is immediately dismissed by it.
- **The waiting spinner is kept off the pointer**, so a hand waiting on a file can click and probe through it.
- **A book whose first page says nothing is previewed from the page behind it** — a book is no longer shown as a rectangle of one flat colour.
- A `.pdb` that is a **Mobipocket book** is no longer handed to the render engine, which cannot read one. A `.pdb` that is the other thing the name means is answered exactly as it was.

## [0.3.3] - 2026-09-24

### Added

- **`Engine → AFK Timer`**: an engine that is not marked `Persistent` is let go once no Explorer window has been reachable for the time you pick — 1 hour down to 15 seconds, 1 minute by default.
- A **`Persistent`** toggle at the top of each **`… TTL`** submenu: on, the engine is kept for its TTL whatever you are doing; off, which is how they all start, the AFK timer bounds it.
- New `config.ini` keys: `afk_timer_seconds`, `office_engine_persistent`, `libreoffice_persistent`, `webview_persistent`.
- **`PeaZip` previews** for the archives nothing here reads — `cab`, `iso`, `udf`, `wim`, `msi`, `deb`, `rpm`, `arj`, `lzh`, `hfs`, `vhd`, `dmg` and the rest of the `[peazip]` list. An installed PeaZip lists them and the app shows the same page of contents a `.zip` is shown as; nothing is bundled and a second hover starts nothing.
- A **`Peazip`** switch in the tray **`Preview Types`** submenu, below **`Magick`**; `[peazip] extensions` in `config.ini` controls which formats are asked about, and a name the engine cannot list is remembered so it costs one launch and never another.
- **`PeaZip`** appears in **`Codecs → Engines`** to show if it is installed.
- Archives are recognized by content as well as by name, so a renamed one still previews.
- **PeaZip previews now drive the rest of the tools PeaZip ships**, each for the format only it reads, so `zst` reports how large the stream was before it was compressed where the 7-Zip console leaves that blank.
- `br`, `bcm` and `lpaq8` — single-stream compressors whose tools print nothing about what is inside — preview as the one member an extraction would write, with nothing started at all.

### Changed

- Version bumped to `0.3.3` in `Cargo.toml` and `Cargo.lock`.
- The update installer is downloaded only after you click the update row and confirm, not as soon as a newer release is found; a check now costs one small request.
- The update check now runs at every startup as well as when the tray menu is opened.
- The update prompt now asks with three answers: **`Auto`** installs the update, **`Manual`** opens the release page in your browser, and **`Cancel`** does nothing.
- **The `Cache` submenu has two entries** where it had five: **`Image (RAM)`** and **`Document (Disk)`**. **Text** and **Ebook** previews are no longer cached at all — they are the cheapest things here to make again and the dearest to hold.
- **The pages both document engines draw are one cache, on disk.** `document_cache_mb` replaces `office_cache_mb` and `libre_cache_mb`, and a `config.ini` holding either older key is read once for the larger of the two and written without them. A page outlives the run and the engine that drew it, and is given up by when it was last _read_, not when it was converted.
- The LibreOffice engine's own profile and stub document moved out of the folder that is deleted at startup to `%LOCALAPPDATA%\rust-hover-preview\libreoffice`.

### Fixed

- A restart no longer hides an available update until the hour is up.
- The update row appears whether or not the installer could be fetched, and a download that fails says so.
- An installation made from a pre-release is offered the stable release its version names.
- A preview of an archive PeaZip listed is no longer drawn stretched to the display instead of at the size the layout planned.
- Hovering a file whose preview an engine makes no longer leaves the spinner up for good when the file is left and taken up again while that engine is still working.
- The `Cache` budget is evicted by when a page was last read rather than when it was converted, which is what "the oldest first" was always meant to be.
- A document an engine would not draw is no longer refused for good: the mark ages out after two minutes.
- The tray's `Cache` submenu no longer says everything it sizes is held in memory, which stopped being true when converted pages landed on disk.

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
- A preview no longer stays on screen when the pointer leaves a **clipped item** — the last row of a scrolled `Extra large icons` or `Large icons` view — for the toolbar above it or past the bottom edge of the window.
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

- Opening the tray menu can check for updates at most once an hour. If a newer version exists, the installer downloads and a row appears above Run at Startup; clicking it asks whether to install.
- `Timing → Settling Delay`: how long the pointer must be still before a preview may open for anything, 0 ms down to 1000 ms, and 0 ms by default. Written to `config.ini` as `settling_delay_ms`.
- `Engine → Select Engine → Office`: which engine an Office document's page is asked of. **Microsoft Office** is the default and asks the application that owns the format; **LibreOffice** asks the render engine for every Office document. The LibreOffice row is greyed out where no LibreOffice is installed.
- `Engine → LibreOffice TTL`: how long the LibreOffice engine is kept after the last document it converted — 10 minutes by default. The next document is handed to an engine that is already up rather than paying for another start: measured on one document, 1.2 s against 0.2 s. `0 seconds` keeps no engine at all.

### Changed

- `Performance → Office Engine TTL` is now `Microsoft Office TTL`, and `SVG Engine TTL` is now `WebView2 TTL`: the same settings under the names of the engines they are about.
- `Select Engine` and the three engine TTLs sit in a new `Engine` section below `Performance`, and `config.ini` is grouped the same way. No key changes and no value moves.
- New Libre preview type for documents LibreOffice can read but this app cannot, such as CorelDRAW `cdr` and older office formats. Previews stay sharp when resized, and use new scaling and a 32 MB cache.
- CorelDRAW `cdr` previews no longer use the small embedded image, and without LibreOffice installed they show no preview.
- Office documents use LibreOffice when Microsoft Office is not installed, under the same Office type.
- LibreOffice conversions run in the background: hovering no longer freezes the preview, tray, or pointer.
- A stuck LibreOffice conversion is ended instead of blocking everything; one bad file no longer breaks later previews.
- The `[libre]` list lost 25 file types LibreOffice does not support; supported names can be added back by hand. `.swf` is now video-only.
- Previews follow a file's content rather than its extension, and formats the signature table does not carry are now recognized by their own bytes — so a renamed file is previewed as what it is whatever it is called.
- File types whose bytes carry no signature are routed by extension instead, and `.pdb` is told apart: a Palm OS ebook is drawn by LibreOffice, a compiler's program database is left alone.
- `Confirm File Type` is on by default; an installation that already exists keeps the value its `config.ini` holds.
- Building from source now needs Rust 1.98.1 or newer.
- Source files are organized into folders by role. No behavior change.
- **A new file under the cursor now gets its preview while the mouse is still moving**, instead of waiting for the pointer to stop first. `Timing → Settling Delay` asks for the old behavior, at any length.
- `Timing → Delay` and `Timing → Rehover Delay` offer more steps — 0 to 1000 ms in 15 steps — in place of Instant, Fast, Medium, Relaxed and Slow. A `config.ini` holding the removed 750 ms step keeps the value.

### Fixed

- A preview of an Office document no longer shows a smaller version of itself first on a machine that also has LibreOffice installed.
- Flash animations (`.swf`) no longer hang previews; they are in the video list only now.
- File types are checked in the same order everywhere, so videos are no longer mistaken for documents.
- A file renamed to another type's extension is now drawn by the whole of the type its content belongs to, size rules included.
- A file whose content is not a document no longer starts Office, and a video whose name is not in the video list is played rather than left on its first frame.
- **Files preview on one monitor while a maximized or fullscreen window is in front on another.**
- A video player left behind by a crash, a forced close, or a Task Manager kill is ended on the next run.

## [0.2.14] - 2026-09-23

### Changed

- New defaults, for new installations: previews are placed at their best position, appear instantly, and wait 200 ms before the same file previews again. Transparency is shown over a checkerboard rather than black, and `DDS Background` offers only Black and White, starting at white. An installation that already exists keeps the values its `config.ini` holds, so if you see black where you expected a checkerboard, that is your file and not a bug.
- `config.ini` is grouped under the same headings as the tray menu, so a setting is easy to find, and the tray marks the item each setting starts at with `(Default)`.
- `config.ini` is checked and tidied every time the app reads it: a missing setting comes back, a value the app cannot read is replaced with the value it is actually using, and any line that is not one the app writes is removed. A list you edited yourself is left exactly as it is.
- Settings that older versions stored under a different name are no longer carried over — `svg_scale`, `svg_background`, `svg_preview_enabled`, `off_trigger_key`, `avoid_filename`, `transparent_background` — so their lines are removed and the settings go back to their defaults.
- The installer no longer writes into `config.ini` and no longer adds the startup entry: the app does both on its first run.
- Bump version to 0.2.14 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A CorelDRAW document shows its drawing instead of a blank white page.
- Editing `config.ini` by hand now takes effect as soon as you save it, and a value the app cannot read is written back as the value in use.

## [0.2.13] - 2026-09-22

### Added

- Preview for design documents: Photoshop `psd` and `psb`, Krita `kra`, OpenRaster `ora`, and the project containers other drawing tools save (`sketch`, `fig`, `xd`). A layered document is shown from the finished picture its format keeps inside it, so it arrives as it was saved.
- Preview for Illustrator `ai` files, drawn as a PDF the same way a PDF is.
- `Design` under the tray's **Preview Types**, with `design_preview_enabled` and a `[design]` extension list in `config.ini`.
- `Design Scaling` under the tray's **Scaling** menu: how much of the screen a design preview covers — Fit to Screen (default), or 75%, 50%, 25%, 10%.
- `Design Background` under the tray's **Background** menu: what a design preview is drawn over, black by default.
- Vector previews: Windows metafiles (`wmf`, `emf`) and Illustrator files saved as encapsulated PostScript (`eps`, `epsi`), drawn by Windows at the preview's size so they stay sharp. An `.eps` shows the picture its writer saved inside it; one without such a picture shows nothing.
- `Vector` under **Preview Types**, with `Vector Scaling` and `Vector Background` in the tray, and `vector_preview_enabled`, `vector_scale`, `vector_background` and a `[vector]` extension list in `config.ini`.
- `ai` files that are not PDFs show the preview saved inside them, under the **Design** kind.

### Changed

- `Vector Scaling` now starts at Fit to Screen rather than half the display.
- `svg` and `svgz` moved from the `[image]` list to the `[vector]` list, where the other drawings are; existing `config.ini` files are updated on the next run.
- Bump version to 0.2.13 in `Cargo.toml` and `Cargo.lock`.

### Fixed

- A name in the text list _and_ in a drawing or document list is previewed as that kind again, not as text.
- SVG previews are back: a document is drawn by the engine again.
- A Photoshop document smaller than the screen shows a preview again, instead of being refused whenever the preview was enlarged past the document's own size.
- EPS files whose preview is a palette picture now preview.

## [0.2.12] - 2026-09-22

### Changed

- The scaling options now live in a `Scaling` menu of their own in the tray, right below `Placement`.
- `Font Background` now defaults to white; `font_background` in `config.ini` still sets it.
- Tray labels renamed for consistency: `Image Scaling`, `Video Scaling`, `Avoid Nothing`, and `Videos` in `Codecs`.
- Preview hover checks now put far less load on Windows Explorer, so the file list stays responsive while previews are running.

### Fixed

- Keyboard previews in `Details` and `Content` views now follow `Avoid Filename` and `Avoid Filename Column` instead of behaving like `Avoid Details`.
- A failed startup of the Explorer lookup could leave hover previews off for the rest of the run; it is now retried until it works.
- The internal timeout that stops a stalled Explorer from blocking the app is now verified instead of assumed.

## [0.2.11] - 2026-09-21

### Added

- Video previews without FFmpeg: where `ffplay` is not installed, videos are decoded by the media engine Windows already has, played in the same preview window as everything else, with sound and looping. **FFmpeg is now optional rather than required.**
- `Codecs` in the tray menu: what this machine has of every engine and codec a preview can lean on, with a check or a cross per row and the missing ones greyed. The rows are for reading only — nothing in the menu does anything, and there is no setting behind it.
- Sample lines for Arabic, Hebrew, Thai and Devanagari, so a font of one of those scripts is drawn as its own script rather than by the first characters its map happens to hold.
- `Font Face` under the tray's **Placement** menu, with `ttc_face` in `config.ini`: which face of a `.ttc` collection a specimen is drawn from. The specimen's heading says which face came out — `(2 of 4)` — so what was asked for and what was drawn cannot be taken for each other.
- DDS texture previews now support almost all DDS formats, including compressed, uncompressed, cubemaps, texture arrays and mip chains; only the first face and first level are shown.
- Signed DDS textures now display correctly: zero is in the middle, and missing color channels show as neutral gray.
- Signed HDR DDS textures now get a preview and use the same HDR tone mapping as other light-based images.
- New `DDS Background`, `hdr_tone_map` and `hdr_exposure` settings for how DDS textures and HDR images are shown, and a new `spinner_delay_ms` for how long to wait before showing the loading spinner.
- New `Videos Scaling` under the tray's **Placement** menu, with `video_scale` in `config.ini`: how large a video is shown, 100% by default.
- New `Animated Scaling` beside it, with `animated_scale` in `config.ini`: how large an animated GIF, WebP or PNG is shown. A check of the file itself tells an animation apart from a still, and a still follows `Images Scaling` like any other picture.

### Changed

- `README.md` describes what FFmpeg adds rather than requiring it, and lists the free Microsoft Store codec extensions.
- Bump version to 0.2.11 in `Cargo.toml` and `Cargo.lock`.
- A specimen's lines carry their own direction, so the Arabic and Hebrew lines are laid out from the right; every other line is drawn as it was.
- A pan-script font is drawn at smaller type rather than past the bottom of the specimen's box.
- Video files are now checked in the background with a spinner, and the result is remembered so they are not checked again every time.
- Video previews show the spinner until the video player window appears, and players that fail to open or hang are closed automatically.
- The loading spinner is now the same small pointer spinner for every preview type, not a preview-sized box, and its delay is unified at 250 ms.
- EXR and HDR previews are now tone mapped instead of clipped, so bright areas fade naturally instead of burning out.
- DDS previews now use the best detail level for the displayed size, and truncated files are handled safely using the real file length.
- DDS is now in the built-in image list, and older config files are updated automatically.
- Animated GIF, WebP and PNG previews follow their own `Animated Scaling`, separate from still images; a `config.ini` written before this gets the animation scale from `preview_scale`.
- `Font Face` is out of the tray menu; `ttc_face` in `config.ini` still picks the face of a `.ttc`.

### Fixed

- **Animations no longer stop at the end of their first play.** This affected GIF, animated WebP and animated PNG previews.
- **Animations with a long frame in them no longer freeze on that frame** — a long hold is now simply a long hold.
- An animation that fits the memory it is kept in stays whole instead of being taken apart frame by frame.

## [0.2.10] - 2026-09-21

### Added

- Font previews for `ttf`, `otf`, `ttc`, `woff` and `woff2`, drawn by the WebView2 engine the SVG previews already use: the name the font calls itself, the pangram _The quick brown fox jumps over the lazy dog._, and a line each for Japanese, Chinese, Korean, Cyrillic and Greek that the font's own character map covers. A line is drawn only where every character of it is in the font, so nothing is ever shown in a system font and passed off as the font.
- `Font Scaling` under the tray's **Placement** menu, with `font_scale` in `config.ini`: how much of the screen a specimen is drawn over — `Fit to Screen`, or 75%, 50% (the default), 25%, 10%.
- `Fonts` under **Preview Types**, and a `[font]` extension list in `config.ini`; both are written on first run.
- `Font Background` under the tray's **Background** menu, with `font_background` in `config.ini`: a specimen is a page of text, so its ink follows the backdrop.
- A `.ttc` previews its first face, which is written out as a font of its own beside the browser's profile folder.

### Changed

- The released executable is now **867 KB smaller** (11.5%) from three build changes taken together: deploy artifacts build with a `github` profile, still WebP is decoded by the Windows codec where it is installed, and SVG previews are drawn by WebView2 rather than in-app.
- `.exr` previews decode single-threaded, and GIF is unified on the same decoder version the rest of the image pipeline uses.
- SVG preview is dropped on machines without WebView2, since no fallback reader remains.
- Show the spinner immediately on cold-engine document hovers, and not at all when warm, and move the engine window with the pointer wait so documents draw where the wait ended.
- Stand down failed engines for one minute instead of five, costing only that hover and clearing its spinner.
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
- The engine keeps its state in a folder per run, cleared at startup; one that fails is retried, then stood down for five minutes so this app's own reader plays the document.
- `trigger_key_enabled` in `config.ini` and a check in the tray `Trigger Key` submenu, to switch the trigger key off without changing its mode.
- `svg_background`, the backdrop an SVG document is drawn over: Transparent, Black, White or Checkerboard, the same four a picture is offered.
- `svg_scale`, how much of the screen an SVG document is drawn over, in `config.ini` and as `Placement → SVG Scaling` in the tray: `Fit to Screen`, or `75`, `50` (default), `25` and `10` percent. A vector is drawn at whatever size it is asked for, so a document's percentage is of the room the display has rather than of the size the file asks for.
- `svg_preview_enabled` and an `SVG` entry under `Preview Types`, below `Office`: a document is its own kind even though `svg` and `svgz` are entries of the image list, so the two are switched independently.
- `decode_budget_gb`, what one hover may decode or read for in gigabytes (default `1`), in `config.ini` and as `Performance → Decode Budget` in the tray: a file past it gets no preview instead of memory the app may not get.

### Changed

- Installing over an older version no longer asks: the previous version is uninstalled first and a running copy of the app is closed instead of prompted for.
- Video previews no longer write `%TEMP%\rust-hover-preview-video.log`, and one left by an earlier version is deleted at startup.
- The tray's `Background` submenu is above `Volume` and holds an `Image Background` and an `SVG Background` half; `transparent_background` is read as both `image_background` and `svg_background`, and written out under the two names.
- Disabling Office/SVG now kills its engine immediately, whether toggled in UI or config.ini, and re-enabling Office previews it without waiting out backoff.

### Fixed

- **An animated SVG no longer stops moving once the engine has been let go for idle** — every document after that was left on its still frame for the rest of the run.
- The WebView2 browser no longer outlives the app when the app is killed, and one left behind by an earlier run is ended by the next launch.
- **Every process the app starts is now put in a Windows job object**, so a crash, a kill from Task Manager or a logoff ends them with the app, and the next launch ends whatever the job could not take. Nothing is acted on by id alone: a record is used only when the process still carries the image _and_ the start time it was recorded with.
- One engine per Office family is enforced rather than assumed, and exiting while a document is mid-render no longer leaves the engine behind.
- **Pictures over 40 megapixels preview again**: the pixel cap is gone, and every reader now asks the decode budget above before it allocates.

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
- Each family keeps its own Office engine, so a folder holding a document, a workbook and a deck starts each application once rather than quitting one and starting another every time the pointer crosses between them.
- The settings a render needs of an Office instance are taken for that render and put back when it is over, instead of being held for as long as the engine lives — which is what makes `Indefinitely` safe for an instance that is the user's own Word or Excel.
- PDF and Office pages use fit-to-screen sizing, but preview scales below `100%` now reduce it.
- The waiting spinner is now a transparent haloed arc, appears immediately, follows the pointer, and is placed flush at the pointer's corner.
- A page arriving over an existing preview replaces it without taking the preview down first.
- `image_cache_mb` defaults to `32` rather than `64` and `pdf_cache_mb` to `32`, so the decodes a folder is swept back over are already done; the text cache still starts at `0`.
- Removed the `Office Preview` submenu; `office_render_enabled` is no longer read.
- An Office engine let go because it refused a page is ended where it stands rather than asked to quit, so the retry reaches the fresh instance without the wait in front of it.

### Fixed

- PowerPoint decks and Excel workbooks on printerless machines now preview.
- Office engines no longer survive quit, and failed or unresponsive instances are replaced.
- A preview that crossed a display boundary is now laid out again at the new display's scale and put back, rather than staying gone until the pointer moved.
- Non-document files no longer start an Office engine.
- `config.ini` is written in a fixed order instead of being reshuffled on every save.
- A page landing over a waiting spinner no longer flashes the spinner stretched across the page's box.
- A pointer that drifts onto a waiting spinner no longer dismisses the hover it belongs to, so the page being waited for is not thrown away with the hover.
- Explorer crashing or being ended no longer leaves previews dead until the app is restarted.
- Office's own "Publishing…" progress window no longer flashes over a hover.

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
