# Changelog

## [Unreleased]

### Added

- **An animated JPEG XL now plays** instead of showing its first frame. It is decoded by a JPEG XL decoder compiled into the app, so it needs no Store package at all, and it plays through the same machinery as a GIF, an APNG or an animated WebP, so **Animated Scaling** and the **Images** gate apply to it. An animated AVIF and a HEIF image sequence are still drawn as their first frame — see **Fixed** below.
- **`Scaling → Text Scaling`** (default **Fit to Screen**): caps how much of the screen a text preview box may take when it opens — a plain text file, code, or Markdown — while a short file still gets the small box its own text needs. `Text Size` still scales the text inside that box.
- **`Text Preview → Render HTML`** (default **Off**): on, a `.htm` or `.html` file is previewed as the page it holds — run by the browser engine, at the share of the screen **Document Scaling** names — rather than as its markup. A page is the one thing this app hands to a browser whole rather than as a picture, and it is the one thing it runs: a page that draws itself with script — a WebGL canvas, a game — has nothing to show without a run, and under the old rule it came up a blank black rectangle. What a page can still do is bounded, and the bound is the same one an SVG is drawn under: the frame a page is shown in is a sandbox that keeps the page to itself — no forms, no popups, no navigation of the top frame — and nothing a page links to is fetched, so a page that reaches for the world reaches nothing. SVG documents and font specimens are unchanged and are still drawn rather than run, because a browser is handed those as an image and an image never runs. A machine without the engine keeps the text preview, as it does for drawings and fonts.
- **`Background → HTML Background`** (default **White**): what a page of HTML is drawn over. It was drawn over whatever **Vector Background** was set to, so a page arrived over the checkerboard unless you had asked for something else. It now offers **White**, **Black** and **Checkerboard** — three of the four backdrops, because transparency is the one a page is not drawn over, and a page is a page: it brings its own colours, and what stands behind it is the page to read it against. Written to `config.ini` as `html_background`; a file naming `transparent` from before the setting existed reads as **White**. **Vector Background** still decides for SVG documents and metafiles, and **Font Background** for a specimen.
- **The pin key now brings a minimized pin back instead of closing it**: pressed while the pin is a bubble, with Explorer in front, the key puts the window back up on the file you have picked since the bubble went down — the same file, in the same place and the same size, `Pin Mode → Update Preview` would have shown it in. With nothing picked since, the key does nothing rather than losing the window. A right-click on the bubble closes it.
- **A pinned window's caption now carries three more buttons**: **Previous** and **Next** step the pin through the sibling files in the folder it was taken up in — that folder only, never a subfolder of it — in the order the Explorer listing is showing them, and **Open With** opens the file in the application Windows has registered for it. Only files this build can actually preview are stepped onto, so a kind switched off under **Preview Types**, or one with no engine installed, is not a step, and the two ends wrap around. A file that is a step but will not open — a corrupted one, a OneDrive placeholder still in the cloud, a kind with no reader, a player that would not start — is stepped over rather than stopped at, so one press is always *the next file there is to look at*; see **Fixed** below. Minimize, maximize and close keep the exact places they have always had; a caption too narrow for the three new ones drops them whole rather than crowding the window's own buttons. The two step buttons work with `Pin Mode → Update Preview` switched off — they are a thing you pressed, and no setting asks for them.
- **The caption now has an `Open With...` button beside `Open With`**, which asks which program to use: it opens the same Windows "How do you want to open this?" dialog you get from Explorer's own context menu, listing the programs installed on this machine that could open the file — the way out of a pin and into a program of your choosing when the registered default is the wrong one. The pinned window steps out from on top of the desktop for as long as that list is standing, so the list can be reached, and takes its place back on top as soon as the list has gone.
- **The two hand-off buttons now say what they are** once the pointer rests on them: **Open With** names the program that would open it (*Open With Adobe Photoshop*), which is read once when the window is pinned rather than asked of the machine on every repaint, and **Open With...** says *Open With...*. A machine with nothing registered for the format has no program to name, so that button says nothing. The name hangs below the caption, over the picture, rather than inside the strip — a name that wide would cover the buttons it describes. Every other button's glyph already says what it is and is left to say it that way.
- **The Left and Right arrow keys walk a pinned window's files**, exactly as the caption's **Previous** and **Next** do, and — like them — whether `Pin Mode → Update Preview` is on or off, since they are a thing you pressed rather than a file you picked behind the window. `Up` and `Down` walk the same way as `Left` and `Right`. What a key does is decided by which window the keyboard is in, so a pinned window has to be clicked before its own keys answer: see **Fixed** below.
- **A pinned window now shows a loading spinner of its own while it is being shown another file.** The read and the decode for a swap run on a thread of their own rather than on the thread that draws, so a file that takes longer than `spinner_delay_ms` — a large picture, a slow disk, a video still being measured — is no longer a window frozen for the length of it. The pin keeps showing the file it already has, which is the pin until the new one has answered, and the arc is drawn *over* that file in the middle of the window rather than in place of it, in the same visual language as a hover's spinner and on the same delay, so **Timing**'s `spinner_delay_ms` now covers a pinned swap as well as a hover: `0` puts the arc up at once, and a file that loads inside the delay still shows none. The window is not hidden while it waits, and the caption and bar are drawn after the arc, so a window showing a slow file is still a window you can read and a caption you can still see.
- **`Pin Mode → Nav File Types`** (default **All**): how wide that walk is. **All** steps through every file this build could preview, so a video sits beside a sound; **Category** narrows the walk to files of the pinned file's own kind of thing — a camera raw is a picture and a book a document, because that is what you call them. It is written to `config.ini` as `pin_nav_file_types` under the **General** heading, and a file written before the setting existed reads as **All**.

### Fixed

- **A step that lands on a file the pin cannot be shown now steps over it.** **Previous** and **Next** (and the arrow keys) used to stop dead on a file that would not open — a corrupted one, a OneDrive placeholder still in the cloud, a kind with no reader, a player that would not start — with no answer at all: the window kept showing the file it already had, the caption went on naming the old one, and pressing **Next** again did exactly the same thing again. One press of a caption button is one gesture, and the gesture is *the next file there is to look at*, so the walk now carries on to the next file and keeps going until it finds one it can show. The walk is bounded by the folder it is made of, so a folder of nothing this build can read ends the walk rather than going round for ever, and a file *you* picked in the Explorer listing is still not stepped over: a file you named that cannot be shown leaves the window on what it had.
- **A pinned window playing a sound no longer closes itself at the end of every pass.** The window asked each tick whether the thing behind it was still there, and for a sound played by FFmpeg's player it read the process — but a sound started part-way through with a seek plays one pass and stops, and the question ran before the tick that starts the next one, so it read the finished player as a window onto nothing and closed the pin over it. The card a sound is drawn as is this app's own text and the player behind it is this app's own, so a sound is no longer asked about at all: a player that is not running is the moment between two passes, and a card whose player never came back is a card with a still clock, which is what a machine with no sound card gives. A pinned **video** whose player has gone, and a document whose browser has gone, still come down as before.
- **The caption's two step buttons now draw their arrows as chevrons**: each of **Previous** and **Next** was drawn as a cross, because the two arms of the mark were measured from the middle of its span and opened on both sides of it, so both buttons showed the same glyph. The point is now at the end each button walks off, and the two point opposite ways, as they should.
- **A still `.avif` written with a sequence brand behind it is now recognised as one.** The ISO base media probe now reads the file's whole list of brands rather than only the one at the front, which is what an animated AVIF is ordinarily written as — `avif` in front, `avis` behind it. It is recognised rather than played, though: the media engine Windows has was asked directly about such a file and refused it, and refused a `msf1` HEIC the same way, while an AV1 `.mp4` and an HEVC `.mp4` through the same calls came back with frames on the same machine — so the codec is installed and what is missing is a demuxer for the image-sequence brands. Both are still drawn as their first frame, exactly as before. `TODO.md` records what a decoder for them would take.
- **A single-image HEIC is no longer at risk of being treated as a sequence.** The `mif1` brand is the generic HEIF image brand an ordinary HEIC is written under rather than a sequence brand, and is now read as the still picture it is.
- **A `.jxl` is recognized in both the forms it arrives in** — a naked codestream and a container — which it was not before; a file whose front settles nothing was left to the tables.
- **The caption's buttons are now all drawn the same size, and centred on their own**: the step buttons' arrows reached a whole span either side of their middle row, so **Previous** and **Next** came out twice as tall as the close cross, the maximize box and the restore pair beside them; and on a scaled screen their two arms no longer met at the point, which is the one thing a chevron is. Both are now the size of the glyph square they are asked for. **Open With** is redrawn as well — it was a box too small to read as a box with a large arrow standing on top of it, which reads as *upload this*, and its middle sat three pixels right of its own button's. It is now a full-height box with its open corner left off and an arrow leaving that corner, which is what every "open in" button is drawn as. This is a later defect than the chevron fix above: the arms were corrected there, but how far they reached was not.
- **`Pin Mode → Update Preview`**: a sound's card is now drawn at its own size, in the middle of the box the window already stands in, rather than filling whatever box the file before it left — a picture's, say, which turned the card into a slab. A window shown a card while maximized now gives the maximize up rather than carrying a restore no card offers a button for.
- **A maximized window no longer shrinks away as you walk its files, and no longer stretches the file you restore down onto**: the box put aside for the way back belonged to whichever file was on screen when the maximize was pressed, and was never re-based when the window walked to another one. Walking next through a folder of mixed shapes fitted each file into the box the one before it left, so a maximized window shrank a little with every press until it was the smallest shape in the folder — the whole way under a caption still drawing the restore button, because nothing ever gave the maximize up. Each file is now fitted to the display afresh, so a walk through shapes is one window of one size showing something else, as every other step already was. Restoring down now measures the file actually on screen rather than playing back the remembered box verbatim, so a window that walked to a differently shaped file comes back in a box of that file's shape rather than stretched into the old one's — at the size you had chosen, which a restore still does not change.
- **The arrow keys no longer walk a pinned preview unless the keyboard is actually in it.** Click the pinned window once and it takes the keyboard the way any window does; `Left` and `Up` then step to the previous file, `Right` and `Down` to the next, exactly as the caption's **Previous** and **Next** do, and `Escape` closes it. Click into another window and the keys belong to whatever is in front again. Before, a pin was a window that could never be focused, so whether a key belonged to it could not be asked and was guessed at — and the guess walked the pin while the keyboard was in *another* program entirely, so a `Left` aimed at your editor stepped the preview as well.
- **Explorer no longer moves its own selection while a pinned preview is on screen.** The arrows were read straight off the keyboard's physical state rather than asked of the window that received them, so one press walked the folder listing *and* the pin — which is what "why does the Explorer selection change when Explorer is not in focus?" was. They are now answered as keystrokes Windows sent to the pinned window, so a listing behind a window you are working in stays where it was. As a consequence, `Pin Mode → Update Preview` no longer follows the keyboard while a pinned window holds it, since the keyboard is genuinely not in the listing until you click back into Explorer — which also stops the pin changing file under the very keys you aimed at it.
- **The pin key no longer hides a pinned preview**, in any situation. Pressing it while a pin is up now does nothing to the pin, wherever the keyboard is: previously it took the pin down when it looked like the key was meant for it, which a `Space` typed into whatever the pin appeared over could easily satisfy. The key still brings back a minimized pin — the bubble — as it always did, with Explorer in front. A pin is ended by the caption's close button, a right-click on the bubble, `Escape` in a window that holds the keyboard, or Alt+F4 on one you have clicked (which used to destroy the preview window outright, leaving no previews for the rest of the session).
- **A left-click on a minimized pin's bubble brings the window back instead of closing it.** The click and the drag of a bubble were both being read as the hand being on the pin, and a press was a close; only a right-click closes a bubble now, as the setting's own description has always said.
- **Clicking a pinned window takes the keyboard.** That is what makes its own keys answer, and it is also what commits an inline rename in progress in Explorer behind it, so a name being typed there is finished by the click. A pinned window also stays out of Alt+Tab — it is a tool window, as it always was — so after alt-tabbing away you click back into it rather than switching to it.

### Changed

- **A page previewed as a page can now be pointed at and typed into.** The pointer arriving on a running page used to count as arriving on the preview and take it down before a click could land, so no page could be touched at all; the page's own rectangle is now a place the hand is held, and a drag that began on the page survives the pointer leaving that rectangle — an orbit carries the pointer well outside the box the page was drawn in. Clicking into a page also gives it the keyboard, where a document, a specimen and every other preview still refuse it: a preview never takes the caret, and this one does not either, since showing a page is not a click and the page takes the keyboard only where the user clicked. A picture or a video is dismissed by the pointer exactly as before, and this applies only to a page that runs. The pin key needed no change: pressed while something other than Explorer is in front it is already answered as another program's key and does nothing. While the page is in front the app's own pointerless gestures are still read — the navigation shortcut under a modifier, a `Backspace`, a `T`, and Explorer's type-ahead — because the read is what spends the press, but nothing they would have done is done. What a page still cannot have: sound of its own, the files beside it, fullscreen, a right-click menu, the tools or a find bar.
- **Dragging a maximized pinned preview now gives the maximize up, the way resizing it already did.** Carrying a maximized window used to take the maximize with it: the window kept the size the restore had put aside, took the new place the hand left it at, and the caption went on drawing the restore glyph over a maximize that was really a box the hand had taken hold of. Any drag that actually moves the box now ends the maximize instead — carried to another place or pulled to a size, it is the same answer either way — so the caption's restore glyph reverts to a maximize and the button maximizes the window again, the file being fitted to the screen as it is for any other. A press that moves the window by nothing is not a drag and leaves the maximize standing. This is a deliberate change of behavior rather than a fix: the two are now one rule, so what a hand takes hold of is never a window there is still a way back from.

## [0.4.0] - 2026-09-27

### Added

- **`Pin Mode → Enable (Space)`**: press the key while a preview is up and the preview becomes a window of its own — captioned, movable, always on top, and still there when the pointer leaves. The key is shown in the menu and set by `pin_key` (`space` by default; `pin_enabled` turns the feature off).
- A pinned window's **caption** carries minimize, maximize and close. Minimize collapses it into a round bubble that can be dragged anywhere and clicked to bring it back; right-clicking the bubble closes it. Maximize fits the media to the screen, centered, keeping its shape.
- A pinned window is **movable by the picture** as well as by the caption. The exceptions are a text preview's own text and scrollbar, and a video FFmpeg's player has.
- A pinned **video** gets a transport bar of this app's own: play/pause, a draggable seek bar and the two clocks. A video the media engine plays can be paused and seeked; one FFmpeg's player has shows a read-out bar instead, because that player can be told nothing.
- **Previews are held back while a pin is up**: nothing is raised or dismissed until the pin closes.
- **Videos now play through the media engine Windows ships with wherever it can decode the file**, and through FFmpeg only where it cannot — the reverse of the previous order. Whether a file can be played is asked of the machine once per file, so a hover never asks twice.
- **The video format list is now two lists in `config.ini`**: `[video]` is what the media engine is asked to play, and `[ffmpeg]` is what only FFmpeg's player reads. An existing `config.ini` is divided between the two on the next run, and a list you have edited is left as it is.
- **A pinned picture now wears its caption and bar on top of it** rather than in bands above and below, so the window is exactly the size of the picture and no longer carries a strip of empty window. Both fade in when the pointer arrives and fade out when it leaves, and **only the strip under the pointer is shown** — the other stays out of the way.
- **A pinned video has a volume of its own**: a speaker button on its transport bar opens a knob over the video, and the level it is set to belongs to that window alone, independent of **Volume → Video** for hovers. The knob stays up while the pointer is on it and closes when the pointer leaves.
- **`Pin Mode → Update Preview`**: a pinned window can now be shown the file you pick while it is up — one you click, or one the keyboard selects — in the box it already has, where it stands. **Enabled** (`default`) turns this on; **On Hover** adds the pointer's own hover as one of the ways a file is picked.
- **`Pin Mode → Pause Preview`**: a window collapsed into its bubble now **pauses what it was playing** instead of going on behind a bubble nobody can see, and starts again at the second it stopped at when the window comes back. **Audio** and **Video** are switches of their own, both on by default.
- **`Timing → Trigger Key → Affect Pin Mode`**: off — where the app starts — the key that holds back hovers is not read while a preview is pinned; on, holding it brings the pin down along with the previews it stops.
- A video's **letterboxed picture is cropped to what is in it** whichever engine plays it, so a black-barred file fills its box instead of growing bars inside it.
- A video shown larger than its own picture is now **scaled up to fill the box**, rather than drawn at its native size in the middle of it.

### Fixed

- A pinned window's media now **scales with the window as an edge is dragged**, instead of standing at its old size inside the new box.
- A pinned window can be **resized from every edge and corner**, and made smaller as well as larger — previously the top edge and both top corners moved the window instead of resizing it, and a picture could be pulled larger but never smaller.
- The pointer's shape now **matches the edge under it**, and returns to the arrow when it is over none.
- **Pinning a video or a sound no longer freezes the app.**
- The caption of a preview pinned at the top of the screen is **no longer drawn off the top of the screen** with its buttons.
- The preview loop now **routes every video the way the loader does** — on a machine with FFmpeg installed, every video used to be played by FFmpeg's own window, so a pinned one could not be resized, maximized or seeked.
- The check that asks whether the media engine can play a file **now returns the right answer** for ordinary `.mp4`, `.mkv`, `.webm` and `.avi` files, which were all being turned down whatever the machine could decode.
- A file the media engine accepts and then cannot draw **falls back to FFmpeg's player** instead of previewing as an empty box.
- The **seek bar of a pinned video works again** at any point in playback, not only in the first three seconds, and dragging it to the far left now starts the file from the beginning.
- The **seek bar follows a seek made while a video is paused**, instead of springing back to where the pause began.
- A pinned video **no longer flashes a sheared picture** when a resize is released.
- Dragging a pinned video is **faster**: moving it no longer repaints the whole window on every pointer move, and resizing it no longer resamples it a pixel at a time.
- A video played by the media engine **no longer flashes a blank backdrop** for a moment at the start of a hover.
- **A pinned window comes back centered** where it is when you bring it out of the bubble, rather than jumping back to where it was before it went away, and a maximized one restores around the center of the screen it is on.
- **The bubble always lands in the same place**: it is put on the minimize button rather than on the window's corner, so repeated collapses no longer wander, and it no longer remembers a position from an earlier move.
- **Dragging a bubble no longer stutters or overshoots**, and it stays under the hand where it was grabbed rather than jumping to its middle.
- **A pinned window being shown another file no longer shrinks as it goes**: the new file is fitted to the window's own longest side, so a portrait pin followed by a widescreen file is not walked down a size at a time.
- A pinned window **kept where the hand put it** when it is shown another file, rather than being placed back beside the cursor.

### Changed

- Version bumped to `0.4.0` in `Cargo.toml` and `Cargo.lock`.
- **`Text Preview → Full Mode` is gone from the tray, and `text_preview_full_mode` with it**: a pinned text preview now comes up in full mode instead, so it scrolls and its text can be selected and copied.
- The play triangle a paused pinned video was drawn with is gone; a paused pin now shows the last frame and nothing over it.
- A pinned picture's **caption and bar now fade in and out** as the pointer arrives and leaves, instead of appearing and vanishing outright.
- `pin_enabled`, `pin_key`, `pin_update_enabled`, `pin_update_on_hover`, `pin_pause_audio`, `pin_pause_video` and `trigger_key_affect_pin_mode` are new keys in `config.ini`; a file written before them keeps the behavior it had.
- The `video_engine_probe` diagnostic (`RHP_VIDEO_PROBE`) now plays the file as well as probing it, and can report a seek.
- `README.md`, `ARCHITECTURE.md` and `PRIVACY.md` describe pinning, its `Pin Mode` submenus, the two video lists and the media engine's role.

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
- **The pages both document engines draw are one cache, on disk.** `document_cache_mb` replaces `office_cache_mb` and `libre_cache_mb`, and a `config.ini` holding either older key is read once for the larger of the two and written without them. A page outlives the run and the engine that drew it, and is given up by when it was last *read*, not when it was converted.
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

- A name in the text list *and* in a drawing or document list is previewed as that kind again, not as text.
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

- Font previews for `ttf`, `otf`, `ttc`, `woff` and `woff2`, drawn by the WebView2 engine the SVG previews already use: the name the font calls itself, the pangram *The quick brown fox jumps over the lazy dog.*, and a line each for Japanese, Chinese, Korean, Cyrillic and Greek that the font's own character map covers. A line is drawn only where every character of it is in the font, so nothing is ever shown in a system font and passed off as the font.
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
- **Every process the app starts is now put in a Windows job object**, so a crash, a kill from Task Manager or a logoff ends them with the app, and the next launch ends whatever the job could not take. Nothing is acted on by id alone: a record is used only when the process still carries the image *and* the start time it was recorded with.
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

