# Changelog

## [Unreleased]

### Added

- **`System → Silent Auto-Update` installs a new release as the app starts, and asks nothing** — the switch sits under `Check for Updates`, off where the app starts, and is greyed while the checks are off. On, the one check a run makes as it starts installs what it finds: the same fetch, the same silent installer and the same `/S /R` hand-over the **Auto** answer of the dialog runs, and the app ends itself so the installer can replace it and start it again. Nothing is installed from a check the tray menu opened, so a preview or a pin on screen is never closed mid-session, and turning the switch on while a release is already on offer installs that one at once, on a thread of its own so the menu does not stall for the download. A silent install that fails — offline, a release carrying no installer, a program the machine will not start — says nothing at all and leaves `Update Available (vX.Y.Z)` in the menu for a hand click. It is `silent_auto_update` in `config.ini`.

## [0.5.0] - 2026-11-09

### Added

- **A pinned sound's card grows a playlist band under itself** — the card's right-click menu gained a **Playlist** check row at its top, and what it opens is the rows of the pin's own walk in the window the card is already in: the file's name without its extension, what the file says about itself (`FLAC · 44.1 kHz · stereo · 320 kbps · 4.2 MB`, read once per row rather than per paint, and read on a thread of its own so a folder's worth of songs opens at once and fills in as each file is read) and how long it plays. A press on a row picks it and a double press shows that file in the card, a name the row has no room for marquees across it while the hand is on it, and a notch of the wheel steps three rows — with a scrollbar in the band's right margin whenever the rows do not all fit. The band's bottom edge is a drag handle that sets its height, the card above it never moves for it, and both its state and its height are remembered for the rest of the run, so the next pin comes up as the last one was left.

- **The band answers the arrows and `Return`, and both its edges and its scrollbar are a hand's** — `↑`/`↓` step the highlight down the rows, starting on the row the pin is showing and holding at either end rather than wrapping, and `Return` shows the row that highlight is on — nothing at all where no row has been picked, and the folder's own walk while the band is down; a band left up walks the folder's sounds whatever `Audio` is set to. Its **top** edge sets the height as the bottom one does, the window's top carried up by the rows the band gained, and every drag is measured on the screen rather than against the window's own rows, which chased itself a pixel at a time. A press on the scrollbar jumps to the row it means where it landed off the thumb and carries the thumb from the part of it the hand took hold of, and the band a hand grabs runs from the last column the length is drawn in to the window's own edge rather than the four pixels of painted bar. A row's name keeps travelling while the pointer is on it whether or not the hand moves inside the window.

- **`Scaling → 3D` gives a model a share of the display of its own** — the tray's `Scaling` submenu gained a `3D` row beside `Design`, listing the same display shares the documents above it list (`Fit to Screen`, `75%`, `50%`, `25%`, `10%`) and starting at `Fit to Screen` like them, written as `model_scale` in `config.ini`. Until now a model was placed by the `Image` share, which asked a photograph's question of a scene: a model has no size of its own — the picture an installed f3d draws of one is drawn at whatever box it is asked for — so what the share is of is the room the display has, the question a page, a drawing and a design document answer. A share of the room is quality the picture gets for free rather than an enlargement, and a hover and a pin both lay a model out at it: the share applies to the next hover and the next take-up, not to a preview already up. A hand-edited value is read the way the documents' are, and a file that names no `model_scale` starts at the whole of the room — the size a model was drawn at before the setting existed.

- **`Ctrl`+`F` over a pinned page seeds the find box with what is selected in the page** — the chord is the one an editor answers the same way: a word double pressed in the page, or a run dragged over, is what the reader is asking to find, so the query comes up holding it rather than the query the bar was last left with, its caret after it and its places already searched and highlighted. The seed is taken as written — two spaces selected are two spaces searched for — and cut at the first line break, because the query is one line. A page with nothing selected is unchanged: the chord brings back the query the bar was left holding, all of it selected.

- **A second `Ctrl`+`F` over a bar that is already up takes the word picked out of the page since** — the reader who double presses another word in the page, or drags over a run of it, and reaches for the chord again is asking for that one, so the query it seeded is replaced and the document searched again rather than the box merely being refocused. A query the reader is still in is untouched: the chord pressed with the field focused selects the whole of it for what is typed next, so the word a page was left selected at is not brought back over what has since been typed.

- **The two switches of a page's find bar say what they are while the pointer rests on them** — `Aa` and `W` are glyphs rather than words, so `Match Case` and `Match Whole Word` are drawn as a small card below the card, centred under the switch the hand is on, in the same panel and shadow a caption button's own name is floated in.

- **`Detach Window` hands a pinned window to a process of its own, so a preview can be pinned and then left completely alone** — the row stands above **To Bubble** in the pin's right-click menus, and the same ask is a caption button between **Always On Top** and **Minimize** (and a fourth window button in a sound's card, between the band and the minimize). What comes up is the same file, at the same place, at the same size, as a window a click, a hover, a key in Explorer or `Tray → Pin Mode → Update Preview` can never change, because it is a second run of this app in which no Explorer watch is running at all: everything a pin is — the caption and its buttons, the transport bar, **Previous**/**Next**, the bubble, the taskbar button, `Escape`, the menu — works there too, and only the file it was handed is on screen. What is handed over is the file, the media box and whether the window is maximized; the pin it came from comes down when the new window is *drawn* rather than when the hand-off is made, so the file is never in neither window, and a child that dies before its window appears leaves the pin standing.

- **Any number of detached windows, one tray icon and one taskbar button** — every detached window is a process of the same executable, which is the identity the taskbar groups buttons by, so they share this app's single button and its hover lists a thumbnail of each window beside the pin's own. They live under the tray icon: closing the app closes all of them (the job every engine of this app's is in also holds every hand-off, and each window waits on the process that began it as well), while closing one leaves the rest alone. Their engines are their own processes, so `Engine → TTL`/`Idle` never reaches them, and a window's own player, browser and render tier end with it.

- **A detached window reads the configuration file the way a pin reads the tray** — what is written down reaches it live, including the two rows the tray *says* rather than only writes (`Text → Font Size` and `Pin Mode → Always On Top`), which a window of another process would otherwise never hear. What a window only reads when it opens a file — the engine a film is played by, the scale its box was measured at — arrives at its next start, which is the behaviour a pin has always had, and a detached window is a separate window rather than a second kind of pin.

### Changed

- **A page's find bar is one row, the widget an editor's find box is** — the query, `Aa`, `W`, the count, the two steps and the close now sit side by side on a single row rather than the query taking a row of its own above the controls: the card is shorter, so it covers less of the page it is searching, and the query is what the rest of the row is left as (a page narrower than the card gives up the field's room rather than a control's).

- **`Ctrl`+`F` selects the text it seeds the box with** — a page's selection is put in the query whole and all of it selected, with the caret at its end, so the next character typed is the new search rather than an append to the word that was picked out: the same box VS Code puts a seeded selection in.

- **`↑` and `↓` step the find box through the queries this run has searched** — the last fifty, most recent first, with a query that is already in it moved to the front rather than repeated, written down when a query is committed (by a step through the matches, by `Esc` emptying it or by the bar being closed). `↑` from a query being typed keeps it as the draft that `↓` comes back to, and either end of the list stops rather than wrapping. The matches are stepped by `Enter`/`Shift`+`Enter` and `F3`/`Shift`+`F3` now — `↑`/`↓` are the field's own keys, an editor's, and `PgUp`/`PgDn` are nothing under a bar at all.

- **Previews are drawn in a viewer process of the run's own** — the tray app boots, owns the tray and the listing's hooks, and begins a viewer (`--viewer`) that owns the window: the hover, the pin, the media, the film host, the player, the document engine and the browser are its, and every crossing between the two is one ring of messages plus a block of published facts the viewer claims for its pid (and where no viewer answers, the tray draws the previews itself as it always did). What that fixes is everything a copy of a fact on the wrong side of the boundary used to answer: a wheel notch over a pin is counted where the tray reads it, the pin key's mouse half is bound in the process that owns the mouse hook, `Trim Caches`, `Stop Engines` and `Hide HTML Page` reach the process that holds the caches and the engines, and a pin asked to come down is read from the block that holds the pin rather than from a flag of whichever window drew last. It is also what a detached window is made of — see **Detach Window** above.

- **A preview nobody asked for is not what a detached window shows while it opens** — the file is opened by the same road a hover is (a film's player begun, a document's page asked for and waited on, a model's viewer started), with the preview itself kept past the corner of every display until the take-up replaces it. A file that never settles ends the process that was waiting for it rather than leaving one with no window in it, and the pin that made the hand-off stays.

- **The caption carries one button more, and a caption that runs out of room still gives up its walk first** — the detach stands at the seam between the pin's own buttons and the window's own three: `Previous`, `Next`, **Open With**, **Open With…**, **Always On Top**, **Detach**, **Minimize**, **Maximize**, **Close**. Close is still the last button on the right and the last to go, and the window's three keep the places they have always had.

- **A detached window's own `Detach Window` says so in the ink a state is read in, and answers nothing** — a window of its own has nothing left to hand on, so the caption’s button and both menu rows stay where they were; the button wears the numbers’ ink, the one the pin beside it wears when the window is kept on top, stands in the wash that side of the caption stands in, and lights a step toward white under the hand. It is not hovered and answers nothing, both menu rows are the theme’s text given up toward the page, and a faded row keeps its panel up where an answering one puts it away. The sound card’s fourth button says the same in its own place.

### Fixed

- **A page's find bar keeps clear of the page's scrollbar** — the card was laid out against the frame's right edge, so a document long enough to scroll had its bar painted over by it at the top and its own column taken away while the card was up. On a page that scrolls, the card's margin is now measured from the scrollbar's left edge rather than from the frame's — the page's own right margin plus a bar's width plus a gap the card stands off the bar by, all in the display's pixels — so the card is clear of the bar at every width and every scale a pin is resized to, and a page short enough not to scroll places it exactly where it was.

- **A page's find count is drawn whole, however many places the query was found** — `1210/19000` was cut after its first digits, because the count was given a fixed column measured for a count of four figures; the column is the room its own text needs now, taken out of the query's field rather than out of the count, so a page of nineteen thousand matches reads as `1210/19000` with every digit on it.

- **A `Ctrl`+`F` over a pinned page is answered on the press rather than depending on which reader of the keyboard saw it first** — the chord was read from the key's "pressed since the last read" bit, and that bit is spent by whoever reads the key first: the Explorer hook's own type-ahead pass reads every letter's, so a chord whose `Ctrl` was let go before the pin's next tick could be spent there and lost. The key being *down* where the last pass found it up is asked as well, the key is read only under `Ctrl` so no reader can spend a press it has no chord for, and a batch that searched always paints the find bar — a bar opened over a query its page does not hold used to find nothing, draw nothing, and stay invisible until something else repainted the window.

- **A caption's two state buttons wear the ink a state is read in** — **Always On Top** while the window is kept on top and **Detach** while it is a window of its own are drawn in the numbers’ ink, the one the same two buttons wear on a sound’s card, standing in the wash that side of the caption stands in and lighting a step toward white under the hand. Every other button of the strip is drawn as it was. The menus’ rows are unchanged: a detached window’s `Detach Window` row is still faded.

- **A detached sound continues where the pin had got to rather than where a hover of that file would start it** — a hand-off opened the file the way a hover is opened, which for a sound is the `Volume → Audio Seek` rule, so a card playing at 1:40 came up in its own window at whatever second a hover of that file was last left at (2:30, say): a seek that jumped the moment the window appeared. The tick that draws the card's clock now writes down the second it drew, and the hand-off carries it, so the new window is the window that was standing with the one thing about it that could have moved put back. A second drawn for another file is not this window's, and a card with no clock behind it hands over none.

- **The detach button in a sound's card lights where it is drawn, not only near the command bar** — the four window buttons stand in a band at the top of the card, and three of them were answered anywhere in it while the detach was asked for only inside the bar's own row below them, so a hand on the button was a hand on the artwork. It is answered with the three beside it now.

- **A sound's card states its facts in full** — the line breaks on a fact boundary, and its box is narrowed to the last fact it holds whole, so a narrow card scrolls the rest rather than clipping one of them; it travels only while a pointer rests on it and comes home when the pointer leaves; the format is drawn in the numbers' ink, the sample rate and the bit rate in the ordinary ink the name is set in, and the channels and the room the file takes in the muted gray; a bit rate the probe states none of is filled in from the file's own size over its length — or from `ffprobe`'s container average where not even that is known — rather than left out of the line; and the card is headed with the file's name rather than with the extension the format fact carries. A scrolled line ends on a character rather than on a parting `·`.

- **A sound's track is answered from the disk before anything is opened, and a read that found nothing is asked again** — a probe's answer was held only for the run that made it, and a failed read was held as an answer for the rest of the run: opening a folder, or a pinned sound's band, re-probed every file it touched, while one read that came back empty left a file this machine can play with no card and no player. The memo holds a `Track` and nothing else, and the track is written to the general disk cache, one file per answer, read back before Media Foundation is opened or `ffprobe` started.

- **A sound Windows cannot decode is read by `ffprobe` even where `ffplay` is not installed** — the probe was gated on the player, so a build shipping `ffprobe` alone gave a FLAC no track at all: no card, no player, and every row of a band left with its extension and its size. `ffprobe_available` is a question of its own now.

- **A pin goes behind the window you work in when you click into it** — a press inside the window that is already active moves no foreground at all, so a lowering wired to "Explorer came to the front" waited for a transition the hand never makes, and the pin stayed over the listing it follows. What asks now is the press itself, read in the process that draws the window and acted on where the point is inside none of this app's windows, behind the *frame* of the window it landed in rather than the child a click in a listing lands on — a press on the desktop gives the band back, and a pin raised over a listing is not lowered by a front it never took.

## [0.4.11] - 2026-11-08

### Added

- **The `Images` group of the `Components` menu reports the camera raws and WebP too** — the two rows the README calls optional are rows about the machine rather than about the app: the camera raws are asked of the Raw Image Extension by the file extensions it claims (`.nef`, `.cr3`, `.dng` and every other name from `3fr` to `x3f` are one decoder to it, so one row stands for them all), and `WebP` asks the WebP Image Extension instead of answering for the libwebp in the binary. A missing row is pickable and opens the Store page its extension is installed from, beside the HEIF, AVIF and JPEG XL rows that already did. Neither changes what a preview does: a still WebP is still decoded in-app and a raw is still developed by ImageMagick where the codec is absent.

- **`Components → Install` installs a missing engine with `winget` in a console window, and the rows say when it is done** — the `Components` submenu (renamed from `Codecs`) gained an `Install` group listing any of the seven engines this machine is missing whose package id is known (mpv, FFmpeg, LibreOffice, ImageMagick, PeaZip, Calibre, f3d): picking one asks first, then launches `winget install` in its own console window unelevated, so the menu stays responsive and progress is visible. Nothing is offered where no `winget` was found by looking rather than by starting it, and a machine that has every engine sees no such group at all. One install at a time: while one is in flight its row reads `… — installing…` and nothing in the group can be picked. A failure winget blames on its sources is answered with `winget source update` and one more attempt, never with a source removed and never with a prompt-driven installer; a failure that stays one leaves the console window open showing why. After a successful install the `PATH` is re-read from both environment keys before the re-probe — winget only broadcasts the change, and a running process keeps the `PATH` it started with — so the engine that was just installed is a preview this run can use rather than one the next start finds. The installer's own engines page shares this same install script.

- **`Engine → F3D TTL` decides how long a hovered model's viewer is kept, and a repeat hover of that model is a re-show rather than a launch** — the fourth TTL beside `MPV TTL` and the two office engines: `Indefinitely` keeps the viewer for the run, a number of seconds keeps it exactly that long, `Until Away (Default)` keeps it while Explorer can be reached and lets it go once the away timer has run, and `Never (0 seconds)` ends it with the hover that began it (`f3d_idle` in `config.ini`). What is kept is the process, the host window and the Shell's registration drawn into it, so the next hover of the same file is one instant picture with no launch and no spinner — the 200 to 400 ms a cold viewer costs, and the host made and registered again, are the whole of what the row buys. It is only ever that one file: f3d draws the file it was begun with and knows no way to open another, so a hover of a different model lets the kept viewer go and begins its own. The row is greyed on a machine with no f3d, and a kept viewer is one resident process and one window off the screen, with nothing of it drawn anywhere.

- **`Preview Types → 3D` and the `[f3d]` list say which names a model is** — the kind is on by default and closes the graphics group of the tray, and `[f3d]` carries the 58 names f3d 4.0.1 reads that no other kind already claims, from `stl`, `3mf`, `gltf` and `ply` to `step`, `las`, `obj` and `vtk`; `f3d_path` names a copy installed anywhere other than beside the executable, on the `PATH` or under Program Files. A model's box is placed by **Image** scaling and asks **Image Background** for the colour the scene clears to (`0.565` for the checkerboard), and those names route to f3d rather than to the picture reader, which is what makes a `.stl` a model instead of a failed picture. A machine without f3d shows no preview for one rather than a broken one.

### Changed

- **A DXF drawing routes to f3d or LibreOffice by the geometry it holds rather than by its name** — `.dxf` sits in both `[f3d]` and `[libre]`: the front of the file (up to 256 KB and 20,000 group-code pairs, stepping past any UTF-8 BOM) is parsed for 3D entities (`3DFACE`, or a `POLYLINE` whose flags mark a 3D polyline, mesh or polyface) to route 3D models to f3d, while 2D drafting sheets or drawings past the bound route to LibreOffice. Where f3d is not installed or the 3D preview gate is off, the drawing is left to LibreOffice alone.

- **The build is aimed at the SSE4.2 level of the instruction set, so a picture preview uses the vector units the machine has without leaving any machine out** — every profile in this repository now compiles with `-C target-cpu=x86-64-v2` (`.cargo/config.toml`) instead of the 2003 x86-64 baseline: SSE4.2, SSSE3, POPCNT and CMPXCHG16B in the decoders (`image`, `jxl-oxide`), the Markdown parser, the highlighter's own scan and the scaling this app does itself. What does not change is everything compiled outside rustc — the standard library is not rebuilt, and the C engines (libwebp, the archive and document converters) keep asking the CPU what it has at run time. v2 rather than the AVX2 a Haswell or an Excavator has, because that is the level with no machine left out: every processor Windows 11 supports is above it, and the AVX2 level would exclude an ARM64 Windows device, whose x64 emulation reaches AVX2 only from 24H2. The installer and the `github` profile that produces it are otherwise as they were: fat LTO, with the merged module now split into sixteen codegen units after its passes, which shortens the build and costs nothing at run time.

- **A model's picture is this app's host window drawing the engine's, not the engine's own window standing at the box** — 0.4.10 put f3d's chrome-less window at the hover box, which is a window the Shell cannot be asked to draw into a pin; the picture is a DWM thumbnail of that window in the hover's host now, so what a pin takes up is the picture already on the screen rather than a second f3d, and f3d's own window is never on a screen at all. Under Pin Mode off no host is taken, and the engine's chrome-less window stands at the box as before.

- **A pinned model's band is the scene's, and every button in it is the engine's** — a left drag orbits the model, a middle drag pans it, a right press opens f3d's own menu and the wheel zooms it, all posted to the engine and none of them activating it. The band answers `HTCLIENT`, so a hand on the picture turns the scene instead of carrying the window it is drawn in, and the caption strip is the only chrome this app answers, its own menu opening there. A middle press over a model is the engine's pan and never the pin's put-up, so a model is pinned with `Space` or the pin key.

- **A model's turn is posted one move at a time rather than one per pointer sample** — the engine draws a whole frame for every message it is handed, measured at about 31 ms at 480x360 and 70 ms at 900x700, so a hand reporting a thousand positions a second queued a scene still turning long after it had let go. The moves are coalesced now: the engine is handed the position the hand has reached when it is ready for one, and the turn stops with the hand.

- **Beginning a model's viewer no longer holds the thread that draws until the engine has made its window** — the launch waited for that window on a 5 ms poll, which measured as about 1.2 s of a frozen preview with nothing on the screen and the cursor spinner only coming up afterwards. The wait is off the drawing thread now, so the spinner is up on the first tick after the hover commits and the window is adopted whenever the engine makes it; a window that arrives after the hover or pin that began it has let the player go is ended unseen rather than adopted.

### Fixed

- **The engines page of the installer opens its `Download` links, and opens them in your own browser** — the seven rows hung their handler off `BN_CLICKED`, which a `SysLink` never sends, so the links were inert; and `ExecShell` hands a URL to the token the installer runs elevated with, which lands it in an administrator's profile or nowhere. The click is a notification now, the address is handed to `explorer.exe` so it opens as the interactive user at medium integrity, and FFmpeg is looked for beside the app before the `PATH`, which is where the app itself looks. The seven rows are one macro.

- **A file whose name is in `[f3d]` is a model only where its own head says it is, so a `.bin` that is not a point cloud no longer starts the viewer for a second only to be turned down with a black window** — the name was the whole of the claim, and `bin` is a name half the machine's own formats share, so a file of that name which was nothing of the sort was handed to f3d anyway, which spent the launch and a window finding out. Every one of the list's names is now held to the front of the file before the kind is claimed: a `.las` to its magic, a `.vtk` to the line that names the format, a `.step` to the header every STEP document opens with, a `.stl` to either its `solid` line or a triangle count that fits its length, a `.3mf`, a `.fcstd` and a `.usdz` to the zip they arrive in, a `.bin` to the two words of the TerraScan header PDAL reads, and four names — a QFIT cloud, an SBET trajectory, a gaussian splat cloud and an OpenCASCADE cache — whose formats declare nothing in their front at all, to nothing at all: their name is the whole of the claim, so the viewer is started and answers. A file that fails the question is not a model: it previews as nothing and starts nothing, which is the answer a name outside every list already had. A name you add to the `[f3d]` list yourself keeps the old answer, since nothing here knows the format it stands for, and so does a head that cannot be read — a file on a cloud volume is asked no question a signature could answer.

- **Hovering a model no longer flashes a black window, at the preview box or in the corner of the screen** — the engine shows the window it made itself, measured on f3d 4.0.1 at 42 to 182 ms after it is begun, and it makes it at the top-left of the desktop as a 300x300 decorated window some 30 ms before the place it was asked for lands: that was a titled black rectangle over the box, and a second one in the corner of the screen. A model's viewer is begun hidden now — a `STARTUPINFO` asking for `SW_HIDE`, which the engine obeys — and past the right edge of every display, so no window of it is ever on a screen before this app has one. The box is populated only by the step that already holds the window: the frame comes off, the window is hidden, placed at the box, made the transparent layered source a thumbnail is cut from, and shown.

- **The picture opens on the model instead of on the engine's blank** — a scene's first render is not its last: the surface is cleared to the backdrop colour before the file is in it, which is the black, then grey, then model a cold hover was seen as. The host is shown only once the viewer has said the file is loaded — its own `100%` loading line, written to standard output and flushed within 100 to 200 ms — plus three frames of grace for the renderer to put a scene in the picture. A viewer that reports nothing at all is sampled for lit pixels after 1.5 s instead, and one whose report never reaches its marker is capped at 8 s inside the 10 s a start is given, so a format this build does not report on still shows, even if it shows the old way. What is left of a cold hover is the spinner, and then the model.

- **A hover that ends leaves nothing of a model on the screen** — a kept viewer's host is taken off the screen with the hover that showed it, and with it the one picture that used to stand over the listing after the pointer had left; the viewer itself stays behind it, off the screen, until its TTL or the pin that took it up ends it.

## [0.4.10] - 2026-11-06

### Added

- **Camera raws and system picture formats decode through Windows natively** — `[image]` handles camera raws (`.raw`, `.nef`, `.cr2`, `.cr3`, `.dng`, etc.) via the Windows Raw Image Extension and system bitmap/cursor formats (`dib`, `rle`, `cur`) with WIC, falling back to ImageMagick only when the native codec is missing.
- **The pinned sound card context menu configures both hover and pin seek modes** — right-clicking a pinned sound card offers separate `Hover Seek` and `Pin Mode Seek` flyout menus.

- **3D models, CAD exchange files, point clouds and scientific datasets preview by hovering or pinning them** — a new `3D` kind beside the pictures, drawn by f3d where one is installed and by nothing where it is not: f3d's own chrome-less window stands at the hover box (or, under Pin Mode, the hover's host mirrors it, which is the window a pin promotes), and the `[f3d]` list in `config.ini` names the extensions asked of it — everything f3d 4.0.1's own reader list carries that no other kind already claims, from `.stl` and `.obj` to `.step`, `.gltf`, `.las` and `.vtk`. `f3d_path` names a copy installed anywhere other than the `PATH` or under Program Files, the tray's **Preview Types → 3D** row switches the kind off and on, a pinned model is orbited by dragging in its band and zoomed by the wheel, and a machine without f3d shows no preview for a model rather than a broken one. Models are drawn with f3d's contact shading (ambient occlusion) on, so a mesh that carries no texture of its own reads as a solid rather than a flat grey cut-out. The picture itself is the engine's to answer: dragging turns the scene, the wheel zooms it, the middle button pans it and the right one opens f3d's own menu, while the pinned window's menu stays on its caption strip.

### Changed

- **The `Engine` submenu reads in the order its rows are used** — `Video Engine`, `Office Engine`, a separator, `Away Timer`, a separator, then the four TTLs: `Microsoft Office TTL`, `LibreOffice TTL`, `WebView2 TTL`, `MPV TTL`.

### Fixed

- **Every `… TTL` submenu lists its idle times longest first** — `Indefinitely` at the top, then `1 hour` down to `30 seconds`, with `Never (0 seconds)` second to bottom and the default now reading `Until Away (Default)` at the bottom.

## [0.4.9] - 2026-11-04

### Changed

- **A pinned caption collapses buttons one by one rather than dropping the strip** — as a window narrows, its title is shortened, then dropped, and buttons are removed from the left (`Previous` first, `Close` last) so the strip remains as a bare hairline bar rather than disappearing; the title reappears only when all buttons fit without shifting under the pointer.

### Fixed

- **A pinned animation keeps playing through a maximize, and through a drag past the size it was decoded at** — a box the frames in hand do not cover is the one decode a resize still owes, and a decode begins a GIF, an animated WebP, an APNG, a HEIF sequence or an animated JPEG XL at its opening frame: the resize meant to sharpen the picture restarted it instead, once in a session for a GIF, whose frames then cover every later box, and on every maximize for one whose frames never cover the display. That decode is now asked for from the place the animation is at, and each frame carries where it is in its file, so the frame that lands is the one carrying that place, with the wait it had already served — whether the pin has played on since the box changed, whether its player has released the frames behind it, and whether the file has looped. A file decoded from a place is answered its tail and then its head, the file in play order, and the pass stops where it began rather than decoding the file on under it, so a pin can no longer loop inside the part of the file a resize happened to catch: a media that has played to the end of what it holds wraps onto the frame carrying the file's first frame. And the answer is not installed before it is the frame the picture is on: a decode lands a tick or more after the box changed, so what it begins with is behind the place the pin has played to, and installing it then was the picture put back by the decode's own duration — the half-second loop a maximize showed and the stutter on every step of a pulled edge. The answer now waits until its own stream reaches that place, the frames drawn for the new box being stretched until it does.

- **A resize no longer puts a pinned animation back at its first frame, and no longer softens one it shrank** — a pinned picture that plays is decoded at the box it is drawn in, and every box the hand left it at was laid out again. A decode begins a GIF, an animated WebP, APNG, a HEIF sequence or an animated JPEG XL at its opening frame, so a drag of an edge, a maximize or a double-press reset put the animation back at the beginning; and the box between was decoded for as well, so a window dragged small and opened again played back out of the thumbnail it had just been decoded into. Two answers now, the pair the other two gestures already give. A box the frames in hand already cover is no decode at all, which is the answer a still picture has always been given: those frames are every pixel the smaller box would have been drawn to, and every later enlargement is stretched from them, so a pin shrunk and pulled back out is as sharp as it was and keeps playing where it was. And a box they do not cover is decoded for — the pin laid out at the size the hand asked for — with the place the animation had got to carried across: the frame the pin was on and the wait that frame had already served, the new media being the same file's frames in the same order (`keep_animation_place`), so the new size sharpens the picture instead of restarting it, and the frame the pin was showing is not held for one more of its own delays.

- **A picture carried inside its pinned window no longer freezes the animation it is showing** — a picture zoomed past its window is carried about by a press and a drag, and that pan painted the whole band on every position the pointer reported. It is the one gesture whose paint is a composed band rather than a copy of rows — 2.3 ms of a 1920x1080 band off a 3840x2160 frame, before the surface is cleared and the layered window is handed it — and a mouse reports up to a thousand positions a second, so the pan asked for several bands per frame the display could show, on the same thread the preview loop's tick runs on. The tick that advances a playing picture therefore did not come while the hand moved, which is seen as an animation that stops, then jumps when the hand stops, with the gap between one frame and the next long enough to read. The move now only writes where the picture has got to and the loop draws it on its own tick, its wait paced by the compositor while a pan is in flight — the same flush that paces a wheel's travel — so the picture is drawn once per frame the display shows, the pointer's positions between them coalesced into the one the hand has reached, and the pan's paint is the animation's own. **A hand pulling an edge is paced the same way**, though its step stays where it was: that step is one `UpdateLayeredWindow` that is also the box's own commit, so it cannot be handed to the loop without splitting the box from the picture in it. It waits on the compositor after it instead, so a step is committed once per frame the display shows, and the tick that follows it paints the pin only if nothing has since the pointer's message did — a resize that used to be two composed bands a frame is one. And the message pump dispatches one pointer position a tick while a resize is in flight, leaving the rest for the tick after it: a step per position is a pump that never empties, and the tick it starves is the one that advances a playing picture — which is why a resize froze the animation for as long as the hand kept moving and it played again the instant the button was let go. The positions left behind are coalesced by the system to where the hand has got to, and the tick reads the pointer live anyway.

- **A wheel zoom no longer puts a pinned animation back at its first frame** — a picture zoomed past its window is decoded again at the size it is drawn at, and a decode of a GIF, an animated WebP or APNG, a HEIF sequence or an animated JPEG XL begins the file at its opening frame. So the decode meant to sharpen the picture restarted the animation on every notch the wheel settled at, and a hand turning the wheel in and out watched the same opening seconds over and over. A frame that *plays* is now the one kind a zoom does not re-read: the frame the pin holds is stretched into the box the picture is drawn at — the answer a picture that fills its window is given, and the one every still picture is drawn from while its decode is in flight — so the animation plays on through the zoom, softened above the size its frames were decoded at. A box change and a take-up still decode, being a pin put at a size the hand asked for and one built for a file it has not stood on.

- **A pinned film is not trimmed of eight rows it does not have** — the crop detector was asked to floor what it measures to a 16-px grid, and 1080 is not a multiple of it, so a whole 1920x1080 picture came back as `1920x1072` with four rows off the top and four off the bottom, inside the slack both guards allow. A pin maximized on a 1080-row screen was drawn 1920x1072 and left a 4-px band of desktop above and below it. The grid is 8, which every common frame height is a multiple of, so a whole frame measures as itself and is refused as the full-frame crop it already was; and a crop that trims under 16 px off *both* axes is refused too, that being codec padding (H.264 pads 1080 to 1088) or rounding rather than a letterbox. A 1.85:1 picture, 43 rows off, is still cropped to; a 2.40:1 one stays outside the 10% cap and is not cropped, as before.

- **A pinned film keeps its crop when carried across displays of different scales** — mpv rescales on `WM_DPICHANGED`, giving a pinned film carried across a DPI boundary a client twice the size it was placed at; the crop is now fitted into the source's live client geometry rather than the initial placement client, so the film is no longer shown zoomed into its top-left quarter.

- **A pinned window survives display and DPI changes instead of closing** — `WM_DPICHANGED` and `WM_DISPLAYCHANGE` no longer tear down a pinned preview or its media. The pin remains alive and is re-boxed for the display it landed on by the next tick: a maximized pin re-maximizes to the new monitor's work area, and captions and transport bars are re-measured at the new scale.

- **A closed or collapsed pin takes its caption popup off the desktop immediately** — the popup caption used over document and specimen previews is a top-level desktop window owned by the pin and is now explicitly hidden when the pin is hidden, collapsed, or closed, preventing an unclickable titlebar from lingering over the desktop.

## [0.4.8] - 2026-11-03

### Added

- **`Pin Mode → Navigation → Image`** — the walk's anchors are four now, and the picture is the first of them: above Video, with the same **All Files** / **Same Category** pair, starting at `Same Category` (`pin_nav_image` in `config.ini`). A file holding only the old `pin_nav_file_types` seeds it with the other three. Pictures used to be Others files, so a walk anchored on one steps through the folder's pictures rather than end to end; a document still walks end to end.

- **`Preview Settings → Text → Preview All Extensionless`** — off by default: on, a file with no extension at all (`README`, `notes`, a source file a repository left bare) is previewed as text, where before only the names list (`LICENSE`, `Makefile`, `.gitignore` and the rest) counted. `text_show_extensionless` in `config.ini`.

- **`Preview Settings → Text → Don't Preview Secret Files`** — on by default: `.env` and its variants, `id_rsa`/`id_ed25519`/`credentials` and their suffixed copies, key, certificate, keystore, VPN and password-database extensions, and anything under `~/.ssh`, `~/.aws`, `~/.gcp`, `~/.kube`, `~/.docker` or a coding agent's credential path are not previewed at all, rather than shown as the text they are. `*.pub` and the `.env.example`/`.sample`/`.template`/`.tpl`/`.defaults` templates are still read. `text_dont_preview_secrets` in `config.ini`.

### Changed

- **A pinned page's caret keys are the page's only once a press inside it has given it the focus** — a page came up answering the arrows, `PgUp`/`PgDn`, `Home`/`End` and their `Ctrl` forms with its caret, so the walk the caption offers could not be stepped from the keyboard. It is drawn unfocused now, with no caret and every one of those keys stepping the folder as it does over a picture; a press in its body focuses it — the same press that puts the caret down or begins a selection — and a press on the caption (its walk, its buttons and the empty drag area) gives the focus up again, clearing none of the caret, its column or the selection, so a page pressed into again comes back where it was left. `Ctrl+A`/`C`, `F3` and `Ctrl+F` read the page rather than move a caret in it and are the page's either way.

### Fixed

- **An image pin moved from the bubble into the taskbar keeps its picture in the button's hover.** A pin that stands in a window of its own wears its own taskbar button, and the row that moves one from the bubble into the taskbar minimized a window that had been hidden since the collapse — a window with no composition for the Shell to freeze, and no picture taken of it — so the hover fell back to the icon. The picture is now taken at the collapse, while that window is still on the screen, and the window is made supplying the still its button is drawn of and answers the Shell's ask for it; a pin restored, ended, or up again gives the still back and forgets the picture, so the direct minimize keeps the Shell's own freeze it has always had.

- **A press on a pin's own menu is the pin's gesture** — the menu is a layered popup that takes no input of its own, so its presses arrive at the card's window and were read by the Explorer hook as presses on the listing underneath: on a menu clamped under the pointer at a screen edge, a press on a row lowered the pin behind that listing and the menu came away. Such a press is neither a pick nor a no-pick now, it spends the rest of the hand, and the menu's own window counts as the pin's when the dismiss ask is read, so a press that lands on the menu puts away nothing.

- **The caption's chevrons step the folder over a pinned page of text** — a chevron asked for `Previous`/`Next`, which are the arrow keys' commands and became the page's caret wherever a page held one, so a press on the chrome moved the caret and never reached the walk. The chrome, the menu's rows, a bubble's rows and the taskbar's buttons now ask `PreviousFile`/`NextFile`, which are the walk whatever the pin holds.

- **The pin's own window and the taskbar stand-in wear this app's icon** — neither was given one, so a pinned picture's button (and the stand-in a bubble's button is made from) drew whatever Windows falls back on. Both sizes are set on both windows, the small one being what a taskbar button draws.

## [0.4.7] - 2026-11-03

### Changed

- **A pin's move roads name no band they drop** — `apply_pin_drag` and `centre_pinned_window` passed `HWND_TOPMOST` beside `SWP_NOZORDER`, which drops that argument: the call moved the window and made no raise, while reading in the source as one it did. The insert-after argument is now empty, and no move claims a raise it never made.

- **The pin road's trace names the window in front, and what stands over one** — `RHP_PIN_TRACE` describes a window rather than leaving it as `other`: `class:pid (image)` for any window Windows can still be asked about, `none` for a handle that names no window and `invalid` for one that names none any more, and this app's own by the words they already had (`pin`, `host`, `preview`, `bubble`, `stand-in`). Every line that reads the front carries it — `set_foreground … -> foreground=… pin above=…`, the band flip with the road that asked for it (`band taken and given back: a press`, `a preview update`, `a press on a running page`, `the pin appeared`, `the setting named the ordinary band`, `the setting was turned off`), the host's show, `pin above:` after a show, the keyboard hand-back at a collapse and at an end, the player's topmost show — and every message a pin's window is handed carries an `above=` beside its `foreground=` and `focus=`. A pin another program put down, or one left behind by a box change, is read from those lines rather than guessed at from the hole it left on the screen. Turning **Always On Top** off is traced on its own: `setting names <band>: pin … caret in … holds-band …`, `appwindow <0|1> of <window>: the user is in it, read at its next show`, `settle after the setting moved: caret in … stands-over … raisable …` and `settled: …`.

### Fixed

- **A pin is not put behind the listing by its own gesture** — a maximize, its restore, a reset size, the end of a resize and a double press on a pin's own picture are all made inside the pin, and the press that ends one used to be read against whatever the window had moved over: a film's host answers `HTTRANSPARENT`, so the hit test under the pin's picture named the listing the pin was fitted over, the point held the gap between the rows, and a press that named no file was answered by lowering the pin behind that listing until a click brought it back. A press whose point is inside the pin's own box is now the pin's own whatever the hit test names under it, and a press read on the very tick this app's own window moved is neither a pick nor a no-pick, since the hand did not move and what reads under the point is whatever the move uncovered — a gesture that changed the box leaves the point *outside* the new one, which is why the box alone cannot refuse that press. Neither reading is answered as a pick of a row either, so a pin carried over a file by its own gesture no longer picks it.

- **A press on a sound card's own controls is not a pick of the file under it** — the hook answered a pin's caption buttons as the pin's chrome and nothing else, and a sound has no caption, so every press on its card — the always-on-top mark included — arrived with the file the card stood over. `pin_button_covers` now asks the card's controls of the same geometry the card's own press is read against.

- **`Pin Mode → Always On Top` turned off leaves the pin in front of the ordinary band, with no hole under the hand** — three steps of that click moved the caret: `HWND_NOTOPMOST` puts a demoted window just under the foreground one, so the demotion is now the band flip and leaves the pin in front of the listing it was picked in; the Shell's button bit costs a visible window a hide and a re-show, and the pin under the hand is the window the user is in, so its bit is written and read at the pin's next show while every other window still pays the show; and a hidden window hands the caret to the window next in the z-order, so the foreground is read before that put-down and the pin's place is settled after it — raised only where the window in front is its own listing, the pin itself or the desktop. A pin behind a program the user chose is left where their own click put it.

## [0.4.6] - 2026-11-03

### Added

- **`Pin Mode → Always On Top`** — whether a pinned window stays above everything else, on by default (`pin_always_on_top`); the same switch is the mark left of the window buttons in a pin's caption. Off, a pin is an ordinary window: it keeps the place the user's own clicks gave it, other programs may cover it, and it wears this app's own taskbar button while it is up. The wheel over it is the pin's only while the pin is the window the user is in, and where mpv draws the film the player is told the same band down its own pipe, so a film under a pin that is not always on top no longer keeps itself above the program the pin stands over.
- **A pin's taskbar button is the pin.** The stand-in is borderless — the Shell's hover is the picture and nothing of Windows' own title bar — wearing this app's own icon and the pinned file's own name (the app's name only while there is no file), and its hover holds the pin's picture: the live one while the pin is on the screen, mirrored with a DWM thumbnail from the window the picture is really in — the player's or the engine's where there is one, this app's own otherwise — and the frame the collapse took once the pin is put away, a sound being pictured by its card. A film's or a sound's hover carries **Previous**, **Play/Pause** and **Next**, which are the pin's own commands and read the same Navigation settings, shuffle and end of file the caption's do.
- **A pinned window's own right-click menu** — the panel a bubble opens, over a window that is still standing: the hold, **Next File**, **Previous File**, **Pause on Minimize** where the kind behind it plays something, the two picks of where a minimize sends the pin (**Minimize to Bubble** / **Minimize to Taskbar**, the one in force carrying the disc), then **To Bubble** and **To Taskbar** — which put the window away at once — and **Close**, the groups parted by rules. A page of text carries **Select All**, **Copy**, **Find** and **Find Next** above them, read against what is on screen, so a walk off the text leaves them out and a walk back brings them; a sound's card keeps its own menu with the flyouts its **End of File** and **Seek** hang from. Where the menu opens is the kind's own answer: a card's body, a picture's body and title bar, and a film's title bar alone, since a band another process's window fills — FFmpeg's player, a rendered page — takes the clicks in the picture for itself. A film is left in its own band while the panel is up rather than raised over it, and any key or a press outside puts the panel away.
- **A pinned page of text is read, not only looked at.** `Up`, `Down`, `Page Up` and `Page Down` work the page — a screenful at a time for the two page keys, and a caret's line for the two arrows since the page became a read-only editor (see below) — `Home` and `End` take the caret to the ends of the line it is on, and the folder still walks for a pin holding anything else. `F3`, and a **Find Next** row in the page's menu, take it to the next place what is selected appears: the search is exact and case-sensitive, starts just after the selection, wraps to end where it began, and leaves the page exactly where it was when there is nowhere to go. **Every pinned page keeps a column of its own line numbers down the left** — source, prose, a listing or a rendered document, always on while the page is pinned and never on a hover — as wide as the widest number the page can reach, so the column does not move as it is scrolled from line 9 to line 10; a wrapped line carries its number on its first row and none on the rest, and a `Ctrl`+`A` and `Ctrl`+`C` of the page still copy the text without them. `Ctrl` and a wheel steps that page's text size through the sizes **Text → Font Size** offers and stops at both ends, and the size is the page's own: it rides across a resize and the next file of a walk, and closing the pin is what gives it up.
- **A pinned page has a find bar.** `Ctrl`+`F` — or the **Find** row above **Find Next** in the page's own right-click menu — opens a card over the page's top-right corner: the query with `current/total` beside it, `Aa` for case and `W` for whole words, and `↑` `↓` `✕` to step through and close. Every keystroke searches the whole document as it is drawn, every place the query appears is highlighted with the one the reader is on drawn apart, and the page is taken to it when it is off screen; `Enter`, `↓` and `F3` step on, `Shift`+`Enter` and `↑` back, and both wrap at the ends. A new query starts at the first match at or after the line the reader was on. While the bar is up it owns the keyboard — the page's own keys and the polled `Ctrl`+`A`/`Ctrl`+`C` are inert. **The query is a real field**: a press in it places the caret (a double press takes the word under the pointer, `Shift` extends the selection), `←`/`→` move a character and `Ctrl`+`←`/`→` a word with `Home`/`End` at the ends, `Shift` grows the selection, typing replaces it, and `Ctrl`+`A`/`C`/`X`/`V` select, copy, cut and paste — a paste becoming one line, capped — while `Ctrl`+`Z`, `Ctrl`+`Y` and `Ctrl`+`Shift`+`Z` undo and redo a run of typing at a time. The caret blinks at the system's pace and is solid from every edit, the selection is washed in the field, `Ctrl`+`F` over an open bar selects the query rather than appending to it, and `Esc` empties a query with text before it closes the bar (a second `Esc`, or the `✕`), leaving the page where it is and giving up the highlights while keeping the query for the next `Ctrl`+`F`; **Find Next** steps that query where the bar is up and the selection where it is not. A box change keeps the bar, and another file closes it.
- **A pinned page is a read-only editor.** `↑`/`↓` move a caret a line at a time and keep their column over the short lines in between, `←`/`→` cross the ends of lines, `Ctrl` and an arrow take a word by the page's own run rule, `Home`/`End` go to the line's two ends and `Ctrl`+`Home`/`End` to the document's, `PgUp`/`PgDn` are a screenful, `Shift` grows the selection and a bare move collapses it — and the page scrolls only as far as the caret needs, where the arrows used to scroll it a screenful at a time. The file is never touched: a typed character, `Backspace`, `Delete`, `Ctrl`+`X` and `Ctrl`+`V` all change nothing, while `Ctrl`+`A` and `Ctrl`+`C` still read the page. And the caption says so, with a **Read only** pill carrying a hand-drawn padlock, immediately left of its buttons.
- **A pinned page is selected with the hand.** A press begins a selection, `Shift` grows it from the anchor, and a second press takes the word under the pointer — letters, digits and the underscore are one run, a mark is its own kind, so `->` comes out whole. A third press inside the same time and box takes the whole document line the word is on, break and all, and the word a double press took stays whole for the drag after it, which grows a run at a time rather than a character. A drag carried above or below the box takes the page three lines a tick toward where it left and grows the selection to the line that comes back at that edge, so more than a screenful can be selected with the pointer. A selection, and the place a copy is cut from, are measured in the document rather than in the frame: a line too long for the box copies whole, **Select All** means the page rather than the screenful, and the selection survives a resize and a scroll.
- **A taken-up sound's `End of File` is answered without rebuilding its player.** A pin that comes up over a sound keeps the player the hover began, so a pass end the pin owes is read where the file comes round rather than by killing and beginning that player again.
- **A pinned sound card's menu offers another program** — **Open in Default App** and **Open With...** below its **Seek** row, handing the file over and ending the pin the way the card's own two buttons do.
- **`Pin Mode → Update Preview → Move To Top`** — on by default (`pin_update_move_to_top`): a file an update shows also puts the pin at the top of its own band, without taking the keyboard from the listing. Greyed while **Follow Selection** is off. And `toggle_key` (`f7` by default) throws `preview_enabled` while Explorer is in front, never swallowing the key; the tray's **Enable Preview** row names it.
- **A pinned SVG, font specimen or rendered HTML page is a pin like any other** — the three kinds the browser engine draws now stand in the pin's own window: it is the window Windows activates, lists and gives a button, its caption floats over the page, and the page follows a drag, a resize, a maximize, a restore and a walk between files. A page that runs — a page of HTML with **Render HTML** on — keeps its own window in front of the pin and its caption in bands above it, a caption drawn over it being one underneath it.
- **A pinned page's picture is drawn into the pin rather than nested in it** — the page's window is lowered, invisible at one part in 255 and click-through, and the band is the Shell's own thumbnail of its client: the picture rides the pin's client through a drag, the pin's taskbar hover holds the page rather than a hole, and a press on the band is the pin's own.
- **The installer, its uninstaller and the installed-apps entry carry the app icon** — the uninstaller kept NSIS's stock glyph, and the entry's icon named the exe an update replaces; the icon now ships beside it as `app.ico`.

### Changed

- **The engine a preview was built under is the engine it keeps.** `Engine → Video Engine`, its **Fallback** switch and `Engine → Office Engine` were written when a row was clicked and read again on every question the preview asked after that, so a resize, a maximize, a seek, a transport key or a step came back with the engine of now: a film playing in one player was ended and begun again in another mid-pass, and a page keyed by the engine that drew it was asked of one that had never seen the document. The choices are read once, where a preview is built — a hover, and a pinned window being shown a file it is not already showing — and a pin taken up over a hover needs no read of its own, the window adopting the player the hover began. A choice reaches the next preview; the probed geometry is given up where the choice is read, since a geometry measured under one engine is not the one the next player wants.
- **A pinned film plays under a hand on the seekbar.** A press on the track used to park and kill FFmpeg's player behind a cover, so the film stopped the moment the hand landed and stayed stopped for the whole of the drag; the press now takes the aim alone, and the kill and the relaunch happen at the release, where the second the hand chose is known — the cover raised over the frame the hand let go on and drawn at that second rather than flashing back to where the thumb was. An mpv film was hidden behind a frozen frame for the whole gesture and for the swap bound after it, over a seek that ends nothing: mpv is told a second down its pipe and goes on playing, so that park is refused for the route and for the volume knob, and a seek whose pipe cannot be reached still falls back to the relaunch that raises the cover.
- **`Engine → Video Engine` lists FFmpeg's player above the engine Windows ships** — **Best**, **MPV**, **FFmpeg**, **Native**, **Native (FFmpeg above 3.2MP)** — so the two players a film is handed to follow the machine's own answer and the engine that draws into this window's band comes after them, which is the order the engines are reached in.
- **Every kind's box is measured off the preview thread, and a hover that comes to nothing ends as the cross.** Text, design pages and pictures now ask for their box the way PDF, comic, engine, audio and video already did, so what a hover — or a pin — is given while the read runs is the wait for the box rather than a tick with nothing to draw: on a file syntect takes seconds to highlight, the spinner the `250 ms` setting asked for used to arrive only as the page was about to be placed. A wait that ends with no preview at all — the readers answering nothing, a load answering no media, the cap on a wait for an engine — no longer blinks the spinner out, which cannot be told from a wait still running, but ends as the cross a pin shows for the same failure, in the box the wait had, and goes three seconds later.
- **Names two kinds share are settled by the file's own bytes.** A `.ts` or `.mts` — the pair both video lists and the text list carry — was called a film by the name-only answer, so a TypeScript source waited on a shape probe that never answered and a pin held its swap on that same wait; a `.vhd` of VHDL source was a disk image the listing engine was started for, a `.dif` spreadsheet was the film the video list claims, and ImageMagick's own `.mpc` cache was an audio card the sound engine was asked for — waits that never answered, over source and over pictures. A kind the bytes name outright is that kind, a format no kind previews is drawn as nothing rather than as the film or the card the name list claims, and a pin takes up only the drawing the loader would have made — which is also what keeps an `.htm` with **Render HTML** off the text page the hover showed rather than a pin onto an engine that was never asked, and a metafile or a PostScript program under a name no vector list claims the drawing the replaying layer draws.
- **The tray's Text settings reach a page a pin has already put up.** **Theme** and **Markdown** lay a pinned page out again from the configuration, at the size the reader put it to; **Font Size** is the one setting that takes a page already up to another size, so the page being read changes under the pick, while a size hand-edited into the file, or one a `Ctrl` and a wheel gave that page, is still left alone.
- **Groups in the tray where exactly one answer is in force carry the disc rather than the tick**, the way the rows that already used a radio mark do; the switches keep their ticks. And the **Open With** mark in a pin's caption is drawn as the external-link mark it stands for, in whole pixels on the glyph square rather than as a blended staircase, with the always-on-top mark's own square filled the same way.
- **A pinned film's level is written to its player's Windows mixer session.** FFmpeg's player used to be ended and begun again for a level, so the film stopped while it was turned; the level now goes to that player's mixer session and the film plays on, where mpv and the media engine were already told theirs while they ran.
- **The wheel over a pinned film's volume control turns it.** A notch steps the level `5%`, each notch re-bases on the level in hand, and a level this app wrote is not read back off the knob under the pointer.
- **A `config.ini` volume edit reaches a pin that is already playing** — the file's watcher and the tray's own reload each publish the half the file moved, and the half this app last wrote is left alone.
- **Every engine TTL defaults to `away`** — `libreoffice_idle`, `office_engine_idle`, `webview_idle` and `video_engine_idle` keep their engine while Explorer is reachable and let it go after **Away Timer**, where a number of seconds is that long after the last file. The old `office_engine_persistent`/`libreoffice_persistent`/`webview_persistent` switches are read once and dropped.
- **The box a pin is given is the box the user last chose** — a maximize, its restore and a double-click's reset each write the pin's bound the way a pulled edge does, so the files picked after one are fitted into that box up as well as down rather than into the size before it; a maximize a caption double-click asks for is the pin's own toggle rather than a native box neither its restore nor its bound ever hears about. A page the browser draws is laid out at its own page box and fitted into the bound, so a pin walking between a page and a picture keeps the size the hand gave it.
- **A film's hovered mpv is kept warm between hovers as one resident**, parked off-screen at 1×1 and ended by **MPV TTL** (`video_engine_idle`), so a hover no longer pays a cold launch. A resident a pin took over leaves a fresh one owed.

### Fixed

- **A caret on an empty line stands inside the text, not over the numbers** — the bar was placed at the frame's own left edge when the line under it had nothing on it, because an empty line has no run to be measured against: the caret was painted across the column of line numbers a pinned page keeps, rows like `21`, `22`, `23`. The column a line's text begins in is now part of the painted line itself — past the numbers and past the indent the line carries — so a place on an empty line stands on that edge: `Home`/`End` and `←`/`→` to it, `↑`/`↓` onto it, a press on the line and a press in the gutter beside it all land at column 0, which is that edge. The same at 150% and at 200%, where the numbers and the room after them are wider.
- **`Ctrl`+`C` under a find bar copies from the surface the reader was last in** — it was decided by the query's selection being empty, so a query with anything selected in it made the page's copy unreachable, and one with nothing selected took the chord even though the reader had just clicked into the page. The query is what a `Ctrl`+`F` or a press in it focuses, the page is what a press or a drag in it focuses, and the focused surface's selection is the one copied; a surface with nothing selected gives the chord back to the other, which is what it has always done with an empty query.
- **A pinned page's arrow is answered when it is pressed** — the key was counted on the keyboard hook's thread and swallowed there, and the loop read the count at the end of its own wait: a pinned static page waits `50 ms` rounded up to the system timer's own tick, measured at 41–62 ms per press. The hook now signals the loop where the count is written, as the wheel already did; the presses a hand leaves behind are drained in one tick rather than one per tick, so a held arrow stops when the hand does rather than replaying a queue of lines; a move reads the page's own document rather than asking the disk whether the file changed; and a press that scrolls the page shifts the frame already on screen instead of laying the screenful out and painting every run of it again. Measured end to end, handle to the frame on the display: 1.0–1.7 ms for an arrow that moves the caret (was 1.6–6.8 ms, and 41–62 ms when the wait swallowed it) and 4.1–4.8 ms for one that scrolls the page (was 8.9–9.3 ms).
- **A pin whose bubble was never carried comes back at its own box** — a restore placed the window beside the bubble every time, and the bubble stands on the minimize button of the window the restore has just brought back, so a pin cycled through the bubble without the hand moving anything walked sideways by most of its own width per click until it was against the edge of the screen.
- **A pin's taskbar button survives what the Shell does to it** — a minimize this app did not ask for is answered by bringing the pin back, a button the Shell took away is made again rather than needing the app restarted, and the Shell's own **Close** on it ends the pin with a fresh button made for the next one. An activation of a put-away stand-in is no click at all, so a collapse no longer queues the restore that showed the window over the desktop and left it standing; a pin restored from the button comes back in front of what was opened over it rather than in the slot it went down in; and the picture the collapse took is the picture the hover holds, rather than one settled again against a player that is no longer on the screen.
- **`Ctrl+A` and `Ctrl+C` answer for a pinned page only while the pin is the window the user is in** — the chords are polled out of the keyboard rather than received as messages, and a keyboard polled from anywhere is a keyboard pressed somewhere else: `Ctrl+A` used to select the page behind Chrome, and `Ctrl+C` put that page's stale selection over what the user was copying there.
- **A find bar's field takes a space and a `t`** — the keyboard hook swallowed the pin's `Space` and `T` for the whole of a bar's life, and `TranslateMessage` only makes the `WM_CHAR` a text field types from out of a key-down nothing swallowed, so a query could hold neither character. The hook now stands down from the pin's keys while a bar is up and the pin's own window answers them, which leaves the bar's one key table the only reader; `PgUp`/`PgDn` join `↑`/`↓` as steps through the matches, `Tab` is inert, and a held arrow repeats at the machine's own key rate.
- **A press on a pin caption button is not a pick of the file under it** — the hook read a press on this app's own window as a press on the listing beneath it, so the Minimize or Close the user aimed at also handed the pin the item under the chrome.
- **A find bar's `↑`/`↓` step on the press and then repeat while held** — a 350 ms delay and a 90 ms interval, disarmed by the release, by a close and by a capture taken away.
- **A find's marks out-rank the selection band in both themes** — a theme naming neither find colour gets a band derived at 0.65 from the page and its foreground, every other match a 0.75 wash of it, and the current one opaque, so a find reads stronger than the selection's 0.25 rather than as the same grey.
- **`To Bubble` from a pin's right-click menu puts the bubble on the point the row was pressed**, not on the minimize button.
- **A hand-off is not a close** — a pin whose file went to another program is noted as handed off, and hovers are held back until the pointer moves, so no card opens over the program that was just launched.
- **A pinned sound's command bar is asked out only by a hand near its own row**, so a hand on the artwork lights nothing.
- **A hand one pixel above a pin's caption asks for the caption** — the near margin is no place a window can stand over the pin, so only a point inside the pin's box is asked what stands under it, and a move that lands on the caption shows the strip on the spot rather than leaving it as state the loop's next tick finds already true.
- **A pin's caption over a page follows the window** — it rides in a popup of its own, mirroring the pin's band, placed on the pin's own motion, hidden with a minimized or invisible pin, and passing the press through to the pin rather than eating it.
- **A pinned page is no longer left behind by a raise** — with **Always On Top** off, a page is put back above the raised pin by claiming topmost and giving it up, and every activation, restore and display change places it again.
- **A hidden hover no longer empties a pinned page's band**, and **a rendered page keeps the box the engine was handed** rather than being measured as text a second time and placed in a frame the engine never saw.
- **Version bumped** — `0.4.6` in `Cargo.toml` and `Cargo.lock`.

## [0.4.5] - 2026-10-18

### Added

- **A wheel zooms a pinned picture.** Over a pinned picture — a still, an animation, a texture, a developed picture, a design, a drawing, or a page (a PDF, a comic's plate, a book's, a page Office, the render engine or Calibre drew) — the wheel is the pin's: the window grows about the picture's centre until it reaches the size **Fit to Screen** would give it, and the picture then grows inside the window that has stopped, about the cursor over it. A wheel back out takes it down the way it came, and the last of it lands with a small swing past its place. A film is the one refusal: its picture belongs to a player's window rather than to the pin. A picture zoomed past the box it is drawn at is the file laid out again at that size — a page is rasterized at it — so the zoom is sharp as the source allows rather than a stretched copy of the last frame.
- **`Preview Settings → Image → Zoom → Background`** — what fills the band a hand carries a zoomed picture off its window leaves behind it: **Black** (default), **White** or **Checkerboard** (`zoom_background` in `config.ini`). The see-through backdrop is not offered: the band is the window's rather than the file's, and a layered window is hit-tested by the shape of its pixels, so a band left transparent is a band the hand clicks through to whatever stands under it.
- **`beat_hz`** — the band a sound's card follows, which the mark drawn beside the name beats with: the bass under the sound sends ripples travelling out from the dot, and a pause brings them home. `off`, or a top of band between `20 Hz` (the fixed floor) and `16 kHz`; `250 Hz` by default. A setting of `config.ini` alone — the tray carries no row for it.
- **The caches go up to `16 GB` and `128 GB`** — **Cache → Image (RAM)** offers `16 GB` at the top and the three disk caches `128 GB`; a size that is a whole number of gigabytes is named in it (`4 GB`, `128 GB`) rather than as the megabytes it is.
- **`audio_mark_offset_y`** — a vertical shift, in pixels, for the mark beside a sound's name: it is centered on the name line's own ink and this moves it up (`-100`) or down (`100`) from there. `0` by default. A setting of `config.ini` alone, like `beat_hz`; a file naming it `audio_mark_vertical_offset`, `audio_mark_y_offset` or `audio_mark_vertical_alignment` is read the same way.
- **`pin_zoom_bounce_ms`, `pin_zoom_bounce_gate`, `pin_zoom_bounce_speed_ms` and `pin_zoom_bounce_power`** — the landing a wheel zoom out ends with, all four in `config.ini`: how long the image takes to swing back to its place once a zoom out has brought the picture within a notch of the fit (`100 ms` by default), whether the swing past its place is on (on by default), how fast it snaps back (`0` by default), and how far past its place it drifts, as a share of the offset it is coming home from (`0.04` by default).

### Changed

- **A wheel over a pinned picture is the pin's, not Explorer's** — the low-level wheel hook counts the notch for the pin and swallows the message, and the preview loop is woken where it stands: a notch counted while the loop waited for a pin used to be read at the end of the wait, which measured 62 ms — four frames of a 144 Hz display, and the whole of what made a zoom start unevenly.
- **A zoom's travel is paced by the compositor rather than by a sleep** — `DwmFlush` answers "when is the next frame" with the display's own period (6.94 ms on a 144 Hz one) instead of the loop's timer-paced wait, so a zoom moves at the rate the screen can show it, on a machine that takes its frames twice as often as well as one that does not.
- **`Image → Zoom → Background` is `Image → Pin Mode → Zoom Background`** — the tray built `Image > Zoom > Background > Background`, a level holding nothing but a level holding the three backdrops; the backdrops now hang directly inside the picture's **Pin Mode**, below **Shuffle** (`zoom_background` in `config.ini` unchanged).
- **A zoomed picture with a transparent pixel keeps its own backdrop** — such a picture is composited over, and banded with, the backdrop **Background → Image** asks for — see-through included — rather than the zoom's own, so a picture meant to show the desktop through it keeps doing so at any size. An opaque picture still gets **Zoom Background** for the band, and a fill of a transparent backdrop is now left as the caller left it rather than written black, which is what lets the band show the desktop through.
- **A caption narrowed past the point of naming the file gives the name up** — and one narrowed past ten pixels gives the strip itself up, rather than drawing a row of window buttons over the top of somebody's picture. Both come back as the window is widened.
- **What earlier versions left on disk is taken by the installer rather than at every start** — the folders older builds cached their pages in, the video log one of them wrote for every hover, and the update check's leftover time and fetched installer are removed once, by the installer's own pre-install step; a start now sweeps only what a render ended mid-flight left in the temp folder and the browser profiles of the runs that are gone.
- **Version bumped** — `0.4.5` in `Cargo.toml` and `Cargo.lock`.

### Fixed

- **A zoom no longer jumps to the size a notch asked for** — the ask is carried out a frame at a time, nine tenths of what is left of it within a tenth of the second and the rest of it behind that, so a spin of notches is a movement the eye can follow rather than a series of steps; a notch in and a notch out leave the picture exactly where it was.
- **A window coming down is placed about its own middle** — a picture a hand has panned sits off its window's centre, and a zoom that placed the smaller window about the picture's carried centre walked the window across the screen with the pan; the window is placed about itself, and the pan goes on moving only the picture within it.
- **The mark beside a sound's name is centred on the name line's own ink** — the middle of the line it marks rather than the middle of the line box, so a short name and a tall one are marked at their own text; `audio_mark_offset_y` above is the hand's own nudge from there.

## [0.4.4] - 2026-10-17

### Added

- **A collapsed film's bubble keeps showing the film.** The picture is a second decode of the file, read beside whichever player is playing it — mpv included — and cropped to the circle, so a pin put away is the film rather than the still grabbed when it went down; a pause holds the frame the film stopped on, and the ring takes that frame's colour.
- **A bubble's bars are read from this app's own sound** — seven bands, thirty times a second, off a process-loopback capture of this app and its players — so a film's bars move with its soundtrack and another window's audio cannot move them. A machine without process loopback keeps the peak-shaped row.
- **`pin_double_click_center` and `pin_double_click_reset_size`**, both on by default: a double-click on a pinned window puts it back to the size its kind's scaling setting asks for, and in the middle of the screen it stands on. Off, the reset is left out and the move stands alone.
- **A raw `.aac` states its own length** — counted from the stream's ADTS frames rather than read off the bit rate, so the card's bar and clock match the sound. A stream cut off inside its last frame, one followed by an ID3v1 tag, and headers naming any object type are all counted.

### Changed

- **Hardware acceleration is mpv's switch alone** — `Performance → Hardware Acceleration (MPV Only)`: FFmpeg's player is given no decode device, and the probe that chose one is gone, every device it walked having died in the graphics driver on a real machine. A machine with no `mpv.exe` shows the row greyed and unchecked, and a `config.ini` that says on is read as the off it means there.
- **A pinned window keeps the pixels it already has when its box shrinks** — a picture, a PDF page, a page Office or the render engine drew, and a comic plate are no longer decoded again for a smaller box; the larger frame is stretched instead, so enlarging back (the double-click above, above all) is sharp at once. A page of text, a listing and a sound's card are still laid out again at the box they are given.
- **`Image (RAM)` holds `128 MB`** rather than `64`, which is a full-screen frame and the three after it, so a hover sweep no longer evicts the frame a pinned window is decoded at.
- **A walk landing behind a collapsed pin begins its player off the desktop** rather than leaving a window where the pin's band was.
- **Version bumped** — `0.4.4` in `Cargo.toml` and `Cargo.lock`.

### Fixed

- **A step of a pin's walk that cannot be shown no longer takes the bubble down** — what a bubble shows is a picture this app holds, so a film whose player has gone leaves the circle standing on that picture rather than coming apart.
- **A held click is given up when the pin's own window moves** — a double-click that centers a pin no longer answers the press against the point the move uncovered.
- **The caption's walk chevrons, the transport bar's marks, the close cross and the collapsed pin's bubble are drawn with coverage**, not in whole columns, so an edge is a curve rather than a staircase at a scaled display.

## [0.4.3] - 2026-10-16

### Added

- **A collapsed pin's bubble holds a `Pause on Minimize` row** — the switch a film's and a sound's own `Pin Mode` carries, now reachable where the pin is away: a right-click on the bubble lists it below the playback rows and above the ways out, and a press turns `pin_pause_video`/`pin_pause_audio` over and leaves the panel up with the mark it now carries. Only a bubble whose kind plays something has the row — a picture's and a document's keep the menu they had, the same rule the hold above it answers — and the row spells the switch the way the tray spells it.

### Changed

- **`Pin Mode → Pause Minimized` is `Pause on Minimize`** — the film's row and the sound's row, one under **Preview Settings → Video** and one under **Preview Settings → Audio**, renamed to the word the bubble's own row spells, so the bubble and the tray name the one switch alike (`pin_pause_video`/`pin_pause_audio` unchanged).
- **`Volume` is a submenu of the tray's own** — one menu with a **Video** half and an **Audio** half, below **Background** and above **Performance**, instead of a level row in each kind's half of **Preview Settings**; the film's own rows are now **Subtitles**, **Loop** and **Pin Mode**, the sound's **Loop**, **Seek** and **Pin Mode**. Every key in `config.ini` is unchanged.
- **`Preview Settings → Image → Pin Mode`** — the picture's one row, **Shuffle**, now hangs under a `Pin Mode` of its own the way a film's and a sound's do.
- **The pinned sound card's right-click menu holds the whole of what a pin does** — the two ways out, the file's own playback, then the **Shuffle** / **End of File** / **Seek** rows it already had, with hairline rules between the groups; the bubble's menu gained the same playback rows. It floats in a popup window of its own, so the card no longer clips its own rows, and its hold row reads `Pause` or `Play` for what it will do rather than what it already does.
- **Version bumped** — `0.4.3` in `Cargo.toml` and `Cargo.lock`.

### Fixed

- **A pinned sound's `End of File` is the rule its player is begun with** — a walk step began the new sound's player with the hover's `Loop` instead of the pin's row, so it went round on its own for ever with the card's clock held at the whole of the file. The rule is now asked of the pin rather than of the file, and a sound a pin takes up over is begun again at the second it is at where the two disagree. (A rule changed while a self-looping player is already going is not read until that player is begun again.)
- **`config.ini` is written whole, or left as it was** — a save cut short left a zero-length file that neither read could read, and the write it then asked for put every default over the whole file. The text is now written beside the file and moved into place, and a file holding no setting of its own is not read as a configuration at all.
- **A read of `config.ini` no longer writes it** — the read that runs while you hold the file open used to put right what it found wrong _in the file_, and an editor writes that file in place, so a read a second later took half of what you were writing and put the whole of the app's own view back over the rest of it (and an outside write was answered with the app's own agreeing with it). What the file says is now taken, and what it has wrong — a value that cannot be read, a key of your own, a name an older build used — is put right in what the app is using; the file keeps what it said until the next start, which is what writes it, the way it always did.
- **A save is on the disk before it is moved, and a move the machine refuses no longer loses the setting** — the staged text is flushed rather than left in the write cache when the move puts it under the app's own name, so a machine that lost the power after the move is not left with a file holding nothing; and where the file is held open where it stands, shared without the delete a move needs, the text is written into the file that is there rather than dropping a setting for the next read to take back off the app.
- **A level moved up on a film begun at `Volume → Video` 0% is heard** — a film begun at nothing is launched with `--aid=no` rather than `--volume=0` (the decoder is saved the whole soundtrack), and mpv takes a level as a property, so a knob dragged up moved a number nothing played at and the film stayed silent for the whole drag — only on mpv, since FFmpeg's player and the media engine took the level in hand. The launch now records that its player was given no audio, and the first level above nothing asks the track back with the level already in hand, so the audio opens at the level dragged to rather than at mpv's own default of full; the rest of the drag and a release's re-assertion are the level alone.
- **A press on a collapsed pin's bubble gives the park back** — with **Pin Mode → Pause on Minimize** on, collapsing a pin parks what is playing, and a press on the bubble's **Play** did nothing because the toggle it reached read the park as a hold the presses had to live with. A collapsed pin now answers a play/pause for itself: a park standing is given back the road a restore gives one back, a bubble with nothing parked — a film the bar had paused before the pin went down — takes the ordinary toggle, and either way the press claims the bubble, so the switches cannot write a park over the playback just asked for. The bubble's own row also names what the press will do.
- **A level is claimed only where the player heard it** — the media engine now answers whether a session took a level, so a pin let go of while its knob was still under the hand no longer leaves the bar claiming a level nothing is playing at; the level stays owed to the next player.
- **A player left behind by a killed run can no longer answer for the next run** — mpv's pipe name carried a counter alone, which begins at one every run, so a player still listening from a killed run held the first name the next run asked for and answered its lines as though they were meant for it (the app reported success for a command the player on screen never heard). The name now carries the app's own process id beside the counter.

## [0.4.2] - 2026-10-14

### Added

- **`Preview Settings`** — every kind's preview settings under one row, below **Preview Types**: **Image**, **Video**, **Audio** and **Text** are its halves, named for the kinds and nothing else. The three top-level `Video Preview`, `Audio Preview` and `Text Preview` rows and the old top-level `Volume` menu are what it holds (the film's own and the sound's own `Volume` halves, with **Subtitles**, **Loop**, **Pin Mode** and the two **Seek** rows under the kinds that have them).
- **`Preview Settings → Image → Shuffle`** — whether a _pinned_ picture's walk takes the next file at random, the question the film's and the sound's own **Shuffle** rows ask (`pin_mode_image_shuffle`, off by default). A picture has no end of file to jump from, so its shuffle is the caption's buttons and the keyboard's alone.

### Changed

- **The pinned sound card's `End of File` choice is called `Jump`** — the word the tray's own `End of File` submenu lists the walk's next file under, so the card and the tray name the one choice alike; the setting in `config.ini` is unchanged (`continue`).
- **`Theme` moved into `System`** — the look is the app's own rather than a text preview's, so it sits between **Run at Startup** and **Edit Config.ini**, beside the other rows about the app (the `theme` key in `config.ini` is unchanged and now reads in the `General` group).
- **`Subtitles` moved out of the film's `Volume` list** — a switch about a film's picture rather than a level, so it stands beside `Volume` under `Preview Settings → Video`, above **Loop**.
- **`Pin Mode → Minimize → Pause` is `Pause Minimized` under each kind's own `Pin Mode`** — one row under **Video** and one under **Audio**, because what should go quiet while a pin is away is a question about one kind of file and each kind's own submenu is the one that knows which of the two is on screen (`pin_pause_video`/`pin_pause_audio` unchanged).
- **`Pin Mode → Navigation Files` is `Pin Mode → Navigation`** — what the three anchors under it are.
- **A sound's two seek rows are both `Seek`** — `Preview Settings → Audio → Seek` (was `Audio Seek`) and `Pin Mode → Seek` (was `Pin Mode Audio Seek`); **Pin Mode** now sits below **Seek**, so the rows a hovered sound plays by come before the ones only a pinned sound has.
- **The walk a pin steps along is the pin's own session** — one playlist spawned when the pin is taken up and stepped along, rather than a fresh draw on every press. A pick in the walk's own folder carries the playlist on (the next press re-anchors to the file picked) and one in another folder starts fresh, the walk rides with the window across a take-up in the same folder, and it is parked while its file's box is still waited on. The playlist is reconciled against the folder on every step, so a file copied in joins it and one moved or deleted leaves it.
- **Version bumped** — `0.4.2` in `Cargo.toml` and `Cargo.lock`.

### Fixed

- **`Volume → Audio` 100% is heard at full** — a sound takes its level through an FFmpeg `volume=` filter at every level now, the way a film's soundtrack already did; the player's own `-volume` knob is silent at 100. A level of nothing is still a player playing silence, so the card's clock moves and the file stays seekable.
- **A pinned sound's play/pause press works again** — the media engine is asked whether it is _holding_ a sound, so the second press lets a held sound go instead of asking it to hold again, and the card's own play button reads the same answer so the glyph says which of the two the sound is at. The bubble parked an engine-held sound as a playing one and played it on the restore; what is not going is not parked now.
- **A tray `Volume` row reaches a pin that is already playing** — the level is applied at once rather than at the next file the pin is shown, pinned and collapsed alike.
- **A sound that loops on its own keeps its clock going round** — a player FFmpeg begins with `-loop 0` never exits, so the card's clock was held at the whole of the file from the second pass on; it wraps the way a pinned film's transport clock does (`elapsed % duration`), read off the two facts the player was begun with so the clock and the flag cannot drift.
- **A swap onto a film FFmpeg's player or mpv draws holds for that player's window** — the wait is asked the way each route can answer it, the engine's own frame for a native film and the player's window for one of this app's own: a portrait file swapped for a landscape one had its last frame stretched across the new box for as long as the replacement took to open, seek and present, and a second pick landing fast behind a first left the same stretched frame behind.
- **A right-click at the middle of a pinned sound's card opens the menu at the click and as wide as its own rows measure to** — the tier that decides which side a flyout opens on reserved a fixed room on each side, so a window too narrow for the two panels split them between its edges and moved the menu to the left edge; the flyout is narrowed into the room the roomier side leaves beside the menu, and the two panels are split only where the window is narrower than the menu's own panel.
- **A press anywhere other than the collapsed pin's bubble puts that menu away** — the bubble is `WS_EX_NOACTIVATE` and only the foreground window's capture carries a press that lands elsewhere to it, so the ask the Explorer hook publishes for a press outside the pin's own window now counts the bubble's window as the pin's, and the loop puts the menu away the way it puts the card's, with the pin left standing.
- **A flyout that opens to the left of the pinned sound card's menu stands tight against it** — the flyout's panel is measured from its own rows, but the side it was placed on was answered in the room the placement reserves, so a panel narrower than that reserve stood a whole reserve of empty space between itself and the menu it came from. It now stands the same few pixels off the menu's own left edge that a flyout to its right stands off its right.
- **`End of File` is taken on the native and FFmpeg routes the way mpv's already was** — a film is taken to its top a tenth of a second _before_ the end of its pass while it is still playing, so what a press then plays is a held film from the top rather than the last frames of a player standing at its end. The engine's own end is read where the walk's end already was, so `End of File → Pause` no longer leaves the bar counting past the film; a film the media engine plays takes its loop from the film on screen rather than from the hover it was adopted from, so a pin set to `Loop` is not rewound by a hover's own loop and one set to `Jump` is never begun again; and a film that has finished is not asked for frames, so a finished film is not read as a file this side cannot draw.
- **`End of File → Pause` holds a pausing FFmpeg film instead of beginning it again** — that player can be told nothing while it runs, so the tick's rewind answered the end of a pass by ending the player and beginning another at zero, playing, which made a `Pause` play the film again pass after pass. The film is now stopped where it stands by the pause key the play button already posts, with the bar held at its own length and the play button a finished film is owed.

## [0.4.1] - 2026-10-14

### Added

- **`Pin Mode → Navigation Files → Video / Audio / Others`** — the single All/Same Category switch is now three anchors, one per kind of first file pinned: Video and Audio start at `Same Category`, Others at `All Files` (`pin_nav_video`, `pin_nav_audio`, `pin_nav_others` in `config.ini`; a file holding only the old `pin_nav_file_types` seeds all three). The walk is anchored on the first file pinned, so `PDF → MP3 → DOCX` walks end to end while `MP3 → FLAC` never reaches a document.
- **Shuffle plays every sound once before repeating** — a shuffled walk draws from a bag and removes each pick, rebuilding minus the file on screen when spent, so one permutation is every sound in some order; a deleted file is tried once then gone.

### Changed

- **The pinned sound card's menu** — both of its rows that open a flyout carry the submenu arrow `▶` rather than the last row alone, each flyout's panel is measured from its own labels instead of being held at one fixed width — so a flyout of short rows hugs them rather than carrying an empty side — and the **End of File** flyout opens beside the **End of File** row it was opened from, where it used to hang off the **Seek** row.
- **`Video Preview` and `Audio Preview` submenus** — a film's settings and a sound's are two menus rather than one `Volume`: each kind's own submenu carries its **Volume** (the levels, with **Normalize (Expensive)** and **Remember Level** above them and the film's own **Subtitles** row at the top of its half), a **Loop** row for whether a _hovered_ one goes round at the end of its file, and a **Pin Mode** submenu for what a _pinned_ one does. The sound's `Audio Seek` is its **Seek** row, and where a _pinned_ sound starts is a second **Seek** inside `Pin Mode`, beside the pinned rows it shares with a film — so every row of a pinned film or sound sits under that kind's own **Pin Mode** rather than beside the rest of the kind's settings.
- **Preview menus carry Loop, End of File and Shuffle rows** — the hover's `Jump → Enabled (Loop)` row is a plain `Loop` check of the kind's own submenu, `Pin Mode Jump` is `End of File` (`Pause` / `Loop` / `Jump`, written in `config.ini` as `pause`/`loop`/`continue`, where `Pause` is what the master switch's off meant), and a new `Shuffle` check is asked both by a forward step and by an end of file that is `Jump` (`video_loop`/`audio_loop`, `pin_mode_*_end_of_file` and `pin_mode_*_shuffle` in `config.ini`; a file that still holds `pin_mode_*_jump_enabled = false` is read as `Pause`, `pin_mode_*_jump = shuffle` as `Continue` with `Shuffle` on, an old `pin_mode_*_loop` as `Continue`, and `video_jump_enabled`/`audio_jump_enabled` as `video_loop`/`audio_loop`). The pinned sound card's own menu carries the same three rows instead of its switch and its rule.

### Fixed

- **`End of File → Pause` takes a pinned mpv film back to its first frame and holds it there** — mpv stands on the last frame instead of exiting, so the end of the pass is read off this app's own clock, and what it does is what a loop does: the film is sent back to the beginning down the player's own pipe and left standing there, with the seek bar and the play button at the start. Pressing play is then the ordinary resume every held film already has — one command, no seek — where a film left standing at its _own_ end had to be taken to the top by the press, and a seek written against a player that has stopped at the end of its file is a restart of the film's whole decoding chain, window and all: the band showed the desktop for as long as that took.
- **Two mpv commands written in one breath both reach the player** — mpv serves one client at a time and frees the next instance of its pipe only once the one before it has been let go of, so the _second_ of the two lines a wrap is written as — the seek, and the line that settles the state the film is left in — was refused with `ERROR_PIPE_BUSY` about once in twenty pairs (measured against mpv 0.41.0). A refused line is a command the player never reads, and for the hold that is a film playing on from its first frame instead of standing there. A busy pipe is now waited out rather than answered `false`.
- **`End of File → Jump` waits for an mpv film's own end** — the five-second grace for reading a _gone_ player's exit as the film finishing was applied to a player still up, so the next file began with five seconds of the last one still playing and the bar sweeping past them. A player that is still up has done no going: the walk is asked at the film's length and no earlier, and the bar is held at the end while the walk is out.
- **An mpv film's bar and clock keep step across a wrap and a seek** — a wrap is a pipe command rather than a relaunch, so the pin's seek bar was never moved with it and stood a whole pass behind the picture after the first wrap, while a seek taken into the last moments of a film was undone by the very next tick, which wrapped the film back to zero. Both clocks are now moved onto the second a pass begins or is taken to, and a pin taken up on a film the hover was playing starts its bar where the picture is rather than at zero.
- **A walk taken under a bubble no longer raises a second window** — with `End of File → Jump`, a pinned sound minimized to its bubble (or its taskbar stand-in) advances to the next file with the window still collapsed: the next file's path, box and media are swapped in behind the bubble, and the restore brings the window up on that file where the take-up used to put the window up beside the bubble.
- **`End of File → Loop` loops a pinned mpv film** — mpv is the one player a take-up does not begin again, so the clock a pass is measured by carried the _hover's_ end of file into the pin: with `Video Preview → Loop` off and `End of File → Loop` on, a pinned film was never begun again, and a pin set to `Jump` was rewound by a hover's own loop. Which end of file a film is looped under is now asked when its end comes, of the film on screen, and the clock is kept for every film mpv is begun with.
- **An mpv film plays again when its loop wraps** — a pass that ended before the wrap reached the player left it paused at the film's first frame, silent, and only a pause-and-play press brought it back: mpv paused a film at the end of its pass and kept that pause across a seek, so the wrap sends the play beside the seek. The loop's clock is also started where the film does — when the player's window comes up, rather than when its process was spawned (`rebase_video_loop_clock`) — so the rewind is posted at the end of the pass instead of cutting the player's own startup time off the film's tail on every pass.
- **A pinned sound shows a play button again when it finishes** — with `End of File → Pause`, a sound that plays to the end of its file now stops there: the card reads `0:00` with its seek bar empty and the button a play button, and pressing it plays the file again from the beginning. Until now the button stayed on pause for a sound that had already ended, on both players. The file is also left at its beginning for `Audio Preview → Seek → Remember`, rather than remembered at the end it stopped at.
- **Pin walk stays anchored on the first file pinned** — every **Next**/**Previous** press reuses the first file's Navigation Files anchor instead of recomputing it from the file the pin is standing on, so a walk taken up on a document (`Others → All Files`) no longer collapses to sounds-only after stepping onto a sound (`PDF → MP3 → MP3` is `PDF → MP3 → DOCX → …` again), and the shuffle bag follows the same anchor. A listing pick (or anything else that clears the walks) re-anchors on the file then pinned.
- **Same-file shuffle no longer leaves the seekbar standing** — a swap onto the file already playing restarts it at the swap's start instead of carrying on.
- **Version bumped** — `0.4.1` in `Cargo.toml` and `Cargo.lock`.

## [0.4.0] - 2026-10-13

### Added

- **MPV backend for video** — where mpv is installed it plays films, pickable under Engine → Video Engine or automatic under Best.
- **PIN MODE (Space) QuickLook** — press Space on any preview to pin it as an always-on-top window; arrows walk sibling files, Esc closes.
- **Best video routing** — Best plays small files with the built-in engine and large ones with FFmpeg, with automatic fallback.
- **Videos play through `mpv` where it is installed** — a third video engine beside the media engine Windows has and FFmpeg's player: `mpv.exe` is looked for beside the app, in your `PATH`, and in the two folders its installers use (`MPV Player` under Program Files, and SVP 4's own `mpv64` — a `winget` install and an SVP install each leave one), **Best** plays every video through it where one is found (a machine without it behaves exactly as it does today), and **Engine → Video Engine** offers **MPV** — greyed where no `mpv.exe` is installed — written as `video_engine=mpv`. Where more than one is found, **Codecs → Videos → MPV** lists them by path and remembers your pick as `mpv_path`. mpv decodes on the graphics card by default (`--hwdec=no` where **Performance → Hardware Acceleration → Video** is off) and draws its own subtitles — the track the film marks as its default and any `.srt` beside it, so the app extracts no subtitle copy for it and the pinned `T` key does nothing on that route (a film whose own track is unmarked, as FFmpeg's muxer writes them, shows none: mpv follows the file).
- **`Scaling → Image / Video / Animated` shares of the display** — the three bitmap submenus now offer the display's fitted size reduced to a share of it — `75%`, `50%`, `25%`, `10%`, listed under a `By Screen` row — beside the shares of the file's own size (`400%`–`25%`), listed under a `By Own Size` row, with a separator above and below that row. A reduced fit is a scale of its own in `config.ini`, written `screen 75` (and read as `75 of screen`, `screen-75` or `screen75` too), so a share picked from the new group is still the share after a restart; a plain number keeps meaning a share of the file's own size, and the document and sound scales are unchanged.
- **`System → Check for Updates` switch** — the update check is now a toggle, on by default (`check_for_updates` in `config.ini`): on, the app asks GitHub for a newer release at startup and when the menu is opened, at most once every six hours (was every hour), offering `Update Available (vX.Y.Z)` above the `System` submenu where one is found; off, nothing is asked at any moment and the row never appears. The "You're up to date." dialog is gone — a check that finds nothing is a menu that stays as it was. Turning the switch on asks at once rather than at the next menu opening.
- **`Scaling → Audio` `7%`** — the sound's-card scale offers `7%` between `10%` and `5%`; the default stays `10%`.
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
- **`Pin Mode → Minimize → To Bubble / To Taskbar`** — Where a pin put away by its Minimize button goes: the round bubble on the desktop (default — the app never appears in the taskbar), or the app's own taskbar button, which brings the window back where it stood when clicked and closes it on a right-click. The preview and pinned windows themselves never touch the taskbar, whichever is chosen; written as `pin_minimize_to`, old files read as `bubble`.
- **`Pin Mode → Minimize → Pause`** — A pin put away (into its bubble or the taskbar) pauses audio/video and resumes at the stopped second when brought back; Video and Audio are separate switches, both on by default.
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

- **`Volume → Remember Level` on by default** — both halves of the row now start on (`remember_audio_volume` and `remember_video_volume` default to `true`), so a level turned with a pinned window's own knob is the level the next preview is played at, and letting go of the knob writes it to `config.ini`.
- **`Volume → Normalize` off by default** — both halves now start off (`normalize_volume` and `normalize_video_volume` default to `false`), so a hover measures nothing until one asks for the single level: the measurement costs a decode of the file — the whole film, for a film — which the rows now say, reading `Normalize (Expensive)`. A `config.ini` that names either key keeps its value.
- **`Codecs` answers kept per process** — the decoder and picture-codec probes behind the `Codecs` submenu are asked of the machine once per process rather than on every menu opening, so opening the menu no longer re-runs the MFT and WIC enumerations; a codec installed while the app is running shows up at the next restart, while a newly installed FFmpeg is still picked up at the next menu opening.
- **`Volume → Video → Subtitles (Expensive)`** — the subtitles row moved from between the halves to the top of the `Video` half, above `Normalize`, and is renamed for the subtitle copy a hover pays for. Labels and positions only — every command id and every `config.ini` key is unchanged.
- **`Scaling` submenus open with a separator** — the greyed `By Screen` row every `Scaling` submenu opens with now has a separator above it, as well as the one below it, so it stands between separators the way the bitmap submenus' `By Own Size` row does.
- **Tray menu regrouped and reworded** — `Pin Mode → Update Preview → Follow Selection` (was `Enabled`), `Pause in Bubble` (was `Pause Preview`), `Navigation Files → All Files` / `Same Category` (was `Nav File Types → All` / `Category`), singular `Preview Types` names in three separated groups (what is looked at and heard, what is read, what is drawn), `Hover Delay` (was `Delay`), `Avoid → Nothing` / `Filename` / `Filename Column` / `Details` (the `Avoid` prefixes dropped), `Scaling` and `Background` submenu names drop their `Scaling`/`Background` suffixes with `Animated` (was `Animated Image`) listed above `Audio` and the seven display-share submenus opening with a greyed `By Screen` row, `Volume → Remember Level` (was `Remember`), `Performance → Explorer Poll` (was `Tick`) with a flattened `Hardware Acceleration (Video)` row (was a submenu holding one `Video` row), `Engine → Away Timer` (was `AFK Timer`) with `Video Engine` and `Office Engine` in place of `Select Engine → Video` / `Select Engine → Office`, and `Update Available (vX.Y.Z)` (was `Update is available! (vX.Y.Z)`) above a new `System (vX.Y.Z)` submenu holding `Check for Updates`, `Run at Startup`, `Edit Config.ini` and the two resets (was the version on `Config.ini (vX.Y.Z)` with the resets inside it and `Run at Startup` at the top level). Labels only — every command id and every `config.ini` key is unchanged.

### Fixed

- **An mpv film opens at the preview box on a display drawn at another scale** — the window mpv put up was twice the box on a display at 200% (twice as wide and twice as tall: 200% of the screen for a fit) and stayed there until the preview tick put it right a second later, which is the huge window that snapped small on a second display and never happened on the primary. mpv creates its window on the primary display and Windows rescales it to the display it is moved to, keeping its _logical_ size, and mpv takes that rectangle as it stands, so the box is now named to mpv on the primary display's scale and Windows' own scaling lands the window on the box (measured on mpv 0.41.0: the box `2864x1611` came up as `5728x3222`, and asked for as `1432x806` it comes up as the box). FFmpeg's player and the media engine were never affected.
- **`Scaling → Video → Fit to Screen` fills the display** — a video's fit was answered as the file's own size, so a film smaller than the screen was watched at the size it was shot at however large the screen was, while a pinned window and a maximize over the same screen enlarged it; a hovered film's fit is the room now, up as well as down, for every engine. The shares below the fit are unchanged: `By Own Size` percentages and `By Screen` shares are still shares of the film's own size.
- **mpv no longer plays a film on the primary display whatever box it was asked for** — a preview box on a monitor left of or above the primary one has a negative coordinate, and the launch dropped the position for such a box, which is what made mpv centre the film on the primary display instead; the geometry names the box whenever it names a size, and `+-2160` is the negative coordinate mpv reads it as (measured against mpv 0.41.0 with a second display above the primary: the size alone opens at the centre of the primary, the position opens on the second display). FFmpeg's player and the media engine were never affected.
- **`Background → Image` `White` and `Checkerboard`** — the two rows carried command ids the `Pin Mode → Navigation Files` rows had also taken, so a click on either backdrop was answered as a click on the navigation rows instead of setting the backdrop; the navigation rows now sit in the free ids above the pin's own rows.
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
- **Open With opens on the pin's display** — The default program is launched with the pin's own window as parent and its monitor as the hint, and the `Open With` list is carried onto the pin's box once it is up, so a pin on a second display hands the file to a program on that display rather than on the primary one; a program that puts its window back where it was last closed still reopens there, and a list the user has already moved is left where they put it.

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
