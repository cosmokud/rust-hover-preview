# Configuration


Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

The file is watched, and changes apply without a restart. A read takes what the file says and puts right in the app what it cannot read — the file itself is left alone until the next start, so a hand edit is never written over while it is being made. A save writes the whole file beside it, flushed to the disk, and moves it into place, so a save cut short leaves the file that was there; a file holding no setting of its own is not read as a configuration, and is left exactly where it is.

Example, trimmed:

```ini
[settings]
; General
check_for_updates=true
pin_always_on_top=true
pin_double_click_center=true
pin_double_click_reset_size=true
pin_enabled=true
pin_key=mouse3,space
pin_nav_audio=category
pin_nav_image=category
pin_nav_others=all
pin_nav_video=category
pin_update_enabled=true
pin_update_on_hover=false
pin_update_move_to_top=true
preview_enabled=true
run_at_startup=true
silent_auto_update=false
theme=light
toggle_key=f7

; Minimize
pin_minimize_to=bubble
pin_pause_audio=true
pin_pause_video=true

; Preview Types
archive_preview_enabled=true
audio_preview_enabled=true
design_preview_enabled=true
document_preview_enabled=true
ebook_preview_enabled=true
font_preview_enabled=true
image_preview_enabled=true
text_preview_enabled=true
vector_preview_enabled=true
video_preview_enabled=true

; Text Preview
markdown_mode=rendered
render_html=false
text_dont_preview_secrets=true
text_font_scale=125
text_show_extensionless=false

; Timing
hover_delay_ms=0
prioritize_keyboard=true
same_file_rehover_delay_ms=200
settling_delay_ms=0
trigger_key=alt
trigger_key_affect_pin_mode=false
trigger_key_enabled=true
trigger_key_mode=disable

; Placement
avoid_mode=filename
follow_cursor=false

; Scaling
animated_scale=100
audio_scale=10
design_scale=fit
document_scale=fit
ebook_scale=fit
font_scale=50
model_scale=fit
preview_scale=100
text_scale=fit
vector_scale=fit
video_scale=100

; Background
dds_background=white
design_background=checkerboard
font_background=white
html_background=white
image_background=checkerboard
vector_background=checkerboard

; Volume
audio_loop=true
audio_seek=remember
audio_volume=10
normalize_video_volume=false
normalize_volume=false
pin_mode_audio_end_of_file=loop
pin_mode_audio_seek=start
pin_mode_audio_shuffle=false
pin_mode_image_shuffle=false
pin_mode_video_end_of_file=loop
pin_mode_video_shuffle=false
remember_audio_volume=true
remember_video_volume=true
video_loop=true
video_volume=0

; Performance
decode_budget_gb=1
document_cache_mb=256
general_disk_cache_mb=128
image_cache_mb=128
image_disk_cache_mb=512
tick_ms=15
video_hw_accel=true

; Engine
afk_timer_seconds=60
libreoffice_idle=away
office_engine=microsoft_office
office_engine_idle=away
video_engine=best
video_engine_idle=away
video_engine_fallback=true
webview_idle=away

; Advanced
hdr_exposure=0
hdr_tone_map=reinhard
spinner_delay_ms=250
```

## Here's what this .INI does:

### Look and text

- `theme`: `light`, `dark`, or `custom:<name>`.
- `markdown_mode`: show Markdown as `rendered` or `source`.
- `text_font_scale`: 1–1000%, default `125`; archive lists follow it too.
- `render_html`: `false` by default.
  - On, `.htm` and `.html` files are run as the pages they hold by the browser engine, at the share of the screen **Document Scaling** names.
  - That is the one kind of preview a pointer and a keyboard reach: the page's own box holds the pointer, and a click into it gives the keys.
  - Without that engine, they are text like before.
- `text_show_extensionless`: whether a file with no extension at all (`README`, `notes`, a bare script) is previewed as text. Default `false`. The `[text] names` list — `LICENSE`, `Makefile`, `.gitignore` and the rest — is read either way, and a folder is never previewed as text whatever its name looks like.
- `text_dont_preview_secrets`: whether secret and credential files are refused rather than shown as the text they are. Default `true`. Refused: `.env` and its variants, `id_rsa`/`id_ed25519`/`credentials` and their suffixed copies, key, certificate and keystore extensions (`.pem`, `.key`, `.p12`, `.pfx`, `.kdbx`, `.ovpn`, `.rdp`, `.pcap` and the rest), and anything under `~/.ssh`, `~/.aws`, `~/.gcp`, `~/.kube`, `~/.docker` or a coding agent's credential path. Still previewed: public keys (`*.pub`) and the `.env.example`/`.sample`/`.template`/`.tpl`/`.defaults` templates.

### What can preview

- `*_preview_enabled`: turn each preview type on or off. File lists stay normal.
- `extensions` under each kind's section: which files that kind previews.
  - Kinds: `image`, `video`, `ffmpeg`, `audio`, `text` (which also carries a `names` key for files with no extension), `archive`, `office`, `font`, `design`, `vector`, `ebook`, `libre`, `magick`, `peazip`, `calibre`.
  - No dots.
  - An entry with a dot, like `tar.gz`, matches the end of the file name.
  - A name two sections hold belongs to the earlier kind, so the kind decides which switch applies to it.
  - A name the text lists hold is still refused while `text_dont_preview_secrets` is on, the lists being a question about names and that setting a question about what the name holds.
- Which engine plays a sound is the machine's answer rather than a setting: the media engine Windows has is asked first and an installed FFmpeg second.

### Video and audio

- `video_loop` / `audio_loop`: whether a hovered film or sound goes round at the end of its file. Both default `true`; off, it stops at the end.
- `pin_mode_video_end_of_file` / `pin_mode_audio_end_of_file`: what a *pinned* film or sound does at the end of its file — `loop` (default), `pause` (the file left where it ended, the card holding it with a play button), or `continue` (the walk's next file).
- `pin_mode_video_shuffle` / `pin_mode_audio_shuffle`: whether a forward step of a *pinned* walk takes a random file of the anchor's kind rather than the one beside it. Both default `false`; a step back is the file the walk came from.
- `pin_mode_image_shuffle`: the same question asked of a *pinned* picture's walk, which has no end of file to jump from — its shuffle is the caption's buttons and the keyboard's alone. Default `false`.
- `pin_mode_audio_seek`: where in a file a *pinned* sound starts playing — `start` (default), `remember`, `middle`, or `random`.
- `beat_hz`: the band a sound's card's mark beats with — `off`, or a top of band in hertz between `20` (the fixed floor) and `16000`. Default `250`. A setting of `config.ini` alone; the tray has no row for it.
- `audio_mark_offset_y`: a vertical shift in pixels for the mark beside the name — `0` default, `-100` to `100`. A file naming it `audio_mark_vertical_offset`, `audio_mark_y_offset` or `audio_mark_vertical_alignment` is read the same way.
- `audio_seek`: where in a file a hovered sound starts playing.
  - `remember` (default) picks it up where the last hover left it, `start` always begins at the beginning, `middle` drops it half way in, and `random` anywhere at all.
  - The remembered positions are kept in a small file under `%TEMP%\rust-hover-preview\audio`, so they survive a restart; nothing is remembered while another mode is chosen.
  - A video is always played from its beginning.
- A video's engine is the `video_engine` setting; the two lists decide which engine is *asked* about a name — the media engine takes only `[video]` names, and an `[ffmpeg]` name is FFmpeg's alone.

### Memory and cache

- `image_cache_mb`: memory for decoded image frames — one pool shared with pinned windows, so a hover sweep does not evict the frame a pin is drawn from. Default `128`, max `16384`; `0` holds nothing.
- `document_cache_mb`: disk cache for rendered document pages. Default `256`, max `131072`; `0` keeps nothing between hovers but still draws the current one. Pages live under `%TEMP%\rust-hover-preview\document`; the least recently read page is removed first.
- `image_disk_cache_mb`: disk cache for the pictures ImageMagick developed. Default `512`, max `131072`; `0` keeps nothing between hovers but still develops the current one. Pictures live under `%TEMP%\rust-hover-preview\image`; the least recently read one is removed first, and they survive a restart.
- `general_disk_cache_mb`: disk cache for the subtitle tracks a film's embedded ones are copied into, and for the track facts a sound was read for — a card and every row of a pinned sound's playlist band are filled from them without opening the file again. Default `128`, max `131072`; `0` extracts nothing, and a film whose subtitles are embedded is drawn without them. Files live under `%TEMP%\rust-hover-preview\general`, least recently used first.
- `decode_budget_gb`: most memory one hover may use. Default `1`, range `0.25`–`64`. A file over the limit shows no preview.

### Performance

- `video_hw_accel`: whether a film mpv plays is decoded on the graphics card — the tray's **Performance → Hardware Acceleration (MPV Only)**.
  - Default `true`.
  - It is mpv's switch alone: what it reaches is mpv's own `--hwdec=auto`, and FFmpeg's player is given no decode device at all. A machine with no `mpv.exe` reads it as off, whatever the file says.

### HDR

- `hdr_tone_map`: how HDR/EXR light becomes screen values: `reinhard` default, `aces`, `srgb`, or `off`. PNG/JPEG are not affected.
- `hdr_exposure`: stops shifted before that curve. Default `0`, range `-10` to `10`.

### Waiting and engines

- `spinner_delay_ms`: wait before showing the loading spinner. Default `250`; `0` shows it immediately.
  - One delay covers all preview types, and a pinned window being shown another file with it.
  - A pin that has to wait shows a spinner of its own in the middle of itself, drawn over the file it is still showing rather than in place of it, and the window is never hidden while it waits.
  - A hover whose wait ends with nothing to show ends as the cross a pinned window shows for the same failure, in the box the wait had, and comes down three seconds later.
- `office_engine`: `microsoft_office` default, falls back to LibreOffice when needed; or `libreoffice`, which always uses LibreOffice. If LibreOffice is missing, it falls back and the tray row is greyed out. Read where a preview begins, the way `video_engine` is.
- `video_engine`: which engine plays a video: `best` (default), `mpv`, `ffmpeg`, `native`, or `hybrid`.
  - **Best** is the machine's own answer — mpv where it is installed, and otherwise `hybrid` where FFmpeg is installed (the media engine draws a film at or under 3.2 megapixels, FFmpeg's player takes a larger one or one it could not measure) and `native` where neither is.
  - `mpv` names mpv alone, `native` names the media engine alone, `ffmpeg` names `ffplay` alone, and `hybrid` states that size rule plainly.
  - A choice this machine cannot supply is read as `best`.
  - Read where a preview begins, so a film already playing keeps the player it was begun under; a change reaches the next hover, or a pin's next file.
- `mpv_path`: which `mpv.exe` plays, written by **Components → Videos → MPV** where the machine has more than one. Empty, the default, means the first one found — beside the app, then the `PATH`, then `MPV Player` under Program Files, then SVP's `mpv64`. The only setting that holds a path.
- `video_engine_fallback`: `true` (default) lets a chosen engine that cannot play a given file fall through to the engines it may — a chosen `mpv` to both of the others, `native` and `ffmpeg` to each other — and `false` makes the engine you chose stand alone. Read only where `video_engine` names an engine; **Best** is its own walk either way.
- `video_engine_idle` / `libreoffice_idle` / `office_engine_idle` / `webview_idle`: how long each engine is kept after its last file — `away` (default) keeps it while Explorer is reachable and lets it go after `afk_timer_seconds`, a number of seconds is that long after the last page (or film, for mpv), `indefinitely` until the app exits, and `0` lets it go at once. The tray's **MPV TTL**, **Microsoft Office TTL**, **LibreOffice TTL** and **WebView2 TTL** write them.
- `afk_timer_seconds`: seconds with no Explorer window before an engine set to `away` is released. Default `60`, max one day. `0` releases immediately.

### Trigger and position

- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled`: the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched. Default `true`.
- `toggle_key`: the key that turns previews on or off while File Explorer is the foreground window. Default `f7`; it takes the same spellings and alternatives as `pin_key` (`ctrl+f7`, `f7,f8`, …).
  - The key is not a global hotkey: it is only acted on when Explorer is in front, and it is never swallowed, so it still reaches Explorer and other programs as usual.
  - Holding it down is one toggle, not one per keyboard repeat. The tray's **Enable Preview** row names the configured key.
- `trigger_key_affect_pin_mode`: whether the trigger key reaches a pinned preview. Default `false` — the key is not read while a preview is pinned, up or collapsed into its bubble, since a pin is a window you put there rather than a hover for the key to hold back. `true` lets holding the key bring the pin down with the previews it stops. It speaks for `disable`; the `enable` mode is left as it is.
- `pin_key` / `pin_enabled`: the key that pins the preview on screen (`mouse3,space` by default — the same spellings the trigger key takes: `alt`, `f8`, `a`, `space`, …), and whether it is watched. See **Pin Mode** above for what a pin does.
  - A binding may name alternatives and chords: `,` separates alternatives (either one matches) and `+` joins keys that must all be held. The last key of a chord is the one whose press counts.
    - `pin_key=space,page_down` — either key pins.
    - `pin_key=ctrl+space` — Space pins while Ctrl is held; `ctrl+alt+space` needs both.
    - `pin_key=space,mouse3` — a mouse button can pin too (`mouse3`, `mouse4`, `mouse5`; `mouse1` and `mouse2` are not allowed).
  - The same syntax applies to `trigger_key`, where an alternative counts as held only while all of its keys are.
  - Names are case-insensitive, spaces around `,` and `+` are ignored, and underscore spellings such as `page_down`, `page_up` and `caps_lock` work.
- `pin_always_on_top`: whether a pin stays above every other window. `true` (default) keeps it topmost; `false` makes it an ordinary window that other programs can cover and that wears this app's own taskbar button while it is up. The caption's own mark toggles the same setting, a film mpv plays is put in the same band, and the wheel over such a pin is the pin's only while the pin is the window you are in.
- `pin_double_click_center` / `pin_double_click_reset_size`: what a double-click on a pinned window does. Both default `true`: the window is put back to the size its kind's scaling setting asks for, and in the middle of the screen it stands on. Off, either half is left out — with the size reset off, a hand's own size stands and the double-click is the move alone.
- `pin_update_enabled`: whether a pin that is up is shown the file picked next, by click or by keyboard, without moving the window. Default `true`.
  - Off, the pin keeps the file it was taken up on until it is closed — a pin is also a window to read in, and one that swapped its file out from under the hand on every keystroke would be unreadable.
  - The keyboard half follows Explorer, so it is inert while a pinned window holds the keyboard — click back into the listing and it resumes.
- `pin_update_on_hover`: whether that following includes the pointer's own hover, or only what a click or a key asks for. Default `false`; greyed in the tray while `pin_update_enabled` is off.
- `pin_update_move_to_top`: whether a file an update shows also puts the pin at the top of its own band, without taking the keyboard from the listing. Default `true`; greyed in the tray while `pin_update_enabled` is off.
- `pin_nav_image` / `pin_nav_video` / `pin_nav_audio` / `pin_nav_others`: what the caption's **Previous** and **Next** buttons step through, by the walk's anchor — the first file pinned. Image, Video and Audio default to `category`, Others to `all`.
  - `all` is every file this build can preview, so a video steps to the sound beside it.
  - `category` narrows the walk to the anchor's category — image, video, audio, or the anchor's own kind for Others, where a camera raw is a picture and a book is a document.
  - The walk is the folder the pin was taken up in and no subfolder of it, is in the order the Explorer listing is showing, wraps at both ends, and is held per folder, read on the first press rather than when the pin is taken up.
  - A file it reaches that the window cannot show is stepped over, and the walk is bounded by the folder: each other file is offered at most once, so a folder of nothing this build can read ends the walk instead of going round for ever.
  - A file holding only the old `pin_nav_file_types` seeds all four with it; a file written before any of them existed walks as the fresh defaults do.
- `pin_zoom_bounce_ms` / `pin_zoom_bounce_gate` / `pin_zoom_bounce_speed_ms` / `pin_zoom_bounce_power`: the landing a wheel zoom out ends with — how long the image takes to swing back to its place (`100` default), whether the swing past its place is on (`true` default), how fast it snaps back (`0` default), and how far past it drifts as a share of the offset it is coming home from (`0.04` default). A hand-edited value past the ceiling is reduced to it.
- `pin_minimize_to`: where a pin put away by the window's own **Minimize** button goes. `bubble` (default) leaves the round bubble on the desktop and keeps this app out of the taskbar; `taskbar` takes the window off the screen and shows this app's own taskbar button — named for the pinned file and wearing this app's icon, its hover holding the pin with **Previous**, **Play/Pause** and **Next** for a film or a sound — which brings the pin back where it stood when clicked and closes it on a right-click. Read when the button is pressed, so a change reaches the next minimize; a file written before this setting existed is read as `bubble`.
- `pin_pause_video` / `pin_pause_audio`: whether a pin put away into its bubble or the taskbar holds what it was playing where it stands, and starts it again at the second it stopped at when the window comes back up. Both default `true`. They are two switches because a video and a sound are two different things to want quiet.
- `follow_cursor`: `true` = Follow Cursor; `false` = Best Position.
- `avoid_mode`: `filename` default, `filename_column`, `details`, or `off` — what the preview avoids.

### Scaling

- `preview_scale`: `fit`, a share of the display's fitted size (`screen 75`, `screen 50`, `screen 25` or `screen 10` — the display fitted to the media, reduced to that share of it), or a share of the picture's own size (`25`–`400`). Default `100`.
- `video_scale`: same for video. Default `100`.
- `audio_scale`: the share of the display a sound's card is laid out over — `25`, `20`, `15`, `10`, `7` or `5`. Default `10`. The whole card scales with the share (font, height, width), so it's the `10%` card uniformly smaller or larger, and `Text Size` no longer resizes it. A hand-edited value is sanitized rather than fatal: `100` or more reads as `100`, `0` reads as `10`, an unrecognized value is ignored.
- `animated_scale`: same for animated GIF, WebP, or PNG — and for an animated JPEG XL, which is played as a picture rather than as a video. Default `100`. Still GIF/PNG follow `preview_scale`.
- `vector_scale`: percentage or `fit`, based on the screen. Default `fit`; `100` or more reads as `fit`. Covers all Vector drawings. Old `svg_scale` is ignored and removed.
- `text_scale`: how much of the screen a text preview box may take when it opens, for plain text, code, and Markdown. Default `fit`; `100` or more reads as `fit`. `Text Size` still scales the text inside that box.
- `ebook_scale`: PDF page scale against the screen. Default `fit`; `100` or more reads as `fit`.
- `document_scale`: document page scale against the screen. Default `fit`. A workbook's fallback bitmap keeps its own size and is never enlarged.
- `font_scale`: font preview scale against the screen. Default `50`; `fit` means all of it; `100` or more reads as `fit`.
- `design_scale`: design document scale against the screen. Default `fit`.
- `model_scale`: 3D model scale against the screen — how much of the display the picture f3d draws of a scene is given. Percentage of `fit`, default `fit`; `100` or more reads as `fit`, since a scene has no size of its own to take a share of.

### Fonts

- `ttc_face`: which face of a `.ttc` is drawn. `1` is first, max `10`. The heading shows which face came out.

### Backgrounds

- `image_background`: `checkerboard` default, `black`, `white`, or `transparent`. Also used for PDF pages, painted text frames, and document pages. Old `transparent_background` is ignored and removed.
- `font_background`: `white` default, `black`, `checkerboard`, or `transparent`.
- `dds_background`: `white` default or `black`, and only those two. Any other value reads as `white`.
- `vector_background`: `checkerboard` default, `white`, `black`, or `transparent`. Used for SVG pages and metafiles. Old `svg_background` is ignored and removed.
- `html_background`: `white` default, `black`, or `checkerboard`, and only those three — a page is drawn on a page, so transparency is not one of the backdrops it is offered. Used for a page of HTML previewed by `render_html`, run or not; any other value reads as `white`.
- `design_background`: `checkerboard` default, `white`, `black`, or `transparent`.
- `zoom_background`: `black` default, `white`, or `checkerboard`, and only those three — the band a hand carries a zoomed picture off its window is the window's to fill rather than a hole to look through, since a layered window is hit-tested by the shape of its pixels. Asked of a zoomed picture with no transparency in it: a picture with a transparent pixel keeps its own kind's backdrop (`image_background` — **Background → Image**), see-through included. Any other value reads as `black`.

### Old or removed names

- `magick_preview_enabled` and `peazip_preview_enabled`: not read. ImageMagick pictures use `image_preview_enabled`; PeaZip archives use `archive_preview_enabled`.
- `libre_preview_enabled` and `libre_scale`: gone. Use `document_preview_enabled` and `document_scale`.
- `office_cache_mb` and `libre_cache_mb`: gone. Use `document_cache_mb`. If both old ones exist, the larger is read once, then both are removed.
- `text_preview_full_mode`: gone. A pinned text preview comes up in full mode instead, so it scrolls and its text can be selected and copied.
- `pin_nav_file_types`: gone. Use `pin_nav_image`, `pin_nav_video`, `pin_nav_audio` and `pin_nav_others`; a file holding only the old key seeds all four with it.
- `video_jump_enabled` / `audio_jump_enabled`: gone. Use `video_loop` / `audio_loop`.
- `pin_mode_video_jump_enabled`, `pin_mode_audio_jump_enabled`, `pin_mode_video_jump`, `pin_mode_audio_jump` and `pin_mode_audio_loop`: gone, replaced by `pin_mode_video_end_of_file`/`pin_mode_audio_end_of_file` beside `pin_mode_video_shuffle`/`pin_mode_audio_shuffle`. A file holding one is read once into the settings that replaced it — `…_jump_enabled = false` as `pause`, `…_jump = shuffle` as `continue` with shuffle on, an old `…_loop` as `continue` — and the retired keys are dropped on the next save.
- `office_engine_persistent`, `libreoffice_persistent` and `webview_persistent`: no longer written and read only once — `false` reads as `away`, `true` as that engine's `*_idle`, or `indefinitely` where it names nothing — and a save drops them.
- No `magick_idle`, `peazip_idle`, or `calibre_idle`. ImageMagick, PeaZip, and Calibre are converters, not engines kept open. A second hover usually costs only a cache hit.

### Extension list behavior

- A deleted `extensions=` line or whole section comes back with built-in entries.
- An `extensions=` line left empty stays empty.

