# Architecture (compact)

> How the app is put together, for a new contributor and for a tool that has to find something fast. `path::Symbol` is ground truth in code — read the code before changing behavior. Prose says what a thing is _for_ and how the pieces talk; the symbol lists are the map you grep from.

**The whole thing in one paragraph.** Hovering a file in Explorer spawns nothing and opens no document. A low-level mouse hook notices the pointer entering an Explorer list item, asks the Explorer window _at that screen position_ what file lives there, and asks a route what kind of file that is and which of five drawing hands should draw it. If the answer is cheap — a still image, a PDF page, a line of text — an in-process reader produces a bitmap on a short-lived worker thread and a layered, non-activating window composites it over Explorer. If the answer is expensive or inherently external — an Office document, a book, an archive, a video — an engine process is started, tracked on disk, and later retired. Pressing the pin key freezes the hover machinery and grows that same window into a persistent panel you can walk a folder with. Everything below is detail on those five steps: sensing input, deciding a kind, staying inside a memory budget, caching what was already made, and cleaning up every process we started.

**Where to start reading.** [Runtime threads](#runtime-threads-srcmainrs-srcapp-srcshell-srcui) for the process shape, [Media pipeline](#media-pipeline-formatsroutingrs-native_formatsrs-listsrs) for "which kind is this file", [Kind notes](#kind-notes-reader--engine-gate-scale-backdrop) for per-format specifics. [Read guards](#read-guards-cloud_filesrs-content_type-pathsrs-budgets) is the section to read before writing any decoder.

## Runtime threads (`src/main.rs`, `src/app/`, `src/shell/`, `src/ui/`)

> Who runs where. Main boots and owns the tray loop; Preview owns the window; hook threads only sense input; heavy work (Office, probes, loads) runs off those loops so a stall never freezes hover or paint.

The app is a hover preview for Windows Explorer: you point at a file, you see it. That single constraint decides the whole process shape. Explorer owns its UI thread and will redraw it whenever it feels like it, so nothing of ours may ever block it — a preview that takes 200 ms to appear is a preview the user has already moved off. So input is _sensed_ on near-free hook threads that do nothing but notice, every decision that can be slow is pushed onto a thread that has no window to freeze, and exactly one thread owns the visible surface.

The long-lived roles, plus a family of per-hover threads that come and go:

| Thread         | File                                                                    | Role / key symbols                                                                                                                                                                                                                                                                           |
| -------------- | ----------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Main           | `main.rs`                                                               | `RUNNING: AtomicBool`, `CONFIG: Lazy<Mutex<AppConfig>>`; claims `Local\rust-hover-preview-single-instance` (`app::single_instance::InstanceGuard`); COM+DPI init; reaps leftovers (`app::engine_processes`); temp clear; tray loop. `RHP_STARTUP_TRACE` → `%TEMP%\rhp-startup-trace.log`     |
| Preview        | `ui/preview_window.rs` + `ui/preview_window/*.rs`                       | Layered window (`WS_EX_LAYERED\|TOOLWINDOW\|TOPMOST\|NOACTIVATE`, `UpdateLayeredWindow`). Channel-driven, idle sleep; cadences: `STATIC_WAIT_MS`=150, `STATIC_PIN_WAIT_MS`=50, 60s anim tick; `STATIC_MEDIA_REFRESH_MS`=500, `POINTER_HOLD_HEARTBEAT_MS`=500, `HOVER_TOPMOST_REASSERT_MS`=1s |
| Explorer hook  | `shell/explorer_hook/*.rs`                                              | UI Automation + Shell COM poll; `tick_ms` / `explorer_pace` / `ExplorerState`; `EnumWindows` CabinetWClass/ExplorerWClass multimonitor rule. Not joined (blocks in shell). Publishes `HoverLocation`, `AvoidRegion`, `preview_screen_rect`, pointer-item box                                 |
| Wheel watcher  | `shell/wheel_input.rs`                                                  | `WH_MOUSE_LL`, wheel-tick counter                                                                                                                                                                                                                                                            |
| Key watcher    | `shell/key_input.rs`                                                    | `WH_KEYBOARD_LL`; pin/trigger keys; `refresh`; press counter drained per tick. Msg-pump only                                                                                                                                                                                                 |
| Hook helper    | `shell/hook_thread.rs`                                                  | `hook_thread::spawn`: shared pump, publishes tid, joinable                                                                                                                                                                                                                                   |
| Config watcher | `config/config.rs` + `shell/key_input::refresh`                         | Reloads `%APPDATA%\rust-hover-preview\config.ini` without restart                                                                                                                                                                                                                            |
| Update check   | `app/updates.rs`                                                        | Start + tray-menu build; 1 req/hour (mem only); `version.txt` @ `releases/latest/download` via WinHTTP; Auto/Manual/Cancel dialog (`WH_CBT`); `MZ`+len check; `/S /R` handover                                                                                                               |
| Office worker  | `engines/office_render/worker.rs`                                       | First render spawns; ends when all families idle. STA + pumped (OLE). Never joined/waited                                                                                                                                                                                                    |
| Probes/loads   | `ui/preview_window/load.rs`, `video_probe.rs`, `audio`/`pin_media_load` | Per-hover threads: `VideoProbed`, audio probe, `PinLoad`, `measured_off_the_tick`. `catch_unwind` → "no preview"                                                                                                                                                                             |

**Window styles.** `WS_EX_LAYERED` means the window has no background of its own — every pixel arrives via `UpdateLayeredWindow` with an alpha channel, which is what lets a preview be a rounded, partly-transparent card floating over a folder view. `TOOLWINDOW` keeps it out of Alt-Tab. `NOACTIVATE` is the important one: the window can be hovered and clicked but must never take keyboard focus, because taking focus from Explorer would cancel the selection the user is about to preview. `TOPMOST` because it has to stay above the Explorer window it is drawn over.

**Cadences.** The preview thread is not an event loop with a 16 ms tick — it sleeps on a message and only wakes when it has something to do. A static preview redraws at most every `STATIC_WAIT_MS`; a pinned one redraws four times as often (`STATIC_PIN_WAIT_MS`) because its caption carries buttons the user is actively aiming at. Media needs a steady heartbeat rather than redraw-on-message, so `STATIC_MEDIA_REFRESH_MS` and `POINTER_HOLD_HEARTBEAT_MS` keep a playing preview alive and a held pointer recognised. Animation runs on a 60-second tick because it only needs to notice a new frame source, not drive one.

**Hook threads never touch pixels and never block.** `WH_MOUSE_LL` / `WH_KEYBOARD_LL` callbacks run inside the OS input path; if one blocks, the whole desktop stutters. They set a flag or bump a counter and return in under a millisecond. Everything expensive — reading the Shell view, probing a video, launching an engine — happens later, on a thread with no window to freeze.

**STA** (single-threaded apartment) is the COM threading model Office requires: every COM object an out-of-process Office app hands back belongs to the thread that called for it and must be released on that thread. The Office worker is therefore STA and runs a message pump (`OLE`), because without a pump an OLE call would deadlock. It is deliberately never joined — shutting an Office app down cleanly is not something this process can be trusted to finish before exit.

Single-instance: mutex first, before DPI/COM/config/threads. 2nd proc exits silently. `Local\` = per-session.
The named mutex is the cheapest possible arbiter of "am I the one?". Taken before anything else, two instances launched at the same instant cannot both load config, both install a hook, and both put a window on screen. `Local\` rather than `Global\` scopes the name to the logon session, so a second user on the same machine gets their own copy instead of being refused by the first user's process.

AFK/engines: `app::afk` clock from hook's `ExplorerState`; non-`Persistent` engines die after `Engine → AFK Timer` (default 1min) with Explorer unreachable; `Persistent` uses idle TTL instead.
Engines are expensive to start, so the app keeps them warm between hovers. The clock is not "time since last hover" — it is "time since Explorer was last reachable and focused", because a preview engine the user cannot see is a process the user cannot tell is running. `Persistent` inverts the rule: a user who marks an engine persistent wants it always up, so it is governed by an idle TTL instead of AFK.

## Source layout

> Where code lives. `app/` = process lifetime, `config/` = settings, `shell/` = input sensing, `ui/` = window, `formats/` = kind decisions, `readers/` = in-process decodes, `engines/` = out-of-process draws, `text/` = painted pages.

```
src/main.rs                 RUNNING, CONFIG, thread boot + tray loop
src/paths.rs                plain_path() — strip \\?\ (\\?\UNC\ keeps server) before Shell/browser/media use
src/app/                    single_instance, startup (registry sync/repair), updates, engine_processes (job+record), afk, dialogs
src/config/                 config (AppConfig, PreviewType/PreviewScale/..., DEFAULT_*, sanitize_*, save/ordered_text/differs/repair_older_lists), theme_files
src/shell/                  explorer_hook/, wheel_input, key_input, hook_thread, cloud_files, pin_navigation, tray/
src/ui/                     preview_window.rs + preview_window/ (pin_*, layout, load*, paint, media_*, video_*, tick, displays, ...); pin_window, displays seam (Desktop real vs recorder)
src/formats/                routing, native_formats::NativeJob/job_for, content_type::{Probe,Content}, head, codecs, lists (16 rows), *_formats (see below)
src/readers/                wic_image, dds_image, bcn/{bc6h,bc7,block_formats}, webp_image, jxl_image, heif_sequence, tone_map::{ToneMap,Curve},
                            psd_image, project_image (+cdr_image), eps_image, metafile_image, svg_preview, font_preview/{font_tables,specimen},
                            pdf_preview::{BookPage,book_page}, comic_preview, archive_listing::entry_bytes, office_preview::{measure,render,WAITING_BOX,source_kind,page_is_workbook_picture},
                            audio_track::{Track,Player,Probed,gain}, video_player/{session,playback::{Crop,Picture},frame_copy}
src/engines/                office_render/{worker,engines,renderers,com} (RenderRequest), libreoffice_render (warm), imagemagick_render::Developed,
                            peazip_render (Backend::of), calibre_render (ebook-convert), webview_preview/{engine,host,environment,pages,api}, document_cache, supervisor::{Adapter,Worker}
src/text/                   text_preview/{document,frame::{TextFrame,TextPreviewOptions,FrameLine,FrameRun,Selection,ScrollBar},layout,markdown,source},
                            text_paint, text_theme::LoadedTheme, archive_preview, audio_preview/{card,page}, pin_chrome/{caption,transport,bubble,primitives}
```

**readers/ vs engines/ is the single most useful distinction in the tree.** A _reader_ runs inside this process and returns a bitmap or a block of text; it is fast, cancellable, and free once it returns. An _engine_ is another process, usually a tool the user installed (LibreOffice, ImageMagick, PeaZip, Calibre, FFmpeg's `ffplay`) or Windows itself (Media Foundation, WebView2). Engines cost a process launch, cannot be cancelled mid-flight, and need a reaper — so the app reaches for one only when a reader genuinely cannot do the job.

**formats/** holds no rendering at all. It answers questions about a file: what is it, who should draw it, is it a list, is this codec present. Keeping it separate is what lets the hook thread, the loader thread, and the layout code ask the identical question and get the identical answer.

**text/** holds the things that are painted rather than decoded — a page of text, an archive listing, an audio card, and all the pin chrome. They share one paint path and one theme, which is why they live together.

## Media pipeline (`formats/routing.rs`, `native_formats.rs`, `lists.rs`)

> Single decision point mapping a file to a kind and a drawer. Hook, loader, and layout all ask the same order so a file never previews as one kind and renders as another; `resolve` answers kind + drawer + backdrop in one call.

Three different callers need three different amounts of this answer. The hook needs the kind alone and must return in microseconds. The layout needs to know whether the thing will be a window of its own before it can place anything. The loader needs to know who draws. Rather than let each ask its own partial question — and drift — there is one kind order and one function that reads it.

Kind order (load-bearing, asked by hook + loader + layout identically): `video, audio, ebook, archive, peazip, calibre, office, libre, magick, design, vector, text, font, picture`.
The order is not alphabetical and not aesthetic; it is a set of tie-breaks for extensions that two systems both claim. Video goes first because only the file's actual bytes can settle whether `.ts` is a transport stream or a TypeScript source, and only one of those two answers can come from the bytes. Pictures go last for the mirror-image reason: `.dds` and `.svg` are claimed by both a name list and a content sniffer, and a picture reader is the safe last answer.

- Video first (only content settles `.ts`/`.mts` vs text); pictures last (`.dds`/`svg` double-claimed).

Three callers want three different amounts of this answer, and all three read the one table. `kind_of` opens the file and is only legal on a hover. `kind_of_name` never does — it exists for the pin walk, which must classify hundreds of siblings without opening any of them. `named_as` is the exhaustive `match` that answers "would this kind claim this name", used by layout to reason about a specific kind it already knows about. A test asserts all three agree on the same inputs, which is how the three cannot rot apart.

- `routing::kind_of` (content+name, hover gate) vs `routing::kind_of_name` (name only, folder walk — no hydration) vs `routing::named_as(kind)` (per-kind check for layout). One map+order; `named_as` exhaustive; test asserts agreement.
  `kind_of` opens the file and is only legal on a hover. `kind_of_name` never does — it exists for the pin walk, which must classify hundreds of siblings without opening any of them. `named_as` is the exhaustive `match` that answers "would this kind claim this name", used by layout to reason about a specific kind it already knows about. A test asserts all three agree on the same inputs, which is how the three cannot rot apart.

- `routing::resolve` → `Route { kind, DrawnBy, backdrop_of, ... }`; `DrawnBy` = who draws; `Page/Drawing/WebPage/Backdrop` enums; `engine_kind_of`, `html_is_engine_drawn`, `page_runs` vs `draws` (program vs picture).

`resolve` exists so a caller who wants the whole route pays for each probe once instead of each sub-question re-finding the file for itself. `Route.content` is what the bytes said, `Route.named` is what the name's lists claim, and `Page`/`Drawing` split the two kinds that are really two kinds each (book vs comic, SVG vs metafile). `DrawnBy` is the answer the window actually cares about — `Native`, `MediaEngine`, `Card`, `WebView`, or `Nothing` — and note it deliberately excludes engine kinds: _which_ engine is a fact about the machine, not about the file. `Backdrop` is a second exhaustive match so the six `current_*_background` settings are read once, in one place.

- `native_formats::job_for` → `NativeJob` (in-app reader for kind); `content_type::Probe` → `Content/head` (16B → 4KB); `head` still-vs-moving: GIF blocks, WebP `ANIM`, PNG `acTL`, ftyp brands, JXL header via decoder.

`content_type::Probe` is the read-and-forget half: sixteen bytes tell you most of a format, and if they do not you escalate to a 4 KB window. The still-vs-moving question is the interesting one, because a GIF and an animated WebP are the same extension to a user and completely different to a preview. GIF has no animation flag at all, so its frames are counted; WebP carries `ANIM`; PNG carries an `acTL` chunk before the image data; MP4 and friends carry an `ftyp` brand list; JPEG XL has no usable header of its own and is asked of its decoder.

- `*_formats` keep only list-unanswerable Qs: `text_formats` (dot-file + name list), `archive_formats` (`tar.gz`), `video_formats` (ffmpeg split + MPEG-TS sync probe), `office_formats::{app_for,page_engine,OfficeApp,OfficeContainer}`, `ebook_formats::page_spelling`, `libre_formats::engine_page_kind`, `magick_formats::is_engine_picture`, `peazip_formats::Backend::of`; `font/calibre/audio` = row/table probes.
  Each `*_formats` module exists because one specific question cannot be answered from an extension table. `tar.gz` is two extensions and one archive. `.xlsx` is both an Office container and a book-shaped thing. `.ai` is either a PDF or an EPS. Anything that genuinely is a table lookup stays a row in `lists`.

- `PreviewType` gate per kind (`text_preview_enabled`, ...); off kind → layout no-size → hover dropped, in-flight replay dropped. `Document` kills Office procs (from tray, worker uncancellable); `Vector/Fonts` releases browser. `render_html` (`htm/html` only) switches Text draw → WebView page, kind stays `Text`.

The gate is the user's on/off switch for a kind, checked before any work happens. Turning a kind off has to unwind cleanly: the layout gets no size, so the hover is dropped and any already-running replay for that hover is thrown away rather than revealed. Turning `Document` off is special — an in-flight Office render cannot be cancelled, so the tray kills the started processes instead of asking politely. `render_html` is the one place the drawer changes without the kind changing: an `.htm` is still "text" to the user, but the Text drawer paints glyphs and cannot do layout, so the browser engine draws it instead.

- Measure off tick where felt: `measured_off_the_tick` keyed by (file,version,against); miss → spinner + replay (`VideoProbed`, `OfficeRenderReady`, box answers). `effective_preview_scale` / `scale_in_room`: single place for size; `preview_scale, video_scale, animated_scale, ebook_scale, document_scale, font_scale, design_scale, vector_scale, text_scale`.

Measuring a file means opening it, and opening it on the preview thread is exactly the stall this architecture exists to prevent. `measured_off_the_tick` is the escape hatch: the loader says "I do not know this one's size yet", paints a spinner in a guessed box, and hands the open to a worker thread. When the answer lands it arrives as `VideoProbed` or `OfficeRenderReady` and _replays_ the hover at the correct box. The scale functions all funnel through `effective_preview_scale` / `scale_in_room` so "what fraction of the screen may this kind take" is one calculation, not nine.

## Read guards (`cloud_files.rs`, `content_type`, `paths.rs`, budgets)

> What keeps a hover cheap and safe: never download cloud files, never allocate unbounded, never hold the global config lock across disk IO, never trust a header size. Both hook and loader ask the same gate so neither can reach what the other refused.

Every one of these is a guard against a file that lies. A size field can claim 400 GB, a cloud placeholder can be a reparse point that triggers a 2 GB download, a header can claim 65535×65535 pixels. The rule is that no allocation happens before the number that justifies it has been read, and no disk read happens without a bound.

- Ceiling: `decode_budget_gb` (1GB, `Performance → Decode Budget`) checked pre-alloc incl. inflate; `catch_unwind` → miss; huge alloc = abort, so budget first (`image` limits + 3 canvas allocators + GIF/WebP/WIC).

The failure mode this prevents is not an exception, it is a process abort. A Rust allocation over the address space the OS will grant terminates the process before any `Result` exists, so `catch_unwind` cannot help and the budget check must happen _before_ the allocation, not after it. "Incl. inflate" means the compressed size is not the size: a 64 MB `.svgz` can inflate to 1 GB, and that number has to be projected before the inflate starts.

- `cloud_files`: attr-only check (no open/download) from hook's single dir-entry read (present/file/version/content-here). Same `HoverFacts` on preview thread: one `content_type::Probe` + `resolve` + `CONFIG` snapshot (drop `CONFIG` lock before IO).

The check is attributes only. Opening a OneDrive placeholder hydrates it — megabytes over the network, on a hook thread, while the user is moving the mouse — so the guard reads what the filesystem already knows (is this a file, does it have content here, what version) and never opens. `HoverFacts` is the struct that carries those answers across the thread boundary so the preview thread re-reads nothing.

- Images: header first for box; budget in `u64` (no `u32` wrap); no-dims → decode but don't cache. DDS: 148B header + first-mip/first-face only; short read = no preview; `dds_background` separate (`MediaType` own kind).

The header is read before the image so the spinner has the right box, and the budget is arithmetic in `u64` because a 32-bit multiply of width × height × 4 wraps below 4 GB and hands the allocator a plausible-looking number. A file with no dimensions at all is decoded but never cached — there is no key that would ever let the app notice the file changed. DDS gets a hard cap of the first mip and first face: a cubemap or a mip chain is hundreds of megabytes of data for a 256×256 hover.

- Archives: TOC only (zip central-dir, 7z header, RAR walk, tar 512B blocks); 20k entries, 300-char names, cancel flag; `.tar.gz` inflate capped 64MB → "stopped early".

A listing is read from the archive's index, never by extracting. Every container has a different index — zip's central directory at the end, 7z a header, RAR a chain of blocks, tar literally 512-byte headers back to back — and each is bounded independently: at most 20k entries, at most 300 characters of name each, and a cancel flag the user can trip.

- Office: 8B OOXML/OLE check only; cloud/zone-id/long-path → temp copy, read-only/hidden/no-MRU/close-unsaved; ask at hover, not rest.

Eight bytes distinguishes a zip-container OOXML file from a legacy OLE compound one. Everything after that is an engine. When the file is remote, carries a Mark-of-the-Web zone identifier, or has a path Office will not open directly, it is copied to a temp file first and opened read-only and hidden so it does not appear in Office's recent-files list, with unsaved-changes prompts suppressed. The check happens at hover time, not at rest time, so a folder of remote documents costs nothing until pointed at.

- UI Automation: `CUIAutomation8` + client timeout (fallback legacy unbounded); timeout = miss.

**UIA** (UI Automation) is Microsoft's cross-process accessibility tree — the only supported way to ask Explorer what is under a point without injecting code into Explorer. `CUIAutomation8` is its COM interface. It is asked with a client-side timeout because a hung accessibility provider would otherwise hang the hook thread forever; the pre-v8 fallback interface has no way to bound the call, which is why it is second choice, not first.

- SVG: whole-read under budget (incl. `.svgz` inflate); fonts: whole-read, parse `cmap`+`name` only; WOFF per-table inflate, WOFF2 Brotli to last-needed table; `.ttc` → copy `ttc_face` (1-10) beside profile (`file,version,face`).

An SVG cannot be header-measured usefully and is small enough to read whole. A font is read whole but parsed for two tables only: `cmap`, which says which Unicode codepoints map to which glyph, and `name`, which is the human-readable family name. WOFF stores each table deflated separately so a single table can be inflated alone; WOFF2 compresses the whole file with Brotli, which cannot be partially decompressed, so the read stops at the last table the preview needs. A `.ttc` holds several faces and must be copied into the WebView2 profile directory as a single-face file, named by file, version, and face index.

- HTML (`render_html`): app reads nothing, budget N/A; browser sandbox `allow-same-origin allow-scripts`, no pops/forms/top-nav/autoplay/file-access/fullscreen; `BROWSER_ARGUMENTS` no host; `plain_path` + `file:///C:/…` URL versioned by mtime.

For an HTML page this app reads nothing at all and delegates the sandbox decision to the browser engine. The two permissions it does grant are the minimum a page needs to run, and everything else a page might reach for is denied. `BROWSER_ARGUMENTS` lets the user add flags but is filtered so a `--host-rules` cannot turn the browser into a general-purpose proxy.

- SVG/font pages: `pointer-events:none` + `draggable=false` (+`user-select:none` specimen); HTML keeps pointer. Pinned engine drag: `settle_pinned_engine_press/drag`, `PinDrag::delivered`, `carry_pin_drag_with_the_hand` + `PIN_DRAG_*`; poll `preview channel` via `Carry` + deadline (no re-entrant pump).

An SVG or font page is drawn by the browser engine but is _not_ a document — it must never eat the click that was meant to select the file behind it, so pointer events and drag are disabled on it. An engine-drawn page (an SVG, a specimen) inside a pin needs the opposite: the drag has to be forwarded to the window that owns it, because the pin's own window only sees a child HWND in the middle. That forwarding polls the preview channel rather than running a nested message loop, which would re-enter paint and deadlock.

## Caches

> What is remembered between hovers so the second look is free, and what key makes it valid. Frames are keyed by file version + pixel size (an edit or resize re-decodes); probes remember misses too so an unplayable file costs one probe, not one per hover.

Caching a hover preview is easy and caching it _correctly_ is the whole problem. The rule throughout: the key must change whenever the output would change. A file's bytes can change (version), the preview box can change (window resize, monitor change), the config can change (scale, backdrop, gate). Any key missing one of those will show the user a stale picture and be indistinguishable from a bug.

- Frames: `image_cache_mb` (64MB/2048/`0`=hold-none, live-read, LRU-counter trim, oversize uncached) keyed (file,version,px). Miss → decode.

The frame cache is memory-only and bounded by `image_cache_mb`, a count ceiling of 2048 entries, and an LRU counter rather than a timestamp. `0` is a real setting, not "unlimited": it means hold nothing and always re-read. An image whose bitmap exceeds the whole budget is decoded for the current hover and then not stored, because caching it would evict everything else to hold one thing.

- Disk: `document_cache` `%TEMP%\rust-hover-preview\document\<key>.{pdf,png,none}` (`document_cache_mb` 256MB, LRU-read table, `hover_ended` at 0); magick `%TEMP%\...\image` (`image_disk_cache_mb`); separate folders so budgets survive restart. Text frames `text_cache_mb` (0 default) keyed (file,ver,opts,box,DPI,line); parsed doc keyed (path,mtime,len,theme,md-mode) + per-doc `styled_window` + `thread_local ParseState` (syntect `!Send`).

Disk caches outlive the process, so each keeps an LRU table it reads on startup. The `.none` extension is the point of the scheme: a file an engine refused to render is itself a cached answer, so the next hover costs a file-existence check instead of a process launch. Text frames additionally key on DPI and line because the same file rendered at a different scale or a different width is a different frame; the parsed (highlighted) document does not key on those. `syntect` is `!Send`, so its parse state cannot move between threads and is therefore `thread_local`.

- Probes remembered: `ProbedGeometry` (incl. miss), `video_player::plays`, `audio_track::{Probed,Track,gain}`, pdf size, svg size, font parse, archive listing, `.none`/refused markers (office 2min, engine-refused run-long). WebView fail 1min; browser `note_unanswered`.

The rule is that a _negative_ answer is cached as carefully as a positive one. `ProbedGeometry` has `Measured` and `Unmeasurable` variants precisely so that "this file cannot be measured" is a memo, not a repeated probe. An unplayable video then costs one `ffprobe` for the life of the app rather than one per hover. Refusal markers expire on different clocks on purpose: a refused Office document may be fixed by the user installing something, so 2 minutes; an engine that refused on its own terms is remembered for the whole run.

- Temp sweep: `%LOCALAPPDATA%\...\update` cleanup; webview per-run profile `%LOCALAPPDATA%\...\webview\<pid>` + startup sweep; `engines\<pid>.state` + job `KILL_ON_JOB_CLOSE` (see Started Processes).

Everything under `%LOCALAPPDATA%` is swept at startup because it is keyed by pid and a crashed run leaves its directory behind forever. A **job object** is a Windows kernel object that owns a set of processes: when its last handle closes, every process in the set is killed. That is what makes `KILL_ON_JOB_CLOSE` the crash-safe half of engine cleanup — this process can die without running a line of cleanup code and the OS still reaps the children.

## Preview window (`ui/preview_window/`)

> The only visible surface: a channel-driven layered window. Hook posts `Show`/`Hide`, slow work replays the same hover when done; layout picks box + display, paint composites the frame. Hide and show race via a generation count so a late frame never resurrects a departed hover.

Three concerns, three files, one window. `layout` decides _where_ and _how big_; `paint` decides _what pixels_; the event loop owns the HWND and the message queue. None of them talks to Explorer, and none of them knows what a video is.

The hard problem is the race. The pointer leaves item A and arrives at item B. The hover thread hides, then shows; meanwhile a probe of A finishes and posts a frame. Naively, A's frame paints over B's preview — a flash of the wrong file, which reads as the app being broken. `HIDDEN_EPOCH` is a counter under a mutex, incremented on every hide. Anything holding a computed answer carries the epoch it was computed under; before revealing, it re-reads the counter, and a mismatch means the user has moved on and the answer is dropped.

- Msgs: `PreviewMessage::{Show,Hide,...}`, `OfficeRenderReady`, `VideoProbed`; `PreviewMessage::PinUpdate` = explorer half (loop no-op). `hide_preview` (hook thread, `HIDDEN_EPOCH` count under lock) vs show (preview thread); newest-wins drain; `ShowWindowAsync`; stall `PREVIEW_STALL_MS` → hook ends warm engines. Reveal gated by `publish_pointer_item_box` + `read_failure_is_the_same_item`.

`newest-wins` drain means the queue is emptied and only the last message acted on — the intermediate hovers between two fast mouse moves are not worth a window operation each. `ShowWindowAsync` is used instead of `ShowWindow` because the call may come from the hook thread. `PREVIEW_STALL_MS` is the backstop for the pathological case: if the preview thread has said nothing for three seconds, the hook concludes the window is wedged and ends the warm engines rather than leaving them running forever.

- Layout: `layout.rs` + `dimensions.rs` + `preview_scale.rs` + `displays.rs::display_at` (`Desktop` real / recorder). `compute_mouse_layout` (Follow=roomiest quadrant; Best=`centered_top` beside cursor/focused item); `waiting_placement/flush_at_cursor` (spinner `WAITING_BOX`/arc box at pointer corner, 1px off); `AvoidMode/AvoidRegion` (Filename default, FilenameColumn, Details; shortest clearing step, sides preferred for columns, mouse held to pointer); keyboard anchors on item tail/middle; text re-measure at fitted width (`text_box_room`, `text_scroll_far_edge_grace_pixels`=40).

`displays.rs` is a seam: `Desktop` enumerates real monitors, and under the recorder it enumerates the recorder's virtual monitor instead, so a screen capture of the app looks correct. `Follow` places the preview in whichever quadrant has the most room; `Best` puts it beside the cursor or the focused item, which is what makes keyboard navigation usable. `AvoidRegion` is the setting that keeps the preview from covering the file name it is previewing: `Filename` clears just the drawn text, `FilenameColumn` clears the whole column as the view reports it, `Details` clears every column. `flush_at_cursor` offsets the placeholder box one pixel from the pointer so the cursor never lands inside the spinner.

- Paint: `paint.rs`, `surfaces.rs`, `pixels.rs`, `loading_paint.rs` (`paint_pin_spinner`, `PIN_ARC`); `WM_DPICHANGED` drops frame → replay; paint-before-reveal; spinner only past `spinner_delay_ms` (250ms); pin wait in media band.

Paint-before-reveal is the rule that prevents a flash of empty window: the frame is composited into the layered surface _first_, and only then does the window become visible. `WM_DPICHANGED` means the monitor's DPI changed (someone dragged the window to a different-scaled display, or changed a setting); the cached frame is now the wrong size, so it is thrown away and replayed at the new scale.

- Anims: `animated.rs` (`animated_scale` vs `preview_scale` via `image_is_animated`); `StreamedFrames→Vec<ImageFrame>` (GIF/WebP/APNG/JXL-live; `heif_sequence` wired but `None` → still fallback); `MIN_ANIMATION_FRAME_DELAY_MS`.

An animated image is not scaled by `preview_scale` — it uses `animated_scale`, because a multi-frame source held to a single preview's size rules is measured differently and would otherwise blow up mid-playback. Animation frames are accumulated into a `Vec<ImageFrame>` and only advance after `MIN_ANIMATION_FRAME_DELAY_MS`, so a source that claims 1 ms per frame does not spin the preview thread.

- Tone map (`tone_map.rs`): `hdr_tone_map=reinhard|aces|srgb|off`, `hdr_exposure`, at `load_static_image` top + `dds_image` per-sample; alpha untouched. Converts scene light to display range instead of clipping.

Tone mapping is the step between "what the file stores" and "what a monitor can show". A raw DDS or HDR image stores light values well above 1.0; without tone mapping those pixels simply clip to white and the highlight detail is gone. `reinhard`, `aces`, and `srgb` are three different curves for compressing that range; exposure is applied before it. Alpha is never touched — the alpha channel is not scene light.

- Codecs (`formats/codecs.rs::Row`, `wic_image`): WIC for HEIF/AVIF/JXL-still/WebP-still/DDS-first; ask at planned box; miss→`None`. `Codecs` submenu: MFTEnumEx (video/audio), WIC codecs+MIME (HEIF=container+decoder), ProgID (Office/Libre), loader version (WebView2), WinRT registration (PDF). FFmpeg presence cached per menu build (tray clears).

**WIC** (Windows Imaging Component) is the OS's built-in image decoder, and using it costs no process and no code. It handles HEIF, AVIF, still JXL, still WebP, and first-frame DDS, but it is asked _at the planned box size_ — WIC decodes to whatever dimensions you ask for, so asking early means decoding at full resolution and throwing most of it away. The tray's `Codecs` submenu enumerates what is actually installed (MFT = Media Foundation Transform, for video and audio decoders; MIME registrations for WIC; ProgID = the registered class id for an application; WinRT registration for the PDF API). Rows are presence-checked when the menu is built and the cache is dropped when the tray menu is next built, so installing FFmpeg mid-session is picked up by reopening the menu.

## Explorer hook (`shell/explorer_hook/`)

> How the hovered file is found. Resolves the item under the cursor (or keyboard focus) to a real path via the Shell view at that position — never by filename lookup or filesystem walk — then gates and notifies the preview thread. Throttles probes by user activity and display state.

This is the part that most needs explaining, because the obvious implementation is wrong. You cannot find a hovered file by listing a directory and matching names — Explorer shows search results, virtual folders, shell namespace items, and un-downloaded cloud placeholders, none of which correspond to a path. The only correct source is the Shell view object at the screen position: ask it "what is at this point", get the item, ask the item for its filesystem path.

Flow: detect window+item → wheel-as-input → nav keys → entry+`plain_path`+gate → `Show/Hide` → resolve-by-position (no fs lookup) → keyboard/mouse precedence.
The order matters. Keyboard input is only consulted after the mouse has been ruled out, because arrow-key navigation must not steal the wheel or mouse behaviour of a folder the user is browsing with the pointer.

- Item: one batched UIA read (box, name, `ItemIndex`, acc value); point→row walk-up; keyboard: focused or selected-of-list + child text boxes (for `Avoid`). `view_bounds` clip; metadata line never resolves.

The UIA read is batched into one call for four values, because four separate calls cost four round-trips into Explorer. A point maps to a row by walking up to the nearest ancestor that reports an `ItemIndex`. `view_bounds` clips so a point outside the list resolves to nothing at all, and the metadata line at the bottom of a details view is excluded — it reports text for the whole selection and would resolve to the wrong file.

- View: match Shell windows by frame HWND; tabs = multi views per frame; showing tab = window-under-pointer/provider window; chrome-pointer → per-view ask + agreement rule; cache per window, drop on loc change. `IFolderView2::GetItem→SIGDN_FILESYSPATH`; claim-then-file; match-with-no-file ≠ fail; pointer double-look agreement; per-(item,window) memo.

One Explorer tabbed window can hold several folder views, so the frame HWND is only half the answer — the view is matched by walking up from the item's own window. `SIGDN_FILESYSPATH` is the shell property that yields the real path, and it can come back empty for shell namespace items that have no file; that is a _match without a file_, which is not a failure, it just means there is nothing to preview. "Agreement rule" and "double-look" are the same safety property asked twice: when the pointer is on Explorer chrome rather than inside a view, two independent asks must both agree before the app acts, so a stale view answer cannot hijack a hover.

- Place: folder+URL facts (`HoverLocation`, `location_fact_key`, canonicalized); missing fact = retry, not change.

The location facts are what let the preview know _where_ it is without re-deriving it. They are canonicalized first so two spellings of one path produce one key. A missing fact is treated as "try again" rather than "the location changed" — the alternative would make every hover of a slow folder look like a folder switch.

- Pace: `ExplorerState` (focus/visible/minimized/absent → `tick_ms`/150/500/1000ms); pin forces `ActiveFocus`; `app::afk`; keyboard poll in active path only (`focus_move_input/NavigationInput`, nav+type-ahead w/o Ctrl/Alt). Tray idle-waits; `ensure_noactivate_monitor` event+timeout.

The poll rate is the whole cost model: 150 ms when Explorer is focused and visible, up to 1000 ms when it is not, and zero work when absent. A pin forces `ActiveFocus` because the pin is a topmost window over Explorer and would otherwise make the app think the user had left. Keyboard polling runs _only_ on the active path, and only for navigation and type-ahead keys — a hook that watched for Ctrl/Alt would fight Explorer's own shortcuts.

- Restart: `TaskbarCreated` → rebuild collection, drop views/caches, hide, re-arm gate; slow-retry missing collection.

Explorer restarts (crash, update, sign-out) and every COM pointer into its windows becomes invalid. `TaskbarCreated` is the broadcast Explorer sends when it comes back. Nothing can be repaired incrementally: the window collection is rebuilt from `EnumWindows`, all view caches are dropped, the preview hides, and the gate re-arms. If the collection is not ready yet the hook retries slowly rather than declaring Explorer gone.

- Wheel: over Explorer or covering preview → reopen latch + restart hover clock; ends keyboard ownership. Text wheel: hook reads `mouseData` delta, swallows inside preview region only.

The wheel doubles as a hover signal: turning the wheel means the pointer is over the window the user cares about, so a preview that was suppressed reopens. It also hands control back from the keyboard to the mouse. The hook reads the wheel delta from `mouseData` directly rather than installing a second low-level hook, and swallows the event only when the pointer is over the preview's own region.

- Folder-change: post-press active cadence; suspension lifts on settle; box-change = gone (`read_failure_is_the_same_item`); programmatic change waits for input; auto-focus never previews.

The last item is the subtle one. When Explorer changes folders _by itself_ — a refresh, an auto-arrange — the app waits for the user to do something rather than previewing whatever is under a pointer that is no longer meaningful. `read_failure_is_the_same_item` distinguishes "this read failed because the item is gone" from "this read failed transiently", which is what stops a deleting file from being retried forever.

## Pin mode (`ui/preview_window/pin_*.rs`, `pin_window.rs`, `pin_chrome/`, `text/`)

> How a hover becomes a persistent topmost window. Pin key freezes hover machinery, grows the same window with caption/transport chrome, and swaps media in place for picks or prev/next walks — same surface, new file, no re-resolve.

The design decision that shapes all of this: the pin is _the preview window_, not a second window. It does not create an HWND, and it does not re-run the hover path. It takes the window that is already there and changes its frame, its height, and which code path fills it. Everything that would require re-resolving a file is therefore free — the pin already knows the folder, the window, the sort order, and the config.

- Take-up: `pin_key` (Space) on settled preview (`pin_screen_is_settled`), never spinner; engine kinds install `pinned_engine_media`. `key_input` LL hook + per-tick counter; `refresh` on config/tray.

A pin can only be taken from a _settled_ preview. Pinning a spinner would mean pinning an unknown box and an unknown kind, and both are needed to lay the pin out. Space is the default pin key because it is the one key that means "hold this" everywhere, but it is configurable, and the low-level key hook only counts presses — the tick drains the count, so a key repeat cannot pin twice.

- State: `pin_window::{PinState::Down/Up(PinUp)/Ending(Reason), PinExit, end_pin, Closed/Asked/SwitchedOff/MediaGone/Hung}`; `PinUp` = window+media+`Option<PinKeyboard>`+cmd queue. One ender; watchdog kills post `Hung` only via loop msg. Focus: `WS_EX_NOACTIVATE` toggled (`pin_set_focusable`, `pin_take_focus` on `WM_LBUTTONDOWN`, release on `WA_INACTIVE`); authority = `GetFocus` (`pin_is_focused`).

`PinState` is a small state machine with exactly one exit path. Everything that can end a pin — the user closing it, asking to end it, a config change turning the mode off, the media disappearing, or a hang — funnels into `PinExit` / `end_pin`, so there is no path where the window survives its own state. `Hung` is the state after which the watchdog is allowed to kill, and only via a message on the loop, because killing from the watchdog thread would leave the loop painting a dead window. Focus is authority-checked with `GetFocus` rather than remembered, since Windows can move focus behind this process's back.

- Isolation: `show_preview/show_preview_keyboard` + hover-msg gate refuse while up; `hide_preview` acts for pin; loop repaints pin band/chrome only. Hook sleeps past trigger gate (except `Update Preview`).

While a pin is up, the entire hover path refuses to do anything, and this is enforced at a gate rather than at each caller. `hide_preview` does still work — it is how the pin closes. The hook thread then goes to sleep past its trigger gate so it is not polling Explorer at all, and the only thing that can wake it is `Update Preview`.

- Follow/swap: `Update Preview` (pick: click+focus-agree+place `press_is_a_listing/click_is_over_a_listing/get_file_under_cursor`, or `follow_selection` when hover off); bubble holds pick (`pin_bubble_pick`). Loop: `pin_media_load` (thread, `PreviewScale::FitToScreen`) → `swap_pinned_media`; `PinnedPreview::bound` (no walk-down); `PinLoad` + `PIN_ARC` spinner over current media; engine docs shown as-asked (`PLACED/PLACE_AS…`, `box_change`; move ≠ `show`).

`Update Preview` is the pin's version of hover: either you click a file in the folder view behind the pin, or — when the hover is off and there is nothing to click — it follows Explorer's selection. Both paths run the same resolve, so the pin never guesses. A pick needs the click, the focus, and the placement to all agree before it is accepted, which is what stops a click on the pin's own caption from being read as a click on the file behind it. The load runs on a thread and the old media stays visible with a spinner drawn over it, so a swap is never a blank flash.

- Walk: `pin_navigation` (`read_dir` flat, `kind_of_name`, `DirEntry::file_type`, no `metadata` except Date/Size walks, per-kind reader check, capped map) + `IFolderView2::GetSortColumns/sort_key_of` else natural order, wrap; `PinStep/pin_walk/pin_walk_of_current/pin_step_off` (skip bad, forward-only, `show_pin_failure→MediaType::Unplayable/PIN_FAILURE_SIDE`); `step_pinned_file` copies config (no lock across IO); arrows only when pin focused (no `GetAsyncKeyState` steal); `pin_command_request/ask_pin` single door (`navigating` flag).

The walk has to reproduce the order the user _sees_, not alphabetical order, so it asks Explorer's own view for its sort columns via `IFolderView2::GetSortColumns` and falls back to natural order only when the view will not say. The file list is read once, flat, and classified by name only — no `metadata` call, because stat-ing hundreds of siblings is the one thing that would make the next-file button feel slow. Files this app cannot read are skipped rather than fatal, and the walk is bounded by a budget of "every file but the one we are on", so a folder of nothing readable ends the walk instead of spinning. `PinStep` carries `at` (where the walk landed) separately from `from` (what the press was made on) because the planner runs on another thread, and a late answer to a press about a file the pin no longer shows must be recognised as such. Arrows are honoured only when the pin itself has focus, and never via `GetAsyncKeyState`, which would steal keys from whatever window actually owns them.

- Chrome: same window grows (caption up, transport down; overlay kinds draw over). `pin_chrome` GDI strokes; window btns fixed slots, pin 4 (prev/next/open/open-with-list) left-packed (`button_boxes`); tooltips (`PinTooltip/PIN_TOOLTIP_DELAY`, `AssocQueryStringW` once); `Open With`=`ShellExecuteW+plain_path` then `request_pin_end` (=`WM_CLOSE` road); `Open With...`=`rundll32 shell32.dll,OpenAs_RunDLL` raw tail, same end; no launch polling.

The window is one rectangle with fixed button slots; a kind that draws to the edge (a video, an SVG) draws _over_ the chrome instead of being inset around it, so there is no second window to keep in sync. `Open With` deliberately ends the pin: it hands the file to another application, and keeping a preview up over a file the user has just opened in Word is not useful. It ends by asking for a close rather than closing directly, so there is exactly one path that tears a pin down.

- Resize/drag: `PinFrame` (shape scales / text reflows / sound+ffplay unframed); 8 bands (corners wider, bands before caption); drag = `SetWindowPos` only; resize stretches opaque (GDI `StretchBlt/HALFTONE`, alpha rewrite) or samples alpha over backdrop. Minimize → bubble (`carry_bubble_drag`, at minimize-btn box, drag by grab-offset, restore at placement mode). Overlay chrome auto-hide (`PIN_CHROME_ARRIVAL_SECONDS`, cursor-proximity, no fade).

`PinFrame` is a band model: eight hit-test bands around the edge, corners wider than the sides so they are easier to grab, and the caption band excluded because the caption already has buttons in it. Dragging is a `SetWindowPos` and nothing else — no tracking loop, no re-layout. Resizing a frame _shape_ (a picture) is a stretch; a frame that carries _text_ reflows instead of stretching, because stretched glyphs are unreadable. Minimizing does not hide: the pin collapses into a small bubble at the position of its own minimize button, and the bubble carries its grab offset so it does not jump when dragged.

- Transport (`pin_playback.rs`, `transport.rs`, `audio_preview/card,page`): `PinTransport{len,pos,stopped}`; native = direct (`SetCurrentTime`, pre-header queue, `SetLoop`); ffplay = clock-over-start + relaunch-at-second (`retire_replaced_player`, `end_retired_player`, `VIDEO_START_WAIT_SECS`), keys via `PostMessageW(VIDEO_HWND,WM_KEYDOWN/UP,'P'/Space,0x00190001)` (scan code, no repeat); bar drag = seek; level at start, `0%`=no player; `remember_audio/video_volume` write on release; sound card = own controls (`control_boxes/control_at`, `BAR_REACH/GAP_PIXELS`), no caption (`pinned_caption_height`=0), `Space`=hold, bar=seek (`PIN_AUDIO_TOGGLE/SEEK`, `toggle/restart_pinned_audio`, `playing_player`); unknown len = block bar, seek refused.

Two playback engines answer these questions in incompatible ways. The Windows media engine reports where it is and seeks directly. FFmpeg's `ffplay` reports nothing at all — it is a window with no API — so the app keeps `PinTransport`: the length from the probe, and a position computed as _this app's clock minus the start time plus the offset it launched at_. Pausing an ffplay means sending it a key and remembering the second; seeking means retiring that player and starting a new one at the new second. Keys go in via `PostMessageW` with an explicit lParam carrying the scan code, and without the repeat flag, because a key-repeat of Space would rapidly toggle.

- Sound start (`Volume→Audio Seek`): beginning/middle/anywhere/memory (`%TEMP%\...\audio\seek.txt`, 1GB cap, `audio_seek::flush`); native `SetCurrentTime` on first ready tick, ffplay `-ss` at launch; loop: native `SetLoop`, ffplay `wrap_audio_player` (2nd pass full-file loop, `AUDIO_CARD_DIRTY` turnover, no backward bar). Normalize: `ebur128` → gain to −14 LUFS vs −1 dBTP (`audio_track::gain`, `Some(1.0)`≠`None`; `AUDIO_GAIN_TIMEOUT_SECS`); per-player gain path.

"Start at" has four meanings and the setting distinguishes them: the beginning, the middle, wherever the pointer was in the file, or where you were last time (remembered in a capped temp file). A native player can seek once it is ready; ffplay has to be told with `-ss` on its command line or it plays from zero. For looping, the native player has a mode flag, and ffplay does not — the app wraps it with a second pass over the whole file. Normalisation is loudness, not volume: `ebur128` measures integrated loudness in LUFS and true peak in dBTP, and the file is adjusted to −14 LUFS without exceeding −1 dBTP, which is a broadcast-style target that stops one quiet video and one loud video from being unusable next to each other. `Some(1.0)` is deliberately not the same as `None`: 1.0 is "measured, needs no change", `None` is "not measured".

- Text pin = full mode (`current_text_options`): scroll/select/clipboard (`Ctrl+C/A`, menu), scrollbar (`bar_row` shared geometry), wrap-to-width + pull-back + end rule, `text_cache_mb` frames, `Refresh` rebuilds from recorded hover; journey+preview regions published per tick.

A pinned text file is a full text preview, not a picture of one: it scrolls, selects, copies, and wraps to the pin's current width, and the frames it re-renders are cached. `Refresh` after a config change works by replaying the _recorded_ hover rather than asking Explorer again, because the pin is the only thing that knows what it is showing.

## Kind notes (reader → engine, gate, scale, backdrop)

> Per-format specifics sharing one pattern: what draws it, which tray gate/scale/backdrop applies, and the one quirk that decides correctness (printer for Excel, flat-cover skip for books, first-plate rule for comics, dual-play alpha for metafiles).

Every kind answers the same six questions: who draws it, which tray gate turns it on, which scale function sizes it, what it composites over, what cache it uses, and what one quirk will make it wrong if ignored. Read this section as a checklist rather than a list of trivia.

- Picture/WIC/DDS/ToneMap: above. `image_background`, `preview_scale`.
  The default kind, and the one everything else falls back to. It has no engine, no probe, and no quirk beyond the header-size trust described under Read guards.

- Design (`psd_image/project_image/eps_image/cdr_image`): never composite; shipped preview (`mergedimage.png`, ORaster `Thumbnails/`, Procreate `QuickLook/Thumbnail.png`, CDR `metadata/thumbnails|previews/`, RIFF chunks; PSD planar+box-filter stream, 4 compressions, no Lab/multichannel/no-compat); `.ai`=PDF if `%PDF-` else EPS; `design_scale/fit`, `design_background`, `design_extensions`, `Design` gate.

A design file's own preview is read from inside the file rather than composed by this app, because Photoshop documents have no meaningful flattened form without the layer stack. The readers read a shipped thumbnail where one exists and otherwise stream the flattened composite — PSD is planar, so it is box-filtered on the fly, and three of the four compressions are supported. `.ai` is genuinely two formats in one extension: a PDF if the first bytes say so, an EPS otherwise.

- Libre (`libreoffice_render`): `[libre]` + Office-no-install fallback; `page_engine` (tray `Engine→Select Engine→Office` + `app_installed` ProgID); headless w/ own doc, `--accept/--invisible` = no-hold; 1.2s cold/0.2s warm; TTL/`warm`; give-up by name+id; `swf` excluded (video). Shares `Document` gate/scale/cache w/ Office.

The second spelling of the same job as the Office engine, for machines with LibreOffice instead of Office. The user picks one in the tray, and the choice is a ProgID registration check, not a path search. LibreOffice is driven headlessly with its own throwaway user profile so a running instance is never disturbed; `--accept` / `--invisible` mean the converter will not block waiting for a dialog. Cold start is a second and a half, warm is a fifth of that, which is why it is kept warm under a TTL.

- Magick (`imagemagick_render`): `[magick]` second spellings+raws; `-auto-orient` (only rotation applied anywhere); box=room ceiling, 8bpc; 1st picture only; raw-dump size via filesize÷px-weight proportions (`-size/-depth`); stdout→frame+disk page; refused remembered run-long; no TTL. `Images` gate.

ImageMagick is here for the camera formats no codec library claims. Two constraints drive everything: the box is the _room ceiling_, because the size is unknown until ImageMagick says, and there is no TTL, because the process is cheap and warm enough to keep. For a raw dump the dimensions are not in the file, so they are computed from the file size and the bytes-per-pixel implied by the requested `-depth` — a proportion, not a fact, which is why the preview may be right in shape and approximate in scale.

- PeaZip (`peazip_render`): `[peazip]` no-reader formats; `Backend::of` (7z console/FreeArc/zpaq/zstd; `.br/.bcm/.lpaq8` name-only); own thread, 1 queued, `CREATE_NO_WINDOW`+null stdin, adopted-not-recorded, give-up→unlistable; same `Listing`+cache. `Archives` gate.

For archive formats nothing in-process can open. `Backend::of` picks which PEaZip executable handles a given extension, and for the most obscure formats it can only guess from the name. One process is queued at a time and started with `CREATE_NO_WINDOW` and a null stdin, because these console tools will otherwise flash a window and wait for a keypress.

- Calibre (`calibre_render`): `[calibre]` azw/mobi/epub/fb2/djvu/lrf/pml/snb/tcr/htmlz; `ebook-convert` → PDF → first-non-flat of 8 (`book_page`); keyed (book,ver,engine); 2× render bound + hover cap; recorded, no TTL. `Ebook` gate, `ebook_scale`, `image_background`.

Calibre converts a book to PDF, and this app then renders the PDF. The one correctness rule is `book_page`: a converted book's pages are not all real content — the cover is a flat image, and several other pages are boilerplate — so the app walks the eight candidates and takes the first that is not flat. Without that a book opens to a blank or a copyright page.

- Comic (`comic_preview`): cbz/cbr/cbc via `[ebook]` + `page_spelling`; headers-only walk, `entry_bytes` single plate (`jpg/jpeg/jpe/jfif/png/gif/bmp/tif/tiff`, skip `__MACOSX`/meta), natural-numeric first, no spread cut, size off-tick, size-only cache; WebP/AVIF plates = none. Same `ebook_scale`/bg/gate.

A comic archive is a zip of images, and it goes through the book gate because it is the same user mental model. The rule that decides correctness is that panels are not pages: the first image in natural numeric order is shown whole, and a two-page spread is never cut in half.

- PDF (`pdf_preview`): `Windows.Data.Pdf` page 1, MTA threads; `%PDF-` in 1KB; `SHCreateStreamOnFileEx` + `CreateRandomAccessStreamOverStream` (verbatim-safe); measure-then-render-at-box (`ebook_scale`, ≥100%=fit); BMP encoder, white-opaque; size-only cache.

The OS PDF API needs a WinRT stream, and WinRT will not open a `\\?\` path — hence `SHCreateStreamOnFileEx` wrapped and converted, which is the "verbatim-safe" note. `MTA` here is the opposite of STA: WinRT streams are free-threaded, so these readers must _not_ be STA. Only page one is rendered, and the render is asked for at the final box size so the bitmap is never oversized.

- Vector (`metafile_image` + `svg_preview`): `Vector` gate/scale/backdrop (`vector_background`; checkerboard in-page). WMF/EMF `PlayEnhMetaFile` into own DIB + dual black/white coverage recovery; EPS preview-only (WIN header or `%BeginPreview`, budget-checked; 8b-index+alpha here). SVG: `svg_preview` (doc?+size via width/height→viewBox→300×150, `.svgz` inflate, verdict cached) → WebView2 image-page (`NO_INTERACTION_STYLE`, mtime URL); `vector_scale`.

A **DIB** (device-independent bitmap) is a plain array of pixels in memory that a program can fill in directly, which is why it is the target here: GDI metafiles are draw-command streams, so `PlayEnhMetaFile` executes them into a memory device context holding one of this app's own DIBs. That produces a bitmap with no alpha, which is why the same drawing is played a second time in white to recover coverage and get a clean edge. EPS is preview-only — this app never implements a PostScript interpreter, it reads the preview bitmap Adobe embedded.

- WebView2 (`webview_preview/{engine,host,environment,pages}`): newest-`want` only (wake, drop supplanted nav); `place` moves; `cursor_preview_hover` close rule; `file://` via `plain_path`; per-run profile; `Persistent↔webview_idle`=10min else AFK; `NAVIGATION/HOST_CREATION_TIMEOUT`, 4× controller retries, `note_unanswered`; `RHP_WEBVIEW_TRACE/RHP_WEBVIEW_PROFILE`. Draws SVG, font specimens, HTML pages.

WebView2 is a whole Chromium instance embedded for the three things GDI and WIC cannot do: SVG layout, font glyph rendering, and HTML. It is the most expensive thing in the app, which is why it is newest-wins — a second navigation request supersedes the first and the first is dropped rather than queued. It runs on a per-run profile directory (see Temp sweep) so no state carries between runs, and a `NAVIGATION` or `HOST_CREATION_TIMEOUT` that expires is recorded via `note_unanswered` so the same dead document does not time out on every hover.

- Font (`font_preview/{font_tables,specimen}`): WebView2 `@font-face` page, `cmap`-gated lines (all-or-nothing + own-chars fallback), `name` (en pref), viewport units, per-bg ink, `direction` per line; `font_scale` (½ room), `font_background`, `[font]` list, own kind slot.

The app reads the font's tables; the browser renders the glyphs. The specimen only shows lines for characters the font actually has — gated on `cmap`, all-or-nothing, because a specimen where half the alphabet is tofu boxes is worse than one showing the font's own characters. `font_scale` gives fonts half the room an image would get, and the ink colour is chosen against the kind's backdrop rather than being fixed.

- Text (`text_preview/*`, `text_paint`, `text_theme`, `text_formats`): `config.ini` lists (ext + namelist for LICENSE/Makefile/.gitignore); measure→`render` at DPI (`GetDpiForMonitor`), `effective_preview_scale`=100% + `text_box_room(text_scale)` + re-measure-on-wrap + `TextMetrics(text_font_scale 0.25–16×)`; `PreviewType::Text` gates incl. `text_preview_enabled` in `is_text_file`; kind-change rebuilds that kind only; themes `theme/`+bundled (`custom:<name>`, menu lists, default fallback, drop on open); producers syntect/two-face + pulldown-cmark + RTF strip + CP437/ANSI NFO; GDI `ExtTextOutW` opaque DIB, alpha forced.

Text is measured at the display's real DPI and then rendered at it, because GDI text at the wrong DPI is visibly wrong and cannot be scaled afterwards. A change to one kind rebuilds only that kind's frames — recompositing everything else would be visible as a flash on formats that have nothing to do with the setting changed. Themes are loaded per menu build, not cached forever, so editing a `.tmTheme` file and reopening the tray menu is enough.

- Archive page (`archive_preview` + `archive_listing`): TOC→`Listing` (zip/7z/UnRAR/tar+flate2); never resolve/join/open; folders-first natural sort, collapse, Fluent/MDL2 glyphs; 100 rows + `… and N more`; unscrollable; cache (path,mtime,len); engine listings separate.

A listing is a card, not a browsable filesystem: it shows one hundred rows and says how many more there are, and it never extracts, resolves a symlink, or opens a child archive. Folders sort before files and the sort is natural (so `page2` precedes `page10`), with collapse arrows drawn from the Fluent/MDL2 glyph set.

- Office page (`office_preview`): render-tier-or-engine page; no embedded thumbnails; Excel manual-calc measured (still recalcs); `CopyPicture→PNG` when `printer_installed` false (retry once, clipboard restored, ≤900×700); `source_kind/page_is_workbook_picture` picks `document_scale`-fit (vector/PNG-export 640–1920px) vs bitmap-size; `WAITING_BOX` spinner; `OfficeRenderReady` upgrade replay (no pre-resize); 25s cap + 2min backoff; Word/Excel→PDF via `pdf_preview`, PPT slide-1 PNG; `Slides.Item(1)`; `ExportAsFixedFormat/Export/CopyPicture` to path.

The decision that matters: this app never uses a document's own embedded thumbnail, because a document can carry a stale thumbnail of a previous save. It renders the real content. Excel is the hard case — rendering a range requires a printer driver, and without one the only route is `CopyPicture`, which goes through the clipboard and is capped at 900×700, retried once, and always restores the clipboard afterwards. The `OfficeRenderReady` upgrade replay is the pattern worth noticing: the preview is first shown at a guessed box from a cheap measure, then _upgraded_ when the real render arrives, without any resize animation in between.

- Office engine (`office_render/{worker,engines,com}`): late `IDispatch` per family, 1 each, `office_engine_idle` (10min/`0`/indef); attach-aware (visible=user's: no hide/quit/kill, settings restored); docs RO/hidden/no-MRU/unsaved; uncancellable → give-up on next hover (kill started procs, gen++, abandon thread, conditional cleanup); quit+pump then kill-if-ours; `EVENT_OBJECT_SHOW` + sweep hides `Publishing…` bar (own procs only); newest-wins.

This engine drives the _user's own already-running_ Office, which is why it is the most careful code in the app. It is attach-aware: if the visible Word window is the user's own document, this app does not hide it, does not quit it, and does not kill the process — it restores whatever settings it changed and walks away. A COM call into Office cannot be cancelled, so if one hangs the only exit is to give up on the next hover: kill the processes _this app started_, bump a generation so stale answers are ignored, abandon the thread, and clean up conditionally.

- Video: `route_video/media_engine_plays`: ffplay present → all video; else `[video]`→engine iff `video_player::plays` (with `MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING`) else none; `[ffmpeg]`→ffplay only. Geometry `ffprobe`+cropdetect parallel, `VIDEO_PROBE_TIMEOUT_SECS`, `ProbedGeometry` (miss held), wait-box + replay (`replay_where_the_pointer_is`); `video_scale` of own px (`animated_scale/preview_scale` for GIF/WebP/PNG via `image_is_animated`); ffplay per-hover window-as-preview (layered hidden; `VideoStart/VIDEO_HWND/player_wait/kill_stray_video_process`, kill-only stop, PID-handle+create-time record, hook+preview-thread re-check, `pinned()` exempts pin player); engine frame-server (`OnVideoStreamTick/TransferVideoFrame`, `MFVideoNormalizedRect` crop, `GetReadyState`, `A8R8G8B8`, `start_video_playback`, `FirstFrameWait/HIDDEN_EPOCH`, `FIRST_FRAME_GIVE_UP→ffplay`, thread-owned teardown). `Videos` gate.

**Two players, chosen per config and per capability.** If `ffplay` is installed, `[ffmpeg]` sends everything to it and the media engine is not used at all. Otherwise `[video]` asks the Windows media engine whether it can actually decode the file — with video processing enabled, which is the difference between "it has a codec" and "it can show me frames" — and falls back to no preview rather than to a black rectangle. The geometry probe runs `ffprobe` and a crop detect in parallel because either alone can take longer than the user will wait. When the engine path is chosen, the engine is used as a _frame server_: frames arrive as `A8R8G8B8` bitmaps with a normalized crop rectangle and are composited into this app's own layered window, so the video respects the same avoid-region, reveal, and pin machinery as everything else. When `ffplay` is chosen, its own window _is_ the preview, hidden and layered, and the app's job is to start it, find its `VIDEO_HWND`, and kill it reliably — which is why the process record includes a handle and a create time rather than a bare PID.

- Audio (`audio_track`, `audio_preview`, `audio_seek`): card (name/format/rate/ch/volume) + engine; machine picks: native stream→PCM else `ffprobe`→ffplay (`plain_name` URL fix); `AUDIO_PROBE_TIMEOUT_SECS`; `Probed::{NotAsked,Track,None}` + `remember_audio_only` (video-probe hands `.mp4/.mka/.ogg`-sound over; router asks first); wait-box + replay + same-thread card measure; memo LRU (never current card; `NotAsked`+card = frozen `make_room`); `AUDIO_CARD_REPAINT` 4Hz + `AUDIO_NAME_REPAINT` 30Hz `NameScroll` ping-pong; `Volume→Video`(mute)/`Audio`(quiet) at start, `0%`=card-only; shared stop path; seek/loop/normalize per Pin section.

An audio preview is a painted _card_ — name, format, sample rate, channel count, volume — plus a player. If the OS can decode it natively, the stream is pulled to PCM; otherwise the same probe-and-ffplay path as video takes over. The card repaints at 4 Hz because volume changes need to be visible promptly; the name repaints at 30 Hz only while `NameScroll` is ping-ponging a too-long filename. `Volume` has three meanings distinguished at start: mute for a video with sound, quiet-but-audible for an audio file, and `0%` meaning _card only, no player started at all_.

## Config / tray / processes / build

> Persistence and packaging. App is the sole writer of `config.ini` (self-repairing + ordered); tray is the UI onto it; every child process is tracked by job object + disk record so a crash still cleans up; installers build from the `github` profile.

`config.ini` is the entire user-facing state, and the app treats rewriting it as routine rather than exceptional. The tray menu is not a UI over a fixed struct — it is generated from `SETTING_GROUPS`, so adding a setting makes it appear in the right place automatically and it is impossible for the menu and the file to disagree about what exists.

- `config.ini` (`%APPDATA%\...`): `[settings]` grouped by tray menu (`SETTING_GROUPS`, `headings_are_old`, `ordered_text` stable order, `; Ungrouped`); then alpha list sections (16 `lists` rows: built-ins, match rule, normalize, field, `before`). App sole writer; installer seeds nothing; `sync_startup_setting` + `repair_startup_entry/same_path` (exe+args parse, nojz resurrection). Sanitize+`differs/to_ini` rewrite (retired keys dropped: `svg_*`, `off_trigger_key`, `transparent_background`, pdf/office/libre splits → `ebook/document`); `DEFAULT_*` single source + `(Default)` marks + fallback reads (`sanitize_dds/html_background`). Lists: missing key→restore+rewrite; empty=none; edited kept+ordered; `repair_older_lists` (row `before`) adopts added formats (image/vector/design/video/libre/archive/calibre/peazip/magick; `[ffmpeg]` order-only); `save` drops unknown. Themes `%APPDATA%\...\theme\*.tmTheme` (`custom:<name>`, listed not read, re-read per menu).

Each list row records five things about itself — the built-ins it shipped with, how a name is matched, how it is normalised, which field it holds, and what it was `before` the user last edited it. That last one is what makes self-repair possible: on load, `repair_older_lists` compares the user's edited list against `before`, and any format the app added in a later version is inserted into the user's list rather than being silently absent, while `[ffmpeg]` is order-only because its order _is_ its meaning. A list that is present but empty means "none", which is different from a missing key, which means "restore the built-ins and rewrite the file". Retired keys are dropped rather than ignored, so an old install's file converges to the current shape on first save.

- Tray (`shell/tray/{menus,submenus,commands,ids,event_loop}`): icon/menu/exit/`TaskbarCreated`; rows: Preview Types, scales, Volume, Codecs (stateless, presence-checked), Engines+TTL+`Persistent`, Pin Mode, Placement, Text Preview, startup. Gate/type change rebuilds that kind (`Refresh` recomposites others, text rebuilds frame).

Changing a gate does not restart anything; it rebuilds the affected kind in place. `Refresh` — a recomposite for most kinds, a full frame rebuild for text — is the operation the tray calls, and it is why a settings change is instantly visible rather than applied at the next hover.

- Started Processes (`app/engine_processes.rs`): Office app + `msedgewebview2.exe` (parent=us, by pre/post walk) + `ffplay` + probes. Job `KILL_ON_JOB_CLOSE` (handle never closed) + `%LOCALAPPDATA%\...\engines\<pid>.state` + profile dirs; PID+image+start-time check (no bare-PID kill); startup sweep first; probes adopted-not-recorded; player recorded but exempt from tier wholesale ends.

The discipline here is _never kill by PID alone_: a bare PID can have been recycled between the record being written and the sweep running, and killing an unrelated process is unrecoverable. Every record carries the image name and the process start time, and a kill is refused unless all three match. The `KILL_ON_JOB_CLOSE` job is the belt to the disk record's braces — the handle is deliberately never closed while the app runs, so only a crash closes it and lets the OS reap everything in the set. Startup sweep runs first because whatever the last run left behind must be gone before this run registers anything of its own.

- Updates: see Runtime table. Checks at start + menu build, hourly mem-only throttle, signed-off installer handover.

An update is an installer, not a patch: the running copy is handed over via `/S /R`, which means silent install plus relaunch, so the check verifies the downloaded file is a PE image (`MZ`) of a plausible length before it is trusted at all.

- Build: MSVC + `rust-lld` (`.cargo/config.toml`); `build.rs` + `winres`; release kills running copy (src-watch; debug exempt); `cargo packager` NSIS from `packaging/nsis/installer.nsi` (silent uninstall, close running); `github` profile (`target/github`, `build-installers.ps1` pinned packager) vs `release` thin-LTO loop.

Two build profiles, two jobs. `release` is the fast inner loop: thin-LTO for quick incremental builds, no installer, and it kills the running copy first because the linker cannot overwrite an executable that is mapped. The `github` profile is the one that produces an artifact: a different target directory so it never collides with the inner loop, and `build-installers.ps1` driving `cargo packager` with a pinned version. NSIS handles the uninstall path, which has to close the running app before it can remove it.
