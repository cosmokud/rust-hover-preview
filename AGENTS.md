# AGENTS.md

## What this is

Windows 11 tray app. Hovering a file in Explorer previews it in-process or through an external engine; middle-click or Space pins that preview to a window of its own.

## Commands

| Purpose | Command | Note |
| --- | --- | --- |
| Build | `cargo build` | links the debug binary the tray runs from |
| Build to ship | `cargo build --release` | `build.rs::stop_running_app` runs `taskkill /F /IM rust-hover-preview.exe` before the link, release profile only |
| Check every target | `cargo check --all-targets` | ~5 s warm; `--all-targets` is what reaches the `src/**/tests` modules |
| Test everything | `cargo test` | 1667 passed, 38 ignored, ~10 s after compile |
| Test a slice | `cargo test <filter>` | `cargo test pin_zoom` = 50 tests in 0.1 s; `cargo test config` = 109 in 0.29 s |
| Path gate | `cargo test --test no_path_leaks` | every tracked file, and every commit message |
| Lint | `cargo clippy --all-targets -- -D warnings` | ~0.3 s warm; the gate is `Cargo.toml`'s `[lints.clippy] all = "deny"`, so a bare `cargo clippy` denies the same lints and the flag adds rustc's warnings |
| Format | `cargo fmt --all -- --check` | `--check` is the half that reports; without it rustfmt rewrites every file it reaches |
| The whole gate | `pwsh -File scripts/verify.ps1` | the real file, not a recipe; `.github/workflows/ci.yml` is gone, so the path gate, the formatter and the linter run from this script and `manual-ci.yml` alone; `-Fast` = check + the three `tests/` targets (~1 s), `-Pure` = the suite without `ui::`/`shell::`, `-Filter <name>` for one test, `-NoTest`; the default run names the desktop-dependent tests out with `--skip` |

## Run it

`cargo run` works. A second instance exits silently: `main` claims `Local\rust-hover-preview-single-instance` (`src/app/single_instance.rs::acquire`) before DPI, COM, config or any hook, and returns. No icon and no window means another copy is already in the tray.

A release build ends the copy you were testing. `build.rs` runs `taskkill /F /IM rust-hover-preview.exe` whenever `PROFILE` is `release`, before the crate is compiled and linked, and says so only through a `cargo:warning`. A debug build does not.

## Traps

- **A key `to_ini` writes and `apply_ini` never reads** (`src/config/config/ini_mapping.rs`). The save
  writes it, the next load loses it, and nothing goes red: `a_file_the_app_wrote_is_one_it_reads_back_as_itself`
  (`src/config/config/tests/persistence.rs`) round-trips `AppConfig::default()` with a short override list,
  so a new field still at its default reads back as itself whether or not `apply_ini` knows it.
- **A format added to `src/formats/lists/built_ins.rs` with no `*_BEFORE_*` constant beside it.**
  `repair_older_lists` (`src/formats/lists/list_store.rs`) rewrites a list only where the file's
  entries match the built-in list or one of the row's `before:` entries, so an installed `config.ini`
  keeps its older list and never receives the format. A fresh install gets it, and no test fails.
- **DPI arithmetic that is not `ui/preview_window/pointer.rs::logical_px` or `text/text_paint.rs::scaled`.**
  A literal in their place is right at 100% and wrong above it, and all sixteen `button_boxes`/`button_at`
  calls the suite makes pass dpi 96 — the 144 and 192 cases in `src/text/pin_chrome/tests/caption.rs`
  repaint the strip rather than place a button — so a caption that stops lining up at 150% fails nothing.
- **A test that shows a window by an activating path.** `tests/no_window_reaches_the_desktop.rs` source-scans
  every test module under `src/ui/preview_window` and fails on any `CreateWindowExW` holding `WS_VISIBLE`
  without `WS_CHILD`, and on any `ShowWindow` outside `SW_SHOWNOACTIVATE`/`SW_HIDE`/`SW_MINIMIZE`/`SW_RESTORE`.
  Make the window through `a_window_of_the_suite` (`src/ui/preview_window/tests.rs`).
- **Ending an engine by pid alone.** `src/app/engine_processes.rs::terminate_verified` refuses unless
  the image name matches and, where the record carries one, the process creation time does too: an id
  read from a `.state` file may belong to something else by the time the sweep runs. The job object's
  `KILL_ON_JOB_CLOSE` is the half that needs no record, its handle never being closed.
- **Work inside a `WH_MOUSE_LL` / `WH_KEYBOARD_LL` callback.** The callback runs in the OS input path,
  where every keystroke and notch on the machine waits for it to return. Set a flag or bump an atomic
  and return; the hooks are `src/shell/key_input.rs` and `src/shell/wheel_input.rs`, and the tick that
  may take a lock is what publishes what they read.
- **Any edit under `src/`.** `build.rs` prints `rerun-if-changed=src`, so one changed line recompiles
  the whole crate and reruns the script; the fingerprint is also keyed on `RHP_WORKTREE`, so a second
  checkout sharing one target dir rebuilds this package on its first build.
- **A box the user chose that is not written to `PinnedPreview::bound`.** A hand's drag, a maximize, its
  restore and a double-click's reset each write it (`apply_pin_drag`, `toggle_pin_maximized`,
  `reset_pinned_size`); one that does not leaves the next file fitted into a stale size, and nothing fails.
- **A window that moves repaints nothing.** Chrome carried in a window of its own is placed on the pin's
  own `WM_MOVE`/`WM_WINDOWPOSCHANGED` (`pin_chrome_strips::follow_the_pin`), not awaited from the paint.
- **A DWM thumbnail composes over its destination's content, children included.** Nothing drawn into a
  child of the pin can stand over the picture, which is why the caption over a page is a `WS_POPUP` —
  and why that popup needs `WS_EX_TRANSPARENT`: a disabled top-level window still takes the press (measured).
- **`HWND_NOTOPMOST` on a window that is already non-topmost is a no-op.** A page under a raised pin is
  put at the top of the ordinary band by claiming topmost and giving it up (`webview_preview::host::raises`).
- **A window sampled by a thumbnail is its client, not the box last asked for.** Compare `GetClientRect`
  (`webview_preview::host::the_window_takes_the_area`), or a band that moved without resizing the window
  leaves the band sampling the older, smaller client.
- **A maximize asked by a caption double-click arrives as `WM_SYSCOMMAND`/`SC_MAXIMIZE`.**
  `DefWindowProcW` gives a native box neither `pin.restore` nor `pin.bound` hears about; answer it with the
  pin's own toggle (`src/ui/preview_window/window_proc.rs`, `toggle_pin_maximized`).

## Where things live

| Path | What it owns | Read more |
| --- | --- | --- |
| `src/main.rs` | `RUNNING`, `CONFIG`, dpi and COM init, thread boot, tray loop | Runtime threads |
| `src/paths.rs` | `plain_path` — the `\\?\` prefix off before the Shell, a browser URL or a media source sees it | Read guards |
| `src/app/` | single instance, startup entry, updates, `engine_processes` (job + record), afk, dialogs | Runtime threads; Config / tray / processes / build |
| `src/config/` | `AppConfig`, defaults, `ini_mapping`, `settings_io`, `engines_in_effect`, `theme_files` | Config / tray / processes / build |
| `src/shell/` | `explorer_hook/`, `key_input`, `wheel_input`, `hook_thread`, `cloud_files`, `pin_navigation`, `tray/` | Explorer hook |
| `src/ui/` | `preview_window.rs` and `preview_window/`: pin, layout, load, paint, media, displays | Preview window |
| `src/formats/` | `routing`, `native_formats::job_for`, `content_type`, `head`, `codecs`, the sixteen `lists` rows | Media pipeline |
| `src/readers/` | in-process decodes: a bitmap or a block of text, cancellable, no process to reap | Kind notes |
| `src/engines/` | out-of-process draws: office, libre, magick, peazip, calibre, webview, the players | Kind notes; Pin mode (the page road) |
| `src/text/` | painted pages, `text_paint`, `text_theme`, `pin_chrome/` | Pin mode; Kind notes |

## Recipes

(a) **Add a setting.**

1. `src/config/config/app_config.rs` — the field.
2. `src/config/config/ini_mapping.rs` — `to_ini`.
3. `src/config/config/ini_mapping.rs` — `apply_ini`. **The pair that fails silently** (trap 1).
4. `src/config/config/settings_io.rs` — `SETTING_GROUPS`.
5. `src/shell/tray/ids.rs` — the id.
6. `src/shell/tray/menus.rs` / `submenus.rs` — the row.
7. `src/shell/tray/commands/toggles.rs` — the handler.
8. `src/shell/tray/event_loop.rs` — the dispatch.

(b) **Add a format.**

1. `src/formats/lists/built_ins.rs` — the extension, in its `DEFAULT_*` list.
2. `src/formats/lists/built_ins.rs` — the `*_BEFORE_*` constant. **The step that fails silently** (trap 2).
3. `src/formats/lists/list_store.rs` — that constant in the row's `before:`.
4. `src/formats/native_formats.rs` — which job draws it.
5. `src/readers/` — a reader of its own, only where the built-in crate cannot take it.

(c) **Add a kind of hover preview.**

1. `src/formats/routing.rs` — `DrawnBy`: does this app read the file, or an engine draw it.
2. `src/formats/native_formats.rs::job_for` — the job that draws it.
3. `src/readers/` or `src/engines/` — the draw itself, behind the thread its kind's work already runs on.
4. `src/ui/preview_window/media_types.rs` — `MediaType`, `media_type`, and `draws_in_a_window_of_its_own`
   (whether the picture in the band is another window's rather than this app's own surface).
5. `src/ui/preview_window/dimensions.rs::media_dimensions` — the box it is measured at; a kind with no
   shape of its own is measured against the room instead, and `pin_swap::pin_keeps_its_box` says whether a
   pin of it is laid out for itself or fitted into the pin's `bound`.
6. `src/ui/preview_window/pin_geometry.rs::pin_frame` — how a pin of it resizes (stretch, reflow, unframed).
7. `src/ui/preview_window/pin_lifecycle.rs::pin_should_stand_in_its_own_window` — whose window a pin of it
   stands in, and `pin_chrome_layout::pin_overlay_chrome` — whether its chrome floats over the picture.
8. `src/config/` and the tray — its `PreviewType` switch, its scaling setting, and whether a cache keeps it.

(d) **Add something to pin mode (a command, a caption button, a menu row).**

1. `src/ui/preview_window/pin_lifecycle.rs` — the `PinCommand` variant.
2. `pin_lifecycle.rs` / `pin_geometry.rs` / `pin_swap.rs` — the road it runs, in the order a pin takes one:
   a box, a media, a window, a state.
3. `src/text/pin_chrome/caption.rs::button_boxes` and `button_at` — the mark and its box; a rendered model
   goes in `page_marks.rs`, and every pixel it places goes through `logical_px` (trap 3).
4. `src/ui/preview_window/pin_input.rs` — the press that lands on it, read in the pin's coordinates.
5. `src/text/pin_chrome/menu.rs` + `src/ui/preview_window/pin_menu.rs` — the row, where the bubble's menu
   and a standing window's share the panel.
6. `src/ui/preview_window/event_loop.rs` — the dispatch; `window_proc.rs` only where a message asks it.
7. `src/ui/preview_window/pin_bubble.rs` — what a collapsed pin owes the same command.

(e) **Windowing: show, place or move a window the pin draws through.**

1. `src/ui/preview_window/pin_real_window.rs` — the pin's own window (the one Windows activates, lists and
   gives a button); `pin_film_host.rs` — the band's host; a page's source is `webview_preview`'s host.
2. `src/ui/preview_window/pin_lifecycle.rs::show_pinned_window` — the show, and
   `place_pinned_siblings` — what stands in the band. **Every road that raises, activates, shows or demotes
   the pin places the siblings again, last**, since they are windows of their own.
3. `src/ui/preview_window/window_proc.rs` — the messages that move it or ask it something:
   `WM_MOVE`/`WM_WINDOWPOSCHANGED`, `WM_ACTIVATE`, `WM_SYSCOMMAND`, `WM_NCHITTEST`.
4. `src/ui/preview_window/paint.rs` — what the band's pixels say: a picture drawn through a thumbnail is an
   alpha-1 claim (`fill_band_hit_testable`) so the press falls to the pin; a page that runs claims nothing.
5. `pin_page_thumbnail.rs` / `pin_film_thumbnail.rs` — the source and its `DwmRegisterThumbnail`, placed once
   and moved after that.
6. `src/ui/preview_window/pin_chrome_strips.rs` — the chrome over the band, a child strip or a popup owned by
   the pin (see ARCHITECTURE.md, "The page road", and the traps above).
7. Tests: a window through `a_window_of_the_suite`; `tests/no_window_reaches_the_desktop.rs` is the rule.

## Ground truth

[ARCHITECTURE.md](ARCHITECTURE.md) says why, and in what order, one section per subsystem; read only the section your change touches.
`codegraph` says where a symbol is and who calls it.
Neither substitutes for reading the code before changing behaviour.