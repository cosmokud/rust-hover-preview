//! What a click on a row of the tray does.
//!
//! One function per command the tray hands out, each of which writes one setting and then
//! tells whatever else reads it. Most of them rebuild nothing: a setting whose truth lives
//! somewhere other than the configuration file is told about the change directly, and one
//! that is read per hover or per frame needs nothing said at all - the next hover answers
//! for itself. The ones that do rebuild are the backdrops, where what stands behind a preview
//! is part of the frame it was composited into, and `toggle_preview_type`, which takes the
//! engine behind a kind of preview with it.
//!
//! The window proc in `event_loop` is what reaches these; the ids that name them are in
//! `tray::ids`, and the lookups that turn an id back into the choice it was listed for are in
//! `submenus`.

use super::ids::{
    AUDIO_SEEK_CHOICES, AVOID_CHOICES, BITMAP_SCALE_CHOICES, CACHE_SIZE_CHOICES_MB,
    DECODE_BUDGET_CHOICES_GB, TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, TRAY_CUSTOM_THEMES,
};
use super::submenus::{
    afk_timer_secs_at, background_at, dds_background_at, document_scale_at, engine_idle_at,
    html_background_at,
};

use crate::app::dialogs;
use crate::config::config::{
    sanitize_decode_budget_gb, sanitize_document_cache_mb, sanitize_image_cache_mb,
    sanitize_image_disk_cache_mb, sanitize_text_font_scale_percent, sanitize_tick_ms, AudioSeek,
    AvoidMode, MarkdownMode, OfficeEngine, PinNavFileTypes, PreviewScale, PreviewType, TextTheme,
    TriggerKeyMode, VOLUME_CHOICES,
};
use crate::engines::document_cache;
use crate::engines::office_render;
use crate::engines::webview_preview;
use crate::ui::preview_window::{
    refresh_pin, refresh_preview, refresh_preview_types, refresh_render_html, trim_image_cache,
};
use crate::{app::startup, CONFIG};
use std::os::windows::ffi::OsStrExt;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

pub(super) fn toggle_startup() {
    // What is flipped is the registry, and what is read to decide which way to flip is the
    // registry too: an entry can be taken away from outside this app, and a toggle that
    // trusted the configuration would then need two clicks to put back — the first turning
    // off something already off. The configuration is the record of the choice, written to
    // agree with what was just done.
    let enable = !startup::is_startup_enabled();

    if enable {
        startup::enable_startup();
    } else {
        startup::disable_startup();
    }

    if let Ok(mut config) = CONFIG.lock() {
        config.run_at_startup = enable;
        config.save();
    }
}

pub(super) fn toggle_preview_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.preview_enabled = !config.preview_enabled;
        config.save();
    }
}

/// Whether the pin key is watched is a setting rather than a view of one, and two
/// things have to be told about a change to it: the hook procedure that watches the
/// key reads a number rather than the configuration, and a preview that is pinned
/// when the feature is switched off is a window nothing would ever take down again.
pub(super) fn toggle_pin_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = !config.pin_enabled;
        config.save();
    }
    crate::shell::key_input::refresh();
    refresh_pin();
}

/// Whether a pin that is up is shown the file the user picks next is a setting rather than a
/// view of one, and it is the one switch here that changes nothing that is already on screen:
/// the pin keeps the file it is showing until the user picks another, and what the switch says
/// is read by the Explorer hook on its next tick (see `pin_update_enabled`).
pub(super) fn toggle_pin_update_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_enabled = !config.pin_update_enabled;
        config.save();
    }
}

/// And whether the pointer's own hover is one of the ways it is told about one, which is read
/// in the same place and changes nothing on screen either (see `pin_update_on_hover`).
pub(super) fn toggle_pin_update_on_hover() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_on_hover = !config.pin_update_on_hover;
        config.save();
    }
}

/// Whether a pin collapsed into its bubble holds the video it is playing is a setting rather than
/// a view of one, and the side that acts on a change is the preview loop on its next tick: a film
/// that is running because the switch was off is held the moment it is switched on, and one that
/// is already held is left where it is until the pin is put back up, because a player started
/// beside a bubble would be a picture on screen next to it (see `settle_bubble_playback`).
pub(super) fn toggle_pin_pause_video() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_pause_video = !config.pin_pause_video;
        config.save();
    }
}

/// The sound's half of the pair above, read on the same tick and acting the same way.
pub(super) fn toggle_pin_pause_audio() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_pause_audio = !config.pin_pause_audio;
        config.save();
    }
}

/// What the trigger key does is a setting rather than a view of one, so the preview
/// on screen is rebuilt: in disable mode a held key is what keeps previews away, and
/// switching to enable mode while it is held should show one.
pub(super) fn set_trigger_key_mode(mode: TriggerKeyMode) {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_mode = mode;
        config.save();
    }
    refresh_preview();
}

/// Whether the key is watched is a setting rather than a view of one, so the preview
/// on screen is rebuilt: switching the key off while it is held lets a preview
/// through, and switching it back on while it is held takes one away.
pub(super) fn toggle_trigger_key_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_enabled = !config.trigger_key_enabled;
        config.save();
    }
    refresh_preview();
}

/// Whether the key reaches a pin is a setting rather than a view of one, and nothing on screen
/// is rebuilt here: the hook reads the switch on its next tick, and a pin it is thrown under is
/// answered from the key's own state as that tick finds it — nothing of the pin belongs to this
/// thread.
pub(super) fn toggle_trigger_key_affect_pin_mode() {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_affect_pin_mode = !config.trigger_key_affect_pin_mode;
        config.save();
    }
}

/// Whether the keyboard driving Explorer holds a parked pointer back is a setting rather
/// than a view of one: nothing on screen is rebuilt and nothing is taken down — the switch
/// is read by the hook on its next tick — so a preview that is up when it is thrown is left
/// where it is.
pub(super) fn toggle_prioritize_keyboard() {
    if let Ok(mut config) = CONFIG.lock() {
        config.prioritize_keyboard = !config.prioritize_keyboard;
        config.save();
    }
}

/// What a picture is drawn over — and every preview that is not a document: a page,
/// a painted frame, a page Office rendered.
///
/// The backdrop is part of the frame a preview was composited into rather than of
/// the file it was drawn from, so the preview on screen is given its frame again
/// rather than left holding the one it has.
pub(super) fn set_image_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_background = background;
        config.save();
    }
    refresh_preview();
}

/// And the same again for a font specimen, which is drawn on a page of its own: the page's
/// colours are part of what the engine draws, so the preview on screen is rebuilt rather
/// than only composited again.
pub(super) fn set_font_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.font_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a page of HTML, which is drawn on a page the engine owns the way a specimen is:
/// the page's colours are part of what the browser draws, so the preview on screen is
/// rebuilt rather than only composited again.
pub(super) fn set_html_background(index: u16) {
    let Some(background) = html_background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.html_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a texture, which is a picture like any other on this side of the answer: its
/// frame is composited by this app, so the preview on screen only needs compositing again.
pub(super) fn set_dds_background(index: u16) {
    let Some(background) = dds_background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.dds_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a design document, which is composited by this app the way a texture is: what
/// stands behind the picture the file keeps of the document is this side's to draw, so the
/// preview on screen only needs compositing again.
pub(super) fn set_design_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.design_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a vector drawing, which is drawn over a backdrop of its own: an SVG document is
/// drawn on a page the engine owns, so the page's colours are part of what it draws and the
/// preview on screen is rebuilt rather than only composited again; a metafile is replayed
/// by this side, so a change to it is composited again like a picture's.
pub(super) fn set_vector_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.vector_background = background;
        config.save();
    }
    refresh_preview();
}

/// A text preview's colors and glyphs are painted into its frame, so the preview
/// on screen is rebuilt rather than only composited again.
pub(super) fn set_theme(theme: TextTheme) {
    if let Ok(mut config) = CONFIG.lock() {
        config.theme = theme;
        config.save();
    }
    refresh_preview();
}

/// The theme a custom item named, by the position the submenu listed it at. A
/// position there is no item for — a click that outlived its menu — selects
/// nothing rather than the wrong theme.
pub(super) fn set_theme_from_menu(index: usize) {
    let theme = TRAY_CUSTOM_THEMES
        .lock()
        .ok()
        .and_then(|themes| themes.get(index).copied());

    if let Some(theme) = theme {
        set_theme(theme);
    }
}

/// Switching a kind of preview off drops the one on screen when it is of that
/// kind — the hover that produced it no longer measures — and switching one back
/// on leaves the preview that is up alone, since a preview of another kind has
/// nothing to do with the gate that changed. The Explorer hook reads the gates
/// fresh on every tick, so a change needs no restart and no cache to clear.
///
/// A kind switched off also takes the engine behind it: what an engine is for is
/// previews of its own kind, and one kept warm for a kind nobody can be shown is a
/// process — and a licence, and a few hundred megabytes — held for nothing.
pub(super) fn toggle_preview_type(kind: PreviewType) {
    if let Ok(mut config) = CONFIG.lock() {
        let enabled = kind.enabled_in(&config);
        kind.set_enabled_in(&mut config, !enabled);
        config.save();
    }

    if !kind.enabled() {
        // An engine only one kind of preview is ever started for goes with that
        // kind's gate. The Office processes are ended from here rather than through
        // the worker — a worker inside a call it cannot cut short would hold one of
        // them for good, and this thread may not wait on one — while the browser
        // needs nothing said to it at all: its own thread reads this gate and lets
        // the engine go.
        if let PreviewType::Document = kind {
            office_render::stop_engines();
        }

        // A page of HTML the browser is drawing rides the text kind's gate: switched off,
        // the page comes down with it (see `webview_preview::hide_html_preview`).
        if let PreviewType::Text = kind {
            webview_preview::hide_html_preview();
        }
    }

    refresh_preview_types();
}

/// Put every setting back at what this build recommends, with the extension lists left
/// exactly as they are.
///
/// The reset itself is the configuration's own (`reset_to_recommended`), and everything
/// after it is the rest of the app being told what a reset means: it is every tray toggle
/// at once, so a setting whose truth lives somewhere other than the file has to be put
/// back on the machine as well, and what a setting was measuring has to be asked again.
/// The question comes first and names what would change, which is the same difference the
/// row is offered on.
pub(super) fn reset_settings_from_tray() {
    let changes = CONFIG
        .lock()
        .map(|config| config.settings_apart_from_recommended())
        .unwrap_or_default();

    if !dialogs::confirm_reset_settings(&changes) {
        return;
    }

    let run_at_startup = {
        let Ok(mut config) = CONFIG.lock() else {
            return;
        };

        config.reset_to_recommended();
        config.save();
        config.run_at_startup
    };

    // The entry is the registry's and the configuration is the record of the choice, the
    // same way round as the toggle beside it: a reset that turns it back on writes the
    // entry, or the file and the machine would disagree about what starts this app.
    if run_at_startup != startup::is_startup_enabled() {
        if run_at_startup {
            startup::enable_startup();
        } else {
            startup::disable_startup();
        }
    }

    // What a reset can turn off as easily as on, and the one setting whose engines are
    // ended from here when it does — the same call its own toggle makes.
    if !PreviewType::Document.enabled() {
        office_render::stop_engines();
    }

    refresh_preview_types();
    refresh_preview();
}

/// Put every extension list back at the built-in one, with every other setting left alone.
///
/// It is the same question over the other half of the configuration. What a kind of
/// preview matches a file against is its list, so what is on screen is asked whether it
/// still measures — and nothing outside the file reads these, which is why this one has no
/// registry to write and no engine to let go.
pub(super) fn reset_lists_from_tray() {
    let sections = CONFIG
        .lock()
        .map(|config| config.lists_apart_from_built_in())
        .unwrap_or_default();

    if !dialogs::confirm_reset_lists(&sections) {
        return;
    }

    if let Ok(mut config) = CONFIG.lock() {
        config.reset_extension_lists();
        config.save();
    }

    refresh_preview_types();
    refresh_preview();
}

/// How long the Office engines are kept after their families' last pages — for an engine
/// marked `Persistent`. One that is not is let go by the AFK timer instead, and this is not
/// consulted for it (see `app::afk`).
///
/// Nothing is rebuilt here and nothing on screen changes. An engine that is being
/// let go sooner is let go by the worker the next time it looks — which is twice a
/// second while it is holding one — and one that is being kept longer is a setting
/// the next render already reads, so neither needs waking.
pub(super) fn set_office_engine_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.office_engine_idle = idle;
        config.save();
    }
}

/// Which engine Office documents are asked of, from `Engine → Select Engine → Office`.
///
/// Nothing is rebuilt here, and nothing has to be: the choice is read live by the side that
/// asks an engine for a page and by the side that draws one (see `office_formats::page_engine`),
/// so the setting a click leaves behind is the one the next hover is answered by. A preview
/// that is already on screen belongs to the engine that drew it and is replaced the next time
/// a hover is raised — and the pointer has left the file to reach the tray by then.
///
/// What is ended here is this app's own Office applications, at the moment the render engine
/// becomes the one to ask: they were started for pages this choice now takes elsewhere, and
/// the one thing this menu is not for is a process kept warm for work it will not be given.
/// A user's own Word or Excel is not one of these and is never touched (see `office_render`).
pub(super) fn set_office_engine(engine: OfficeEngine) {
    if let Ok(mut config) = CONFIG.lock() {
        config.office_engine = engine;
        config.save();
    }

    if engine == OfficeEngine::LibreOffice {
        office_render::stop_engines();
    }
}

/// How long the browser engine is kept after the last document it drew — for a browser that
/// is marked `Persistent`; one that is not is let go by the AFK timer instead, and this is
/// not consulted for it. Nothing is rebuilt here either: the engine reads the setting every
/// time it decides whether to let itself go, so a shorter time applies to the engine that is
/// already warm.
pub(super) fn set_webview_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.webview_idle = idle;
        config.save();
    }
}

/// How long the LibreOffice engine is kept after the last page it drew — for an engine marked
/// `Persistent`; one that is not is let go by the AFK timer instead, and this is not consulted
/// for it.
///
/// Nothing is rebuilt here either, and nothing has to be: the engine thread reads the setting
/// every second while it waits for documents, so a shorter time applies to the engine that is
/// already running, and `0 seconds` — the bottom of the list — lets go of one within the
/// second. An engine that is kept is a process this app holds and ends itself; a setting of
/// `indefinitely` keeps it for the rest of the run (see `libreoffice_render`).
pub(super) fn set_libreoffice_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.libreoffice_idle = idle;
        config.save();
    }
}

/// How long Explorer may be out of reach before an engine that is not marked `Persistent` is
/// let go.
///
/// Nothing is rebuilt here and nothing on screen changes for the same reason the idle times
/// rebuild nothing: an engine that is not persistent reads this on the look it already takes
/// — the Office worker twice a second, the engine thread once a second, the browser every
/// quarter of one while it is up — so a shorter time lets go of an engine that is already
/// warm, and a longer one keeps an engine the next look would have let go of.
pub(super) fn set_afk_timer(index: u16) {
    let Some(seconds) = afk_timer_secs_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.afk_timer_seconds = seconds;
        config.save();
    }
}

/// Whether one of the three engines is kept whatever the user is doing, from the
/// `Persistent` toggle at the top of its TTL submenu.
///
/// The toggle is which submenu it heads rather than which engine it names, because the three
/// are listed in one order and the ids are handed out in it. Nothing is rebuilt here either:
/// both sides of the setting are read live by the engine that decides with them, so a toggle
/// turned on keeps the engine the next look would have let go of, and one turned off lets go
/// of an engine that is already up as soon as the AFK timer says it may.
pub(super) fn toggle_engine_persistent(index: u16) {
    let Ok(mut config) = CONFIG.lock() else {
        return;
    };

    match index {
        0 => config.office_engine_persistent = !config.office_engine_persistent,
        1 => config.libreoffice_persistent = !config.libreoffice_persistent,
        2 => config.webview_persistent = !config.webview_persistent,
        _ => return,
    }

    config.save();
}

/// Whether a video is decoded on the graphics card, turned the other way round.
///
/// Nothing *already on screen* changes when this is switched: the setting is read when a player is
/// launched and a film already playing was launched with the other answer, so it is kept playing
/// rather than being relaunched under the user's feet. That is the same bargain every other switch
/// in this menu makes, and it is why the row is a question about the next preview rather than about
/// the one under the pointer.
///
/// The *next* preview, though, is the next one in this run of the app and not the next one after a
/// restart, which is what the answer had to be told for it to be true: the device is found by a
/// probe that asks FFmpeg to decode a tenth of a second of a synthetic pattern on each candidate in
/// turn, and a probe that ran with the setting off names no device at all. So the answer found
/// before the switch is an answer to the question as it stood then, and `forget` is what makes every
/// answer so far one to a question no longer being asked. Without it this row would take effect on
/// the next run of the app and its own note above would be quietly wrong.
pub(super) fn toggle_video_hw_accel() {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_hw_accel = !config.video_hw_accel;
        config.save();
    }

    crate::ui::preview_window::forget_video_hw_accel_answer();
}

/// The size an item of the `Cache` submenu stands for, by the position it was
/// listed at. An id past the last size the menu offered is one that is not there.
fn cache_size_at(index: u16) -> Option<u32> {
    CACHE_SIZE_CHOICES_MB.get(index as usize).copied()
}

/// How much memory the decoded-image cache may hold.
///
/// A preview on screen is not drawn from the cache but from the frame it was
/// loaded as, so nothing on screen changes — what changes is how much is freed, and
/// a smaller size frees it now rather than at the next decode.
pub(super) fn set_image_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_cache_mb = sanitize_image_cache_mb(megabytes);
        config.save();
    }

    trim_image_cache();
}

/// How much of what an engine drew may be kept, between hovers.
///
/// Unlike the image cache beside it this does not switch anything off: a page is drawn for the
/// hover that asks for it whatever the size, and a size of nothing means it is given up when
/// that hover ends. So nothing is rebuilt here either — the next hover answers for itself.
pub(super) fn set_document_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.document_cache_mb = sanitize_document_cache_mb(megabytes);
        config.save();
    }

    document_cache::trim_now();
}

/// How much of what the image converter developed may be kept, between hovers.
///
/// The same shape as the document cache beside it, and for the same reason: a picture is
/// developed for the hover that asks for it whatever the size, and a size of nothing means the
/// hover after it pays for the development again. Nothing is rebuilt here either — the pages are
/// read by the hover that wants one, so what a size does is bound what is left for it to read.
pub(super) fn set_image_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_disk_cache_mb = sanitize_image_disk_cache_mb(megabytes);
        config.save();
    }

    document_cache::trim_image_now();
}

/// What one hover may decode or read for, in gigabytes.
///
/// Nothing is rebuilt here, and nothing already on screen changes: every reader asks
/// for the budget as it runs, so the next hover is answered under the new ceiling
/// whatever it is — a smaller one simply refuses more files than a larger one did.
pub(super) fn set_decode_budget_gb(index: u16) {
    let Some(gigabytes) = DECODE_BUDGET_CHOICES_GB.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.decode_budget_gb = sanitize_decode_budget_gb(gigabytes);
        config.save();
    }
}

/// A text preview's own keys and pointer are the pin's business now: the row that used to
/// switch full mode on for every hover of a text file is gone, and pinning one with the pin key
/// is what brings it up as something to work in (see `current_text_options`).
pub(super) fn set_text_font_scale(percent: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        config.text_font_scale_percent = sanitize_text_font_scale_percent(percent);
        config.save();
    }
    refresh_preview();
}

pub(super) fn set_markdown_mode(mode: MarkdownMode) {
    if let Ok(mut config) = CONFIG.lock() {
        config.markdown_mode = mode;
        config.save();
    }
    refresh_preview();
}

/// Which files the previous/next buttons on a pin's own caption step through: every file this
/// build can preview, or only those of the kind of thing the pinned file is.
///
/// Nothing on screen changes and nothing is rebuilt. The walk is read per step rather than
/// held, and what it is made of is a question answered off the folder when a button is
/// pressed — so the next step of a pin already up is the first one under the new setting,
/// and a pin with no folder of its own to walk is not a walk at all either way.
pub(super) fn set_pin_nav_file_types(mode: PinNavFileTypes) {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_nav_file_types = mode;
        config.save();
    }
}

/// Whether a page of HTML is drawn by the browser engine rather than shown as its markup.
///
/// What changes is which of the two things draws a `.htm`, and the whole menu is read from
/// the configuration, so the row is ticked from the setting on the next open rather than
/// kept in step by hand. The two renderers are not one swapped in place of the other — the
/// page is a window of the engine's and the markup a frame of this app's — so the preview
/// on screen is rebuilt from the hover it came from rather than left standing. The direction
/// that leaves the engine has its page taken down with the switch, though a pin's is left
/// as it is (see `refresh_render_html` and `refresh_preview`).
pub(super) fn toggle_render_html() {
    let mut turned_off = false;
    if let Ok(mut config) = CONFIG.lock() {
        config.render_html = !config.render_html;
        turned_off = !config.render_html;
        config.save();
    }

    // The direction that leaves the engine is the one that has to say so: a preview
    // rebuilt from its hover is what a painted page owes (see `refresh_preview`), while
    // a page the engine is drawing for a hover is a window nothing else would take down
    // — and a pin's page is left as it is (see `refresh_render_html`).
    if turned_off {
        refresh_render_html();
    }

    refresh_preview();
}

pub(super) fn set_video_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.video_volume = volume;
        config.save();
    }
}

/// A level of the `Volume → Audio` submenu, by the position it was listed at: the volume a
/// sound file is played at, which is the setting a card is drawn against as well as the one its
/// player is started with.
pub(super) fn set_audio_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_volume = volume;
        config.save();
    }
}

/// The switch above a sound's levels: whether a file's measured loudness is brought to one level
/// before it is played, which is read where a player is started the way the level
/// beside it is — nothing on screen is rebuilt, and a sound already playing is left where it is.
///
/// It is the file that is measured rather than the playing that is re-scaled, so a file switched
/// on for while it was already known is measured off the tick, and what that measurement is for is
/// the hover after the one that asked for it (see `spawn_gain_scan`).
pub(super) fn toggle_normalize_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.normalize_volume = !config.normalize_volume;
        config.save();
    }
}

/// And the video's own switch, which is the same question asked about a film's soundtrack: read
/// where its player is started, like the level beside it, and off where the app starts — a film is
/// looked at, and a measurement of its audio is a decode of the film.
pub(super) fn toggle_normalize_video_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.normalize_video_volume = !config.normalize_video_volume;
        config.save();
    }
}

/// The row under a sound's `Normalize`: whether a level turned on a pinned window's own knob is
/// the level the next sound is previewed at.
///
/// Nothing on screen is rebuilt, and a sound already playing is left where it is: what the switch
/// changes is where the *next* player starts, which is the same bargain the level beside it makes.
pub(super) fn toggle_remember_audio_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.remember_audio_volume = !config.remember_audio_volume;
        config.save();
    }
}

/// And the video's own, which is the same switch asked about a soundtrack and off where the app
/// starts, like the video's level it decides.
pub(super) fn toggle_remember_video_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.remember_video_volume = !config.remember_video_volume;
        config.save();
    }
}

/// A way of starting a sound, by the position it was listed at: where in a file a hover drops
/// the needle, which is read as a player is started the way the volume beside it is. An id past
/// the last way the menu offered is one that is not there.
///
/// Nothing on screen is rebuilt: a sound already playing is left where it is, and what a click
/// here changes is where the *next* sound starts — the same bargain the volume beside it makes.
pub(super) fn set_audio_seek(index: u16) {
    let Some(seek) = audio_seek_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_seek = seek;
        config.save();
    }
}

/// The way an item of the `Volume → Audio Seek` submenu stands for, by the position it was
/// listed at. An id past the last way the menu offered is one that is not there.
pub(super) fn audio_seek_at(index: u16) -> Option<AudioSeek> {
    AUDIO_SEEK_CHOICES.get(index as usize).copied()
}

pub(super) fn set_follow_cursor(follow: bool) {
    if let Ok(mut config) = CONFIG.lock() {
        config.follow_cursor = follow;
        config.save();
    }
}

/// The way an item of the `Avoid` submenu stands for, by the position it was listed
/// at. An id past the last way the menu offered is one that is not there.
pub(super) fn avoid_mode_at(index: u16) -> Option<AvoidMode> {
    AVOID_CHOICES.get(index as usize).copied()
}

/// How far a preview is kept off the item it is about.
///
/// Where that item's name is drawn is read with the hover, so the region is part of
/// the placement that was made when the preview was opened — and, like the position
/// setting beside it, this applies to the next hover rather than moving the preview
/// that is already up.
pub(super) fn set_avoid_mode(index: u16) {
    let Some(mode) = avoid_mode_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.avoid_mode = mode;
        config.save();
    }
}

/// The tick an item of the `Tick` submenu stands for, by the position it was listed at.
/// An id past the last tick the menu offered is one that is not there.
fn tick_ms_at(index: u16) -> Option<u64> {
    TICK_CHOICES_MS.get(index as usize).copied()
}

/// How often the app looks at the pointer's world while Explorer has focus.
///
/// Nothing on screen changes and nothing is rebuilt: the preview that is up was placed
/// when it was opened, and the tick is what says when the next look happens — so a
/// slower tick is a preview that stays a moment longer after the pointer has left it,
/// and a faster one is an answer that arrives sooner, at the cost of more crossings into
/// Explorer (see `DEFAULT_TICK_MS`).
pub(super) fn set_tick_ms(index: u16) {
    let Some(tick_ms) = tick_ms_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.tick_ms = sanitize_tick_ms(tick_ms);
        config.save();
    }
}

/// How large a picture is drawn, by the position the item was listed at.
///
/// The size a bitmap is drawn at is part of the placement that was made when the preview
/// was opened — the box is sized, and the frame is scaled into it — so, like the position
/// beside it, this applies to the next hover rather than resizing the preview that is up.
pub(super) fn set_preview_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.preview_scale = scale;
        config.save();
    }
}

/// The same for a video, at the share `video_scale` names: the frame that stands in for
/// one is a bitmap like a picture, so the shares its submenu offers are the picture's
/// shares, and the setting written is the video's own.
pub(super) fn set_video_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.video_scale = scale;
        config.save();
    }
}

/// And the same for an animated picture, at the share `animated_scale` names: the frames
/// an animation decodes into are bitmaps like a picture's, so this submenu offers the same
/// shares as the two beside it, and the setting written is the animation's own. Which
/// files are animated is not a setting at all: a GIF, a WebP or a PNG is one when the file
/// itself holds more than a single frame, and a still one keeps the picture scale.
pub(super) fn set_animated_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.animated_scale = scale;
        config.save();
    }
}

/// How much of the display a PDF page — the `Ebook` kind — is drawn over, by the position
/// the item was listed at.
///
/// The size a document is drawn at is part of the placement that was made when the preview
/// was opened — the box is sized, and the document is drawn into it — so, like the position
/// and the picture scale beside it, this applies to the next hover rather than resizing the
/// preview that is already up.
pub(super) fn set_ebook_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.ebook_scale = scale;
        config.save();
    }
}

/// How much of the display a page of the `Document` kind is shown over, by the position the
/// item was listed at — the one setting behind both halves of the kind: a page an Office
/// document's own application exported, and a page the render engine drew.
pub(super) fn set_document_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.document_scale = scale;
        config.save();
    }
}

/// How much of the display a font specimen is drawn over, by the position the item was
/// listed at. The same rule as the three beside it: the next hover, not the one that is up.
pub(super) fn set_font_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.font_scale = scale;
        config.save();
    }
}

/// How much of the display a design document is drawn over, by the position the item was
/// listed at. The same rule as the four beside it: what a document is previewed from is
/// the picture its own format keeps of the whole thing, so the share is of the display
/// rather than of the document, and it applies to the next hover rather than resizing a
/// preview that is already up.
pub(super) fn set_design_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.design_scale = scale;
        config.save();
    }
}

/// How much of the display a vector drawing is replayed over, by the position the item was
/// listed at — the same rule as the documents beside it: the share is of the display, and
/// it applies to the next hover rather than resizing a preview that is already up.
pub(super) fn set_vector_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.vector_scale = scale;
        config.save();
    }
}

/// How much of the display a page of text is measured in, by the position the item was listed
/// at. The same rule as the documents beside it: the box a text page is measured for is part of
/// the placement that was made when the preview was opened, so this applies to the next hover
/// rather than resizing the page that is already up.
pub(super) fn set_text_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.text_scale = scale;
        config.save();
    }
}

/// The share of its own size an item of the `Image Scaling` or `Video Scaling` submenu
/// stands for, by the position it was listed at. An id past the last choice the menu
/// offered is one that is not there.
pub(super) fn bitmap_scale_at(index: u16) -> Option<PreviewScale> {
    BITMAP_SCALE_CHOICES.get(index as usize).copied()
}

/// The delay an item of a `Timing` submenu stands for, by the position it was listed at.
/// The three submenus list the same delays, so one table answers for all of them — and
/// each submenu's own range is what says which setting the click was meant for. An id
/// past the last item is not one the menu offered.
pub(super) fn timing_delay_at(index: u16) -> Option<u64> {
    TIMING_DELAY_CHOICES_MS.get(index as usize).copied()
}

/// How long the pointer must rest on a file before a preview is put up for it, by the
/// position its item was listed at.
pub(super) fn set_hover_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.hover_delay_ms = delay_ms;
        config.save();
    }
}

/// How long the same file waits before a preview of it is put up again, by the position
/// its item was listed at.
pub(super) fn set_same_file_rehover_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.same_file_rehover_delay_ms = delay_ms;
        config.save();
    }
}

/// How long the pointer must be still before a preview may open for anything, by the
/// position its item was listed at.
pub(super) fn set_settling_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.settling_delay_ms = delay_ms;
        config.save();
    }
}

/// A page, opened in the browser the user already has: the same call the release page is
/// opened with, and nothing is fetched or run here — what becomes of the page is the
/// browser's own business (see `updates::open_release_page`).
pub(super) fn open_link(url: &str) {
    let wide_url: Vec<u16> = url.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let _ = ShellExecuteW(
            HWND(std::ptr::null_mut()),
            w!("open"),
            PCWSTR(wide_url.as_ptr()),
            PCWSTR(std::ptr::null()),
            PCWSTR(std::ptr::null()),
            SW_SHOWNORMAL,
        );
    }
}

pub(super) fn open_config_file() {
    if let Ok(config) = CONFIG.lock() {
        config.save();
    }

    if let Some(path) = crate::config::config::AppConfig::config_path() {
        let wide_path: Vec<u16> = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();

        unsafe {
            let _ = ShellExecuteW(
                HWND(std::ptr::null_mut()),
                w!("open"),
                PCWSTR(wide_path.as_ptr()),
                PCWSTR(std::ptr::null()),
                PCWSTR(std::ptr::null()),
                SW_SHOWNORMAL,
            );
        }
    }
}
