//! The rows that name a value: what a click on one of them writes, and who is told about it.
//!
//! A setter here turns the position or the id a submenu listed an item at into the choice that
//! item stands for, writes it into the configuration, and then says whatever else reads it.
//! Most of them rebuild nothing: a setting whose truth lives somewhere other than the file is
//! told about the change directly, and one read per hover or per frame needs nothing said at all
//! — the next hover answers for itself. The ones that do rebuild are the backdrops, where what
//! stands behind a preview is part of the frame it was composited into, and the scales, which
//! are part of the placement the preview was opened with.
//!
//! The four lookups an id is turned back into a choice by live here too, next to the setters
//! that use them: the submenus hand the answer to whoever asks (see `tray::tests`, which reads
//! them the way the tests of a menu do). The switches are in `toggles` beside this file; the
//! window proc in `event_loop` reaches both through their parent.
//!
//! Every public item here is `pub(in super::super)` rather than `pub(super)` for the reason the
//! sibling file gives: the parent re-exports each of them under the name the window proc already
//! knows it by, and a re-export cannot be wider than the thing it names. `super::super` is
//! `tray`, which is exactly as wide as the `pub(super)` these were written as when they all
//! lived in one file.

use super::super::ids::{
    AUDIO_SEEK_CHOICES, AVOID_CHOICES, BITMAP_SCALE_CHOICES, CACHE_SIZE_CHOICES_MB,
    DECODE_BUDGET_CHOICES_GB, TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, TRAY_CUSTOM_THEMES,
};
use super::super::submenus::{
    afk_timer_secs_at, audio_scale_at, background_at, dds_background_at, document_scale_at,
    engine_idle_at, html_background_at,
};

use crate::app::dialogs;
use crate::config::config::{
    sanitize_decode_budget_gb, sanitize_document_cache_mb, sanitize_general_disk_cache_mb,
    sanitize_image_cache_mb, sanitize_image_disk_cache_mb, sanitize_text_font_scale_percent,
    sanitize_tick_ms, AudioSeek, AvoidMode, MarkdownMode, OfficeEngine, PinNavFileTypes,
    PreviewScale, PreviewType, TextTheme, TriggerKeyMode, VideoEngine, VOLUME_CHOICES,
};
use crate::engines::document_cache;
use crate::engines::office_render;
use crate::ui::preview_window::{
    forget_video_geometry, refresh_preview, refresh_preview_types, trim_image_cache,
    trim_subtitle_cache,
};
use crate::{app::startup, CONFIG};
use std::os::windows::ffi::OsStrExt;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

/// What the trigger key does is a setting rather than a view of one, so the preview
/// on screen is rebuilt: in disable mode a held key is what keeps previews away, and
/// switching to enable mode while it is held should show one.
pub(in super::super) fn set_trigger_key_mode(mode: TriggerKeyMode) {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_mode = mode;
        config.save();
    }
    refresh_preview();
}

/// What a picture is drawn over — and every preview that is not a document: a page,
/// a painted frame, a page Office rendered.
///
/// The backdrop is part of the frame a preview was composited into rather than of
/// the file it was drawn from, so the preview on screen is given its frame again
/// rather than left holding the one it has.
pub(in super::super) fn set_image_background(index: u16) {
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
pub(in super::super) fn set_font_background(index: u16) {
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
pub(in super::super) fn set_html_background(index: u16) {
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
pub(in super::super) fn set_dds_background(index: u16) {
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
pub(in super::super) fn set_design_background(index: u16) {
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
pub(in super::super) fn set_vector_background(index: u16) {
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
pub(in super::super) fn set_theme(theme: TextTheme) {
    if let Ok(mut config) = CONFIG.lock() {
        config.theme = theme;
        config.save();
    }
    refresh_preview();
}

/// The theme a custom item named, by the position the submenu listed it at. A
/// position there is no item for — a click that outlived its menu — selects
/// nothing rather than the wrong theme.
pub(in super::super) fn set_theme_from_menu(index: usize) {
    let theme = TRAY_CUSTOM_THEMES
        .lock()
        .ok()
        .and_then(|themes| themes.get(index).copied());

    if let Some(theme) = theme {
        set_theme(theme);
    }
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
pub(in super::super) fn reset_settings_from_tray() {
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
pub(in super::super) fn reset_lists_from_tray() {
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
pub(in super::super) fn set_office_engine_idle(index: u16) {
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
pub(in super::super) fn set_office_engine(engine: OfficeEngine) {
    if let Ok(mut config) = CONFIG.lock() {
        config.office_engine = engine;
        config.save();
    }

    if engine == OfficeEngine::LibreOffice {
        office_render::stop_engines();
    }
}

/// Which engine plays a video, from `Engine -> Select Engine -> Video`.
///
/// The choice is read live by the router on the next hover, so nothing on screen is rebuilt: a
/// preview that is already up belongs to the engine playing it and is replaced the next time it
/// is laid out. What is given up is the probe cache, because a geometry read by FFprobe is not
/// the geometry a media-engine preview is placed by (or the reverse), so the next hover of a
/// file measures it again (see `forget_video_geometry`).
pub(in super::super) fn set_video_engine(engine: VideoEngine) {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_engine = engine;
        config.save();
    }
    forget_video_geometry();
}

/// How long the browser engine is kept after the last document it drew — for a browser that
/// is marked `Persistent`; one that is not is let go by the AFK timer instead, and this is
/// not consulted for it. Nothing is rebuilt here either: the engine reads the setting every
/// time it decides whether to let itself go, so a shorter time applies to the engine that is
/// already warm.
pub(in super::super) fn set_webview_idle(index: u16) {
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
pub(in super::super) fn set_libreoffice_idle(index: u16) {
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
pub(in super::super) fn set_afk_timer(index: u16) {
    let Some(seconds) = afk_timer_secs_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.afk_timer_seconds = seconds;
        config.save();
    }
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
pub(in super::super) fn set_image_cache_mb(index: u16) {
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
pub(in super::super) fn set_document_cache_mb(index: u16) {
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
pub(in super::super) fn set_image_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_disk_cache_mb = sanitize_image_disk_cache_mb(megabytes);
        config.save();
    }

    document_cache::trim_image_now();
}

/// How much of what a film's own subtitle tracks were copied into may be kept, between hovers.
///
/// The same shape as the two beside it and for the same reason: a film is copied for the hover
/// that asks for it whatever the size, and a size of nothing means nothing is copied at all —
/// such a hover is answered without subtitles rather than with the whole film streamed for
/// them, which is the answer the first hover of a film gets anyway while its copy is coming
/// (see `subtitle_files`).
pub(in super::super) fn set_general_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.general_disk_cache_mb = sanitize_general_disk_cache_mb(megabytes);
        config.save();
    }

    trim_subtitle_cache();
}

/// What one hover may decode or read for, in gigabytes.
///
/// Nothing is rebuilt here, and nothing already on screen changes: every reader asks
/// for the budget as it runs, so the next hover is answered under the new ceiling
/// whatever it is — a smaller one simply refuses more files than a larger one did.
pub(in super::super) fn set_decode_budget_gb(index: u16) {
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
pub(in super::super) fn set_text_font_scale(percent: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        config.text_font_scale_percent = sanitize_text_font_scale_percent(percent);
        config.save();
    }
    refresh_preview();
}

pub(in super::super) fn set_markdown_mode(mode: MarkdownMode) {
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
pub(in super::super) fn set_pin_nav_file_types(mode: PinNavFileTypes) {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_nav_file_types = mode;
        config.save();
    }
}

pub(in super::super) fn set_video_volume(index: u16) {
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
pub(in super::super) fn set_audio_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_volume = volume;
        config.save();
    }
}

/// A way of starting a sound, by the position it was listed at: where in a file a hover drops
/// the needle, which is read as a player is started the way the volume beside it is. An id past
/// the last way the menu offered is one that is not there.
///
/// Nothing on screen is rebuilt: a sound already playing is left where it is, and what a click
/// here changes is where the *next* sound starts — the same bargain the volume beside it makes.
pub(in super::super) fn set_audio_seek(index: u16) {
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
pub(in super::super) fn audio_seek_at(index: u16) -> Option<AudioSeek> {
    AUDIO_SEEK_CHOICES.get(index as usize).copied()
}

pub(in super::super) fn set_follow_cursor(follow: bool) {
    if let Ok(mut config) = CONFIG.lock() {
        config.follow_cursor = follow;
        config.save();
    }
}

/// The way an item of the `Avoid` submenu stands for, by the position it was listed
/// at. An id past the last way the menu offered is one that is not there.
pub(in super::super) fn avoid_mode_at(index: u16) -> Option<AvoidMode> {
    AVOID_CHOICES.get(index as usize).copied()
}

/// How far a preview is kept off the item it is about.
///
/// Where that item's name is drawn is read with the hover, so the region is part of
/// the placement that was made when the preview was opened — and, like the position
/// setting beside it, this applies to the next hover rather than moving the preview
/// that is already up.
pub(in super::super) fn set_avoid_mode(index: u16) {
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
pub(in super::super) fn set_tick_ms(index: u16) {
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
pub(in super::super) fn set_preview_scale(index: u16) {
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
pub(in super::super) fn set_video_scale(index: u16) {
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
pub(in super::super) fn set_animated_scale(index: u16) {
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
pub(in super::super) fn set_ebook_scale(index: u16) {
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
pub(in super::super) fn set_document_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.document_scale = scale;
        config.save();
    }
}

/// How much of the display a sound's card is laid out over, by the position the
/// item was listed at — the card is filled with the room the share names rather
/// than drawn from a bitmap of the file, so the question is the one the document
/// scales beside it answer. The same rule as the ones beside it: the next hover,
/// not the one that is up.
pub(in super::super) fn set_audio_scale(index: u16) {
    let Some(scale) = audio_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_scale = scale;
        config.save();
    }
}

/// How much of the display a font specimen is drawn over, by the position the item was
/// listed at. The same rule as the three beside it: the next hover, not the one that is up.
pub(in super::super) fn set_font_scale(index: u16) {
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
pub(in super::super) fn set_design_scale(index: u16) {
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
pub(in super::super) fn set_vector_scale(index: u16) {
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
pub(in super::super) fn set_text_scale(index: u16) {
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
pub(in super::super) fn bitmap_scale_at(index: u16) -> Option<PreviewScale> {
    BITMAP_SCALE_CHOICES.get(index as usize).copied()
}

/// The delay an item of a `Timing` submenu stands for, by the position it was listed at.
/// The three submenus list the same delays, so one table answers for all of them — and
/// each submenu's own range is what says which setting the click was meant for. An id
/// past the last item is not one the menu offered.
pub(in super::super) fn timing_delay_at(index: u16) -> Option<u64> {
    TIMING_DELAY_CHOICES_MS.get(index as usize).copied()
}

/// How long the pointer must rest on a file before a preview is put up for it, by the
/// position its item was listed at.
pub(in super::super) fn set_hover_delay(index: u16) {
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
pub(in super::super) fn set_same_file_rehover_delay(index: u16) {
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
pub(in super::super) fn set_settling_delay(index: u16) {
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
pub(in super::super) fn open_link(url: &str) {
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

pub(in super::super) fn open_config_file() {
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
