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
    ANIMATED_WEBP_BUDGET_CHOICES_MB, AUDIO_SEEK_CHOICES, AVOID_CHOICES, BITMAP_SCALE_CHOICES,
    DECODE_BUDGET_CHOICES_GB, DISK_CACHE_SIZE_CHOICES_MB, IMAGE_CACHE_SIZE_CHOICES_MB,
    OTHERS_CACHE_SIZE_CHOICES_MB, TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, TRAY_CUSTOM_THEMES,
};
use super::super::submenus::{
    afk_timer_secs_at, audio_scale_at, background_at, dds_background_at, document_scale_at,
    end_of_file_at, engine_idle_at, html_background_at, zoom_background_at,
};

use crate::app::dialogs;
use crate::app::seam::{
    refresh_preview, refresh_preview_types, send_preview, trim_subtitle_cache, PreviewMessage,
};
use crate::config::config::{
    sanitize_animated_webp_budget_mb, sanitize_decode_budget_gb, sanitize_document_cache_mb,
    sanitize_general_disk_cache_mb, sanitize_image_cache_mb, sanitize_image_disk_cache_mb,
    sanitize_others_cache_mb, sanitize_text_font_scale_percent, sanitize_tick_ms, AudioSeek,
    AvoidMode, MarkdownMode, OfficeEngine, PinMinimizeTo, PinNavFileTypes, PinScalingMode,
    PreviewScale, PreviewType, TextTheme, TriggerKeyMode, VideoEngine, VOLUME_CHOICES,
};
use crate::config::{commit_config, publish_config_change, snapshot};
use crate::formats::codecs;
use crate::{app::startup, CONFIG};

/// What the trigger key does is a setting rather than a view of one, so the preview
/// on screen is rebuilt: in disable mode a held key is what keeps previews away, and
/// switching to enable mode while it is held should show one.
pub(in super::super) fn set_trigger_key_mode(mode: TriggerKeyMode) {
    publish_config_change(
        |config| {
            config.trigger_key_mode = mode;
        },
        &[PreviewMessage::Refresh],
    );
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

    publish_config_change(
        |config| {
            config.image_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// The band beside a picture the wheel has taken past its window, which is filled by a
/// backdrop of the window's rather than of the file's.
///
/// The frame is composited by this app the way a picture's own backdrop is, so the preview
/// on screen is given its frame again rather than left holding the one it has.
pub(in super::super) fn set_zoom_background(index: u16) {
    let Some(background) = zoom_background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.zoom_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// And the same again for a font specimen, which is drawn on a page of its own: the page's
/// colours are part of what the engine draws, so the preview on screen is rebuilt rather
/// than only composited again.
pub(in super::super) fn set_font_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.font_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// And for a page of HTML, which is drawn on a page the engine owns the way a specimen is:
/// the page's colours are part of what the browser draws, so the preview on screen is
/// rebuilt rather than only composited again.
pub(in super::super) fn set_html_background(index: u16) {
    let Some(background) = html_background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.html_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// And for a texture, which is a picture like any other on this side of the answer: its
/// frame is composited by this app, so the preview on screen only needs compositing again.
pub(in super::super) fn set_dds_background(index: u16) {
    let Some(background) = dds_background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.dds_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// And for a design document, which is composited by this app the way a texture is: what
/// stands behind the picture the file keeps of the document is this side's to draw, so the
/// preview on screen only needs compositing again.
pub(in super::super) fn set_design_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.design_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// And for a vector drawing, which is drawn over a backdrop of its own: an SVG document is
/// drawn on a page the engine owns, so the page's colours are part of what it draws and the
/// preview on screen is rebuilt rather than only composited again; a metafile is replayed
/// by this side, so a change to it is composited again like a picture's.
pub(in super::super) fn set_vector_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    publish_config_change(
        |config| {
            config.vector_background = background;
        },
        &[PreviewMessage::Refresh],
    );
}

/// A text preview's colors and glyphs are painted into its frame, so the preview
/// on screen is rebuilt rather than only composited again.
pub(in super::super) fn set_theme(theme: TextTheme) {
    publish_config_change(
        |config| {
            config.theme = theme;
        },
        &[PreviewMessage::Refresh],
    );
    crate::ui::custom_preview_settings::notify_theme_changed();
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
    let changes = snapshot().settings_apart_from_recommended();

    if !dialogs::confirm_reset_settings(&changes) {
        return;
    }

    let (run_at_startup, doc_enabled) = commit_config(|config| {
        config.reset_to_recommended();
        (
            config.run_at_startup,
            PreviewType::Document.enabled_in(config),
        )
    });

    // The entry is the registry's and the configuration is the record of the choice, the
    // same way round as the toggle beside it: a reset that turns it back on writes the
    // entry, or the file and the machine would disagree about what starts this app.
    if run_at_startup != startup::is_startup_enabled() {
        let outcome = if run_at_startup {
            startup::enable_startup()
        } else {
            startup::disable_startup()
        };

        // And the file keeps what the machine holds rather than what the reset asked for, for the
        // reason the toggle does (see `toggles::toggle_startup`).
        let held = startup::is_startup_enabled();
        if held != run_at_startup {
            commit_config(|config| {
                config.run_at_startup = held;
            });
        }

        if !outcome.settled() {
            dialogs::report_startup_failure(outcome);
        }
    }

    // What a reset can turn off as easily as on, and the one setting whose engines are
    // asked to go when it does — the same ask its own toggle makes, in the process holding
    // them (see `PreviewMessage::StopEngines`).
    if !doc_enabled {
        send_preview(PreviewMessage::StopEngines);
    }

    refresh_preview_types();
    refresh_preview();
    crate::ui::custom_preview_settings::notify_theme_changed();
}

/// Put every extension list back at the built-in one, with every other setting left alone.
///
/// It is the same question over the other half of the configuration. What a kind of
/// preview matches a file against is its list, so what is on screen is asked whether it
/// still measures — and nothing outside the file reads these, which is why this one has no
/// registry to write and no engine to let go.
pub(in super::super) fn reset_lists_from_tray() {
    let sections = snapshot().lists_apart_from_built_in();

    if !dialogs::confirm_reset_lists(&sections) {
        return;
    }

    publish_config_change(
        |config| {
            config.reset_extension_lists();
        },
        &[PreviewMessage::RefreshTypes, PreviewMessage::Refresh],
    );
}

/// How long the mpv a film is played by is kept between films, from `Engine → MPV TTL`.
///
/// Nothing is rebuilt here: the setting is stored for the menu that just wrote it, and what
/// keeps an mpv warm is not written yet.
pub(in super::super) fn set_mpv_engine_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    commit_config(|config| {
        config.video_engine_idle = idle;
    });
}

/// How long the Office engines are kept after their families' last pages, from
/// `Engine → Microsoft Office TTL`.
///
/// Nothing is rebuilt here and nothing on screen changes. An engine that is being
/// let go sooner is let go by the worker the next time it looks — which is twice a
/// second while it is holding one — and one that is being kept longer is a setting
/// the next render already reads, so neither needs waking.
pub(in super::super) fn set_office_engine_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    commit_config(|config| {
        config.office_engine_idle = idle;
    });
}

/// Which engine Office documents are asked of, from `Engine → Office Engine`.
///
/// Which engine draws an Office document's page, from `Engine → Office Engine`.
///
/// **Nothing on screen is rebuilt for it.** The choice is written here and read where a preview
/// is begun, so a document already up belongs to the engine that drew it, and the next hover —
/// and the next document a pinned window is shown — is the one drawn by the engine named here
/// (see `engines_in_effect`). What the row said before that was that the choice *is read live
/// by the side that asks an engine for a page*, which was true of the next hover and was not
/// true of the document on screen: every box change and every step asked again, and a page
/// keyed by the engine that drew it was asked of the other one.
///
/// What is ended here is this app's own Office applications, at the moment the render engine
/// becomes the one to ask: they were started for pages this choice now takes elsewhere, and
/// the document on screen keeps the page the engine it was begun with gave it rather than
/// being asked for again (see `engines_in_effect`, `document_cache`). A user's own Word or
/// Excel is not one of these and is never touched (see `office_render`).
pub(in super::super) fn set_office_engine(engine: OfficeEngine) {
    commit_config(|config| {
        config.office_engine = engine;
    });

    if engine == OfficeEngine::LibreOffice {
        send_preview(PreviewMessage::StopEngines);
    }
}

/// Which engine plays a video, from `Engine → Video Engine`.
///
/// **Nothing on screen is rebuilt for it.** The choice is written here and read where a
/// preview is begun, so a film already playing belongs to the player that started it, and the
/// next hover — or the next file a pinned window is shown — is the one played by the engine
/// named here (see `engines_in_effect`). What the row said before that was that a preview is
/// *replaced the next time it is laid out*, which is true of the next preview and was not true
/// of the one on screen: a resize, a maximize, a gesture and a step are each a question about
/// the file that is up, and each of them came back with the engine of now.
///
/// What is given up is the probe cache, because a geometry read by FFprobe is not the
/// geometry a media-engine preview is placed by (or the reverse) — given up with that read
/// rather than here, since the preview on screen is still placed by the geometry it was built
/// under (see `forget_video_geometry`).
pub(in super::super) fn set_video_engine(engine: VideoEngine) {
    commit_config(|config| {
        config.video_engine = engine;
    });
}

/// Which `mpv.exe` a launch runs, from `Components → Videos → MPV`, by the position its install was
/// listed at.
///
/// The installs are asked of the machine again rather than remembered from the menu build, so an
/// id means the same thing on both sides: the same list, in the same order, is what the menu
/// listed (see `codecs::mpv_installs`). A position with no install behind it — the machine
/// changed between the menu opening and the click — writes nothing.
///
/// Nothing on screen is rebuilt: the next launch reads the setting, and a film already playing
/// belongs to the player that started it.
pub(in super::super) fn set_mpv_path(index: u16) {
    let installs = codecs::mpv_installs();

    let Some(install) = installs.get(index as usize) else {
        return;
    };

    commit_config(|config| {
        config.mpv_path = install.to_string_lossy().into_owned();
    });
}

/// How long the browser engine is kept after the last document it drew, from
/// `Engine → WebView2 TTL`. Every value is the setting the engine is kept by, and `Until away`
/// is the AFK timer (see `EngineIdle`). Nothing is rebuilt here either: the engine reads the
/// setting every time it decides whether to let itself go, so a shorter time applies to the
/// engine that is already warm.
pub(in super::super) fn set_webview_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    commit_config(|config| {
        config.webview_idle = idle;
    });
}

/// How long the viewer a model was hovered by is kept after that hover ends, from
/// `Engine → F3D TTL`. `Until away` is the AFK timer and a time is the engine's own
/// (see `EngineIdle`).
///
/// Nothing is rebuilt here either, and nothing has to be: the preview tick reads the setting every
/// time it asks whether a warm model is due to be let go, so a shorter time applies to the viewer
/// that is already kept and `Never (0 seconds)` ends it on the next pass (see `hover_film_host`).
pub(in super::super) fn set_f3d_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    commit_config(|config| {
        config.f3d_idle = idle;
    });
}

/// How long the LibreOffice engine is kept after the last page it drew, from
/// `Engine → LibreOffice TTL`. `Until away` is the AFK timer and a time is the engine's own
/// (see `EngineIdle`).
///
/// Nothing is rebuilt here either, and nothing has to be: the engine thread reads the setting
/// every second while it waits for documents, so a shorter time applies to the engine that is
/// already running, and `Never (0 seconds)` — the bottom of the list — lets go of one within the
/// second. An engine that is kept is a process this app holds and ends itself; a setting of
/// `Indefinitely` keeps it for the rest of the run (see `libreoffice_render`).
pub(in super::super) fn set_libreoffice_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    commit_config(|config| {
        config.libreoffice_idle = idle;
    });
}

/// How long Explorer may be out of reach before an engine set to `Until away` is let go.
///
/// Nothing is rebuilt here and nothing on screen changes for the same reason the idle times
/// rebuild nothing: an engine set to `Until away` reads this on the look it already takes
/// — the Office worker twice a second, the engine thread once a second, the browser every
/// quarter of one while it is up — so a shorter time lets go of an engine that is already
/// warm, and a longer one keeps an engine the next look would have let go of.
pub(in super::super) fn set_afk_timer(index: u16) {
    let Some(seconds) = afk_timer_secs_at(index) else {
        return;
    };

    commit_config(|config| {
        config.afk_timer_seconds = seconds;
    });
}

/// The size an item of the `Cache` submenu stands for, by the position it was
/// listed at and the list its own cache is offered in. An id past the last size
/// the menu offered is one that is not there.
fn cache_size_at(index: u16, choices: &[u32]) -> Option<u32> {
    choices.get(index as usize).copied()
}

/// How much memory the decoded-image cache may hold.
///
/// A preview on screen is not drawn from the cache but from the frame it was
/// loaded as, so nothing on screen changes — what changes is how much is freed, and
/// a smaller size frees it now rather than at the next decode.
pub(in super::super) fn set_image_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index, &IMAGE_CACHE_SIZE_CHOICES_MB) else {
        return;
    };

    publish_config_change(
        |config| {
            config.image_cache_mb = sanitize_image_cache_mb(megabytes);
        },
        &[PreviewMessage::TrimCaches],
    );
}

/// How much of what an engine drew may be kept, between hovers.
///
/// Unlike the image cache beside it this does not switch anything off: a page is drawn for the
/// hover that asks for it whatever the size, and a size of nothing means it is given up when
/// that hover ends. So nothing is rebuilt here either — the next hover answers for itself.
pub(in super::super) fn set_document_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index, &DISK_CACHE_SIZE_CHOICES_MB) else {
        return;
    };

    publish_config_change(
        |config| {
            config.document_cache_mb = sanitize_document_cache_mb(megabytes);
        },
        &[PreviewMessage::TrimCaches],
    );
}

/// How much of what the image converter developed may be kept, between hovers.
///
/// The same shape as the document cache beside it, and for the same reason: a picture is
/// developed for the hover that asks for it whatever the size, and a size of nothing means the
/// hover after it pays for the development again. Nothing is rebuilt here either — the pages are
/// read by the hover that wants one, so what a size does is bound what is left for it to read.
pub(in super::super) fn set_image_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index, &DISK_CACHE_SIZE_CHOICES_MB) else {
        return;
    };

    publish_config_change(
        |config| {
            config.image_disk_cache_mb = sanitize_image_disk_cache_mb(megabytes);
        },
        &[PreviewMessage::TrimCaches],
    );
}

/// How much of what a film's own subtitle tracks were copied into may be kept, between hovers.
///
/// The same shape as the two beside it and for the same reason: a film is copied for the hover
/// that asks for it whatever the size, and a size of nothing means nothing is copied at all —
/// such a hover is answered without subtitles rather than with the whole film streamed for
/// them, which is the answer the first hover of a film gets anyway while its copy is coming
/// (see `subtitle_files`).
pub(in super::super) fn set_general_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index, &DISK_CACHE_SIZE_CHOICES_MB) else {
        return;
    };

    commit_config(|config| {
        config.general_disk_cache_mb = sanitize_general_disk_cache_mb(megabytes);
    });

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

    commit_config(|config| {
        config.decode_budget_gb = sanitize_decode_budget_gb(gigabytes);
    });
}

/// The room one pinned picture's held run of frames may take, in megabytes.
///
/// Nothing on screen is rebuilt by it: every ask for the run reads the setting as it runs, so the
/// next pin is answered under the new room whatever it is — a smaller one simply holds its frames
/// at a smaller size than a larger one did (see `pin_index`).
pub(in super::super) fn set_animated_webp_budget_mb(index: u16) {
    let Some(megabytes) = ANIMATED_WEBP_BUDGET_CHOICES_MB.get(index as usize).copied() else {
        return;
    };

    commit_config(|config| {
        config.animated_webp_budget_mb = sanitize_animated_webp_budget_mb(megabytes);
    });
}

/// What the caches that are not stacks of decoded frames may hold between hovers, in megabytes —
/// the parsed pages, the archive listings and the one picture the image converter develops.
///
/// Unlike the two budgets beside it this one *is* told to the viewer, because a figure lowered has
/// to bite at the moment it is set rather than at whatever parse happens to pass through the caches
/// next: what is over the new size goes now, the way a smaller `image_cache_mb` does (see
/// `PreviewMessage::TrimCaches` and `others_cache::trim`).
pub(in super::super) fn set_others_cache_mb(index: u16) {
    let Some(megabytes) = OTHERS_CACHE_SIZE_CHOICES_MB.get(index as usize).copied() else {
        return;
    };

    publish_config_change(
        |config| {
            config.others_cache_mb = sanitize_others_cache_mb(megabytes);
        },
        &[PreviewMessage::TrimCaches],
    );
}

/// A text preview's own keys and pointer are the pin's business now: the row that used to
/// switch full mode on for every hover of a text file is gone, and pinning one with the pin key
/// is what brings it up as something to work in (see `current_text_options`).
pub(in super::super) fn set_text_font_scale(percent: u32) {
    let percent = sanitize_text_font_scale_percent(percent);
    publish_config_change(
        |config| {
            config.text_font_scale_percent = percent;
        },
        &[
            PreviewMessage::TrayTextSize { percent },
            PreviewMessage::Refresh,
        ],
    );
}

pub(in super::super) fn set_markdown_mode(mode: MarkdownMode) {
    publish_config_change(
        |config| {
            config.markdown_mode = mode;
        },
        &[PreviewMessage::Refresh],
    );
}

/// Which files the previous/next buttons on a pin's own caption step through: every file this
/// build can preview, or only those of the kind of thing the pinned file is.
///
/// Nothing on screen changes and nothing is rebuilt. The walk is read per step rather than
/// held, and what it is made of is a question answered off the folder when a button is
/// pressed — so the next step of a pin already up is the first one under the new setting,
/// and a pin with no folder of its own to walk is not a walk at all either way.
/// Which files a walk anchored on a picture steps through (see `Navigation → Image`).
pub(in super::super) fn set_pin_nav_image(mode: PinNavFileTypes) {
    commit_config(|config| {
        config.pin_nav_image = mode;
    });
}

/// Which files a walk anchored on a video steps through (see `Navigation → Video`).
pub(in super::super) fn set_pin_nav_video(mode: PinNavFileTypes) {
    commit_config(|config| {
        config.pin_nav_video = mode;
    });
}

/// Which files a walk anchored on a sound steps through (see `Navigation → Audio`).
pub(in super::super) fn set_pin_nav_audio(mode: PinNavFileTypes) {
    commit_config(|config| {
        config.pin_nav_audio = mode;
    });
}

/// Which files a walk anchored on any other file steps through (see `Navigation → Others`).
pub(in super::super) fn set_pin_nav_others(mode: PinNavFileTypes) {
    commit_config(|config| {
        config.pin_nav_others = mode;
    });
}

/// Where a pinned window's minimize puts it: the round bubble on the desktop, or the app's own
/// button in the taskbar.
///
/// Nothing on screen changes and nothing is rebuilt. The answer is read the moment the minimize
/// button is pressed, so a change made while a pin is up is the next minimize's — a pin already
/// away is left where it was put (see `collapse_pin`).
pub(in super::super) fn set_pin_minimize_to(target: PinMinimizeTo) {
    commit_config(|config| {
        config.pin_minimize_to = target;
    });
}

pub(in super::super) fn set_pin_scaling_mode(mode: PinScalingMode) {
    commit_config(|config| {
        config.pin_scaling_mode = mode;
    });
}

pub(in super::super) fn set_video_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    publish_config_change(
        |config| {
            config.video_volume = volume;
        },
        &[PreviewMessage::TrayVolume { audio: false }],
    );
}

/// A level of the `Volume → Audio` submenu, by the position it was listed at: the volume a
/// sound file is played at, which is the setting a card is drawn against as well as the one its
/// player is started with.
pub(in super::super) fn set_audio_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    publish_config_change(
        |config| {
            config.audio_volume = volume;
        },
        &[PreviewMessage::TrayVolume { audio: true }],
    );
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

    commit_config(|config| {
        config.audio_seek = seek;
    });
}

/// The way an item of the `Volume → Audio Seek` submenu stands for, by the position it was
/// listed at. An id past the last way the menu offered is one that is not there.
pub(in super::super) fn audio_seek_at(index: u16) -> Option<AudioSeek> {
    AUDIO_SEEK_CHOICES.get(index as usize).copied()
}

/// The `Volume → Pin Mode Audio Seek` submenu's command: where a
/// *pinned* sound starts, which is the pin's own setting rather
/// than the hover's.
///
/// Nothing on screen is rebuilt: a sound already playing is left
/// where it is, and what a click here changes is where the *next*
/// pinned sound starts — the same bargain the hover's `Audio Seek`
/// beside it makes.
pub(in super::super) fn set_pin_mode_audio_seek(index: u16) {
    let Some(seek) = audio_seek_at(index) else {
        return;
    };

    commit_config(|config| {
        config.pin_mode_audio_seek = seek;
    });
}

/// What a pinned film or sound does at the end of its file, by the position it was
/// listed at: the film's own end of file or the sound's, written into the pin's setting
/// rather than the hover's loop.
///
/// Nothing on screen is rebuilt: a file already playing is left where it is, and what
/// a click here changes is what the *next* end does — stop there, go round, or step on
/// to the next file — the same bargain the seek beside it makes.
pub(in super::super) fn set_pin_end_of_file(index: u16, video: bool) {
    let Some(end_of_file) = end_of_file_at(index) else {
        return;
    };

    commit_config(|config| {
        if video {
            config.pin_mode_video_end_of_file = end_of_file;
        } else {
            config.pin_mode_audio_end_of_file = end_of_file;
        }
    });
}

/// Whether a hovered film or sound goes round at the end of it, which is what the
/// `Loop` row throws the other way round.
///
/// Nothing on screen is rebuilt: a file already playing is left where it is, and what a
/// click here changes is what the *next* end does.
pub(in super::super) fn toggle_loop(video: bool) {
    commit_config(|config| {
        if video {
            config.video_loop = !config.video_loop;
        } else {
            config.audio_loop = !config.audio_loop;
        }
    });
}

/// The kind of pinned walk a `Shuffle` row asks about: the three anchors that have a
/// shuffle of their own, one setting apiece, because the question is the same question asked
/// of three different things to step through (see `PinStep`).
pub(in super::super) enum PinShuffle {
    Video,
    Audio,
    Image,
}

/// Whether a *pinned* film, sound or picture takes the next file of its walk at random rather
/// than the one beside it, which is what the `Shuffle` row throws the other way round.
///
/// Nothing on screen is rebuilt either: a walk already in hand carries the shuffle it was
/// planned with, and what a click here changes is the walk the next step is planned as —
/// whether that step is a caption's next press, a forward key, or the end of a file the
/// kind's own `End of File` row calls `Jump`. A step back is the file the walk came from,
/// shuffle or no shuffle: the order a shuffle picks the *next* file in is the only thing it
/// is (see `PinStep::step`).
pub(in super::super) fn toggle_pin_shuffle(kind: PinShuffle) {
    commit_config(|config| match kind {
        PinShuffle::Video => config.pin_mode_video_shuffle = !config.pin_mode_video_shuffle,
        PinShuffle::Audio => config.pin_mode_audio_shuffle = !config.pin_mode_audio_shuffle,
        PinShuffle::Image => config.pin_mode_image_shuffle = !config.pin_mode_image_shuffle,
    });
}

/// Whether a pinned sound's playlist band puts its camera on the file its walk stepped onto — the
/// `Camera Follows on Seek` row.
///
/// Nothing on screen moves for the click itself: the row decides what the *next* step costs the
/// band's scroll, and the band already up is where the file the hand hears already is. Off, the
/// rows are the same list of the same files and the pin the same walk; what is given up is the
/// scroll that goes looking for the row the step landed on (see `pin_playlist`).
pub(in super::super) fn toggle_pin_camera_follows() {
    commit_config(|config| {
        config.pin_mode_audio_camera_follows = !config.pin_mode_audio_camera_follows;
    });
}

/// Whether a sound taken up pinned comes up with its playlist band open — the `Always Open
/// Playlist` row.
///
/// Nothing on screen moves for the click itself: the setting is read at the next take-up, and the
/// band that a pin is already carrying is the one the hand left (see `pin_playlist::open_for`).
/// A band is only ever a pin's, so this is the one road that opens one with no press behind it.
pub(in super::super) fn toggle_pin_playlist_always_open() {
    commit_config(|config| {
        config.pin_mode_audio_playlist_always_open = !config.pin_mode_audio_playlist_always_open;
    });
}

pub(in super::super) fn set_follow_cursor(follow: bool) {
    commit_config(|config| {
        config.follow_cursor = follow;
    });
}

pub(in super::super) fn toggle_bypass_placement() {
    commit_config(|config| {
        config.bypass_placement = !config.bypass_placement;
    });
}

pub(in super::super) fn toggle_custom_preview_enabled() {
    publish_config_change(
        |config| {
            config.custom_preview_enabled = !config.custom_preview_enabled;
        },
        &[PreviewMessage::Refresh],
    );
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

    commit_config(|config| {
        config.avoid_mode = mode;
    });
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

    commit_config(|config| {
        config.tick_ms = sanitize_tick_ms(tick_ms);
    });
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

    commit_config(|config| {
        config.preview_scale = scale;
    });
}

/// The same for a video, at the share `video_scale` names: the frame that stands in for
/// one is a bitmap like a picture, so the shares its submenu offers are the picture's
/// shares — of the display's fitted size and of the frame's own size both — and the
/// setting written is the video's own.
pub(in super::super) fn set_video_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.video_scale = scale;
    });
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

    commit_config(|config| {
        config.animated_scale = scale;
    });
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

    commit_config(|config| {
        config.ebook_scale = scale;
    });
}

/// How much of the display a page of the `Document` kind is shown over, by the position the
/// item was listed at — the one setting behind both halves of the kind: a page an Office
/// document's own application exported, and a page the render engine drew.
pub(in super::super) fn set_document_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.document_scale = scale;
    });
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

    commit_config(|config| {
        config.audio_scale = scale;
    });
}

/// How much of the display a font specimen is drawn over, by the position the item was
/// listed at. The same rule as the three beside it: the next hover, not the one that is up.
pub(in super::super) fn set_font_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.font_scale = scale;
    });
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

    commit_config(|config| {
        config.design_scale = scale;
    });
}

/// How much of the display a vector drawing is replayed over, by the position the item was
/// listed at — the same rule as the documents beside it: the share is of the display, and
/// it applies to the next hover rather than resizing a preview that is already up.
pub(in super::super) fn set_vector_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.vector_scale = scale;
    });
}

/// How much of the display a model is drawn over, by the position the item was listed at.
/// The same rule as the documents beside it: a scene has no size of its own — the picture
/// an installed f3d draws of one is drawn at whatever box it is asked for — so the share is
/// of the display, and it applies to the next hover rather than resizing a preview that is
/// already up.
pub(in super::super) fn set_model_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.model_scale = scale;
    });
}

/// How much of the display a page of text is measured in, by the position the item was listed
/// at. The same rule as the documents beside it: the box a text page is measured for is part of
/// the placement that was made when the preview was opened, so this applies to the next hover
/// rather than resizing the page that is already up.
pub(in super::super) fn set_text_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    commit_config(|config| {
        config.text_scale = scale;
    });
}

/// The share an item of the `Image`, `Video` or `Animated`
/// submenu stands for, by the position it was listed at — a share of
/// the display's fitted size or of a bitmap's own size, which is the
/// one question the three submenus ask with one table. An id past the
/// last choice the menu offered is one that is not there.
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

    commit_config(|config| {
        config.hover_delay_ms = delay_ms;
    });
}

/// How long the same file waits before a preview of it is put up again, by the position
/// its item was listed at.
pub(in super::super) fn set_same_file_rehover_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    commit_config(|config| {
        config.same_file_rehover_delay_ms = delay_ms;
    });
}

/// How long the pointer must be still before a preview may open for anything, by the
/// position its item was listed at.
pub(in super::super) fn set_settling_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    commit_config(|config| {
        config.settling_delay_ms = delay_ms;
    });
}

/// A page, opened in the browser the user already has: the same call the release page is
/// opened with, and nothing is fetched or run here — what becomes of the page is the
/// browser's own business (see `updates::open_release_page`).
pub(in super::super) fn open_link(url: &str) {
    // Through the one road this app hands anything over: on an elevated run the browser is started
    // by the Shell, so the page comes up in the user's own browser rather than in an administrator's
    // (see `app::hand_off`).
    crate::app::hand_off::open(url, None);
}

pub(in super::super) fn open_config_file() {
    if let Ok(config) = CONFIG.lock() {
        config.save();
    }

    if let Some(path) = crate::config::config::AppConfig::config_path() {
        // The file is handed to whatever edits an `.ini`, which on an elevated run is the user's own
        // editor rather than one running as an administrator — the same rule every hand-off of this
        // app's follows (see `app::hand_off`).
        crate::app::hand_off::open(&crate::paths::plain_path(&path), None);
    }
}
