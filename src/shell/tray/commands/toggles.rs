//! The switches: what a row of the tray turns the other way round.
//!
//! A toggle writes a setting the same way a setter in `setters` beside it does, and tells
//! whatever else reads it the same way too. The one thing that differs is where the answer
//! comes from: a toggle asks the setting what it is before putting it the other way, and
//! what it asks is not always the file. `toggle_startup` asks the registry, because an entry
//! can be taken away from outside this app and a flip that trusted the record would then need
//! two clicks to put back.
//!
//! A switch read per hover or per frame needs nothing said at all — the next hover answers for
//! itself — and the ones that do rebuild are the backdrops and the kinds of preview, where
//! what stands behind a picture, or the engine behind a kind of it, is part of what is on
//! screen. The lookups the setters write through live in `setters`; the window proc in
//! `event_loop` reaches both through their parent.
//!
//! Every function here is `pub(in super::super)` rather than `pub(super)`: the parent re-exports
//! each of them under the name the window proc already knows it by, and a re-export cannot be
//! wider than the thing it names. `super::super` is `tray`, which is exactly as wide as the
//! `pub(super)` these were written as when they all lived in one file.

use crate::app::startup;
use crate::config::config::PreviewType;
use crate::engines::office_render;
use crate::engines::webview_preview;
use crate::ui::preview_window::{
    refresh_pin, refresh_preview, refresh_preview_types, refresh_render_html,
};
use crate::CONFIG;

pub(in super::super) fn toggle_startup() {
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

pub(in super::super) fn toggle_preview_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.preview_enabled = !config.preview_enabled;
        config.save();
    }
}

/// Whether the pin key is watched is a setting rather than a view of one, and two
/// things have to be told about a change to it: the hook procedure that watches the
/// key reads a number rather than the configuration, and a preview that is pinned
/// when the feature is switched off is a window nothing would ever take down again.
pub(in super::super) fn toggle_pin_enabled() {
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
pub(in super::super) fn toggle_pin_update_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_enabled = !config.pin_update_enabled;
        config.save();
    }
}

/// And whether the pointer's own hover is one of the ways it is told about one, which is read
/// in the same place and changes nothing on screen either (see `pin_update_on_hover`).
pub(in super::super) fn toggle_pin_update_on_hover() {
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
pub(in super::super) fn toggle_pin_pause_video() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_pause_video = !config.pin_pause_video;
        config.save();
    }
}

/// The sound's half of the pair above, read on the same tick and acting the same way.
pub(in super::super) fn toggle_pin_pause_audio() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_pause_audio = !config.pin_pause_audio;
        config.save();
    }
}

/// Whether the key is watched is a setting rather than a view of one, so the preview
/// on screen is rebuilt: switching the key off while it is held lets a preview
/// through, and switching it back on while it is held takes one away.
pub(in super::super) fn toggle_trigger_key_enabled() {
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
pub(in super::super) fn toggle_trigger_key_affect_pin_mode() {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_affect_pin_mode = !config.trigger_key_affect_pin_mode;
        config.save();
    }
}

/// Whether the keyboard driving Explorer holds a parked pointer back is a setting rather
/// than a view of one: nothing on screen is rebuilt and nothing is taken down — the switch
/// is read by the hook on its next tick — so a preview that is up when it is thrown is left
/// where it is.
pub(in super::super) fn toggle_prioritize_keyboard() {
    if let Ok(mut config) = CONFIG.lock() {
        config.prioritize_keyboard = !config.prioritize_keyboard;
        config.save();
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
pub(in super::super) fn toggle_preview_type(kind: PreviewType) {
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

/// Whether one of the three engines is kept whatever the user is doing, from the
/// `Persistent` toggle at the top of its TTL submenu.
///
/// The toggle is which submenu it heads rather than which engine it names, because the three
/// are listed in one order and the ids are handed out in it. Nothing is rebuilt here either:
/// both sides of the setting are read live by the engine that decides with them, so a toggle
/// turned on keeps the engine the next look would have let go of, and one turned off lets go
/// of an engine that is already up as soon as the AFK timer says it may.
pub(in super::super) fn toggle_engine_persistent(index: u16) {
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
pub(in super::super) fn toggle_video_hw_accel() {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_hw_accel = !config.video_hw_accel;
        config.save();
    }

    crate::ui::preview_window::forget_video_hw_accel_answer();
}

/// Whether an explicitly chosen video engine falls through to the others when it cannot play a
/// file, from the `Fallback` row at the top of `Engine -> Select Engine -> Video`.
///
/// Nothing on screen changes: the switch is read by the router on the next hover, and what the
/// geometry probe holds is given up for the reason the choice moving gives it up — the answer
/// the next hover wants may be another engine's (see `forget_video_geometry`).
pub(in super::super) fn toggle_video_engine_fallback() {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_engine_fallback = !config.video_engine_fallback;
        config.save();
    }
    crate::ui::preview_window::forget_video_geometry();
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
pub(in super::super) fn toggle_render_html() {
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

/// The switch above a sound's levels: whether a file's measured loudness is brought to one level
/// before it is played, which is read where a player is started the way the level
/// beside it is — nothing on screen is rebuilt, and a sound already playing is left where it is.
///
/// It is the file that is measured rather than the playing that is re-scaled, so a file switched
/// on for while it was already known is measured off the tick, and what that measurement is for is
/// the hover after the one that asked for it (see `spawn_gain_scan`).
pub(in super::super) fn toggle_normalize_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.normalize_volume = !config.normalize_volume;
        config.save();
    }
}

/// And the video's own switch, which is the same question asked about a film's soundtrack: read
/// where its player is started, like the level beside it, and off where the app starts — a film is
/// looked at, and a measurement of its audio is a decode of the film.
pub(in super::super) fn toggle_normalize_video_volume() {
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
pub(in super::super) fn toggle_remember_audio_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.remember_audio_volume = !config.remember_audio_volume;
        config.save();
    }
}

/// And the video's own, which is the same switch asked about a soundtrack and off where the app
/// starts, like the video's level it decides.
pub(in super::super) fn toggle_remember_video_volume() {
    if let Ok(mut config) = CONFIG.lock() {
        config.remember_video_volume = !config.remember_video_volume;
        config.save();
    }
}
