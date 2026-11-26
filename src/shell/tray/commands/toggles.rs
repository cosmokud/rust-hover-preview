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

use crate::app::dialogs;
use crate::app::seam::{
    refresh_pin, refresh_preview, refresh_preview_types, refresh_render_html, send_preview,
    the_rings_are_the_road, PreviewMessage,
};
use crate::app::startup;
use crate::config::config::PreviewType;
use crate::config::{commit_config, publish_config_change};
use std::sync::atomic::Ordering;
use std::time::{Duration, Instant};

pub(in super::super) fn toggle_startup() {
    // What is flipped is the machine, and what is read to decide which way to flip is the machine
    // too: an entry or a task can be taken away from outside this app, and a toggle that trusted the
    // configuration would then need two clicks to put back — the first turning off something already
    // off. The configuration is the record of the choice, written to agree with what was just done.
    let enable = !startup::is_startup_enabled();

    let outcome = if enable {
        startup::enable_startup()
    } else {
        startup::disable_startup()
    };

    // **What is written down is what the machine holds, not what was asked for.** A write the
    // registry refused — a `Run` key a policy or another tool holds, a hive this process may not
    // open — used to leave `run_at_startup=true` in the file beside no entry at all, which is the
    // one failure that reads as the app forgetting the choice it was given. The row says so as
    // well, since a switch that flips back by itself is a switch nobody trusts.
    let (held, elevated) = {
        let held = startup::machine();
        (held.started().runs(), held.elevates())
    };
    commit_config(|config| {
        config.run_at_startup = held;
        config.run_as_administrator = elevated;
    });

    match outcome {
        // The logon that is to start this app has to start it elevated, which is a scheduled task,
        // and a task is a machine write this token may not make. What an unelevated run does about
        // that is hand the write to an elevated copy of itself (see `hand_the_write_to_an_elevated_copy`).
        startup::Outcome::NeedsElevation | startup::Outcome::TaskRefused => {
            hand_the_write_to_an_elevated_copy(if enable {
                startup::Applied::TurnOn
            } else {
                startup::Applied::TurnOff
            });
        }
        other if !other.settled() => dialogs::report_startup_failure(other),
        _ => {}
    }
}

/// Hand the one write an unelevated process may not make — registering or removing the scheduled
/// task that starts this app **elevated** at a logon — to an elevated copy of this app, and then
/// become what that write leaves behind rather than keeping two of them.
///
/// **The elevated copy is started first and this one leaves second**, which is the order the single
/// instance mutex allows: a second tray begun while this one still holds the mutex would exit, and
/// the elevated copy would then be the app's only process — which is exactly what the row asked for
/// and what the tray the user is looking at would no longer be. So the helper is begun (it writes
/// one thing and exits, and takes no mutex), and this process then starts whatever the machine now
/// holds and ends itself: the task it just registered for a turn-on, or the ordinary `Run` value it
/// just wrote for a turn-off. What the user sees either way is the tray icon coming back as the copy
/// a logon would start.
fn hand_the_write_to_an_elevated_copy(apply: startup::Applied) {
    let helper = match startup::raise_to_elevation(apply) {
        startup::Elevated::Started(helper) => helper,
        startup::Elevated::Refused => {
            // The prompt was answered with no, or no elevated copy could be started at all: the row
            // says the machine did not move, and nothing of this process's changes.
            dialogs::report_startup_failure(startup::Outcome::NotWritten {
                held: startup::is_startup_enabled(),
            });
            return;
        }
    };

    // **The wait ends with the helper, not with a clock.** Waiting is what keeps this process alive
    // until the machine holds what was asked for, so that the tray the user is looking at does not
    // blink out of the task list before there is anything to come back to — but what is waited on is
    // the helper's own process, which exits the moment its one write is behind it. A poll that ran
    // to a deadline instead was this thread frozen for the whole of it, and a write the machine
    // refused cost the user half a minute of a tray that answered no click.
    let written = match helper {
        Some(helper) => startup::wait_for_the_helper(helper) && the_machine_took_the_write(apply),
        // The one Shell that begins the copy without leaving a handle leaves nothing to wait on, and
        // the machine's own state is then what is asked, until it moves.
        None => wait_for_the_write(apply),
    };

    if !written {
        dialogs::report_startup_failure(startup::Outcome::TaskRefused);
        return;
    }

    hand_over_to_the_successor();
}

/// Wait until the machine holds what the elevated copy was begun to write, for the Shell that began
/// it without leaving a process to wait on.
///
/// It is the poll the process handle replaced, kept for the one case where there is no handle: the
/// helper is one COM call and an exit, so this is a short deadline on a machine under a scanner's
/// thumb rather than an expectation.
fn wait_for_the_write(apply: startup::Applied) -> bool {
    let deadline = Instant::now() + Duration::from_secs(10);

    while Instant::now() < deadline {
        if the_machine_took_the_write(apply) {
            return true;
        }

        std::thread::sleep(Duration::from_millis(100));
    }

    the_machine_took_the_write(apply)
}

/// Whether the machine now holds what the elevated copy was begun to write: the task that starts this
/// copy at the highest run level for a turn-on, and the ordinary entry with no task beside it for a
/// turn-off.
///
/// The helper says nothing back beyond exiting, so what it did is read off the machine rather than
/// taken on trust — the same two reads the tray's rows are drawn from, asked once the process that
/// was to make the write is gone.
fn the_machine_took_the_write(apply: startup::Applied) -> bool {
    match apply {
        startup::Applied::TurnOn => {
            let exe = std::env::current_exe().ok();
            startup::task_facts()
                .starts_this_copy(exe.as_deref().unwrap_or(std::path::Path::new("")))
        }
        startup::Applied::TurnOff => {
            let machine = startup::machine();
            machine.started().runs() && !machine.elevates()
        }
    }
}

/// Start the copy the machine now holds and end this one — the half of a hand-over that both
/// directions share.
///
/// **The single-instance name is given up before the successor is started, not after.** A copy begun
/// while this process still holds it would find the app already running and exit, and the tray that
/// was meant to come back would be no tray at all. So the name goes first (see
/// `single_instance::let_go_early`), the successor is asked for, and only then does this process
/// leave.
///
/// **And a name given up for a successor that did not come is taken back.** The machine is right —
/// the task or the entry was written — and what failed is only this copy's replacement: a tray that
/// stayed as the app's only window without being the app's only instance would be one a second copy
/// could stand beside, two trays over one name nobody holds. Taking it back leaves the app exactly
/// as it was, and the copy the machine now names is the one the next logon starts.
fn hand_over_to_the_successor() {
    crate::app::single_instance::let_go_early();

    let started = {
        let machine = startup::machine();

        match machine.started() {
            // The elevated answer: the record Windows itself will run at the next logon.
            startup::Started::Task => startup::start_the_task(),
            // The ordinary one: the value a logon runs, started through the Shell's own token so
            // that it is the unelevated copy rather than a second elevated one.
            _ => startup::start_through_the_entry(),
        }
    };

    // A successor that did not start leaves this copy as the app: the name it gave up is taken
    // again, and only a copy that got there first — which is then the app — keeps it.
    if !started && crate::app::single_instance::acquire() {
        return;
    }

    crate::RUNNING.store(false, Ordering::SeqCst);
    crate::shell::tray::leave_the_tray();
}

/// Whether the logon that starts this app starts it **elevated** — the row under `Run at Startup`,
/// and the one thing in this app that asks Windows for a token it does not have.
///
/// Turning it on registers a scheduled task with the highest run level, which is a machine write
/// that needs the elevated token, so an unelevated run hands it to an elevated copy of itself
/// exactly as the startup row does (see `hand_the_write_to_an_elevated_copy`) and comes back elevated.
///
/// **Turning it off is the mirror of turning it on, all the way down.** The mark on the exe is
/// cleared first (a `RUNASADMIN` layer is the same request, and Windows starts no `Run` value whose
/// exe carries it), then the machine is left with the ordinary entry, and *then* the tray restarts
/// itself through that entry so that the icon that comes back is the unelevated copy — the one the
/// next logon will start. Turning it off without that restart left an administrator's tray running
/// under a row that said otherwise, which is the state this is written to avoid.
pub(in super::super) fn toggle_run_as_administrator() {
    let machine = startup::machine();
    // The mark on the file counts as the row being on: it is the same request, made in Windows'
    // own properties, and it is what the row says when it is read back from the machine.
    let on = !(machine.elevates() || startup::compat_elevation_is_set());

    // **The mark is cleared before the file is written, whichever way the row is going.** A
    // `RUNASADMIN` layer on the exe is the same request as this row and Winsows starts no `Run`
    // value whose exe carries it, so the two would otherwise argue: turning the row on would leave a
    // mark that the mechanism underneath no longer needed, and turning it off would leave a machine
    // asking for elevation with nothing that starts it. Off, it is the thing being turned off; on,
    // it is replaced by the task, which is the answer the row keeps.
    let mark_gone = startup::clear_compat_elevation();

    commit_config(|config| {
        config.run_as_administrator = on;
    });

    // Nothing at a logon is nothing to elevate yet: the choice is kept in the file, and the row
    // above writes the task — and raises the one prompt — when it is turned on (see `app::startup`).
    let wanted = crate::CONFIG
        .lock()
        .map(|config| config.run_at_startup)
        .unwrap_or(false);

    if !wanted {
        return;
    }

    // **The row above decides what a logon does, and this one decides how.** On, the mechanism is
    // the task; off, it is the ordinary `Run` value — because `run_at_startup` is still on, and a
    // startup that was left with neither would be a switch that turned off more than its own row.
    // `enable_startup` is what writes the value and takes the task away in one pass, so both
    // directions go through it while the row above says yes; only a startup that is not wanted at
    // all is the call that removes both.
    let outcome = if on || wanted {
        startup::enable_startup()
    } else {
        startup::disable_startup()
    };

    match outcome {
        startup::Outcome::NeedsElevation | startup::Outcome::TaskRefused => {
            hand_the_write_to_an_elevated_copy(if on {
                startup::Applied::TurnOn
            } else {
                startup::Applied::TurnOff
            });
            return;
        }
        other if !other.settled() => {
            dialogs::report_startup_failure(other);
            return;
        }
        _ => {}
    }

    if !mark_gone {
        // The row asked for the mark to be gone and the machine kept it: what a logon does is not
        // what the row says it does, and that is said out loud rather than left to a checkmark.
        dialogs::report_startup_failure(startup::Outcome::NotWritten {
            held: startup::is_startup_enabled(),
        });
        return;
    }

    // And the hand-over, in the one direction that needs it: turning the elevated start off leaves
    // this process with an administrator's token, and an unelevated copy is what a logon will now
    // start — so the tray is replaced by that copy rather than left standing as the old state.
    if !on {
        // The entry is written, so the copy to be looking at is the one it names: this one hands the
        // single-instance name over, asks for that copy and leaves, and the tray that comes back is
        // the unelevated one a logon starts.
        hand_over_to_the_successor();
    }
}

/// Whether the programs this app hands a file over to while it is elevated are started with this
/// process's own token — the row a user ticks after being warned about what that means.
///
/// Nothing is restarted by it: a hand-off is read at the moment a program is started, so the next
/// file opened takes the new answer and the programs already on screen keep the token they have.
pub(in super::super) fn toggle_elevate_hand_offs() {
    commit_config(|config| {
        config.elevate_hand_offs = !config.elevate_hand_offs;
    });
}

pub(crate) fn toggle_preview_enabled() {
    commit_config(|config| {
        config.preview_enabled = !config.preview_enabled;
    });
}

/// Whether the app asks GitHub for a newer release on its own: the `System → Check
/// for Updates` row, on where the app starts. Turning it on asks at once rather than
/// at the next opening of the menu or the next turn of the tray's timer, so the row
/// above the submenu reports what the check found without waiting for one; turning it
/// off asks nothing further, and the row is gone while there is nothing already on
/// offer to say.
pub(in super::super) fn toggle_check_for_updates() {
    let enabled = commit_config(|config| {
        config.check_for_updates = !config.check_for_updates;
        config.check_for_updates
    });

    if enabled {
        crate::app::updates::request_check();
    }
}

/// Whether an update a check finds is put on without asking: the `System → Silent Auto-Update`
/// row under the switch above it, off where the app starts. Turning it on while one is already on
/// offer puts that one on at once, so the switch is not one that waits for a later check to do
/// anything — and it waits for the screen to empty where a pin or a menu is on it, since nothing
/// of the user's may be closed by an install nobody asked for (see `updates::the_app_is_idle`).
/// Turning it off leaves an install already in flight alone, as a switch reading a file cannot
/// take back a program that has been started, and drops one it was owed.
pub(in super::super) fn toggle_silent_auto_update() {
    let enabled = commit_config(|config| {
        config.silent_auto_update = !config.silent_auto_update;
        config.silent_auto_update
    });

    if enabled {
        crate::app::updates::install_silently_in_background();
    }
}

/// Whether the pin key is watched is a setting rather than a view of one, and two
/// things have to be told about a change to it: the hook procedure that watches the
/// key reads a number rather than the configuration, and a preview that is pinned
/// when the feature is switched off is a window nothing would ever take down again.
pub(in super::super) fn toggle_pin_enabled() {
    commit_config(|config| {
        config.pin_enabled = !config.pin_enabled;
    });
    crate::shell::key_input::refresh();
    refresh_pin();
}

pub fn toggle_pin_always_on_top() {
    publish_config_change(
        |config| {
            config.pin_always_on_top = !config.pin_always_on_top;
        },
        &[PreviewMessage::PinChanged],
    );
}

/// Whether a preview is made where a **browse dialog** is showing is a setting rather than a view of
/// one, and it is read by the Explorer hook on its next tick like the pick switches above (see
/// `browse_dialog_preview_enabled`).
///
/// What it also does is tell the run's published fact, because the press road and the pin's own
/// placement may take no lock and answer for a dialog they find where neither the configuration nor
/// a tick has been read: a switch thrown off with a pin standing in a dialog would otherwise leave
/// the pin owing a front to a listing the app no longer answers for until the next tick noticed
/// (see `app::seam::BROWSE_DIALOG_PREVIEWS`).
pub(in super::super) fn toggle_browse_dialog_preview() {
    let enabled = commit_config(|config| {
        config.browse_dialog_preview_enabled = !config.browse_dialog_preview_enabled;
        config.browse_dialog_preview_enabled
    });

    crate::app::seam::publish_browse_dialog_previews(enabled);
    refresh_preview();
}

/// Whether a preview is made where **voidtools Everything** is showing its results is a setting rather
/// than a view of one, and it is read by the Explorer hook on its next tick like the dialog switch above
/// (see `everything_previews`).
///
/// What it also does is tell the run's published fact, because the press road and the pin's own
/// placement may take no lock and answer for an Everything window they find before any tick has been
/// read (see `app::seam::EVERYTHING_PREVIEWS`).
pub(in super::super) fn toggle_everything_previews() {
    let enabled = commit_config(|config| {
        config.everything_previews = !config.everything_previews;
        config.everything_previews
    });

    crate::app::seam::publish_everything_previews(enabled);
    refresh_preview();
}

/// Whether a preview is made where **Q-Dir** is showing its panes is a setting rather than a view of one,
/// and it is read by the Explorer hook on its next tick like the two above (see `q_dir_previews`).
///
/// What it also does is tell the run's published fact, because the press road and the pin's own placement
/// may take no lock and answer for a Q-Dir frame they find before any tick has been read (see
/// `app::seam::Q_DIR_PREVIEWS`).
pub(in super::super) fn toggle_q_dir_previews() {
    let enabled = commit_config(|config| {
        config.q_dir_previews = !config.q_dir_previews;
        config.q_dir_previews
    });

    crate::app::seam::publish_q_dir_previews(enabled);
    refresh_preview();
}

/// Whether a pin that is up is shown the file the user picks next is a setting rather than a
/// view of one, and it is the one switch here that changes nothing that is already on screen:
/// the pin keeps the file it is showing until the user picks another, and what the switch says
/// is read by the Explorer hook on its next tick (see `pin_update_enabled`).
pub(in super::super) fn toggle_pin_update_enabled() {
    commit_config(|config| {
        config.pin_update_enabled = !config.pin_update_enabled;
    });
}

/// And whether the pointer's own hover is one of the ways it is told about one, which is read
/// in the same place and changes nothing on screen either (see `pin_update_on_hover`).
pub(in super::super) fn toggle_pin_update_on_hover() {
    commit_config(|config| {
        config.pin_update_on_hover = !config.pin_update_on_hover;
    });
}

/// Whether an update also brings the pin in front of what it was behind, which is read by the
/// preview loop at the next pick and so changes nothing that is already on screen (see
/// `pin_update_move_to_top`).
pub(in super::super) fn toggle_pin_update_move_to_top() {
    commit_config(|config| {
        config.pin_update_move_to_top = !config.pin_update_move_to_top;
    });
}

/// Whether a pin collapsed into its bubble holds the video it is playing is a setting rather than
/// a view of one, and the side that acts on a change is the preview loop on its next tick: a film
/// that is running because the switch was off is held the moment it is switched on, and one that
/// is already held is left where it is until the pin is put back up, because a player started
/// beside a bubble would be a picture on screen next to it (see `settle_bubble_playback`).
pub(in super::super) fn toggle_pin_pause_video() {
    commit_config(|config| {
        config.pin_pause_video = !config.pin_pause_video;
    });
}

/// The sound's half of the pair above, read on the same tick and acting the same way.
pub(in super::super) fn toggle_pin_pause_audio() {
    commit_config(|config| {
        config.pin_pause_audio = !config.pin_pause_audio;
    });
}

/// Whether the key is watched is a setting rather than a view of one, so the preview
/// on screen is rebuilt: switching the key off while it is held lets a preview
/// through, and switching it back on while it is held takes one away.
pub(in super::super) fn toggle_trigger_key_enabled() {
    commit_config(|config| {
        config.trigger_key_enabled = !config.trigger_key_enabled;
    });
    // The hook's own chord table carries whether the key is watched at all, so the switch is
    // handed to it here or a key turned off would keep its latch (see `key_input::refresh`).
    crate::shell::key_input::refresh();
    refresh_preview();
}

/// Whether the key reaches a pin is a setting rather than a view of one, and nothing on screen
/// is rebuilt here: the hook reads the switch on its next tick, and a pin it is thrown under is
/// answered from the key's own state as that tick finds it — nothing of the pin belongs to this
/// thread.
pub(in super::super) fn toggle_trigger_key_affect_pin_mode() {
    commit_config(|config| {
        config.trigger_key_affect_pin_mode = !config.trigger_key_affect_pin_mode;
    });
}

/// Whether the keyboard driving Explorer holds a parked pointer back is a setting rather
/// than a view of one: nothing on screen is rebuilt and nothing is taken down — the switch
/// is read by the hook on its next tick — so a preview that is up when it is thrown is left
/// where it is.
pub(in super::super) fn toggle_prioritize_keyboard() {
    commit_config(|config| {
        config.prioritize_keyboard = !config.prioritize_keyboard;
    });
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
    let now_enabled = commit_config(|config| {
        let enabled = kind.enabled_in(config);
        kind.set_enabled_in(config, !enabled);
        !enabled
    });

    if !now_enabled {
        // An engine only one kind of preview is ever started for goes with that
        // kind's gate. The Office processes are asked for as a whole rather than
        // through the worker — a worker inside a call it cannot cut short would hold
        // one of them for good, and this thread may not wait on one — while the
        // browser needs nothing said to it at all: its own thread reads this gate and
        // lets the engine go. Both are the viewer's processes, ended where they run
        // (see `PreviewMessage::StopEngines`).
        if let PreviewType::Document = kind {
            send_preview(PreviewMessage::StopEngines);
        }

        // A page of HTML the browser is drawing rides the text kind's gate: switched off,
        // the page comes down with it, in the process drawing it (see
        // `refresh_render_html`).
        if let PreviewType::Text = kind {
            refresh_render_html();
        }
    }

    refresh_preview_types();
}

/// Whether a video is decoded on the graphics card, turned the other way round.
///
/// Nothing *already on screen* changes when this is switched: the setting is read when a player is
/// launched and a film already playing was launched with the other answer, so it is kept playing
/// rather than being relaunched under the user's feet. That is the same bargain every other switch
/// in this menu makes, and it is why the row is a question about the next preview rather than about
/// the one under the pointer.
///
/// It is mpv's switch and nobody else's, so the file is the whole of what this writes: there is no
/// answer of this app's own to give up with it. A machine whose `config.ini` says on and which has
/// no `mpv.exe` reads it as off all the same, which is what keeps every config written before the
/// row said so working (see `video_hw_accel_wanted`).
pub(in super::super) fn toggle_video_hw_accel() {
    commit_config(|config| {
        config.video_hw_accel = !config.video_hw_accel;
    });
}

/// Whether the away timer reaches the app's memory caches as well as its engines, from
/// `Performance → Memory → Respect Away Timer`.
///
/// **Nothing on screen changes and nothing is told.** The row is read where the flush is decided,
/// which is the viewer's tick, and the setting is read there live rather than captured — so the
/// first tick after the row is turned on is the first one that answers the new value, and a machine
/// that is away when it is turned on flushes on that tick (see `app::away_flush`). A refresh would
/// be a repaint of a preview for a setting that changes none of its pixels.
pub(in super::super) fn toggle_respect_away_timer() {
    commit_config(|config| {
        config.respect_away_timer = !config.respect_away_timer;
    });
}

/// Whether an explicitly chosen video engine falls through to the others when it cannot play a
/// file, from the `Fallback` row at the top of `Engine → Video Engine`.
///
/// **Nothing on screen changes: the switch is half of the engine choice**, read with it where a
/// preview is begun, so a film already playing keeps the fall-through it was started under —
/// the switch was already *said* to apply on the next hover and was not, because the router
/// asked it again on every question a resize, a gesture or a step put to it. It is one move on
/// the choice of the file, and what the geometry probe holds is given up where that choice is
/// read rather than here (see `engines_in_effect`, `forget_video_geometry`).
pub(in super::super) fn toggle_video_engine_fallback() {
    commit_config(|config| {
        config.video_engine_fallback = !config.video_engine_fallback;
    });
}

/// Whether a film's own subtitle tracks are probed for and copied out at all,
/// from the `Subtitles` row under `Preview Settings → Video`.
///
/// Nothing on screen is rebuilt: the switch is read where a probe runs and
/// where a player is started, on the next hover. What the probe holds is
/// given up for the reason the engine choice moving gives it up — a geometry
/// answered while subtitles were wanted carries subtitle streams, a sidecar
/// and a copy that a probe under the new answer would not have (see
/// `forget_video_geometry`), and an extraction running for a film is dropped
/// rather than left to read a film nobody will draw subtitles from (see
/// `keep_extraction_for`).
pub(in super::super) fn toggle_video_subtitles() {
    publish_config_change(
        |config| {
            config.video_subtitles = !config.video_subtitles;
        },
        &[
            PreviewMessage::ForgetVideoGeometry,
            PreviewMessage::KeepExtraction { path: None },
        ],
    );
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
    let turned_off = commit_config(|config| {
        config.render_html = !config.render_html;
        !config.render_html
    });

    // The direction that leaves the engine is the one that has to say so: a preview
    // rebuilt from its hover is what a painted page owes (see `refresh_preview`), while
    // a page the engine is drawing for a hover is a window nothing else would take down
    // — and a pin's page is left as it is (see `refresh_render_html`).
    if turned_off {
        refresh_render_html();
    }

    refresh_preview();
}

/// Whether an extensionless file is opened as raw text.
pub(in super::super) fn toggle_text_show_extensionless() {
    publish_config_change(
        |config| {
            config.text_show_extensionless = !config.text_show_extensionless;
        },
        &[PreviewMessage::Refresh],
    );
}

/// Whether secret and sensitive files are refused for text preview.
pub(in super::super) fn toggle_text_dont_preview_secrets() {
    publish_config_change(
        |config| {
            config.text_dont_preview_secrets = !config.text_dont_preview_secrets;
        },
        &[PreviewMessage::Refresh],
    );
}

/// Whether a pinned document's viewer takes the caret as it is revealed — `Preview Settings → PDF →
/// Pin Mode → Autofocus on PDF`.
///
/// **Nothing on screen is rebuilt for it, and nothing is told.** The row is read where a viewer is
/// revealed and where one is resumed, which is the pin's or the browser's to do and not this
/// process's (see `pdf_autofocus` and `engines::pdf_viewer::Host::take_the_caret`): a viewer that is
/// up keeps the keyboard it has, and the reveal after this one — or the resume after a pin put away
/// and brought back — is the first to answer the new value. A refresh would be a pin laid out again
/// for a setting that changes nothing of what is on screen.
pub(in super::super) fn toggle_pdf_autofocus() {
    publish_config_change(
        |config| {
            config.pdf_autofocus = !config.pdf_autofocus;
        },
        &[PreviewMessage::Refresh],
    );
}

/// Whether a pinned document is handed to Chromium's own viewer — `Preview Settings → PDF → Pin Mode →
/// Full Mode`.
///
/// **The row is the whole of what a pin does with a document, so it is answered where the pin is.**
/// On, the next tick asks for a viewer for the file the pin is standing on; off, the viewer the pin is
/// holding is let go of at once and what stands in the band is the single page this app draws — no
/// browser, no second window, and nothing of Chromium in the pin's path (see `pdf_full_mode`,
/// `close_the_pin` and `the_bands_pdf_viewer_is_owed`).
///
/// The pin is not laid out again for it: a box, a size and a file are the same ones either way, and
/// what changed is which window fills the band.
pub(in super::super) fn toggle_pdf_full_mode() {
    let full = commit_config(|config| {
        config.pdf_full_mode = !config.pdf_full_mode;
        config.pdf_full_mode
    });

    if !full {
        // In-process mode closes local PDF pins immediately; two-process mode propagates across
        // the seam (serving viewer receives ClosePdfPins; detached viewers observe config_generation).
        if !the_rings_are_the_road() {
            crate::engines::pdf_viewer::close_the_pins();
        }
        send_preview(PreviewMessage::ClosePdfPins);
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
    commit_config(|config| {
        config.normalize_volume = !config.normalize_volume;
    });
}

/// And the video's own switch, which is the same question asked about a film's soundtrack: read
/// where its player is started, like the level beside it, and off where the app starts — a film is
/// looked at, and a measurement of its audio is a decode of the film.
pub(in super::super) fn toggle_normalize_video_volume() {
    commit_config(|config| {
        config.normalize_video_volume = !config.normalize_video_volume;
    });
}

/// The row under a sound's `Normalize`: whether a level turned on a pinned window's own knob is
/// the level the next sound is previewed at.
///
/// Nothing on screen is rebuilt, and a sound already playing is left where it is: what the switch
/// changes is where the *next* player starts, which is the same bargain the level beside it makes.
pub(in super::super) fn toggle_remember_audio_volume() {
    commit_config(|config| {
        config.remember_audio_volume = !config.remember_audio_volume;
    });
}

/// And the video's own, which is the same switch asked about a soundtrack and off where the app
/// starts, like the video's level it decides.
pub(in super::super) fn toggle_remember_video_volume() {
    commit_config(|config| {
        config.remember_video_volume = !config.remember_video_volume;
    });
}
