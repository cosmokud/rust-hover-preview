//! The one mpv a hover's film is loaded into, kept warm between hovers.
//!
//! **A film per hover costs a player per hover**, and a player's first frame costs what the
//! process takes to come up: measured here, `loadfile` on a player already running puts its first
//! frame up in 21–32 ms, where a cold spawn takes 200–400 ms. So the hover road keeps one mpv —
//! the *resident* — begun with a hover's own argv and no file, and every hover after it is a file
//! loaded down the pipe this app already talks to mpv on (see `mpv_ipc`).
//!
//! **Only the hover road.** A pin's player is begun by `pin_playback` and `pin_install` exactly as
//! it always was, and FFmpeg's player and the media engine are untouched: what a hover is saved is
//! the wait a hover was paying.
//!
//! **The resident's own window is never seen and never takes the keyboard.** It is begun at
//! `-32000,-32000,1x1` — off every display — with the launch's `--focus-on*` switch where this
//! build has one, and it is noted as a hover about to show before mpv puts the window up, so the
//! monitor's eager pass gives it `WS_EX_NOACTIVATE` first (see
//! `ffplay_window::note_a_hover_about_to_show`). Its size is not what the film is drawn at: every
//! hover that shows it places or thumbnails it at the box it belongs in, which is a window resize
//! mpv answers like any other.
//!
//! **What ends a hover does not end the resident.** `stop_video_playback` stops the film on it
//! (mpv goes idle with its window still up) rather than killing it, and every other pid — a pin's,
//! FFmpeg's — is killed exactly as before (see `video_playback::kill_player_process`, which is the
//! one place that rule lives). What ends the resident is the `MPV TTL` setting (`video_engine_idle`,
//! an `EngineIdle`): `Away` keeps it while the user is at the keyboard, a number of seconds keeps it
//! that long after its last hover, `Indefinitely` never lets it go and `0` lets it go the moment the
//! hover ends. A resident that is showing a hover, or that a pin has taken over, is never reaped.
//!
//! **A pin that takes a hover's player takes the resident with it, and it is not this app's from
//! there.** The player a pin adopts *is* the resident (same process), so its record here is cleared
//! at the take-over and a fresh resident is begun once that pin has closed (see
//! `a_pin_took_it_over` and `tick`).

use super::pin_trace::note_pin_trace;
use super::*;
use crate::config::config::EngineIdle;

/// Where the resident's window is parked: off every display, at the smallest window there is, so
/// that nothing of it is ever on the screen before a hover places it (see the module note).
const PARKED_AT: (i32, i32) = (-32000, -32000);
const PARKED_SIZE: i32 = 1;

/// How long the resident is given to say the file it was handed is open, and how often it is asked.
///
/// The answer arrives in 21–32 ms when it comes at all, and a player that has not said it inside
/// half a second is one this app stops waiting for: the hover takes a cold player instead, which
/// costs the wait the resident exists to save and nothing else (see `serve`).
const FILE_LOADED_WAIT: Duration = Duration::from_millis(500);
const FILE_LOADED_POLL: Duration = Duration::from_millis(2);

/// What a resident was begun with, as far as a later hover can disagree with it.
///
/// **Every field is a thing the command line names and a hover cannot change once it is running**:
/// the program itself, the subtitle track answer (`--sid`/`--sub-auto`), the decode (`--hwdec`),
/// the band the window is begun in rather than the thumbnail's (`--ontop`), and the mouse's own
/// binding file. A setting changed under a standing resident is a resident begun again on the next
/// hover (see `the_signature_still_stands`).
#[derive(Clone, PartialEq, Eq, Debug)]
pub(super) struct LaunchSignature {
    pub(super) program: PathBuf,
    pub(super) subtitles: bool,
    pub(super) hw_accel: bool,
    pub(super) behind_the_host: bool,
    pub(super) input_conf: Option<PathBuf>,
}

impl LaunchSignature {
    /// What the settings ask for at this moment, for a hover whose picture is (or is not) a
    /// thumbnail of the player's own window.
    fn of_the_machine(behind_the_host: bool) -> Option<Self> {
        let chosen = CONFIG
            .lock()
            .map(|config| config.mpv_path.clone())
            .unwrap_or_default();

        Some(Self {
            program: video_hw::mpv_program(&chosen, &codecs::mpv_installs())?,
            subtitles: subtitles_wanted(),
            hw_accel: video_hw_accel_wanted(),
            behind_the_host,
            input_conf: video_launch::mpv_input_conf().map(Path::to_path_buf),
        })
    }
}

/// The player this app keeps for a hover's film: the process, the pipe it is told on, what it was
/// begun with, and how it stands — serving a hover, or standing by with the clock its TTL runs on.
struct Resident {
    player: PlayerProcess,
    pipe: String,
    signature: LaunchSignature,
    /// Whether a hover is showing on it. A resident that is serving is never reaped.
    serving: bool,
    /// When it last stopped serving, which is when its TTL started.
    idle_since: Option<Instant>,
}

/// The one resident, or nothing while there is none.
static RESIDENT: Mutex<Option<Resident>> = Mutex::new(None);

/// Whether the tick owes a fresh resident after a pin took the last one over (see
/// `a_pin_took_it_over`).
static RE_WARM: AtomicBool = AtomicBool::new(false);

/// Whether the hover road keeps a resident at all: on in a run, and in a test only where the test
/// asks for it.
///
/// **The suite's other hover tests must keep the launch they were written for**, and one of them
/// sharing a resident with the next is one noted player standing in for two launches — so the seam
/// is off by default and the tests that are about the resident turn it on (see
/// `keep_a_resident_in_a_test`). The release binary carries no part of it.
#[cfg(not(test))]
fn keeping_a_resident() -> bool {
    true
}

#[cfg(test)]
fn keeping_a_resident() -> bool {
    // The thread's own answer rather than the process's: a test that keeps a resident is a test of
    // the resident, and every other test's hover must keep the launch it was written for even while
    // that one runs beside it.
    RESIDENT_ASKED.with(|held| held.get())
}

#[cfg(test)]
thread_local! {
    static RESIDENT_ASKED: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

/// Keep a resident in a test build, or stop doing so (see `keeping_a_resident`).
#[cfg(test)]
pub(super) fn keep_a_resident_in_a_test(keep: bool) {
    RESIDENT_ASKED.with(|held| held.set(keep));
    if let Ok(mut slot) = RESIDENT.lock() {
        if let Some(resident) = slot.take() {
            end(resident);
        }
    }
}

/// The resident that is serving a hover: the handle the launch hands on, which is the same type
/// every other road's player is (see `video_playback::PlayerProcess`).
pub(super) struct Served {
    pub(super) player: PlayerProcess,
}

/// Whether a hover may be loaded on the resident at all: a pin that stands is a picture of its own,
/// and a film begun under it is the one it has today.
pub(super) fn the_resident_may_serve(a_pin_stands: bool) -> bool {
    !a_pin_stands
}

/// Whether this resident may be let go: only a standby one, whose TTL has run out.
///
/// A resident that is serving a hover is the picture on screen, and one a pin has taken over is the
/// pin's — neither is this app's to end, however long the clock has run.
pub(super) fn the_resident_may_be_reaped(
    serving: bool,
    pin_owned: bool,
    idle: EngineIdle,
    idle_for: Duration,
) -> bool {
    !serving && !pin_owned && idle.has_expired(idle_for)
}

/// Whether the tick begins a fresh resident: one is owed by a pin that took the last over, no pin
/// stands, and none is held.
pub(super) fn the_resident_is_owed_a_re_warm(
    owed: bool,
    a_pin_stands: bool,
    a_resident_is_held: bool,
) -> bool {
    owed && !a_pin_stands && !a_resident_is_held
}

/// The word a reaped resident's line says: which half of `MPV TTL` let it go.
fn the_ttl_word(idle: EngineIdle) -> &'static str {
    match idle {
        EngineIdle::Away => "away",
        EngineIdle::Seconds(_) | EngineIdle::Indefinite => "ttl",
    }
}

/// The `MPV TTL` setting as it stands.
fn the_ttl() -> EngineIdle {
    CONFIG
        .lock()
        .map(|config| config.video_engine_idle)
        .unwrap_or(EngineIdle::Away)
}

/// Whether the resident that was begun with `begun` is the player the settings ask for now.
pub(super) fn the_signature_still_stands(
    begun: &LaunchSignature,
    wanted: &LaunchSignature,
) -> bool {
    begun == wanted
}

/// Whether a player this app holds has ended.
fn the_player_has_ended(player: &mut PlayerProcess) -> bool {
    matches!(player.try_wait(), Ok(Some(_)))
}

/// The pid of the resident, or nothing while there is none.
pub(super) fn the_resident_pid() -> Option<u32> {
    RESIDENT
        .lock()
        .ok()
        .and_then(|slot| slot.as_ref().map(|resident| resident.player.id()))
}

/// Whether `pid` is the resident this app is holding.
pub(super) fn holds_the_pid(pid: u32) -> bool {
    the_resident_pid() == Some(pid)
}

/// Begin a resident, and answer what was begun. It is stored by the caller, so that the record and
/// the lock are written in one place.
fn begun(signature: LaunchSignature, volume: u32, audio_off: bool) -> Option<Resident> {
    let ipc = mpv_ipc::next_pipe_name();

    // **The hover's own argv with no file in it.** `mpv_args` writes the file it is given as its
    // first argument, so the resident is built from it with an empty one and that argument dropped:
    // every flag a hover's player is begun with is the one the resident is begun with, which is what
    // makes a file loaded on it the same player a hover would have had.
    let mut args = video_launch::mpv_args(
        Path::new(""),
        PARKED_AT.0,
        PARKED_AT.1,
        PARKED_SIZE,
        PARKED_SIZE,
        0.0,
        false,
        volume,
        None,
        signature.hw_accel,
        signature.subtitles,
        signature.input_conf.as_deref(),
        &ipc,
        // A window that is never seen keeps the bitblt swapchain: the thumbnail road makes a source
        // layered to keep it off the screen, and a layered window cannot have a flip-model one.
        false,
        !signature.behind_the_host,
    );
    args.remove(0);
    // The two that make it a resident rather than a player: no file to play, and a window to draw
    // into while there is none (see the module note).
    args.push("--idle=yes".to_string());
    args.push("--force-window=yes".to_string());
    if let Some(switch) = video_launch::hover_focus_switch(video_launch::the_focus_on_support()) {
        args.push(switch.to_string());
    }

    let mut command = engine_processes::hidden_command(&signature.program);
    command
        .args(args)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());

    let player = spawn_player(&mut command)?;
    let pid = player.id();

    // **The window is a hover's window before mpv shows it**: the eager pass gives it
    // `WS_EX_NOACTIVATE` first, which is the discipline every hover player's window is begun under
    // (see `ffplay_window::note_a_hover_about_to_show`).
    ffplay_window::note_a_hover_about_to_show(pid);
    mpv_ipc::remember_pipe(pid, &ipc, audio_off);
    note_pin_trace!("mpv resident begun pid {pid}");

    Some(Resident {
        player,
        pipe: ipc,
        signature,
        serving: false,
        idle_since: None,
    })
}

/// End a resident: the process goes, and the pipe a command would reach it on goes with it.
fn end(resident: Resident) {
    let pid = resident.player.id();
    let mut player = resident.player;
    mpv_ipc::forget_pipe(pid);
    let _ = player.kill();
}

/// Begin a resident for a test, with nothing loaded on it: the pid a test answers a read for before
/// the hover that is to be served by it (see `mpv_ipc::answer_a_read` and `serve`).
#[cfg(test)]
pub(super) fn begin_a_resident_for_a_test(behind_the_host: bool) -> Option<u32> {
    let signature = LaunchSignature::of_the_machine(behind_the_host)?;
    let volume = current_video_volume();
    let mut slot = RESIDENT.lock().ok()?;
    *slot = begun(
        signature,
        volume,
        video_launch::mpv_launched_without_audio(volume),
    );
    slot.as_ref().map(|resident| resident.player.id())
}

/// Begin the resident if there is none, or hand back the one standing, and load `path` on it.
///
/// `None` is every way a hover is not served, and every one of them leaves the caller to begin a
/// player of its own: no resident could be begun at all, a pin stood, the file was not accepted, or
/// the player did not say the file was open inside `FILE_LOADED_WAIT` — in which case the resident
/// is ended and forgotten rather than left behind as a player this app cannot talk to.
pub(super) fn serve(path: &Path, volume: u32, behind_the_host: bool) -> Option<Served> {
    if !keeping_a_resident() || !the_resident_may_serve(pinned()) {
        return None;
    }

    let signature = LaunchSignature::of_the_machine(behind_the_host)?;
    let audio_off = video_launch::mpv_launched_without_audio(volume);
    let (pid, pipe) = {
        let mut slot = RESIDENT.lock().ok()?;

        // A resident that has ended is no resident, and the record goes rather than answering for a
        // player that is not there.
        if slot
            .as_mut()
            .is_some_and(|resident| the_player_has_ended(&mut resident.player))
        {
            if let Some(resident) = slot.take() {
                mpv_ipc::forget_pipe(resident.player.id());
                note_pin_trace!("mpv resident let go (exit)");
            }
        }
        if let Some(resident) = slot.as_ref() {
            if !the_signature_still_stands(&resident.signature, &signature) {
                if let Some(resident) = slot.take() {
                    end(resident);
                    note_pin_trace!("mpv resident let go (mismatch)");
                }
            }
        }
        if slot.is_none() {
            *slot = begun(signature, volume, audio_off);
        }

        let resident = slot.as_mut()?;
        resident.serving = true;
        resident.idle_since = None;
        (resident.player.id(), resident.pipe.clone())
    };

    // **The level is asked again per file**, because the command line's per-file options are given
    // to every file the player loads — `--aid=no` among them — so the pipe is remembered again with
    // what the launch said of its audio before the level is moved (see `mpv_ipc::set_volume`).
    let started = Instant::now();
    let asked =
        mpv_ipc::load_file(pid, path) && {
            mpv_ipc::remember_pipe(pid, &pipe, audio_off);
            mpv_ipc::set_volume(pid, volume)
        };
    // A file is not held where the last one was left: a hover shows its film playing, from its
    // beginning, which is what `loadfile ... replace` and this one ask say together.
    let asked = asked && mpv_ipc::set_paused(pid, false);

    if asked && await_the_file(pid) {
        note_pin_trace!(
            "hover loadfile {} first frame in {} ms",
            pin_trace::path_name(path),
            started.elapsed().as_millis()
        );
        return Some(Served {
            player: PlayerProcess::for_the_resident(pid),
        });
    }

    note_pin_trace!(
        "hover loadfile failed ({}), began a cold player",
        if asked { "no frame" } else { "the pipe" }
    );
    if let Ok(mut slot) = RESIDENT.lock() {
        let failed = slot.as_ref().is_some_and(|resident| resident.player.id() == pid);
        if failed {
            if let Some(resident) = slot.take() {
                end(resident);
            }
        }
    }

    None
}

/// Wait for the player's own word that the file it was handed has a frame in it, bounded by
/// `FILE_LOADED_WAIT`.
///
/// **The answer is asked on a thread of its own** so that a player which never speaks costs this
/// app the wait and not the thread that is holding the hover — the same bound `mpv_ipc::get_property`
/// keeps for one read, asked again until the file is up. `video-params` is what a player reports
/// once it has decoded a frame of the file, which is the whole of what a served hover needs: the
/// frame is the picture, and a player that has not drawn one is not a hover's film yet.
fn await_the_file(pid: u32) -> bool {
    let (said, heard) = channel();
    std::thread::spawn(move || {
        let deadline = Instant::now() + FILE_LOADED_WAIT;
        loop {
            if mpv_ipc::get_property(pid, "video-params").is_some() {
                break said.send(true).is_ok();
            }
            if Instant::now() >= deadline {
                break said.send(false).is_ok();
            }
            std::thread::sleep(FILE_LOADED_POLL);
        }
    });

    heard
        .recv_timeout(FILE_LOADED_WAIT + Duration::from_millis(100))
        .unwrap_or(false)
}

/// The hover that was served has ended: the resident is left standing by rather than ended, and the
/// clock its TTL runs on starts here.
///
/// It answers whether `pid` is the resident — the one question `video_playback::kill_player_process`
/// asks, so that every other player is killed exactly as it always was.
pub(super) fn finish(pid: u32) -> bool {
    let Ok(mut slot) = RESIDENT.lock() else {
        return false;
    };
    let Some(resident) = slot
        .as_mut()
        .filter(|resident| resident.player.id() == pid)
    else {
        return false;
    };

    let _ = mpv_ipc::stop_playing(pid);
    resident.serving = false;
    resident.idle_since = Some(Instant::now());
    true
}

/// A pin has taken the hover's player over as its own picture (see
/// `pin_film_host::promote_the_hover_film_host`): the resident is the pin's from here.
///
/// The record goes and the process is not touched — the pin is drawing it — so nothing loads a
/// hover's file on it, nothing reaps it, and nothing but the pin ends it. A fresh resident is begun
/// once that pin has closed (see `tick`).
pub(super) fn a_pin_took_it_over() {
    let Ok(mut slot) = RESIDENT.lock() else {
        return;
    };
    let took_it = slot
        .as_ref()
        .is_some_and(|resident| resident.player.id() == VIDEO_PID.load(Ordering::SeqCst));
    if !took_it {
        return;
    }
    if let Some(resident) = slot.take() {
        note_pin_trace!(
            "mpv resident let go (a pin took it over) pid {}",
            resident.player.id()
        );
    }
    RE_WARM.store(true, Ordering::Release);
}

/// One pass of the preview tick's own work on the resident: a fresh one that a pin's take-over owes,
/// and one standing by whose TTL has run out.
pub(super) fn tick() {
    if !keeping_a_resident() {
        return;
    }
    let Ok(mut slot) = RESIDENT.lock() else {
        return;
    };

    // A resident that has ended — a pin's relaunch ended it, or it came apart — is no resident, and
    // the next one is begun when it is wanted rather than left as a record of a process that is
    // gone.
    if slot
        .as_mut()
        .is_some_and(|resident| the_player_has_ended(&mut resident.player))
    {
        if let Some(resident) = slot.take() {
            mpv_ipc::forget_pipe(resident.player.id());
            note_pin_trace!("mpv resident let go (exit)");
        }
        RE_WARM.store(true, Ordering::Release);
    }

    if the_resident_is_owed_a_re_warm(
        RE_WARM.load(Ordering::Acquire),
        pinned(),
        slot.is_some(),
    ) {
        RE_WARM.store(false, Ordering::Release);
        let ttl_volume = current_video_volume();
        if let Some(signature) = LaunchSignature::of_the_machine(false) {
            *slot = begun(
                signature,
                ttl_volume,
                video_launch::mpv_launched_without_audio(ttl_volume),
            );
        }
    }

    let idle = the_ttl();
    let expired = slot.as_ref().is_some_and(|resident| {
        !resident.serving
            && resident
                .idle_since
                .is_some_and(|since| the_resident_may_be_reaped(false, false, idle, since.elapsed()))
    });
    if expired {
        if let Some(resident) = slot.take() {
            end(resident);
            note_pin_trace!("mpv resident let go ({})", the_ttl_word(idle));
        }
    }
}

/// End the resident, because the app is ending.
pub(crate) fn shut_down_the_resident() {
    let Ok(mut slot) = RESIDENT.lock() else {
        return;
    };
    if let Some(resident) = slot.take() {
        end(resident);
        note_pin_trace!("mpv resident let go (the app is ending)");
    }
}
