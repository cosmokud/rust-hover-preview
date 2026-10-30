//! The one mpv a hover's film is loaded into, kept warm between hovers.
//!
//! **A film per hover costs a player per hover**, and a player's first frame costs what the
//! process takes to come up: measured here, `loadfile` on a player already running puts its first
//! frame up in 21–32 ms, where a cold spawn takes 200–400 ms. So the hover road keeps one mpv —
//! the *resident* — begun with a hover's own argv and no file, and every hover after it is a file
//! loaded down the pipe this app already talks to mpv on (see `mpv_ipc`).
//!
//! **A resident is a process a good while before it is a pipe.** mpv makes its end of the JSON IPC
//! while it starts — measured here, 40–211 ms after the spawn — so the first hover a resident is
//! begun for has nowhere to write its `loadfile` yet. That hover waits `PIPE_UP_WAIT` for the pipe
//! and is served the moment it is there; where it is not there in time, the resident is left
//! standing by rather than ended, so the next hover finds a pipe that is long open and is served
//! down it (see `serve`).
//!
//! **Only the hover road.** A pin's player is begun by `pin_playback` and `pin_install` exactly as
//! it always was, and FFmpeg's player and the media engine are untouched: what a hover is saved is
//! the wait a hover was paying.
//!
//! **The resident's own window is never seen and never takes the keyboard.** It is begun at
//! `-32000,-32000,1x1` — off every display — with the launch's `--focus-on*` switch where this
//! build has one, and it is noted as a hover about to show before mpv puts the window up, so the
//! monitor's eager pass gives it `WS_EX_NOACTIVATE` first (see
//! `ffplay_window::note_a_player_about_to_show`). Its size is not what the film is drawn at: every
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

/// How long a resident that has just been begun is waited on for its pipe to be there.
///
/// **The process is not the pipe.** On this machine with mpv 0.41.0 the pipe appears 40 ms after
/// the spawn with the file cache warm and 211 ms after it on a first run, and a `loadfile` offered
/// before that is answered `ERROR_FILE_NOT_FOUND` at once — which was every first hover of every
/// resident, since a resident is begun idle and the first line it is ever sent is a file. So the
/// first hover of a resident waits this long for somewhere to write: long enough to catch a
/// warm-cache start several times over, and short enough that it is well under the cold launch a
/// hover falls back to, so even a miss costs less than the player it then takes (see `serve` and
/// `mpv_ipc::deliver_within`).
pub(super) const PIPE_UP_WAIT: Duration = Duration::from_millis(120);

/// How long the resident is given to say the file it was handed is open, and how often it is asked.
///
/// The answer arrives in 21–32 ms when it comes at all, and a player that has not said it inside
/// half a second is one this app stops waiting for: the hover takes a cold player instead, and the
/// resident is left standing by with its film stopped, so the next hover is served by it rather
/// than by a player begun for that one too (see `serve`).
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

/// How long the file a hover asks for is waited on: some of `PIPE_UP_WAIT` for a resident this
/// hover is the one that began, and nothing at all for a resident that was already standing.
///
/// **A standing resident's pipe is long open**, so a line offered to it is a line mpv hears at once
/// and waiting on it would be the wait `deliver_within` keeps for a pipe that may not exist, paid
/// by every hover for nothing. It is a rule of its own rather than an `if` beside the call so that
/// it can be asked directly: the line either road offers is the same bytes (see
/// `mpv_ipc::load_file_line`), and a test build notes that line instead of writing it, so which of
/// the two functions ran is not a thing the suite could tell from the lines alone.
pub(super) fn the_wait_for_the_pipe(begun_now: bool) -> Option<Duration> {
    begun_now.then_some(PIPE_UP_WAIT)
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
    // (see `ffplay_window::note_a_player_about_to_show`).
    //
    // **And it is taken out of the Shell's lists before mpv shows it, because it has no file in it.**
    // The resident is begun with `--force-window=yes` so that a hover has a window to draw into, and
    // that window is a real top-level window of the user's until it is told otherwise: the eager pass
    // above is written for a window that is not on the screen yet, and a resident's window is shown no
    // matter how quickly the monitor looks. So the pid is named for the one road that writes the bits
    // on a window that may already be listed — the Shell reads them again on the show it pays for
    // there — and the pause between this line and that pass is a button titled `No file` at worst (see
    // `ffplay_window::note_a_window_to_unlist` and `eager_ex_style`).
    ffplay_window::note_a_player_about_to_show(pid, true);
    ffplay_window::note_a_window_to_unlist(pid);
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
/// the player did not say the file was open inside `FILE_LOADED_WAIT`. **None of them ends the
/// resident.** A resident whose pipe had not opened yet, or which never drew a frame, is one the
/// next hover can be served by: it is stopped and left standing *by* — not serving, its TTL
/// running — rather than killed for a failure this hover is already paying for.
pub(super) fn serve(path: &Path, volume: u32, behind_the_host: bool) -> Option<Served> {
    if !keeping_a_resident() || !the_resident_may_serve(pinned()) {
        return None;
    }

    let signature = LaunchSignature::of_the_machine(behind_the_host)?;
    let audio_off = video_launch::mpv_launched_without_audio(volume);
    // Whether this hover is the one the resident was begun for, which is the whole of the
    // difference between a pipe that is there and one that may not be yet (see `PIPE_UP_WAIT`).
    let begun_now;
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
        begun_now = slot.is_none();
        if begun_now {
            *slot = begun(signature, volume, audio_off);
        }

        let resident = slot.as_mut()?;
        resident.serving = true;
        resident.idle_since = None;
        (resident.player.id(), resident.pipe.clone())
    };

    // **The pipe is named again before the file is offered, and the level is asked again per file**
    // because the command line's per-file options are given to every file the player loads —
    // `--aid=no` among them (see `mpv_ipc::set_volume`). The naming is not only for the level: the
    // pipe this app holds is remembered against one player at a time, so a resident that stood by
    // while a cold player was begun for the hover before it is one whose name the record no longer
    // holds — and without naming it here a resident left warming would never be reached again (see
    // `mpv_ipc::remember_pipe`).
    let started = Instant::now();
    mpv_ipc::remember_pipe(pid, &pipe, audio_off);
    let offered = match the_wait_for_the_pipe(begun_now) {
        Some(within) => mpv_ipc::load_file_within(pid, path, within),
        None => mpv_ipc::load_file(pid, path),
    };
    let asked = offered && mpv_ipc::set_volume(pid, volume);
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
    note_pin_trace!(
        "hover resident left warming ({})",
        if asked { "no frame" } else { "the pipe" }
    );

    // **It stands by rather than being ended.** A pipe that had not opened yet, or a frame that had
    // not landed inside `FILE_LOADED_WAIT`, is not a player this app has lost: it is one the next
    // hover can be served by. So the film is stopped on it and the clock its TTL runs on is started
    // — which is what lets `tick` reap it where it is genuinely of no use, and what leaves it
    // holding no hover while it waits for the one that can use it.
    if let Ok(mut slot) = RESIDENT.lock() {
        if let Some(resident) = slot.as_mut().filter(|resident| resident.player.id() == pid) {
            let _ = mpv_ipc::stop_playing(pid);
            resident.serving = false;
            resident.idle_since = Some(Instant::now());
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
    let Some(resident) = slot.as_mut().filter(|resident| resident.player.id() == pid) else {
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

    if the_resident_is_owed_a_re_warm(RE_WARM.load(Ordering::Acquire), pinned(), slot.is_some()) {
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
            && resident.idle_since.is_some_and(|since| {
                the_resident_may_be_reaped(false, false, idle, since.elapsed())
            })
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
