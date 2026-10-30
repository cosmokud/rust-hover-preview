//! The one mpv a hover's film is loaded into, kept warm between hovers: what it is begun with, what
//! a served hover asks it, what a hover ending does to it, and what its `MPV TTL` lets go of.
//!
//! No real player is begun here — the suite's `spawn_player` notes one instead — and the seam that
//! keeps a resident at all is off for every other test, so that one hover's resident is never another
//! test's launch (see `keep_a_resident_in_a_test`).

use super::*;
use crate::config::config::EngineIdle;

/// The hover road's resident, kept for the length of one test and taken down on the way out.
struct ResidentStand {
    previous_pin_enabled: bool,
}

impl ResidentStand {
    fn new(pin_enabled: bool) -> Self {
        stand_pin(None);
        VIDEO_PID.store(0, Ordering::SeqCst);
        mpv_ipc::clear_noted();
        mpv_ipc::clear_reads();
        pin_trace::clear_noted();
        let previous_pin_enabled = CONFIG
            .lock()
            .map(|config| config.pin_enabled)
            .unwrap_or(true);
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_enabled = pin_enabled;
        }
        stand_mpv_in(Some(true));
        stand_ffplay_in(Some(false));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));
        mpv_resident::keep_a_resident_in_a_test(true);
        Self {
            previous_pin_enabled,
        }
    }
}

impl Drop for ResidentStand {
    fn drop(&mut self) {
        mpv_resident::keep_a_resident_in_a_test(false);
        stand_mpv_in(None);
        stand_ffplay_in(None);
        stand_video_engine_in(None);
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_enabled = self.previous_pin_enabled;
        }
        VIDEO_PID.store(0, Ordering::SeqCst);
        stand_pin(None);
    }
}

fn a_film() -> PathBuf {
    PathBuf::from("a-hovered-film.mkv")
}

/// A hover's launch, the way the loop makes it: from the beginning, not held, at the hover's level,
/// and off the user's desktop, which is where every window a test makes stands.
fn a_hover(path: &PathBuf, volume: u32) -> Option<PlayerProcess> {
    start_hover_video_playback(
        path, -32000, -32000, 320, 180, 0.0, false, volume, None, None,
    )
}

/// A resident begun for a test, with nothing loaded on it yet: the pid a test answers a read for
/// before the hover that is to be served by it (see `mpv_ipc::answer_a_read`).
fn a_resident(behind_the_host: bool) -> u32 {
    mpv_resident::begin_a_resident_for_a_test(behind_the_host).expect("a noted resident")
}

/// A hover's media, with the player it is showing.
fn hover_media(player: PlayerProcess) -> MediaData {
    let mut media = create_loading_media(16, 16);
    media.media_type = MediaType::Video;
    media.video_process = Some(player);
    media
}

/// **A hover with no pin up is loaded on the resident**, which is the whole of what the road is for:
/// the file and the hover's own level go down the pipe this app already holds, and the player that
/// comes back is the one already standing rather than a process begun for this hover.
#[test]
fn a_resident_serves_a_hover_when_no_pin_stands() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    assert!(
        mpv_resident::the_resident_may_serve(false),
        "a hover is served while no pin stands"
    );

    let pid = a_resident(true);
    mpv_ipc::clear_noted();
    mpv_ipc::answer_a_read(pid, "video-params", Some("{}"));
    let _ = take_last_spawned_command();

    let served = a_hover(&a_film(), 55).expect("the resident, serving this hover");

    assert_eq!(served.id(), pid, "the player is the resident that stood");
    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "and it is still the resident, holding the hover"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "no player was begun for this hover"
    );
}

/// **A pin that stands is a film a hover is not the best picture of**: nothing is loaded on the
/// resident while one is up, and a hover begun under a pin is the player it always was.
#[test]
fn a_resident_never_serves_while_a_pin_stands() {
    assert!(
        !mpv_resident::the_resident_may_serve(true),
        "a pin's film is never loaded on the hover's resident"
    );
}

/// The command a served hover is, exactly: the file and the hover's own level, and nothing else.
#[test]
fn loadfile_asks_for_the_file_and_the_hover_volume() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    mpv_ipc::clear_noted();
    mpv_ipc::answer_a_read(pid, "video-params", Some("{}"));

    let _ = a_hover(&a_film(), 55).expect("the resident, serving this hover");

    let lines = mpv_ipc::noted_lines();
    assert!(
        lines
            .iter()
            .any(|line| line.contains(r#""command":["loadfile""#)
                && line.contains("a-hovered-film.mkv")
                && line.contains(r#""replace""#)),
        "the file is loaded on the player that is already there: {lines:?}"
    );
    assert!(
        lines
            .iter()
            .any(|line| line.contains(r#""command":["set_property","volume",55]"#)),
        "and the hover's own level is asked for it: {lines:?}"
    );
    assert!(
        lines
            .iter()
            .any(|line| line.contains(r#""command":["set_property","pause",false]"#)),
        "a hover's film plays from its beginning rather than held where the last one was: {lines:?}"
    );
}

/// **A hover that cannot be served is the launch this app always had, and it costs the resident
/// nothing.** A player that never says the file is open is stopped and left standing by rather than
/// ended, the failure is written down, a cold player is begun for the hover — and the TTL can still
/// let that resident go, which is what keeps a resident that never serves anything from standing
/// for the rest of the run.
#[test]
fn a_hover_that_cannot_be_served_falls_back_to_a_cold_player() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();
    let _ = take_last_spawned_command();

    let cold = a_hover(&a_film(), 55).expect("a player of the hover's own");

    assert_ne!(
        cold.id(),
        pid,
        "a player begun for the hover, not the resident"
    );
    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "and the resident that could not serve it is kept rather than killed for it"
    );
    assert!(
        take_last_spawned_command().is_some(),
        "the cold player is the launch the hover road had before the resident"
    );
    let traced = pin_trace::noted_lines();
    assert!(
        traced
            .iter()
            .any(|line| line.contains("hover loadfile failed")),
        "and the run says the file was not taken: {traced:?}"
    );
    assert!(
        traced
            .iter()
            .any(|line| line.contains("hover resident left warming")),
        "and that the resident was left standing by rather than let go: {traced:?}"
    );
    assert!(
        mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains(r#""command":["stop"]"#)),
        "with its film stopped, so it stands by idle rather than holding a file it never took: {:?}",
        mpv_ipc::noted_lines()
    );

    // **And the TTL still has it**, which is what keeps a resident that can never serve from
    // standing for the rest of the run: standing by is the state a resident is reaped from.
    let previous_idle = CONFIG
        .lock()
        .map(|config| config.video_engine_idle)
        .unwrap_or(EngineIdle::Away);
    if let Ok(mut config) = CONFIG.lock() {
        config.video_engine_idle = EngineIdle::Seconds(0);
    }
    mpv_resident::tick();
    if let Ok(mut config) = CONFIG.lock() {
        config.video_engine_idle = previous_idle;
    }

    assert_eq!(
        mpv_resident::the_resident_pid(),
        None,
        "a TTL of nothing lets the resident it left standing by go at once"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.contains("mpv resident let go (ttl)")),
        "and the run says which half of the TTL let it go: {:?}",
        pin_trace::noted_lines()
    );
}

/// **A resident whose first hover it could not serve is the one that serves the second.** The
/// failure used to end the resident, so every hover began a player, found its pipe not yet there,
/// killed that player and began another — the warm road that never once ran outside the suite. A
/// resident left standing by is still the resident, and the hover after it is loaded on it.
#[test]
fn a_resident_whose_first_load_failed_serves_the_next_hover() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();
    let _ = take_last_spawned_command();

    // The first hover: nothing ever answers for the file, so the resident cannot say it is open.
    let first = a_hover(&a_film(), 55).expect("a player of the hover's own");
    assert_ne!(first.id(), pid, "the resident could not serve this hover");
    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "and is kept rather than ended for it"
    );
    let _ = take_last_spawned_command();

    // The next hover, on a resident whose pipe is by now long open: it is the one that serves it.
    mpv_ipc::clear_noted();
    mpv_ipc::answer_a_read(pid, "video-params", Some("{}"));
    let served = a_hover(&a_film(), 55).expect("the resident, serving this hover");

    assert_eq!(
        served.id(),
        pid,
        "the resident the failed hover left warming is the one that serves the next hover"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "so no second player is begun for it"
    );
    assert!(
        mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains(r#""command":["loadfile""#)),
        "and the file is loaded down its pipe: {:?}",
        mpv_ipc::noted_lines()
    );
}

/// **Only a resident that this hover itself began is waited on for its pipe.** A resident that was
/// already standing has a pipe that is long open, and waiting on it would be the wait kept for a
/// pipe that may not exist, paid by every hover for nothing.
///
/// Which of the two asks `serve` made is not something a test build can see: the pipe under both is
/// the one seam a test run never touches — `mpv_ipc` notes the line and writes nothing, before
/// either function opens anything — and the line is the same bytes either way (see
/// `mpv_ipc::load_file_line`). So it is the rule that is asserted, and `serve` reads the rule from
/// here rather than keeping its own copy of it.
#[test]
fn only_a_just_begun_resident_is_waited_on_for_its_pipe() {
    assert_eq!(
        mpv_resident::the_wait_for_the_pipe(true),
        Some(mpv_resident::PIPE_UP_WAIT),
        "a resident begun for this hover may not have its pipe yet, so its file is waited on"
    );
    assert_eq!(
        mpv_resident::the_wait_for_the_pipe(false),
        None,
        "a resident that was already standing is offered its line at once, with no wait at all"
    );
}

/// **A hover ending stops the resident and does not kill it**: mpv goes idle with nothing playing,
/// and the player is still there for the next hover.
#[test]
fn a_finished_hover_stops_the_resident_and_does_not_kill_it() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    mpv_ipc::answer_a_read(pid, "video-params", Some("{}"));
    let served = a_hover(&a_film(), 55).expect("the resident, serving this hover");
    mpv_ipc::clear_noted();

    let mut media = hover_media(served);
    stop_video_playback(&mut media);

    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "the hover ending leaves the resident standing"
    );
    assert!(media.video_process.is_none(), "with no player on screen");
    assert!(
        mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains(r#""command":["stop"]"#)),
        "and the film stopped on it rather than the player ended: {:?}",
        mpv_ipc::noted_lines()
    );
}

/// **Every other player is killed exactly as it was**, which is the half of the rule above that keeps
/// a pin's player, FFmpeg's, and a stray from being treated as the resident: the pipe a command would
/// reach it on goes with it.
#[test]
fn any_other_player_is_killed_as_before() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let _resident = a_resident(true);
    let mut command = engine_processes::hidden_command("cmd");
    let player = video_playback::spawn_player(&mut command).expect("a noted player");
    let pid = player.id();
    mpv_ipc::remember_pipe(pid, "a-pipe-of-its-own", false);
    assert!(mpv_ipc::answers_properties(pid), "the player is held");

    let mut media = hover_media(player);
    kill_player_process(&mut media);

    assert!(
        media.video_process.is_none(),
        "a player that is not the resident is ended"
    );
    assert!(
        !mpv_ipc::answers_properties(pid),
        "and the pipe it was told on went with it"
    );
}

/// **The TTL item says when a standby resident expires**, in all four shapes the setting takes: a
/// number of seconds after its last hover, the away timer for `Away`, never for `Indefinitely`, and at
/// once for `0`.
#[test]
fn the_ttl_item_says_when_a_standby_resident_expires() {
    let _room = the_room_to_itself();

    let minute = Duration::from_secs(60);
    let year = Duration::from_secs(365 * 24 * 60 * 60);

    assert!(
        !mpv_resident::the_resident_may_be_reaped(
            false,
            false,
            EngineIdle::Seconds(60),
            minute - Duration::from_secs(1)
        ),
        "a resident kept for a minute is kept until it has stood by that long"
    );
    assert!(mpv_resident::the_resident_may_be_reaped(
        false,
        false,
        EngineIdle::Seconds(60),
        minute
    ));
    assert!(
        mpv_resident::the_resident_may_be_reaped(
            false,
            false,
            EngineIdle::Seconds(0),
            Duration::ZERO
        ),
        "a resident let go as soon as its hover ends is let go at once"
    );
    assert!(
        !mpv_resident::the_resident_may_be_reaped(false, false, EngineIdle::Indefinite, year),
        "one kept for the life of the app never expires"
    );

    // `Away` is the away timer's answer rather than the resident's own clock, which is what makes it
    // the one shape a duration cannot answer (see `EngineIdle::has_expired`).
    let timer = CONFIG
        .lock()
        .map(|config| config.afk_timer_seconds)
        .unwrap();
    crate::app::afk::note_explorer_reachable(true);
    assert!(
        !mpv_resident::the_resident_may_be_reaped(false, false, EngineIdle::Away, year),
        "a user at the keyboard keeps the resident however long it has stood by"
    );
    if let Ok(mut config) = CONFIG.lock() {
        config.afk_timer_seconds = 0;
    }
    crate::app::afk::note_explorer_reachable(false);
    assert!(
        mpv_resident::the_resident_may_be_reaped(false, false, EngineIdle::Away, Duration::ZERO),
        "and one who is away lets it go"
    );
    if let Ok(mut config) = CONFIG.lock() {
        config.afk_timer_seconds = timer;
    }
    crate::app::afk::note_explorer_reachable(true);
}

/// **A resident that is serving a hover, or that a pin has taken over, is never reaped** — whatever
/// the TTL says and however long it has stood by. Only a standby one is ever let go.
#[test]
fn a_serving_or_pin_owned_resident_is_never_reaped() {
    let year = Duration::from_secs(365 * 24 * 60 * 60);

    for idle in [
        EngineIdle::Seconds(0),
        EngineIdle::Seconds(60),
        EngineIdle::Indefinite,
    ] {
        assert!(
            !mpv_resident::the_resident_may_be_reaped(true, false, idle, year),
            "the picture on screen is never let go: {idle:?}"
        );
        assert!(
            !mpv_resident::the_resident_may_be_reaped(false, true, idle, year),
            "and neither is a pin's: {idle:?}"
        );
    }
    assert!(
        mpv_resident::the_resident_may_be_reaped(false, false, EngineIdle::Seconds(60), year),
        "the one that is let go is the standby resident its TTL has run out on"
    );
}

/// **A setting the resident's command line names is a resident begun again**: a hover that asks for a
/// player its resident was not begun as takes the old one down and loads on a fresh one.
#[test]
fn a_changed_launch_signature_restarts_the_resident() {
    let _room = the_room_to_itself();
    // No host wanted, so the launch's signature is the one the resident below was begun with…
    let _stand = ResidentStand::new(false);
    let pid = a_resident(false);

    // …and with Pin Mode off the very next hover is still the same signature.
    mpv_ipc::answer_a_read(pid, "video-params", Some("{}"));
    let _ = a_hover(&a_film(), 55).expect("the resident, serving this hover");
    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "the same settings are the same resident"
    );

    // A hover whose picture is a thumbnail of the player's own window is a player begun in another
    // band, which is not a command line this resident has.
    assert!(
        !mpv_resident::the_signature_still_stands(
            &pinned_nowhere_signature(),
            &a_signature_behind_a_host()
        ),
        "the two bands are two launches"
    );
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = true;
    }
    pin_trace::clear_noted();

    let _ = a_hover(&a_film(), 55);

    assert_ne!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "a hover that asks for another player is not served by this one"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.contains("mpv resident let go (mismatch)")),
        "and the run says why the standing one was let go: {:?}",
        pin_trace::noted_lines()
    );
}

/// The signature a hover with no host is begun with, as a test builds it.
fn pinned_nowhere_signature() -> mpv_resident::LaunchSignature {
    mpv_resident::LaunchSignature {
        program: PathBuf::from("mpv.exe"),
        subtitles: false,
        hw_accel: true,
        behind_the_host: false,
        input_conf: None,
    }
}

/// The same, for a hover whose picture is a thumbnail of the player's window.
fn a_signature_behind_a_host() -> mpv_resident::LaunchSignature {
    mpv_resident::LaunchSignature {
        behind_the_host: true,
        ..pinned_nowhere_signature()
    }
}

/// **The leftover sweep never sweeps the resident.** It is the player kept between hovers, and a
/// sweep is for the player that outlived the hover it was begun for.
#[test]
fn the_stray_sweep_spares_the_resident() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    // The sweep acts on the published id, which a served hover publishes as the resident's.
    VIDEO_PID.store(pid, Ordering::SeqCst);

    kill_stray_video_process();

    assert_eq!(
        mpv_resident::the_resident_pid(),
        Some(pid),
        "the resident outlives the sweep that would have taken a stray"
    );
}

/// **A pin that took the resident over re-warms one after it closes.** The player is the pin's from
/// the take-over, so the record goes then; the fresh one is begun by the tick, once no pin stands.
#[test]
fn a_pin_that_took_the_resident_over_re_warms_one_after_it_closes() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);

    let pid = a_resident(true);
    VIDEO_PID.store(pid, Ordering::SeqCst);
    mpv_resident::a_pin_took_it_over();

    assert_eq!(
        mpv_resident::the_resident_pid(),
        None,
        "the resident is the pin's now, and no longer this app's to serve with"
    );
    assert!(
        !mpv_resident::the_resident_is_owed_a_re_warm(true, true, false),
        "nothing is begun while the pin that took it stands"
    );

    scene_the_pin_closed();
    mpv_resident::tick();

    let warmed =
        mpv_resident::the_resident_pid().expect("a fresh resident once the pin has closed");
    assert_ne!(warmed, pid, "one begun anew, not the pin's");
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.contains("mpv resident begun pid")),
        "and the run says one was begun: {:?}",
        pin_trace::noted_lines()
    );
}

/// The pin that took the resident over has closed: its state is down and nothing of it stands.
fn scene_the_pin_closed() {
    stand_pin(None);
    VIDEO_PID.store(0, Ordering::SeqCst);
}

/// **A resident names the window it is begun with for the road that takes it out of the Shell's
/// lists, at the moment it is begun.**
///
/// The fault this is a regression test for is the user's: the warm mpv is begun with `--idle=yes
/// --force-window=yes` so that a hover has a window to draw into, and that window is a real top-level
/// window of the desktop's with no file in it — so the Shell gave it a taskbar button titled `No
/// file`, from the spawn until the resident was first picked up, or for ever where it never was. The
/// eager pass a launch writes its pid to is the pass that catches a window before its show, and a
/// resident's window is shown as soon as mpv has configured it: the pid is named for the one road that
/// can take a window out of the lists after it is on the screen, and this holds the name being written
/// where the resident is begun (see `ffplay_window::note_a_window_to_unlist` and `eager_ex_style`).
#[test]
fn a_resident_names_its_window_for_the_shells_lists() {
    let _room = the_room_to_itself();
    let _stand = ResidentStand::new(true);
    ffplay_window::forget_the_window_to_unlist();

    let pid = a_resident(true);

    assert_eq!(
        ffplay_window::the_window_to_unlist(),
        Some(pid),
        "the window the resident is begun with is the one the Shell must not list: a button titled \
         `No file` is not one the user asked for"
    );

    ffplay_window::forget_the_window_to_unlist();
}
