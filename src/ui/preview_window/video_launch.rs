//! How a video is asked of FFmpeg's player, as arguments rather than as a process.
//!
//! Everything here is a question with an answer that does not need a player to be running, which
//! is what makes these the seam the six video faults were fixed at. Three of them are faults in
//! *which arguments are passed* and no amount of window handling would have answered them:
//!
//! - The loop. `ffplay`'s own `-loop 0` and an input-side `-ss` cannot both be asked for. Measured
//!   on FFmpeg 9.0.2: `-loop 0` alone wraps an eight-second clip from `0` back to `0`, while
//!   `-loop 0 -ss 6` wraps it from `5.208` back to `5.208` — so the last two and a half seconds of
//!   the film A-B for ever and the first six are never seen. The seek is what breaks the loop, not
//!   where the seek is written: putting `-ss` after the input behaves the same way. So a player
//!   begun at a second is not given a loop at all, and the rewind is this app's own (see
//!   `loop_is_ours` and `rewind_due`).
//!
//! - Hardware acceleration, which is deliberately *not* among these arguments. `-hwaccel` names a
//!   *device* rather than asking for a decoder, and ffplay brings up its Vulkan renderer the moment
//!   the flag is present: measured here on FFmpeg 9.0.2, `dxva2`, `d3d11va`, `qsv` and `cuda` each
//!   die of an access violation in the graphics driver with no picture at all, while the same
//!   command without the flag plays the file through in software. So the switch is mpv's alone,
//!   where the safe list and the software fallback are the player's own (see
//!   `video_hw_accel_wanted`).
//!
//! - Subtitles. `-sst` is inert on this build: the subtitle stream is demuxed, no subtitle filter
//!   is put in the graph, and nothing is drawn. The route that works is `-vf subtitles=<file>`, which
//!   initialises libass — and which has three sharp edges, all measured rather than assumed. The
//!   filter's own argument is colon-separated, so a Windows drive letter is read as an option
//!   (`subtitles=D:/x.srt` fails with `Unable to parse "original_size" option value`) and has to be
//!   escaped inside a quoted value. And an input-side `-ss` makes the filter draw *nothing at
//!   all*: measured on a clip with a cue over 0.5s–3.0s, `-ss 2.0` before the input leaves the frame
//!   byte-identical to the same frame without the filter, while the same `-ss` after the input
//!   changes it. The third is the file it is pointed at: the filter draws text subtitles and refuses
//!   everything else, and a file it refuses — or one that is not there — is a filtergraph that fails
//!   to build, which takes the player down with it rather than falling back to no subtitles (see
//!   `subtitle_filter`, and `subtitle_files` for what it is therefore named).
//!
//! **And the filter is measured here to be nearly free, which is worth saying because it is the
//! obvious suspect for a slow start.** Timed on this machine at first window: a 1080p MKV with an
//! external sidecar came up in 303ms against 300ms for the same file with no filter at all, and a
//! 2160p HEVC MKV with its own embedded track came up in 346ms against 342ms — noise, at both sizes,
//! and the same for the bare spelling as for the `filename=` one, which are byte-identical frames.
//! Subtitles are a per-frame cost, not a per-start one.
//!
//! - The loop's rewind. FFmpeg's player has no key that goes to the beginning of a file, so a loop
//!   this app gives cannot be given by posting one (see `rewind_launch`).

use super::media_types::DerivedSubtitles;
use super::subtitles_wanted;
use super::{codecs, video_hw};
use crate::CONFIG;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;

/// The one binding an mpv video is given, which is the key this app already pauses a film with.
///
/// It is needed at all because of what mpv has *baked in*: `MBTN_LEFT_DBL` cycles fullscreen and
/// the wheel moves the volume, and a hover overlay must have neither — while `p cycle pause` does
/// have to be there, because the pin pauses an external player by posting that key (see
/// `video_key_pause`). `--input-conf` alone does not remove the built-ins, so the file is paired
/// with `--no-input-default-bindings` and the pair is dropped together where the file cannot be
/// written: a player with a live mouse beats a player that cannot be paused.
pub(super) const MPV_PAUSE_BINDING: &str = "p cycle pause\n";

/// The binding file written for a run of the app, if it could be written, and nothing where it
/// could not.
///
/// It is written once per run rather than once per launch — it is one line and nothing about it
/// changes — and it goes where the app puts every file it derives, under the temp folder
/// (`subtitle_files`' own two files are `general`, and the pages are `document`), which is the
/// folder a run's leftovers are cleared from at the next start (see
/// `document_cache::discard_leftovers`). A file left behind by a run that was killed is therefore
/// not read by the next one: both flags are dropped with this answer, so a stale file is not used.
static MPV_INPUT_CONF: std::sync::OnceLock<Option<PathBuf>> = std::sync::OnceLock::new();

pub(super) fn mpv_input_conf() -> Option<&'static Path> {
    MPV_INPUT_CONF
        .get_or_init(|| {
            let file = crate::engines::document_cache::temp_folder().join(MPV_INPUT_CONF_NAME);
            (write_input_conf(&file)).then_some(file)
        })
        .as_deref()
}

/// The name the binding file is written under.
const MPV_INPUT_CONF_NAME: &str = "mpv-input.conf";

/// Write the one binding out, answering whether it landed.
///
/// A write that cannot be made is not a failure worth announcing and not one worth stopping for:
/// what it costs is the pair of input flags, and what stands instead is mpv's own bindings with a
/// pause key that still works (see `MPV_PAUSE_BINDING`).
fn write_input_conf(file: &Path) -> bool {
    std::fs::write(file, MPV_PAUSE_BINDING).is_ok()
}

/// Whether a launch at `volume` is given no audio at all rather than a level of zero, which is
/// `--aid=no` and nothing else — no track at all, so the decoder is saved the whole soundtrack.
///
/// It is one rule asked in two places, and it is one function so that the two cannot come apart:
/// `mpv_args`, which pushes the flag from it, and the launch itself, which records the answer
/// beside the pipe so that a level moved on to such a player has a track to ask back for before it
/// can be heard (see `mpv_ipc::set_volume`).
pub(super) fn mpv_launched_without_audio(volume: u32) -> bool {
    volume == 0
}

/// The arguments an mpv video is begun with, which is the command line the spec settles on. A video
/// output is named, which it was not for years — why is on the flag itself (see below).
///
/// **No `--focus-on*` is among them**, which is not an omission: this list is the launch every road
/// gets, and a `--focus-on` is a hard start failure on a build old enough not to have it. What a
/// *hover* is told about the foreground is asked of the installed build once and appended after this
/// list, where nothing of a pin's launch can see it (see `hover_focus_switch` and
/// `the_focus_on_support`).
///
/// The switches are passed in rather than read here, because reading one at a hover's beginning
/// reads it for the *next* hover: a film already playing was launched with the answer that stood
/// then and is kept playing with it (see `video_hw_accel_wanted`).
///
/// **A muted film is given no audio rather than a level of zero**, which is the ffplay route's
/// `-an` and saves the decoder the whole soundtrack; and **a film with no measured gain is given
/// no filter at all**, because none has measured one yet and what measures it is a read on a
/// thread of its own (see `spawn_gain_scan`). What that first half leaves for a later hand is
/// `mpv_launched_without_audio`, which the launch records beside the pipe as well (see
/// `mpv_ipc::set_volume`).
#[allow(clippy::too_many_arguments)] // Each one is a fact about the launch, and a struct of them would be a type for one call.
pub(super) fn mpv_args(
    path: &Path,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    start: f64,
    hold: bool,
    volume: u32,
    gain: Option<f64>,
    hw_accel: bool,
    subtitles: bool,
    input_conf: Option<&Path>,
    ipc: &str,
    flip: bool,
    ontop: bool,
) -> Vec<String> {
    let mut args: Vec<String> = vec![
        // Read before the configuration and before any file the user has: the app's own command
        // line is the whole of what an mpv video is launched with, so a user's `mpv.conf` cannot
        // turn on an OSC or a mouse binding for a preview that is meant to have neither.
        path.to_string_lossy().into_owned(),
        "--no-config".to_string(),
    ];

    // The mouse: mpv's own bindings are what a hover overlay must not have, and they survive
    // `--input-conf` on their own — see `MPV_PAUSE_BINDING`, which is where both flags are
    // explained and where they are dropped together.
    if let Some(input_conf) = input_conf {
        args.push("--no-input-default-bindings".to_string());
        args.push(format!("--input-conf={}", input_conf.to_string_lossy()));
    }

    args.push("--no-osc".to_string());
    args.push("--no-border".to_string());
    // **And its window is not mpv's to drag, whatever happens to be over it.** mpv moves its own
    // window with the left button by default, and this app carries the pin — so a press that ever
    // reached mpv's window is a film dragged out of the hand and put back by the tick, the magnet
    // the film's own borders were reported as. The pin's window is the drag handle for every film
    // this app shows (see `pinned_press`), and it takes the press there because its band is the
    // pin's own pixel — which is a fact about where the two windows stand. This refusal is not.
    args.push("--window-dragging=no".to_string());
    // **And mpv holds its own window above everything from the frame it is born on**, which is the
    // half of the raise this app cannot answer for: the film it draws under a pin is *seen* through
    // the pin's own band rather than being a window of this app's, and a window left in the
    // ordinary band is covered by Explorer the moment the user clicks into it — while the band goes
    // on taking the press (see `player_order`). The app's own raise is the other half, and it is
    // asked for on every tick and on every monitor pass (see `place_player_window`): what makes
    // the two halves two rather than one is that this side holds the window up before this app has
    // ever found its handle, which is what the first frames of a pinned film are drawn under.
    //
    // **The order the player is held up in, where it is to hold one.** A top-level film is a
    // window of the desktop's that has to stay over whatever the pin stands in front of; a film
    // the pin shows through a DWM thumbnail is one that has to stay *behind* the window it is
    // seen in, so its launch asks for no band at all (see `pin_film_thumbnail`).
    if ontop {
        args.push("--ontop".to_string());
    }
    // The size and the position, always both, and always written `+x+y` whether x and y are
    // negative or not: a box on a monitor left of or above the primary one has a negative
    // coordinate, and `+-2160` is how that coordinate is spelled for a geometry whose offsets
    // are introduced by a sign — a bare `-2160` is a distance from the far edge instead, so the
    // film would be put somewhere other than the box. Measured on mpv 0.41.0, on a primary
    // display 3440x1440 with a second one at (3440,-2160), preview box (3540,-2060):
    // `+3540+-2060` puts the window on the second display, and the size alone centres it on the
    // primary at (1520,570).
    args.push(format!("--geometry={width}x{height}+{x}+{y}"));
    // **`--vo=gpu-next` is being measured here, not chosen.** It is libplacebo's renderer rather than
    // `vo=gpu`'s, with its own swapchain and its own present path, and what it is being tried for is
    // the two things the bitblt swapchain below has not settled: a band that goes black while a grown
    // host waits for the player to redraw its child at the new size, and a picture that judders inside
    // a box that is otherwise standing still. `--gpu-api=d3d11` and `--d3d11-flip=no` apply to it as
    // they did to `vo=gpu`. One flag, reverted by deleting it; what
    // a run of it wants looked at is the resize, the colours and the scale against `vo=gpu`, HDR files,
    // and `--hwdec=auto`.
    args.push("--vo=gpu-next".to_string());
    args.push("--gpu-api=d3d11".to_string());
    // **And the swapchain is the bitblt model rather than the flip one**: a flip-model swapchain
    // re-derives its buffers for a new size — a `ResizeBuffers` and a drain of the present queue,
    // at a cost proportional to the buffer's size — and a pin is resized on every pointer move, so
    // that re-derivation is a hitch on every step in a child window this app does not own (see
    // `pin_film_host`). The bitblt model presents by copying at the window's current size and has
    // no such queue to drain, which is the classic cure for resize shimmer in an embedded child
    // window. What it gives up is the independent flip the compositor hands straight to the
    // display: a film here costs a copy per frame and can sit a step behind a moving pin.
    //
    // **Except on the thumbnail road, where `flip` asks for the other model as an experiment.**
    // There the player makes a top-level window of its own and the pin shows it through a DWM
    // thumbnail, and a flip-model swapchain has a property the bitblt one does not: the compositor
    // scales the buffer to the window's client area, so the picture fills the window *by
    // construction* rather than because mpv caught a `WM_SIZE` and resized its buffer. With the
    // bitblt model a buffer that is a size behind the window is presented 1:1 — cropped where it is
    // larger, a small picture in black where it is smaller (see `pin_film_thumbnail`). It is asked
    // for by the knob and never by default. No other flag conflicts: `--gpu-api=d3d11` is the
    // context the flag names,
    // `--vo=gpu-next` presents through the same DXGI swapchain, and `--hwdec=auto`'s d3d11va
    // interop is what a flip model is designed for.
    //
    // **And never on a window made invisible either**: a layered window cannot have a flip-model
    // swapchain, so `flip` is `false` there by the caller's own decision and the launch says the knob
    // was overruled rather than dropping it (see `the_thumbnail_swapchain_flip`).
    args.push(
        if flip {
            "--d3d11-flip=yes"
        } else {
            "--d3d11-flip=no"
        }
        .to_string(),
    );

    // mpv decodes on the card by its own safe list with its own software fallback, so the switch
    // is the whole of the answer here: there is no device for this side to measure and none to
    // name (see `video_hw_accel_wanted`, which is mpv's switch alone).
    args.push(if hw_accel {
        "--hwdec=auto".to_string()
    } else {
        "--hwdec=no".to_string()
    });

    // Subtitles are mpv's own to find and choose — its own tracks and the sidecars beside the
    // film — so the `Subtitles` switch is answered by turning them off rather than by naming a
    // track: with it off the film must show none, or the tray row is a lie for every user who has
    // mpv. `--sub-auto=no` is what stops mpv loading the sidecars as tracks at all, and `--sid=no`
    // is what stops it selecting one of the film's own.
    if subtitles {
        args.push("--sub-auto=fuzzy".to_string());
    } else {
        args.push("--sid=no".to_string());
        args.push("--sub-auto=no".to_string());
    }

    // The film stops on its last frame rather than exiting, which is what the app's own loop is
    // for: nothing this app can observe would say a pass had ended otherwise, so the tick that
    // measures the film off its own clock is what begins it again (see `video_loop_tick`).
    args.push("--keep-open=yes".to_string());

    // **And it *stops* there rather than pausing there.** mpv's default end of file sets `pause`,
    // which makes the press that plays a finished film again a seek *and* an unpause: the seek is
    // answered against a player that is paused, and the picture blinks out for the moment the
    // restart takes before the film is playing again. That is the one visible difference between
    // that press and the loop's wrap, which seeks a film that is playing (see
    // `toggle_pinned_playback` and `video_loop_tick`). A film that is stopped and not paused is
    // resumed by the seek alone — mpv's own answer for this option — and the end of a pass is read
    // off this app's clock either way, so nothing here wants the player paused (see `pass_ended`).
    args.push("--keep-open-pause=no".to_string());

    // A player begun at the beginning is given no `--start` rather than `--start=0`, so that
    // nothing is asked for that would be answered with the position it is already at.
    if start > 0.0 {
        args.push(format!("--start={start:.3}"));
    }

    // **A player begun on a frame another player is showing is begun standing on it.** A pin taken
    // up from a hover begins its own player at the exact second the hover was frozen on, and the
    // hover's frame is the picture on screen until that player can show the same one — so the seek
    // is exact rather than to the keyframe before the second, and the player is paused, because a
    // player begun playing would have advanced off the frame by the time the handover asks for it
    // (see `pin_film_host::settle_pin_handover`). Both are per-file options, which is why they are
    // written beside `--start` and not at the top of this list.
    if hold {
        args.push("--hr-seek=yes".to_string());
        args.push("--pause".to_string());
    }

    if mpv_launched_without_audio(volume) {
        args.push("--aid=no".to_string());
    } else {
        args.push(format!("--volume={volume}"));

        // The measured gain, as the same libavfilter `volume` filter the ffplay route applies it
        // with — named through mpv's bridge to libavfilter, which is the documented spelling of a
        // libavfilter graph: the bare `--af=volume=` reaches the same filter, but only through
        // mpv's automatic bridge for a name it has no filter of its own for, and the manual says
        // that bridge neither verifies its arguments nor answers for them (see `--vf`'s `lavfi`).
        // A muted film is not normalised, since there is nothing to hear.
        if let Some(gain) = gain {
            args.push(format!("--af=lavfi=[volume={gain:.4}]"));
        }
    }

    // The pipe this player is told things on while it runs, which is the whole of what makes the
    // mpv route not FFmpeg's: a volume, a seek, a pause and a track step are written down it rather
    // than paid for with a player ended and begun again (see `mpv_ipc`). The name is unique to this
    // launch, so two previews can never be told things on one pipe.
    args.push(format!(r#"--input-ipc-server=\\.\pipe\{ipc}"#));

    args
}

/// **A prototype's knob, and not a setting: `RHP_THUMBNAIL_FLIP=1`.**
///
/// The thumbnail road asks mpv for a flip-model swapchain instead of the bitblt one (see `mpv_args`),
/// which makes the picture fill its window by construction rather than by catching a `WM_SIZE`: the
/// compositor scales the buffer to the client area. The value is exactly `1`, because this is an A/B
/// of two runs and anything else must be the default road — a run that spells it otherwise is a run
/// whose flags did not change, which is a wasted measurement and not a surprise to be discovered
/// later.
const THUMBNAIL_FLIP_VARIABLE: &str = "RHP_THUMBNAIL_FLIP";

/// Whether the machine asked for the flip-model swapchain on the thumbnail road (see
/// [`THUMBNAIL_FLIP_VARIABLE`]). Read once per launch, which is the moment the answer is true for:
/// a player already begun keeps the model it was begun with.
pub(super) fn the_thumbnail_flip_knob() -> bool {
    thumbnail_flip_from(std::env::var_os(THUMBNAIL_FLIP_VARIABLE).as_deref())
}

/// The knob's answer for one value: `1` is on and every other value — including none — is off.
fn thumbnail_flip_from(value: Option<&std::ffi::OsStr>) -> bool {
    value.is_some_and(|value| value == std::ffi::OsStr::new("1"))
}

/// **The swapchain a thumbnail road's player is really asked for**, which is the knob *unless* the source
/// is made invisible.
///
/// A flip-model swapchain is not a thing a layered window can have: DXGI creates a flip-model swapchain
/// for the window and the compositor hands it straight to the display, which is the arrangement
/// `WS_EX_LAYERED`'s own blending is the opposite of. So where the invisible road is on
/// (`RHP_THUMBNAIL_INVISIBLE`, see `pin_film_thumbnail`) the bitblt swapchain is asked for whatever
/// `RHP_THUMBNAIL_FLIP` says, and the launch traces that the knob was read and overruled rather than
/// dropping it on the floor (see `video_playback::start_mpv_video_playback`).
pub(super) fn the_thumbnail_swapchain_flip(asked: bool, invisible: bool) -> bool {
    asked && !invisible
}

/// Which of the two spells for "do not take the foreground" the installed mpv recognises.
///
/// **It is a measurement of the build on this machine rather than of any documented option set.**
/// `--focus-on` is the current spelling (its `never` is one of its choices), and `--focus-on-open`
/// is the older one mpv now lists as removed — a build may still reject it, and this machine's does
/// (see `the_builds_focus_on_answer`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum FocusOnSupport {
    /// The installed mpv accepted `--focus-on=never`.
    Never,
    /// It refused that and accepted the old `--focus-on-open=no`.
    OpenOnly,
    /// It refused both: a hover's launch asks for nothing.
    Neither,
}

/// The one switch a hover's mpv launch names for the foreground, or nothing where this build
/// recognises neither spelling (see `FocusOnSupport`).
///
/// Hover-only, and a pure answer about the build so that what the launch appends is a value a test
/// can hold: a pin's launch is the command line every road had before this existed
/// (see `video_playback::start_mpv_video_playback`).
pub(super) fn hover_focus_switch(support: FocusOnSupport) -> Option<&'static str> {
    match support {
        FocusOnSupport::Never => Some("--focus-on=never"),
        FocusOnSupport::OpenOnly => Some("--focus-on-open=no"),
        FocusOnSupport::Neither => None,
    }
}

/// What the installed build answered when it was asked once, or nothing while it is being asked (see
/// `probe_focus_on_once`).
static FOCUS_ON_SUPPORT: Mutex<Option<FocusOnSupport>> = Mutex::new(None);

/// Whether the one question has been asked, so a launch can never start a second probe.
static FOCUS_ON_ASKED: AtomicBool = AtomicBool::new(false);

/// Ask the installed mpv, once, which `--focus-on*` switch it recognises, and keep the answer.
///
/// **It is asked off every hover's road** — the answer is the same for the whole run, and a preview
/// has no time to spend on a question about the player it is about to begin. It is asked at the
/// preview window's start, and it finds the player on a thread of its own there, so that by the time
/// a hand is over a film the answer is standing (see `event_loop::run_preview_window`); a launch that
/// finds no answer yet adds nothing, which is the command line this app always had.
///
/// shortcut: the answer is kept for the run whatever build it was measured on, so a machine that
/// swaps `mpv.exe` mid-run keeps the old build's answer; re-probe when the setting can change while
/// the app runs.
pub(super) fn probe_focus_on_once() {
    if FOCUS_ON_ASKED.swap(true, Ordering::AcqRel) {
        return;
    }

    std::thread::spawn(|| {
        let chosen = CONFIG
            .lock()
            .map(|config| config.mpv_path.clone())
            .unwrap_or_default();
        let Some(program) = video_hw::mpv_program(&chosen, &codecs::mpv_installs()) else {
            return;
        };
        let support = the_builds_focus_on_answer(&program);
        if let Ok(mut held) = FOCUS_ON_SUPPORT.lock() {
            *held = Some(support);
        }
    });
}

/// The answer the installed build gave, or `Neither` while it is still being asked (see
/// `probe_focus_on_once`).
pub(super) fn the_focus_on_support() -> FocusOnSupport {
    FOCUS_ON_SUPPORT
        .lock()
        .ok()
        .and_then(|held| *held)
        .unwrap_or(FocusOnSupport::Neither)
}

/// Stand the cached answer where a test needs it, without beginning a player (see
/// `the_focus_on_support`).
#[cfg(test)]
pub(super) fn pretend_the_build_answers(support: FocusOnSupport) {
    FOCUS_ON_ASKED.store(true, Ordering::Release);
    if let Ok(mut held) = FOCUS_ON_SUPPORT.lock() {
        *held = Some(support);
    }
}

/// The build's own answer, asked with the switch on a command line it has to parse.
///
/// **`--version` is no use as the probe**: measured on this machine's mpv 0.41.0, `--version`
/// short-circuits before the option parser, so `mpv --no-config --version --focus-on=never` and the
/// same line with a `--focus-on-open=no` this build does not have both exit 0. `--input-test` parses
/// the options and then ends, which is the run that answers the question: measured here, that line
/// with `--focus-on=never` exits 0, and with `--focus-on-open=no` — and with any option this build
/// does not know — exits 1. It is a start of the player and nothing else, so it costs about a tenth
/// of a second and decodes nothing.
fn the_builds_focus_on_answer(program: &Path) -> FocusOnSupport {
    if mpv_accepts_the_switch(program, "--focus-on=never") {
        FocusOnSupport::Never
    } else if mpv_accepts_the_switch(program, "--focus-on-open=no") {
        FocusOnSupport::OpenOnly
    } else {
        FocusOnSupport::Neither
    }
}

/// Whether the player at `program` parsed `switch` rather than refusing it: the run ends cleanly when
/// the option exists and fails when it does not (see `the_builds_focus_on_answer`).
fn mpv_accepts_the_switch(program: &Path, switch: &str) -> bool {
    super::engine_processes::hidden_command(program)
        .args(["--no-config", "--no-terminal", "--input-test", switch])
        .stdin(std::process::Stdio::null())
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .status()
        .map(|status| status.success())
        .unwrap_or(false)
}

/// The size a film's box is asked for, so that the window the film is drawn in is that box.
///
/// **A window is asked for on the primary display's scale, and Windows rescales it to the display
/// it is moved to — so the size mpv is told is not the size the window ends up.** mpv creates its
/// window at `CW_USEDEFAULT`, which is the primary whatever display the film is watched on, and
/// applies the geometry exactly there; a window crossing to a display of another scale then gets
/// `WM_DPICHANGED` from Windows with a suggested rectangle that keeps the window's *logical* size,
/// and mpv's handler takes that rectangle as it stands (`w32->windowrc = *rc;` under
/// `case WM_DPICHANGED` in `video/out/w32_common.c` — it does not put the geometry back, and no
/// option of that version makes it). The window therefore comes out at the size asked for times
/// the display's scale over the primary's, which is the film at 200% of the preview box.
///
/// Measured on mpv 0.41.0 with the primary display 3440x1440 at 96 DPI and a second one 3840x2160
/// at 192 DPI at (3440,-2160), reading `GetWindowRect` every millisecond from a per-monitor-v2
/// process with the pointer on the second display: `--geometry=2864x1611+4416+-2160` is *visible*
/// at 5728x3222 and stays there — twice the box, wider than the display, and the wrong size is the
/// one shown, never a frame of the right one — and `--geometry=480x270+4160+-1755` is visible at
/// 960x540. Asked for on the primary's scale the window is the box from the first frame it is
/// visible: 1432x805 opens at 2864x1610, 240x135 at 480x270 exactly, and the same box on the
/// primary display opens at the box unchanged. The three flags that look like they should matter —
/// `--hidpi-window-scale=yes`, `--force-window-position=yes`, `--auto-window-resize=no` — each
/// leave it doubled (one run each), and `--screen=1` moves the offsets onto that display's origin
/// without touching the scale.
///
/// shortcut: the divisor is the ratio of the two displays' DPIs and nothing finer, so a box whose
/// size does not halve evenly lands a pixel out until the tick asserts the box (see
/// `ensure_video_window_topmost`); a display drawn at the primary's scale — the primary itself
/// among them — is asked for its box as it stands, which is every box on a machine of one scale;
/// and this rests on mpv's window being created on the primary display, which is the one thing
/// here that is mpv's to change. Upgrade when mpv re-applies the geometry after a DPI change, or
/// when Windows stops suggesting a rescaled rectangle.
pub(super) fn mpv_box_size(
    width: i32,
    height: i32,
    display_dpi: u32,
    primary_dpi: u32,
) -> (i32, i32) {
    if display_dpi == primary_dpi {
        return (width, height);
    }

    let on_the_primarys_scale = f64::from(primary_dpi) / f64::from(display_dpi);

    (
        (f64::from(width) * on_the_primarys_scale).round().max(1.0) as i32,
        (f64::from(height) * on_the_primarys_scale).round().max(1.0) as i32,
    )
}

/// The extensions a sidecar subtitle file is looked for under, in the order they are tried.
///
/// SubRip first because it is what a sidecar is nearly always, and `.ass`/`.ssa` after it because
/// an anime release brought its own styles alongside the video is the case where a sidecar is
/// written in them. The order is a preference and not a requirement: at most one of the three is
/// present for any given name, and where two are, the SubRip one is the plainer of the two.
const SIDECAR_EXTENSIONS: [&str; 3] = ["srt", "ass", "ssa"];

/// The subtitle file a video is shown with, if it has one at all.
///
/// **Two sources, and only ever small files: a sidecar beside the film, or a track this app
/// copied out of it** (see `subtitle_files`). The first of the two that exists wins, and where
/// neither exists the answer is nothing — a film played without subtitles.
///
/// **The film's own stream is never named, and that is a refusal with a measurement behind it.**
/// A `subtitles` filter pointed at the film opens it a second time and streams the whole
/// container to its first subtitle before the first frame is drawn: measured, a cold 1.4 GB MKV
/// took 14 904 ms and read 1 423 MB that way, against 492 ms and 41 MB drawn from the 30 KB
/// `.ass` copied out of it. The player's start is waited for under `VIDEO_START_WAIT_SECS` (see
/// `player_wait`), so a format that costs fifteen seconds before its first frame is a player
/// this app gives up on and kills — a preview that never appears, every time the film is
/// hovered, and — the failure being cached — permanently. A copy that is not there yet, a copy
/// that has failed and a film whose only track has no small form are all answered with no
/// subtitles: the fast hover is the user's own choice, and no subtitle is worth the picture.
///
/// **And only a file libass can draw is ever named.** The filter renders text subtitles and
/// refuses every other kind outright (its own check is `AV_CODEC_PROP_TEXT_SUB`, logged as "Only
/// text based subtitles are currently supported"), and a filter that refuses the file it was
/// given is a filtergraph that fails to build: the player exits without ever putting a window
/// up, which is this same no-preview fault by a second road. Both sources are therefore gated on
/// a name the filter can read — the sidecar by the extensions it is looked for under (see
/// `sidecar_for`), and the copy by the extensions the pass writes (see `subtitle_files::resolve`).
///
/// **The named file is checked before it is named**, and that check is not the folder walk the
/// cached answer exists to save: a sidecar renamed away beside an unchanged film leaves a cached
/// answer that names a file no longer there, and a filter naming a missing file fails to load —
/// the player exits without a window, which is the preview that never appears. A stat of a named
/// path is the whole of the check; it opens no directory and reads no file.
///
/// The index handed over is the *subtitle-relative* one, counted from zero among subtitle
/// streams (see `SubtitleStreams`), and it chooses which copied file is named rather than a `si=`
/// specifier: a copied track is a whole file with no streams of its own, exactly as a sidecar is.
/// An index the answer does not carry — a track that was never copied, or a number the cache
/// remembers from before the film changed — falls back to the first copy rather than to nothing,
/// because the choice is worth less than the subtitles (see
/// `a_track_the_answer_does_not_carry_is_answered_with_the_first_copy_rather_than_with_nothing`
/// in the tests).
pub fn subtitle_filter(
    sidecar: Option<&Path>,
    derived: Option<&DerivedSubtitles>,
    track: Option<usize>,
) -> Option<String> {
    // A sidecar carries no streams of its own for a track to name: the whole of that file is the
    // subtitle, which is why the filter beside it has no `si=` at all. It is named only where it
    // is still there, which is the check the rename fault is closed with (see the note above).
    if let Some(sidecar) = sidecar
        .filter(|sidecar| sidecar.is_file())
        .and_then(spelled)
    {
        return Some(format!("subtitles='{sidecar}'"));
    }

    // The track this app copied out of the film, where the extraction wrote one and it is still
    // there. Like a sidecar it is a whole file with no streams of its own, so it carries no
    // `si=`; unlike a sidecar, the container's own fonts are drawn with it out of the folder the
    // same pass dumped them into, because the styles an embedded track is written in name faces
    // the film carries and the machine may not have (see `subtitle_files`).
    let derived = derived?;
    let track = track
        .filter(|track| *track < derived.tracks.len())
        .unwrap_or(0);
    let track_file = derived
        .track(track)
        .filter(|track_file| track_file.is_file())
        .and_then(spelled)?;

    let fonts = derived
        .fonts
        .as_deref()
        .filter(|fonts| fonts.is_dir())
        .and_then(spelled);

    Some(match fonts {
        Some(fonts) => format!("subtitles='{track_file}':fontsdir='{fonts}'"),
        None => format!("subtitles='{track_file}'"),
    })
}

/// The sidecar beside `path` that carries subtitles, if one does.
///
/// The extension is swapped rather than the name appended: a video called `episode.mkv` beside
/// `episode.srt` is the arrangement every player and every tool on this machine reads, and one
/// beside `episode.mkv.srt` is not read by any of them. The video's own extension is put back
/// first, so `a.b.mkv` looks for `a.b.srt` rather than for `a.srt` beside `b.mkv`.
///
/// **Where the swap finds nothing, the folder is read once for a sidecar carrying a language in its
/// own name** — `episode.en.srt`, `episode.eng.ass` — because that is the spelling a release
/// actually ships and the plain swap never once looked at one. A release writes the language into
/// the file precisely because there is more than one language in the folder, and a folder written
/// that way was answered with the container's own track at best and with nothing at all where the
/// container held none: a subtitle file sitting right there, spelled correctly, never opened.
///
/// The language is read off the name rather than off a list of codes, because the codes are
/// unbounded — a release that writes `episode.en.srt` will as happily write `episode.eng.srt`, and
/// both are the same arrangement — so guessing at a table would be a rule that quietly does not
/// apply to the next folder. One read of the video's own folder is the whole of it, measured at
/// 1.9ms over a folder of 1200 files against a player launch that costs 300ms, and it is reached
/// only once the plain swap has already found nothing. There is no walk: a hover must not cost a
/// tree, and a sidecar beside some *other* film is not this film's subtitles.
///
/// The appended spelling is refused even though it has the very shape this rule matches, and the
/// video's own extension after the dot is what tells the two apart: `episode.mkv.srt` is read by no
/// player, so drawing it would be drawing a file the user did not mean.
///
/// It is asked of the probe thread and held with the geometry rather than asked of at every launch
/// (see `probe_video_geometry`), which is the only change made to it: the walk itself is a
/// `read_dir` of one folder, and what it costs is the size of that folder rather than anything
/// about the film.
pub(super) fn sidecar_for(path: &Path) -> Option<std::path::PathBuf> {
    // The gate: the probe is this function's only caller and asks it only where
    // subtitles are wanted (see `probe_video_geometry`), and the answer is kept
    // honest here as well, so that the walk of the film's folder is one no other
    // road can start for a file nothing will draw beside the film.
    if !subtitles_wanted() {
        return None;
    }

    let stem = path.with_extension("");

    if let Some(sidecar) = SIDECAR_EXTENSIONS
        .iter()
        .map(|extension| stem.with_extension(extension))
        .find(|candidate| candidate.is_file())
    {
        return Some(sidecar);
    }

    // Case-blind throughout, because a Windows path opens the file either way and a rule that only
    // held for one spelling of a name would be a rule that stopped working when a folder was
    // re-saved by something that title-cased it.
    let own_extension = path
        .extension()
        .map(|extension| extension.to_string_lossy().to_lowercase());
    let prefix = format!("{}.", stem.file_name()?.to_string_lossy().to_lowercase());
    let folder = path
        .parent()
        .filter(|folder| !folder.as_os_str().is_empty())
        .unwrap_or(Path::new("."));

    let mut named: Vec<std::path::PathBuf> = std::fs::read_dir(folder)
        .ok()?
        .filter_map(|entry| {
            let entry = entry.ok()?;
            if !entry.file_type().is_ok_and(|kind| kind.is_file()) {
                return None;
            }

            let candidate = entry.path();
            let extension = candidate.extension()?.to_string_lossy().to_lowercase();
            if !SIDECAR_EXTENSIONS.contains(&extension.as_str()) {
                return None;
            }

            let stem = candidate.file_stem()?.to_string_lossy().to_lowercase();
            let tag = stem.strip_prefix(&prefix)?;

            (!tag.is_empty() && Some(tag) != own_extension.as_deref()).then_some(candidate)
        })
        .collect();

    // Sorted, because a folder holding `episode.de.srt` and `episode.en.srt` has two right answers
    // and no preference to choose between them: nothing here reads a language off a track, so the
    // order is the one that does not change when the folder is read again.
    named.sort();
    named.into_iter().next()
}

/// The characters of a path that the filtergraph reads as its own punctuation, and which therefore
/// have to be escaped in one.
///
/// The set is measured rather than assembled from the grammar, and it was measured by the one
/// instrument that settles the question: `vf_subtitles` logs the filename it was handed, verbatim,
/// when it cannot open it. So each of these was put in a path and read back out of that log. A
/// colon, a comma, a semicolon, a bracket, an equals sign and a percent all come back as themselves
/// once backslash-escaped, and so do a space, an ampersand, a hash and a plus — the last four are
/// named here because a Windows path may hold any of them and none of them means anything here,
/// which is worth knowing rather than assuming.
///
/// The percent is the one whose escaping is belt and braces: measured on this build it survives
/// either way, because the expansion a timeline expression gets is not applied to a value inside
/// single quotes. It is escaped anyway, because `%` is the filtergraph's own expansion character
/// and a lone one is not a spelling worth depending on.
///
/// **The apostrophe is missing from this list on purpose.** It is the eighth character the two
/// parsers read as punctuation and the only one of them that cannot be spelled — see
/// `escape_filter_path`, which is where that is settled, and `spelled`, which is what refuses.
const FILTERGRAPH_PUNCTUATION: [char; 7] = [':', ',', ';', '[', ']', '=', '%'];

/// A path as the `subtitles` filter's argument has to have it: escaped, and not quoted.
///
/// The escaping is the filter's own rather than the shell's, and there are two of them stacked: the
/// filtergraph parser reads the chain and strips the quotes, and the filter's own option parser
/// then reads what is left — where the quotes are gone and a bare `\` is an escape again. So the
/// backslash put in here is spent by the second parser rather than by the first, which is exactly
/// why the value is worth quoting and why the two spellings do not behave alike: `subtitles='D\:/x.srt'`
/// draws and `subtitles=D\:/x.srt` fails with `Unable to parse "original_size" option value`
/// (measured on FFmpeg 9.0.2), because outside quotes the first parser consumes the backslash as
/// its own escape and hands the second parser a bare colon to read as the option separator.
///
/// Separators become forward slashes rather than being escaped: a Windows path opens with a
/// backslash, every path API on Windows takes a forward one, and a backslash left in the value would
/// be spent as an escape by the second parser rather than read as part of a name.
///
/// **An apostrophe comes back out of here as a backslash followed by nothing this can fix**, and
/// that is not a formality: inside the quotes a `\` is literal, so a `\'` ends the quoted run and
/// leaves a bare apostrophe behind, and a bare apostrophe is what the second parser then reads as
/// the start of another quoted run. Measured, the outcome is worse than a failure — `it's.mkv`
/// arrives as `its.mkv`, so the filter does not complain, it opens a file that is not there. So this
/// says nothing about how to escape one, and `spelled` is what refuses it.
pub fn escape_filter_path(path: &Path) -> String {
    let mut escaped = String::with_capacity(path.to_string_lossy().len());

    for character in path.to_string_lossy().chars() {
        match character {
            '\\' => escaped.push('/'),
            punctuation if FILTERGRAPH_PUNCTUATION.contains(&punctuation) => {
                escaped.push('\\');
                escaped.push(punctuation);
            }
            plain => escaped.push(plain),
        }
    }

    escaped
}

/// A path as the filter's argument can be spelled, or nothing where it cannot.
///
/// The one refusal is the apostrophe (see `escape_filter_path`), and it is a refusal rather than a
/// substitution because there is no second spelling to offer: a filter naming a file that is not
/// there fails to load, and one naming a *different* file that is not there does the same while
/// looking like it worked.
fn spelled(path: &Path) -> Option<String> {
    (!path.to_string_lossy().contains('\'')).then(|| escape_filter_path(path))
}

/// Whether the player this video is begun with has to be given its loop by this app.
///
/// True exactly when the launch named a second for it. `-loop 0` is correct on its own and is
/// still asked for in that case, because a player that wraps itself never reaches its own end and
/// so never exits out from under a hover — which is the ordinary case, a hover always starting at
/// the beginning.
///
/// It is wrong the moment `-ss` is in the same command, and it is wrong *badly* rather than
/// subtly: the wrap lands near the seek rather than at the beginning, so an eight-second clip begun
/// at six plays its last two and a half seconds for ever. Nothing in FFmpeg's own options expresses
/// "start at six, then loop the whole film", so the loop is taken away rather than corrected.
pub fn loop_is_ours(seeked: bool) -> bool {
    seeked
}

/// The second a loop this app gives is begun again at, which is the only second a loop can restart
/// at and the whole of why the rewind is a relaunch rather than a key.
///
/// A relaunch, and nothing else, because FFmpeg's player has no key that goes to the beginning of
/// a file. That is not an absence this app has to work around by guessing which key is nearest; it
/// is the shape of the player's whole seek table, read out of its own source for the exact build on
/// this machine (`fftools/ffplay.c`, tag `n9.0.2`, `event_loop`): every key bound to a seek is a
/// *fixed increment* — `Left` and `Right` are ten seconds, `Up` and `Down` are sixty, `PageUp` and
/// `PageDown` are six hundred or a chapter — and not one of them names a second. There is no `Home`,
/// no `Ctrl`-home, nothing absolute at all.
///
/// So a rewind posted as a key is a step backwards, and how far backwards depends on the length of
/// the file rather than on where in it the film is. `Down` on a clip shorter than a minute clamps to
/// the start and looks exactly like the go-to-start everyone expects it to be; on a film of ten
/// minutes it is a seek to nine minutes, and the pass after that is a seek to eight, so the film
/// walks backwards a minute at a time and never plays the first nine minutes again — while this
/// app's own clock, having been zeroed, says it is at the beginning the whole time. A bar that
/// disagrees with the picture by nine minutes is worse than no bar.
///
/// So the film is begun again, at zero, with no `-ss` — which is the launch a hover already gets and
/// the one FFmpeg's own `-loop 0` cannot be given beside a seek (see `loop_is_ours`). It costs a
/// window coming down and going up once per pass, which is a gap in the loop on a large file. That
/// cost is paid on purpose: it is one gap per pass in exchange for the pass being the whole film,
/// and a gap is at least honest about where the picture is.
pub fn rewind_launch_seconds() -> f64 {
    0.0
}

/// How far before the end of the file the rewind is posted, in seconds.
///
/// A rewind is a relaunch, and a relaunch is begun rather than posted — but it is begun *early*, for
/// the reason that holds however it is done: a player begun at the last frame of a file has already
/// closed its window by the time anything could be told, so a rewind at the end is not a rewind but
/// a relaunch of a player that has exited. A tenth of a second is wide enough to cover a tick of
/// this app's own loop and narrow enough that the film is not visibly cut short on a loop a user is
/// watching: at the cadence the pinned window is kept in step at, that margin is several ticks.
pub const REWIND_MARGIN_SECS: f64 = 0.1;

/// Whether the player should be sent back to the beginning now.
///
/// The four facts are the whole of the question, and three of them are refusals rather than
/// permissions. A player that is not playing is not rewound, because a held film that jumps to its
/// beginning on the tick is a film the user's hold did nothing to. A file whose length nothing has
/// read cannot be rewound, because the whole question is "how near the end is it" and a film of
/// unknown length has no end to be near — and a player begun at zero of such a file has nothing to
/// be saved from. And a file shorter than the margin is never rewound at all, because every second
/// of it is inside the margin and the player would be sent back the instant it began.
pub fn rewind_due(playing: bool, duration: Option<f64>, position: f64) -> bool {
    if !playing {
        return false;
    }

    let Some(duration) = duration.filter(|d| *d > REWIND_MARGIN_SECS) else {
        return false;
    };

    position >= duration - REWIND_MARGIN_SECS
}

#[cfg(test)]
mod tests {
    use super::super::subtitle_files::VideoSubtitlesSetting;
    use super::*;
    use std::path::PathBuf;

    /// The pipe name every `mpv_args` call below is handed, and the flag that name becomes — the
    /// one flag the mpv route has that FFmpeg's player has never been given.
    const IPC: &str = "test-pipe";
    const IPC_ARG: &str = r"--input-ipc-server=\\.\pipe\test-pipe";

    /// A Windows path survives the filter that reads it, which is a statement about the escaping
    /// rather than about the filter: the drive letter is a colon, and the filter's own separator is
    /// a colon.
    #[test]
    fn a_windows_path_is_escaped_for_the_filter_that_reads_it() {
        assert_eq!(
            escape_filter_path(&PathBuf::from(r"D:\video\clip.srt")),
            "D\\:/video/clip.srt",
            "the drive colon must be escaped and the separators need no escape of their own"
        );
    }

    /// Every character of the filtergraph's own punctuation survives the escaping, and every other
    /// character a Windows path may hold is left exactly as it was.
    ///
    /// This is the regression test for a path that quietly stops working. The drive colon was the
    /// only one that had been handled, on the reasoning that a Windows path opens with a colon and
    /// the filter's own separator is a colon — which is true, and is not the whole of it. A
    /// filtergraph reads a comma as the end of a chain, a semicolon as the end of a chain, a bracket
    /// as a filter's label, an equals sign as a filter's label, and a percent as its own expansion
    /// character, and a path holding any of them takes the whole chain down with it.
    ///
    /// Every one of these was measured against FFmpeg 9.0.2 by the filename the filter logs when it
    /// cannot open it, which is the only reading of the question that is not a guess: each character
    /// was put in a path and read back out of that log, and each came back as itself. The four in
    /// the second half were measured the same way and are asserted here because a Windows path may
    /// hold any of them and they must be left alone.
    #[test]
    fn every_character_the_filtergraph_reads_as_punctuation_survives_the_escaping() {
        for (character, escaped) in [
            (':', "\\:"),
            (',', "\\,"),
            (';', "\\;"),
            ('[', "\\["),
            (']', "\\]"),
            ('=', "\\="),
            ('%', "\\%"),
        ] {
            let path = PathBuf::from(format!("D:\\films\\a{character}b.srt"));

            assert_eq!(
                escape_filter_path(&path),
                format!("D\\:/films/a{escaped}b.srt"),
                "the filtergraph reads `{character}` as its own punctuation, so it has to reach the \
                 filter escaped; unescaped it ends the chain and takes the picture down with it"
            );
        }

        for character in [' ', '&', '#', '+', '(', ')'] {
            let path = PathBuf::from(format!("D:\\films\\a{character}b.srt"));

            assert_eq!(
                escape_filter_path(&path),
                format!("D\\:/films/a{character}b.srt"),
                "`{character}` is in a Windows path and means nothing to a filtergraph, so escaping \
                 it would put a backslash in the filename and name a file that is not there"
            );
        }
    }

    /// A subtitle file with an apostrophe in it is refused rather than spelled, which is the whole
    /// of what an unspellable path gets.
    ///
    /// This is the one character of the seven that cannot be escaped, and the reason is that there
    /// are two parsers and the apostrophe is punctuation for both. Inside the quotes a backslash is
    /// literal, so escaping one ends the quoted run and hands the second parser a bare `'` to read
    /// as the start of another one — and the result is not an error. Measured against FFmpeg 9.0.2,
    /// `it's.mkv` arrives at the filter as `its.mkv`, so it opens a different file and says nothing
    /// about it. A filter that draws the wrong file is worse than no filter, which is why this
    /// refuses the path rather than looking for a second spelling.
    #[test]
    fn a_path_that_cannot_be_spelled_is_refused_rather_than_silently_changed() {
        let dir = std::env::temp_dir().join(format!("hl-apostrophe-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let sidecar = dir.join("episode's cut.srt");
        std::fs::write(&sidecar, b"1\n").expect("a stand-in sidecar is writable");

        assert_eq!(
            subtitle_filter(Some(&sidecar), None, None),
            None,
            "an apostrophe cannot be escaped through both parsers: escaped, it survives as a bare \
             `'` and the filter opens `episodes cut.srt` instead, quietly. A film with subtitles \
             this app cannot name is a film played without them, which is not the same fault as a \
             film played with the wrong ones"
        );
        assert_eq!(
            spelled(&sidecar),
            None,
            "the refusal belongs to the spelling, so nothing that asks for one can be given one"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A subtitle file is named inside quotes, which is not decoration but the difference between
    /// the two ways of escaping a drive letter working and one of them silently not.
    ///
    /// Measured against FFmpeg 9.0.2: `subtitles='D\:/dir/x.srt'` draws and `subtitles=D\:/dir/x.srt`
    /// does not. Outside quotes the filtergraph parser consumes the backslash as its own escape and
    /// hands the option parser the bare colon, which is then read as the separator.
    #[test]
    fn a_sidecar_is_named_inside_quotes_so_the_escape_survives() {
        let dir = std::env::temp_dir().join(format!("hl-quoted-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let sidecar = dir.join("episode.srt");
        std::fs::write(&sidecar, b"1\n").expect("a stand-in sidecar is writable");

        let filter = subtitle_filter(Some(&sidecar), None, None)
            .expect("a clip with a sidecar beside it has a subtitle filter to draw it with");

        assert!(
            filter.starts_with("subtitles='") && filter.ends_with("'"),
            "the value has to be quoted or the filtergraph parser eats the escape that protects \
             the drive colon, which is the difference between the two spellings: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A video with a sidecar beside it is shown with the sidecar, and the sidecar's own name is
    /// the one asked for rather than the video's.
    #[test]
    fn a_sidecar_beside_the_video_is_preferred_over_the_files_own_tracks() {
        let dir = std::env::temp_dir().join(format!("hl-sidecar-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let sidecar = dir.join("episode.srt");
        std::fs::write(&sidecar, b"1\n").expect("a stand-in sidecar is writable");

        // The sidecar is the probe's answer now, handed in rather than looked for — which is
        // what the finding of it is, and it is found and cached once per file version (see
        // `probe_video_geometry`). This is the filter's half of that arrangement.
        let filter = subtitle_filter(Some(&sidecar), None, Some(2))
            .expect("a sidecar is a subtitle file to show");

        assert!(
            filter.contains("episode.srt"),
            "the sidecar is what is drawn, not the file's own track: {filter}"
        );
        assert!(
            !filter.contains("si="),
            "a sidecar carries no streams of its own for a track to name — the whole of that file \
             is the subtitle — so the track the caller asked for has nothing to apply to: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A sidecar whose name cannot be spelled falls through to the copy made out of the film
    /// rather than to the film itself.
    ///
    /// The sidecar wins over the copy beneath it wherever it can be named, so refusing it must not
    /// mean refusing the subtitles too: the copy is the *next source down*, not another attempt at
    /// the same one. And where there is no copy either, the answer is no subtitles at all — the
    /// film's own track is never named (see `the_films_own_track_is_never_named`).
    #[test]
    fn a_sidecar_that_cannot_be_spelled_falls_through_to_the_copy_of_the_films_own_track() {
        let dir = std::env::temp_dir().join(format!("hl-sidecar-fall-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");

        // A sidecar whose own name cannot be spelled: the copy beneath it answers, and the
        // sidecar is not silently renamed into something else.
        let cut = dir.join("episode's cut.srt");
        std::fs::write(&cut, b"1\n").expect("a stand-in sidecar is writable");
        let copy = dir.join("sub1.ass");
        std::fs::write(&copy, b"[Script Info]\n").expect("a stand-in copy is writable");
        let copied = DerivedSubtitles {
            tracks: vec![None, Some(copy)],
            fonts: None,
        };

        let filter = subtitle_filter(Some(&cut), Some(&copied), Some(1))
            .expect("the copy beneath an unspellable sidecar is still a subtitle file to draw");

        assert!(
            !filter.contains(".srt"),
            "the sidecar here is named `episode's cut.srt`, whose apostrophe cannot be spelled, so a \
             filter naming it would open `episodes cut.srt` instead — quietly, and wrongly: {filter}"
        );
        assert!(
            filter.contains("sub1.ass"),
            "and the copy made out of the film is what answers instead: {filter}"
        );

        // With no copy beneath it, the same sidecar is a film played without subtitles rather
        // than a filter naming the film.
        assert_eq!(
            subtitle_filter(Some(&cut), None, None),
            None,
            "an unspellable sidecar with nothing under it is no subtitles rather than the film"
        );

        // A sidecar that can be spelled is still preferred over the copy, which is what makes the
        // fall-through above the second answer rather than the first.
        let plain = dir.join("plain.srt");
        std::fs::write(&plain, b"1\n").expect("a stand-in sidecar is writable");

        let filter = subtitle_filter(Some(&plain), Some(&copied), Some(1))
            .expect("a spellable sidecar is a subtitle file");

        assert!(
            filter.contains("plain.srt"),
            "a sidecar that can be named wins over the copy, as it always has: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// The track the caller names is the copy that is drawn, which is the whole of what a track
    /// choice is worth to a relaunch.
    ///
    /// A player reports nothing about what it did with the number, and every relaunch — a seek, a
    /// resize, a change of track — begins a player again, so this is the only place a choice survives
    /// one: the copy the filter is given is the one the bar's next-track press landed on.
    ///
    /// The index is the subtitle-relative one, counted from zero among subtitle streams, which is
    /// the numbering the copies were written under (see `subtitle_files`).
    #[test]
    fn the_track_the_caller_names_is_the_copy_the_filter_is_given() {
        let dir = std::env::temp_dir().join(format!("hl-track-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let copies = ["sub0.ass", "sub1.ass", "sub2.srt"].map(|name| {
            let file = dir.join(name);
            std::fs::write(&file, b"[Script Info]\n").expect("a stand-in copy is writable");
            Some(file)
        });
        let copied = DerivedSubtitles {
            tracks: copies.to_vec(),
            fonts: None,
        };

        for chosen in 0..3 {
            let filter = subtitle_filter(None, Some(&copied), Some(chosen))
                .expect("a film with copies has a filter to draw them with");

            assert!(
                filter.contains(&format!("sub{chosen}.")),
                "the copy the caller chose is the one that has to be drawn, or the choice never \
                 left the bar: {chosen} chosen, got {filter}"
            );
        }

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A track the answer does not carry is answered with the first copy rather than with nothing.
    ///
    /// A number with no copy behind it is a stale cache entry rather than a user's choice — the
    /// probe's answer is per file and version — so the fallback costs the choice and not the
    /// subtitles, which is the asymmetry worth having: the subtitles are what the user asked for,
    /// and which track among them is the smaller of the two facts.
    #[test]
    fn a_track_the_answer_does_not_carry_is_answered_with_the_first_copy_rather_than_with_nothing()
    {
        let dir = std::env::temp_dir().join(format!("hl-range-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let first = dir.join("sub0.ass");
        std::fs::write(&first, b"[Script Info]\n").expect("a stand-in copy is writable");
        let copied = DerivedSubtitles {
            tracks: vec![Some(first)],
            fonts: None,
        };

        for out_of_range in [1, 2, 7, usize::MAX] {
            let filter = subtitle_filter(None, Some(&copied), Some(out_of_range))
                .expect("a film with a copy keeps its subtitles whatever number the cache held");

            assert!(
                filter.contains("sub0.ass"),
                "`{out_of_range}` names no copy and the first copy answers for it: {filter}"
            );
        }

        // A caller that has chosen nothing is the same answer, because the player's own first
        // stream is what a pin writes down for itself (see `SubtitleStreams::chosen`).
        let filter =
            subtitle_filter(None, Some(&copied), None).expect("a film with a copy has a filter");
        assert!(
            filter.contains("sub0.ass"),
            "nothing chosen draws the first copy: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// **The film's own track is never named, whatever the state of the copy** — the regression
    /// this refusal exists for.
    ///
    /// A `subtitles` filter pointed at the film opens it and streams the whole container before the
    /// first frame is drawn — measured at 14 904 ms and 1 423 MB on a cold 1.4 GB MKV — while the
    /// player's start is waited for under `VIDEO_START_WAIT_SECS`. So naming the film is not a slow
    /// preview but a preview the loop abandons and kills: a spinner that ends in nothing, every
    /// time the film is hovered, and — the answer being cached — permanently. A copy that is
    /// still coming, a copy that has failed and a film nothing can copy are all a film played
    /// without subtitles.
    #[test]
    fn the_films_own_track_is_never_named() {
        for (streams, track) in [(1usize, Some(0)), (2, Some(1)), (3, Some(2))] {
            assert_eq!(
                subtitle_filter(None, None, track),
                None,
                "a film with {streams} subtitle streams of its own and no small file to draw is a \
                 film played without subtitles: naming the film is the whole-container read that \
                 takes the preview down with it"
            );
        }

        // And a caller that chose nothing is the same answer: a film with no small file and no
        // sidecar has no filter at all, however the extraction ended.
        assert_eq!(
            subtitle_filter(None, None, None),
            None,
            "no copy and no sidecar is no filter"
        );
    }

    /// A sidecar named for the language it is in is drawn, which is the naming a release actually
    /// ships and the one this used to walk straight past.
    ///
    /// `episode.en.srt` beside `episode.mkv` is how an anime release writes an external track: the
    /// language is in the name because there is more than one of them, and the file beside it that
    /// carries the picture says nothing about which language the words are in. This looked for
    /// `episode.srt` and stopped there, so a folder written that way was answered with the
    /// container's own track at best and with nothing at all where the container had none — which is
    /// a sidecar sitting right there, spelled correctly, never once looked at.
    ///
    /// The directory is read rather than a list of language codes guessed at, because the codes are
    /// unbounded and a release that writes `episode.en.srt` will as happily write `episode.eng.srt`
    /// or a code this has never heard of. One read of the video's own folder is the whole of it:
    /// measured at 1.9ms over a folder of 1200 files, against a player launch that costs 300ms, and
    /// it is reached only once the plain swap has already found nothing.
    #[test]
    fn a_sidecar_named_for_its_language_is_drawn_rather_than_walked_past() {
        // The switch on: these are the answers the walk has always given,
        // which is the answer under the setting a film's subtitles are
        // wanted for (see `VideoSubtitlesSetting`).
        let _setting = VideoSubtitlesSetting::stood_at(true);
        let dir = std::env::temp_dir().join(format!("hl-lang-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let video = dir.join("episode.mkv");
        std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");
        let english = dir.join("episode.en.srt");
        std::fs::write(&english, b"1\n").expect("a stand-in sidecar is writable");

        let found = sidecar_for(&video);
        assert_eq!(
            found.as_deref(),
            Some(english.as_path()),
            "`episode.en.srt` is the spelling a release ships, and it has to be what is found: \
             {:?}",
            found
        );

        // And it has to reach the filter as that answer does: the probe asks once and the launch
        // is handed it (see `probe_video_geometry`), so the drawing of it is a separate half and
        // is checked here rather than assumed from the finding above.
        let filter = subtitle_filter(found.as_deref(), None, None)
            .expect("a sidecar carrying the language in its own name is still a sidecar");

        assert!(
            filter.contains("episode.en.srt"),
            "a sidecar the probe found has to be the file the filter names: {filter}"
        );
        assert!(
            !filter.contains("si="),
            "and it is still drawn as a whole file rather than as one of the video's own streams: \
             {filter}"
        );

        // The longer code is the same arrangement, and pinning it here is what stops a fix for
        // `en` from quietly being a fix for two letters only.
        let video = dir.join("other.mkv");
        std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");
        let longer = dir.join("other.eng.ass");
        std::fs::write(&longer, b"1\n").expect("a stand-in sidecar is writable");

        let found = sidecar_for(&video);
        let filter = subtitle_filter(found.as_deref(), None, None)
            .expect("a three-letter language code is the same case");

        assert_eq!(
            found.as_deref(),
            Some(longer.as_path()),
            "the code's length is not the question — whether the name carries one at all is: \
             {:?}",
            found
        );
        assert!(
            filter.contains("other.eng.ass"),
            "and it reaches the filter as a sidecar does: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// The plain swap still wins over the named one, because it is the arrangement every tool reads.
    ///
    /// A folder can hold both, and the one every player would pick is `episode.srt` — so a second
    /// rule added underneath must not take the first answer away, or the file that has worked for
    /// years starts depending on which other files happen to be beside it.
    #[test]
    fn the_plain_sidecar_still_wins_over_one_named_for_its_language() {
        // The switch on, as the test beside this one stands it: the answers
        // here are the ones the walk has always given (see
        // `VideoSubtitlesSetting`).
        let _setting = VideoSubtitlesSetting::stood_at(true);
        let dir = std::env::temp_dir().join(format!("hl-lang-priority-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let video = dir.join("episode.mkv");
        std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");
        let plain = dir.join("episode.srt");
        std::fs::write(&plain, b"1\n").expect("a stand-in sidecar is writable");
        std::fs::write(dir.join("episode.en.srt"), b"1\n").expect("a stand-in sidecar is writable");

        let found = sidecar_for(&video);
        let filter = subtitle_filter(found.as_deref(), None, None).expect("a sidecar is a sidecar");

        assert_eq!(
            found.as_deref(),
            Some(plain.as_path()),
            "`episode.srt` is the one every tool on this machine reads, so it is asked for first \
             and the answer must not depend on what else is in the folder: {:?}",
            found
        );
        assert!(
            filter.contains("episode.srt") && !filter.contains("episode.en.srt"),
            "and the filter names the answer it is handed and nothing else: {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A sidecar that is not beside the video is not a sidecar, and a video's own name with an
    /// extension appended is not one either — the two boundaries a broader rule has to keep.
    ///
    /// Both are pins on `sidecar_for` having been widened from three stats into a read of the
    /// folder, which is the change that could plausibly have swallowed either: the folder of a film
    /// is full of other films' subtitles, and `episode.mkv.srt` sits in the very shape the wider
    /// rule matches. No player reads the appended form, so drawing it would be drawing a file the
    /// user did not mean.
    #[test]
    fn a_sidecar_is_only_the_one_beside_the_video_and_never_the_video_s_own_name_appended() {
        // The switch on, so the walk these boundaries pin is the one that runs
        // rather than the gate answering the question before it is asked (see
        // `VideoSubtitlesSetting`).
        let _setting = VideoSubtitlesSetting::stood_at(true);
        let dir = std::env::temp_dir().join(format!("hl-lang-bounds-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let elsewhere = dir.join("other-folder");
        std::fs::create_dir_all(&elsewhere).expect("a scratch folder is creatable");

        // `episode.mkv.srt`: the appended spelling, in the very shape the wider rule matches.
        let video = dir.join("episode.mkv");
        std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");
        std::fs::write(dir.join("episode.mkv.srt"), b"1\n")
            .expect("a stand-in sidecar is writable");

        let found = sidecar_for(&video);
        assert_eq!(
            found, None,
            "`episode.mkv.srt` is not read by any player, so drawing it would be drawing a file \
             the user did not mean — and the video's own extension after the dot is what tells the \
             two apart"
        );
        assert_eq!(
            subtitle_filter(found.as_deref(), None, None),
            None,
            "and with nothing found and no streams of its own, the film is drawn without subtitles"
        );

        // A sidecar in another folder, named for this video exactly. No walk: a hover must not
        // cost a tree, and a sidecar beside some *other* film is not this film's subtitles.
        let video = dir.join("lonely.mkv");
        std::fs::write(&video, b"stand-in").expect("a stand-in file is writable");
        std::fs::write(elsewhere.join("lonely.en.srt"), b"1\n").expect("a sidecar is writable");

        assert_eq!(
            sidecar_for(&video),
            None,
            "a subtitle file in a different folder is not this film's subtitles, and looking for one \
             would put a directory walk on the thread that has to keep answering Explorer"
        );
        assert_eq!(
            subtitle_filter(None, None, None),
            None,
            "so the filter below it is built from nothing, which is the only answer that leaves the \
             film playing"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A video with no subtitles of its own and none beside it is played without a filter, which is
    /// the whole of what "no subtitles" asks for.
    #[test]
    fn a_video_with_no_subtitles_anywhere_is_given_no_subtitle_filter() {
        // The switch on, so the walk runs and answers nothing, rather than the
        // gate answering the question before the walk is asked (see
        // `VideoSubtitlesSetting`).
        let _setting = VideoSubtitlesSetting::stood_at(true);
        let dir = std::env::temp_dir().join(format!("hl-none-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let video = dir.join("silent.mkv");
        std::fs::write(&video, b"not really a video").expect("a stand-in file is writable");

        assert_eq!(
            subtitle_filter(sidecar_for(&video).as_deref(), None, Some(0)),
            None,
            "a filter naming a file that is not there is a filter that fails to load"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A film whose own tracks are still being copied is drawn without subtitles for that hover,
    /// which is the whole of the fault this arrangement was built for.
    ///
    /// The user chose this against the alternative of drawing the slow ones: the first hover of
    /// a film whose subtitles are embedded is fast and has none, and every hover after the one
    /// copying pass finishes draws the small files it wrote (see `subtitle_files`). Before this,
    /// the filter named the film itself, and the player streamed the whole container to the
    /// track's first packet before drawing a frame — measured cold, 14 904 ms and 1 423 MB read
    /// on a 1.4 GB episode.
    #[test]
    fn a_film_whose_own_tracks_are_still_being_copied_is_drawn_without_subtitles() {
        assert_eq!(
            subtitle_filter(None, None, Some(0)),
            None,
            "naming the film is the read this exists to take off the hover, so a copy that is \
             still coming means no filter this time rather than the slow one"
        );
    }

    /// Once the one copying pass has failed, the film is still drawn without subtitles rather than
    /// from its own track: a slow hover is not what a failure buys, a broken preview is.
    ///
    /// This is the other half of the regression `the_films_own_track_is_never_named` is written
    /// for. The failure is remembered in the probe's answer — the film is not asked for a copy
    /// again until the next run — and what that remembered failure must not become is the
    /// whole-container read: that route overruns the ten-second wait a player's start is given,
    /// so the preview is killed and every later hover repeats it. No subtitles and a fast film is
    /// the answer, and the film never stops being previewable for it.
    #[test]
    fn a_film_whose_copy_has_failed_is_still_drawn_without_subtitles() {
        assert_eq!(
            subtitle_filter(None, None, Some(1)),
            None,
            "the film's own track would stream the whole container before a first frame — measured \
             at 14 904 ms — so a failed copy is a film played without subtitles, and the preview is \
             never the thing that is lost"
        );
    }

    /// A track this app copied out of the film is drawn like a sidecar — a whole file, no track
    /// number — with the container's own fonts beside it, which is what keeps an embedded
    /// track's styling exactly as it was while the film itself is never opened.
    #[test]
    fn a_copied_track_is_drawn_with_its_fonts_and_no_track_number() {
        let dir = std::env::temp_dir().join(format!("hl-copied-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
        let ass = dir.join("sub0.ass");
        std::fs::write(&ass, b"[Script Info]\n").expect("a stand-in copy is writable");
        let fonts = dir.join("fonts");
        std::fs::create_dir_all(&fonts).expect("the fonts folder for this test is creatable");
        std::fs::write(fonts.join("font0.ttf"), b"stand-in").expect("a stand-in font is writable");

        let copied = DerivedSubtitles {
            tracks: vec![Some(ass.clone())],
            fonts: Some(fonts),
        };
        let filter = subtitle_filter(None, Some(&copied), None)
            .expect("a copied track is a subtitle file to draw");

        assert!(
            filter.contains("sub0.ass"),
            "the copied file is what is opened, not the film: {filter}"
        );
        assert!(
            filter.contains("fontsdir="),
            "the container's own fonts are named beside it, or the styles are drawn with the \
             faces the machine happens to have: {filter}"
        );
        assert!(
            !filter.contains("si="),
            "a copied track is a whole file with no streams of its own, exactly as a sidecar is: \
             {filter}"
        );

        // A copy whose container attached no fonts is named alone.
        let copied = DerivedSubtitles {
            tracks: vec![Some(ass)],
            fonts: None,
        };
        let filter = subtitle_filter(None, Some(&copied), None)
            .expect("a copied track is a subtitle file to draw with or without its fonts");

        assert!(
            !filter.contains("fontsdir="),
            "a `fontsdir` naming a folder that holds nothing is an argument that says nothing: \
             {filter}"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A sidecar renamed away after the probe answered is not named again — the cached answer is
    /// stale, and a filter naming a file that is not there fails to load, which takes the
    /// preview down with it. This is the fault a renamed sidecar used to brick a preview with.
    #[test]
    fn a_sidecar_that_went_away_is_not_named() {
        let dir = std::env::temp_dir().join(format!("hl-stale-{}", std::process::id()));
        std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");

        // The sidecar the probe saw is gone; a copy of the film's own track answers instead.
        let missing = dir.join("episode.srt");
        let ass = dir.join("sub0.ass");
        std::fs::write(&ass, b"[Script Info]\n").expect("a stand-in copy is writable");
        let copied = DerivedSubtitles {
            tracks: vec![Some(ass)],
            fonts: None,
        };
        let filter = subtitle_filter(Some(&missing), Some(&copied), Some(0))
            .expect("the copied track is still a subtitle file to draw");

        assert!(
            filter.contains("sub0.ass"),
            "the copied track answers where the sidecar used to: {filter}"
        );
        assert!(
            !filter.contains("episode.srt"),
            "the stale path must not be named: a filter naming a file that is not there fails to \
             load and the preview never appears: {filter}"
        );

        // Gone with nothing else to draw: no filter at all rather than the missing file, and
        // never the film's own track (see `the_films_own_track_is_never_named`).
        assert_eq!(
            subtitle_filter(Some(&missing), None, Some(0)),
            None,
            "a sidecar that is not there is not a sidecar, and nothing else is drawn while the \
             film's own copy is coming"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// The whole command line an mpv video is begun with, asserted in one piece: an edit to any
    /// flag in it is then a failing test rather than a surprise in a preview nobody on this
    /// machine can run (there is no mpv here — see the report on this ticket).
    ///
    /// **And no `--focus-on*` is among them**: the switch a hover is asked for the foreground with is
    /// appended by the hover's launch alone, so that a pin's command line is the one every road had
    /// (see `hover_focus_switch` and `video_playback::start_mpv_video_playback`).
    #[test]
    fn a_hover_asks_for_the_focus_switch_the_build_parses_and_nothing_where_it_parses_neither() {
        assert_eq!(
            hover_focus_switch(FocusOnSupport::Never),
            Some("--focus-on=never"),
            "the current spelling, where the installed build parses it"
        );
        assert_eq!(
            hover_focus_switch(FocusOnSupport::OpenOnly),
            Some("--focus-on-open=no"),
            "and the older one where that is the one it knows"
        );
        assert_eq!(
            hover_focus_switch(FocusOnSupport::Neither),
            None,
            "a build that parses neither is asked for nothing, rather than for a flag it would \
             refuse and refuse the whole launch with"
        );
    }

    #[test]
    fn the_mpv_route_is_launched_with_the_command_line_the_spec_settles_on() {
        assert_eq!(
            mpv_args(
                Path::new("films/episode.mkv"),
                10,
                20,
                640,
                360,
                0.0,
                false,
                80,
                None,
                true,
                true,
                Some(Path::new("scratch/mpv-input.conf")),
                IPC,
            false,
            true
            ),
            [
                "films/episode.mkv",
                "--no-config",
                "--no-input-default-bindings",
                "--input-conf=scratch/mpv-input.conf",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=640x360+10+20",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=80",
                IPC_ARG,
            ],
            "the film, then the configuration off, the mouse off, the chrome off, the order it is\
             held up in, the box, the video output being measured, the three Direct3D options \
             and the level — the output is named at all only as a measurement, and no \
             `--start` is asked for at the beginning of a file (see `mpv_args`)"
        );
    }

    /// **The flip-model swapchain is the thumbnail road's, and the knob is the only thing that asks
    /// for it.** A pinned mpv whose window is a window of the desktop's is the one place a flip-model
    /// swapchain fixes what the bitblt model gets wrong: the compositor scales the buffer to the
    /// window's client area, so the picture fills the window by construction rather than by mpv
    /// catching a `WM_SIZE` (see `mpv_args` and `pin_film_thumbnail`).
    #[test]
    fn the_flip_model_swapchain_is_asked_for_only_on_the_thumbnail_road() {
        let with_flip = mpv_args(
            Path::new("a.mkv"),
            0,
            0,
            640,
            360,
            0.0,
            false,
            80,
            None,
            true,
            true,
            None,
            IPC,
            true,
            false);
        assert!(
            with_flip.iter().any(|arg| arg == "--d3d11-flip=yes"),
            "a player that makes a window of its own on the thumbnail road is asked for the flip \
             model when the knob is on: {with_flip:?}"
        );

    }

    /// The flip knob is exactly `1`: the A/B is two runs, and a value that is not the one named must
    /// be the default road rather than a third one nobody wrote down (see
    /// `RHP_THUMBNAIL_FLIP`).
    #[test]
    fn the_flip_knob_is_on_for_one_and_off_for_everything_else() {
        assert_eq!(THUMBNAIL_FLIP_VARIABLE, "RHP_THUMBNAIL_FLIP");
        assert!(!thumbnail_flip_from(None));
        assert!(thumbnail_flip_from(Some(std::ffi::OsStr::new("1"))));
        assert!(!thumbnail_flip_from(Some(std::ffi::OsStr::new("0"))));
        assert!(!thumbnail_flip_from(Some(std::ffi::OsStr::new("yes"))));
        assert!(!thumbnail_flip_from(Some(std::ffi::OsStr::new(""))));
    }

    /// **A flip-model swapchain is not a thing a layered window can have**, so the invisible road asks
    /// for the bitblt one whatever the knob says — and the knob is overruled out loud, in the launch's
    /// own trace, rather than dropped on the floor (see `the_thumbnail_swapchain_flip`). A knob nobody
    /// can read back is a run nobody can compare, and this is the one case where the run's flag is not
    /// the one asked for.
    #[test]
    fn a_layered_source_is_never_asked_for_the_flip_model_whatever_the_knob_says() {
        assert!(!the_thumbnail_swapchain_flip(false, true));
        assert!(!the_thumbnail_swapchain_flip(false, false));
        assert!(
            the_thumbnail_swapchain_flip(true, false),
            "a visible source with the knob on is the run the knob is for"
        );
        assert!(
            !the_thumbnail_swapchain_flip(true, true),
            "and an invisible one cannot have the model at all"
        );
    }

    /// **The one flag the drag story rests on without anything else**: mpv is launched told not to
    /// move its own window.
    ///
    /// `--window-dragging` defaults to yes, and a player that moves its own window while this app
    /// carries the pin is two things writing one position — which is the film being pulled out of
    /// the hand and put back every tick. Whatever keeps a press off mpv's window (the pin's own
    /// band, its chrome, the click-through style) is a fact about where that window happens to be,
    /// and the refusal here is not.
    ///
    /// It is asserted apart from the whole-list test because it has to hold with the input flags
    /// *dropped* as well: those two go together where the generated binding file cannot be written,
    /// and this one is not part of that pair.
    #[test]
    fn an_mpv_player_is_launched_told_not_to_drag_its_own_window() {
        let args = mpv_args(
            Path::new("a.mkv"),
            0,
            0,
            8,
            8,
            0.0,
            false,
            100,
            None,
            true,
            true,
            None,
            IPC,
        false,
        true
        );
        assert!(
            args.iter().any(|arg| arg == "--window-dragging=no"),
            "mpv drags its own window with the left button by default, so a launch that leaves it \
             at mpv's default leaves a second writer of the position this app is carrying: the \
             film would be moved by the hand and put back by the app every tick, which is the \
             magnet the left and right borders were reported as; the launch was {args:?}"
        );
    }

    /// Subtitles are the tray's `Subtitles` row, and with it off the film is asked for none at all:
    /// the row is answered by mpv as it is by the ffplay route, rather than left to a player that
    /// would find and select the film's own track whatever the row says.
    #[test]
    fn an_mpv_video_is_shown_without_subtitles_where_the_row_is_off() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                0.0,
                false,
                100,
                None,
                true,
                false,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sid=no",
                "--sub-auto=no",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=100",
                IPC_ARG,
            ],
            "with the row off nothing is shown and nothing is looked for: `--sid=no` selects no \
             track at all, embedded or beside the film, and `--sub-auto=no` keeps the sidecar \
             beside the film from being read as one in the first place"
        );
    }

    /// The switch under `Performance → Hardware Acceleration` is the whole of what mpv is told
    /// about decoding, and its off position asks for a core.
    #[test]
    fn an_mpv_video_is_decoded_on_the_card_only_where_the_switch_says_so() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                0.0,
                false,
                50,
                None,
                false,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=no",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=50",
                IPC_ARG,
            ],
            "the switch off is the whole of the answer: mpv is asked for no device at all rather \
             than for one this app measured, and every other flag is the one the default launch \
             carries"
        );
    }

    /// A level of zero is a film with no audio rather than a film played at nothing, which is the
    /// ffplay route's `-an` — and a film with nothing to hear is not normalised either.
    #[test]
    fn a_muted_mpv_video_is_given_no_audio_and_no_filter() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                0.0,
                false,
                0,
                Some(2.5),
                true,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--aid=no",
                IPC_ARG,
            ],
            "a muted film is a film with its audio turned off, which saves the decoder the whole \
             soundtrack, and a gain for a soundtrack nothing will play is a filter nothing asked \
             for"
        );
    }

    /// A measured gain with `Normalize` on reaches mpv as the same libavfilter `volume` filter the
    /// ffplay route applies, through mpv's own bridge to libavfilter.
    ///
    /// **The bridge is not a preference, it is the spelling the manual gives.** The gain is
    /// libavfilter's `volume` filter, and mpv reaches a libavfilter filter through its `lavfi`
    /// bridge: `lavfi=[volume=…]` is what the manual's `--af` syntax asks for and what mpv's own
    /// `--af=help` example prints for the audio bridge, where a bare `--af=volume=` reaches the
    /// same filter only by the automatic bridge mpv falls back on for a name it has no filter of
    /// its own for — one the manual says neither verifies arguments nor supports help output, so
    /// the graph is named rather than left to the fallback (checked against mpv's manual and
    /// `filters/f_lavfi.c` — see the report on this ticket).
    #[test]
    fn a_measured_gain_reaches_mpv_as_the_libavfilter_volume_filter() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                0.0,
                false,
                100,
                Some(2.5),
                true,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=100",
                "--af=lavfi=[volume=2.5000]",
                IPC_ARG,
            ],
            "the gain is the filter ffplay is given the same gain through, and `--volume=100` \
             stands beside it as the level the two multiply with"
        );
    }

    /// A film begun at a second is asked for that second, in the spelling the manual gives.
    #[test]
    fn an_mpv_video_begun_at_a_second_is_asked_for_that_second() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                12.5,
                false,
                100,
                None,
                true,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--start=12.500",
                "--volume=100",
                IPC_ARG,
            ],
            "the second a pinned bar was dragged to is what the relaunch asks mpv for, at the same \
             precision the ffplay route writes it in"
        );
    }

    /// **A player begun on a frame another player is showing is begun standing on it**: the seek is
    /// exact rather than to the keyframe before the second, and the player is paused so that the frame
    /// it presents is that one and stays that one until the handover lets it go (see
    /// `pin_film_host::settle_pin_handover`). A player begun any other way is given neither flag, and
    /// an ordinary launch is unchanged by this.
    #[test]
    fn a_player_begun_on_a_frozen_frame_is_held_there_and_seeks_to_it_exactly() {
        let held = mpv_args(
            Path::new("a.mkv"),
            0,
            0,
            8,
            8,
            12.5,
            true,
            100,
            None,
            true,
            true,
            None,
            IPC,
            false,
            true
        );
        assert!(
            held.contains(&"--start=12.500".to_string()),
            "the frame to stand on is named: {held:?}"
        );
        assert!(
            held.contains(&"--hr-seek=yes".to_string()),
            "and the seek is exact rather than to the keyframe before it, or the pin opens on a \
             frame the hover was never showing: {held:?}"
        );
        assert!(
            held.contains(&"--pause".to_string()),
            "and the player stands there rather than playing off the frame the hover froze on \
             before the handover can ask for it: {held:?}"
        );

        let playing = mpv_args(
            Path::new("a.mkv"),
            0,
            0,
            8,
            8,
            12.5,
            false,
            100,
            None,
            true,
            true,
            None,
            IPC,
            false,
            true
        );
        assert!(
            !playing
                .iter()
                .any(|arg| arg == "--hr-seek=yes" || arg == "--pause"),
            "and a player begun to play from a second is given neither: {playing:?}"
        );
    }

    /// A preview box on a monitor left of or above the primary one has a negative coordinate, and
    /// that coordinate is passed on as `+-100`: the geometry's offsets are introduced by a sign, so
    /// a negative one is `+-100` rather than `-100` — which mpv reads as a distance from the right
    /// or bottom edge of the display instead of a coordinate on the desktop, and which on a machine
    /// whose second display sits above the primary puts the film on the primary display every time.
    ///
    /// Measured against mpv 0.41.0 on a primary display 3440x1440 with a second one at
    /// (3440,-2160), for the preview box (3540,-2060): `--geometry=400x300+3540+-2060` puts the
    /// window on the second display, and `--geometry=400x300` centres it on the primary at
    /// (1520,570) — which is the fault this is the regression test for, since the box the app
    /// asked for was on the second display.
    #[test]
    fn an_mpv_video_on_a_monitor_left_of_the_primary_one_is_asked_for_its_own_box() {
        for (x, y, which) in [(-100, 20, "left of the primary"), (20, -100, "above it")] {
            let geometry = format!("--geometry=640x360+{x}+{y}");
            assert_eq!(
                mpv_args(
                    Path::new("a.mkv"),
                    x,
                    y,
                    640,
                    360,
                    0.0,
                    false,
                    100,
                    None,
                    true,
                    true,
                    None,
                    IPC,
                false,
                true
                ),
                [
                    "a.mkv",
                    "--no-config",
                    "--no-osc",
                    "--no-border",
                    "--window-dragging=no",
                    "--ontop",
                    &geometry,
                    "--vo=gpu-next",
                    "--gpu-api=d3d11",
                    "--d3d11-flip=no",
                    "--hwdec=auto",
                    "--sub-auto=fuzzy",
                    "--keep-open=yes",
                    "--keep-open-pause=no",
                    "--volume=100",
                    IPC_ARG,
                ],
                "a box {which} is asked for as the box it is, negative coordinate and all"
            );
        }

        // And a position both of whose coordinates are non-negative is spelled the same way, so
        // there is one spelling of a box rather than two: `+0+0` is the primary display's origin.
        assert!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                640,
                360,
                0.0,
                false,
                100,
                None,
                true,
                true,
                None,
                IPC,
            false,
            true
            )
            .contains(&"--geometry=640x360+0+0".to_string()),
            "a box at the primary display's own origin is passed as it stands"
        );
    }

    /// **The size a film's box is asked for on a display whose scale is not the primary's**, which
    /// is the regression test for the window that came up at twice the box and was put right a
    /// second later by the tick — the fault a hover on the second display shows and a hover on the
    /// primary one cannot.
    ///
    /// The measurement the division rests on is written out over `mpv_box_size`. What is pinned
    /// here is the answer this machine's second display gives it together with the command line
    /// that answer reaches mpv as, because a size asked for wrongly is a command line right up to
    /// the preview — and there is no preview on a machine of one scale that can fail.
    ///
    /// **The trace the answer rests on is not and cannot be a test**: it reads `GetWindowRect` off
    /// a player's window on a second display, so a suite holding it would need that display, a
    /// player window on it and a pointer there, and what it measured would be this machine's two
    /// displays rather than the code. The window itself was read, and the readings are in the note
    /// over `mpv_box_size` and in the report on this ticket.
    #[test]
    fn a_box_on_a_scaled_display_is_asked_for_on_the_scale_of_the_display_the_window_is_born_on() {
        // The fit a hover of a 3840x2160 room gives beside the pointer on a display drawn at 200%,
        // beside a 96 DPI primary: the box is what the layout placed, and mpv is told that box on
        // the primary's scale, which is the size Windows doubles on its way to the second display.
        let (asked_w, asked_h) = mpv_box_size(2864, 1611, 192, 96);
        assert_eq!(
            (asked_w, asked_h),
            (1432, 806),
            "the window comes out at the size asked for times the display's scale over the \
             primary's, so the box is asked for at that share of itself"
        );
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                4416,
                -2160,
                asked_w,
                asked_h,
                0.0,
                false,
                100,
                None,
                true,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=1432x806+4416+-2160",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=100",
                IPC_ARG,
            ],
            "the box the layout placed is 2864x1611 and the box mpv is given for it is that size \
             on the primary's scale; the position is the box's own, because Windows keeps a \
             window's top-left corner where a DPI change finds it (measured)"
        );

        // Every box on a machine of one scale, and every box on a display drawn at the primary's
        // scale, is asked for exactly as it stands — which is the whole of the common case, and it
        // cannot be made worse by this: two displays at 100% are two displays at one scale.
        assert_eq!(mpv_box_size(1720, 967, 96, 96), (1720, 967));
        assert_eq!(mpv_box_size(480, 270, 144, 144), (480, 270));

        // A display scaled *below* the primary's is the same division the other way, because the
        // window would come out smaller than the box there rather than larger.
        assert_eq!(mpv_box_size(960, 540, 96, 192), (1920, 1080));

        // And a box that does not halve evenly is a pixel out rather than a window twice the size:
        // 967 is the fit this machine's second display gives a 3840x2160 room and it is odd, which
        // is the trade `mpv_box_size`'s note names — the tick asserts the box itself a moment
        // later (see `ensure_video_window_topmost`).
        assert_eq!(mpv_box_size(1720, 967, 192, 96), (860, 484));
        assert_eq!(mpv_box_size(480, 270, 192, 96), (240, 135));
    }

    /// The two input flags are dropped together where the binding file could not be written, so a
    /// player is never given half of the arrangement.
    ///
    /// `--input-conf` alone does not remove mpv's baked-in mouse bindings, and
    /// `--no-input-default-bindings` alone takes `p` away with them — so a launch holding one of
    /// the two is either a hover with a live mouse or a film the pin cannot pause. Both flags or
    /// neither is the whole of the answer.
    #[test]
    fn a_binding_file_that_could_not_be_written_drops_both_input_flags() {
        assert_eq!(
            mpv_args(
                Path::new("a.mkv"),
                0,
                0,
                8,
                8,
                0.0,
                false,
                100,
                None,
                true,
                true,
                None,
                IPC,
            false,
            true
            ),
            [
                "a.mkv",
                "--no-config",
                "--no-osc",
                "--no-border",
                "--window-dragging=no",
                "--ontop",
                "--geometry=8x8+0+0",
                "--vo=gpu-next",
                "--gpu-api=d3d11",
                "--d3d11-flip=no",
                "--hwdec=auto",
                "--sub-auto=fuzzy",
                "--keep-open=yes",
                "--keep-open-pause=no",
                "--volume=100",
                IPC_ARG,
            ],
            "no file is no `--input-conf`, and mpv's own bindings stand rather than its bindings \
             minus every key — the two flags are dropped together or not at all"
        );
    }

    /// The binding file holds the one line that pauses the film, written where the app puts what
    /// it derives — and a write that cannot be made answers so rather than pretending.
    #[test]
    fn the_binding_file_holds_the_one_binding_and_a_write_that_fails_says_so() {
        let dir = std::env::temp_dir().join(format!("hl-mpv-conf-{}", std::process::id()));
        let file = dir.join("nested").join("mpv-input.conf");

        assert!(
            !write_input_conf(&file),
            "a file in a folder that is not there is a write that failed, and what that costs is \
             the pair of input flags rather than a player that will not start"
        );

        std::fs::create_dir_all(file.parent().expect("the folder the file is written into"))
            .expect("the scratch folder for this test is creatable");

        assert!(
            write_input_conf(&file),
            "and a write that lands says so, which is what the two flags are asked for"
        );
        assert_eq!(
            std::fs::read_to_string(&file).expect("the binding file is readable"),
            MPV_PAUSE_BINDING,
            "the file's whole content is the one binding"
        );
        assert_eq!(
            MPV_PAUSE_BINDING, "p cycle pause\n",
            "the key is the one the pin posts and mpv's own binding is the one that answers it \
             (see `VIDEO_PAUSE_KEY`)"
        );

        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A loop is begun again at zero whatever the length of the file, which is the whole of why the
    /// rewind is a relaunch rather than a key.
    ///
    /// FFmpeg's player has no key that goes to the beginning: every seek key it binds is a *fixed
    /// increment*, so `Down` is sixty seconds back and `Left` is ten, and the mistake is invisible on
    /// a clip shorter than a step because a step past the start clamps to the start. On a film of
    /// ten minutes the same key is a seek to nine minutes, then to eight — the film walks backwards a
    /// minute a pass and never plays its first nine minutes again, while this app's own clock, having
    /// been zeroed, says it is at the beginning throughout. A bar nine minutes out with the picture
    /// is worse than no bar.
    ///
    /// So the answer is a second rather than a key, and it is the beginning for every file. That
    /// `rewind_due` says a film of any length is due still says nothing about *where* it goes, which
    /// is why this is the other half of the decision and the half that has to be pinned down too.
    #[test]
    fn a_loop_is_begun_again_at_the_beginning_whatever_the_length_of_the_file() {
        assert_eq!(
            rewind_launch_seconds(),
            0.0,
            "a loop that restarted anywhere but the beginning is a loop over the part of the film \
             nearest its end, and on a film longer than a seek step that is a walk backwards rather \
             than a repeat"
        );

        // The step that made the mistake invisible on a short clip, asserted here so the reason the
        // key was rejected is on the record rather than only in the note above.
        for duration in [8.0, 61.0, 600.0, 7_200.0] {
            let position = duration - REWIND_MARGIN_SECS / 2.0;

            assert!(
                rewind_due(true, Some(duration), position),
                "a film of {duration}s within the margin of its end is due a rewind"
            );
            assert_eq!(
                rewind_launch_seconds(),
                0.0,
                "and it goes to the beginning at {duration}s as it does at 8s: a key would move it \
                 sixty seconds back at 600s and ten at 61s, so the same code would be right on the \
                 short clip and wrong on the long one"
            );
        }
    }

    /// A player begun at a second of the file is given its loop by this app, and one begun at the
    /// beginning is left to loop itself — which is the whole of the arrangement, and the difference
    /// between a film that repeats and a film that repeats its last two seconds.
    #[test]
    fn only_a_player_was_seeked_is_the_loop_this_apps_to_give() {
        assert!(
            !loop_is_ours(false),
            "a player begun at the beginning wraps to the beginning on its own, and never exits out from under a hover"
        );
        assert!(
            loop_is_ours(true),
            "a player begun at a second of the file wraps to that second on its own, so the loop has to be taken away from it"
        );
    }

    /// The rewind is posted before the end rather than at it, and never for a file nothing has
    /// measured — the three answers that decide whether a player is asked for one.
    #[test]
    fn the_rewind_is_posted_short_of_the_end_and_only_for_a_file_that_has_one() {
        assert!(
            rewind_due(true, Some(8.0), 7.95),
            "a player within the margin of a measured end is sent back before it reaches it"
        );
        assert!(
            !rewind_due(true, Some(8.0), 4.0),
            "a player in the middle of the file is left playing it"
        );
        assert!(
            !rewind_due(false, Some(8.0), 7.95),
            "a held film is not rewound, because a hold that jumped to the beginning would be a hold that did nothing"
        );
        assert!(
            !rewind_due(true, None, 7.95),
            "a file whose length nothing has read has no end to be near, and a player at zero of it has nothing to be saved from"
        );
        assert!(
            !rewind_due(true, Some(0.05), 0.04),
            "a file shorter than the margin is every second inside it, so rewinding it would send the player back the instant it began"
        );
    }
}
