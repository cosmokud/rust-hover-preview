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
//! - Hardware acceleration. `-hwaccel` names a *device*, not a decoder, and FFmpeg falls back to
//!   its own software decoding by itself when the device cannot take a stream — so the setting is
//!   one name rather than a capability probe (see `hw_accel_device`).
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
//! Subtitles are a per-frame cost, not a per-start one. The thing that *is* on the way to a first
//! window is `-hwaccel`, measured at +330ms on its own — and, on a machine whose probe finds a
//! device that survives, fatal beside this filter: `-hwaccel dxva2` hands the graph `dxva2_vld`
//! frames, `subtitles` is software-only, and the graph fails to build with `Impossible to convert
//! between the formats supported by the filter 'Parsed_setsar_0' and the filter 'auto_scale_0'`.
//! The player then exits without ever putting a window up, which is a spinner to `VIDEO_START_WAIT_SECS`
//! and no picture at all — and which `-loglevel quiet` over a null stderr says nothing about (see
//! `HWACCEL_DEVICES`, and the note on `video_playback::start_video_playback`).
//!
//! **That silence is the open item — propose-only, pending and not accepted.** The failure is
//! unobservable where it happens: the launch is `-loglevel quiet` over a null stderr (the launch
//! at `video_playback::start_video_playback`, a file another branch owns), so a dxva2 machine
//! reads the fault as a spinner that times out and a frame that never comes, with nothing
//! written to say why. The proposal, for whoever owns that launch: let the graph failure be
//! heard — the player's stderr read through the window wait, or a device refused before the
//! film is launched that cannot share a graph with `subtitles` at all — which is why it is
//! written down here rather than fixed: the launch itself is that file's to make.
//!
//! - The loop's rewind. FFmpeg's player has no key that goes to the beginning of a file, so a loop
//!   this app gives cannot be given by posting one (see `rewind_launch`).

use super::media_types::DerivedSubtitles;
use std::path::Path;

/// The devices worth asking a video to decode on, in the order they are worth asking.
///
/// Ordered by what was measured on this machine rather than by what is fashionable, and every one
/// of them was measured *on FFmpeg's player* rather than on `ffmpeg`, which is the distinction that
/// matters and the one that is easy to get wrong: `ffmpeg` will happily decode through
/// `-hwaccel d3d11va` and say `src: d3d11`, while `ffplay` cannot. `ffplay` takes `-hwaccel` only
/// alongside its Vulkan renderer, which it switches on by itself and quietly when the option is
/// present — so the flag meant to save a core is also the flag that changes how every frame reaches
/// the screen, and on a machine whose Vulkan cannot do what this build asks of it that is an
/// access violation rather than a slower film.
///
/// Measured on FFmpeg 9.0.2 here, playing a 2560x1440 HEVC file, with everything else the app
/// passes:
/// - no `-hwaccel` — plays the file through, software decoded, and is the answer that works.
/// - `-hwaccel d3d11va` — dies of an access violation after four frames.
/// - `-hwaccel auto` — draws no frames at all, having silently brought up the renderer for nothing.
/// - `-hwaccel cuda` — draws no frames at all; there is no NVIDIA device to derive.
/// - `-hwaccel dxva2` — the only name that reaches the picture, and not reliably: it survives on a
///   plain launch and dies of the same access violation on the same file with the seek written after
///   the input, which is the shape this app uses.
///
/// `dxva2` is first because it is both the one that worked and the one present on every Windows
/// machine with a GPU driver at all; the rest are there for machines where it is not enough, and
/// cost nothing on a machine that never reaches them. A name FFmpeg does not know is refused before
/// it decodes a frame, so an unknown one in this list costs a probe and nothing else.
pub const HWACCEL_DEVICES: [&str; 4] = ["dxva2", "d3d11va", "qsv", "cuda"];

/// The devices a video is to be asked for, given what the setting says — and none at all for a
/// setting that is off.
///
/// An empty list is a whole answer rather than a flag on the argument: `-hwaccel` is either there or
/// it is not, and the walk that turns this list into one name is what decides whether anything is
/// named (see `preview_window::probe_hwaccel_device`).
pub fn hw_accel_candidates(enabled: bool) -> &'static [&'static str] {
    if enabled {
        &HWACCEL_DEVICES
    } else {
        &[]
    }
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
    use super::*;
    use std::path::PathBuf;

    /// A setting that is on asks for the devices worth trying, and one that is off asks for none.
    #[test]
    fn hardware_acceleration_asks_for_the_devices_worth_trying_and_only_when_it_is_on() {
        assert_eq!(
            hw_accel_candidates(true),
            &HWACCEL_DEVICES,
            "a setting that is on has to have somewhere to start: the device is measured rather \
             than assumed, so this is a list to walk rather than one name to pass"
        );
        assert!(
            hw_accel_candidates(false).is_empty(),
            "a setting that is off must ask for no device at all, because `-hwaccel` is either \
             there or it is not"
        );
    }

    /// The list is ordered by what was measured to work, and `auto` is not on it.
    ///
    /// `auto` is the trap the ordering exists to avoid: measured, it silently brings up FFmpeg's
    /// Vulkan renderer and draws no frames at all, having decoded nothing. A name that is *fatal*
    /// is worse and is much harder to see, which is why the list is walked against a probe rather
    /// than taken on faith — `d3d11va` dies of an access violation after four frames on the
    /// machine this was written on.
    #[test]
    fn the_devices_are_ordered_by_what_worked_and_never_include_auto() {
        assert_eq!(
            HWACCEL_DEVICES[0], "dxva2",
            "the name that reached a picture goes first, and it is also the one every Windows \
             machine with a GPU driver has"
        );
        assert!(
            !HWACCEL_DEVICES.contains(&"auto"),
            "`-hwaccel auto` silently brings up the Vulkan renderer and draws nothing: the trap \
             this ordering exists to avoid"
        );
        assert!(
            HWACCEL_DEVICES.contains(&"d3d11va"),
            "d3d11va stays in the list for the machines where it works, behind the names that are \
             cheaper to be wrong about"
        );
    }

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
