//! A sound: the track read out of the file, the loudness measured off it, the gain it is
//! normalised to, the player it is given, and the clock it is asked against.

use super::dimensions::audio_font_scale_percent;
use super::*;

/// How loud a video is played, which is read when one is started rather than when the
/// setting changes: a preview is a few seconds long, and the next one is played at
/// whatever the volume is by then.
pub(super) fn current_video_volume() -> u32 {
    CONFIG.lock().map(|cfg| cfg.video_volume).unwrap_or(0)
}

/// The volume a sound is previewed at, read the way the video's is: from the configuration at
/// the moment a player is started, so a change in the tray reaches the next hover.
pub(super) fn current_audio_volume() -> u32 {
    CONFIG.lock().map(|cfg| cfg.audio_volume).unwrap_or(0)
}

/// Where a sound starts, read the way the volume is and at the same moment: from the
/// configuration as a player is started, so a change in the tray reaches the next hover — and
/// a sound already playing is left where it is rather than dropped somewhere else, which is
/// what a seek asked of a running engine would be.
pub(super) fn current_audio_seek() -> AudioSeek {
    CONFIG
        .lock()
        .map(|cfg| cfg.audio_seek)
        .unwrap_or(DEFAULT_AUDIO_SEEK)
}

/// Where a pinned sound starts, read the way the hover's is and at
/// the same moment: from the pin's own setting as a player is
/// started, so a change in the tray reaches the next pin — and a
/// sound already playing is left where it is.
pub(super) fn current_pin_mode_audio_seek() -> AudioSeek {
    CONFIG
        .lock()
        .map(|cfg| cfg.pin_mode_audio_seek)
        .unwrap_or(DEFAULT_PIN_MODE_AUDIO_SEEK)
}

/// What a sound's card is built with: the theme, which is the one setting a painted
/// preview answers to, and the font size the Audio Scaling share names — the default
/// text size at the 10% anchor, scaled by the share's fraction of it (see
/// `audio_font_scale_percent`). The Audio Scaling setting decides the room the card
/// is laid out over and the size the card is drawn at together (see `audio_box_room`),
/// so the text font size no longer moves the card at all: the card's font is the
/// share's own, and what a change to Text Preview → Font Size does to a text preview
/// is something a sound's card never does.
pub(super) fn current_audio_options() -> AudioPreviewOptions {
    CONFIG
        .lock()
        .map(|cfg| AudioPreviewOptions {
            theme: cfg.theme,
            font_scale_percent: audio_font_scale_percent(cfg.audio_scale),
        })
        .unwrap_or(AudioPreviewOptions {
            theme: TextTheme::Light,
            font_scale_percent: audio_font_scale_percent(DEFAULT_AUDIO_SCALE),
        })
}

/// The theme a sound's card is painted with, read the way
/// `current_audio_options` reads it: from the configuration, for the
/// one setting that still reaches a card already on screen. The Audio
/// Scaling share a card is drawn at is the share it was drawn with for
/// as long as that drawing is on screen (see `MediaData::audio_scale`),
/// but the theme is not a size — it is a re-styling, and the repaints
/// of a card on screen compose it over the font the remembered share
/// names (see `MediaData::refresh_audio_card` and
/// `MediaData::relayout_audio_card`).
pub(super) fn current_audio_theme() -> TextTheme {
    CONFIG
        .lock()
        .map(|cfg| cfg.theme)
        .unwrap_or(TextTheme::Light)
}

/// The options a sound's card is drawn with, from the share its card
/// is drawn at: the theme the configuration now answers — the Theme
/// setting still re-styles a card on screen — over the font the share
/// names, which is the size the card keeps for as long as that drawing
/// is on screen (see `MediaData::audio_scale` and
/// `PinnedPreview::audio_scale`).
pub(super) fn audio_options_of(scale: PreviewScale) -> AudioPreviewOptions {
    AudioPreviewOptions {
        theme: current_audio_theme(),
        font_scale_percent: audio_font_scale_percent(scale),
    }
}

/// The share of the display a sound's card is laid out over, read the way the
/// options are: from the configuration at the moment a box is measured, so a
/// change in the tray reaches the next hover rather than a card already on
/// screen (see `audio_box_room`).
pub(super) fn current_audio_scale() -> PreviewScale {
    CONFIG
        .lock()
        .map(|cfg| cfg.audio_scale)
        .unwrap_or(DEFAULT_AUDIO_SCALE)
}

/// How long a sound's probe is given — the source reader's own read of a file, or FFmpeg's
/// container probe — before the wait for it is over. The same cap a video's probe has, and for
/// the same reason: a probe that has run this long is a file nothing is coming back from.
pub(super) const AUDIO_PROBE_TIMEOUT_SECS: u64 = 10;

/// How often a sound's card is painted again while a player is running. The clock changes once
/// a second and the bar creeps by a few pixels in that time, so four times a second is smooth
/// to the eye and a fraction of what a video's own frames cost.
pub(super) const AUDIO_CARD_REPAINT: Duration = Duration::from_millis(250);

/// How often a card whose name does not fit is painted again, which is the same question asked
/// for the one thing about a card that moves faster than a clock: a name is scrolled across it
/// at about an advance a 100 ms, and a cadence coarse enough for a clock to read smoothly would
/// show that as a slideshow. A card is some four hundred pixels square, so what thirty of them
/// a second costs is a fraction of what the spinner's own overlay costs at the same rate (see
/// `AUDIO_CARD_REPAINT`).
pub(super) const AUDIO_NAME_REPAINT: Duration = Duration::from_millis(33);

/// A card that is stale the moment it is asked about, rather than at the next of its own
/// repaints: a sound a key has held or let go, a second a press on the card's own bar has taken
/// the file to, and a pass that has gone round under a card still drawn at the end of the last
/// one.
///
/// The first two are answered by the loop rather than by the window that asked — the clock a card
/// is drawn from is this loop's (see `toggle_pinned_audio` and `settle_pinned_audio_seek`) — and
/// the third by the loop alone, since the player is what ends a pass and the card is then at the
/// other end of the file. A card a quarter of a second behind the hand that just used it, or a
/// quarter of a second behind the sound that just went round, is a card that has not caught up,
/// and the quarter of a second is the whole of what the cadence above is.
pub(super) static AUDIO_CARD_DIRTY: AtomicBool = AtomicBool::new(false);

/// Whether the preview of `path` is a sound — the form a caller with no hover of its own asks.
/// See `HoverFacts::is_audio` for what the answer is.
///
/// What this used to do and no longer does is worth a line: it copied the audio list out of the
/// configuration under one lock, read the file's entry, took the lock a second time for the
/// content answer, and then compared the list itself — a deep copy of sixteen `Vec<String>` and
/// a `PathBuf` each, two hundred allocations, to read one list and one flag, on every hover of
/// a sound, and two locks where one was asked for.
pub(super) fn drawn_as_audio(path: &Path) -> bool {
    HoverFacts::read(path).is_audio()
}

/// What a sound's card says, with the clock as it stands — or nothing for a file with no track
/// behind it, which is a file no probe has answered for or one nothing here can play.
///
/// `name_offset` is how far a name the card has no room for has been scrolled: it is nothing
/// for the card a hover is measured with and for the first frame of one, and what the repaints
/// of a moving card hand over (see `audio_preview::NameScroll`).
///
/// `chrome` is what the card's own controls are saying, and nothing at all for a card a hover
/// shows — a hover's own window is a window nobody is in, so a click on one of its buttons would
/// land on the file behind it (see `audio_preview::Card::controls`). It is handed over rather than
/// asked for under the media's own lock, which is what `refresh_audio_card` and
/// `relayout_audio_card` are both reached holding (see `pinned_audio_chrome`).
pub(super) fn audio_card(
    path: &Path,
    elapsed: Option<f64>,
    duration: Option<f64>,
    name_offset: i32,
    chrome: Option<CardChrome>,
) -> Option<Card> {
    let Probed::Track(track) = audio_track::probed(path) else {
        return None;
    };

    Some(Card {
        name: audio_preview::name_of(path),
        facts: audio_preview::facts_of(&track, path),
        duration: duration.or(track.duration),
        elapsed,
        name_offset,
        controls: chrome,
    })
}

/// The card a sound is previewed as, painted into the box the layout settled on.
pub(super) fn load_audio_card(path: &Path, width: u32, height: u32, dpi: u32) -> Option<MediaData> {
    // No controls: this is the card a *hover* shows, whose own window is a window nobody is in
    // (see `audio_card`).
    let card = audio_card(path, None, None, 0, None)?;
    // The share the frame is painted at is the frame's own from
    // here on: every repaint of it paints with the options that
    // share names rather than asking the configuration again (see
    // `MediaData::audio_scale`), so a change to Audio Scaling
    // reaches the next hover and the next pin only — and the pin
    // taken up over this hover remembers the very same share (see
    // `take_up_pinned_window`).
    let scale = current_audio_scale();
    let (pixels, width, height) =
        audio_preview::render(&card, width, height, dpi, audio_options_of(scale))?;

    Some(MediaData {
        frames: vec![Arc::new(ImageFrame::new(pixels, width, height, 0))],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Audio,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
        audio_scale: Some(scale),
    })
}

/// What this machine has for playing a sound: Windows' own decoders where one of them reaches
/// the format, and FFmpeg's player where none does.
///
/// The two are asked in the order the chain names them, and the first that answers is the
/// player the file is played by: the engine answers for a file it can decode — which is the
/// whole of what its probe is for — and everything else is FFmpeg's, where FFmpeg is installed
/// at all. A file neither answers for is a file with no preview.
///
/// The loudness of a file that plays is measured here as well, where the tray's `Normalize` asks for
/// it: one read of one file, on the thread the hover is waiting on anyway, and the gain every
/// hover after this one is played at (see `measure_audio_gain`).
pub(super) fn probe_audio_track(path: &Path) -> Option<audio_track::Track> {
    let track = match video_player::audio_probe(path) {
        Some(track) => Some(track),
        None => ffprobe_audio_track(path),
    };

    if track.is_some() {
        measure_audio_gain(path);
    }

    track
}

/// What FFmpeg's own probe reports about a file, and the player that would play it.
///
/// The container is opened and its streams read rather than the file played to find out, which
/// is the pair of answers this side wants: whether there is a sound in the file at all, and
/// what the card beside it says. A machine without FFmpeg is answered by the engine above or
/// not at all.
pub(super) fn ffprobe_audio_track(path: &Path) -> Option<audio_track::Track> {
    if !codecs::ffplay_available() {
        return None;
    }

    // Spawned rather than run through `Command::output` so that the probe is in the job before
    // it is waited on, and waited for under a cap rather than for as long as it takes — the
    // same arrangement the video path's own probes have (see `wait_bounded`).
    let child = engine_processes::hidden_command("ffprobe")
        .args([
            "-v",
            "error",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-show_entries",
            "format=duration:stream=codec_type,codec_name,sample_rate,channels,bit_rate",
            "-of",
            "default=noprint_wrappers=1",
        ])
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    engine_processes::adopt(child.id());

    let output = wait_bounded(child, Duration::from_secs(AUDIO_PROBE_TIMEOUT_SECS))?;
    let report = String::from_utf8_lossy(&output.stdout);

    audio_track_from_report(&report)
}

/// The track an `ffprobe` report describes, or nothing where the file holds no sound.
///
/// The entries arrive one stream at a time, so what follows a `codec_type=audio` line is that
/// stream's own fields and nothing of the streams before it — which is what makes a film with a
/// soundtrack distinguishable from a song.
pub(super) fn audio_track_from_report(report: &str) -> Option<audio_track::Track> {
    let mut in_audio = false;
    let mut heard_audio = false;
    let mut codec: Option<String> = None;
    let mut rate: Option<u32> = None;
    let mut channels: Option<u16> = None;
    let mut bitrate: Option<u32> = None;
    let mut duration: Option<f64> = None;

    for line in report.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        let value = value.trim();

        match key.trim() {
            "codec_type" => {
                in_audio = value == "audio";
                heard_audio |= in_audio;
            }
            "codec_name" if in_audio => codec = Some(codec_label(value)),
            "sample_rate" if in_audio => rate = value.parse().ok(),
            "channels" if in_audio => channels = value.parse().ok(),
            "bit_rate" if in_audio => bitrate = value.parse().ok(),
            "duration" => duration = value.parse().ok(),
            _ => {}
        }
    }

    heard_audio.then_some(audio_track::Track {
        player: Player::Ffmpeg,
        codec,
        rate: rate.filter(|rate| *rate > 0),
        channels: channels.filter(|channels| *channels > 0),
        bitrate: bitrate.filter(|bitrate| *bitrate > 0),
        duration: duration.filter(|duration| *duration > 0.0),
    })
}

/// What a codec is called, by the name FFmpeg writes it under: the words a person reads on a
/// label where the codec has one, and the name itself where it does not.
pub(super) fn codec_label(name: &str) -> String {
    match name {
        "mp3" => "MP3",
        "flac" => "FLAC",
        "alac" => "ALAC",
        "aac" => "AAC",
        "opus" => "Opus",
        "vorbis" => "Vorbis",
        "speex" => "Speex",
        "wmav1" | "wmav2" | "wmapro" => "WMA",
        "wmalossless" => "WMA Lossless",
        "ac3" => "Dolby Digital",
        "eac3" => "Dolby Digital Plus",
        "dts" => "DTS",
        "ape" => "Monkey's Audio",
        "wavpack" => "WavPack",
        "tta" => "True Audio",
        "musepack" | "mpc7" | "mpc8" => "Musepack",
        "shorten" => "Shorten",
        "tak" => "TAK",
        "amrnb" => "AMR",
        "amrwb" => "AMR-WB",
        "cook" | "atrac3" | "atrac3p" | "sipr" => "RealAudio",
        "dsd_lsbf" | "dsd_msbf" | "dsd_lsbf_planar" | "dsd_msbf_planar" => "DSD",
        name if name.starts_with("pcm_") => "PCM",
        name => return name.to_uppercase(),
    }
    .to_string()
}

/// How long a sound's loudness is given — a decode of every sample in the file, which is what a
/// loudness that is the file's own costs — before the wait for it is over.
///
/// It is longer than the probe's own cap for what it reads: a probe reads a header, and this reads
/// the file. The cap is for the file no meter is coming back from, and one this side gives up on
/// is played as it holds rather than waited for (see `finish_gain_scan`).
pub(super) const AUDIO_GAIN_TIMEOUT_SECS: u64 = 15;

/// The sounds whose loudness is being measured right now, which is what keeps a file hovered
/// twice in the time one measurement takes from being read twice (see `begin_gain_scan`).
pub(super) static MEASURING_GAIN: Lazy<Mutex<Vec<PathBuf>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Reads in flight this list holds before it is emptied, the same bound and the same reasoning as
/// the measure list's own (see `MEASURING_MAX_ENTRIES`).
pub(super) const MEASURING_GAIN_MAX_ENTRIES: usize = 64;

/// Whether a sound's loudness is measured and applied at all: the setting is on and this machine
/// has the two programs that make it possible (see `codecs::normalize_available`).
pub(super) fn normalizing_audio() -> bool {
    let wanted = CONFIG
        .lock()
        .map(|config| config.normalize_volume)
        .unwrap_or(DEFAULT_NORMALIZE_VOLUME);

    normalize_available_for(wanted)
}

/// The same question for a video's soundtrack, which is a setting of its own and off where the app
/// starts (see `normalize_video_volume`).
pub(super) fn normalizing_video() -> bool {
    let wanted = CONFIG
        .lock()
        .map(|config| config.normalize_video_volume)
        .unwrap_or(DEFAULT_NORMALIZE_VIDEO_VOLUME);

    normalize_available_for(wanted)
}

/// Whether a loudness may be measured and applied at all, for whichever kind asked for it.
pub(super) fn normalize_available_for(wanted: bool) -> bool {
    wanted && codecs::normalize_available()
}

/// Say that this file's loudness is being measured, answering whether one already is.
pub(super) fn begin_gain_scan(path: &Path) -> bool {
    let Ok(mut measuring) = MEASURING_GAIN.lock() else {
        return false;
    };

    if measuring.iter().any(|running| running == path) {
        return false;
    }

    if measuring.len() >= MEASURING_GAIN_MAX_ENTRIES {
        measuring.clear();
    }

    measuring.push(path.to_path_buf());

    true
}

/// Say that the measurement of this file's loudness is done with.
pub(super) fn end_gain_scan(path: &Path) {
    let Ok(mut measuring) = MEASURING_GAIN.lock() else {
        return;
    };

    measuring.retain(|running| running != path);
}

/// Measure a sound's loudness where `Normalize` asks for it and nothing has measured it yet, on the
/// thread that asked — which is a thread a hover is waiting on and never the tick (see
/// `probe_audio_track`).
pub(super) fn measure_audio_gain(path: &Path) {
    measure_gain(path, normalizing_audio());
}

/// The measurement itself, for whichever kind asked for it: one read per file, held whether or not
/// it answered (see `finish_gain_scan`).
pub(super) fn measure_gain(path: &Path, wanted: bool) {
    if !wanted || audio_track::gain(path).is_some() {
        return;
    }

    if begin_gain_scan(path) {
        finish_gain_scan(path);
    }
}

/// The same measurement on a thread of its own, for a sound that was probed before `Normalize`
/// was on — or probed on a machine that had no FFmpeg to measure it with.
///
/// Nothing waits for it and nothing is held up by it: the sound being hovered is played as the
/// file holds it, and what it measures is measured for the hover after this one (see
/// `start_audio_playback`).
pub(super) fn spawn_gain_scan(path: &Path) {
    if audio_track::gain(path).is_some() || !begin_gain_scan(path) {
        return;
    }

    let path = path.to_path_buf();

    std::thread::spawn(move || finish_gain_scan(&path));
}

/// Measure a sound's loudness and hold the gain it asked for, which is where every measurement of
/// one ends.
///
/// A measurement that answered nothing — a meter that failed, a file with no sound stream in it,
/// a read that ran past its cap — is held as a gain of one rather than left unmeasured: what such
/// a file is played at is what it holds, and a file that cannot be measured is not measured again
/// on every hover of it.
pub(super) fn finish_gain_scan(path: &Path) {
    let gain = measure_audio_loudness(path).unwrap_or(1.0);

    audio_track::remember_gain(path, gain);
    end_gain_scan(path);
}

/// The level a sound is brought to before it is played, in the units the meter reads: ITU-R
/// BS.1770 integrated loudness, where `0 LUFS` is a full-scale sine and a number is how loud the
/// file *sounds* rather than where its loudest sample happens to sit.
///
/// It is `-14 LUFS` because that is the level everything a sound is likely to be played beside is
/// already at — every music service normalizes to it, and it is what a listener's ear has been
/// trained on rather than what the format's ceiling allows. That last part is the whole of the
/// difference between this and what the switch used to do, and it is deliberate: the target is a
/// property of the app, so two files of the same loudness are given the same gain whatever their
/// peaks happen to be. A target read off the file — the gain that brings *its* loudest sample to
/// full scale — is a target the file chooses, and it makes the gain a statement about how hard the
/// file was limited: two Suno renders at the same `-13.3 LUFS`, one limited to a peak of `0 dBFS`
/// and one to `-2 dBFS`, came out 2 dB apart, which is the bug.
///
/// A file already at the target is therefore played exactly as it holds — a gain of one — and
/// that is an answer rather than a lack of one: nothing has to be measured for a caller to tell
/// it apart from a file nothing has measured (see `audio_track::gain`).
pub(super) const NORMALIZE_TARGET_LUFS: f64 = -14.0;

/// How far above a file's own loudest inter-sample peak it may be lifted: the ceiling a gain that
/// carries a file's peaks up is not carried past.
///
/// Loudness is an average and a peak is a worst case, so the two disagree — a heavily limited
/// master measures loud and peaks at `0 dBFS`, and lifting it to `-14 LUFS` from a quieter one is
/// worth twenty decibels that land on samples that are already at the ceiling, where a decoder
/// clips them into the flat top nobody mixed on purpose. `-1 dBTP` is the ceiling every streaming
/// codec works to and is short of that.
///
/// It is a clamp and not a limit on what a file may be played at: a file that needs more gain
/// than its headroom allows is given the headroom and stops there, left quieter than the target
/// rather than torn. The cost is that two files needing more than they have get gains decided by
/// their peaks again, in the one case where there is nothing else to decide them by — a clipped
/// file and a clipped file are both already torn.
pub(super) const NORMALIZE_PEAK_CEILING_DBFS: f64 = -1.0;

/// The gain that brings a sound to the level the tray's `Normalize` plays files at, as FFmpeg's own
/// `ebur128` measures it — or nothing where it could not be asked.
///
/// `ebur128` is the ITU-R BS.1770 scanner, and what this asks of it is the `Summary` it writes at
/// the end of the pass: `I:` is the integrated loudness of the whole file and `Peak:` its true
/// peak, and the gain is what carries the first to the target without carrying the second past the
/// ceiling (see `NORMALIZE_TARGET_LUFS` and `NORMALIZE_PEAK_CEILING_DBFS`). `peak=true` is what
/// asks for that peak at all, and `framelog=quiet` is what stops the scanner writing a line for
/// every hundredth of a second on the way — a report of forty lines rather than one of two
/// thousand, which is a file's own worth of work thrown away on a hover.
///
/// `loudnorm` is the other meter FFmpeg has and it answers the same question with a JSON object
/// rather than a summary, and it was not the one to ask: its own pass runs the whole of a dynamic
/// normalization to produce the numbers, which is 8× the decode for an answer this throws away —
/// 0.45s against 3.8s for a three-minute file.
///
/// Only the file's own sound stream is handed to the meter, and what it writes is nothing — the
/// pass is the measurement (see `Normalize`).
pub(super) fn measure_audio_loudness(path: &Path) -> Option<f64> {
    if !codecs::normalize_available() {
        return None;
    }

    // Spawned and waited for under a cap, the arrangement every probe here has: a meter left
    // behind by a crash is one the job ends, and one that has not answered by the deadline is
    // killed rather than waited for (see `wait_bounded`).
    let child = engine_processes::hidden_command("ffmpeg")
        .args(["-v", "info", "-nostdin", "-hide_banner"])
        .arg("-i")
        .arg(path)
        .args(["-map", "0:a:0", "-af", "ebur128=peak=true:framelog=quiet"])
        .args(["-f", "null", "-"])
        .stdout(Stdio::null())
        .stderr(Stdio::piped())
        .spawn()
        .ok()?;

    engine_processes::adopt(child.id());

    let output = wait_bounded(child, Duration::from_secs(AUDIO_GAIN_TIMEOUT_SECS))?;
    let report = String::from_utf8_lossy(&output.stderr);

    audio_gain_from_report(&report)
}

/// The gain an `ebur128` report asks for: what brings a file's measured loudness to the target,
/// held short of its own true peak's ceiling.
///
/// What is read is the `Summary` block rather than the first line that looks like it: a scanner
/// told to log every frame writes `I:` and `TPK:` on each of them as it goes, and those are the
/// momentary readings of a moment — the summary is the whole of the file.
///
/// Nothing finite to read is nothing to apply. A file of silence reports a true peak of `-inf` dB
/// and an integrated loudness at the floor of the scale, and what either of those would ask for is
/// every sample of it times an infinity — a file with nothing to hear is left as it is rather than
/// lifted to whatever the format's noise floor happens to be.
pub(super) fn audio_gain_from_report(report: &str) -> Option<f64> {
    let summary = report.split_once("Summary:")?.1;

    let integrated = read_ebur128(summary, "I:")?;
    let peak = read_ebur128(summary, "Peak:")?;

    if !integrated.is_finite() || !peak.is_finite() {
        return None;
    }

    let wanted = 10f64.powf((NORMALIZE_TARGET_LUFS - integrated) / 20.0);
    // What the file's own headroom allows, and unity where it has none to give: a gain below one
    // makes the file quieter and cannot clip it, so the ceiling does not bind on a reduction —
    // which is what lets a file whose peaks already stand past it still be brought to the target
    // rather than pushed further from it.
    let ceiling = 10f64.powf((NORMALIZE_PEAK_CEILING_DBFS - peak) / 20.0);

    let gain = wanted.min(ceiling.max(1.0));

    (gain.is_finite() && gain > 0.0).then_some(gain)
}

/// One number out of an `ebur128` summary: the first line under the name, which is a field whose
/// width moves with the number in it and so is read by whitespace rather than by column.
pub(super) fn read_ebur128(summary: &str, field: &str) -> Option<f64> {
    summary
        .lines()
        .find_map(|line| line.split_once(field))
        .and_then(|(_, value)| value.split_whitespace().next())
        .and_then(|value| value.parse::<f64>().ok())
}

/// Which player is playing `path` now: the file's own answer, unless `Normalize` has measured a
/// gain for it, which is FFmpeg's to apply and not the engine's.
///
/// The probe's answer is the machine's own question — does this engine have a decoder for the
/// file — and it is the right answer to *start* a sound in on its own. It is the wrong answer to
/// every question asked while the sound plays, because `start_audio_playback` plays a file with a
/// gain through FFmpeg whatever the probe said: a gain is a filter, and the engine Windows has
/// cannot be handed one. So a quiet MP3 — anything `Normalize` measures as short of the target —
/// is probed `Native` and played by FFmpeg, and a site that branches on the probe alone then
/// speaks to an engine with no session while the sound is being heard from a process.
///
/// Every question about a playing sound therefore asks this rather than `track.player`: the pause
/// a key asks for, the seek a press on the bar asks for, and the clock a card is drawn from.
pub(super) fn playing_player(path: &Path, track: &audio_track::Track) -> Player {
    player_for_gain(track.player, normalizing_audio(), audio_track::gain(path))
}

/// The same answer with the three facts it is made of handed in rather than asked for, which is
/// what makes it testable on a machine with no FFmpeg in it — the whole question is whether a gain
/// is in play, and where the gain came from is the caller's business.
///
/// A gain of one is not the absence of one: it is what a file already standing at the target
/// measures as, and it is a filter like any other, so it goes to FFmpeg the same way.
pub(super) fn player_for_gain(probed: Player, normalizing: bool, gain: Option<f64>) -> Player {
    if normalizing && gain.is_some() {
        return Player::Ffmpeg;
    }

    probed
}

/// Start the player a sound's card is drawn against, answering whether a player that was
/// expected arrived.
///
/// A card is drawn whether or not anything plays, and a level of nothing is a player like any other
/// level: what the caller is told is whether the player that *was* asked for came up — a sound no
/// engine here will actually play is a hover answered with nothing rather than a card whose clock
/// can never move.
///
/// `start` is where in the file the sound is dropped, and it is the caller's answer: it is a
/// question about the file's length and the tray's `Volume → Audio Seek`, both of which are
/// read where the hover is answered (see `audio_seek::start_position`).
///
/// Which player a sound is started in is settled here too, and the gain measured for the file is
/// the whole of that question: where the tray's `Normalize` is on and a gain has been measured for
/// the file, FFmpeg's player is the one it is started in — a gain is a filter there, and the engine
/// Windows has cannot be handed one — while every other file is played by whichever engine its own
/// probe answered for.
pub(super) fn start_audio_playback(path: &Path, media: &mut MediaData, start: f64) -> bool {
    start_audio_playback_at(path, media, start, current_audio_volume())
}

/// The same, at a level the caller names rather than the one the tray names: what a pin's own
/// player is begun at, where the level belongs to the window it was moved on and not to the setting
/// a hover reads (see `PinVolume`).
///
/// `volume` is a parameter for that reason alone — the rest of this is one function either way, and
/// two copies of it would be two answers to one question about which player a file is started in.
pub(super) fn start_audio_playback_at(
    path: &Path,
    media: &mut MediaData,
    start: f64,
    volume: u32,
) -> bool {
    let Some(track) = audio_track::playable(path) else {
        return false;
    };

    // A player the hover before this one left behind is ended before this one starts, which is
    // the check the video path makes before it spawns its own: a sound must not go on playing
    // over the sound of the file the pointer has moved to. The engine's own session is stopped
    // by `play_audio` rather than here, and this is what answers for the other engine — a
    // player that has not been confirmed gone, whose process handle the hover that started it
    // took with it when it ended.
    kill_stray_video_process();

    // A level of nothing is handed to the player like any other level rather than answered here:
    // silence is what a player at zero is, not the absence of one, so the sound goes on playing and
    // stays seekable with nothing to hear. Both players already read a zero as silence while they
    // run — the engine as a mute (`video_player::Session::begin`) and FFmpeg as a level of its own
    // scale, or a gain of nothing where a file has been measured (`start_audio_player`).
    if normalizing_audio() {
        match audio_track::gain(path) {
            // A gain of one is a gain: a file measured as already standing at the target is played
            // through the filter like every other measured file, rather than being read as a file
            // that has nothing to apply.
            Some(gain) => {
                // The engine's own session is let go first, which is what a file that has just been
                // handed to the other player needs: a sound already playing natively — the hover
                // that started before this file's loudness was measured — must not go on playing
                // over the gain it asked for (see `video_player::stop`).
                video_player::stop();

                media.video_process = start_audio_player(path, volume, start, Some(gain));

                return media.video_process.is_some();
            }
            // Nothing has measured the file, so this hover is played as the file holds it: a decode
            // of every sample is not work for the tick, and what the scan is started for is the
            // hover after this one (see `spawn_gain_scan`).
            None => spawn_gain_scan(path),
        }
    }

    // Which player this file is played by, which is the answer the rest of the app asks
    // `playing_player` for while the sound is up (see it).
    match playing_player(path, &track) {
        Player::Native => {
            // A hover that lands on the file already playing leaves it playing, the same way
            // the FFmpeg path compares the file it last started.
            if video_player::playing_path().as_deref() == Some(path) && video_player::is_playing() {
                return true;
            }

            video_player::play_audio(path, volume, start);
            video_player::is_playing()
        }
        Player::Ffmpeg => {
            // A gain measured for the file is FFmpeg's filter and goes with it; the branch above
            // is where a file with one was already taken, so this is a player started with no gain.
            media.video_process = start_audio_player(path, volume, start, None);
            media.video_process.is_some()
        }
    }
}

/// Start FFmpeg's player on a sound: no window at all, which is the whole of what this side asks
/// of it, and the level the sound is played at — the player's own scale, `0` to `100`, where the
/// file is played as it holds it, and the filter that carries the gain the tray's `Normalize`
/// measured for it (see `Normalize`).
///
/// A sound is looped while it is hovered, as a video is: what a hover is for is the file, and a
/// sound that stopped under a pointer that had not moved would be a preview that ended on its
/// own. The card's clock wraps with it (see `audio_clock`).
///
/// Where the sound starts is `-ss`, and it is an option of the *input* rather than of the
/// player: what it does is seek the file before anything of it is read, which is a player that
/// begins at that second rather than one that plays its way there. What it costs is nothing:
/// the seek is the player's own, and nothing is decoded before it.
///
/// A player given such a second is given no loop, and that is `ffplay`'s arrangement rather than
/// this side's: the position its own loop seeks back to *is* the position it was started at —
/// `-ss` is the one value that seek-back reads — so a sound dropped half way into a file would
/// go round the second half of it for as long as it was hovered. That pass is played once
/// instead, `-autoexit` being what ends it at the end of the file, and what this side starts in
/// its place plays the whole file and loops from the beginning of it (see
/// `wrap_audio_player`): every pass after the first goes back to 0:00 whatever the file is and
/// whatever the setting started it at. It is the same rule the engine Windows has is held to,
/// where `SetLoop` restarts the whole presentation rather than a seek into it — one answer for
/// both players, and this side's own hand under the half that has none.
///
/// A sound that starts at the beginning of its file is that second player already, so it is
/// given that player's own loop rather than a pass for this side to restart after.
pub(super) fn start_audio_player(
    path: &Path,
    volume: u32,
    start: f64,
    gain: Option<f64>,
) -> Option<Child> {
    let mut command = engine_processes::hidden_command("ffplay");
    command.args(["-nodisp", "-autoexit", "-loglevel", "quiet"]);

    // The level the sound is played at, with the file's own measured gain folded into it where
    // one was: FFmpeg's player has a scale of its own for a level and a filter for a gain, and the
    // two multiply — what is handed to the filter is the whole of what the file is scaled by, so a
    // normalized file at `Volume → Audio` 10% is heard at a tenth of the level rather than at ten
    // times it.
    match gain {
        Some(gain) => {
            let level = format!("volume={:.4}", f64::from(volume) / 100.0 * gain);
            command.args(["-af", &level]);
        }
        None => {
            command.args(["-volume", &volume.min(100).to_string()]);
        }
    }

    if start.is_finite() && start > 0.0 {
        command.args(["-ss", &format!("{start:.3}")]);
    } else {
        // The whole file, looped by the player itself: with no position given to it, the
        // position its loop returns to is the beginning of the file.
        command.args(["-loop", "0"]);
    }

    let child = command
        .arg(path)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .ok()?;

    // The player is this app's own child, taken charge of the way every other one is: the job
    // ends it when the app does, and the record answers for a run that never got to end it. It
    // is recorded as the player rather than as an engine, because what ends it is its hover
    // ending.
    engine_processes::record_player(VIDEO_PROCESS_IMAGE_NAME, child.id());
    VIDEO_PID.store(child.id(), Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);

    Some(child)
}

/// How long a player has to have lived before its exit is read as the end of the file it was
/// given, where the file's own length says nothing shorter than this.
///
/// A player that has stopped is either one that played its file through or one that never played
/// it at all — a machine with no output device, a decoder that will not have the file — and the
/// second of those stops within a moment of starting. Time is the only thing that tells the two
/// apart, and what it is spent on is the difference between a sound that goes round again and a
/// process spawned a tick apart for as long as the file is hovered. What this sits at is past
/// every failure a player of these files has and under the pass any file is hovered for, and it
/// is a ceiling on what is asked rather than the bar itself: a pass shorter than it — the last
/// moment of a short file, which `Random` can land in — is asked only to have been played (see
/// `reached_the_end`).
pub(super) const AUDIO_WRAP_MINIMUM: f64 = 1.0;

/// Whether a player that has stopped is one that reached the end of the file it was handed.
///
/// What the player was given is the file from the second the sound was dropped in to its own
/// end, so the length the file says it has is what that pass takes — except that a length is a
/// container's own reading of a header and is a little out for some formats, which is why it is
/// used as a ceiling on what is asked for rather than as the answer itself: a stop is read as the
/// end of the file where the player lived for [`AUDIO_WRAP_MINIMUM`], or for the whole of a pass
/// shorter than that. A player cannot have lived past the end of the pass it was given, and the
/// last moment of a short file is a pass of its own.
pub(super) fn reached_the_end(played: Duration, length: Option<f64>, offset: f64) -> bool {
    let pass = length
        .map(|length| (length - offset).max(0.0))
        .unwrap_or(AUDIO_WRAP_MINIMUM);

    played.as_secs_f64() >= pass.min(AUDIO_WRAP_MINIMUM)
}

/// Put a sound FFmpeg plays round to the beginning of its file where the player it was given has
/// reached the end of it.
///
/// The player this side starts for a sound dropped into the middle of a file plays that pass and
/// stops — see `start_audio_player` for why the loop cannot be the player's own — so the end of
/// the file is the player's exit, and what is started in its place is a player of the whole file
/// which loops from the beginning of it. Every pass after the first therefore goes back to 0:00,
/// whatever the file is and whatever the setting started it at.
///
/// What the card's clock is drawn from moves with the player, and that is the whole of what a
/// wrap is on this side: the moment the sound was put in at is replaced by the moment the new
/// player started, and the position the clock is counted from by the beginning of the file. A
/// player that reports nothing at all is a clock of this app's, and a clock still counted from
/// the second the *old* pass was dropped in at would have the card say the sound was a minute
/// into a file whose playing was heard to begin.
///
/// A player whose stop is not read as the end of its file by `reached_the_end` — one that never
/// played anything — is not started again: the card is left with no clock rather than with
/// another player, which is the answer a file this machine will not play gets.
pub(super) fn wrap_audio_player(
    media: &mut MediaData,
    path: &Path,
    started: &mut Option<Instant>,
    offset: &mut f64,
) {
    let Some(process) = media.video_process.as_mut() else {
        return;
    };

    // The field holds the player this hover started and no other, and a process that has been
    // waited on is a process that has ended: a player that is still going is not one to replace.
    if matches!(process.try_wait(), Ok(None)) {
        return;
    }

    let played = started.map(|at| at.elapsed()).unwrap_or_default();
    let length = audio_track::playable(path).and_then(|track| track.duration);

    // The player is confirmed gone — a process that has been waited on has ended — so the record
    // of it goes with it rather than being left for the leftover-process sweep to find, exactly
    // as the end of a video's player is answered for (see `is_video_process_running`).
    let pid = process.id();
    media.video_process = None;
    VIDEO_HWND.store(0, Ordering::SeqCst);
    VIDEO_PID.store(0, Ordering::SeqCst);
    engine_processes::forget(pid);

    if !reached_the_end(played, length, *offset) {
        *started = None;
        AUDIO_CARD_DIRTY.store(true, Ordering::Release);
        return;
    }

    // The pass after this one is the whole file, started the way the hover started the player it
    // replaces: the same volume, the same care about a process left behind, and the same answer
    // as to whether a player arrived at all.
    start_audio_playback(path, media, 0.0);

    if media.video_process.is_some() {
        *started = Some(Instant::now());
        *offset = 0.0;
    } else {
        *started = None;
    }

    // The card is at the other end of the file from where it was drawn a moment ago, and this
    // side's clock no longer wraps itself, so nothing would ask for it again until the next of
    // its own repaints: a card sitting at the whole of a file for a quarter of a second after
    // the sound has gone round (see `AUDIO_CARD_DIRTY`).
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
}

/// Where the sound is and how long it is: the engine's own clock where Windows plays it, and
/// this app's clock over the player's start where FFmpeg does. The whole is the file's own
/// answer either way, and either half is nothing where there is nothing to say it — which is
/// what a card with no player behind it is drawn with.
///
/// A sound loops for as long as it is hovered, and where in the file it is *now* is the player's
/// own: a pass ends when the player this side started is seen to have ended, and the pass after
/// it begins at 0:00 with a bar that is empty (see `wrap_audio_player`).
///
/// The clock over the player's start is counted from the second the sound was put in at as well
/// — `from` — because `ffplay` reports nothing at all: a file dropped half way into itself draws
/// its clock and its bar at its middle only if this side counts the first half as already played,
/// and what a hover would otherwise show is a sound playing from its middle with a card saying it
/// has just begun.
///
/// `paused_at` is what a key answered by holding a pinned sound has to say in its place: a player
/// this app ended is a player that measures nothing, and a card with no clock at all is a card
/// that has forgotten where in the file the sound was left. A sound the engine Windows has is not
/// in this branch, because that engine holds where it is told to and keeps reporting that second
/// for itself (see `toggle_pinned_audio`).
pub(super) fn audio_clock(
    path: &Path,
    started: Option<Instant>,
    from: f64,
    paused_at: Option<f64>,
) -> (Option<f64>, Option<f64>) {
    let Some(track) = audio_track::playable(path) else {
        return (None, None);
    };

    if let Some(at) = paused_at {
        return (Some(at), track.duration);
    }

    // The engine's own clock is what a sound the engine plays is measured by, and a sound this side
    // started a player for is measured by the clock over that player's start — whichever engine the
    // machine's own decoders would have made of the file: a gain puts a file the engine could have
    // played into FFmpeg's hands, and a player of this side's reports nothing at all (see
    // `start_audio_playback` and `audio_started`).
    if started.is_none() && playing_player(path, &track) == Player::Native {
        // The engine's own clock has the seek in it — it is the engine that was taken to where
        // the sound starts — so what it reports is the position with nothing added to it.
        return (
            video_player::position(),
            video_player::duration().or(track.duration),
        );
    }

    let elapsed = started.map(|at| from + at.elapsed().as_secs_f64());

    // A sound still inside its file is held at the whole of it rather than wrapped round it, and
    // the end of a pass is the player's exit rather than the length: a container's length is a
    // header's reading of itself and is a little out for some formats (see `reached_the_end`), so
    // the two are two answers to one question that do not always agree — and where they disagree
    // the length is the one that is wrong, because a sound outlives its own header.
    //
    // Wrapping here instead meant a file whose header read short went back to 0:00 while the last
    // moment of it was still playing, and the card's bar with it: a sub-second position is a
    // pixel or two of bar and a clock still reading zero, so the bar stood a little way along a
    // card saying 0:00 and then stepped back to nothing as the player's exit turned the pass over
    // for real. A progress bar going backwards, once in a while, for a file whose header is a
    // little out — and this side's clock is not the thing that gets to say the pass ended.
    let position = match (elapsed, track.duration) {
        (Some(elapsed), Some(duration)) if duration > 0.0 => Some(elapsed.min(duration)),
        (elapsed, _) => elapsed,
    };

    (position, track.duration)
}
