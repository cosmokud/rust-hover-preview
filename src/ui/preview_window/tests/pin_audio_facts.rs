use super::*;

/// The report `ffprobe` writes for a sound that carries its cover art inside
/// itself: an Ogg or Opus file, whose art is a `METADATA_BLOCK_PICTURE`
/// comment entry, which FFmpeg names as a video stream of its own beside the
/// sound — an attached picture, which is not the sound. The fields of each
/// stream arrive with the stream's codec name ahead of the line that says
/// what the stream is, which is the order the report is read in.
const OPUS_WITH_JPEG_ART: &str = "\
codec_name=opus
codec_type=audio
sample_rate=48000
channels=2
bit_rate=N/A
codec_name=mjpeg
codec_type=video
bit_rate=N/A
duration=222.227500
";

/// The same shape with the art a PNG carries: the picture stream's codec
/// name is the one a reader that takes the first name it is handed would
/// name the sound by.
const OPUS_WITH_PNG_ART: &str = "\
codec_name=opus
codec_type=audio
sample_rate=48000
channels=2
bit_rate=N/A
codec_name=png
codec_type=video
bit_rate=N/A
duration=358.960000
";

/// A film's report: the picture stream first and the soundtrack behind it,
/// which is the same order read from the other end — the sound's own name
/// arrives ahead of the line that says the stream is the sound.
const FILM_WITH_SOUNDTRACK: &str = "\
codec_name=h264
codec_type=video
bit_rate=1200000
codec_name=ac3
codec_type=audio
sample_rate=48000
channels=6
bit_rate=448000
duration=7331.000000
";

/// The track a report of a sound with its art inside it describes: the
/// sound's own codec, rate and channels — not the art's, which the report
/// names as a video stream beside the sound.
#[test]
fn a_sound_with_art_inside_it_is_named_by_its_own_codec() {
    let track = audio_track_from_report(OPUS_WITH_JPEG_ART).expect("a track");

    assert_eq!(track.player, audio_track::Player::Ffmpeg);
    assert_eq!(track.codec.as_deref(), Some("Opus"));
    assert_eq!(track.rate, Some(48_000));
    assert_eq!(track.channels, Some(2));
    assert_eq!(track.duration, Some(222.2275));
}

/// The same answer for the art a PNG carries: the picture's codec name is
/// not the sound's, whichever of the two the report names first.
#[test]
fn a_sound_with_a_png_inside_it_is_named_by_its_own_codec() {
    let track = audio_track_from_report(OPUS_WITH_PNG_ART).expect("a track");

    assert_eq!(track.codec.as_deref(), Some("Opus"));
    assert_eq!(track.rate, Some(48_000));
    assert_eq!(track.channels, Some(2));
    assert_eq!(track.duration, Some(358.96));
}

/// A film's track is its soundtrack's: the picture stream that arrives first
/// is not the stream the card names, and the soundtrack behind it is.
#[test]
fn a_film_is_named_by_its_soundtrack_not_its_picture() {
    let track = audio_track_from_report(FILM_WITH_SOUNDTRACK).expect("a track");

    assert_eq!(track.codec.as_deref(), Some("Dolby Digital"));
    assert_eq!(track.rate, Some(48_000));
    assert_eq!(track.channels, Some(6));
    assert_eq!(track.bitrate, Some(448_000));
    assert_eq!(track.duration, Some(7331.0));
}

/// The whole chain a hover runs, on the files it went wrong on: the track
/// the probes answer and the facts line the card draws from it. Ignored, and
/// driven by `RHP_AUDIO_FACTS_PROBE` —
/// `$env:RHP_AUDIO_FACTS_PROBE = "C:\art\sounds"; cargo test -- --ignored --nocapture audio_facts_probe`
/// — a folder of real sounds that carry their cover art inside themselves, or
/// one or more of their paths, separated by ';'.
#[test]
#[ignore = "reads the files named in RHP_AUDIO_FACTS_PROBE"]
fn audio_facts_probe() {
    let Ok(list) = std::env::var("RHP_AUDIO_FACTS_PROBE") else {
        println!("set RHP_AUDIO_FACTS_PROBE to a folder of files, or to one or more paths, separated by ';'");
        return;
    };

    // A folder is every file in it, sorted, which is the form a sweep of real
    // sounds takes; anything else is the paths themselves.
    let entry = std::fs::read_dir(&list).ok();
    let paths: Vec<PathBuf> = match entry {
        Some(files) => {
            let mut paths: Vec<_> = files.filter_map(|file| file.ok().map(|file| file.path())).collect();
            paths.sort();
            paths
        }
        None => list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
            .collect(),
    };

    // The readouts are gathered rather than asserted one file at a time, so
    // one run says what every file in the folder holds.
    let mut misnamed: Vec<String> = Vec::new();

    for path in paths {
        let name = path
            .file_name()
            .unwrap_or_default()
            .to_string_lossy()
            .into_owned();
        println!("\n--- {name} ---");

        let Some(track) = probe_audio_track(&path) else {
            panic!("{name}: no probe answers a playable track");
        };

        let facts: Vec<_> = audio_preview::facts_of(&track, &path)
            .into_iter()
            .map(|fact| fact.text)
            .collect();
        println!(
            "player: {:?}, codec: {:?}, facts: {}",
            track.player,
            track.codec,
            facts.join(" · ")
        );

        // The format fact is the sound's own codec, which is what a card
        // names a sound by — not the picture the file carries beside the
        // sound, which the report names as a video stream.
        if facts.first().map(String::as_str) != Some("Opus") {
            misnamed.push(format!("{name}: {:?}", facts.first()));
        }
    }

    assert!(
        misnamed.is_empty(),
        "the format fact names the sound, not its art: {}",
        misnamed.join("; ")
    );
}
