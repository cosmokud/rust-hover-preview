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

/// A file whose sound stream states no rate of its own and whose container
/// does: the container's entries are written after every stream's and behind
/// the length, which is the order that tells its rate from a stream's.
const OPUS_WITHOUT_A_STREAM_RATE: &str = "\
codec_name=opus
codec_type=audio
sample_rate=48000
channels=2
bit_rate=N/A
duration=222.227500
bit_rate=96000
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

/// A track of the shape a probe answers for a sound, at a length: what the test below starts
/// from, since what it is about is the length a probe's own answer is replaced by.
fn a_track(codec: &str, duration: f64) -> audio_track::Track {
    audio_track::Track {
        player: audio_track::Player::Native,
        codec: Some(codec.to_string()),
        rate: Some(44_100),
        channels: Some(2),
        bitrate: Some(128_000),
        duration: Some(duration),
    }
}

/// A rate a container states where its sound stream states none of its own: the file's own
/// average is the one rate there is, and a card with no rate draws no `kbps` fact at all.
#[test]
fn a_rate_the_container_states_is_read_where_the_stream_states_none() {
    let track = audio_track_from_report(OPUS_WITHOUT_A_STREAM_RATE).expect("a track");

    assert_eq!(
        track.bitrate,
        Some(96_000),
        "the file's own rate is the answer where the stream's is nothing"
    );
}

/// And a stream's own rate is drawn where the file states one as well: what the sound is encoded
/// at is the number the card is about, and the file's average is that plus the container's own.
#[test]
fn a_streams_own_rate_is_drawn_rather_than_the_files_average() {
    let report = "codec_name=ac3\ncodec_type=audio\nsample_rate=48000\nchannels=6\n\
                  bit_rate=448000\nduration=7331.000000\nbit_rate=12000000\n";

    let track = audio_track_from_report(report).expect("a track");

    assert_eq!(track.bitrate, Some(448_000));
}

/// A file whose rate the probe left out and whose length it read: the average a card draws is the
/// file's own size over that length, which is the rate its container would have stated.
#[test]
fn a_length_the_probe_read_is_enough_for_the_file_average() {
    let folder = std::env::temp_dir().join("rust-hover-preview-audio-bitrate");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("no-rate.flac");
    std::fs::write(&path, vec![0u8; 320_000]).expect("a written file");

    // Twenty seconds of three hundred and twenty thousand bytes, which is a hundred and
    // twenty-eight kilobits a second.
    let track = audio_track::Track {
        bitrate: None,
        ..a_track("FLAC", 20.0)
    };
    let filled = with_a_bit_rate(track, &path);

    assert_eq!(
        filled.bitrate,
        Some(128_000),
        "the size over the length the probe stated, in bits a second"
    );

    let facts: Vec<String> = audio_preview::facts_of(&filled, &path)
        .into_iter()
        .map(|fact| fact.text)
        .collect();
    assert!(
        facts.iter().any(|fact| fact == "128 kbps"),
        "a card whose probe states no rate still draws one: {facts:?}"
    );

    // And a rate the probe did state is kept as it is, whatever the file's size says.
    let stated = with_a_bit_rate(a_track("FLAC", 20.0), &path);
    assert_eq!(stated.bitrate, Some(128_000), "the probe's own answer");

    let _ = std::fs::remove_dir_all(&folder);
}

/// A file nothing states a rate for and nothing states a length for draws no rate at all: a size
/// on its own is no rate, and a zero would be a rate the file does not have.
#[test]
fn a_file_with_no_rate_and_no_length_draws_no_rate() {
    let folder = std::env::temp_dir().join("rust-hover-preview-audio-bitrate-none");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("junk.mp3");
    std::fs::write(&path, b"\0\0\0\x18ftypM4A ").expect("a written file");

    let track = audio_track::Track {
        bitrate: None,
        duration: None,
        ..a_track("MP3", 180.0)
    };
    let filled = with_a_bit_rate(track, &path);

    assert_eq!(
        filled.bitrate, None,
        "a length nothing read leaves the size unusable, and FFmpeg has no sound to answer for"
    );

    let facts: Vec<String> = audio_preview::facts_of(&filled, &path)
        .into_iter()
        .map(|fact| fact.text)
        .collect();
    assert!(
        !facts.iter().any(|fact| fact.contains("bps")),
        "nothing is not zero: {facts:?}"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A raw AAC stream's length is the one thing neither engine can read off the file — both read an
/// estimate from the bit rate, which is out by seconds on ordinary streams and by minutes on long
/// ones — so it is counted where the probe calls the file AAC. What a card's bar is a share of is
/// then the file's own length rather than that estimate (see `aac_stream`).
#[test]
fn a_raw_aac_streams_length_is_counted_rather_than_estimated() {
    let folder = std::env::temp_dir().join("rust-hover-preview-aac-counted");
    std::fs::create_dir_all(&folder).expect("a test folder");

    // Three frames of forty-four one, which is a stream of three times 1024 samples of it, and
    // each frame written at a size of its own so that the count is not read off the file's.
    let mut bytes = Vec::new();
    for payload in [20, 300, 90] {
        bytes.extend(aac_stream::a_frame(4, payload));
    }
    let stream = folder.join("stream.aac");
    std::fs::write(&stream, &bytes).expect("a written file");

    assert_eq!(
        counted_length(a_track("AAC", 999.0), &stream).duration,
        Some(3.0 * 1024.0 / 44_100.0),
        "the length a card is drawn at is the frames the file holds, not the estimate"
    );

    // A sound that is not AAC is left exactly as the probe read it, whatever file is beside it.
    assert_eq!(
        counted_length(a_track("MP3", 180.0), &stream).duration,
        Some(180.0),
        "the walk is about a raw AAC stream and about nothing else"
    );

    // And a file the walk cannot answer for — here, one that is not ADTS frames at all — keeps
    // the length the probe read, which is the answer every other file gets.
    let other = folder.join("not-a-stream.aac");
    std::fs::write(&other, b"\0\0\0\x18ftypM4A \0\0\0\0M4A mp42isom").expect("a written file");

    assert_eq!(
        counted_length(a_track("AAC", 999.0), &other).duration,
        Some(999.0),
        "a file that is not a stream of frames keeps the length its container states"
    );

    let _ = std::fs::remove_dir_all(&folder);
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
            let mut paths: Vec<_> = files
                .filter_map(|file| file.ok().map(|file| file.path()))
                .collect();
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
