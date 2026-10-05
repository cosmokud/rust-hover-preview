use super::*;

/// Every track whose codec has a small form is copied by the one pass, at the name the reader
/// looks for — `sub<i>.<ext>` under the film's folder, numbered the way `-map 0:s:<i>` numbers —
/// and a codec with no small form is left out rather than failing the pass.
#[test]
fn the_one_pass_copies_every_track_that_has_a_small_form() {
    let folder = Path::new("cache");
    let film = Path::new("film.mkv");
    let codecs = [
        "ass".to_string(),
        "dvd_subtitle".to_string(),
        "subrip".to_string(),
        "hdmv_pgs_subtitle".to_string(),
    ];

    let args =
        extraction_args(folder, film, &codecs, &[]).expect("three copyable tracks make a command");

    assert_eq!(
        &args[..6],
        ["-y", "-v", "error", "-hide_banner", "-i", "film.mkv"],
        "the pass is quiet and the film is its input: {args:?}"
    );

    // `dvd_subtitle` has no small form this app draws, so it has no output — while its index
    // still belongs to the tracks that do, which is the whole of why a name carries the track's
    // own number rather than a count of what was written.
    for (index, extension) in [(0usize, "ass"), (2, "srt"), (3, "sup")] {
        let position = args
            .iter()
            .position(|argument| argument == &format!("0:s:{index}"))
            .expect("every copyable track has its own map");

        assert_eq!(
            args[position - 1],
            "-map",
            "the map names a stream: {args:?}"
        );
        assert_eq!(
            args[position + 1],
            "-c:s",
            "and the stream is copied: {args:?}"
        );
        assert_eq!(
            args[position + 2],
            "copy",
            "copied, not converted: {args:?}"
        );
        assert_eq!(
            Path::new(&args[position + 3]),
            folder.join(format!("sub{index}.{extension}")),
            "the output is the file the reader looks for, so the extension has to be the codec's \
             own: {args:?}"
        );
    }

    assert!(
        !args.iter().any(|argument| argument.contains("sub1.")),
        "the track with no small form is left out rather than written as something it is not: \
         {args:?}"
    );

    // A film none of whose tracks has a small form has no command at all: the pass would be the
    // whole read this exists to avoid and would produce nothing.
    assert_eq!(
        extraction_args(folder, film, &["dvd_subtitle".to_string()], &[]),
        None,
        "a read that produces nothing is the one thing the caller must never start"
    );
}

/// The container's own fonts are dumped by the same pass, at the index the container numbers its
/// attachment streams at, and only the fonts: a cover image is not a face `fontsdir` would read.
#[test]
fn the_one_pass_dumps_the_containers_fonts_and_nothing_else() {
    let folder = Path::new("cache");
    let film = Path::new("film.mkv");
    let codecs = ["ass".to_string()];
    let attachments = ["mjpeg".to_string(), "ttf".to_string(), "otf".to_string()];

    let args = extraction_args(folder, film, &codecs, &attachments)
        .expect("a copyable track makes a command");

    for (index, name) in [(1usize, "font1.ttf"), (2, "font2.ttf")] {
        let position = args
            .iter()
            .position(|argument| argument == &format!("-dump_attachment:t:{index}"))
            .expect("every font is dumped");

        assert_eq!(
            args[position + 1],
            name,
            "under its own relative name, because an absolute one is refused outright by FFmpeg \
             9.0.2 — measured, `Filename ... is unsafe` — and the subtitle files are not written \
             either when it is: {args:?}"
        );
    }

    assert!(
        !args.iter().any(|argument| argument.contains("font0")),
        "the first attachment is a cover image, and `fontsdir` would never read it: {args:?}"
    );
}

/// Whether an extraction is worth starting: a sidecar is already a small file to draw, files
/// already copied are the answer itself, a film with no subtitle streams has nothing to copy,
/// and a film none of whose tracks has a small form would be the whole read for nothing.
#[test]
fn an_extraction_is_due_only_for_a_film_that_has_something_to_copy_and_nowhere_to_read_it() {
    let copyable = ["ass".to_string()];
    let uncopyable = ["dvd_subtitle".to_string()];
    let copied = DerivedSubtitles {
        tracks: vec![Some(PathBuf::from("sub0.ass"))],
        fonts: None,
    };

    assert!(
        extraction_due(None, None, 1, &copyable),
        "an embedded track with no sidecar and nothing copied yet is the case this exists for"
    );
    assert!(
        !extraction_due(Some(Path::new("episode.srt")), None, 1, &copyable),
        "a sidecar beside the film is already a small file to draw"
    );
    assert!(
        !extraction_due(None, Some(&copied), 1, &copyable),
        "files already copied are the answer itself"
    );
    assert!(
        !extraction_due(None, None, 0, &copyable),
        "a film with no subtitle streams has nothing to copy"
    );
    assert!(
        !extraction_due(None, None, 1, &uncopyable),
        "a film whose only track has no small form would be a whole read producing nothing"
    );
}

/// What the folder already holds is read back off the names its files carry, by
/// subtitle-relative index — a track with no small form keeps its place as a hole rather than
/// shifting its neighbours — and a folder with nothing is the answer that nothing has been
/// copied yet.
#[test]
fn what_is_already_in_the_folder_is_answered_by_track_with_its_fonts() {
    let dir = std::env::temp_dir().join(format!("subtitle-files-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
    let film = dir.join("episode.mkv");
    std::fs::write(&film, b"stand-in").expect("a stand-in film is writable");

    assert!(
        resolve(&film, &["ass".to_string()]).is_none(),
        "nothing written yet is the answer the spawning hover needs"
    );

    let folder = film_folder(&film);
    std::fs::create_dir_all(&folder).expect("the film's own folder is creatable");
    std::fs::create_dir_all(folder.join(FONTS_FOLDER)).expect("its fonts folder is creatable");
    let copy = folder.join("sub1.ass");
    std::fs::write(&copy, b"[Script Info]\n").expect("a stand-in copy is writable");

    // The second track was copied and the first was not — the film's first subtitle track being
    // one whose codec has no small form, which is the shape the answers are checked against.
    let codecs = ["dvd_subtitle".to_string(), "ass".to_string()];
    let derived = resolve(&film, &codecs).expect("a copy means an answer");

    assert_eq!(
        derived.track(0),
        None,
        "a track whose codec has no small form has no file, and its place is a hole rather than \
         its neighbour's answer"
    );
    assert_eq!(
        derived.track(1),
        Some(copy.as_path()),
        "and the copied track is answered under its own number"
    );
    assert_eq!(
        derived.fonts, None,
        "a fonts folder that holds nothing is not named"
    );

    std::fs::write(folder.join(FONTS_FOLDER).join("font0.ttf"), b"stand-in")
        .expect("a stand-in font is writable");
    let derived = resolve(&film, &codecs).expect("a copy means an answer");
    assert_eq!(
        derived.fonts.as_deref(),
        Some(folder.join(FONTS_FOLDER).as_path()),
        "a fonts folder that holds a face is named with the track"
    );

    let _ = std::fs::remove_dir_all(&dir);
}

/// The name a film's files are kept under moves when the film does: the length and the
/// modification time are in the hash, so a film replaced in place is copied again rather than
/// drawn with the subtitles of the file it used to be.
#[test]
fn the_name_of_a_films_folder_moves_when_the_film_does() {
    let dir = std::env::temp_dir().join(format!("subtitle-key-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
    let film = dir.join("episode.mkv");
    std::fs::write(&film, b"first").expect("a stand-in film is writable");
    let first = key(&film);

    std::fs::write(&film, b"a different length altogether").expect("the stand-in is rewritable");

    assert_ne!(
        first,
        key(&film),
        "a film whose bytes changed is a film whose subtitles are copied again"
    );

    let _ = std::fs::remove_dir_all(&dir);
}
