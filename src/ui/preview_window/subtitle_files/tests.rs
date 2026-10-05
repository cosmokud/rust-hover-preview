use super::*;

/// Every track whose codec has a drawable small form is copied by the one pass, at the name the
/// reader looks for — `sub<i>.<ext>` under the film's folder, numbered the way `-map 0:s:<i>`
/// numbers — and a codec with no small form, or one the filter cannot draw, is left out rather
/// than failing the pass.
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
        extraction_args(folder, film, &codecs, &[]).expect("two copyable tracks make a command");

    assert_eq!(
        &args[..6],
        ["-y", "-v", "error", "-hide_banner", "-i", "film.mkv"],
        "the pass is quiet and the film is its input: {args:?}"
    );

    // `dvd_subtitle` has no small form this app draws and PGS is a form the filter refuses, so
    // neither has an output — while their indices still belong to the tracks that do, which is
    // the whole of why a name carries the track's own number rather than a count of what was
    // written.
    for (index, extension) in [(0usize, "ass"), (2, "srt")] {
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

    for index in [1usize, 3] {
        assert!(
            !args.iter().any(|argument| argument
                .contains(&format!("sub{index}."))
                || argument == &format!("0:s:{index}")),
            "the track with no drawable small form is left out rather than written as something it \
             is not — a PGS copy above all, which the filter refuses and takes the picture with it: \
             {args:?}"
        );
    }

    // A film none of whose tracks has a small form has no command at all: the pass would be the
    // whole read this exists to avoid and would produce nothing.
    assert_eq!(
        extraction_args(folder, film, &["dvd_subtitle".to_string()], &[]),
        None,
        "a read that produces nothing is the one thing the caller must never start"
    );
    assert_eq!(
        extraction_args(folder, film, &["hdmv_pgs_subtitle".to_string()], &[]),
        None,
        "and a PGS-only film is the same nothing: a copy of a track the filter refuses is a copy \
         that would take the film's picture down with it"
    );
}

/// A `sup` copy an older build left in a film's folder is not an answer — neither read back by
/// the resolver nor named by the filter — while the drawable copies beside it still are.
///
/// The name filter is the whole of the repair for a cache a build before this one poisoned: the
/// derived files are read by every probe after the one that wrote them, in memory and on disk, so
/// ignoring a stale answer here is what lets such a film be previewed again without the user
/// having to clear anything.
#[test]
fn a_stale_pgs_copy_is_not_read_back_as_an_answer() {
    let dir = std::env::temp_dir().join(format!("subtitle-sup-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
    let film = dir.join("episode.mkv");
    std::fs::write(&film, b"stand-in").expect("a stand-in film is writable");

    let folder = film_folder(&film);
    std::fs::create_dir_all(&folder).expect("the film's own folder is creatable");
    std::fs::write(folder.join("sub0.sup"), b"stand-in").expect("a stand-in copy is writable");

    assert!(
        resolve(&film, &["hdmv_pgs_subtitle".to_string()]).is_none(),
        "a folder holding only a copy the filter cannot draw holds nothing this app may name"
    );

    // And with a drawable copy beside it, the drawable one is the answer and the refused one is
    // not read at all.
    let ass = folder.join("sub1.ass");
    std::fs::write(&ass, b"[Script Info]\n").expect("a stand-in copy is writable");
    let derived = resolve(&film, &["hdmv_pgs_subtitle".to_string(), "ass".to_string()])
        .expect("a drawable copy means an answer");

    assert_eq!(
        derived.track(0),
        None,
        "the PGS track's slot stays a hole, whatever a folder from an older build holds"
    );
    assert_eq!(
        derived.track(1),
        Some(ass.as_path()),
        "while its neighbour is answered under its own number"
    );

    let _ = std::fs::remove_dir_all(&dir);
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

/// One extraction slot: a film already being copied is not copied again, another film's pass is
/// dropped by the claim that displaces it, and only the roads out of a preview may keep one
/// running.
///
/// This is the single-flight rule the whole of `Slot` exists for, asserted where it can be
/// asserted without a film to read: the state machine itself, on a slot of the test's own. The
/// processes are what a claim leads to, and what is checked here is what every later claim and
/// drop is built on — that a second film's request ends the first's, that a window that has gone
/// ends whatever it was showing, and that an ending task cannot clear a claim that is no longer
/// its own.
#[test]
fn one_extraction_at_a_time_is_the_slot_the_newest_window_takes() {
    let first = Path::new("hl-slot-first.mkv");
    let second = Path::new("hl-slot-second.mkv");
    let mut slot = Slot::default();

    // Nothing is running to begin with, so both a keep for a film and a keep for no film are
    // nothing to drop.
    slot.keep(None);
    slot.keep(Some(first));

    let (one, displaced) = slot.claim(first).expect("the empty slot is claimed");
    assert!(displaced.is_none(), "nothing was running to displace");
    assert!(
        slot.claim(first).is_none(),
        "a film already being copied is not copied again"
    );

    let (two, displaced) = slot.claim(second).expect("another film takes the slot");
    let displaced = displaced.expect("the task it displaced is answered with it");
    assert!(
        displaced.dropped.load(Ordering::Acquire),
        "and the displaced task is dropped: the film on screen is the one worth reading"
    );
    assert!(!two.dropped.load(Ordering::Acquire));

    // A keep for the film in the slot leaves it alone; a keep for another film drops it — the
    // roads out of a preview are a hover that ended and a pin that came down or was shown
    // another file.
    slot.keep(Some(second));
    assert!(
        !two.dropped.load(Ordering::Acquire),
        "the window showing the film keeps its copy running"
    );
    slot.keep(Some(first));
    assert!(
        two.dropped.load(Ordering::Acquire),
        "a window showing another film drops it"
    );

    let (three, displaced) = slot
        .claim(first)
        .expect("a dropped task does not hold the slot");
    assert!(
        displaced.is_some_and(|displaced| Arc::ptr_eq(&displaced, &two)),
        "and the dropped task is what this one displaces"
    );

    slot.keep(None);
    assert!(
        three.dropped.load(Ordering::Acquire),
        "a preview that has gone asks for no copy at all"
    );

    // An ending task gives up the slot only where it is still its own: the first task ended long
    // ago, and its release must not clear the claim of the task that displaced it.
    let (four, _) = slot.claim(second).expect("the slot is free again");
    slot.release(&one);
    assert!(
        slot.claim(second).is_none(),
        "the ending first task left the fourth task's claim standing"
    );
    slot.release(&two);
    assert!(
        slot.claim(second).is_none(),
        "and the displaced second task left it standing as well"
    );
    slot.release(&four);
    assert!(
        slot.claim(second).is_some(),
        "while the task that actually owns it gives it up"
    );
}

/// A pass that failed or was dropped leaves no folder behind: a half-written answer is one a later
/// probe would read and trust, which is worse than no answer at all.
#[test]
fn a_discarded_film_leaves_no_folder_for_a_later_probe_to_trust() {
    let dir = std::env::temp_dir().join(format!("subtitle-discard-{}", std::process::id()));
    std::fs::create_dir_all(&dir).expect("the scratch folder for this test is creatable");
    let film = dir.join("episode.mkv");
    std::fs::write(&film, b"stand-in").expect("a stand-in film is writable");

    let folder = film_folder(&film);
    std::fs::create_dir_all(folder.join(FONTS_FOLDER)).expect("the copy folder is creatable");
    std::fs::write(folder.join("sub0.ass"), b"[Script Info]\n")
        .expect("a stand-in copy is writable");

    discard(&film);

    assert!(
        !folder.exists(),
        "what a failed or dropped pass wrote goes with it: a later probe reads the folder as an \
         answer, and half a copy is not one"
    );

    let _ = std::fs::remove_dir_all(&dir);
}
