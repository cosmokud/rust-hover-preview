use super::*;

/// A folder of a test's own with three pictures in it and nothing else, so that a walk of
/// it is the three and the three in the order they are named: the name order is what a
/// folder nothing has been hovered in is walked in, and a fresh folder of a test's own is
/// exactly that (see `pin_navigation`).
fn walkable_folder(name: &str) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-pin-walk-tests")
        .join(name);
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a folder a test can write to");

    for picture in ["a.png", "b.png", "c.png"] {
        std::fs::write(folder.join(picture), vec![0u8; 8]).expect("a file a test can write");
    }

    folder
}

/// The walk is read under the app's own configuration, which is loaded from `config.ini` —
/// so a walk test says what it wants walked rather than taking whatever the machine it runs
/// on has set. Only that switch is touched, and it is put back when the guard is dropped,
/// so a test running beside this one is not answered with this one's setting.
struct WalkEveryFile(PinNavFileTypes);

impl WalkEveryFile {
    fn set() -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = config.pin_nav_file_types;
        config.pin_nav_file_types = PinNavFileTypes::All;

        WalkEveryFile(was)
    }
}

impl Drop for WalkEveryFile {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_nav_file_types = self.0;
        }
    }
}

/// a walk bounded by the list it came from, standing on the file it last /// landed on.
#[test]
fn a_walk_offers_every_other_file_of_the_folder_once_and_then_ends() {
    let _every_file = WalkEveryFile::set();
    let folder = walkable_folder("bounded");
    let list = vec![
        folder.join("a.png"),
        folder.join("b.png"),
        folder.join("c.png"),
    ];

    let mut walk = PinStep {
        at: list[0].clone(),
        from: list[0].clone(),
        step: 1,
        // The bound a walk of this folder is given by the button that started it, so that
        // what is tested here is the bound the caption actually walks under.
        left: walk_budget(&list),
        // The list is what the planner read off the preview thread, handed in rather
        // than looked up again on every step (see `PinStep`).
        list: list.clone(),
        shuffle: false,
        sounds: Vec::new(),
    };

    // Every step, and the one that ends the walk: the end is part of what is being asked
    // about, so it is a step like any other and is recorded as one.
    let mut offered = Vec::new();
    loop {
        let next = walk.step();
        let ends = next.is_none();
        offered.push(next);
        if ends {
            break;
        }
    }

    assert_eq!(
        offered,
        vec![Some(list[1].clone()), Some(list[2].clone()), None],
        "every file of the folder but the one the pin is showing is offered, and then the \
             walk ends rather than going round for ever"
    );
    assert_eq!(
        walk.at, list[2],
        "the walk stands on the file it last landed on, which is where the next step is \
             taken from rather than from the file the pin is still showing"
    );
}

/// a file the pin cannot be shown is not where a walk stops, and the walk /// comes back rather than being spent.
#[test]
fn stepping_over_a_file_the_pin_cannot_show_carries_the_walk_on() {
    let _every_file = WalkEveryFile::set();
    let folder = walkable_folder("carried-on");

    let mut held: Option<PinStep> = None;
    let walk = PinStep {
        at: folder.join("a.png"),
        from: folder.join("a.png"),
        step: 1,
        left: 1,
        list: vec![folder.join("a.png"), folder.join("b.png")],
        shuffle: false,
        sounds: Vec::new(),
    };

    pin_step_off(Some(walk), &mut held);

    let walk = held
        .clone()
        .expect("the walk is handed back rather than spent");
    assert_eq!(
        walk.at,
        folder.join("b.png"),
        "standing on the next file, which is the one the walk carries on to"
    );

    // And a walk that has nothing left is not held: a folder of files this app cannot show
    // is a walk that ends rather than one that goes for ever, and nothing is queued in its
    // place — which is what the mark for the file it stopped on is for.
    let mut spent: Option<PinStep> = None;
    pin_step_off(Some(walk), &mut spent);
    assert!(
        spent.is_none(),
        "the walk that has run out of files is not held, so nothing is asked for again"
    );

    // And so is a failure that had no walk to begin with: a file a pick in the listing
    // failed at, which is the same case with nothing to step from. Nothing is queued, and
    // what a pin is left standing over is the mark rather than the file it came from — the
    // file before a pin is on screen is not somewhere the walk goes *back* to (see
    // `show_pin_failure`).
    let mut none_left: Option<PinStep> = None;
    pin_step_off(None, &mut none_left);
    assert!(
        none_left.is_none(),
        "a failure with no walk under it queues nothing: navigation moves forward only, so a \
             folder with one broken file in it can be walked past with Next"
    );
}

/// a file the pin could not draw is shown as the mark rather than as the file it /// came from, and the mark is a live kind.
#[test]
fn a_failed_file_with_nowhere_to_step_to_is_left_standing_as_the_mark() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let broken = PathBuf::from("C:\\folder\\broken.mp4");
    let mut pin = overlay_pin((0, 0, 320, 240), PinChrome::always());
    pin.path = broken.clone();
    stand_pin(Some(pin));

    // The mark itself, drawn against a palette of its own rather than the one this run
    // happens to be reading: a cross over an empty buffer would pass the liveness read below
    // and show nothing at all, which is the one way this could be wrong with every other
    // assertion here still green.
    let palette = pin_chrome::ChromePalette {
        background: [10, 20, 30],
        foreground: [240, 240, 240],
        accent: [0, 0, 0],
        dark: true,
    };
    let mut pixels = vec![0u8; 128 * 128 * 4];
    pin_chrome::paint_failure_mark(&mut pixels, 128, 128, &palette);

    assert_eq!(
        pixels.len(),
        128 * 128 * 4,
        "a 128 by 128 mark is a whole bitmap and not a quarter of one"
    );
    assert!(
        pixels.chunks(4).all(|pixel| pixel[3] == 255),
        "every pixel of it is written, so it is a surface the pin's own compositor may copy \
             rather than blend, and the tray's backdrop is never seen through it"
    );
    let panel = [30u8, 20, 10, 255];
    let ink = [240u8, 240, 240, 255];
    let painted = pixels.chunks(4).filter(|pixel| **pixel == ink).count();
    assert!(
        painted > 0 && painted < pixels.len() / 4,
        "and something is drawn on the panel in the theme's ink: a cross fills part of it \
             rather than all of it, and a flat panel says nothing at all about the file that \
             failed"
    );
    assert!(
        pixels
            .chunks(4)
            .all(|pixel| *pixel == panel || *pixel == ink),
        "and the two of them are the only colours in it, so the cross has an edge rather \
             than a blend into whatever was behind it"
    );
    // Two strokes crossing, which is the whole of what the mark is. A row through the middle
    // carries both strokes and so does a row well above and below it: a bar or a ring would
    // carry one at the middle and none at the edges, and a single diagonal would carry one at
    // the top left and none at the bottom right.
    let row = |y: usize| {
        pixels
            .chunks(4)
            .skip(y * 128)
            .take(128)
            .filter(|pixel| **pixel == ink)
            .count()
    };
    assert!(
        row(64) >= 8,
        "the middle of the mark is where the strokes cross"
    );
    assert!(
        row(32) >= 8 && row(96) >= 8,
        "and the strokes reach out towards all four corners of the panel, which is what says \
             *could not be shown* rather than *empty*"
    );

    // The media that mark becomes: a square of its own, and not the shape of a file whose
    // header survived.
    let media = unplayable_media((PIN_FAILURE_SIDE, PIN_FAILURE_SIDE));
    assert_eq!(
        (media.current_width(), media.current_height()),
        (PIN_FAILURE_SIDE, PIN_FAILURE_SIDE),
        "the mark is the shape the pin is given, so installing it never depends on measuring \
             a file whose measurement is what cannot be trusted"
    );

    // A mark on screen is a frame this app holds, so the liveness read — which is what
    // closed a pin put straight onto a corrupted file — has to find it alive with nothing
    // queued behind it.
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }
    assert!(
        pin_media_is_alive(false),
        "a window with the mark in it is a window with something in it: a file that fails and \
             has nowhere to step to is shown, not closed"
    );

    // Which is a property of the kind and not of the media's pixels: an engine kind, a film
    // and a sound each answer the question in their own terms, and a mark is not one of those
    // — nothing outside this thread can take a cross away.
    assert!(
        pin_media_failed_before_a_frame(None).is_none(),
        "and the mark is not a failure a second time, so it is not stepped over or re-marked \
             for ever"
    );

    // And the whole of how it gets there: a load that has already answered, so that nothing
    // is waited for between the tick that notices the failure and the tick that shows it.
    let mut load: Option<PinLoad> = None;
    assert!(
        show_pin_failure(&broken, &mut load),
        "a pin that is up is given the mark"
    );
    let answer =
        take_pin_load(&mut load).expect("the mark's load is answered before it is handed over");
    assert_eq!(
        answer.path, broken,
        "for the file that failed, and not for any other"
    );
    assert!(
        answer.walk.is_none(),
        "and it is not a step of a walk, there being none to step"
    );
    let answer = answer.media.expect("a mark is always a frame");
    assert_eq!(
        answer.media_type,
        MediaType::Unplayable,
        "of the kind that reads as alive"
    );
    assert!(
        !answer.is_streaming(),
        "and it is not something still being decoded: a cross this app drew is not queued"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A folder of a test's own with sounds and pictures in it, so that a
/// shuffled walk has sounds to pick from and files that are not sounds
/// beside them: which extension is a sound is the `[audio]` list of
/// `config.ini` (see `lists::AUDIO`), and a folder of a test's own is
/// one nothing has hovered, so the walk of it is the name order.
fn sounded_folder(name: &str) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-pin-walk-tests")
        .join(name);
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a folder a test can write to");

    for sound in ["a.mp3", "b.mp3", "c.mp3"] {
        std::fs::write(folder.join(sound), vec![0u8; 8]).expect("a file a test can write");
    }
    for picture in ["a.png", "b.png"] {
        std::fs::write(folder.join(picture), vec![0u8; 8]).expect("a file a test can write");
    }

    folder
}

/// The pin's own shuffle switch, set for a test and put back when the
/// guard is dropped, for the same reason `WalkEveryFile` is: the walk is
/// read under the app's own configuration, so a test says what it wants
/// walked rather than taking whatever the machine it runs on has set.
struct ShuffleThePin(bool);

impl ShuffleThePin {
    fn set() -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = config.pin_mode_audio_shuffle;
        config.pin_mode_audio_shuffle = true;

        ShuffleThePin(was)
    }
}

impl Drop for ShuffleThePin {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_mode_audio_shuffle = self.0;
        }
    }
}

/// A shuffled step lands on a sound of the folder other than the one the
/// pin is showing, and on nothing else: a shuffle is a walk through the
/// folder's sounds in an order of their own choosing, so every roll names
/// one of the sounds, and the file on screen is not a file a step is for
/// (see `walk_budget`).
#[test]
fn a_shuffled_step_lands_on_a_sound_other_than_the_one_pinned() {
    let folder = sounded_folder("shuffle-pick");
    let sounds = vec![
        folder.join("a.mp3"),
        folder.join("b.mp3"),
        folder.join("c.mp3"),
    ];
    let pinned = folder.join("b.mp3");

    for roll in 0..64 {
        let landed = shuffle_step(&pinned, &sounds, roll)
            .expect("a folder of sounds has one to land on");

        assert!(
            sounds.contains(&landed),
            "a shuffled step lands on a sound of the folder: {landed:?}"
        );
        assert_ne!(
            landed, pinned,
            "the file the pin is showing is not a file a step is for"
        );
    }

    // Every sound of the folder but the pinned one is one some roll names,
    // which is what makes the step a shuffle rather than a walk that
    // skips files.
    let mut landed_on = Vec::new();
    for roll in 0..64 {
        landed_on.push(shuffle_step(&pinned, &sounds, roll).expect("a sound"));
    }
    for sound in sounds.iter().filter(|sound| *sound != &pinned) {
        assert!(
            landed_on.contains(sound),
            "every sound but the pinned one is one a roll lands on: {sound:?}"
        );
    }
}

/// A folder whose only sound is the one the pin is showing is a folder a
/// shuffled step has nowhere to go in: the file on screen is not a file a
/// step is for, and there is no other sound to step onto, so the walk ends
/// rather than stepping onto a file that is not a sound — the one answer a
/// shuffle has that keeps it inside the `[audio]` category.
#[test]
fn a_folder_whose_only_sound_is_the_pinned_one_has_nowhere_to_shuffle() {
    let folder = sounded_folder("shuffle-alone");
    let only = folder.join("only.mp3");
    let sounds = vec![only.clone()];

    assert_eq!(
        shuffle_step(&only, &sounds, 0),
        None,
        "the one sound of the folder is the one on screen"
    );
    assert_eq!(
        shuffle_step(&only, &[], 0),
        None,
        "and a folder of no sounds at all is the same answer"
    );
}

/// The walk a shuffled pin's planner answers carries the folder's sounds
/// and nothing else: the list is every file the walk steps through, and
/// the sounds are the `[audio]` category of it — the files whose extension
/// is in the `[audio]` list of `config.ini` — so a shuffled step is a
/// random one of the sounds and never a file that is not a sound.
#[test]
fn a_walk_that_shuffles_carries_the_sounds_of_the_folder() {
    let _every_file = WalkEveryFile::set();
    let _shuffle = ShuffleThePin::set();
    let folder = sounded_folder("planner-shuffle");

    let answer = answer_pin_job(&PinJob::Walk {
        at: folder.join("b.mp3"),
        step: 1,
        config: Box::new(CONFIG.lock().expect("the configuration").clone()),
    });
    let Some(PinPlanned::Walk(walk)) = answer else {
        panic!("a walk is always answered");
    };

    assert!(walk.shuffle, "the walk shuffles when the pin's shuffle is on");
    assert_eq!(
        walk.sounds,
        vec![
            folder.join("a.mp3"),
            folder.join("b.mp3"),
            folder.join("c.mp3"),
        ],
        "the sounds are the folder's `[audio]` files, and no file of another kind is one"
    );

    // And a walk that does not shuffle carries no sounds to pick from,
    // which is the walk every step of a pin's was before the setting.
    let mut config = CONFIG.lock().expect("the configuration");
    config.pin_mode_audio_shuffle = false;
    let answer = answer_pin_job(&PinJob::Walk {
        at: folder.join("b.mp3"),
        step: 1,
        config: Box::new(config.clone()),
    });
    let Some(PinPlanned::Walk(walk)) = answer else {
        panic!("a walk is always answered");
    };
    assert!(!walk.shuffle, "and the walk does not shuffle with the shuffle off");
    assert!(
        walk.sounds.is_empty(),
        "a walk that steps to the file beside it has no sounds to pick from"
    );
}

/// the mark behind a bubble, and behind a pin that has come apart for some other /// reason, is still a kind this app draws.
#[test]
fn the_mark_is_drawn_wherever_a_frame_is() {
    assert!(
        MediaType::Unplayable.has_bubble_picture(),
        "a bubble collapsed over a file that failed shows the failure, not a mark standing in \
             for a picture that is not there"
    );
    assert_eq!(
        MediaType::Unplayable.kind(),
        None,
        "and it is behind no preview gate: a window already standing over a file it could \
             not draw does not become hideable the moment the user hides films"
    );
    assert_eq!(
        pin_frame(Some(MediaType::Unplayable)),
        PinFrame::Shaped,
        "a square of its own is a shape a window scales both sides of"
    );
    assert!(
        !pin_transport_kind(Some(MediaType::Unplayable)),
        "and there is nothing playing behind it, so there is no bar to carry a playhead"
    );
}

/// the short-circuit that keeps a pin up while it is being shown another file, over every /// kind there is — including the new one.
///
/// It is asked here for every kind rather than for the one it was written for because the
/// kinds it is a short-circuit *over* are the whole of it: `NativeVideo`, `Video` and the
/// two the engine draws all answer "gone" on this machine, with no player running and no
/// browser standing, and a kind that answers "gone" wrongly is the defect this rule exists
/// to prevent. Every other kind answers "there" either way, so listing them is what makes
/// the four above it a real answer rather than an accident of where they are written.
#[test]
fn a_pin_being_shown_another_file_is_alive_whatever_is_on_screen() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((0, 0, 80, 60), PinChrome::always());
    pin.path = PathBuf::from("C:\\folder\\whatever.mp4");
    stand_pin(Some(pin));

    for media_type in [
        MediaType::StaticImage,
        MediaType::Dds,
        MediaType::EngineSvg,
        MediaType::EngineFont,
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
        MediaType::AnimatedHeif,
        MediaType::AnimatedJxl,
        MediaType::Video,
        MediaType::NativeVideo,
        MediaType::Audio,
        MediaType::Pdf,
        MediaType::Text,
        MediaType::Archive,
        MediaType::Peazip,
        MediaType::Office,
        MediaType::Design,
        MediaType::Vector,
        MediaType::Libre,
        MediaType::Magick,
        MediaType::Calibre,
        MediaType::Comic,
        MediaType::Unplayable,
        MediaType::Loading,
    ] {
        let mut media = create_loading_media(8, 8);
        media.media_type = media_type;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        assert!(
            pin_media_is_alive(true),
            "{media_type:?} with a next file already in hand is a pin with something to be a \
                 window onto, whatever is behind it right now"
        );
    }

    // And the four that make the rule worth anything: asked without a next file in hand,
    // each of them says the player or the engine behind it has gone, which is what the
    // short-circuit is there to overrule.
    for media_type in [
        MediaType::EngineSvg,
        MediaType::EngineFont,
        MediaType::Video,
        MediaType::NativeVideo,
    ] {
        let mut media = create_loading_media(8, 8);
        media.media_type = media_type;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        assert!(
            !pin_media_is_alive(false),
            "{media_type:?} with nothing queued and no player or engine behind it really has \
                 come apart: the answer above is the short-circuit and not the kind"
        );
    }

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

#[test]
fn a_pin_wait_puts_its_arc_up_only_once_it_has_outlasted_the_delay() {
    let mut load = PinLoad {
        path: PathBuf::from("C:\\wait\\slow.png"),
        update: PinUpdate {
            content: (0, 0, 100, 100),
            dpi: 96,
            volume: 50,
        },
        arc: PinArc {
            started: Instant::now(),
            spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
            turned: None,
        },
        answer: channel().1,
        walk: None,
    };

    assert!(
        !load.arc.due(),
        "a wait that has not run its delay shows no arc"
    );

    load.arc.started = Instant::now() - Duration::from_millis(DEFAULT_SPINNER_DELAY_MS + 1);
    assert!(load.arc.due(), "and a wait that has outlasted it shows one");

    load.arc.spun();
    assert!(
        !load.arc.due(),
        "the turn is a cadence away from the last one rather than due at once, so the \
             moment the delay runs out is one paint and not one paint a tick"
    );
    assert!(
        load.arc.turned.is_some(),
        "and a wait that has had its arc is a wait that does not ask for a second one"
    );

    load.arc.turned = Some(Instant::now() - Duration::from_millis(u64::from(SPINNER_TURN_MS) + 1));
    assert!(
        load.arc.due(),
        "until the spinner's own cadence has come round again"
    );
}

/// the wait a walk is asked under is painted on the pin's own terms — /// nothing before the delay, then the arc at the spinner's cadence — because the folder /// read is on a thread of its own now and a walk is only slow f
#[test]
fn a_walk_wait_is_answered_on_the_same_terms_a_load_is() {
    let mut wait = PinWait {
        started: Instant::now(),
        spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
        turned: None,
    };

    assert!(
        !wait.due(),
        "a folder that reads inside the delay shows no arc at all"
    );

    wait.started = Instant::now() - Duration::from_millis(DEFAULT_SPINNER_DELAY_MS + 1);
    assert!(wait.due(), "and a read that outlasts it shows one");

    wait.spun();
    assert!(
        !wait.due(),
        "and the turn is a cadence away from the last rather than due at once, so a \
             long folder read does not repaint the window as fast as the loop runs"
    );
}

/// the walk a caption button asks for is asked of the planner, and the /// planner answers with a walk that already holds the folder's list.
#[test]
fn a_walk_the_planner_answers_carries_the_list_it_walked() {
    let folder = walkable_folder("planner-walked");
    let list = vec![folder.join("a.png"), folder.join("b.png")];

    let mut walk = PinStep {
        at: list[0].clone(),
        from: list[0].clone(),
        step: 1,
        left: walk_budget(&list),
        list: list.clone(),
        shuffle: false,
        sounds: Vec::new(),
    };

    assert_eq!(
        walk.step(),
        Some(list[1].clone()),
        "the walk steps through the list it was handed, with no second read of the folder"
    );
    assert_eq!(
        walk.from(),
        list[0],
        "and it still says which file it was asked from after it has moved on, which is \
             what a late answer is matched against"
    );
    assert_eq!(
        walk.left, 0,
        "and the bound came with it, so a folder of files this app cannot show is a walk \
             that ends rather than one that goes round for ever"
    );
    assert!(
        walk.step().is_none(),
        "which is what ends it: nothing further is asked of the planner for this walk"
    );
}

/// the Shell is asked once per *format*, and a name with no extension /// is keyed by itself rather than filed under no format at all.
#[test]
fn the_shell_is_asked_once_per_format_and_a_dot_file_is_its_own() {
    assert_eq!(shell_format(Path::new("C:\\a\\b.PNG")), ".png");
    assert_eq!(shell_format(Path::new("C:\\a\\b.png")), ".png");
    assert_eq!(shell_format(Path::new("C:\\a\\.gitignore")), ".gitignore");
    assert_eq!(shell_format(Path::new("C:\\a\\LICENSE")), "license");

    assert_ne!(
        shell_format(Path::new("C:\\a\\.gitignore")),
        shell_format(Path::new("C:\\a\\Makefile")),
        "a name with no extension is its own format rather than a shared empty one, so \
             two such files cannot trade associations"
    );
}

/// a walk is *always* answered, so a press that had nowhere to go ends /// the wait at once rather than leaving the pin's arc turning out a bound.
#[test]
fn a_walk_with_nowhere_to_go_is_answered_rather_than_left_outstanding() {
    let _every_file = WalkEveryFile::set();
    let config = CONFIG
        .lock()
        .expect("the configuration is not poisoned")
        .clone();

    // The case the planner most often finds, and the least dramatic: a folder holding
    // nothing but the file the pin is standing on. `walkable_folder` makes three, so
    // this is a folder of its own with one.
    let folder = walkable_folder("nowhere-to-go-alone");
    for picture in ["b.png", "c.png"] {
        let _ = std::fs::remove_file(folder.join(picture));
    }

    let answered = answer_pin_job(&PinJob::Walk {
        at: folder.join("a.png"),
        step: 1,
        config: Box::new(config.clone()),
    });
    let Some(PinPlanned::Walk(walk)) = answered else {
        panic!("a walk is always answered, so the pin's arc can be taken down");
    };
    assert_eq!(
        walk.left, 0,
        "and a walk with nothing to step to is a spent one, which is how the loop is told \
             the press is over rather than still out"
    );
    assert_eq!(
        walk.at,
        folder.join("a.png"),
        "standing on the file it was asked from"
    );
    let _ = std::fs::remove_dir_all(&folder);

    // And a folder this build cannot read at all is the same answer rather than silence.
    let gone = answer_pin_job(&PinJob::Walk {
        at: PathBuf::from("C:\\no-such-folder-here\\a.png"),
        step: 1,
        config: Box::new(config),
    });
    let Some(PinPlanned::Walk(walk)) = gone else {
        panic!("a folder that cannot be read is answered like a folder with nothing in it");
    };
    assert_eq!(
        walk.left, 0,
        "and it ends the wait on the tick it lands, rather than the pin waiting out \
             PIN_JOB_GIVEUP for a question that was answered the moment it was asked"
    );
}

/// the arc is put over a band that is already there, in the middle of it, /// and nowhere else — a band cleared for a spinner would be a window with nothing in it for /// the length of a decode, which is the freeze the arc
#[test]
fn the_pin_arc_leaves_the_band_it_is_drawn_over_except_where_it_is() {
    let (width, height) = (240u32, 200u32);
    // A band as a file leaves it: an even, opaque surface, so a pixel that has changed is a
    // pixel the arc can only have changed.
    let band = |alpha: u8| {
        let mut surface = vec![40u8; (width as usize) * (height as usize) * 4];
        for pixel in surface.as_chunks_mut::<4>().0 {
            pixel[3] = alpha;
        }
        surface
    };
    let changed = |surface: &[u8], untouched: &[u8]| {
        surface
            .as_chunks::<4>()
            .0
            .iter()
            .zip(untouched.as_chunks::<4>().0.iter())
            .filter(|(drawn, was)| drawn != was)
            .count()
    };

    let untouched = band(255);
    pin_arc_set(Some(Duration::from_millis(500)));

    let mut out = untouched.clone();
    paint_pin_spinner(&mut out, width, 0, height as i32);

    let drawn = changed(&out, &untouched);
    assert!(
        drawn > 0,
        "a window waiting for a file is painted with the arc, in the middle of its band"
    );

    // Every pixel of it is inside the arc's own ring, and every pixel of the ring is
    // inside the box the drawing walks: the band outside that box is the file the pin is
    // showing, untouched right to the edges of the window.
    let outside = untouched
        .as_chunks::<4>()
        .0
        .iter()
        .zip(out.as_chunks::<4>().0.iter())
        .enumerate()
        .filter(|(_, (was, drawn))| was != drawn)
        .all(|(index, _)| {
            let (x, y) = (index as i32 % width as i32, index as i32 / width as i32);
            (x - width as i32 / 2).abs() <= PIN_ARC_REACH
                && (y - height as i32 / 2).abs() <= PIN_ARC_REACH
        });
    assert!(
        outside,
        "and it is drawn nowhere else: the rest of the band is the file the pin is showing"
    );

    // A band with nothing behind it — the hole a player's window of its own stands in —
    // takes the arc as premultiplied coverage over nothing, which is what a layered
    // window's surface is read as.
    let clear = band(0);
    let mut over_nothing = clear.clone();
    paint_pin_spinner(&mut over_nothing, width, 0, height as i32);
    assert!(
        changed(&over_nothing, &clear) > 0,
        "an arc over an empty band is still an arc: the window says what it is waiting for"
    );

    // And a band with no room for one is left alone, rather than drawn into a half of a
    // ring that would read as a mark rather than a wait.
    let narrow = band(255);
    let mut cramped = narrow.clone();
    paint_pin_spinner(&mut cramped, width, 0, PIN_ARC_REACH * 2 - 1);
    assert_eq!(
        cramped, narrow,
        "and a band with no room for the arc keeps the file it is showing"
    );

    pin_arc_set(None);
    let mut after = untouched.clone();
    paint_pin_spinner(&mut after, width, 0, height as i32);
    assert_eq!(
        after, untouched,
        "and a window that is waiting for nothing is painted exactly as it was"
    );
}
