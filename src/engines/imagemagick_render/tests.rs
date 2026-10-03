use super::*;
use std::time::Instant;

/// The engine is never started for a name its own list does not hold: a picture this
/// app decodes itself, a document an engine of its own draws and a video are all
/// somebody else's, and asking an image converter about one would cost a launch and
/// answer nothing.
#[test]
fn asks_the_engine_only_about_the_names_it_reads() {
    for name in [
        "photo.png",
        "photo.jpg",
        "photo.heic",
        "texture.dds",
        "render.exr",
        "drawing.svg",
        "report.pdf",
        "notes.txt",
        "font.ttf",
        "clip.mp4",
        "drawing.cdr",
    ] {
        assert!(
            !crate::formats::magick_formats::is_engine_picture(Path::new(name)),
            "`{name}` is not one of its formats"
        );
    }

    for name in ["shot.nef", "shot.CR3", "shot.arw", "shot.dng", "scan.dcm"] {
        assert!(
            crate::formats::magick_formats::is_engine_picture(Path::new(name)),
            "`{name}` is one of its formats"
        );
    }
}

/// And a name it does not read is not queued either: nothing about a file is asked of
/// the engine that its list has not claimed.
#[test]
fn queues_nothing_for_a_name_it_does_not_read() {
    assert!(
        !WORKER.take_queued(),
        "the slot is empty before anything is asked of the engine"
    );

    request(Path::new("photo.png"), (1920, 1080), 7);

    assert!(
        !WORKER.take_queued(),
        "the engine is not asked about a name no list of its own holds"
    );
}

/// What the engine wrote is read for the size in its own header and for nothing else:
/// the signature every PNG opens with, and the width and height of the picture that
/// follows it. What is not a whole picture is not one to place a preview by — a
/// conversion that was cut short, an error message, a file of another format.
#[test]
fn reads_the_size_a_png_says_it_is() {
    let mut png = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
    png.extend_from_slice(&[0, 0, 0, 13]); // the length of the first chunk
    png.extend_from_slice(b"IHDR");
    png.extend_from_slice(&1600u32.to_be_bytes());
    png.extend_from_slice(&1200u32.to_be_bytes());

    assert_eq!(png_dimensions(&png), Some((1600, 1200)));

    // A conversion that wrote nothing, one that wrote something else, one that was cut
    // short inside its header, and a picture that claims no size at all.
    assert_eq!(png_dimensions(b""), None);
    assert_eq!(png_dimensions(b"%PDF-1.4"), None);
    assert_eq!(png_dimensions(&png[..18]), None);

    let mut empty = png.clone();
    empty[16..20].copy_from_slice(&0u32.to_be_bytes());
    assert_eq!(png_dimensions(&empty), None);

    // The size is what the picture says rather than what it was asked for: what the
    // engine writes is a picture fitted into the room, which is the size it is shown at
    // only when the file is larger than the display.
    let mut smaller = png.clone();
    smaller[16..20].copy_from_slice(&800u32.to_be_bytes());
    smaller[20..24].copy_from_slice(&600u32.to_be_bytes());
    assert_eq!(png_dimensions(&smaller), Some((800, 600)));
}

/// What the engine developed is held for the hover that asked and handed over once: it
/// is a picture waiting to be drawn, not a cache, and the second caller — a hover that
/// was replayed twice — is answered with nothing rather than with the same picture again.
#[test]
fn hands_the_picture_it_developed_over_once() {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-magick-tests")
        .join("holding");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let raw = folder.join("shot.nef");
    let other = folder.join("other.nef");
    std::fs::write(&raw, b"a raw, of a sort").expect("a written file");
    std::fs::write(&other, b"another").expect("a written file");

    let held = || Held {
        key: key_of(&raw),
        width: 1600,
        height: 1200,
        png: vec![0x89, b'P', b'N', b'G'],
    };

    hold(held());
    assert!(developed(&raw), "the picture is in hand");
    assert!(!developed(&other), "and it is not another file's");

    let taken = take_developed(&raw).expect("the picture");
    assert_eq!((taken.width, taken.height), (1600, 1200));
    assert!(!developed(&raw), "and there is one of it");
    assert!(
        take_developed(&raw).is_none(),
        "which cannot be taken twice"
    );

    // A picture developed for one file is left where it is by a caller asking about
    // another: what it is waiting for is its own hover.
    hold(held());
    assert!(take_developed(&other).is_none());
    assert!(
        developed(&raw),
        "so the picture it is holding is still there"
    );
    take_developed(&raw);

    let _ = std::fs::remove_dir_all(&folder);
}

/// What a file's size and its refusals are remembered by: the file and the version of it
/// that was read. A file saved again is a file to develop again — and one the engine
/// would not draw is asked about once, not once per hover.
#[test]
fn remembers_an_answer_for_the_version_of_the_file_it_was_read_from() {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-magick-tests")
        .join("remembering");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let raw = folder.join("shot.nef");
    std::fs::write(&raw, b"a raw, of a sort").expect("a written file");

    remember(&key_of(&raw), (1600, 1200));
    refuse(&key_of(&raw));
    assert_eq!(dimensions(&raw), Some((1600, 1200)));
    assert!(refused(&raw));

    assert_eq!(dimensions(&folder.join("other.nef")), None);
    assert!(!refused(&folder.join("other.nef")));

    // Saved again: what was known about the file it was is not what it is now.
    std::fs::write(&raw, b"a raw, edited and then some").expect("a written file");
    assert_eq!(dimensions(&raw), None);
    assert!(!refused(&raw));

    let _ = std::fs::remove_dir_all(&folder);
}

/// What a raw sample dump is, which is the one kind of file the engine cannot measure for
/// itself: the name says what one pixel weighs and nothing says how many there are, so the
/// shape comes out of the file's own length.
///
/// The arithmetic is checked on the lengths pictures are actually written at, and on the ones
/// that settle nothing: a length that is not a whole number of pixels, and a length no
/// proportion of the table divides exactly, are both answered with nothing rather than with a
/// guess — while a length whose proportion is a landscape photograph's is answered even where
/// it is also a square's, which is what the order of the table is for.
#[test]
fn works_a_raw_samples_shape_out_of_its_own_length() {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-magick-tests")
        .join("geometry");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let dump = |name: &str, bytes: u64| {
        let path = folder.join(name);
        // A length rather than its contents: what is measured is the file's size, and a
        // sparse file gets one without writing gigabytes of nothing to a disk.
        let file = std::fs::File::create(&path).expect("a written dump");
        file.set_len(bytes).expect("a dump of that length");
        path
    };

    // The lengths a photograph comes in: three bytes a pixel, four, and one.
    assert_eq!(
        raw_geometry(&dump("a.rgb", 640 * 480 * 3)),
        Some((640, 480))
    );
    assert_eq!(raw_geometry(&dump("b.gray", 800 * 600)), Some((800, 600)));
    assert_eq!(
        raw_geometry(&dump("c.rgba", 1920 * 1080 * 4)),
        Some((1920, 1080))
    );
    assert_eq!(
        raw_geometry(&dump("d.rgb", 2560 * 1440 * 3)),
        Some((2560, 1440))
    );
    assert_eq!(
        raw_geometry(&dump("e.rgb", 3840 * 2160 * 3)),
        Some((3840, 2160))
    );

    // A 16:9 picture is a square's worth of pixels — 1920x1080 is 1440 of them to a side —
    // and the picture is the answer rather than the square.
    let wide = raw_geometry(&dump("f.gray", 1920 * 1080)).expect("a shape");
    assert_eq!(wide, (1920, 1080));
    assert_ne!(wide.0, wide.1, "the longer side is the width");

    // A bi-level bitmap is bits rather than bytes, and a fax bitstream is read the same way.
    assert_eq!(
        raw_geometry(&dump("g.mono", 1024 * 768 / 8)),
        Some((1024, 768))
    );

    // A name that is not a dump has no shape to work out at all — a camera raw is a container
    // with a picture in it, which is a different thing entirely — and neither has a length
    // that is not a whole number of pixels, or one that no proportion of the table divides.
    assert!(
        !is_raw_sample(&dump("h.NEF", 1024)),
        "a camera raw is a container rather than a dump"
    );
    assert!(!is_raw_sample(&dump("j.png", 64)));
    assert!(
        is_raw_sample(&dump("i.RGB", 64)),
        "whatever case it is written in"
    );
    assert_eq!(raw_geometry(&dump("k.nef", 640 * 480 * 3 + 64)), None);
    assert_eq!(raw_geometry(&dump("l.rgb", 100)), None, "not a whole pixel");
    assert_eq!(raw_geometry(&dump("m.gray", 0)), None);
    assert_eq!(
        raw_geometry(&dump("n.rgb", 7)),
        None,
        "too small to be a picture"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// The integer square root the shape is worked out with: exact at the ends a file length
/// reaches, and never a unit off.
#[test]
fn takes_an_exact_square_root() {
    for (value, root) in [
        (0u64, 0u64),
        (1, 1),
        (3, 1),
        (4, 2),
        (8, 2),
        (9, 3),
        (10_000, 100),
        (10_001, 100),
        (4_294_967_296, 65_536),
    ] {
        assert_eq!(integer_sqrt(value), root, "the square root of {value}");
    }
}

/// What the engine ends is only what it started, and only once that conversion has had
/// its chance. A process that stays up stands in for the engine — a test is not going
/// to make ImageMagick spin on a file — recorded the way the engine is, by image name,
/// which is the check that keeps an id from being acted on by itself.
///
/// The bound is the one this engine gives itself rather than the one the give-up is
/// decided by (`supervisor`): what is being asked here is that the two ends of the
/// decision reach the process, and the number that decides is that module's business and
/// tested there.
#[test]
fn ends_a_conversion_only_once_it_has_outrun_the_give_up() {
    let _stand_in = crate::app::engine_processes::STAND_IN
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let _in_flight = crate::engines::supervisor::IN_FLIGHT_TAKEN
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let mut engine = crate::app::engine_processes::hidden_command("ping")
        .args(["-n", "30", "127.0.0.1"])
        .stdout(std::process::Stdio::null())
        .spawn()
        .expect("a process to stand in for the engine");
    let pid = engine.id();

    assert!(
        crate::app::engine_processes::processes_named("ping.exe").contains(&pid),
        "the stand-in runs the image the record below names it by"
    );
    crate::app::engine_processes::record("ping.exe", pid);

    // A conversion that has just started is a file being read, and is left to it.
    supervisor::begin(Adapter::ImageMagick, Path::new("shot.nef"), pid);
    supervisor::end_hung(Adapter::ImageMagick, CONVERSION_GIVE_UP);
    assert!(
        crate::app::engine_processes::is_running(pid),
        "a conversion that has just started is not an engine to end"
    );

    // One that has outrun the give-up is an engine that has stopped answering. A bound of
    // nothing is the shortest a run can be outrun by, which says the same thing about
    // the decision without waiting half a minute to say it.
    supervisor::end_hung(Adapter::ImageMagick, Duration::ZERO);
    assert!(
        !crate::app::engine_processes::is_running(pid),
        "the engine a conversion has outrun is ended"
    );

    supervisor::stop(Adapter::ImageMagick);
    let _ = engine.wait();
}

/// The engine's own program is found by name, and the older name only ever where
/// ImageMagick is installed under it: a `convert.exe` on the `PATH` is Windows' own
/// file system converter, and the one that is there on every machine must not be
/// mistaken for the engine. Whatever this machine has is what the expectations are
/// written from, and the rule is what is asserted.
#[test]
fn finds_the_engine_where_it_installs_and_never_a_system_tool() {
    let found = find_engine();

    if let Some(program) = &found {
        let name = program
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or_default();
        assert!(
            name.eq_ignore_ascii_case(ENGINE_IMAGE)
                || name.eq_ignore_ascii_case(LEGACY_ENGINE_IMAGE),
            "the engine runs under one of its own two names"
        );

        if name.eq_ignore_ascii_case(LEGACY_ENGINE_IMAGE) {
            let folder = program
                .parent()
                .and_then(|folder| folder.file_name())
                .and_then(|folder| folder.to_str())
                .unwrap_or_default();
            assert!(
                folder.starts_with("ImageMagick"),
                "the older name is only ever ImageMagick's own folder, not a system tool"
            );
        }
    }

    assert_eq!(
        available(),
        found.is_some(),
        "and whether an engine is installed is the answer this app goes by"
    );
}

/// What a conversion is asked for, measured against the installed engine: the picture
/// comes back as bytes with the size its header says, a file the engine cannot read
/// comes back as nothing at all, and neither leaves anything on the disk. Ignored because
/// it starts the installed ImageMagick, and run when the command is being looked at:
/// `cargo test -- --ignored --nocapture engine_conversion_probe`.
#[test]
#[ignore = "starts the installed ImageMagick"]
fn engine_conversion_probe() {
    let Some(program) = ENGINE.as_ref() else {
        println!("no ImageMagick installed: nothing to measure");
        return;
    };
    println!("engine: {}", program.display());

    let folder = std::env::temp_dir().join("rust-hover-preview-magick-probe");
    std::fs::create_dir_all(&folder).expect("a probe folder");

    // A picture to convert, made with the same engine: what is being measured is the
    // conversion rather than the format it is asked about, and every format is asked
    // about the same way.
    let sample = folder.join("probe-source.png");
    let made = crate::app::engine_processes::hidden_command(program)
        .args(["-size", "1200x800", "gradient:red-blue"])
        .arg(format!("png:{}", sample.display()))
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status();
    println!("a picture to convert: {made:?}");

    let started = Instant::now();
    let first = convert(program, &sample, (1920, 1080));
    println!(
        "first conversion: {:?} bytes in {:?}",
        first.as_ref().map(Vec::len),
        started.elapsed()
    );
    println!(
        "and its size, from its own header: {:?} — the picture fitted into the room",
        first.as_deref().and_then(png_dimensions)
    );

    let started = Instant::now();
    let again = convert(program, &sample, (600, 400));
    println!(
        "again, at a smaller room: {:?} bytes in {:?}, size {:?}",
        again.as_ref().map(Vec::len),
        started.elapsed(),
        again.as_deref().and_then(png_dimensions)
    );

    // And a file the engine cannot read: nothing is written and nothing is held.
    let broken = folder.join("probe.xcf");
    std::fs::write(&broken, b"not a picture at all").expect("a written file");
    println!(
        "a file it cannot read: {:?}",
        convert(program, &broken, (1920, 1080))
    );

    let _ = std::fs::remove_dir_all(&folder);
}
