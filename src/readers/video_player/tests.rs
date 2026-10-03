use super::*;

/// The name a media source is read by is the Shell's path with the verbatim prefix
/// taken off, which is the whole of what the engine was failing on: a verbatim path
/// is not a URL, and the engine resolves this name as one.
#[test]
fn reads_the_plain_form_of_a_verbatim_path() {
    assert_eq!(
        plain_path(Path::new(r"\\?\C:\Music\track.mp3")),
        r"C:\Music\track.mp3",
        "the verbatim form of a local path is the path it is written around"
    );
    assert_eq!(
        plain_path(Path::new(r"\\?\UNC\server\share\track.mp3")),
        r"\\server\share\track.mp3",
        "and the verbatim form of a share keeps its server"
    );
    assert_eq!(
        plain_path(Path::new(r"C:\Music\track.mp3")),
        r"C:\Music\track.mp3",
        "a path that was never verbatim is left exactly as it is"
    );
    assert_eq!(
        plain_path(Path::new(r"\\server\share\track.mp3")),
        r"\\server\share\track.mp3",
        "and so is a share written the ordinary way"
    );
}

/// A file the engine has been watched failing at is a file FFmpeg's player plays from then
/// on: the mark is written where the probe's own answer is kept and read by the same
/// question, so what the routing asks about the file is the answer this side learned by
/// watching it rather than the one it guessed at.
#[test]
fn a_file_the_engine_failed_at_is_ffmpegs_from_here_on() {
    // A file of this module's own holding nothing a decoder could read, so that the answer
    // the mark corrects is one about a file no engine can play rather than about whatever
    // this machine happens to have decoders for.
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-video-tests")
        .join("unplayable");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join("not-a-film.mp4");
    std::fs::write(&path, b"not a film at all").expect("the file");

    mark_unplayable(&path);

    assert!(
        !plays(&path),
        "the mark is what the router is answered with, not the probe"
    );
}

/// The pixel at a row and column of a frame of the box's own width.
fn pixel_at(bgra: &[u8], width: usize, row: usize, column: usize) -> [u8; 4] {
    let at = (row * width + column) * 4;

    bgra[at..at + 4].try_into().expect("four bytes to a pixel")
}

/// A picture read above its own size is read *between* its pixels, which is the whole of
/// what this side scales a preview above 100% for: the edge of a picture that was two pixels
/// wide arrives as the four steps between them rather than as the two it was made of.
#[test]
fn an_enlarged_edge_is_read_between_its_pixels() {
    // A 2x2 picture of one edge: black down the left of it and blue down the right.
    let source = [
        0u8, 0, 0, 255, 255, 0, 0, 255, //
        0, 0, 0, 255, 255, 0, 0, 255,
    ];
    let mut out = vec![0u8; 4 * 4 * 4];
    let mut rows = [Vec::new(), Vec::new()];

    assert!(scale_rows(&source, 8, (2, 2), (4, 4), &mut rows, &mut out));

    // The outermost pixels stand on the two the picture has, and each of the two between them
    // is a quarter of the way from one to the other — where a picture read a source pixel at a
    // time would have arrived as `0, 0, 255, 255`.
    let expected = [0u8, 64, 191, 255];

    for row in 0..4 {
        for (column, level) in expected.iter().enumerate() {
            assert_eq!(
                pixel_at(&out, 4, row, column),
                [*level, 0, 0, 255],
                "row {row}, column {column} of the edge, and every pixel of it opaque"
            );
        }
    }
}

/// The other direction is read the same way — a picture read below its own size lands between
/// four of its pixels and is answered with their average — which is what a file whose pixels
/// are not square asks of the one axis it is stretched along at every size.
#[test]
fn a_shrunken_picture_is_read_between_its_pixels_too() {
    // Black above blue, read into the one pixel between them.
    let source = [
        0u8, 0, 0, 255, 0, 0, 0, 255, //
        255, 0, 0, 255, 255, 0, 0, 255,
    ];
    let mut out = vec![0u8; 4];
    let mut rows = [Vec::new(), Vec::new()];

    assert!(scale_rows(&source, 8, (2, 2), (1, 1), &mut rows, &mut out));

    assert_eq!(out, [128, 0, 0, 255], "the pixel standing between them");
}

/// Who scales the picture is one question and these are its two sizes: a box meaningfully larger
/// than the picture on either axis is a file shown above its own size and is this side's to
/// scale, and everything at or below the picture's own size is the engine's, as it always
/// was.
#[test]
fn a_box_larger_than_the_picture_is_the_one_this_side_scales() {
    assert!(
        scales_here((640, 480), 1280, 960),
        "shown at twice its size"
    );

    assert!(
        !scales_here((640, 480), 640, 480),
        "a preview at 100% is the picture itself, and nothing is scaled by anybody"
    );
    assert!(!scales_here((640, 480), 320, 240), "a preview below 100%");

    assert!(
        scales_here((720, 480), 872, 480),
        "a pixel wider than it is tall is a stretch along one axis, which is still this side's"
    );
    assert!(
        !scales_here((0, 0), 320, 240),
        "a sound has no picture to scale, and a box of nothing is not one either"
    );
}

/// A box a hair bigger than the picture is not a file shown above its own size, and the
/// difference is the whole of this side's argument: the two sizes do not come out of the
/// layout equal even at 100%, and a comparison rather than a share resamples a whole film
/// for a hundredth of a pixel. So the three cases are the three answers, and the middle one
/// is the one that used to be the first.
#[test]
fn only_an_enlargement_worth_the_resampler_comes_to_this_side() {
    assert!(
        !scales_here((1920, 1080), 1920, 1080),
        "an exact match is the picture itself, and is what a preview at 100% is"
    );

    // One per cent: a 4K file on a 4K display, which is what `fit` places, and the largest
    // enlargement that is not worth resampling for by any distance.
    assert!(
        !scales_here((3840, 2160), 3878, 2182),
        "a box a per cent bigger than a 4K picture is left to the engine"
    );
    assert!(
        !scales_here((1920, 1080), 1939, 1094),
        "and the same a per cent on either axis, however the rounding fell"
    );

    assert!(
        !scales_here((640, 480), 641, 480),
        "nor is one pixel, which is what the rounding of a box that was meant to be the picture gives"
    );
    assert!(!scales_here((640, 480), 640, 481), "on either axis alone");

    assert!(
        scales_here((640, 480), 960, 720),
        "while one and a half times is a file deliberately shown larger than it is"
    );
    assert!(
        scales_here((1920, 1080), 2560, 1440),
        "and a 1080p file enlarged onto a 4K display, which is the case the resampler exists for"
    );
    assert!(
        scales_here((1920, 1080), 1920, 2200),
        "along one axis only, which is what a picture of non-square pixels asks of at every size"
    );
}

/// The two rows a row of the box is mixed from are kept between the rows that need them,
/// and the only way to see whether they are is to count what is read: a resampler that
/// reads each row twice and one that reads each once draw exactly the same frame. At twice
/// the picture's size each source row is wanted by two rows of the box and at four times by
/// four, so the two numbers are one half and one quarter — and a cache that keeps the upper
/// row but reads the lower one again on every tick comes to one and a bit either way.
#[test]
fn every_row_of_the_picture_is_read_once_however_far_it_is_stretched() {
    // A picture big enough for the mapping's arithmetic to be the whole of the cost and
    // small enough to be written out by hand.
    let picture = (100u32, 100u32);
    let source = vec![0u8; picture.0 as usize * picture.1 as usize * 4];

    let doubled = rows_read_per_destination_row(&source, picture, (200, 200));
    assert!(
        doubled <= 0.5,
        "a row of the picture is wanted by two of the box, so it is read once each: {doubled:.3}"
    );

    let quadrupled = rows_read_per_destination_row(&source, picture, (400, 400));
    assert!(
        quadrupled <= 0.25,
        "and by four at four times the size, for the same reason: {quadrupled:.3}"
    );

    // A box the same size as the picture is the other end of it: one row wanted by one row,
    // and one read for it. It is also the last row of the picture under every other box,
    // where the pair is a row and the same row again and so is read once rather than twice.
    let matched = rows_read_per_destination_row(&source, picture, (100, 100));
    assert!(
        matched <= 1.0,
        "the last row of a picture is its own lower half, so a whole frame reads one row fewer than it has: {matched:.3}"
    );

    // One axis equal and the other stretched is what a file of non-square pixels asks for at
    // every size, and it is asked along the height here: the row cache does not care which
    // axis it is, and the horizontal half of this is a copy rather than a resample.
    let stretched = rows_read_per_destination_row(&source, picture, (100, 200));
    assert!(
        stretched <= 0.5,
        "and a box stretched on one axis only reads each row of the picture once for it as well: {stretched:.3}"
    );

    let stretched_across = rows_read_per_destination_row(&source, picture, (200, 100));
    assert!(
        stretched_across <= 1.0,
        "and one row for one row is one read for one row, whatever the columns are doing: {stretched_across:.3}"
    );
}

/// A box as wide as the picture needs no resampling across, and the row it is given is the row
/// it was handed rather than a mix of it with itself: which is the difference between reading
/// a row and copying it, and is what a picture of non-square pixels gets on the axis that is
/// not being stretched.
#[test]
fn a_box_as_wide_as_the_picture_is_given_the_row_unchanged() {
    // One row of a picture, with the alpha bytes a converter chose rather than this app.
    let source = [
        10u8, 20, 30, 0, //
        40, 50, 60, 253, 70, 80, 90, 254, 100, 110, 120, 255,
    ];
    let mut row = [0u8; 16];

    interpolate_row(&source, 16, 4, 0, A_WHOLE_PIXEL, &mut row);

    let expected = [
        [10u8, 20, 30, 255],
        [40, 50, 60, 255],
        [70, 80, 90, 255],
        [100, 110, 120, 255],
    ];

    for (column, pixel) in expected.iter().enumerate() {
        assert_eq!(
            pixel_at(&row, 4, 0, column),
            *pixel,
            "column {column} arrived at from the source pixel and no other, and is opaque"
        );
    }
}

/// The alpha byte the colour converter writes into `ARGB32` is not one anybody could draw
/// with, so the copy writes it rather than reading it: which is what a row of the engine's
/// surface has to be for the preview to be opaque at all.
#[test]
fn a_copied_row_is_opaque_whatever_the_converter_wrote_in_it() {
    // Three pixels carrying the three alpha values a transfer into `MFVideoFormat_ARGB32`
    // has been seen to leave behind, plus the all-zero alpha of a frame nothing has been
    // written into yet.
    let source = [
        10u8, 20, 30, 0, //
        40, 50, 60, 253, 70, 80, 90, 254, 100, 110, 120, 255,
    ];
    let mut row = [0u8; 16];

    copy_row_opaque(&source, &mut row);

    let expected = [
        [10u8, 20, 30, 255],
        [40, 50, 60, 255],
        [70, 80, 90, 255],
        [100, 110, 120, 255],
    ];

    for (column, pixel) in expected.iter().enumerate() {
        assert_eq!(
            pixel_at(&row, 4, 0, column),
            *pixel,
            "pixel {column} keeps the colour it arrived with and is handed an alpha that can be composited"
        );
    }
}

/// A picture the engine wrote less of than it said it would is answered with no frame rather
/// than with one drawn from the rows it managed: half a picture is not a picture, and the
/// frame the preview is holding is a better answer than a half-filled one.
#[test]
fn a_picture_short_of_its_own_size_is_no_frame_at_all() {
    // One row of a two-row picture, read into a box with room for it.
    let source = [0u8; 8];
    let mut out = vec![9u8; 4 * 4 * 4];
    let mut rows = [Vec::new(), Vec::new()];

    assert!(!scale_rows(&source, 8, (2, 2), (4, 4), &mut rows, &mut out));
    assert!(
        out.iter().all(|byte| *byte == 9),
        "and what was there is left as it was"
    );
}

/// The engine offering the frame it offered last is not a frame, and a session that has
/// drawn nothing is owed the first picture whatever the engine has to say about it. This is
/// the whole of the question, and it is asked in hundred-nanosecond units because that is
/// what the tick is answered in.
#[test]
fn a_frame_the_engine_is_still_holding_is_not_drawn_again() {
    assert!(
        is_a_new_frame(None, 0),
        "the first frame of a session is new however the engine times it"
    );

    assert!(
        is_a_new_frame(Some(0), 333_333),
        "and so is the frame that follows the one drawn, a hundredth of a second later"
    );

    assert!(
        !is_a_new_frame(Some(333_333), 333_333),
        "while the same time is the same picture however many ticks it is offered over"
    );

    // What a file that loops hands back, and what a seek hands back that a bar was
    // dragged to: the beginning of the file arriving again, and the time of a frame
    // that did change arriving as the time of one that did not.
    assert!(
        is_a_new_frame(Some(12_000_000), 0),
        "the loop of a file starts its times over without starting its frames"
    );
    assert!(
        is_a_new_frame(None, 12_000_000),
        "and a seek leaves the caller owed the time it lands on, exactly as it was"
    );
}

/// A run of transfers that all failed is a session that has failed, and a run that does not
/// is not: which is the whole of what the bound is for, and the only thing standing between
/// a video that is not moving and a video this side has stopped being able to ask for.
///
/// The two ends are what a file on a network drive is and what this module's own frame path
/// was: a gap, a seek and a loop are a tick or two of this, while a pipeline that cannot give
/// a frame up at all is three seconds of it and never anything else. A bound anywhere near
/// the first number would hand a film over to another engine every time the network hiccuped,
/// and a bound nowhere near the second would leave a preview frozen on a frame from a second
/// ago reporting nothing whatever was wrong.
#[test]
fn a_transfer_that_keeps_failing_is_a_session_that_has_failed() {
    assert!(
        !a_run_of_refusals_is_a_failure(0),
        "a session that has transferred every frame is not failing"
    );

    assert!(
        !a_run_of_refusals_is_a_failure(1),
        "one refused transfer is a seek, a loop landing on its own first time, or a gap"
    );
    assert!(
        !a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP - 1),
        "and so is a run a tick short of the bound, however long the film has been playing"
    );

    assert!(
        a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP),
        "while a run of the whole bound is a pipeline that will not give a frame up"
    );
    assert!(
        a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP * 2),
        "and a longer one is the same fault rather than a worse one"
    );

    // The bound is three seconds of the preview loop's own ticks, which is the same length of
    // time as the wait a session gets for its first frame. A file this side cannot take
    // frames out of is a file another engine has to take over, and the two answers ought to
    // arrive in about the same time whichever order the two mistakes happen in — within a
    // tick, which is the resolution a count of ticks has.
    let bound = Duration::from_millis(TRANSFER_FAILURES_GIVE_UP as u64 * 16);
    assert!(
        bound >= FIRST_FRAME_GIVE_UP && bound < FIRST_FRAME_GIVE_UP + Duration::from_millis(16),
        "the bound is FIRST_FRAME_GIVE_UP counted in the loop's ticks: {bound:?}"
    );
}

/// The four answers a tick can have are four different facts and the caller is given one
/// answer for three of them, which is right — a tick with nothing to paint is a tick with
/// nothing to say — but only because the session keeps them apart for itself. The two
/// faults this module now has to tell apart are a session that never drew and a session
/// that cannot take a frame out of the one it has, and the difference between reporting the
/// second and not the first is this condition.
#[test]
fn a_file_whose_frames_cannot_be_taken_out_is_a_failure_after_the_first_frame_too() {
    let give_up = FIRST_FRAME_GIVE_UP;
    let fresh = FIRST_FRAME_GIVE_UP * 2;
    let short = TRANSFER_FAILURES_GIVE_UP - 1;
    let long = TRANSFER_FAILURES_GIVE_UP;

    assert!(
        engine_is_failing(true, false, 0, give_up),
        "an engine that has been up long enough with nothing on the screen is failing at the \
         file, which is what it was always asked about"
    );
    assert!(
        !engine_is_failing(true, false, 0, Duration::ZERO),
        "while one that has only just started is a file that is merely slow, and taking it \
         away would be a real loss"
    );

    // The case this is for. A session that has drawn a frame and cannot draw another keeps
    // the frame it had, and every question the app asks about it is answered well: the
    // picture on screen is the picture from three seconds ago and nothing is wrong with the
    // preview as far as anything outside can see.
    assert!(
        engine_is_failing(true, true, long, Duration::ZERO),
        "a session whose transfers keep being refused is failing whatever it has drawn"
    );

    // And what must not move with it. `drew` means a frame has been handed over, and the
    // reason is written on `failing_before_a_frame`: a film that played and then met a bad
    // sector is a film the probe got right about, and handing it to another engine over a
    // fault that has nothing to do with what plays it costs the engine its decoder for the
    // rest of a file it was getting on with.
    assert!(
        !engine_is_failing(true, true, short, fresh),
        "a run of refusals short of the bound is a hiccup, and a session that has drawn a \
         frame is not a session that cannot draw"
    );
    assert!(
        !engine_is_failing(true, true, 0, fresh),
        "and a session that is playing normally is not one no matter how long it has been up"
    );

    assert!(
        !engine_is_failing(false, false, long, fresh),
        "a sound has no frames to hand over and is never waiting for one, however it is asked"
    );
}

/// What this side costs to take a frame of a video: frames actually drawn, wall time, and
/// the CPU the process burned while it did — against a real file, played at the box the
/// layout would have placed it at.
///
/// It is the measurement this module's own numbers are argued from, and it is here rather
/// than in a harness of its own because the thing being measured is a thread-local session
/// and a static this module owns; nothing outside can reach them. Three things are
/// measured because a change to this path moves them independently: how many frames the
/// caller was actually handed, how long the window took, and what the process's own clock
/// says it spent. The first says whether the picture moved at all, the second whether the
/// loop kept up with the file, and the third is the whole of the question this module
/// exists to answer — a preview that costs a core of the CPU is a preview that makes every
/// other hover on the desktop stutter.
///
/// The box matters as much as the file. A 320 x 240 box measures the *engine* and nothing
/// this side does: the copies, the resample and the hand to the compositor are all
/// proportional to the number of pixels, so a small box hides the entire cost. The default
/// is a 2560 x 1440 file at `fit` on this machine's display, which is a little over
/// 2493 x 1400 — a box this side does *not* scale into (see `ENLARGEMENT_WORTH_RESAMPLING`),
/// so what it measures is the copy and not the resampler. `RHP_VIDEO_PERF_BOX` overwrites it
/// as `WIDTHxHEIGHT` for a machine whose display is another size.
///
/// The window is eight seconds with the first one spent settling, which is where a first
/// frame's decoder comes up and where a hardware pipeline's device is made; a run that
/// started its clock at `Play` would be measuring setup rather than a preview. The tick is
/// sixteen milliseconds, which is `preview_window::FRAME_WAIT_MS` — that constant belongs to
/// the loop and is not this file's to read, and a measurement taken at any other cadence is
/// not the loop being measured.
///
/// The CPU is the process's own, read from `GetProcessTimes` on this process: kernel and
/// user together, in 100-nanosecond units, over the same window. Shelling out to an external
/// timer would be a second program and its own precision to measure a first.
///
/// Run it by hand:
///
/// ```text
/// cargo test video_take_cost -- --ignored --nocapture
/// ```
///
/// The file it is most worth running on is one that decodes as a rate the display cannot
/// show — a 144 fps file of near-duplicate frames — because that is where the difference
/// between decoding on the GPU and decoding in software is the largest thing in the
/// measurement, and where the waste of decoding frames nobody is shown is at its height.
#[test]
#[ignore = "plays the file named in RHP_VIDEO_PERF at the size of the display"]
fn video_take_cost() {
    let Ok(path) = std::env::var("RHP_VIDEO_PERF") else {
        println!("set RHP_VIDEO_PERF to the path of a video to play");
        return;
    };

    let path = PathBuf::from(path);
    let (width, height) = box_from_env((2493, 1400));
    let picture = dimensions(&path).expect("the file's own frame size");

    println!("\n--- {} ---", path.display());
    println!(
        "the picture is {}x{}, shown in a box of {width}x{height}",
        picture.0, picture.1
    );

    // The crop the geometry probe would have settled on is not asked for: this is a
    // measurement of the frame pipeline and the file's own bars are its own business. A
    // file with a crop is measured a few thousand pixels narrower than this box, which is
    // what it is played at in the app too.
    play(
        &path,
        width,
        height,
        0,
        Picture {
            width: picture.0,
            height: picture.1,
            crop: None,
        },
    );

    assert!(is_playing(), "the engine took the file at all");
    println!(
        "this side is the one that scales into the box: {}",
        scales_here(picture, width, height)
    );

    // The first frame is waited for rather than assumed, because a file the engine cannot
    // draw has none to arrive and a window measured over a session that never drew is a
    // measurement of nothing (see `failing_path`).
    let mut pixels = Vec::new();
    let ticks = Duration::from_millis(TICK_MS);
    let waiting = Instant::now();
    let mut drew = false;

    while waiting.elapsed() < FIRST_FRAME_GIVE_UP {
        if copy_frame_into(&mut pixels).is_some() {
            drew = true;
            break;
        }

        std::thread::sleep(ticks);
    }

    if !drew {
        println!(
            "no frame in {} ms — nothing to measure",
            waiting.elapsed().as_millis()
        );
        stop();
        return;
    }

    // A second of the file is played before the clock is read, so that what is measured is
    // a preview and not a pipeline coming up: a first frame's decoder, and a device that
    // has not been made yet if there is a GPU path to make one.
    let settling = Instant::now();
    let mut settling_frames = 0;
    while settling.elapsed() < Duration::from_secs(1) {
        if copy_frame_into(&mut pixels).is_some() {
            settling_frames += 1;
        }

        std::thread::sleep(ticks);
    }

    // And what the process's own clock says before and after: the kernel's and the user's
    // together, since a decode that is handed to a driver runs partly in one and partly in
    // the other and neither on its own is the cost. A process whose own CPU time cannot be
    // read has no measurement to make, and says so rather than reporting a zero.
    let Some(before) = cpu_hundred_nanoseconds() else {
        println!("this process's own CPU time cannot be read — nothing to measure");
        stop();
        return;
    };

    let started = Instant::now();
    let mut frames = 0;
    let mut refused = 0;

    while started.elapsed() < WINDOW {
        if copy_frame_into(&mut pixels).is_some() {
            frames += 1;
        } else if failing_path().is_some() || !is_playing() {
            refused += 1;
        }

        std::thread::sleep(ticks);
    }

    let wall = started.elapsed();
    let cpu = cpu_hundred_nanoseconds().map_or(0, |now| now.saturating_sub(before));
    let seconds = wall.as_secs_f64();

    println!(
        "\n{frames} frames drawn in {:.2} s ({:.1} a second, over {settling_frames} more \
         while it settled)",
        seconds,
        frames as f64 / seconds,
    );
    println!(
        "process CPU {:.2} s over {:.2} s — {:.1}% of one core",
        cpu as f64 / 10_000_000.0,
        seconds,
        cpu as f64 / 10_000_000.0 / seconds * 100.0,
    );
    println!(
        "each frame cost {:.2} ms of CPU to hand over",
        if frames > 0 {
            cpu as f64 / 10_000_000.0 / frames as f64 * 1000.0
        } else {
            f64::NAN
        }
    );
    println!("{refused} ticks came back with nothing to draw");
    println!("the session is still playing: {}", is_playing(),);

    stop();
}

/// How long the preview loop waits between its ticks, which is the cadence the measurement
/// above is taken at: `preview_window::FRAME_WAIT_MS`, written out here because that
/// constant belongs to the loop and this module does not read it.
const TICK_MS: u64 = 16;

/// How long `video_take_cost` measures a preview for, once it has settled: long enough that
/// the answer is a number rather than a rounding, and short enough to run before a film is
/// over — which matters, because a 144 fps file of eight seconds is over a thousand frames
/// of decode and the file the measurement is most worth running on is twelve minutes long.
const WINDOW: Duration = Duration::from_secs(8);

/// The box a preview is played at, from `RHP_VIDEO_PERF_BOX` as `WIDTHxHEIGHT` or the
/// default this module's own measurements have been taken at: a 1440p file at `fit` on a
/// 2560-wide display, which is the box the layout hands `play` and therefore the size at
/// which every copy and every hand to the compositor is paid for.
fn box_from_env(default: (u32, u32)) -> (u32, u32) {
    let Ok(setting) = std::env::var("RHP_VIDEO_PERF_BOX") else {
        return default;
    };

    let Some((width, height)) = setting.split_once('x') else {
        println!("RHP_VIDEO_PERF_BOX is not WIDTHxHEIGHT — using {default:?}");
        return default;
    };

    match (width.trim().parse(), height.trim().parse()) {
        (Ok(width), Ok(height)) if width > 0 && height > 0 => (width, height),
        _ => {
            println!("RHP_VIDEO_PERF_BOX is not two whole numbers — using {default:?}");
            default
        }
    }
}

/// What this process has spent on the CPU, in 100-nanosecond units: kernel and user
/// together, since a preview that hands its decoding to a driver spends the time in
/// whichever of the two the driver spends it in.
fn cpu_hundred_nanoseconds() -> Option<u64> {
    use windows::Win32::Foundation::FILETIME;
    use windows::Win32::System::Threading::{GetCurrentProcess, GetProcessTimes};

    let mut created = FILETIME::default();
    let mut exited = FILETIME::default();
    let mut kernel = FILETIME::default();
    let mut user = FILETIME::default();

    // SAFETY: all four out-parameters are the caller's own, initialised, and live for the
    // duration of the call; `GetCurrentProcess` is a pseudo-handle that is always valid and
    // is what the process's own times are read from. A refusal reads as no measurement
    // rather than as a zero, which would flatter everything this is used for.
    unsafe {
        GetProcessTimes(
            GetCurrentProcess(),
            &mut created,
            &mut exited,
            &mut kernel,
            &mut user,
        )
    }
    .ok()?;

    let as_units =
        |time: FILETIME| ((time.dwHighDateTime as u64) << 32) | time.dwLowDateTime as u64;

    Some(as_units(kernel) + as_units(user))
}

/// The engine's report of a first frame is the one thing a frame transfer may be asked for,
/// and it arrives as an event on the notify callback rather than as anything to be read off
/// the engine's own state.
///
/// It is the gate `copy_into` waits on: until this flag is set the engine is not ticked, no
/// transfer is asked for, nothing is written to the caller's pixels and the session is not
/// recorded as having drawn. `GetReadyState` cannot stand in for it, because a cold read
/// reaches `HAVE_CURRENT_DATA` before the engine has drawn anything, and the surface a
/// transfer would read is a bitmap of zeros until it does.
#[test]
fn only_the_engines_own_report_of_a_first_frame_is_taken_as_one() {
    let failed = Arc::new(AtomicBool::new(false));
    let first_frame = Arc::new(AtomicBool::new(false));

    // The callback as the engine reaches it: behind a COM object, called by the name the
    // vtable calls it by.
    let notify = windows::core::ComObject::new(Notify {
        failed: Arc::clone(&failed),
        first_frame: Arc::clone(&first_frame),
    });

    notify
        .EventNotify(
            windows::Win32::Media::MediaFoundation::MF_MEDIA_ENGINE_EVENT_PLAY.0 as u32,
            0,
            0,
        )
        .expect("a callback with nothing to record answers with S_OK");
    assert!(
        !first_frame.load(Ordering::Acquire),
        "an engine that has begun to play has not said it has a frame to hand over"
    );

    notify
        .EventNotify(MF_MEDIA_ENGINE_EVENT_FIRSTFRAMEREADY.0 as u32, 0, 0)
        .expect("the callback");
    assert!(
        first_frame.load(Ordering::Acquire),
        "a frame the engine decoded and handed over is the one thing the take waits for"
    );
    assert!(
        !failed.load(Ordering::Acquire),
        "and it is not a failure: the file is one the engine can play"
    );

    notify
        .EventNotify(MF_MEDIA_ENGINE_EVENT_ERROR.0 as u32, 0, 0)
        .expect("the callback");
    assert!(
        failed.load(Ordering::Acquire),
        "an error is still an error, recorded on its own flag beside this one"
    );
}
