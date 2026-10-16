use super::*;

use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};

/// Opens a real capture of this process's own sound, which no ordinary test run may do: it takes a
/// stream on the machine's audio service, and what it answers is a fact about the Windows it is run
/// on as much as about this code.
///
/// It is here because everything *else* about the meter is arithmetic this file can check on its
/// own, while the conversation with the audio service is not: the activation of a process-loopback
/// stream, the format the engine is asked to convert to, and the reading of the packets it hands
/// back are the half of `audio_meter` that no test with a made-up window can reach, and they are the
/// half where a mistake is a bubble that never draws a bar at all. Run by hand where there is an
/// audio device: `cargo test a_real_capture -- --ignored --nocapture`, and what it prints is what
/// the capture answered.
#[test]
#[ignore = "opens a real audio stream, which a test run must never do"]
fn a_real_capture_answers_with_bands_and_not_a_refusal() {
    // Safety: the apartment is the thread's own and the call is the one that takes it; the answer
    // is let go rather than matched, because a thread that already has one is answered the same way
    // it would be by taking one (see `office_render::worker`).
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let bands = audio_meter::output_bands();
    println!("this process's own sound was read as: {bands:?}");

    let bands = bands.expect("this Windows would not let this app's own sound be captured");
    for (band, level) in bands.iter().enumerate() {
        assert!(
            level.is_finite() && (0.0..=1.0).contains(level),
            "the {band} band of a window of real samples read as {level}"
        );
    }
}

/// A full-scale sine at `hz`, as the samples a capture window is read as.
///
/// It is the one signal whose answer can be worked out by hand — a tone is in the band that
/// holds its frequency and it is the loudest thing there — so a band that is read as anything
/// else is a band read wrong, and the arithmetic of a meter is checked against something other
/// than the meter (see `audio_meter::band_levels`).
///
/// The phase is turned in double precision and only the sample is a float: a f32 holding the
/// argument of the last sample of a window is holding four million radians, two radians apart,
/// and the tone it would make of that is noise rather than the tone that was asked for.
fn tone(hz: f32, rate: u32, samples: usize) -> Vec<f32> {
    (0..samples)
        .map(|index| {
            let turn = 2.0 * std::f64::consts::PI * hz as f64 * index as f64 / rate as f64;
            turn.sin() as f32
        })
        .collect()
}

/// The rate the capture asks the engine to convert everything it plays to (see `audio_meter`).
const RATE: u32 = 48_000;
/// The window one band is read over (see `audio_meter::WINDOW_SAMPLES`).
const WINDOW: usize = audio_meter::WINDOW_SAMPLES;

#[test]
fn a_window_of_one_level_is_all_in_the_first_bin() {
    // The transform's own arithmetic, read where a meter cannot see it: a window that holds one
    // level and no sound at all has every bit of its energy at no frequency, and a transform that
    // scattered it — a bit-reversal read backwards, a turn off by one — puts it everywhere.
    let mut real = [1.0f32; WINDOW];
    let mut imag = [0.0f32; WINDOW];
    audio_meter::fft(&mut real, &mut imag);

    // A window of ones is four thousand and ninety-six of them, which is what the first bin is.
    assert!(
        (real[0] - WINDOW as f32).abs() < 1.0,
        "the first bin held {}",
        real[0]
    );
    // And every other bin is nothing beside it: what is left there is the rounding of twelve
    // passes of f32 arithmetic and not a frequency.
    assert!(
        real[1..].iter().all(|bin| bin.abs() < 1.0),
        "a bin away from the first held a frequency"
    );
    assert!(
        imag.iter().all(|bin| bin.abs() < 1.0),
        "a bin away from the first held a frequency"
    );
}

#[test]
fn a_tone_leaves_the_bands_that_do_not_hold_it_alone() {
    // The other half of the same answer, and the one every reading of a real capture rests on: a
    // lone tone is a bar in its own band and nothing in the bands beside it. A row that answered
    // everywhere would read as a meter that moves with everything, which is the one thing a bubble
    // forty-four pixels across cannot afford to say (see `band_levels`).
    let levels = audio_meter::band_levels(&tone(1_000.0, RATE, WINDOW), RATE);

    assert!(
        levels[3] > 0.9,
        "the band that holds a thousand hertz read as {:.3}",
        levels[3]
    );

    for (band, level) in levels.iter().enumerate() {
        if band != 3 {
            assert!(
                *level < 0.1,
                "a lone thousand hertz read as {level:.3} in the {band} band"
            );
        }
    }
}

#[test]
fn a_tone_falls_in_the_band_that_holds_it() {
    // One tone per band the app draws, named by the band it must land in — the lowest, the one
    // the app's own two-thousand hertz band straddles, and the highest. A tower that answered
    // with a band of its own choosing is a row of bars that moves for the wrong reason.
    for (hz, band) in [(60.0, 0usize), (1_000.0, 3), (12_000.0, 6)] {
        let levels = audio_meter::band_levels(&tone(hz, RATE, WINDOW), RATE);
        let loudest = (0..audio_meter::BARS)
            .max_by(|left, right| levels[*left].total_cmp(&levels[*right]))
            .expect("seven bands have a loudest");

        assert_eq!(loudest, band, "a {hz} Hz tone was read as band {loudest}");
        // A full-scale tone is the top of a bar: the window scatters a little of it (the tone
        // is between two bins as often as on one, which costs under two decibels) and nothing
        // else about it is quiet.
        assert!(
            levels[band] > 0.9,
            "a full-scale {hz} Hz tone read as {:.3}",
            levels[band]
        );
    }
}

#[test]
fn silence_is_no_bars_at_all() {
    // A pin over a sound nobody is hearing must draw no bars rather than a row of them at
    // nothing, which is what the drawing asks of this answer and what the mixer's own peak
    // answers today (see `draw_histogram`).
    assert_eq!(
        audio_meter::band_levels(&[0.0; WINDOW], RATE),
        [0.0; audio_meter::BARS]
    );
}

#[test]
fn a_tone_at_half_scale_reads_as_a_lower_bar() {
    // Half the amplitude is six decibels down, which is a tenth of the sixty a bar is drawn
    // over — so a tone played at half scale is a bar a little under the top one, and never
    // above it. The two values come from the decibels in this comment and not from the meter.
    let full = audio_meter::band_levels(&tone(1_000.0, RATE, WINDOW), RATE)[3];
    let half: Vec<f32> = tone(1_000.0, RATE, WINDOW)
        .into_iter()
        .map(|sample| sample * 0.5)
        .collect();
    let half = audio_meter::band_levels(&half, RATE)[3];

    assert!(
        (0.85..0.95).contains(&half),
        "a half-scale tone read as {half:.3}, which is not six decibels down"
    );
    assert!(half < full, "a quieter tone read as a taller bar");
}
