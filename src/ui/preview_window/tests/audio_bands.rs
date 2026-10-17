use super::*;

use std::time::Duration;

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
/// than the meter (see `audio_meter::readings`).
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
    // forty-four pixels across cannot afford to say (see `readings`).
    let (levels, _) = audio_meter::readings(&tone(1_000.0, RATE, WINDOW), RATE, None);

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
        let (levels, _) = audio_meter::readings(&tone(hz, RATE, WINDOW), RATE, None);
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
        audio_meter::readings(&[0.0; WINDOW], RATE, None).0,
        [0.0; audio_meter::BARS]
    );
}

#[test]
fn a_tone_at_half_scale_reads_as_a_lower_bar() {
    // Half the amplitude is six decibels down, which is a tenth of the sixty a bar is drawn
    // over — so a tone played at half scale is a bar a little under the top one, and never
    // above it. The two values come from the decibels in this comment and not from the meter.
    let full = audio_meter::readings(&tone(1_000.0, RATE, WINDOW), RATE, None).0[3];
    let half: Vec<f32> = tone(1_000.0, RATE, WINDOW)
        .into_iter()
        .map(|sample| sample * 0.5)
        .collect();
    let half = audio_meter::readings(&half, RATE, None).0[3];

    assert!(
        (0.85..0.95).contains(&half),
        "a half-scale tone read as {half:.3}, which is not six decibels down"
    );
    assert!(half < full, "a quieter tone read as a taller bar");
}

#[test]
fn a_bass_tone_reads_as_a_pulse_and_a_tone_above_it_as_none() {
    // The card's ring travels by the bass alone: a tone inside the range is the whole of the
    // pulse, and one above it — what a bass line's harmonics share the air with — is nothing at
    // all, or the ring would beat for cymbals and vocals as much as for the drum (see
    // `readings`).
    let (_, bass) = audio_meter::readings(&tone(100.0, RATE, WINDOW), RATE, DEFAULT_BEAT_HZ.top());
    assert!(
        bass > 0.9,
        "a full-scale hundred hertz tone read as {bass:.3}"
    );

    let (_, treble) =
        audio_meter::readings(&tone(6_000.0, RATE, WINDOW), RATE, DEFAULT_BEAT_HZ.top());
    assert_eq!(treble, 0.0, "a tone above the bass moved the pulse");
}

#[test]
fn silence_is_no_pulse_at_all() {
    // Nothing is playing, so the mark stands at rest — the same answer a machine that refuses
    // the capture gets, which is what the ring is drawn without.
    let (_, beat) = audio_meter::readings(&[0.0; WINDOW], RATE, DEFAULT_BEAT_HZ.top());
    assert_eq!(beat, 0.0);
}

#[test]
fn a_quieter_bass_reads_as_a_shorter_travel() {
    // Six decibels down is a sixth of the thirty-six the pulse travels over (see `BEAT_DB`), so
    // a tone at half scale stands at five sixths of the ring's travel. The two values come from
    // the decibels in this comment and not from the meter.
    let (_, full) = audio_meter::readings(&tone(100.0, RATE, WINDOW), RATE, DEFAULT_BEAT_HZ.top());
    let half: Vec<f32> = tone(100.0, RATE, WINDOW)
        .into_iter()
        .map(|sample| sample * 0.5)
        .collect();
    let (_, half) = audio_meter::readings(&half, RATE, DEFAULT_BEAT_HZ.top());

    assert!(
        (0.78..0.89).contains(&half) && half < full,
        "a half-scale bass read as {half:.3} against {full:.3}"
    );
}

#[test]
fn the_pulse_jumps_to_a_louder_window_and_falls_by_the_clock() {
    // A hit stands the ring at itself, and a drop falls back through the reads that follow — by
    // the time between them, so that a card's reads and a bubble's agree (see `beat_pulse`).
    let landed = audio_meter::beat_pulse(0.2, 0.8, Duration::from_millis(33));
    assert!(
        (landed - 0.8).abs() < 1e-6,
        "a hit of eight tenths left the pulse at {landed:.3}"
    );

    // One time constant of the hundred and fifty milliseconds the release names is 1/e of what
    // was held — a tenth place of eight tenths is 0.294 — and that is where the range below
    // comes from rather than from the meter.
    let fallen = audio_meter::beat_pulse(0.8, 0.0, Duration::from_millis(150));
    assert!(
        (0.27..0.32).contains(&fallen),
        "one time constant of falling left {fallen:.3} of eight tenths"
    );

    // And a window still holding a louder bass is never fallen through: the ring stands at what
    // the sound is doing rather than dropping below it.
    let held = audio_meter::beat_pulse(0.3, 0.5, Duration::from_millis(300));
    assert!(
        (held - 0.5).abs() < 1e-6,
        "a sustained bass left the pulse at {held:.3}"
    );
}

#[test]
fn a_tone_over_the_settings_band_moves_no_pulse() {
    // The top of the band is the setting's (see `BeatHz`), so one window is a pulse under one
    // setting and nothing under another: fifteen hundred hertz stands well past the top of the
    // default band and under the top of the widest one the app offers, and a band that is off
    // is no pulse whatever the window holds.
    let inside =
        audio_meter::readings(&tone(100.0, RATE, WINDOW), RATE, DEFAULT_BEAT_HZ.top()).1;
    assert!(
        inside > 0.9,
        "a bass tone under the default top read as {inside:.3}"
    );

    let over =
        audio_meter::readings(&tone(1_500.0, RATE, WINDOW), RATE, DEFAULT_BEAT_HZ.top()).1;
    assert_eq!(over, 0.0, "a tone above the default top moved the pulse");

    // The same tone is the pulse where the band reaches it, which is what makes the line above
    // a reading of the band and not of the tone.
    let kept = audio_meter::readings(&tone(1_500.0, RATE, WINDOW), RATE, Some(16_000)).1;
    assert!(kept > 0.9, "a tone under a band that reaches it read as {kept:.3}");

    let off = audio_meter::readings(&tone(100.0, RATE, WINDOW), RATE, None).1;
    assert_eq!(off, 0.0, "a band that is off read a pulse");
}

#[test]
fn a_fall_comes_home_by_the_clock_and_reaches_rest() {
    // What a pause draws the mark by (see `beat_fade`): the pulse the card last held is at 1/e
    // of itself one BEAT_FADE later — the value in this comment and not from the meter — and
    // the fall is slower than the release a live pulse falls by, which is what makes a pause a
    // fall rather than a snap.
    let held = 0.8f32;
    let at_once = audio_meter::beat_fade(held, Duration::from_millis(0));
    assert!(
        (at_once - held).abs() < 1e-6,
        "a fall of no time left {at_once:.3}"
    );

    let one = audio_meter::beat_fade(held, Duration::from_millis(400));
    assert!(
        (0.27..0.32).contains(&one),
        "one time constant of falling left {one:.3} of eight tenths"
    );

    let released = audio_meter::beat_pulse(held, 0.0, Duration::from_millis(400));
    assert!(one > released, "a fall was not slower than the release");

    // Under a hundredth the mark is at rest for good, which is what ends the painting a fall
    // is drawn by.
    assert_eq!(audio_meter::beat_fade(held, Duration::from_secs(2)), 0.0);
}

#[test]
fn a_wait_the_loop_missed_falls_by_a_step_and_not_by_the_whole_of_it() {
    // A card painted at its own cadence counts the whole of the wait between its frames, which is
    // the clock a fall goes home by (see `beat_step`).
    assert_eq!(
        audio_meter::beat_step(Duration::from_millis(33)),
        Duration::from_millis(33)
    );
    // A wait longer than that is a frame the loop came back late for, and it is counted as the one
    // step: a quarter of the fall's own clock, and so the share of what is left of it that one
    // frame can take.
    assert_eq!(
        audio_meter::beat_step(Duration::from_millis(1_500)),
        Duration::from_millis(100)
    );
    assert_eq!(
        audio_meter::beat_step(Duration::from_secs(5)),
        Duration::from_millis(100)
    );

    // The value in this comment and not from the arithmetic again: eight tenths a quarter of a
    // clock later is 0.8 * e^-0.25.
    let held = 0.8f32;
    let stepped =
        audio_meter::beat_fade(held, audio_meter::beat_step(Duration::from_millis(1_500)));
    assert!(
        (0.60..0.65).contains(&stepped),
        "a wait the loop missed took the beat to {stepped:.3} of eight tenths"
    );

    // Where the same wait counted whole is the whole of the travel gone in one frame — the jump
    // this step is here to stop (0.8 * e^-3.75).
    let whole = audio_meter::beat_fade(held, Duration::from_millis(1_500));
    assert!(
        whole < 0.05,
        "the uncounted wait left {whole:.3} of eight tenths"
    );
}
