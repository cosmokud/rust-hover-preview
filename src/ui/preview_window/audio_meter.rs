//! The level the sound on screen is being heard at: the seven bands of this app's own sound, and
//! the mixer's own peak for the machine that cannot read them.
//!
//! **The bands are what this app is playing and nothing else.** They are read from a process
//! loopback capture of this process and its children — the media engine plays in this process,
//! FFmpeg's player in one of its own — so the same row of bars is filled whether the file on
//! screen is played by one player or the other, and nothing the user is playing in another window
//! moves them. The sound itself is read and turned into seven numbers and is not kept, written or
//! sent anywhere (see `ProcessLoopback`).
//!
//! Where the capture cannot be had the bars fall back to the mixer's own peak: one number, read
//! from the default output endpoint, which is every sound the machine is making rather than this
//! app's own. It is the other machine's answer — process loopback arrived after Windows 10 — and
//! it is what a bubble has always drawn its bars from, so a machine without the capture keeps the
//! meter it has rather than a row of bars at nothing (see `output_peak` and `output_bands`).
//!
//! Nothing here asks for an apartment of its own. The meter belongs to the thread that made it, so
//! it is cached per thread — and the one caller is the preview loop, whose thread has already taken
//! an apartment at the top of the loop by way of the two engines that need one (see
//! `pdf_preview::initialize_apartment` and the head of `run_preview_window`).
//!
//! Nothing here asks for an apartment of its own. The meter belongs to the thread that made it, so
//! it is cached per thread — and the one caller is the preview loop, whose thread has already taken
//! an apartment at the top of the loop by way of the two engines that need one (see
//! `pdf_preview::initialize_apartment` and the head of `run_preview_window`).

use super::*;

use std::sync::OnceLock;

use windows::Win32::Media::Audio::Endpoints::IAudioMeterInformation;
use windows::Win32::Media::Audio::{eConsole, eRender, IMMDeviceEnumerator, MMDeviceEnumerator};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_ALL};

thread_local! {
    /// The meter this thread reads the level from, opened on the first ask.
    ///
    /// It is opened once and kept, rather than asked for per bubble, because opening one is a walk
    /// of the machine's device tree: a bubble over a sound is painted every thirty-three
    /// milliseconds while its bars animate, and an endpoint walked that often is work with nothing
    /// to show for it — the default endpoint is the one the files are already going out of.
    ///
    /// What it is *not* held across is a refusal: a machine whose walk fails has an empty slot
    /// again on the next ask, so a device that appears later (a headset plugged in, a driver that
    /// finished starting) is found by the next bubble rather than by the next run of the app. The
    /// cost of that is the walk itself, which is the case this arrangement is least often in.
    static METER: RefCell<Option<IAudioMeterInformation>> = const { RefCell::new(None) };
}

/// The peak the sound on screen is being heard at, 0.0..=1.0, or nothing where the machine cannot
/// be asked.
///
/// Nothing here panics and nothing here reports: a bubble is a picture of a file and not a promise
/// about the mixer, so a machine with no output endpoint, a mixer that refuses the meter and a
/// thread without an apartment are all the same answer — no level — and the caller draws its bars
/// from zero, which is the truth about a sound nothing is hearing (see `bubble_art`).
pub(super) fn output_peak() -> Option<f32> {
    METER.with(|slot| {
        let mut slot = slot.borrow_mut();

        if slot.is_none() {
            *slot = open_meter();
        }

        // Safety: the meter is this thread's own — opened here and never handed to another thread
        // — and `GetPeakValue` reads one number out of it and writes nothing anywhere.
        let peak = unsafe { slot.as_ref()?.GetPeakValue() };
        peak.ok()
    })
}

/// Open the meter on this machine's default output endpoint.
///
/// The default endpoint is asked for rather than a device chosen by name or a session hunted out on
/// one: what a bubble over a sound has to say is how loud the sound is in the room, and the device
/// the room is listening through is the one the user told Windows to play through — `eRender` for
/// the direction, `eConsole` for the role, which is the device a machine's own sounds come out of
/// rather than the one a chat program asks for so its notifications keep arriving over the music.
///
/// The meter rather than the session's own peak on purpose: a session meter would name the process
/// playing the file, and the file is played by whichever of two players this machine happens to
/// have — the engine Windows has plays in this process, FFmpeg's plays in one of its own — so the
/// device meter is the one question that is the same answer either way (see the module note).
fn open_meter() -> Option<IAudioMeterInformation> {
    unsafe {
        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL).ok()?;
        let device = enumerator.GetDefaultAudioEndpoint(eRender, eConsole).ok()?;

        device
            .Activate::<IAudioMeterInformation>(CLSCTX_ALL, None)
            .ok()
    }
}

// ------------------------------------------------------------------------ the bands

/// The number of bands a bubble's bars are drawn for, which is the number of bars.
pub(super) const BARS: usize = 7;

/// The seven bands an equalizer's own sliders stand for — 60, 150, 400, 1000, 2400, 6000 and
/// 15000 hertz — taken as ranges: each neighbouring pair's geometric mean is the border between
/// them, and the row as a whole is bracketed by 40 hertz and 20 kilohertz so that a sound has
/// somewhere to be at every frequency it can carry.
///
/// These numbers are the whole of what makes a row of bars read *low on the left and high on the
/// right* to anyone who has seen a player's own meter, which is the only thing a bubble forty-four
/// pixels across can say about frequency (see `pin_chrome::draw_histogram`).
const BAND_EDGES: [f32; BARS + 1] = [
    40.0, 95.0, 245.0, 632.0, 1_550.0, 3_800.0, 9_500.0, 20_000.0,
];

/// How many samples a band is read over: an eighty-fifth of a second at the rate the capture asks
/// for.
///
/// Long enough that the lowest band is a measurement rather than a guess — 40 hertz is three and a
/// half of its own cycles inside this window — and short enough that the row still follows what is
/// being played now rather than reporting what was played a fifth of a second ago.
pub(super) const WINDOW_SAMPLES: usize = 4_096;

/// How far down a bar reaches: sixty decibels between a sound at the top of the scale and a bar at
/// nothing, which is a thousandth of its pressure.
///
/// It is the range rather than a wider one because what a person hears at all should move the bar
/// that stands for it: a range of ninety decibels would leave everything quieter than a full-scale
/// tone within a third of a bar of the floor.
const RANGE_DB: f32 = 60.0;

/// The level of each band of a window of samples, 0.0..=1.0, the lowest band first.
///
/// The whole of the meter is here, and it is arithmetic on numbers rather than anything about a
/// machine: what a window of mono samples says about seven ranges of frequency. The window is
/// transformed once (`fft`) and a band is the loudest bin that falls inside it — seven readings of
/// one spectrum rather than seven filters, so that a band stands for everything in its range and
/// not only for the frequencies a filter happened to be pointed at.
///
/// The samples are the last `WINDOW_SAMPLES` of what is given and are padded at the front where
/// there are fewer of them, so what is read is the sound being played now and a capture that has
/// just begun reads as a quiet window rather than as a window of the wrong length.
pub(super) fn band_levels(samples: &[f32], rate: u32) -> [f32; BARS] {
    let count = samples.len().min(WINDOW_SAMPLES);
    if count < 2 {
        return [0.0; BARS];
    }

    // The samples read through one window for the whole row rather than through one window per
    // band: what the seven readings have in common is the spectrum they are read from.
    let window = hann();
    let padding = WINDOW_SAMPLES - count;
    let mut real = [0.0f32; WINDOW_SAMPLES];
    let mut imag = [0.0f32; WINDOW_SAMPLES];
    for (index, sample) in samples[..count].iter().enumerate() {
        real[padding + index] = sample * window[padding + index];
    }

    fft(&mut real, &mut imag);

    // What a full-scale sine reads as, so that a bar reaches its top at a sound at the top of the
    // scale. The window's own magnitude for such a sine is a quarter of its length — a half-cosine
    // keeps half of the tone, and half of that is the tone's own positive frequencies — so a
    // magnitude is scaled by four over the window to come out as the sine's own amplitude.
    let scale = 4.0 / WINDOW_SAMPLES as f32;
    let per_hertz = WINDOW_SAMPLES as f32 / rate as f32;
    // Half the window is half the rate, and a spectrum has nothing above that to read.
    let nyquist = WINDOW_SAMPLES / 2;

    let mut levels = [0.0f32; BARS];
    for (level, edges) in levels.iter_mut().zip(BAND_EDGES.windows(2)) {
        // The first bin rather than the zeroth, which is a sound at no frequency at all: a capture
        // offset from zero would otherwise light the lowest bar with a number and not with a sound.
        let low = (edges[0] * per_hertz).round().max(1.0) as usize;
        let high = ((edges[1] * per_hertz).round() as usize).min(nyquist);
        let mut loudest = 0.0f32;

        for bin in low..=high {
            loudest = loudest.max(real[bin].hypot(imag[bin]));
        }

        *level = bar_level(scale * loudest);
    }

    levels
}

/// The spectrum of a window of samples, in place: `real` and `imag` come back as the bins of the
/// transform, the lowest frequency first.
///
/// A radix-2 Cooley-Tukey transform, which is what a spectrum at a length that is a power of two
/// is: the samples read in the bit-reversed order the butterfly stages assume, and then one pass a
/// doubling of the length in which two half-length transforms are combined with a complex
/// multiplication by a turn from a table made once (see `turns`).
pub(super) fn fft(real: &mut [f32], imag: &mut [f32]) {
    let count = real.len();
    debug_assert_eq!(count, WINDOW_SAMPLES);

    // The order the samples are read in, which is the index read backwards a bit at a time.
    let mut reversed = 0usize;
    for index in 1..count {
        let mut bit = count >> 1;
        while reversed & bit != 0 {
            reversed ^= bit;
            bit >>= 1;
        }
        reversed |= bit;

        if index < reversed {
            real.swap(index, reversed);
            imag.swap(index, reversed);
        }
    }

    let turns = turns();
    let mut length = 2;
    while length <= count {
        let half = length / 2;
        let step = count / length;

        for start in (0..count).step_by(length) {
            for index in 0..half {
                let (cosine, sine) = turns[index * step];
                let next = start + index + half;
                let odd = (real[next], imag[next]);
                let turned = (odd.0 * cosine - odd.1 * sine, odd.0 * sine + odd.1 * cosine);
                let even = (real[start + index], imag[start + index]);

                real[start + index] = even.0 + turned.0;
                imag[start + index] = even.1 + turned.1;
                real[next] = even.0 - turned.0;
                imag[next] = even.1 - turned.1;
            }
        }

        length <<= 1;
    }
}

/// The turns the transform multiplies by, in the order one stage of it reads them: `e^(-2πik/n)`
/// for the kth of half a window.
///
/// The whole of what a transform of this length ever asks for, made once rather than as a cosine
/// a butterfly — the length is fixed and the call is thirty times a second (see `fft`).
fn turns() -> &'static [(f32, f32)] {
    static TURNS: OnceLock<Vec<(f32, f32)>> = OnceLock::new();

    TURNS.get_or_init(|| {
        (0..WINDOW_SAMPLES / 2)
            .map(|index| {
                let turn = -2.0 * std::f32::consts::PI * index as f32 / WINDOW_SAMPLES as f32;
                (turn.cos(), turn.sin())
            })
            .collect()
    })
}

/// The top of a bar an amplitude is drawn at: the decibels it stands for, over the range a bar
/// spans, so that a sound a tenth as loud is a bar a tenth as tall rather than a hundredth.
fn bar_level(amplitude: f32) -> f32 {
    let decibels = 20.0 * amplitude.max(1e-6).log10();
    ((decibels + RANGE_DB) / RANGE_DB).clamp(0.0, 1.0)
}

/// The window every band is read through, made once and kept.
///
/// A Hann window, which is what keeps a loud low note out of the bars that stand for high
/// frequencies: a band read through a rectangle of samples hears a tone a couple of hundred hertz
/// below it at around a twentieth of its strength, and through a half-cosine at around a
/// thousandth. It costs one multiplication a sample, once, for all seven bands, and what it buys
/// is a row of bars that does not all move together every time a drum is hit.
fn hann() -> &'static [f32; WINDOW_SAMPLES] {
    static WINDOW: OnceLock<[f32; WINDOW_SAMPLES]> = OnceLock::new();

    WINDOW.get_or_init(|| {
        let mut window = [0.0f32; WINDOW_SAMPLES];
        for (index, value) in window.iter_mut().enumerate() {
            let turn = 2.0 * std::f32::consts::PI * index as f32 / (WINDOW_SAMPLES - 1) as f32;
            *value = 0.5 - 0.5 * turn.cos();
        }
        window
    })
}
