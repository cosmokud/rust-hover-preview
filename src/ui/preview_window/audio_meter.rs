//! The level the sound on screen is being heard at: the seven bands of this app's own sound, and
//! the mixer's own peak for the machine that cannot read them.
//!
//! **The bands are what this app is playing and nothing else.** They are read from a process
//! loopback capture of this process and its children — the media engine plays in this process,
//! FFmpeg's player in one of its own — so the same row of bars is filled whether the file on
//! screen is played by one player or the other, and nothing the user is playing in another window
//! moves them. The sound itself is read and turned into seven numbers and is not kept, written or
//! sent anywhere (see `open_capture`).
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

use std::mem::ManuallyDrop;
use std::sync::{Arc, Condvar, Mutex, OnceLock};
use std::time::{Duration, Instant};

use windows::core::{implement, Interface, HRESULT};
use windows::Win32::Media::Audio::Endpoints::IAudioMeterInformation;
use windows::Win32::Media::Audio::{
    eConsole, eRender, ActivateAudioInterfaceAsync, IActivateAudioInterfaceAsyncOperation,
    IActivateAudioInterfaceCompletionHandler, IActivateAudioInterfaceCompletionHandler_Impl,
    IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator, MMDeviceEnumerator,
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
    AUDCLNT_STREAMFLAGS_LOOPBACK, AUDIOCLIENT_ACTIVATION_PARAMS, AUDIOCLIENT_ACTIVATION_PARAMS_0,
    AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK, AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
    PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE, VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
    WAVEFORMATEX, WAVE_FORMAT_PCM,
};
use windows::Win32::System::Com::{CoCreateInstance, IAgileObject, IAgileObject_Impl, CLSCTX_ALL};
use windows::Win32::System::Variant::VT_BLOB;
use windows_core::PROPVARIANT;

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

/// The number of bands a bubble's bars are drawn for, which is the number of bars — the row the
/// chrome draws, so that the two cannot come apart (`pin_chrome::BUBBLE_BARS`).
pub(super) const BARS: usize = pin_chrome::BUBBLE_BARS;

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

// --------------------------------------------------------------------- the capture

/// The rate the capture asks the engine to convert everything this app plays to.
///
/// It is this side's own number rather than the machine's, because the format is this side's own
/// too (see `open_capture`): a shared-mode stream is converted to whatever its client asks for,
/// and 48000 is a rate every sound card's engine already runs at.
const RATE: u32 = 48_000;

/// The channels the capture asks for — the two a mix has, averaged into the one window the bands
/// are read from (see `Bands::drain`).
const CHANNELS: usize = 2;

/// What one frame of the capture is: the two channels of sixteen bits each (see `CHANNELS`).
const FRAME_BYTES: usize = CHANNELS * 2;

/// How long the capture's own buffer is, in the hundred-nanosecond units a stream is asked in.
///
/// Two hundred milliseconds, which is six repaints of a bubble: a client that does not drain its
/// buffer in time loses what was in it, and this is what makes the polling in `Bands::read` a
/// cadence rather than a deadline.
const BUFFER: i64 = 2_000_000;

/// How long a bubble may stop asking before the stream is stopped.
///
/// The capture is asked for while a bubble over a playing sound is being painted and not at all
/// while the pin is open, so this is what decides that a window nobody is looking at is not a
/// stream left running.
const IDLE: Duration = Duration::from_secs(1);

/// How much of a bar is left after one repaint with nothing behind it.
///
/// A bar jumps to what its band is doing and falls back through this over the repaints that
/// follow, which is what makes a row of bars move like a meter rather than flicker like a
/// readout: a bar is at a tenth of its height a quarter of a second later.
const FALL: f32 = 0.75;

/// How long an activation is given before the bubble gives up on it.
///
/// A machine that answers at all answers within a millisecond or two. What this bounds is a system
/// that has taken the ask and means never to answer it, which would otherwise be a bubble that is
/// never painted again.
const ACTIVATION_WAIT: Duration = Duration::from_millis(250);

/// The level of each band of the sound this app is playing now, 0.0..=1.0, the lowest band first,
/// or nothing where the machine will not let the capture be had.
///
/// Nothing here panics, fails or reports: a bubble is a picture of a file rather than a promise
/// about the machine, so a Windows without process loopback, a refusal from the audio service and
/// a system that never answers are all the same answer — no bands — and the caller falls back to
/// the mixer's peak, which is what a bubble has always drawn its bars from (see `output_peak`).
pub(super) fn output_bands() -> Option<[f32; BARS]> {
    CAPTURE.with(|slot| {
        let mut slot = slot.borrow_mut();

        if slot.is_none() {
            *slot = Some(match open_capture() {
                Some(capture) => Reading::Bands(capture),
                // Kept rather than tried again on the next ask: this is not a device that may be
                // plugged in later, it is a Windows that does not have process loopback at all.
                None => Reading::Refused,
            });
        }

        match slot.as_mut()? {
            Reading::Bands(capture) => capture.read(),
            Reading::Refused => None,
        }
    })
}

thread_local! {
    /// The capture the thread the bubble is drawn on has, opened on the first ask and never asked
    /// for again (see `Reading`).
    static CAPTURE: RefCell<Option<Reading>> = const { RefCell::new(None) };
}

/// What this thread has to say about the capture: the stream it has opened, or that the machine
/// has refused one.
///
/// The refusal is remembered for the life of the thread, which is the opposite of what the meter
/// above does with its own slot — and the two failures are not the same kind. An endpoint missing
/// today can be plugged in by tomorrow: the walk that finds it is worth repeating and costs
/// nothing. A capture of this process's own sound is either something this Windows does at all or
/// something it has never heard of (see `open_capture`), so asking again thirty times a second
/// would be an activation and a device walk per repaint, for the same nothing.
enum Reading {
    /// The stream, and the window the bands are read from (see `Bands`).
    Bands(Bands),
    /// The machine will not have it, and will not have it the next time either.
    Refused,
}

/// A running capture of this app's own sound, and the window the bands are read from.
struct Bands {
    client: IAudioClient,
    capture: IAudioCaptureClient,
    /// The mono samples of the last window, oldest first.
    ///
    /// A `Vec` that is pushed to and trimmed rather than a ring that the reading is stitched back
    /// out of: a window is sixteen kilobytes, and a ring would be bought with an index, a length
    /// and a copy per read to save one move of them.
    window: Vec<f32>,
    /// The bars as they stand, which is what the falling is: a band quieter than it was is read
    /// through the bar that is already up (see `FALL`).
    shown: [f32; BARS],
    /// When a bubble last asked, which is the whole of what tells an idle capture from a busy one.
    asked: Instant,
    /// When a packet last brought samples, which is what tells a stream that has stopped giving
    /// anything from one that is playing silence — the engine hands silence over rather than
    /// handing nothing (see `read`).
    last_packet: Instant,
    /// Whether the stream is running, since it is stopped — and not closed — when nothing asks.
    running: bool,
}

impl Bands {
    /// What the seven bands are doing now, read from every packet the capture is holding.
    fn read(&mut self) -> Option<[f32; BARS]> {
        let idle = self.asked.elapsed() > IDLE;
        self.asked = Instant::now();
        let taken: usize;

        // Safety: the client and the capture are this thread's own, made here and never handed to
        // another thread, and every call below reads or writes the stream this struct owns.
        unsafe {
            if self.running && idle {
                let _ = self.client.Stop();
                self.running = false;
                // A stopped stream leaves its last samples behind, and they are older than
                // anything the bars should be drawn from.
                self.window.clear();
            }

            if !self.running {
                self.client.Start().ok()?;
                self.running = true;
            }

            taken = self.drain();
        }

        if taken > 0 {
            self.last_packet = Instant::now();
        } else if self.last_packet.elapsed() > IDLE {
            // A stream that has brought nothing at all for a second is a stream that has stopped —
            // the device went, the service went, the file ended — rather than one making no sound,
            // and a window it is no longer filling would leave the bars standing at the last thing
            // the sound was doing for as long as the pin is up.
            self.window.clear();
        }

        if self.window.len() > WINDOW_SAMPLES {
            let oldest = self.window.len() - WINDOW_SAMPLES;
            self.window.drain(..oldest);
        }

        let levels = band_levels(&self.window, RATE);
        for (shown, level) in self.shown.iter_mut().zip(levels) {
            *shown = level.max(*shown * FALL);
        }

        Some(self.shown)
    }

    /// Take everything the capture has to give, as the mono samples a window is read from, and
    /// answer how many of them there were.
    ///
    /// Every packet is taken rather than the newest, because there is no asking for one: a packet
    /// is however much the engine chose to hand over — the same routine can find ten milliseconds
    /// or forty-eight of them waiting — and what is not taken now is what the buffer loses when
    /// the next one arrives. They are averaged to one channel and read as the window's own.
    ///
    /// A packet that cannot be read is where this stops rather than a refusal of the whole
    /// capture: what is already in the window is what the bars are drawn from, and how long ago
    /// the last one arrived is the caller's to read (see `read`).
    ///
    /// # Safety
    ///
    /// The capture must belong to this thread, which is the case for the one it is called on
    /// (see `read`).
    unsafe fn drain(&mut self) -> usize {
        let mut taken = 0usize;

        loop {
            let Ok(frames) = self.capture.GetNextPacketSize() else {
                return taken;
            };
            if frames == 0 {
                return taken;
            }

            let mut data: *mut u8 = std::ptr::null_mut();
            let mut given = 0u32;
            let mut flags = 0u32;
            let Ok(()) = self
                .capture
                .GetBuffer(&mut data, &mut given, &mut flags, None, None)
            else {
                return taken;
            };

            // A packet marked silent holds no numbers worth reading — the engine says so rather
            // than writing zeros — and it is the same silence the bars are drawn flat for.
            if flags & AUDCLNT_BUFFERFLAGS_SILENT.0 as u32 == 0 && !data.is_null() {
                let bytes = std::slice::from_raw_parts(data, given as usize * FRAME_BYTES);
                let (frames, _) = bytes.as_chunks::<FRAME_BYTES>();

                for frame in frames {
                    let left = i16::from_le_bytes([frame[0], frame[1]]) as i32;
                    let right = i16::from_le_bytes([frame[2], frame[3]]) as i32;
                    self.window.push((left + right) as f32 / 65_536.0);
                }
            } else {
                self.window.extend(std::iter::repeat_n(0.0, given as usize));
            }

            let Ok(()) = self.capture.ReleaseBuffer(given) else {
                return taken;
            };
            taken += given as usize;
        }
    }
}

/// Open a capture of this app's own sound, or nothing where the machine will not have it.
///
/// **It is a capture of this process and its children, which is exactly what this app plays.** The
/// media engine plays in this process and FFmpeg's player in one of its own, so one capture answers
/// for both, and the film an mpv child plays is in it too — while nothing another window is playing
/// is, which is the whole of the difference between this and the meter above (see `output_peak`).
///
/// The device named is a *virtual* one — `VAD\Process_Loopback`, a name the audio service knows —
/// rather than a real endpoint, which is why nothing here asks the machine for a default output
/// device and why a capture works for a machine with none. The parameters that pick the process
/// travel as a blob inside a variant, which is the one shape an activation is asked in.
///
/// What comes back is asked for in the app's own format — 16-bit stereo at `RATE` — rather than in
/// the engine's, and the engine converts: this is the short way round because the mix format is not
/// readable on such a client at all (`GetMixFormat` answers `E_NOTIMPL` on one of these, which is a
/// fact of the interface rather than of a machine) and because a format this side chose is one the
/// reading of a window can be sure about (see `band_levels`).
fn open_capture() -> Option<Bands> {
    let params = AUDIOCLIENT_ACTIVATION_PARAMS {
        ActivationType: AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK,
        Anonymous: AUDIOCLIENT_ACTIVATION_PARAMS_0 {
            ProcessLoopbackParams: AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS {
                TargetProcessId: std::process::id(),
                ProcessLoopbackMode: PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE,
            },
        },
    };

    // The blob points at `params` rather than holding it, so `params` outlives the activation —
    // which it does, being a local of this function and the activation being waited for inside it.
    let raw = windows_core::imp::PROPVARIANT {
        Anonymous: windows_core::imp::PROPVARIANT_0 {
            Anonymous: windows_core::imp::PROPVARIANT_0_0 {
                vt: VT_BLOB.0,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: windows_core::imp::PROPVARIANT_0_0_0 {
                    blob: windows_core::imp::BLOB {
                        cbSize: std::mem::size_of::<AUDIOCLIENT_ACTIVATION_PARAMS>() as u32,
                        pBlobData: &params as *const _ as *mut u8,
                    },
                },
            },
        },
    };
    // Safety: the variant is a blob of bytes pointing at a local that outlives it, and the wrapper
    // is the whole of what an activation parameter is passed as.
    //
    // **And it is never dropped.** A `PROPVARIANT` is cleared when it goes — the blob freed the way
    // the blob of an activation parameter is meant to be freed — and this blob is a local of this
    // function rather than anything `CoTaskMemAlloc` handed over, so the clearing would be a free of
    // the stack: the crash it causes is the heap's, and it is the next allocation anywhere that
    // takes the blame for it. It is held in a `ManuallyDrop` for that and leaks nothing: what it
    // points at is released with the frame it belongs to, which is what holds it alive until the
    // activation has read it.
    let property = ManuallyDrop::new(unsafe { PROPVARIANT::from_raw(raw) });

    let shared = Arc::new((Mutex::new(None::<usize>), Condvar::new()));
    let handler: IActivateAudioInterfaceCompletionHandler = Activated {
        address: Arc::clone(&shared),
    }
    .into();

    // Safety: the activation is handed the virtual device's own name, the interface this side
    // wants, a blob it may read for the length of the call, and a handler that outlives both the
    // call and the wait below — which is what the two locals are for.
    let operation = unsafe {
        ActivateAudioInterfaceAsync(
            VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
            &IAudioClient::IID,
            Some(&*property),
            &handler,
        )
    }
    .ok()?;

    // The answer arrives on a thread of the system's own rather than as this call's return, so what
    // is waited on here is the address that thread left: a bounded wait, because a bubble that is
    // never painted again is worse than bars drawn the old way.
    let (address, ready) = &*shared;
    let address = {
        let left = address.lock().ok()?;
        let (left, waited) = ready
            .wait_timeout_while(left, ACTIVATION_WAIT, |left| left.is_none())
            .ok()?;

        if waited.timed_out() {
            return None;
        }

        *left
    }?;

    // Both are dropped only now: the system holds its own reference to the handler until it has
    // called it, and the operation is what carries the answer.
    drop(operation);
    drop(handler);

    // Safety: the address is the one the completion handler left, and the reference it belongs to
    // was never released — it is this side's from here on.
    let client: IAudioClient = unsafe { Interface::from_raw(address as *mut std::ffi::c_void) };

    let format = WAVEFORMATEX {
        wFormatTag: WAVE_FORMAT_PCM as u16,
        nChannels: CHANNELS as u16,
        nSamplesPerSec: RATE,
        nAvgBytesPerSec: RATE * CHANNELS as u32 * 2,
        nBlockAlign: (CHANNELS * 2) as u16,
        wBitsPerSample: 16,
        cbSize: 0,
    };

    // Safety: the stream is a shared-mode loopback of this side's own process tree, initialized
    // once and never again, and the format outlives the call that reads it.
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                BUFFER,
                0,
                &format,
                None,
            )
            .ok()?;
        let capture = client.GetService::<IAudioCaptureClient>().ok()?;

        Some(Bands {
            client,
            capture,
            window: Vec::with_capacity(WINDOW_SAMPLES),
            shown: [0.0; BARS],
            asked: Instant::now(),
            last_packet: Instant::now(),
            running: false,
        })
    }
}

/// The activation's own callback, which the system calls on a thread of its own.
///
/// **What it hands the thread that asked for it is an address and never the interface**, and that
/// is a fact about an interface rather than a preference: one holds a pointer and is not `Send`, so
/// a callback on another thread cannot pass one on. An address can, and the reference it stands for
/// was never released — the thread that asked takes it back (see `open_capture`).
///
/// `IAgileObject` is named with the completion handler and is the free-threaded marker the sample
/// this call comes from names too (as `FtmBase`): the waiter is a thread that is not pumping its
/// own message queue, so a handler that had to be marshalled into the apartment it was made in
/// would be a callback nothing ever answers.
#[implement(IActivateAudioInterfaceCompletionHandler, IAgileObject)]
struct Activated {
    /// Where the answer is left: the address of the interface the system activated, or nothing
    /// where the activation refused.
    address: Arc<(Mutex<Option<usize>>, Condvar)>,
}

impl IActivateAudioInterfaceCompletionHandler_Impl for Activated_Impl {
    fn ActivateCompleted(
        &self,
        operation: Option<&IActivateAudioInterfaceAsyncOperation>,
    ) -> windows_core::Result<()> {
        let address = answered(operation);

        // Whatever the answer was, the thread waiting for it is woken: a refusal that left it to
        // the timeout would be a bubble painted from the peak a quarter of a second late.
        if let Ok(mut left) = self.address.0.lock() {
            *left = address;
            self.address.1.notify_all();
        }

        Ok(())
    }
}

impl IAgileObject_Impl for Activated_Impl {}

/// The interface an activation answered with, as the address of its own reference, or nothing
/// where the answer was a failure of any kind.
fn answered(operation: Option<&IActivateAudioInterfaceAsyncOperation>) -> Option<usize> {
    let mut result = HRESULT(0);
    let mut unknown: Option<windows_core::IUnknown> = None;

    // Safety: the operation is the one the system handed this callback, and the two out-parameters
    // are its own to write.
    unsafe {
        operation?
            .GetActivateResult(&mut result, &mut unknown)
            .ok()?;
    }
    result.ok().ok()?;

    // The cast is an interface of its own with a reference of its own, and the one it was cast from
    // is released as it goes out of scope: what is named here is one reference, which the waiting
    // thread takes over.
    let client: IAudioClient = unknown?.cast().ok()?;

    Some(Interface::into_raw(client) as usize)
}
