//! The level the sound on screen is being heard at: one number, read from the mixer the machine is
//! playing through.
//!
//! It is the meter on the default output endpoint rather than anything this app decodes, and that is
//! the whole of why it is here. A bubble standing for a sound is asked how loud the sound is *in the
//! room*, and the same file is played by FFmpeg's player on one machine and by the media engine on
//! another — so a level taken from this app's own player would be silent for a sound the engine
//! Windows has and wrong for one the user turned down in the mixer. The mixer's own peak is the one
//! answer that is right for both, and for the sound of anything else playing over them (see
//! `audio_preview` and `video_player`).
//!
//! One call and no transform: `IAudioMeterInformation::GetPeakValue` is the mixer's own peak over
//! the sound being heard now, which is exactly what a row of bars wants. A frequency analysis would
//! be a window of samples per repaint on the thread that draws the bubble, for a shape that at
//! forty-four pixels across reads the same either way (see `pin_chrome::draw_histogram`).
//!
//! Nothing here asks for an apartment of its own. The meter belongs to the thread that made it, so
//! it is cached per thread — and the one caller is the preview loop, whose thread has already taken
//! an apartment at the top of the loop by way of the two engines that need one (see
//! `pdf_preview::initialize_apartment` and the head of `run_preview_window`).

use super::*;

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
