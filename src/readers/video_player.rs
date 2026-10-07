//! The media engine Windows has, for the previews it plays.
//!
//! **Where this runs at all.** A video is played by this only on a machine with no `ffplay`
//! installed. Where FFmpeg's player *is* installed, that player plays every video, whatever its
//! name and whatever this engine could decode — because on that machine this engine is not the
//! fast path, it is the slow one: it decodes in software, and on a 1440p 144 fps HEVC file it
//! spends several hundred percent of a core to deliver frames a hardware decoder delivers for
//! nothing (see "Where the decoding happens" below). The engine is therefore the *fallback* for
//! video, and it is reached only when there is no fallback but it.
//!
//! A sound is a different question and keeps the order it always had: this plays a sound where
//! its own decoders reach the format, and `ffplay` plays it where they do not — the same question
//! asked of both, which is whether this machine has a decoder for the file (see [`plays`] and
//! [`audio_probe`]).
//!
//! What this buys a video, now that it is the fallback rather than the default, is the ergonomics
//! FFmpeg's player cannot report: a position that is real rather than a clock over a launch, a
//! pause that is a pause rather than a process ended. What it costs is the thing the measurements
//! below are about — speed, and hardware decoding — and that cost is what the ordering above
//! accepts on a machine that has FFmpeg at all. The one thing FFmpeg's player still cannot be
//! asked for is subtitles: it renders embedded and sidecar ones itself, while frame-server mode
//! returns video frames and nothing else, so a film shown here has no subtitles on it.
//!
//! The two are the same preview to the rest of the app — the same window, the same
//! placement, the same box — because what comes out of here is frames, drawn where every
//! other frame is drawn, rather than a player's window standing in for the preview. That
//! is the shape the media engine is asked for: it is created in *frame-server* mode,
//! which is what it is by default when no playback window is named, and then
//!
//!   * it decodes, paces itself, plays the audio and loops — none of which this app does;
//!   * this side asks, once a tick, whether a frame is due (`OnVideoStreamTick`) and takes it
//!     (`TransferVideoFrame`) into a bitmap of its own, from the crop the probe settled on
//!     where it settled on one (see [`Crop`]).
//!
//! # Where the decoding happens, and why it is not where it should be
//!
//! Every video this app plays is decoded in software, and the first thing to be clear about why
//! is that attaching a DXGI device manager cannot fix it on this machine. A media engine handed
//! one through `MF_MEDIA_ENGINE_DXGI_MANAGER` is documented to use the display's own hardware for
//! the decoding instead, and it does not: the media stack here has no hardware HEVC decoder
//! registered at all, so there is nothing for a manager to select.
//!
//! That makes the measurement this starts from a trap, and it is worth stating plainly because it
//! has been read the wrong way twice. Attaching a manager on a 144 fps 1440p HEVC file shown at
//! a preview box of a display's own size takes the engine from 44.02 s of a core over eight
//! seconds to 0.11 s — which reads as a four-hundred-fold win and is not one. The 0.11 s is an
//! idle engine: it is decoding nothing, because every transfer fails and no frame is ever
//! produced. The control settles it. Remove the device manager and the cost per frame is
//! identical; only the buffer the samples arrive in changes. So the manager was never selecting a
//! hardware decoder here, and the only thing it was selecting was a way for the engine to stop.
//!
//! That is why one is not attached, and not because it would be dearer to build: on this machine
//! the media engine will not then hand a frame over at all. `TransferVideoFrame` accepts a DXGI
//! surface *or* a WIC bitmap, a hardware-decoded frame lives in a texture on the GPU, and the
//! obvious reading is that the destination has to become a texture too — a buffer made by
//! `MFCreateDXGISurfaceBuffer` over a D3D11 texture, read back through a staging copy. That was
//! built, and it is refused. Measured, on this machine, over a 1440p HEVC file and a 1920x800
//! H.264 one:
//!
//!   * a WIC bitmap destination — the one this app already uses — answers the *first* transfer
//!     and every transfer after it with `E_NOINTERFACE` (`0x80004002`). The first frame is the one
//!     decoded before the DXGI pipeline is up; every frame after it is refused, whatever decoded
//!     it. That was true with the output format set to `ARGB32`, to `RGB32`, to `NV12` and unset
//!     entirely.
//!   * a `MFCreateDXGISurfaceBuffer` destination is refused on the first transfer as well as the
//!     rest, over a plain `D3D11_USAGE_DEFAULT` texture, over the same texture bound as a render
//!     target, over a surface from `IDXGIDevice::CreateSurface`, at the box's size and at the
//!     picture's own size, and over all of those again with the same four output formats.
//!
//! So the engine renders into a bitmap only while no manager is attached, and into nothing at all
//! once one is. Attaching a manager therefore is not a slower
//! preview: it is one frame, then three seconds of refusals, then the file handed to FFmpeg's
//! player (see [`TRANSFER_FAILURES_GIVE_UP`](playback::TRANSFER_FAILURES_GIVE_UP)) — which is a working preview on a machine that has
//! FFmpeg's player and no preview at all on a machine that does not.
//!
//! What is left of it, and the reason this paragraph is here rather than a branch somewhere, is
//! that the freeze is precisely the failure that used to be invisible: it was measured once as a
//! frame pipeline that sat on the first frame and reported nothing. Anyone reaching for this again
//! should read this first and [`failing_path`] second, because that is what stands between the
//! attempt and a preview that hangs.
//!
//! # The one other way a media engine is asked for frames, and why it is not this one
//!
//! There is a second way, and it was built and measured and thrown away, and what it found is
//! worth more than the code was. `IMFMediaEngineEx::EnableWindowlessSwapchainMode` has the engine
//! create a swap chain and present into it, and `GetVideoSwapchainHandle` hands that chain over
//! as an opaque handle; the handle is wrapped with
//! `IDXGIFactoryMedia::CreateSwapChainForCompositionSurfaceHandle`, `GetBuffer(0)` is the engine's
//! own back buffer, and it is read through a staging copy. No `TransferVideoFrame` anywhere in
//! it, so the refusal above is never asked — and the engine is still an `IMFMediaEngine`, which
//! is what makes it look like the answer to everything in this file.
//!
//! It is not, and the reason is narrower and duller than "Media Foundation does not work":
//!
//!   * The mode itself is accepted. So is the handle, the chain, `GetBuffer(0)` and the
//!     `CopyResource` out of it. Every step answers `S_OK`.
//!   * What answers is empty. `Map` on the staging copy succeeds and hands back a row pitch of
//!     *zero* and no address at all, over and over, for as long as the file plays. Nothing is
//!     ever written into the engine's back buffer: zero frames, zero non-zero bytes, and the
//!     clock standing still at zero while the engine believes it is playing.
//!   * The engine will only render into the chain if something is *compositing* the chain. This
//!     is the part that decides it, and it is not a bug in the probe: with no visual and no
//!     window attached, the swap chain has no consumer, so the engine has nothing to present
//!     into and writes nothing. Presenting the chain by hand — recycling a buffer and turning
//!     the rotation on, which puts no pixels anywhere and is not "showing" anything — is what a
//!     flip-model chain needs to advance at all, and it deadlocks the session on the second
//!     read rather than producing a frame.
//!
//! So the path needs a DirectComposition visual to live inside, which means the app's own
//! window and the compositor that draws it, which is a different and much larger project than
//! the one this file is. Nothing about the decoding is unreachable: what is unreachable is
//! *reading* a frame back out of a chain that nobody is displaying.
//!
//! Two of the things learned along the way are worth writing down because they cost the most and
//! are the sort of thing that would be learned again from scratch:
//!
//!   * A `DXGI_SWAP_CHAIN_DESC1` written with `..Default::default()` carries a
//!     `DXGI_SAMPLE_DESC` whose `Count` is **zero**, and every composition swap chain made over
//!     one is refused `DXGI_ERROR_INVALID_CALL` on this machine — every swap effect, every buffer
//!     count, every alpha mode, both formats, and with no media engine in the process at all.
//!     `Count` of one is what makes it, and the same is true of a `D3D11_TEXTURE2D_DESC` for the
//!     staging copy, which is refused `E_INVALIDARG` without it. C++ samples leave both at zero
//!     and get away with it, which is exactly why this reads as a Media Foundation problem and
//!     is not one.
//!   * `IMFDXGIDeviceManager` has no `SetDevice`. The only call that binds a device to one is
//!     `ResetDevice`, which is what Chromium's media foundation renderer does with the manager it
//!     locks; and the device has to be the engine's device rather than merely this side's, since
//!     a staging copy of a buffer on one device cannot be made with another's context.
//!
//! And the third, which is the one that cost the most time and is the reason this section
//! exists at all: `EnableWindowlessSwapchainMode` and `GetVideoSwapchainHandle` have to be asked
//! for *after* `Play`, not before `Load`, and the handle has to be polled rather than asked for
//! once. Before the engine is running the first is accepted and never acted on and the second
//! answers `S_OK` and a null handle. And `CreateSwapChainForCompositionSurfaceHandle` is a race:
//! asked for promptly it answers `S_OK`, and asked for a few hundred milliseconds later — the
//! engine having made its own chain over the same handle in the meantime — it answers
//! `DXGI_ERROR_ALREADY_EXISTS`. Which is the same finding as the one above, wearing a different
//! hat: there is always already a chain there, and it is not one anybody is showing.
//!
//! # The third way, which works and is still not the answer
//!
//! There is a third, and unlike the two above it does hand pixels back: an `IMFSourceReader`
//! handed an `IMFDXGIDeviceManager` through `MF_SOURCE_READER_D3D_MANAGER`, with advanced video
//! processing on, decoding the file through the same software HEVC decoder and answering
//! `ReadSample` with an `IMFDXGIBuffer` over a texture that a staging copy maps and reads. No
//! compositor is involved at any point and the pixels come back as decoded picture in system
//! memory, which is the thing the engine above would not do for anything. So the refusal is
//! narrower than it reads: it belongs to that engine rather than to Media Foundation, which is
//! worth having written down before anyone concludes the latter.
//!
//! Measured on this machine over eight seconds with the process's own clock, five runs of each
//! because the number moves a few milliseconds between them: 20.3–24.3 ms of CPU a frame to
//! decode alone, 25.3–31.2 ms a frame to decode and copy the staging texture back, against the
//! 128.99 ms a frame this side spends handing one over (345 frames in 8.01 s for 44.50 s of CPU,
//! the same instrument as `tests::video_take_cost`). Four to six times cheaper, which is a real
//! result and not a rounding error. It is also nowhere near the two milliseconds or so a hardware
//! decode of a 1440p picture would cost, and it could not be, because the decoder underneath is
//! the software one and no arrangement of attributes changes which decoder is registered. Only
//! `NV12` would negotiate as an output type at all — `BGRA` was refused `MF_E_INVALIDMEDIATYPE` at
//! every size and every frame rate tried — and the NV12 staging layout that had the most reason
//! to be trouble was not trouble at all.
//!
//! **So the frame pipeline is not being rewritten on a source reader**, and the measurements are
//! why: a source reader would hand this app a presentation clock, a WASAPI audio client and an
//! error channel that the engine gives away for free, and would buy four to six times on a
//! machine whose decoder is software whatever asks for it. The spike that measured all of this is
//! gone and the numbers are what is left of it, which is the arrangement the rest of this file
//! takes anyway. What it is worth is the last thing it found, and that is the way to read every
//! other measurement here: **a claim that hardware decode engaged is gated on CPU cost and never
//! on having been handed a GPU buffer.** A device manager wraps a software decoder's output in an
//! `IMFDXGIBuffer` as readily as a hardware one's, so the buffer says which path produced the frame
//! and says nothing whatever about what decoded it.
//!
//! A tick is not a frame. The clock this side ticks on is a vertical blank, sixty times a
//! second whatever the file runs at, so for a file at or below that rate most ticks find the
//! engine offering the picture it offered the last of them, and a tick that finds one is answered
//! without taking it at all (see [`is_a_new_frame`](session::is_a_new_frame)). A file above that rate is the other way
//! round: a 144 fps one is decoded at 144 fps and drawn at sixty, so most of what the engine
//! decodes is never shown at all — and on a machine whose decoder is the software one, which this
//! is, that is the largest single cost there is (see `tests::video_take_cost`). A 4K picture is
//! four times the frame of a 1080p one throughout, which is why that cost is where it is felt.
//!
//! Who *scales* the picture is settled by the two sizes alone: a box meaningfully larger than
//! the picture is one this side scales into it ([`scale_rows`](frame_copy::scale_rows)), because which filter the
//! engine reads a picture at is not a question this app can ask it, let alone choose. What each
//! side of that choice costs is written where it is made (see [`play`]).
//!
//! What that buys over the ffplay path is the whole of the window machinery a player's
//! own window needs — the style monitor, the topmost re-assertion, the PID record, the
//! job object — none of which exists here, because there is no second window and no
//! second process. What it costs is that the frames are copied once per frame.
//!
//! The engine is what decides which files it can play: whatever Windows 11 decodes out of
//! the box, plus whatever a codec extension from the Microsoft Store has added — which is
//! the question the tray's `Codecs` submenu answers for the machine it is running on.
//!
//! The questions asked of the engine from outside are all answered here. [`dimensions`]
//! is a video's shape, asked of a source reader before there is anything to play. [`plays`]
//! is whether a decoder here can hand a video's own frames back at all — what the router
//! asks before it hands a file to this engine or to FFmpeg's player.
//! [`audio_track`](crate::readers::audio_track) is the same question about a sound, asked
//! the same way — the reader is asked for *decoded* PCM, which it can only give where the
//! decoder is registered, and the media type the file declares is read for the rest. And
//! [`position`] and [`duration`] are what the running session says about itself, for the
//! card's clock and a pinned window's transport bar.

mod frame_copy;
mod playback;
mod session;

pub use frame_copy::audio_probe;
pub(crate) use frame_copy::{force_opaque, open_stream, unpack_pair};
pub use playback::{
    apply_seek, audio_ended, copy_frame_into, dimensions, duration, failing_before_a_frame,
    failing_path, is_playing, mark_unplayable, play, play_audio, playing_path, plays, position,
    resize, seek, set_paused, set_volume, stop, Crop, Picture,
};

// Reached from the preview window's own tests rather than from the app, so it is only here
// where they are compiled.
#[cfg(test)]
pub use playback::can_play;

#[cfg(test)]
use frame_copy::{
    copy_row_opaque, interpolate_row, rows_read_per_destination_row, scale_rows, A_WHOLE_PIXEL,
};
#[cfg(test)]
use playback::{
    a_run_of_refusals_is_a_failure, a_sound_that_has_played_to_its_end, engine_is_failing,
    scales_here, Notify, FIRST_FRAME_GIVE_UP, TRANSFER_FAILURES_GIVE_UP,
};
#[cfg(test)]
use session::is_a_new_frame;

// What the tests beside this module reach with `use super::*`: the path the engine resolves and
// the types and constants only they ask about.
#[cfg(test)]
use crate::paths::plain_path;
#[cfg(test)]
use std::path::{Path, PathBuf};
#[cfg(test)]
use std::sync::atomic::{AtomicBool, Ordering};
#[cfg(test)]
use std::sync::Arc;
#[cfg(test)]
use std::time::{Duration, Instant};
#[cfg(test)]
use windows::Win32::Media::MediaFoundation::{
    IMFMediaEngineNotify_Impl, MF_MEDIA_ENGINE_EVENT_ENDED, MF_MEDIA_ENGINE_EVENT_ERROR,
    MF_MEDIA_ENGINE_EVENT_FIRSTFRAMEREADY,
};

#[cfg(test)]
mod tests;
