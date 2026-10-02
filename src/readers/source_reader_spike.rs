//! A source reader and a DXGI device manager, asked the one question the media engine cannot.
//!
//! **THROWAWAY SPIKE.** Nothing in here is wired to the app, and none of it is meant to survive
//! whatever the answer turns out to be. It exists because two attempts to get hardware-decoded
//! pixels out of the media engine failed in ways that both ended at the same wall, and the next
//! step in the frame pipeline is somebody rewriting it — which is not a thing to do on the
//! strength of a hunch. See [`video_player`]'s module documentation for those two attempts; the
//! short version is that the engine decodes on the GPU happily (process CPU over an eight-second
//! window fell from 44.50 s to 0.11 s) and then refuses to hand a single frame to anything that
//! is not a compositor: `TransferVideoFrame` answers `E_NOINTERFACE` on every destination tried,
//! and a windowless swap chain's back buffer stays empty forever.
//!
//! A source reader is not that. `IMFSourceReader::ReadSample` hands back an `IMFMediaBuffer` with
//! no compositor anywhere in it, and when an `IMFDXGIDeviceManager` has been handed to the reader
//! through `MF_SOURCE_READER_D3D_MANAGER` that buffer is `IMFDXGIBuffer`-backed — the decoder
//! owns a texture and the reader gives it over. So the question this file answers is narrow and
//! measurable, and nothing else:
//!
//! > Can `IMFSourceReader` plus a DXGI device manager hand this application **CPU-readable BGRA
//! > pixels** from a 1440p HEVC file, at a decode cost low enough to matter?
//!
//! "Low enough to matter" is one number, and it is the number the comparison was taken with. The
//! software path in `video_player` — a media engine with no manager on it — measured on the same
//! machine, through the same instrument, at **128.99 ms of process CPU per frame handed over**
//! (345 frames in 8.01 s for 44.50 s of CPU). Anything under about 2 ms per frame is a different
//! kind of thing from that, and it is rate-independent, so it does not matter whether a reader
//! delivers 60 frames a second or 144.
//!
//! Two things are kept apart on purpose, because conflating them is how a measurement like this
//! lies:
//!
//!   * **Frames read**, and **CPU per frame**. Those say whether the route is cheap. They are
//!     measured over a fixed window with the process's own clock, exactly as
//!     `video_player::tests::video_take_cost` measures the thing this is a candidate to replace.
//!   * **Which buffer the reader actually handed over**. This says whether hardware decode ran.
//!     `S_OK` off `ReadSample` does not: a software decoder produces an `IMF2DBuffer` in system
//!     memory and a hardware decoder produces an `IMFDXGIBuffer` over a texture, and there is no
//!     third answer. The count of each is reported.
//!
//! `S_OK` is not believed anywhere in here. Every frame that is read is checked for actually
//! containing pixels — the sampled bytes' mean and a signature over them — and a run that reads
//! frames of zeroes is reported as the failure it is rather than as a pass.
//!
//! Run it by hand, in release, because decode cost is meaningless in a debug build:
//!
//! ```text
//! cargo test --release source_reader_spike -- --ignored --nocapture
//! ```
//!
//! The box size is 2493 x 1400 unless `RHP_SPIKE_BOX` says `WIDTHxHEIGHT`, which is the display's
//! own size and the size `video_take_cost` was measured at.
//!
//! # Nothing here is called by the application
//!
//! Every function below is reachable only from this module's own `#[ignore]`d test. That is why
//! the module carries `#![allow(dead_code)]`: a non-test build genuinely has nothing to call any
//! of this, and the spike is not going to be wired into a preview. The allowance is in one place
//! rather than scattered over twenty items, so that the file stops carrying annotations the day it
//! is deleted.
//!
//! # What came of it
//!
//! **The route works and it is not worth anything, because this machine cannot hardware-decode
//! HEVC.** Read that in two halves, because they are separate findings and only the first one is
//! good news.
//!
//! **Half one: the source reader does hand over CPU-readable pixels, and no compositor is
//! involved.** This is the thing the two dead attempts were reaching for and it is real. The
//! frames arrive as `IMFDXGIBuffer`, `GetResource` gives a genuine `ID3D11Texture2D` over
//! `DXGI_FORMAT_NV12` on the device manager's own device, a `USAGE_STAGING` copy of it maps and
//! reads, and the bytes are decoded picture: 543 frames, every one with a nonzero sampled mean
//! (~110.8), 542 distinct signatures out of 543, none all zero. Nothing here needed a window.
//!
//! **Half two: the decode was in software.** Three independent pieces of evidence, and the
//! third is the one that settles it:
//!
//!   * This machine has **no hardware HEVC decoder**. The decoder registry below lists
//!     `HEVCVideoExtension` — and its hardware URL is `false`. The only video decoder on the
//!     machine that claims hardware is `AMD D3D11 Hardware MFT Playback Decoder`, and asking for
//!     hardware decoders that accept HEVC as *input* returns zero, for HEVC, for H.264 and for
//!     AV1 alike.
//!   * `IMFDXGIBuffer` is therefore not evidence of a hardware decode here. It is evidence that
//!     the source reader wrapped a frame in a GPU surface, which it will do for a software
//!     decoder's output too once a device manager and advanced video processing are in the
//!     pipeline.
//!   * **With the device manager removed the cost does not change.** `RHP_SPIKE_MANAGER=0`
//!     answers `IMF2DBuffer` in system memory at **20.26 ms** a frame; with it, a texture on the
//!     GPU at **21.9–24.3 ms** a frame. A hardware decode of 1440p HEVC is on the order of a tenth
//!     of a millisecond. Both numbers are software decode, and the DXGI buffer is a surface the
//!     decoded bytes were copied into.
//!
//! Measured, on this file, at 60 fps paced, over eight seconds, with the process's own clock.
//! Five runs, because the number moves by a few milliseconds between them and quoting one of them
//! as *the* number would be the same kind of mistake as quoting `S_OK` as a frame:
//!
//! ```text
//! decode only, frames taken and dropped:  536-556 frames, 20.3-24.3 ms of CPU each
//! decode + read back into system memory: 543-558 frames, 25.3-31.2 ms of CPU each
//! with the device manager removed:        556 frames,    20.3 ms of CPU each, IMF2DBuffer
//! ```
//!
//! The last line is the one that matters, and it is the control: **removing the DXGI device
//! manager changes what the frame is and does not change what it costs.** That is the measurement
//! that says the surface is not the decode.
//!
//! Against the **128.99 ms** a frame the software path in `video_player` costs, this is 4x to 6x
//! cheaper — a real and somewhat surprising result for a source reader against a media engine, and
//! most likely the engine's own per-frame overhead rather than anything about hardware. It is
//! nowhere near the **under 2 ms** the gate asks for, and no arrangement of attributes will get
//! there on a machine whose HEVC decoder is the software one.
//!
//! **So the gate fails**, on two counts: CPU per frame is an order of magnitude over the budget,
//! and hardware decode could not be shown to be engaged because it is not engaged.
//!
//! What this means for the frame pipeline, which is the decision this spike was for: **do not
//! rewrite it on the strength of a source reader.** A source reader is not the answer this machine
//! needs. What would be worth knowing is whether the machine that *does* have a hardware HEVC
//! decoder behaves like the media engine did — and the honest answer is that the premise the whole
//! line of enquiry rests on has never been established on this machine.
//!
//! # The premise the whole line of enquiry rests on, which is false
//!
//! The brief for this spike, and the reasoning behind the two dead attempts, is the measured claim
//! that on this machine "hardware decode genuinely engages — process CPU for an 8-second window of
//! the test file fell from **44.50 s (555.9% of a core) to 0.11 s (1.4%)**" when a DXGI device
//! manager is attached to the media engine.
//!
//! **0.11 s of CPU over eight seconds is what an idle engine costs, not what decoding costs.**
//! Those two experiments are the ones in `video_player`'s module documentation, and in both of
//! them a device manager was attached and then *no frames were handed over at all*: every
//! `TransferVideoFrame` answered `E_NOINTERFACE` from the first frame onwards, and the windowless
//! swap chain's back buffer was empty for the whole eight seconds. An engine that decoded one
//! frame and then spent eight seconds refusing to hand over any more costs almost nothing, and the
//! 0.11 s is a measurement of the refusals.
//!
//! It is consistent with this spike's own finding, which is that there is no hardware HEVC decoder
//! here for an engine to use either. Nothing in this file disproves hardware decoding on Windows
//! in general — but on *this machine*, for *this file*, it is not happening, and the 44.50 s to
//! 0.11 s figure should not be cited as evidence that it is.
//!
//! # What had to be different from the documented recipe
//!
//! Four things, three of which contradict the recipe or the research that came with it:
//!
//!   1. **BGRA does not negotiate at all.** `SetCurrentMediaType` was refused
//!     `MF_E_INVALIDMEDIATYPE` (`0xC00D36B4`) for BGRA at the box size at 60/1, at the box size at
//!     the file's own rate, at the file's own size at the file's own rate, and at the file's own
//!     size with no frame rate stated at all. Only **NV12** was accepted. The recipe's preferred
//!     outcome — BGRA via XVP — is not available here, and the research's warning about NV12
//!     staging-texture layout being driver-restricted did **not** materialise: the NV12 texture
//!     copied and mapped without complaint.
//!   2. **`IMFActivate` is its own `IMFAttributes`.** Reading an MFT's friendly name and hardware
//!     URL through `ActivateObject`/`IMFTransform::GetAttributes` fails on every one of them,
//!     because nothing is activated yet. `MFTEnumEx` returns *factories*, and a factory's
//!     attributes are the attributes.
//!   3. **`MFTEnumEx` with an input-type filter and `MFT_ENUM_FLAG_HARDWARE` returns an empty
//!     list on this machine**, for HEVC and for H.264 and for AV1 alike. An enumeration that
//!     answers zero where the answer is "one software and no hardware" is not evidence of
//!     anything, and reading it as "no hardware decoder" would have been a wrong conclusion drawn
//!     from a right observation.
//!   4. **`IID_ID3D11Texture2D` is not in the `windows` crate**, and neither is
//!     `MFVideoFormat_BGRA`. Both are written out here as `GUID` constants read out of
//!     `d3d11.h` and `mfapi.h` on this machine rather than from memory, which matters: the first
//!     attempt got the texture IID wrong and `GetResource` answered `E_NOINTERFACE` on all 965
//!     frames — **the same HRESULT the media engine's `TransferVideoFrame` gives, from an
//!     unrelated cause.** Anyone who sees `E_NOINTERFACE` from this file should check the
//!     constant before concluding anything about Media Foundation.
//!
//! The DXGI descriptor traps the previous two attempts documented did not reproduce: every field of
//! the staging texture is written explicitly, `SampleDesc.Count` is `1`, and `CopyResource` was
//! accepted on the first try against a texture this file had not created.

use std::collections::BTreeSet;
use std::path::Path;
use std::time::{Duration, Instant};

use windows::core::{Interface, GUID, PWSTR};
use windows::Win32::Foundation::{FILETIME, HMODULE};
use windows::Win32::Graphics::Direct3D::{D3D_FEATURE_LEVEL, D3D_DRIVER_TYPE_HARDWARE};
use windows::Win32::Graphics::Direct3D11::{
    D3D11CreateDevice, D3D11_CPU_ACCESS_READ, D3D11_CREATE_DEVICE_BGRA_SUPPORT,
    D3D11_CREATE_DEVICE_VIDEO_SUPPORT, D3D11_MAP_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_SDK_VERSION,
    D3D11_TEXTURE2D_DESC, D3D11_USAGE_STAGING, ID3D11Device, ID3D11DeviceContext, ID3D11Texture2D,
};
use windows::Win32::Graphics::Dxgi::Common::{DXGI_FORMAT, DXGI_SAMPLE_DESC};
use windows::Win32::Graphics::Dxgi::IDXGIDevice;
use windows::Win32::Media::MediaFoundation::{
    IMF2DBuffer, IMFActivate, IMFAttributes, IMFDXGIDeviceManager, IMFDXGIBuffer, IMFMediaBuffer,
    IMFMediaType, IMFSourceReader, MFCreateAttributes, MFCreateDXGIDeviceManager,
    MFCreateMediaType, MFCreateSourceReaderFromByteStream, MFMediaType_Video, MFTEnumEx,
    MFVideoFormat_AV1, MFVideoFormat_H264, MFVideoFormat_HEVC, MFVideoFormat_NV12,
    MFVideoInterlace_Progressive, MF_MT_FRAME_RATE,
    MF_MT_FRAME_SIZE, MF_SOURCE_READERF_ENDOFSTREAM,
    MF_MT_INTERLACE_MODE, MF_MT_MAJOR_TYPE, MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SUBTYPE,
    MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS, MF_SOURCE_READER_D3D_MANAGER,
    MF_SOURCE_READER_DISABLE_DXVA, MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING,
    MF_SOURCE_READER_FIRST_VIDEO_STREAM, MFT_CATEGORY_VIDEO_DECODER, MFT_ENUM_FLAG,
    MFT_ENUM_FLAG_ALL, MFT_ENUM_FLAG_HARDWARE, MFT_ENUM_FLAG_SORTANDFILTER,
    MFT_ENUM_HARDWARE_URL_Attribute, MFT_FRIENDLY_NAME_Attribute, MFT_REGISTER_TYPE_INFO,
};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::Threading::{GetCurrentProcess, GetProcessTimes};

use crate::formats::codecs;
use crate::readers::video_player::{open_stream, unpack_pair};

/// `MFVideoFormat_BGRA`, which `windows` 0.58 does not carry.
///
/// Every other `MFVideoFormat_*` the spike needs is in the crate, and this one is missing from
/// it — which is a data constant rather than a missing binding, so it is written out here rather
/// than being treated as the stop condition. Nothing in this file is declared `extern "system"`:
/// this is a `GUID` built from its own bytes, the same thing `shell::explorer_hook` does for the
/// format table it prints into.
///
/// Which is moot, in the end: this machine refuses BGRA from a source reader at every size and
/// every frame rate, and hands back NV12 instead. See the module documentation.
/// <https://learn.microsoft.com/en-us/windows/win32/api/mfapi/ne-mfapi-mfvideoformat>
const MF_VIDEO_FORMAT_BGRA: GUID = GUID::from_u128(0x0000000d_0000_0010_8000_00aa00389b71);

/// `IID_ID3D11Texture2D`, which `windows` 0.58 does not carry either, and which
/// `IMFDXGIBuffer::GetResource` cannot be asked for a texture without.
///
/// The crate's `define_interface!` macro generates no `IID_*` constants, and `GetResource` takes
/// the interface identifier as a raw pointer rather than a generic, so there is no way to reach
/// the texture through the typed wrappers. Same reasoning as [`MF_VIDEO_FORMAT_BGRA`]: a value,
/// not a declaration.
///
/// The value is read out of `d3d11.h` on this machine rather than written from memory, because
/// writing it from memory got it wrong: `GetResource` then answered `E_NOINTERFACE` on all 965
/// frames, which is the same HRESULT the media engine's `TransferVideoFrame` gives and would have
/// looked like a repeat of the same failure. It is cross-checked twice — against
/// `MIDL_INTERFACE("6f15aaf2-d208-4e89-9ab4-489535d34f9c")` in the header, and against the
/// crate's own `define_interface!(ID3D11Texture2D, …, 0x6f15aaf2_d208_4e89_9ab4_489535d34f9c)`.
/// <https://learn.microsoft.com/en-us/windows/win32/api/d3d11/nn-d3d11-id3d11texture2d>
const IID_ID3D11_TEXTURE2D: GUID = GUID::from_u128(0x6f15aaf2_d208_4e89_9ab4_489535d34f9c);

/// How long the frames are read for once the pipeline is up.
///
/// Eight seconds, which is the window `video_player::tests::video_take_cost` measured the
/// software path over, because a number is only comparable with another number taken over the
/// same window on the same machine.
const WINDOW: Duration = Duration::from_secs(8);

/// How long is spent reading before the clock is started.
///
/// A first frame's decoder and a hardware pipeline's device are both made in it, and a run whose
/// clock started at the first `ReadSample` would be measuring setup — which is the mistake
/// `video_take_cost` already makes deliberately not.
const SETTLING: Duration = Duration::from_secs(1);

/// How far apart the bytes sampled out of a frame are.
///
/// A frame at the box size is about fourteen million bytes, and reading all of it on the CPU to
/// prove it is not a field of zeroes would put the measurement's own cost inside the number being
/// measured — at a two-millisecond budget that matters. Fourteen thousand samples spread across
/// the frame is enough to say the picture is real and to tell one frame from the next, and costs
/// about ten microseconds.
const SAMPLE_STRIDE: usize = 997;

/// The frame rate the output type asks the reader for.
///
/// Sixty, which is the refresh of the display the preview is drawn on and the rate `video_player`
/// draws at; the file itself runs at 143.9. Asking for less than the file has is the point — a
/// reader asked for 60 will not hand over frames nobody is going to see — but it is also the one
/// attribute that can put a *frame rate converter* in the pipeline rather than only a colour
/// converter, so the frame count is printed against what went in.
///
/// It is a candidate rather than the only one, because `SetCurrentMediaType` refused `60/1`
/// against this file with `MF_E_INVALIDMEDIATYPE` — and whether the refusal was about the frame
/// rate, the size or the format is exactly the question the spike exists to answer. Each is tried
/// and its answer printed; see [`negotiate`].
const PREFERRED_FRAME_RATE: (u32, u32) = (60, 1);

/// What a spike run found, in the shape the test prints and the gate is read off.
///
/// Only the three numbers the question is actually about, so that the shape of the answer cannot
/// quietly grow a field nobody looks at.
pub(crate) struct Summary {
    /// Frames the reader handed over inside the measured window.
    pub frames: u64,
    /// Milliseconds of this process's own CPU per frame handed over: the number that matters.
    pub cpu_per_frame_ms: f64,
    /// Which of the two buffers came back, which is the hardware-decode evidence.
    pub buffer_kind: &'static str,
    /// Frames that carried something other than zeroes, and frames that differed from the one
    /// before them — the "the pixels are decoded and they are moving" half of the gate.
    pub frames_with_pixels: u64,
    /// Frames whose sampled bytes differed from the frame immediately before.
    pub frames_changed: u64,
}

/// The spike, run against one file at one box size. Every failure path returns `None` and says
/// what it was on the way out; nothing in here can panic, because a measurement that panics is a
/// measurement with no number in it.
pub(crate) fn measure(path: &Path, width: u32, height: u32) -> Option<Summary> {
    if !codecs::mf_started() {
        println!("Media Foundation will not start on this machine — nothing to ask of it");
        return None;
    }

    // The two facts about this machine that have to hold before the reading means anything, and
    // which are worth printing whether or not they do: is there a hardware HEVC decoder at all,
    // and can a video device be made here. A machine with neither answers "zero frames" and says
    // nothing about the route.
    // What this machine has, before anything is asked of it, and the whole decoder list rather than
// only the HEVC part of it.
//
// The first version asked for HEVC hardware decoders and was answered "zero", on a machine whose
// media engine had demonstrably hardware-decoded the same file — so either the enumeration is
// lying or the hardware decoder is registered somewhere this is not looking. Both are worth
// knowing before a number is read off it, so the whole category is listed and each entry says
// whether it claims to be hardware.
    println!("--- video decoders this machine has registered ---");
    for name in all_video_decoders() {
        println!("  {name}");
    }
    for (label, subtype) in [
        ("H.264", MFVideoFormat_H264),
        ("HEVC", MFVideoFormat_HEVC),
        ("AV1", MFVideoFormat_AV1),
    ] {
        let hardware = hardware_decoders(&subtype, MFT_ENUM_FLAG_HARDWARE).unwrap_or_default();
        let sorted = hardware_decoders(
            &subtype,
            MFT_ENUM_FLAG_HARDWARE | MFT_ENUM_FLAG_SORTANDFILTER,
        )
        .unwrap_or_default();
        println!(
            "  {label}: {} hardware decoders by input-type filter, {} with sorting on",
            hardware.len(),
            sorted.len()
        );
    }

    let Some((device, context, feature_level)) = hardware_device() else {
        return None;
    };
    println!(
        "hardware D3D11 device created at feature level 0x{:x}",
        feature_level.0
    );

    // The manager, and then `ResetDevice` on it before anything at all is asked of it. There is no
    // `SetDevice`: `ResetDevice` is the only call that binds a device to a manager, and a manager
    // handed to a reader with no device behind it refuses every stream rather than falling back
    // to software — which would read as a working answer for the wrong reason.
    let Some((manager, reset_token)) = device_manager(&device) else {
        return None;
    };
    println!("device manager reset on token {reset_token} — device bound before the reader exists");

    let byte_stream = open_stream(path)?;
    let mut attributes: Option<IMFAttributes> = None;
    // SAFETY: `attributes` is the caller's own out-parameter, initialised to `None`, and MF is
    // being asked for the default number of elements rather than for a count of its own.
    if let Err(error) = unsafe { MFCreateAttributes(&mut attributes, 4) } {
        println!("MFCreateAttributes refused: {error}");
        return None;
    }
    let attributes = attributes?;

    // The attributes that turn the reader into a hardware pipeline.
    // `MF_SOURCE_READER_D3D_MANAGER` is the one that matters — it is what the source reader hands
    // the decoder as its `MFT_MESSAGE_SET_D3D_MANAGER`, and without it the decoder has no device
    // to decode into.
    //
    // The other three are switches rather than instructions, because a spike that can only be run
    // one way cannot answer "which of these is responsible" when it answers badly. Each is read
    // from the environment so that one build answers the whole matrix; what the gate asked for —
    // all four on — is what the run in this file's own documentation used.
    let (want_xvp, want_hardware_transforms, want_manager) = switches();

    println!(
        "attributes: D3D manager {want_manager}, hardware transforms {want_hardware_transforms}, XVP {want_xvp}"
    );

    // SAFETY: each call writes a value into the attribute store rather than through a pointer of
    // ours; the keys are the `windows` crate's own constants and the one value that is not a
    // scalar, the device manager, is a live binding passed by reference.
    let set = unsafe {
        [
            want_manager.then(|| attributes.SetUnknown(&MF_SOURCE_READER_D3D_MANAGER, &manager)),
            Some(attributes.SetUINT32(
                &MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS,
                u32::from(want_hardware_transforms),
            )),
            // `MF_SOURCE_READER_DISABLE_DXVA` is set to FALSE rather than left alone, because the
            // recipe asks for it explicitly and an attribute the app means to set has no business
            // being inferred from a default.
            Some(attributes.SetUINT32(&MF_SOURCE_READER_DISABLE_DXVA, 0)),
            want_xvp.then(|| {
                attributes.SetUINT32(&MF_SOURCE_READER_ENABLE_ADVANCED_VIDEO_PROCESSING, 1)
            }),
        ]
    };

    for outcome in set {
        if let Some(Err(error)) = outcome {
            println!("a reader attribute was refused: {error}");
            return None;
        }
    }

    let Ok(reader) =
        (unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, Some(&attributes)) })
    else {
        println!("MFCreateSourceReaderFromByteStream refused with the manager attached");
        return None;
    };

    let video_stream = MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32;
    let Ok(native) = (unsafe { reader.GetNativeMediaType(video_stream, 0) }) else {
        println!("the first video stream has no native media type at all");
        return None;
    };
    print_media_type("native", native.clone());

    // The native frame rate is read off the file rather than assumed, because "the output type
    // was refused" and "the output type was refused *because* its frame rate was not the file's"
    // are different findings and the ladder below needs to be able to tell them apart.
    let native_rate = unsafe { native.GetUINT64(&MF_MT_FRAME_RATE) }
        .ok()
        .map(unpack_pair)
        .unwrap_or(PREFERRED_FRAME_RATE);
    let native_size = unsafe { native.GetUINT64(&MF_MT_FRAME_SIZE) }
        .ok()
        .map(unpack_pair)
        .unwrap_or((width, height));

    if negotiate(
        &reader,
        video_stream,
        (width, height),
        native_size,
        native_rate,
    )
    .is_none()
    {
        println!("\nno output type at all was accepted — nothing to read frames with");
        return None;
    }

    if let Ok(negotiated) = unsafe { reader.GetCurrentMediaType(video_stream) } {
        print_media_type("negotiated", negotiated);
    }

    // Two measurements, because "CPU per frame" is two different numbers and reporting only the
    // larger one would say the route is 26 ms a frame when the decoder is not.
    //
    // The first takes each frame, identifies what kind of buffer it is, and drops it: that is the
    // decoder and the source reader and nothing else. The second then reads the same frames back
    // through a staging copy, which is what this application would have to do to draw one, and is
    // a cost proportional to the picture's size rather than to the decode.
    let decode_only = measure_window(&reader, video_stream, &context, false);
    let with_readback = measure_window(&reader, video_stream, &context, true);

    let Some(full) = with_readback else {
        return None;
    };
    if full.frames == 0 {
        println!("\nno frames at all — the route does not work on this file");
        return None;
    }

    // Everything the report needs is lifted out of the two windows by value first, so that the
    // printing below is a read of plain numbers and cannot be got wrong by a moved-from field.
    let report = Report {
        decode_only: decode_only.map(|w| (w.frames, w.cpu_per_frame_ms)),
        frames: full.frames,
        fps: full.frames as f64 / full.elapsed.as_secs_f64(),
        cpu_per_frame_ms: full.cpu_per_frame_ms,
        buffer_kind: full.run.buffer_kind(),
        dxgi_frames: full.run.dxgi_frames,
        system_frames: full.run.system_frames,
        frames_with_pixels: full.run.frames_with_pixels,
        frames_all_zero: full.run.frames_all_zero,
        frames_changed: full.run.frames_changed,
        distinct: full.run.distinct.len(),
        mean_of_means: full.run.mean_of_means,
        texture_format: full.run.texture_format.map(|f| f.0),
        refusals: full.run.refusals.clone(),
    };

    report.print(native_rate);

    Some(Summary {
        frames: report.frames,
        cpu_per_frame_ms: report.cpu_per_frame_ms,
        buffer_kind: report.buffer_kind,
        frames_with_pixels: report.frames_with_pixels,
        frames_changed: report.frames_changed,
    })
}

/// What a run of the spike found, as plain numbers, so that the printing of them is separate from
/// the collecting of them.
struct Report {
    /// Frames and CPU-per-frame for the window that took each frame and dropped it.
    decode_only: Option<(u64, f64)>,
    /// Frames and CPU-per-frame for the window that read every frame back into system memory.
    frames: u64,
    fps: f64,
    cpu_per_frame_ms: f64,
    buffer_kind: &'static str,
    dxgi_frames: u64,
    system_frames: u64,
    frames_with_pixels: u64,
    frames_all_zero: u64,
    frames_changed: u64,
    distinct: usize,
    mean_of_means: f64,
    texture_format: Option<i32>,
    refusals: Vec<String>,
}

impl Report {
    fn print(&self, native_rate: (u32, u32)) {
        println!("\n=== the reading ===");
        if let Some((frames, each)) = self.decode_only {
            println!(
                "decode only, frames taken and dropped: {frames} frames, {:.3} ms of CPU each",
                each
            );
        }
        println!(
            "decode + read back into system memory:  {} frames, {:.3} ms of CPU each ({:.1} fps)",
            self.frames, self.cpu_per_frame_ms, self.fps
        );
        println!(
            "buffer type: {} — {} on a texture (IMFDXGIBuffer), {} in system memory (IMF2DBuffer)",
            self.buffer_kind, self.dxgi_frames, self.system_frames
        );
        println!(
            "pixels: {} frames carried bytes, {} were all zero, {} differed from the frame before",
            self.frames_with_pixels, self.frames_all_zero, self.frames_changed
        );
        println!(
            "        {} distinct signatures out of {} read, mean sampled byte {:.1}",
            self.distinct, self.frames, self.mean_of_means
        );
        if let Some(format) = self.texture_format {
            println!(
                "        texture format {format} (103 is DXGI_FORMAT_NV12, 87 is DXGI_FORMAT_B8G8R8A8_UNORM)"
            );
        }
        println!(
            "        delivered at {:.1} fps against the file's own {:.1}",
            self.fps,
            (native_rate.0 as f64 / native_rate.1 as f64).max(1.0),
        );
        println!(
            "\nthe software path in video_player, same machine and instrument: 128.99 ms per frame"
        );
        for (label, each) in [
            ("decode only ", self.decode_only.map(|(_, e)| e)),
            ("with readback", Some(self.cpu_per_frame_ms)),
        ] {
            if let Some(each) = each {
                if each > 0.0 {
                    println!("  {label}: {each:.3} ms/frame — {:.0}x cheaper", 128.99 / each);
                }
            }
        }

        for refusal in self.refusals.iter().take(6) {
            println!("refused along the way: {refusal}");
        }
    }
}

/// One measured window: a settling window nobody counts, then the window itself with the process's
/// own clock read either side of it.
struct Window {
    frames: u64,
    elapsed: Duration,
    cpu_per_frame_ms: f64,
    run: Run,
}

/// The window, measured. `readback` says whether frames are actually read out of their GPU
/// resource, which is what separates the decoder's cost from this application's.
fn measure_window(
    reader: &IMFSourceReader,
    video_stream: u32,
    context: &ID3D11DeviceContext,
    readback: bool,
) -> Option<Window> {
    let mut run = Run::new(context, readback);
    if let Some(interval) = run.frame_interval {
        println!(
            "paced at {:.0} fps (RHP_SPIKE_FPS)",
            1.0 / interval.as_secs_f64()
        );
    }
    run.read(reader, video_stream, SETTLING);

    let before = cpu_hundred_nanoseconds()?;
    let started = Instant::now();
    run.read(reader, video_stream, WINDOW);
    let elapsed = started.elapsed();
    let after = cpu_hundred_nanoseconds()?;

    let cpu_ms = (after.saturating_sub(before)) as f64 / 10_000.0;
    let cpu_per_frame_ms = if run.frames == 0 {
        f64::NAN
    } else {
        cpu_ms / run.frames as f64
    };

    let frames = run.frames;
    Some(Window {
        frames,
        elapsed,
        cpu_per_frame_ms,
        run,
    })
}

/// What a media type says, read off the reader rather than guessed: the subtype is what makes the
/// hardware-decoder count above the right question rather than a general one, and the frame rate
/// is what the output rate is a conversion away from.
fn print_media_type(label: &str, media_type: IMFMediaType) {
    let subtype = unsafe { media_type.GetGUID(&MF_MT_SUBTYPE) }.ok();
    let size = unsafe { media_type.GetUINT64(&MF_MT_FRAME_SIZE) }.ok().map(unpack_pair);
    let rate = unsafe { media_type.GetUINT64(&MF_MT_FRAME_RATE) }.ok().map(unpack_pair);

    println!(
        "{label}: subtype {:?}, {}x{}, {} / {} fps",
        subtype.map(|g| format!("{g:?}")).unwrap_or_else(|| "none".into()),
        size.map(|s| s.0).unwrap_or(0),
        size.map(|s| s.1).unwrap_or(0),
        rate.map(|r| r.0).unwrap_or(0),
        rate.map(|r| r.1).unwrap_or(1),
    );
}

/// What one run of the read loop found, and the only place a frame is judged.
///
/// It is a struct rather than a set of locals because a frame that came back has to be counted
/// and signed and meaned, and every one of those is a place the answer could quietly not be
/// recorded.
struct Run {
    context: ID3D11DeviceContext,

    /// Frames read inside the measured window.
    frames: u64,

    /// How many of those arrived as a texture on the device manager's device, which is what a
    /// hardware decoder produces and the only thing in this file that proves one was used.
    dxgi_frames: u64,

    /// How many arrived in system memory, which is what a software decoder produces.
    system_frames: u64,

    /// Whether the GPU-to-CPU staging copy and the read through it are actually done, which is
    /// the only way to tell "what the decoder cost" apart from "what this file's readback cost".
    ///
    /// Off, a frame is taken, identified as a texture, and dropped — which measures the decoder
    /// and the reader and nothing else. On, the same frame is copied into a staging texture and
    /// read, which is what the application would have to do to draw it, and is a per-frame cost
    /// proportional to the picture's size.
    readback: bool,

    /// Frames whose sampled bytes were not all zero.
    frames_with_pixels: u64,

    /// Frames whose sampled bytes were all zero — a decode that "succeeded" and produced nothing.
    frames_all_zero: u64,

    /// Frames whose sampled bytes differed from the frame before, which is the difference between
    /// a decoder and a buffer.
    frames_changed: u64,

    /// The signatures of the sampled bytes, so "the same frame over and over" and "a picture
    /// changing" are different answers.
    distinct: BTreeSet<u64>,

    /// The last signature seen, for the frame-to-frame comparison.
    last_signature: Option<u64>,

    /// The mean of each frame's sampled bytes, averaged over the frames. A frame that is all
    /// zeroes and a frame of black are the same number and are told apart by `frames_all_zero`.
    mean_of_means: f64,

    /// The format the hardware decoder actually produced, read off the texture rather than
    /// assumed from the output type that was asked for.
    texture_format: Option<DXGI_FORMAT>,

    /// The staging texture the readback copies into, kept across frames.
    ///
    /// Making one per frame is a cost the measurement would then be reporting as the route's: at
    /// five and a half megabytes a frame and two hundred frames a second, a fresh allocation per
    /// frame is eleven gigabytes a second of driver allocations, and the first version of this
    /// file measured exactly that and called it the cost of reading video. One texture, reused,
    /// is what an application would do and is what makes the number mean anything.
    staging: Option<(u32, u32, i32, ID3D11Texture2D)>,

    /// The interval the reader is held to between frames, when one is imposed.
    ///
    /// A source reader is not a player: it hands over what it has as fast as it is asked, so
    /// uncapped it will decode this file at over two hundred frames a second — faster than the
    /// display can show and faster than the media engine's paced path ever does. Comparing an
    /// uncapped per-frame cost against a paced one is not a like-for-like comparison, so a rate
    /// can be imposed and the same measurement taken under it.
    frame_interval: Option<Duration>,

    /// When the next frame is due, so the rate is imposed by waiting rather than by polling.
    next_frame_at: Option<Instant>,

    /// The first few things that were refused, so a route that is mostly working and occasionally
    /// not can say which way.
    refusals: Vec<String>,
}

impl Run {
    fn new(context: &ID3D11DeviceContext, readback: bool) -> Self {
        Run {
            // A COM handle is a reference count, so taking another one here is free and lets a
            // window be returned by value rather than borrowing the caller's context — which is
            // what lets the two measured windows sit side by side in one `println!`.
            context: context.clone(),
            readback,
            frames: 0,
            dxgi_frames: 0,
            system_frames: 0,
            frames_with_pixels: 0,
            frames_all_zero: 0,
            frames_changed: 0,
            distinct: BTreeSet::new(),
            last_signature: None,
            mean_of_means: 0.0,
            texture_format: None,
            staging: None,
            frame_interval: frame_interval(),
            next_frame_at: None,
            refusals: Vec::new(),
        }
    }

    /// The evidence, in one line: a texture on the GPU, or bytes in this process's own memory.
    ///
    /// Deliberately does **not** say "hardware decode". A texture on a device is what a hardware
    /// decoder produces, and it is also what a device manager plus advanced video processing
    /// produce for a *software* decoder's output — the frame is copied into a GPU surface either
    /// way, and on this machine it is the second. The number that distinguishes them is the CPU,
    /// and that is [`Report`]'s to report.
    fn buffer_kind(&self) -> &'static str {
        if self.dxgi_frames > 0 {
            "IMFDXGIBuffer — a texture on the device manager's device"
        } else if self.system_frames > 0 {
            "IMF2DBuffer — system memory"
        } else {
            "neither: no frame arrived"
        }
    }

    /// `ReadSample` in a loop for `for_how_long`, judging every frame that comes back.
    ///
    /// The loop is a tight one with no sleep in it, and that is deliberate: a source reader is not
    /// a player, it hands over what is there as fast as it is asked, and a sleep here would
    /// measure the sleep. The clock outside it is what bounds the run.
    fn read(&mut self, reader: &IMFSourceReader, video_stream: u32, for_how_long: Duration) {
        let started = Instant::now();
        let mut means: f64 = 0.0;

        while started.elapsed() < for_how_long {
            // The wait is outside the clock the CPU is measured over only by accident, which is the
            // point: `GetProcessTimes` counts a sleeping thread as nothing, so pacing does not
            // inflate the number being measured. It is here so that the frames counted are the
            // frames a preview would actually ask for.
            if let Some(due) = self.next_frame_at {
                let now = Instant::now();
                if now < due {
                    std::thread::sleep(due - now);
                }
            }
            self.next_frame_at = self
                .frame_interval
                .map(|interval| Instant::now() + interval);

            let mut stream_flags = 0u32;
            let mut timestamp = 0i64;
            let mut sample = None;
            // SAFETY: `sample`, `stream_flags` and `timestamp` are the caller's own
            // out-parameters, initialised and live for the call, and the reader is alive across
            // it. `None` for the actual-stream-index out-parameter is what the existing
            // `heif_sequence` reader passes too, and is the shape this project has already got
            // frames out of a source reader with.
            let result = unsafe {
                reader.ReadSample(
                    video_stream,
                    0,
                    None,
                    Some(&mut stream_flags),
                    Some(&mut timestamp),
                    Some(&mut sample),
                )
            };

            // The end of the stream is `S_OK`, no sample, and this flag rather than an error —
            // which is why it is read off the flags rather than left to the `sample` check below.
            if result.is_err() || stream_flags & MF_SOURCE_READERF_ENDOFSTREAM.0 as u32 != 0 {
                let end = stream_flags & MF_SOURCE_READERF_ENDOFSTREAM.0 as u32 != 0;
                self.refusals.push(format!(
                    "ReadSample stopped ({end} end-of-stream, {stream_flags:#x}) after {} frames",
                    self.frames
                ));
                break;
            }

            // `ReadSample` answers `S_OK` and hands back no sample at the end of a file, which is
            // documented behaviour and not a decode that failed.
            let Some(sample) = sample else {
                self.refusals
                    .push(format!("S_OK with no sample — end of stream after {}", self.frames));
                break;
            };

            // SAFETY: the reader wrote the sample into the slot above and it came back non-null,
            // so it is a live `IMFSample` for as long as this scope holds it.
            let Ok(buffer) = (unsafe { sample.GetBufferByIndex(0) }) else {
                self.refusals.push("GetBufferByIndex(0) failed".to_string());
                continue;
            };

            self.frames += 1;
            means += self.judge(&buffer);
        }

        if self.frames > 0 {
            self.mean_of_means += means / self.frames as f64;
        }
    }

    /// One frame, down whichever of the two branches it arrives by, judged the same way either
    /// way, and its mean returned for the caller to average.
    ///
    /// The branch is the whole point and is made at run time rather than assumed: with a device
    /// manager attached a frame can come back as a texture, and on a machine where the hardware
    /// path silently declined the very same reader hands back system memory instead, and the
    /// difference is a factor of a thousand in cost. A spike that assumed the first would report
    /// a hardware number for a software run.
    fn judge(&mut self, buffer: &IMFMediaBuffer) -> f64 {
        // `cast` asks the buffer for another interface out of its own vtable and takes no
        // ownership of anything.
        if let Ok(dxgi) = buffer.cast::<IMFDXGIBuffer>() {
            self.dxgi_frames += 1;
            if !self.readback {
                return 0.0;
            }
            let sample = self.texture_sample(&dxgi);
            return self.note(sample);
        }

        self.system_frames += 1;
        if !self.readback {
            return 0.0;
        }

        let Ok(two_d) = buffer.cast::<IMF2DBuffer>() else {
            self.refusals
                .push("a buffer that is neither IMFDXGIBuffer nor IMF2DBuffer".to_string());
            return 0.0;
        };

        // SAFETY: `buffer` is a live `IMFMediaBuffer` and this asks it how many bytes it holds,
        // which is what bounds the read below.
        let length = unsafe { buffer.GetCurrentLength() }.unwrap_or(0);

        let mut scanline: *mut u8 = std::ptr::null_mut();
        let mut pitch: i32 = 0;
        // SAFETY: both out-parameters are the caller's own and live for the call. `Unlock2D` is
        // called on the way out of the block whatever the reads find, which is the only thing
        // that lets the buffer go when the sample does.
        let locked = unsafe { two_d.Lock2D(&mut scanline, &mut pitch) }.is_ok();
        if !locked || scanline.is_null() || pitch == 0 {
            if locked {
                // SAFETY: the lock taken above is released exactly once on the way out. A
                // refusal here is not actionable — there is nothing left to release by hand and
                // the buffer goes when the sample does.
                unsafe { two_d.Unlock2D() }.ok();
            }
            self.refusals
                .push(format!("IMF2DBuffer::Lock2D gave no pointer (pitch {pitch})"));
            return 0.0;
        }

        // A negative pitch is a buffer whose rows run backwards from the pointer; the bytes are
        // the same bytes either way, so the count is taken from the buffer's own length and the
        // walk starts at the first row.
        let stride = pitch.unsigned_abs() as usize;
        let bytes = length as usize / stride * stride;
        // SAFETY: `scanline` is the pointer the lock just reported and `bytes` is that stride
        // over only as many whole rows as the buffer's own length contains, so the slice is
        // inside the lock. `Unlock2D` below is what lets it out.
        let sample = sample_frame(unsafe {
            std::slice::from_raw_parts(scanline as *const u8, bytes)
        });
        // SAFETY: the lock taken above is released exactly once on the way out, whatever the
        // sampled bytes turned out to be — which is the whole reason the read is above this line
        // rather than deferred.
        unsafe { two_d.Unlock2D() }.ok();

        self.note(Some(sample))
    }

    /// A frame that came back as a texture: read out through a staging copy, because a texture
    /// on a device cannot be `Map`ped and a staging texture cannot be sampled from, and this is
    /// the only crossing between the two.
    fn texture_sample(&mut self, dxgi: &IMFDXGIBuffer) -> Option<(f64, u64)> {
        let mut raw = std::ptr::null_mut();
        // SAFETY: the out-parameter is the caller's own; `GetResource` fills it with a reference
        // to the buffer's own resource, which this takes over as the texture below.
        if let Err(error) = unsafe { dxgi.GetResource(&IID_ID3D11_TEXTURE2D, &mut raw) } {
            self.refusals.push(format!("IMFDXGIBuffer::GetResource: {error:?}"));
            return None;
        }

        if raw.is_null() {
            self.refusals
                .push("IMFDXGIBuffer::GetResource gave a null texture".to_string());
            return None;
        }

        // SAFETY: MF answered `S_OK` for a non-null pointer to the `ID3D11Texture2D` whose
        // identifier was asked for, and this wraps it in exactly one owner — the texture below,
        // which is dropped with this binding and releases the reference MF held.
        let texture = unsafe { ID3D11Texture2D::from_raw(raw) };

        // SAFETY: the texture is alive and `desc` is the caller's own, zeroed, and the call only
        // writes into it.
        let mut desc = D3D11_TEXTURE2D_DESC::default();
        unsafe { texture.GetDesc(&mut desc) };
        self.texture_format = Some(desc.Format);

        // One staging texture, reused for every frame of the same shape. The first version of this made a
        // fresh one per frame and measured twenty-seven milliseconds a frame — five and a half
        // megabytes of driver allocation, two hundred times a second, reported as if it were the
        // cost of reading video.
        let shape = (desc.Width, desc.Height, desc.Format.0);
        let staging = match &self.staging {
            Some((w, h, f, texture)) if (*w, *h, *f) == shape => texture.clone(),
            _ => {
                let Some(texture) = staging_texture(&self.context, &desc) else {
                    self.refusals.push(format!(
                        "no staging texture for {}x{} format {}",
                        desc.Width, desc.Height, desc.Format.0
                    ));
                    return None;
                };
                self.staging = Some((shape.0, shape.1, shape.2, texture.clone()));
                texture
            }
        };
        let staging = &staging;

        // SAFETY: `CopyResource` wants the two resources to have identical descriptors, which is
        // what `staging_texture` is built to guarantee — it mirrors this one and changes only the
        // three fields a staging texture is allowed to change. Both are on this same device, since
        // the staging copy is made on the device manager's context.
        unsafe { self.context.CopyResource(&*staging, &*texture) };

        let mut mapped = D3D11_MAPPED_SUBRESOURCE::default();
        // SAFETY: the staging texture is on this context's device and was just written to by this
        // context, `D3D11_MAP_READ` is what a `USAGE_STAGING` texture with CPU read access takes,
        // and `mapped` is the caller's own. It is unmapped on every path out of the block that
        // mapped it.
        if let Err(error) =
            unsafe { self.context.Map(&*staging, 0, D3D11_MAP_READ, 0, Some(&mut mapped)) }
        {
            self.refusals.push(format!("staging Map: {error:?}"));
            return None;
        }

        let length = desc.Height as usize * mapped.RowPitch as usize;
        let sample = if mapped.pData.is_null() || mapped.RowPitch == 0 || length == 0 {
            self.refusals.push(format!(
                "Map gave no memory (pitch {} over {} rows)",
                mapped.RowPitch, desc.Height
            ));
            None
        } else {
            // SAFETY: `Map` just reported the address and the row stride of the whole surface, and
            // `length` is that stride over the height the descriptor declares, so the whole
            // slice is inside the mapping. `Unmap` below is what lets it out.
            Some(sample_frame(unsafe {
                std::slice::from_raw_parts(mapped.pData as *const u8, length)
            }))
        };

        // SAFETY: unmapping the subresource that was mapped, on the same resource, on the same
        // context. Nothing below reads the mapping after this point.
        unsafe { self.context.Unmap(&*staging, 0) };

        sample
    }

    /// Record what a frame's sampled bytes turned out to be, and hand the mean back for the
    /// caller to average.
    fn note(&mut self, sample: Option<(f64, u64)>) -> f64 {
        let Some((mean, signature)) = sample else {
            return 0.0;
        };

        self.distinct.insert(signature);
        if self.last_signature != Some(signature) {
            self.frames_changed += 1;
        }
        self.last_signature = Some(signature);

        if mean == 0.0 {
            self.frames_all_zero += 1;
        } else {
            self.frames_with_pixels += 1;
        }
        mean
    }
}

/// A frame's sampled bytes, as `(mean, signature)`.
///
/// The mean says whether there is anything in the frame at all. The signature is a hash over the
/// same sampled bytes, and exists so that consecutive frames can be told apart without keeping any
/// of them: two reads that agree on the signature are the same picture whatever the reader did in
/// between.
fn sample_frame(frame: &[u8]) -> (f64, u64) {
    let mut sum: u64 = 0;
    let mut signature: u64 = 0xcbf2_9ce4_8422_2325;
    let mut taken: u64 = 0;

    for (index, byte) in frame.iter().enumerate() {
        if index % SAMPLE_STRIDE != 0 {
            continue;
        }
        let byte = *byte as u64;
        sum += byte;
        taken += 1;
        signature ^= byte;
        signature = signature.wrapping_mul(0x0000_0100_0000_01b3);
    }

    if taken == 0 {
        return (0.0, signature);
    }
    (sum as f64 / taken as f64, signature)
}

/// The staging copy's descriptor: the source texture's with only the three fields a staging
/// texture is permitted to differ in.
///
/// **Every field is written.** `CopyResource` refuses a pair whose descriptors differ, and
/// `CreateTexture2D` on a staging texture refuses a `SampleDesc` whose `Count` is zero — the trap
/// that cost two previous attempts a day each, and the reason nothing here is built with
/// `..Default::default()`. `MiscFlags` is written as zero rather than copied, because a resource
/// carrying any of them is not one a staging copy can be made of.
fn staging_texture(
    context: &ID3D11DeviceContext,
    source: &D3D11_TEXTURE2D_DESC,
) -> Option<ID3D11Texture2D> {
    let desc = D3D11_TEXTURE2D_DESC {
        Width: source.Width,
        Height: source.Height,
        MipLevels: source.MipLevels,
        ArraySize: source.ArraySize,
        Format: source.Format,
        SampleDesc: DXGI_SAMPLE_DESC {
            Count: source.SampleDesc.Count.max(1),
            Quality: 0,
        },
        Usage: D3D11_USAGE_STAGING,
        BindFlags: 0,
        CPUAccessFlags: D3D11_CPU_ACCESS_READ.0 as u32,
        MiscFlags: 0,
    };

    // SAFETY: an immediate context is always created over a device that outlives it, so asking
    // one back out of the context to create a resource on is the device this context is on.
    let device = unsafe { context.GetDevice() }.ok()?;

    let mut texture = None;
    // SAFETY: `desc` is a fully specified descriptor living on this stack for the call, there is
    // no initial data because a staging texture is written by a copy rather than by the CPU, and
    // `texture` is the caller's own out-parameter.
    match unsafe { device.CreateTexture2D(&desc, None, Some(&mut texture)) } {
        Ok(()) => texture,
        Err(_) => None,
    }
}

/// A hardware D3D11 device with both flags a video pipeline needs: `VIDEO_SUPPORT` because a
/// decoder will not be given a device without it, and `BGRA_SUPPORT` because the frames this
/// application composes in are BGRA.
fn hardware_device() -> Option<(ID3D11Device, ID3D11DeviceContext, D3D_FEATURE_LEVEL)> {
    let mut device = None;
    let mut context = None;
    let mut level = D3D_FEATURE_LEVEL(0);

    // SAFETY: every out-parameter is the caller's own and initialised, `None` for the adapter asks
    // for the default one and `None` for the HMODULE asks for the hardware driver rather than
    // WARP, and the feature level array is absent so the call takes its own default list.
    let created = unsafe {
        D3D11CreateDevice(
            None,
            D3D_DRIVER_TYPE_HARDWARE,
            None::<&HMODULE>,
            D3D11_CREATE_DEVICE_BGRA_SUPPORT | D3D11_CREATE_DEVICE_VIDEO_SUPPORT,
            None,
            D3D11_SDK_VERSION,
            Some(&mut device),
            Some(&mut level),
            Some(&mut context),
        )
    };

    if let Err(error) = created {
        println!("D3D11CreateDevice refused: {error}");
        return None;
    }

    Some((device?, context?, level))
}

/// The device manager, and then the device bound to it.
///
/// `ResetDevice` is the whole of "bind a device", and it is called before anything else is asked
/// of the manager on purpose: a manager with no device behind it does not fall back to software,
/// it refuses.
fn device_manager(device: &ID3D11Device) -> Option<(IMFDXGIDeviceManager, u32)> {
    let mut token: u32 = 0;
    let mut manager = None;
    // SAFETY: `token` is the caller's own out-parameter and `manager` the caller's own `Option`,
    // initialised to `None`; MF fills both, and the manager binding owns its reference from here.
    if let Err(error) = unsafe { MFCreateDXGIDeviceManager(&mut token, &mut manager) } {
        println!("MFCreateDXGIDeviceManager refused: {error}");
        return None;
    }
    let manager = manager?;

    // SAFETY: `IMFDXGIDeviceManager::ResetDevice` documents an `IUnknown` that is really the
    // `IDXGIDevice` the decoder will allocate on, and unwrapping the D3D11 device to it is what
    // Chromium's media foundation renderer does with the manager it locks. The cast is a query on
    // a live device, and the manager holds its own reference afterwards.
    let dxgi_device = device.cast::<IDXGIDevice>().ok()?;
    if let Err(error) = unsafe { manager.ResetDevice(&dxgi_device, token) } {
        println!("IMFDXGIDeviceManager::ResetDevice refused: {error}");
        return None;
    }

    Some((manager, token))
}

/// One output type to ask the reader for: a subtype, a size, and a frame rate, written out field
/// by field.
///
/// Every field is written for the same reason every descriptor in this file is: an output type
/// missing a field is one the reader is free to answer differently, and this is the type the
/// whole measurement is about.
fn output_type(subtype: &GUID, size: (u32, u32), rate: Option<(u32, u32)>) -> Option<IMFMediaType> {
    // SAFETY: MF allocates and returns the media type; there are no pointers involved, so the only
    // thing to say about it is that it has to be started before it is used.
    let media_type = match unsafe { MFCreateMediaType() } {
        Ok(media_type) => media_type,
        Err(error) => {
            println!("MFCreateMediaType refused: {error}");
            return None;
        }
    };

    let major = MFMediaType_Video;
    let interlace = MFVideoInterlace_Progressive;
    // SAFETY: every call writes into the media type rather than through a pointer of ours. The
    // `GUID`s are the caller's own `Copy`s on this stack, live for the whole block, and the
    // `UINT64`s are packed the way Media Foundation packs a pair — width over height, numerator
    // over denominator, the first in the high half; see `video_player::unpack_pair`, which is
    // where the transposed-frame mistake lives.
    let set = unsafe {
        [
            media_type.SetGUID(&MF_MT_MAJOR_TYPE, &major),
            media_type.SetGUID(&MF_MT_SUBTYPE, subtype),
            media_type
                .SetUINT64(&MF_MT_FRAME_SIZE, ((size.0 as u64) << 32) | size.1 as u64),
            media_type.SetUINT64(&MF_MT_PIXEL_ASPECT_RATIO, (1u64 << 32) | 1),
            media_type.SetUINT32(&MF_MT_INTERLACE_MODE, interlace.0 as u32),
        ]
    };

    // The frame rate is optional rather than defaulted: "no frame rate stated" is not a frame rate
    // of zero, and a reader asked for `0/0` refuses on its own terms rather than inferring the
    // file's. So a ladder rung that wants no rate simply does not write the field.
    if let Some((numerator, denominator)) = rate {
        // SAFETY: as above, writing one more value into the same media type.
        if let Err(error) = unsafe {
            media_type.SetUINT64(&MF_MT_FRAME_RATE, ((numerator as u64) << 32) | denominator as u64)
        } {
            println!("MF_MT_FRAME_RATE on the output type was refused: {error}");
            return None;
        }
    }

    for outcome in set {
        if let Err(error) = outcome {
            println!("a field of the output type was refused: {error}");
            return None;
        }
    }

    Some(media_type)
}

/// What a subtype goes by in the printed output, so a refusal in the ladder below says *which*
/// format was refused rather than only that one was.
fn subtype_name(subtype: &GUID) -> &'static str {
    match *subtype {
        MF_VIDEO_FORMAT_BGRA => "BGRA",
        g if g == MFVideoFormat_NV12 => "NV12",
        g if g == MFVideoFormat_HEVC => "HEVC",
        _ => "unknown",
    }
}

/// The output type the reader will actually take, out of a ladder tried in the order the
/// application most wants them, and a note of which rung it landed on.
///
/// This exists because `SetCurrentMediaType(BGRA, 2493x1400, 60/1)` was refused with
/// `MF_E_INVALIDMEDIATYPE` (`0xC00D36B4`) on the first run, and "BGRA does not negotiate" is a
/// very different conclusion from "60/1 does not negotiate and the decoder wanted the file's own
/// rate". Rather than guess which of the three fields caused it, each is varied on its own and the
/// answers are printed.
///
/// The order is the order of what a preview wants:
///   1. BGRA at the box size at the display's rate — the exact frame the app would draw.
///   2. BGRA at the box size at the *file's* rate — is the frame rate the refusal.
///   3. BGRA at the file's own size at the file's rate — is the resize the refusal.
///   4. NV12 at the file's own size at the file's rate — is the format the refusal, which is the
///      finding the research flagged as driver-restricted.
///   5. BGRA at the file's own size at no stated rate at all, in case a rate is the problem in a
///      way the reader would accept silently rather than by refusing.
///
/// Whichever rung lands, the reader is left configured with it — the accepted media type is what
/// the frames that follow will be. The accepted media type itself is not returned: the reader owns
/// a copy of it, and the frames are the proof that it took, which is what the rest of this file
/// measures.
fn negotiate(
    reader: &IMFSourceReader,
    video_stream: u32,
    box_size: (u32, u32),
    native_size: (u32, u32),
    native_rate: (u32, u32),
) -> Option<String> {
    let ladder: [(&GUID, (u32, u32), Option<(u32, u32)>); 6] = [
        (&MF_VIDEO_FORMAT_BGRA, box_size, Some(PREFERRED_FRAME_RATE)),
        (&MF_VIDEO_FORMAT_BGRA, box_size, Some(native_rate)),
        (&MF_VIDEO_FORMAT_BGRA, native_size, Some(native_rate)),
        (&MF_VIDEO_FORMAT_BGRA, native_size, None),
        (&MFVideoFormat_NV12, native_size, Some(native_rate)),
        (&MFVideoFormat_NV12, box_size, Some(PREFERRED_FRAME_RATE)),
    ];

    println!("output types tried, in order:");
    for (subtype, size, rate) in ladder {
        let Some(asked) = output_type(subtype, size, rate) else {
            continue;
        };

        let label = format!(
            "{}, {}x{}, {}",
            subtype_name(subtype),
            size.0,
            size.1,
            rate.map(|r| format!("{}/{}", r.0, r.1))
                .unwrap_or_else(|| "no stated rate".into())
        );

        // SAFETY: `asked` is a live media type built by this function and the reader is the live
        // one created a few lines above; `None` for the decoder index means "the type implies
        // its own".
        match unsafe { reader.SetCurrentMediaType(video_stream, None, &asked) } {
            Ok(()) => {
                println!("  {label} — accepted");
                return Some(label);
            }
            Err(error) => println!("  {label} — refused: {error}"),
        }
    }

    None
}

/// Video decoders this machine has registered for one subtype, by name, under whatever flags.
fn hardware_decoders(subtype: &GUID, flags: MFT_ENUM_FLAG) -> Option<Vec<String>> {
    let input = MFT_REGISTER_TYPE_INFO {
        guidMajorType: MFMediaType_Video,
        guidSubtype: *subtype,
    };

    let mut activates: *mut Option<IMFActivate> = std::ptr::null_mut();
    let mut count: u32 = 0;

    // SAFETY: `activates` and `count` are the caller's own out-parameters, `input` is a fully
    // specified input type for the call, and MF allocates the array it fills.
    let enumerated = unsafe {
        MFTEnumEx(
            MFT_CATEGORY_VIDEO_DECODER,
            flags,
            Some(&input),
            None,
            &mut activates,
            &mut count,
        )
    };
    enumerated.ok()?;

    // SAFETY: on success MF filled `count` slots of `activates`, each a live `IMFActivate` this
    // owns exactly one reference to, and `activates` itself is one allocation of `count` of them.
    let mut names = Vec::new();
    if !activates.is_null() {
        for slot in unsafe { std::slice::from_raw_parts_mut(activates, count as usize) } {
            // SAFETY: each slot is a distinct, initialised `Option<IMFActivate>` MF wrote, and
            // `take` leaves it empty so the loop cannot release anything twice. Dropping the value
            // releases MF's reference, which is how the array is meant to be emptied.
            let Some(activate) = slot.take() else {
                continue;
            };
            names.push(
                activate
                    .cast::<IMFAttributes>()
                    .map_or_else(|_| "(no attributes)".to_string(), |a| friendly_name(&a)),
            );
        }
        // SAFETY: `activates` is the single MF-allocated block the loop above just emptied, and
        // it is freed exactly once, with the allocator MF allocated it from.
        unsafe { CoTaskMemFree(Some(activates as *const core::ffi::c_void)) };
    }

    Some(names)
}

/// Every MFT registered as a video decoder on this machine, named, and marked with whether it
/// claims to be a hardware one.
///
/// Enumerated with no input-type filter at all, because the filtered enumeration answered "zero
/// hardware decoders" for HEVC, for H.264 and for AV1 alike on a machine whose only hardware video
/// decoder is an AMD one. An enumeration that answers zero where the answer is "one software
/// decoder and no hardware decoder" is worse than no enumeration at all: it looks like evidence,
/// and it would have read as proof that hardware HEVC decode is impossible here when what it
/// actually shows is that the filter matches nothing. Each entry is asked for
/// `MFT_ENUM_HARDWARE_URL_Attribute` as well, which is how a hardware MFT declares itself, and the
/// name is reported with it so a decoder that turns out to be named
/// something else is visible rather than hidden behind a count.
fn all_video_decoders() -> Vec<String> {
    let mut activates: *mut Option<IMFActivate> = std::ptr::null_mut();
    let mut count: u32 = 0;

    // SAFETY: `activates` and `count` are the caller's own out-parameters, both input types are
    // `None` so nothing is filtered out, and MF allocates the array it fills.
    let enumerated = unsafe {
        MFTEnumEx(
            MFT_CATEGORY_VIDEO_DECODER,
            MFT_ENUM_FLAG_ALL,
            None,
            None,
            &mut activates,
            &mut count,
        )
    };
    if enumerated.is_err() || activates.is_null() {
        return Vec::new();
    }

    // SAFETY: on success MF filled `count` slots of `activates`, each a live `IMFActivate` this
    // owns exactly one reference to, and `activates` itself is one allocation of `count` of them.
    let mut names = Vec::new();
    for slot in unsafe { std::slice::from_raw_parts_mut(activates, count as usize) } {
        // SAFETY: each slot is a distinct, initialised `Option<IMFActivate>` MF wrote, and `take`
        // leaves it empty so the loop cannot release anything twice. Dropping the value releases
        // MF's reference, which is how the array is meant to be emptied.
        let Some(activate) = slot.take() else {
            continue;
        };

        // An `IMFActivate` is itself an `IMFAttributes` — it is its factory's attribute store, and
        // it carries the friendly name and the hardware URL before anything is activated. Casting
        // through `IMFTransform` instead, which is the obvious way to read an MFT's attributes,
        // fails on every single one of them: the object does not exist until it is activated.
        let attributes = activate.cast::<IMFAttributes>().ok();

        let Some(attributes) = attributes else {
            names.push("(a decoder whose attributes could not be read)".to_string());
            continue;
        };

        let name = friendly_name(&attributes);
        // A hardware MFT carries a hardware URL; there is no separate flag, and its presence is
        // the only thing that distinguishes it from a software decoder in the same category.
        let mut url = PWSTR::null();
        let mut url_length = 0u32;
        // SAFETY: both out-parameters are the caller's own and live for the call; the block `url`
        // is left pointing at is freed below whichever way this goes.
        let hardware = unsafe {
            attributes
                .GetAllocatedString(&MFT_ENUM_HARDWARE_URL_Attribute, &mut url, &mut url_length)
                .is_ok()
        };
        // SAFETY: the block `GetAllocatedString` just filled, freed exactly once with the
        // allocator MF allocated it from.
        unsafe { CoTaskMemFree(Some(url.0 as *const core::ffi::c_void)) };

        names.push(format!("{name} — hardware URL: {hardware}"));
    }

    // SAFETY: `activates` is the single MF-allocated block the loop above just emptied, and it
    // is freed exactly once, with the allocator MF allocated it from.
    unsafe { CoTaskMemFree(Some(activates as *const core::ffi::c_void)) };

    names
}

/// What an MFT calls itself, from its own attributes.
///
/// `GetAllocatedString` hands back a block this has to free, which is why this is a function
/// rather than a line: a leaked friendly name per decoder would be a small and permanent leak in a
/// process that is about to run a preview for hours.
fn friendly_name(attributes: &IMFAttributes) -> String {
    let mut name = PWSTR::null();
    let mut length = 0u32;

    // SAFETY: both out-parameters are the caller's own, initialised and live for the call, and
    // the block `name` ends up pointing at is freed on every path out below.
    let read = unsafe {
        attributes.GetAllocatedString(&MFT_FRIENDLY_NAME_Attribute, &mut name, &mut length)
    };

    let found = read
        .ok()
        .and_then(|()| unsafe { name.to_string() }.ok())
        .unwrap_or_else(|| "(unnamed)".to_string());

    // SAFETY: `name` is the MF-allocated block `GetAllocatedString` just filled, and it is freed
    // exactly once with the allocator MF allocated it from.
    unsafe { CoTaskMemFree(Some(name.0 as *const core::ffi::c_void)) };
    found
}

/// How many hardware video decoders this machine has registered for a subtype.
///
/// The independent half of "hardware decode was actually engaged". The buffer type is the direct
/// evidence — a frame on a texture cannot come from a software decoder — but a machine with no
/// hardware decoder for the file at all would read as a working route with a wrong explanation
/// for it, and this is what tells the two apart.

/// The three reader attributes the recipe fixes, read from the environment so one build can be run
/// with any of them off.
///
/// A spike that can only be run one way cannot answer "which of these is responsible" when it
/// answers badly, and the first run did answer badly — `MF_E_INVALIDMEDIATYPE` on BGRA and then
/// `E_POINTER` from `ReadSample`. Both of those have more than one possible cause, and guessing
/// between them by editing and rebuilding is how a half-day question becomes a two-day one.
///
/// `RHP_SPIKE_XVP`, `RHP_SPIKE_HW` and `RHP_SPIKE_MANAGER` are each off when set to `0`, and on
/// when set to anything else or not set at all — so the gate's configuration is the default.
fn switches() -> (bool, bool, bool) {
    let off = |name: &str| std::env::var(name).as_deref() == Ok("0");
    (
        !off("RHP_SPIKE_XVP"),
        !off("RHP_SPIKE_HW"),
        !off("RHP_SPIKE_MANAGER"),
    )
}

/// The rate the reader is held to, from `RHP_SPIKE_FPS`, or nothing at all when that is not set.
///
/// This exists because a source reader paces nothing. Left alone it decodes this file at over two
/// hundred frames a second — faster than the display refreshes, faster than the media engine's own
/// paced path ever manages, and a workload no preview ever asks for. The per-frame CPU of such a
/// run is not comparable with the per-frame CPU of a paced one, so both are measured and both are
/// reported rather than whichever looks better.
///
/// Sixty is the rate `video_take_cost` measured the software path at, and is the default here for
/// the same reason.
fn frame_interval() -> Option<Duration> {
    match std::env::var("RHP_SPIKE_FPS") {
        Ok(rate) => match rate.trim().parse::<f64>() {
            Ok(rate) if rate > 0.0 => Some(Duration::from_secs_f64(1.0 / rate)),
            _ => {
                println!("RHP_SPIKE_FPS is not a rate — reading at whatever speed it goes");
                None
            }
        },
        Err(_) => Some(Duration::from_millis(1000 / 60)),
    }
}

/// What this process has spent on the CPU, in 100-nanosecond units: kernel and user together,
/// since a decode handed to a driver spends its time in whichever of the two the driver spends it
/// in.
///
/// Copied rather than shared with `video_player`'s own tests because a test module's helper is
/// not visible from here, and because this file is throwaway: whatever it borrows from the app it
/// does not take a dependency on.
fn cpu_hundred_nanoseconds() -> Option<u64> {
    let mut created = FILETIME::default();
    let mut exited = FILETIME::default();
    let mut kernel = FILETIME::default();
    let mut user = FILETIME::default();

    // SAFETY: all four out-parameters are the caller's own, initialised, and live for the
    // duration of the call; `GetCurrentProcess` is a pseudo-handle that is always valid and is
    // what the process's own times are read from. A refusal reads as no measurement rather than as
    // a zero, which would flatter everything this is used for.
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

#[cfg(test)]
#[allow(dead_code)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    /// The spike, behind `RHP_SPIKE_FILE`, in release.
    ///
    /// The gate is three conditions and all three are printed: frames read, CPU per frame under
    /// two milliseconds, and hardware decode actually engaged.
    ///
    /// The third condition is **not** "the buffer was an `IMFDXGIBuffer`", which is what the brief
    /// suggested and what the first version of this test checked. A frame on a device is what a
    /// hardware decoder produces, and it is also what a source reader produces for a *software*
    /// decoder's output once a device manager and advanced video processing are in the pipeline.
    /// On this machine it is unambiguously the second — `HEVCVideoExtension` carries no hardware
    /// URL, and removing the device manager changes the buffer type without changing the cost — so
    /// a gate that read "texture" as "hardware" would have reported a hardware pass for a software
    /// decode that happened to be measured at 27 ms.
    ///
    /// What is checked instead is the thing that cannot be argued with: **a hardware decode of
    /// 1440p HEVC does not cost twenty milliseconds of CPU.** Anything over a millisecond or so
    /// is a software decode wearing a GPU surface, and this test says so rather than reading the
    /// surface as the evidence.
    #[test]
    #[ignore = "decodes the file named in RHP_SPIKE_FILE through a source reader"]
    fn source_reader_hw_decode() {
        let Ok(file) = std::env::var("RHP_SPIKE_FILE") else {
            println!("set RHP_SPIKE_FILE to the path of a video to read");
            return;
        };

        let path = PathBuf::from(file);
        let (width, height) = box_from_env((2493, 1400));

        println!("\n=== source reader + DXGI device manager ===");
        println!("{}", path.display());
        println!("box {width}x{height}\n");

        let Some(summary) = measure(&path, width, height) else {
            println!("\nFAIL: the route produced no frames at all");
            return;
        };

        let pixels = summary.frames > 0
            && summary.frames_with_pixels > 0
            && summary.frames_changed > 1;
        let cheap = summary.cpu_per_frame_ms < 2.0;
        // A hardware decode is a fraction of a millisecond of CPU a frame. Two milliseconds is the
        // gate's own budget and is also the loosest number that could be called hardware decode;
        // a frame costing more than that is being decoded on this machine's CPU.
        let hardware = summary.cpu_per_frame_ms < 2.0;

        println!("\n=== the gate ===");
        println!("frames read ................ {}", summary.frames);
        println!(
            "CPU per frame ............... {:.4} ms  (under 2 ms: {})",
            summary.cpu_per_frame_ms,
            if cheap { "yes" } else { "NO" }
        );
        println!(
            "pixels decoded and moving .. {} frames with bytes, {} differing from the one before",
            summary.frames_with_pixels, summary.frames_changed
        );
        println!("buffer type ................. {}", summary.buffer_kind);
        println!(
            "hardware decode engaged ..... {}",
            if hardware {
                "yes"
            } else {
                "NO — see the module documentation: this machine has no hardware HEVC decoder"
            }
        );
        println!(
            "verdict ..................... {}",
            if pixels && cheap && hardware {
                "PASS"
            } else {
                "FAIL"
            }
        );
    }

    /// The box size, from `RHP_SPIKE_BOX`, which is the display's own size and the one
    /// `video_player::tests::video_take_cost` was measured at.
    fn box_from_env(default: (u32, u32)) -> (u32, u32) {
        let Ok(box_) = std::env::var("RHP_SPIKE_BOX") else {
            return default;
        };
        let Some((width, height)) = box_.split_once('x') else {
            println!("RHP_SPIKE_BOX is not WIDTHxHEIGHT — using {default:?}");
            return default;
        };

        match (width.trim().parse(), height.trim().parse()) {
            (Ok(width), Ok(height)) if width > 0 && height > 0 => (width, height),
            _ => {
                println!("RHP_SPIKE_BOX is not two whole numbers — using {default:?}");
                default
            }
        }
    }
}
