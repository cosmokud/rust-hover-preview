//! What this machine has: every engine and every codec extension a preview leans on.
//!
//! Two answers are asked of the machine as a whole, and one question is asked per hover.
//! The tray's `Codecs` submenu asks for the whole list, once, when the menu is built —
//! that is the moment a user is looking at what is installed, and the moment an answer is
//! allowed to change. The video path asks the question of its own, which engine plays a
//! video, and asks it once per hover — which is what the cache below is for. And the
//! question of whether a document's own application is installed is asked per hover as
//! well: it is what decides where an Office document's page comes from, whether from
//! Office or from the render engine beside it (see `office_formats::app_installed`). It is
//! a registry read like every other probe here, cheap enough to be asked per hover, and
//! asking it again is also what lets an application installed while the app is running be
//! used by the next hover rather than by the next restart.
//!
//! Nothing here starts anything. An Office engine is asked for by its ProgID, which is a
//! registry read rather than a process; a decoder is asked for by enumerating what is
//! registered rather than by decoding; a browser is asked for through the loader's own
//! version query. A probe that started a program would cost more than the answer is worth,
//! and the one for Office would leave an application running behind a menu.

use once_cell::sync::Lazy;
use std::path::{Path, PathBuf};
use std::sync::{Mutex, OnceLock};
use windows::core::{IUnknown, Interface, GUID};
use windows::Win32::Graphics::Imaging::{
    IWICBitmapCodecInfo, WICComponentEnumerateDefault, WICDecoder,
};
use windows::Win32::Media::MediaFoundation::{
    IMFActivate, MFMediaType_Audio, MFMediaType_Video, MFStartup, MFTEnumEx, MFVideoFormat_AV1,
    MFVideoFormat_H264, MFVideoFormat_HEVC, MFVideoFormat_MP4V, MFVideoFormat_MPEG2,
    MFVideoFormat_Theora, MFVideoFormat_VP90, MFVideoFormat_WMV3, MFAudioFormat_AAC,
    MFAudioFormat_ADTS, MFAudioFormat_ALAC, MFAudioFormat_DTS, MFAudioFormat_Dolby_AC3,
    MFAudioFormat_FLAC, MFAudioFormat_MP3, MFAudioFormat_Opus, MFAudioFormat_Vorbis,
    MFAudioFormat_WMAudioV8, MFAudioFormat_WMAudioV9,
    MFAudioFormat_WMAudio_Lossless, MFSTARTUP_FULL, MFT_CATEGORY_AUDIO_DECODER,
    MFT_CATEGORY_VIDEO_DECODER, MFT_ENUM_FLAG, MFT_ENUM_FLAG_ASYNCMFT, MFT_ENUM_FLAG_HARDWARE,
    MFT_ENUM_FLAG_LOCALMFT, MFT_ENUM_FLAG_SYNCMFT, MFT_REGISTER_TYPE_INFO, MF_VERSION,
};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoInitializeEx, CoTaskMemFree, COINIT_MULTITHREADED,
};
use windows::Win32::System::Registry::{
    RegCloseKey, RegOpenKeyExW, HKEY, HKEY_LOCAL_MACHINE, KEY_READ,
};

/// The player FFmpeg is recognized by. One probe answers for the whole install: `ffplay`,
/// `ffprobe` and `ffmpeg` arrive together, and the player is the one whose absence a
/// video preview has to answer for.
const FFPLAY_NAME: &str = "ffplay.exe";

/// And the program a sound's peak is measured by: the other name normalization asks for, since
/// the install the player comes from is the install the meter comes from (see
/// `normalize_available`).
const FFMPEG_NAME: &str = "ffmpeg.exe";

/// The WinRT class the Windows PDF engine is activated through. What is asked of it is
/// whether it is registered at all, which is what the system itself answers activation
/// from: the class is opened with a stream rather than with nothing, so there is no
/// cheaper way to ask it something.
const PDF_ENGINE_CLASS: &str =
    r"SOFTWARE\Microsoft\WindowsRuntime\ActivatableClassId\Windows.Data.Pdf.PdfDocument";

/// The MIME types a codec extension claims a picture format under. A component that
/// reports one of these is the extension being installed, which is the same question the
/// decoder answers when it is handed a file of that kind.
const HEIF_MIME_TYPES: &[&str] = &["image/heic", "image/heif"];
const AVIF_MIME_TYPES: &[&str] = &["image/avif"];
const JXL_MIME_TYPES: &[&str] = &["image/jxl"];

/// What a decoder may be registered as. The transcode-only and post-process flags are
/// left out: neither is a decoder a preview could play a file through.
const DECODER_ENUM_FLAGS: MFT_ENUM_FLAG = MFT_ENUM_FLAG(
    MFT_ENUM_FLAG_SYNCMFT.0
        | MFT_ENUM_FLAG_ASYNCMFT.0
        | MFT_ENUM_FLAG_HARDWARE.0
        | MFT_ENUM_FLAG_LOCALMFT.0,
);

/// Where each thing a machine can be missing is got from: the pages the README names for them,
/// kept beside the rows so that picking one opens exactly what the README tells a user to
/// install rather than a search for it.
///
/// The `apps.microsoft.com` ones are the Store's own pages and are opened as pages: what a row
/// hands over is a link, and what answers it is the browser the user already has.
const FFMPEG_PAGE: &str = "https://ffmpeg.org/download.html";
const MPEG2_PAGE: &str = "https://apps.microsoft.com/detail/9N95Q1ZZPMH4";
const HEVC_PAGE: &str = "https://apps.microsoft.com/detail/9N4WGH0Z6VHQ";
const VP9_PAGE: &str = "https://apps.microsoft.com/detail/9N4D0MSMP0PT";
const AV1_PAGE: &str = "https://apps.microsoft.com/detail/9MVZQVXJBQ9V";
/// The one package that carries both halves of the Ogg family: Theora for a video and Vorbis
/// and Opus for a sound are the same extension, which is why the rows for all three name it.
const THEORA_PAGE: &str = "https://apps.microsoft.com/detail/9N5TDP8VCMHS";
const HEIF_PAGE: &str = "https://apps.microsoft.com/detail/9PMMSR1CGPWG";
const JXL_PAGE: &str = "https://apps.microsoft.com/detail/9MZPRTH5C0TB";
const LIBREOFFICE_PAGE: &str = "https://www.libreoffice.org/download/";
const IMAGEMAGICK_PAGE: &str = "https://imagemagick.org/download/";
const PEAZIP_PAGE: &str = "https://peazip.github.io/peazip-64bit.html";
const CALIBRE_PAGE: &str = "https://calibre-ebook.com/download_windows";

/// One row of the tray's `Codecs` submenu: what the engine or codec is called, whether this
/// machine has it, and where it is got from where it does not.
pub struct Row {
    pub name: &'static str,
    pub available: bool,
    /// Where a row this machine does not have is got from: the page the README names for it,
    /// opened in the browser when the row is picked. `None` where there is nowhere to send
    /// anyone — a component Windows ships, and an application this app cannot point at — and,
    /// for the two picture formats that arrive in two packages, the page of whichever package
    /// this machine is missing.
    pub link: Option<&'static str>,
}

/// The engines and the codecs a video preview leans on.
///
/// H.264, MPEG-4 and WMV are what Windows 11 decodes out of the box, and they are listed
/// so that the rows under them read as what they are: what a machine usually has to be
/// given. Those four are the Microsoft Store's free video extensions, one package each.
pub fn video() -> Vec<Row> {
    vec![
        Row {
            name: "FFmpeg (ffplay)",
            available: ffplay_available(),
            link: Some(FFMPEG_PAGE),
        },
        Row {
            name: "Windows Media Foundation",
            available: mf_started(),
            link: None,
        },
        Row {
            name: "H.264",
            available: video_decoder(&MFVideoFormat_H264),
            link: None,
        },
        // One row for the two, because one decoder answers for both: Windows' MPEG-4
        // Part 2 decoder is what plays a DivX or Xvid file as well.
        Row {
            name: "MPEG-4 / WMV",
            available: video_decoder(&MFVideoFormat_MP4V) || video_decoder(&MFVideoFormat_WMV3),
            link: None,
        },
        Row {
            name: "MPEG-2",
            available: video_decoder(&MFVideoFormat_MPEG2),
            link: Some(MPEG2_PAGE),
        },
        Row {
            name: "HEVC (H.265)",
            available: video_decoder(&MFVideoFormat_HEVC),
            link: Some(HEVC_PAGE),
        },
        Row {
            name: "VP9",
            available: video_decoder(&MFVideoFormat_VP90),
            link: Some(VP9_PAGE),
        },
        Row {
            name: "AV1",
            available: video_decoder(&MFVideoFormat_AV1),
            link: Some(AV1_PAGE),
        },
        Row {
            name: "Theora (Ogg)",
            available: video_decoder(&MFVideoFormat_Theora),
            link: Some(THEORA_PAGE),
        },
    ]
}

/// The engines and the codecs a sound preview leans on.
///
/// The same shape as the video list above it: what Windows decodes out of the box is listed
/// first, and what a machine usually has to be given after it. The decoders are asked of the
/// machine rather than told to it — one package carries the whole Ogg family, and FFmpeg is
/// what plays everything neither it nor Windows reads — and the rows are the codecs a file is
/// most likely to hold, not every format the list of sound names has an entry for.
///
/// The one row that is not a decoder is the engine Windows has, which is what plays a sound
/// where a decoder above it reaches one, and which is what a hover on such a file is played
/// by: a machine with no Media Foundation at all is a machine whose sounds are played by
/// FFmpeg or not at all.
pub fn audio() -> Vec<Row> {
    vec![
        Row {
            name: "FFmpeg (ffplay)",
            available: ffplay_available(),
            link: Some(FFMPEG_PAGE),
        },
        Row {
            name: "Windows Media Foundation",
            available: mf_started(),
            link: None,
        },
        Row {
            name: "MP3",
            available: audio_decoder(&MFAudioFormat_MP3),
            link: None,
        },
        // One row for the two spellings of AAC, because one decoder answers for both: an
        // `.aac` file is the raw stream and an `.m4a` the same codec in a container.
        Row {
            name: "AAC / M4A",
            available: audio_decoder(&MFAudioFormat_AAC) || audio_decoder(&MFAudioFormat_ADTS),
            link: None,
        },
        // And one for the WMA family, whose three decoders are one format to a person.
        Row {
            name: "WMA",
            available: audio_decoder(&MFAudioFormat_WMAudioV8)
                || audio_decoder(&MFAudioFormat_WMAudioV9)
                || audio_decoder(&MFAudioFormat_WMAudio_Lossless),
            link: None,
        },
        Row {
            name: "FLAC",
            available: audio_decoder(&MFAudioFormat_FLAC),
            link: None,
        },
        Row {
            name: "ALAC",
            available: audio_decoder(&MFAudioFormat_ALAC),
            link: None,
        },
        // The one row here that is a package rather than something Windows ships: Vorbis and
        // Opus are the Ogg family's audio half, and the extension that carries Theora carries
        // them (see `THEORA_PAGE`).
        Row {
            name: "Vorbis & Opus (Ogg)",
            available: audio_decoder(&MFAudioFormat_Vorbis) || audio_decoder(&MFAudioFormat_Opus),
            link: Some(THEORA_PAGE),
        },
        Row {
            name: "Dolby Digital (AC-3)",
            available: audio_decoder(&MFAudioFormat_Dolby_AC3),
            link: None,
        },
        Row {
            name: "DTS",
            available: audio_decoder(&MFAudioFormat_DTS),
            link: None,
        },
    ]
}

/// The picture formats this app has no decoder of its own for — the ones a codec
/// extension is needed for.
///
/// A `.heic` and an `.avif` are both a HEIF container, so both need the HEIF extension;
/// what the two do not share is the compression inside it — HEVC for one, AV1 for the
/// other — which is why each asks for the container *and* its own decoder. It is the
/// two-package answer the README gives, asked of the machine rather than told to it.
pub fn images() -> Vec<Row> {
    let heif = image_codec(HEIF_MIME_TYPES);
    let hevc = video_decoder(&MFVideoFormat_HEVC);
    let av1 = video_decoder(&MFVideoFormat_AV1);

    let heif_available = heif && hevc;
    let avif_available = image_codec(AVIF_MIME_TYPES) || (heif && av1);

    vec![
        Row {
            name: "HEIF (HEIC)",
            available: heif_available,
            // Two packages carry this format — the container and the compression inside it —
            // and a row opens one page, so what is offered is the page for whichever of the
            // two is missing: the container's extension where there is no container, and the
            // codec's where only the codec is gone.
            link: if heif_available {
                None
            } else if !heif {
                Some(HEIF_PAGE)
            } else {
                Some(HEVC_PAGE)
            },
        },
        Row {
            name: "AVIF",
            available: avif_available,
            // The same two-package answer as the row above, with AV1 in the codec's place.
            link: if avif_available {
                None
            } else if !heif {
                Some(HEIF_PAGE)
            } else {
                Some(AV1_PAGE)
            },
        },
        Row {
            name: "JPEG XL",
            available: image_codec(JXL_MIME_TYPES),
            link: Some(JXL_PAGE),
        },
        // The one picture here that needs nothing: this app carries its own libwebp,
        // which is what plays an animated one and what decodes a still one where the
        // codec is missing. The row stands for an engine that is in the binary, so it is
        // always there — and it is listed so that the group reads as the whole of what a
        // picture can be.
        Row {
            name: "WebP",
            available: true,
            link: None,
        },
    ]
}

/// The engines a preview leans on rather than a codec: the ones that draw a document, a
/// specimen or a page, the one that develops a picture, and the ones that draw a page of a
/// document Office owns.
pub fn engines() -> Vec<Row> {
    vec![
        Row {
            name: "WebView2 Runtime",
            available: crate::engines::webview_preview::is_available(),
            link: None,
        },
        Row {
            name: "LibreOffice",
            available: crate::engines::libreoffice_render::available(),
            link: Some(LIBREOFFICE_PAGE),
        },
        Row {
            name: "ImageMagick",
            available: crate::engines::imagemagick_render::available(),
            link: Some(IMAGEMAGICK_PAGE),
        },
        Row {
            name: "PeaZip",
            available: crate::engines::peazip_render::available(),
            link: Some(PEAZIP_PAGE),
        },
        Row {
            name: "Calibre",
            available: crate::engines::calibre_render::available(),
            link: Some(CALIBRE_PAGE),
        },
        // The Office applications are the one group here with no page to open: they are not
        // free downloads, so a row that is missing one stays a row that cannot be picked.
        Row {
            name: "Microsoft Word",
            available: prog_id_installed("Word.Application"),
            link: None,
        },
        Row {
            name: "Microsoft Excel",
            available: prog_id_installed("Excel.Application"),
            link: None,
        },
        Row {
            name: "Microsoft PowerPoint",
            available: prog_id_installed("PowerPoint.Application"),
            link: None,
        },
        Row {
            name: "Windows PDF Engine",
            available: winrt_class_registered(PDF_ENGINE_CLASS),
            link: None,
        },
    ]
}

/// Ask again, which is what the tray's menu build does before it lists the answers.
///
/// Every other probe is asked on demand and answers from the machine at that moment. The
/// answers that are kept are FFmpeg's own — the player a video is played by, and the meter and
/// the player a sound's peak is measured and applied with — because they are asked on every
/// video hover and every sound hover, and a hover must not go looking through the `PATH` for
/// them: opening the menu is what lets a machine that has just been given FFmpeg start using
/// it, rather than a restart.
pub fn refresh() {
    if let Ok(mut cached) = FFPLAY.lock() {
        *cached = None;
    }

    if let Ok(mut cached) = NORMALIZE.lock() {
        *cached = None;
    }
}

/// Whether `ffplay.exe` is on this machine, and so whether a video preview is played by
/// FFmpeg rather than by the engine Windows has.
///
/// The answer is kept between hovers: it is asked for every video a pointer lands on, and
/// the `PATH` is a list of directories rather than a question. [`refresh`] is what asks it
/// again.
pub fn ffplay_available() -> bool {
    if let Ok(cached) = FFPLAY.lock() {
        if let Some(answer) = *cached {
            return answer;
        }
    }

    let answer = find_ffplay().is_some();

    if let Ok(mut cached) = FFPLAY.lock() {
        *cached = Some(answer);
    }

    answer
}

/// Whether a sound's peak can be measured and applied on this machine: the meter that measures
/// it and the player that applies the gain to it are one install, and both of them have to be
/// here for the tray's `Normalize` row to be anything but a switch that cannot act.
///
/// It is asked the way the player's own answer is asked and kept the same way — an install that
/// arrives while the app is running is picked up by the next opening of the menu rather than at
/// the next restart (see [`refresh`]).
pub fn normalize_available() -> bool {
    if let Ok(cached) = NORMALIZE.lock() {
        if let Some(answer) = *cached {
            return answer;
        }
    }

    let answer = find_ffplay().is_some() && find_program(FFMPEG_NAME).is_some();

    if let Ok(mut cached) = NORMALIZE.lock() {
        *cached = Some(answer);
    }

    answer
}

/// Whether a video is played by the engine Windows has, which is what the layout and the
/// renderer both have to agree about: the one asks whether a video has a size to be placed
/// at, and the other asks which of the two engines draws it.
pub fn plays_video_natively() -> bool {
    !ffplay_available()
}

/// Whether the media stack is up, and with it every decoder this machine has.
///
/// Media Foundation is started once for the process and never shut down: it is the library
/// every codec lives in, the app may need it at any moment, and the count it keeps is the
/// process's own. Starting it is also what asks the question — a machine where it will not
/// start is a machine with no media engine, and every video below that depends on one is
/// answered with no.
pub fn mf_started() -> bool {
    initialize_apartment();

    *MEDIA_FOUNDATION.get_or_init(|| unsafe { MFStartup(MF_VERSION, MFSTARTUP_FULL).is_ok() })
}

/// The apartment a thread has to be in before it can hold anything the media stack makes.
///
/// It is asked for per thread, which is why it is not inside the `OnceLock` above: the
/// library is started once for the process, while an apartment is the thread's own. Two
/// threads reach the media stack from here — the worker that measures a file and the
/// preview thread that plays it — and a thread that never asked for an apartment is a
/// thread whose objects cannot be made at all.
///
/// A thread that has already asked for another apartment is answered with
/// `RPC_E_CHANGED_MODE`, which is not a reason to give up: the tray's own thread is a
/// single-threaded apartment, and enumerating decoders from it works.
fn initialize_apartment() {
    thread_local! {
        static DONE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    }

    DONE.with(|done| {
        if done.replace(true) {
            return;
        }

        unsafe {
            let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
        }
    });
}

/// Where `ffplay.exe` is, asked the way `Command::new("ffplay")` asks it: the folder this
/// app runs from first, then every folder on the `PATH`.
///
/// Spawning the player to ask whether it exists would be a program started to answer a
/// question about a program, and the `PATH` already holds the answer.
fn find_ffplay() -> Option<PathBuf> {
    find_program(FFPLAY_NAME)
}

/// Where a program of FFmpeg's is, by the name it is installed under: the same two places
/// `Command::new` looks for it, and nothing started to ask.
///
/// Spawning the program to ask whether it exists would be a program started to answer a
/// question about a program, and the `PATH` already holds the answer.
fn find_program(name: &str) -> Option<PathBuf> {
    let mut folders: Vec<PathBuf> = Vec::new();

    if let Some(own_folder) = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.parent().map(Path::to_path_buf))
    {
        folders.push(own_folder);
    }

    if let Some(path) = std::env::var_os("PATH") {
        folders.extend(std::env::split_paths(&path));
    }

    folders
        .into_iter()
        .map(|folder| folder.join(name))
        .find(|program| program.is_file())
}

/// Whether a decoder for one compressed video format is registered.
///
/// The input type is what is asked for, because that is the side a decoder is named by: an
/// extension package adds an MFT whose *input* is HEVC, AV1, VP9 or MPEG-2, and what it
/// produces is not this app's business.
fn video_decoder(subtype: &GUID) -> bool {
    if !mf_started() {
        return false;
    }

    let input = MFT_REGISTER_TYPE_INFO {
        guidMajorType: MFMediaType_Video,
        guidSubtype: *subtype,
    };

    let mut activates: *mut Option<IMFActivate> = std::ptr::null_mut();
    let mut count: u32 = 0;

    let enumerated = unsafe {
        MFTEnumEx(
            MFT_CATEGORY_VIDEO_DECODER,
            DECODER_ENUM_FLAGS,
            Some(&input),
            None,
            &mut activates,
            &mut count,
        )
    };

    // The array is the caller's to free whether or not anything was found, and each entry
    // holds a reference until it is dropped.
    unsafe {
        if !activates.is_null() {
            for index in 0..count as usize {
                drop((*activates.add(index)).take());
            }
            CoTaskMemFree(Some(activates as *const std::ffi::c_void));
        }
    }

    enumerated.is_ok() && count > 0
}

/// Whether a decoder for one audio codec is registered.
///
/// The same question as the video one above and asked the same way, of the audio category
/// rather than the video's: an input type of the codec and the machine's own answer, with the
/// activation array freed whether or not anything was found. What it answers for a hover is
/// asked by playing rather than by enumerating (see `video_player::audio_probe`), because a
/// registered decoder is not the whole of whether a file plays.
fn audio_decoder(subtype: &GUID) -> bool {
    if !mf_started() {
        return false;
    }

    let input = MFT_REGISTER_TYPE_INFO {
        guidMajorType: MFMediaType_Audio,
        guidSubtype: *subtype,
    };

    let mut activates: *mut Option<IMFActivate> = std::ptr::null_mut();
    let mut count: u32 = 0;

    let enumerated = unsafe {
        MFTEnumEx(
            MFT_CATEGORY_AUDIO_DECODER,
            DECODER_ENUM_FLAGS,
            Some(&input),
            None,
            &mut activates,
            &mut count,
        )
    };

    unsafe {
        if !activates.is_null() {
            for index in 0..count as usize {
                drop((*activates.add(index)).take());
            }
            CoTaskMemFree(Some(activates as *const std::ffi::c_void));
        }
    }

    enumerated.is_ok() && count > 0
}

/// Whether a WIC codec claims one of `mime_types`.
///
/// A codec extension announces itself as a component of the imaging factory rather than as
/// a file, so this is "is a decoder for this kind of picture installed" asked of the place
/// it is installed to. Nothing is decoded and no picture is needed: what is listed is what
/// a `.heic` would be opened by.
fn image_codec(mime_types: &[&str]) -> bool {
    // The imaging factory is a COM object, so the thread asking for it has to be in an
    // apartment. This is asked here rather than left to whichever probe ran first: the
    // three groups are three calls, and a caller is free to ask for this one alone.
    initialize_apartment();

    let Some(factory) = crate::readers::wic_image::factory() else {
        return false;
    };

    let Ok(components) = (unsafe {
        factory
            .CreateComponentEnumerator(WICDecoder.0 as u32, WICComponentEnumerateDefault.0 as u32)
    }) else {
        return false;
    };

    let mut batch: [Option<IUnknown>; 8] = std::array::from_fn(|_| None);

    loop {
        let mut fetched: u32 = 0;
        let result = unsafe { components.Next(&mut batch, Some(&mut fetched)) };

        if result.is_err() || fetched == 0 {
            return false;
        }

        for entry in batch.iter().take(fetched as usize).flatten() {
            let Ok(info) = entry.cast::<IWICBitmapCodecInfo>() else {
                continue;
            };

            if mime_types.iter().any(|mime| component_reports(&info, mime)) {
                return true;
            }
        }

        // `S_FALSE` is the last batch there is.
        if result == windows::Win32::Foundation::S_FALSE {
            return false;
        }
    }
}

/// Whether one component's own MIME types include `mime`.
fn component_reports(info: &IWICBitmapCodecInfo, mime: &str) -> bool {
    let mut buffer = [0u16; 256];
    let mut written: u32 = 0;

    if unsafe { info.GetMimeTypes(&mut buffer, &mut written) }.is_err() {
        return false;
    }

    let end = buffer
        .iter()
        .position(|unit| *unit == 0)
        .unwrap_or(buffer.len());
    let reported = String::from_utf16_lossy(&buffer[..end]);

    // A component reports the types it claims as one comma-separated list — the HEIF
    // extension answers with `image/heic,image/heif,image/avci,…` and a JPEG one with
    // `image/jpeg,image/jpe,image/jpg` — so a match is against a whole entry rather than
    // against the string. The comparison is case-insensitive because the names are not
    // written the same way twice: the JPEG XL codec reports `image/JXL`.
    reported
        .split(',')
        .any(|entry| entry.trim().eq_ignore_ascii_case(mime))
}

/// Whether an application that answers to `prog_id` is registered, which is the same
/// question the render tier asks before it drives one — and, unlike asking for an engine,
/// it does not start one.
///
/// It is asked about the Office applications by the tray's menu and by the question that
/// decides where a document's page comes from: one whose own application is here is drawn
/// by it, and one whose application is missing is drawn by the render engine beside it
/// (see `office_formats::app_installed`).
pub(crate) fn prog_id_installed(prog_id: &str) -> bool {
    let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe { CLSIDFromProgID(windows::core::PCWSTR(wide.as_ptr())).is_ok() }
}

/// Whether a WinRT class is registered on this machine.
fn winrt_class_registered(class: &str) -> bool {
    let wide: Vec<u16> = class.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let mut key = HKEY::default();
        if RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            windows::core::PCWSTR(wide.as_ptr()),
            0,
            KEY_READ,
            &mut key,
        )
        .is_err()
        {
            return false;
        }

        let _ = RegCloseKey(key);
    }

    true
}

/// The media stack, started once for the process or never.
static MEDIA_FOUNDATION: OnceLock<bool> = OnceLock::new();

/// Whether FFmpeg is installed, kept between hovers and cleared by [`refresh`].
static FFPLAY: Lazy<Mutex<Option<bool>>> = Lazy::new(|| Mutex::new(None));

/// Whether both of FFmpeg's programs a sound's peak needs are here — the meter that measures it
/// and the player that applies it — kept the way the player's own answer is and cleared by the
/// same [`refresh`].
static NORMALIZE: Lazy<Mutex<Option<bool>>> = Lazy::new(|| Mutex::new(None));
