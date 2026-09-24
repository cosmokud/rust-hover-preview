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
    IMFActivate, MFMediaType_Video, MFStartup, MFTEnumEx, MFVideoFormat_AV1, MFVideoFormat_H264,
    MFVideoFormat_HEVC, MFVideoFormat_MP4V, MFVideoFormat_MPEG2, MFVideoFormat_Theora,
    MFVideoFormat_VP90, MFVideoFormat_WMV3, MFSTARTUP_FULL, MFT_CATEGORY_VIDEO_DECODER,
    MFT_ENUM_FLAG, MFT_ENUM_FLAG_ASYNCMFT, MFT_ENUM_FLAG_HARDWARE, MFT_ENUM_FLAG_LOCALMFT,
    MFT_ENUM_FLAG_SYNCMFT, MFT_REGISTER_TYPE_INFO, MF_VERSION,
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

/// One row of the tray's `Codecs` submenu: what the engine or codec is called, and
/// whether this machine has it.
pub struct Row {
    pub name: &'static str,
    pub available: bool,
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
        },
        Row {
            name: "Windows Media Foundation",
            available: mf_started(),
        },
        Row {
            name: "H.264",
            available: video_decoder(&MFVideoFormat_H264),
        },
        // One row for the two, because one decoder answers for both: Windows' MPEG-4
        // Part 2 decoder is what plays a DivX or Xvid file as well.
        Row {
            name: "MPEG-4 / WMV",
            available: video_decoder(&MFVideoFormat_MP4V) || video_decoder(&MFVideoFormat_WMV3),
        },
        Row {
            name: "MPEG-2",
            available: video_decoder(&MFVideoFormat_MPEG2),
        },
        Row {
            name: "HEVC (H.265)",
            available: video_decoder(&MFVideoFormat_HEVC),
        },
        Row {
            name: "VP9",
            available: video_decoder(&MFVideoFormat_VP90),
        },
        Row {
            name: "AV1",
            available: video_decoder(&MFVideoFormat_AV1),
        },
        Row {
            name: "Theora (Ogg)",
            available: video_decoder(&MFVideoFormat_Theora),
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

    vec![
        Row {
            name: "HEIF (HEIC)",
            available: heif && video_decoder(&MFVideoFormat_HEVC),
        },
        Row {
            name: "AVIF",
            available: image_codec(AVIF_MIME_TYPES) || (heif && video_decoder(&MFVideoFormat_AV1)),
        },
        Row {
            name: "JPEG XL",
            available: image_codec(JXL_MIME_TYPES),
        },
        // The one picture here that needs nothing: this app carries its own libwebp,
        // which is what plays an animated one and what decodes a still one where the
        // codec is missing. The row stands for an engine that is in the binary, so it is
        // always there — and it is listed so that the group reads as the whole of what a
        // picture can be.
        Row {
            name: "WebP",
            available: true,
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
        },
        Row {
            name: "LibreOffice",
            available: crate::engines::libreoffice_render::available(),
        },
        Row {
            name: "ImageMagick",
            available: crate::engines::imagemagick_render::available(),
        },
        Row {
            name: "PeaZip",
            available: crate::engines::peazip_render::available(),
        },
        Row {
            name: "Microsoft Word",
            available: prog_id_installed("Word.Application"),
        },
        Row {
            name: "Microsoft Excel",
            available: prog_id_installed("Excel.Application"),
        },
        Row {
            name: "Microsoft PowerPoint",
            available: prog_id_installed("PowerPoint.Application"),
        },
        Row {
            name: "Windows PDF Engine",
            available: winrt_class_registered(PDF_ENGINE_CLASS),
        },
    ]
}

/// Ask again, which is what the tray's menu build does before it lists the answers.
///
/// Every other probe is asked on demand and answers from the machine at that moment. The
/// one answer that is kept is whether FFmpeg is installed, because it is asked on every
/// video hover and a hover must not go looking through the `PATH` for it — so opening the
/// menu is what lets a machine that has just been given FFmpeg start using it, rather than
/// a restart.
pub fn refresh() {
    if let Ok(mut cached) = FFPLAY.lock() {
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
        .map(|folder| folder.join(FFPLAY_NAME))
        .find(|player| player.is_file())
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
