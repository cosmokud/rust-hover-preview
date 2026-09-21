//! The picture formats this app's own decoder does not read and Windows does: HEIF
//! (`.heic`, `.heif`), AVIF (`.avif`) and JPEG XL (`.jxl`).
//!
//! Nothing is bundled for them and nothing is installed beside the app. What decodes
//! them is the Windows Imaging Component — the codec surface Windows itself draws
//! Explorer's thumbnails with — through whichever codec the machine has: the **HEIF
//! Image Extension** for `.heic` and `.avif`, which is what needs the **HEVC Video
//! Extensions** for one and the **AV1 Video Extension** for the other, and the
//! **JPEG XL Image Extension** for `.jxl`. A machine without the codec is a machine
//! with no preview for that format, which is the answer a file that will not decode
//! gets; the extensions are named in the README, and the app never sends anyone to
//! the Store by itself.
//!
//! A codec is asked for the box the layout planned rather than for the file's own
//! size, so a hover onto a forty-megapixel photograph costs what its preview costs.
//! That is the same shape the PDF page and the Office page take, and it is the reason
//! this path exists at all rather than a decoder being compiled in: one of these
//! formats is patent-encumbered, a bundled codec for another is the size of this
//! application, and the codec Windows ships is already on the machine, is
//! hardware-accelerated where the machine has that, and will decode at the size it is
//! asked for.
//!
//! Nothing here is allowed to take the app down with it. A codec that is missing, a
//! file that will not open, a frame that will not decode and a picture past the
//! budget are all `None`, which is the answer a decoder that fails gives; nothing in
//! this module panics, and the budget every other reader is handed is asked before
//! this one allocates.
//!
//! What is *not* done here is orientation. A photograph from a phone carries the
//! rotation it is to be shown at — in the container, in its EXIF, or in both — and
//! neither this module nor the rest of the app applies one to any format yet, so a
//! picture whose bytes are stored sideways is drawn sideways. What shape that fix
//! takes is settled by what the codecs do with the container's own transform, which
//! is a question for a file written by a phone rather than for a guess.

use crate::config::frame_bytes_within_budget;
use std::cell::RefCell;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::GENERIC_READ;
use windows::Win32::Graphics::Imaging::{
    CLSID_WICImagingFactory2, GUID_WICPixelFormat32bppBGRA, IWICBitmap, IWICBitmapFrameDecode,
    IWICBitmapScaler, IWICBitmapSource, IWICBitmapSourceTransform, IWICFormatConverter,
    IWICImagingFactory, WICBitmapDitherTypeNone, WICBitmapInterpolationModeHighQualityCubic,
    WICBitmapPaletteTypeCustom, WICBitmapTransformRotate0, WICDecodeMetadataCacheOnDemand,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
};

/// The names a codec is asked for.
///
/// They are entries of the image list as well — that list is what decides a hover is
/// a picture at all, and these are pictures — so this is only which pictures the
/// codec has to be asked about rather than the decoder every other picture goes
/// through.
const CODEC_EXTENSIONS: &[&str] = &["avif", "heic", "heif", "jxl"];

thread_local! {
    /// The imaging factory this thread asks its codecs through.
    ///
    /// WIC is a COM API: the object belongs to the apartment that made it, so it is
    /// made per thread and kept for as long as the thread lives. It is not a decode
    /// state — a factory is what asks for codecs, and a codec is created, used and
    /// dropped within one decode.
    static FACTORY: RefCell<Option<IWICImagingFactory>> = const { RefCell::new(None) };
}

/// Initialize the apartment this module's calls need.
///
/// Every thread that opens a codec has to call this once before its first call, the
/// way the PDF module's threads do for `Windows.Data.Pdf`.
pub fn initialize_apartment() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
    }
}

/// Whether the format is one Windows has to be asked about.
pub fn is_codec_file(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase())
        .is_some_and(|extension| CODEC_EXTENSIONS.contains(&extension.as_str()))
}

/// A picture's own size, which is the size the layout places and the frame is decoded
/// into.
///
/// `None` is a file no codec claims — the extension is not installed — or one that is
/// not a picture of that kind at all, and both are the same answer: no preview.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let frame = open_frame(path)?;
    frame_size(&frame)
}

/// A picture decoded to exactly `width` by `height`, in the order the preview's own
/// frame is composed in: BGRA, top-down, four bytes to the pixel.
///
/// The size is the box the layout came out with rather than the picture's own, so
/// what the codec is asked to produce is the preview and not the file.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let frame = open_frame(path)?;
    let (source_width, source_height) = frame_size(&frame)?;

    // The budget every other reader is handed before it allocates: what a file may
    // ask for is bounded the same way whatever the file turned out to be, so a
    // picture whose own frame is past it is answered with no preview. It is asked of
    // the file's shape rather than of the preview's, because the preview's is this
    // side's to choose and the file's is not.
    frame_bytes_within_budget(source_width, source_height, 4)?;

    scaled_pixels(&frame, width, height)
}

/// The imaging factory this thread asks through, made on first use.
fn factory() -> Option<IWICImagingFactory> {
    FACTORY.with(|slot| {
        let mut slot = slot.borrow_mut();

        if slot.is_none() {
            *slot = unsafe {
                CoCreateInstance(&CLSID_WICImagingFactory2, None, CLSCTX_INPROC_SERVER).ok()
            };
        }

        slot.clone()
    })
}

/// The file's first frame, read by whichever codec claims the file.
///
/// What a file is, is its own header's business rather than its name's: WIC is asked
/// for the file, so a `.heic` holding AV1 and a `.avif` holding HEVC are each read by
/// the codec that knows them — and one no codec claims, which is what a machine
/// without the extension answers with, is `None` rather than an error to be told
/// apart from a file that will not decode.
///
/// The path is handed over as it is, verbatim prefix and all, because the codec opens
/// a file with the file API rather than through a broker: every path form the rest of
/// the app produces is one it accepts, the long one and the share included — which is
/// what the module's test asks, in the form the Shell hands a hovered path over.
fn open_frame(path: &Path) -> Option<IWICBitmapFrameDecode> {
    let factory = factory()?;
    let wide: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    unsafe {
        let decoder = factory
            .CreateDecoderFromFilename(
                PCWSTR(wide.as_ptr()),
                None,
                GENERIC_READ,
                WICDecodeMetadataCacheOnDemand,
            )
            .ok()?;

        decoder.GetFrame(0).ok()
    }
}

/// A frame's own size in pixels.
fn frame_size(frame: &IWICBitmapFrameDecode) -> Option<(u32, u32)> {
    let (mut width, mut height) = (0u32, 0u32);
    unsafe { frame.GetSize(&mut width, &mut height) }.ok()?;

    (width > 0 && height > 0).then_some((width, height))
}

/// The frame's pixels at the size the preview wants.
///
/// The codec is asked to do the scaling where it can, because that is the whole point
/// of going through it: a phone's photograph is tens of megapixels and a preview is a
/// fraction of one, and a codec asked for the preview's own size decodes the fraction.
/// A codec without that interface is scaled by WIC instead, and a codec whose own
/// pixels are not the format the preview is composed in is converted on the way — so
/// what comes back is four bytes to the pixel either way.
fn scaled_pixels(frame: &IWICBitmapFrameDecode, width: u32, height: u32) -> Option<Vec<u8>> {
    match native_scaled_bitmap(frame, width, height) {
        // The codec produced the preview's own size, so what is left is the copy.
        Some(bitmap) if bitmap_size(&bitmap) == Some((width, height)) => {
            copy_pixels(&bitmap, width, height)
        }

        // The codec produced a size nearer the file's than the preview's, so the
        // scaler takes the rest of the way. What it produced is already the format
        // the preview wants, which is why no converter is made for it.
        Some(bitmap) => {
            let factory = factory()?;
            scale_pixels(&factory, &bitmap, width, height)
        }

        // No scaled decode to be had: the frame itself is converted into the
        // preview's format and scaled from there.
        None => {
            let factory = factory()?;
            let converter = convert_to_bgra(&factory, frame)?;

            scale_pixels(&factory, &converter, width, height)
        }
    }
}

/// The codec's own scaled decode, where it has one.
///
/// `None` is every way a codec can decline, and every one of them is answered by the
/// frame's own pixels instead: no transform interface at all, a size it can produce
/// that is smaller than the preview asked for, a pixel format it will not produce,
/// or a decode that fails.
fn native_scaled_bitmap(
    frame: &IWICBitmapFrameDecode,
    width: u32,
    height: u32,
) -> Option<IWICBitmap> {
    let transform = frame.cast::<IWICBitmapSourceTransform>().ok()?;

    // The size it will decode to that is nearest the preview's: the interface is
    // asked for the size wanted and answers with the one it will do, which is
    // usually a step of the codec's own — half, a quarter — and never the file's
    // own size where the file is larger.
    let (mut scaled_width, mut scaled_height) = (width, height);
    unsafe { transform.GetClosestSize(&mut scaled_width, &mut scaled_height) }.ok()?;

    // A size below the preview's would be the codec doing the scaler's job worse
    // than the scaler does it, so it is left to the scaler.
    if scaled_width < width || scaled_height < height {
        return None;
    }

    // The format it will produce: the preview is composed in BGRA, and a codec that
    // would hand back something else is taken the long way round instead.
    let mut format = GUID_WICPixelFormat32bppBGRA;
    unsafe { transform.GetClosestPixelFormat(&mut format) }.ok()?;
    if format != GUID_WICPixelFormat32bppBGRA {
        return None;
    }

    let stride = scaled_width.checked_mul(4)?;
    let mut pixels = vec![0u8; stride as usize * scaled_height as usize];

    unsafe {
        transform
            .CopyPixels(
                std::ptr::null(),
                scaled_width,
                scaled_height,
                &GUID_WICPixelFormat32bppBGRA,
                WICBitmapTransformRotate0,
                stride,
                &mut pixels,
            )
            .ok()?
    };

    let factory = factory()?;
    unsafe {
        factory
            .CreateBitmapFromMemory(
                scaled_width,
                scaled_height,
                &GUID_WICPixelFormat32bppBGRA,
                stride,
                &pixels,
            )
            .ok()
    }
}

/// The frame through WIC's format converter: whatever the codec's own pixels are —
/// ten-bit, wide-gamut, with an alpha channel or without one — turned into the 32-bit
/// BGRA the preview's frame is composed in.
fn convert_to_bgra(
    factory: &IWICImagingFactory,
    frame: &IWICBitmapFrameDecode,
) -> Option<IWICFormatConverter> {
    let converter: IWICFormatConverter = unsafe { factory.CreateFormatConverter() }.ok()?;

    unsafe {
        converter
            .Initialize(
                frame,
                &GUID_WICPixelFormat32bppBGRA,
                WICBitmapDitherTypeNone,
                None,
                0.0,
                WICBitmapPaletteTypeCustom,
            )
            .ok()?
    };

    Some(converter)
}

/// A source scaled into exactly the preview's box, which is the last step of every
/// path: a codec that decoded nearer the file's size, and one whose pixels had to be
/// converted.
fn scale_pixels(
    factory: &IWICImagingFactory,
    source: &IWICBitmapSource,
    width: u32,
    height: u32,
) -> Option<Vec<u8>> {
    let scaler: IWICBitmapScaler = unsafe { factory.CreateBitmapScaler() }.ok()?;

    unsafe {
        scaler
            .Initialize(
                source,
                width,
                height,
                WICBitmapInterpolationModeHighQualityCubic,
            )
            .ok()?
    };

    copy_pixels(&scaler, width, height)
}

/// The pixels of a source that is exactly the preview's size.
fn copy_pixels(source: &IWICBitmapSource, width: u32, height: u32) -> Option<Vec<u8>> {
    let stride = width.checked_mul(4)?;
    let mut pixels = vec![0u8; stride as usize * height as usize];

    unsafe { source.CopyPixels(std::ptr::null(), stride, &mut pixels) }.ok()?;

    Some(pixels)
}

/// A source's own size, asked the way a frame's is.
fn bitmap_size(source: &IWICBitmapSource) -> Option<(u32, u32)> {
    let (mut width, mut height) = (0u32, 0u32);
    unsafe { source.GetSize(&mut width, &mut height) }.ok()?;

    Some((width, height))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;
    use windows::Graphics::Imaging::{BitmapAlphaMode, BitmapEncoder, BitmapPixelFormat};
    use windows::Storage::Streams::{DataReader, InMemoryRandomAccessStream};

    /// The color the probe picture is, and what it is allowed to come back as: the
    /// codec is a lossy one, so a solid color is read back near its own value rather
    /// than at it.
    const PROBE_RED: u8 = 200;
    const PROBE_GREEN: u8 = 90;
    const PROBE_BLUE: u8 = 40;
    const PROBE_TOLERANCE: i32 = 12;

    /// The names the codec is asked about, and the ones it is not.
    #[test]
    fn claims_the_names_windows_has_a_codec_for() {
        for name in [
            "picture.avif",
            "picture.HEIC",
            "picture.heif",
            "picture.jxl",
        ] {
            assert!(is_codec_file(Path::new(name)), "{name}");
        }

        for name in ["picture.png", "picture.jpg", "picture.heic.txt", "heic"] {
            assert!(!is_codec_file(Path::new(name)), "{name}");
        }
    }

    /// A file no codec claims is answered with no picture rather than with a guess,
    /// which is what a machine without the extension installed answers with as well.
    #[test]
    fn a_file_no_codec_can_read_is_no_picture() {
        let folder = std::env::temp_dir().join("rust-hover-preview-wic-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("not-a-picture.heic");
        std::fs::write(&path, b"this is not a HEIF file at all").expect("a written file");

        assert_eq!(dimensions(&path), None);
        assert_eq!(decode(&path, 16, 16), None);
        assert_eq!(dimensions(&folder.join("missing.heic")), None);
    }

    /// A picture of a format the codec Windows has wrote is read back by this module,
    /// at the box the preview asked for rather than at the size the file is.
    ///
    /// The file is made here rather than carried in the repository, because the
    /// encoders belong to the same extension package the decoders do: a machine that
    /// can read a `.heic` is a machine that can write one. A machine that cannot is a
    /// machine with no codec to test, so the test has nothing to say about it and
    /// returns rather than failing — the same way the Office diagnostic stands down on
    /// a machine without the application.
    ///
    /// What it wrote is left where it is, so that the app-wide hover probe can be
    /// pointed at a file of a kind this machine can otherwise only get from a phone:
    /// `$env:RHP_APP_PROBE = "$env:TEMP\rust-hover-preview-wic-tests\probe.heic"`.
    #[test]
    fn decodes_a_heic_written_by_the_windows_codec() {
        initialize_apartment();

        let (file_width, file_height) = (64u32, 48u32);
        let Some(path) = write_probe_heic(file_width, file_height) else {
            return;
        };

        assert_eq!(
            dimensions(&path),
            Some((file_width, file_height)),
            "the frame's own size is what the layout is placed from"
        );

        // A quarter of the file in each direction, which is the point of the path: what
        // the codec is asked to produce is the preview rather than the file.
        let (preview_width, preview_height) = (16u32, 12u32);
        let pixels = decode(&path, preview_width, preview_height).expect("a decoded preview");

        assert_eq!(
            pixels.len(),
            (preview_width * preview_height * 4) as usize,
            "four bytes to the pixel, at the box that was asked for"
        );

        for (index, pixel) in pixels.chunks_exact(4).enumerate() {
            let (blue, green, red, alpha) = (pixel[0], pixel[1], pixel[2], pixel[3]);
            let distance =
                |channel: u8, expected: u8| (i32::from(channel) - i32::from(expected)).abs();

            assert!(
                distance(red, PROBE_RED) <= PROBE_TOLERANCE
                    && distance(green, PROBE_GREEN) <= PROBE_TOLERANCE
                    && distance(blue, PROBE_BLUE) <= PROBE_TOLERANCE,
                "pixel {index} came back as ({red}, {green}, {blue})"
            );
            assert_eq!(
                alpha, 255,
                "an opaque picture stays opaque at pixel {index}"
            );
        }

        // The Shell hands a path over in its verbatim form — `\\?\C:\…` — and a codec
        // is opened with the file API rather than through a broker, so the form the app
        // actually hovers is asked for as well as the plain one.
        if path.is_absolute() && !path.to_string_lossy().starts_with(r"\\") {
            let verbatim = PathBuf::from(format!(r"\\?\{}", path.display()));

            assert_eq!(
                decode(&verbatim, preview_width, preview_height).as_deref(),
                Some(pixels.as_slice()),
                "a verbatim path is read the same way"
            );
        }
    }

    /// A `.heic` of a solid color, written by the HEIF codec's own encoder.
    ///
    /// `None` is a machine whose codec cannot write one — the extension is not
    /// installed, or it is a version that reads without writing — which is what the
    /// test above stands down on.
    fn write_probe_heic(width: u32, height: u32) -> Option<PathBuf> {
        let format = BitmapEncoder::HeifEncoderId().ok()?;
        let pixels: Vec<u8> = std::iter::repeat([PROBE_RED, PROBE_GREEN, PROBE_BLUE, 255])
            .take((width * height) as usize)
            .flatten()
            .collect();

        let stream = InMemoryRandomAccessStream::new().ok()?;
        let encoder = BitmapEncoder::CreateAsync(format, &stream)
            .ok()?
            .get()
            .ok()?;
        encoder
            .SetPixelData(
                BitmapPixelFormat::Rgba8,
                BitmapAlphaMode::Straight,
                width,
                height,
                96.0,
                96.0,
                &pixels,
            )
            .ok()?;
        encoder.FlushAsync().ok()?.get().ok()?;

        let bytes = read_stream_bytes(&stream)?;
        let folder = std::env::temp_dir().join("rust-hover-preview-wic-tests");
        std::fs::create_dir_all(&folder).ok()?;
        let path = folder.join("probe.heic");
        std::fs::write(&path, bytes).ok()?;

        Some(path)
    }

    /// The bytes a WinRT stream holds, read the way the PDF module reads its own.
    fn read_stream_bytes(stream: &InMemoryRandomAccessStream) -> Option<Vec<u8>> {
        let length = stream.Size().ok()? as usize;
        if length == 0 {
            return None;
        }

        let input = stream.GetInputStreamAt(0).ok()?;
        let reader = DataReader::CreateDataReader(&input).ok()?;
        reader.LoadAsync(length as u32).ok()?.get().ok()?;

        let mut bytes = vec![0u8; length];
        reader.ReadBytes(&mut bytes).ok()?;

        Some(bytes)
    }
}
