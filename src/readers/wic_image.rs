//! The picture formats this app's own decoder does not read and Windows does: HEIF
//! (`.heic`, `.heif`, `.heics`, `.heifs`, `.hif`), AVIF (`.avif`, `.avifs`, `.avcs`), the AVC
//! still (`.avci`), JPEG XL (`.jxl`), JPEG XR (`.jxr`, `.wdp`), a still WebP (`.webp`), a
//! DirectDraw Surface (`.dds`), the camera raws the Raw Image Extension decodes
//! (`.dng`, `.nef`, `.cr3` and the rest), and the bitmap, icon and cursor resources
//! (`.bmp`, `.dib`, `.rle`, `.ico`, `.icon`, `.exif`, `.cur`).
//!
//! They are pictures the image list carries — the master bucket for everything Windows
//! decodes that *is* a picture — and this is the reader that asks the codec for them (see
//! `lists::DEFAULT_IMAGE_EXTENSIONS`).
//!
//! Nothing is bundled for them and nothing is installed beside the app. What decodes
//! them is the Windows Imaging Component — the codec surface Windows itself draws
//! Explorer's thumbnails with — through whichever codec the machine has: the **HEIF
//! Image Extension** for `.heic` and `.avif`, which is what needs the **HEVC Video
//! Extensions** for one and the **AV1 Video Extension** for the other, the **JPEG XL
//! Image Extension** for `.jxl`, and the **WebP Image Extension** for a WebP still. A
//! machine without the codec is a machine with no preview for that format, which is
//! the answer a file that will not decode gets; the extensions are named in the
//! README, and the app never sends anyone to the Store by itself.
//!
//! The WebP one is the codec a machine is most likely to be missing — Windows 11
//! comes with it and Windows 10 usually has to be given it — so a still WebP is the
//! one picture here that is not simply lost to a missing package: `webp_image` decodes
//! it with the libwebp the binary already carries, and this module is only the reader
//! asked first. The WebP that *moves* is not this module's at all: libwebp is asked
//! for it before anything else (see the animated reader in `preview_window`), so what
//! arrives here is the picture that does not move.
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
//!
//! A `.dds` is the one format here this app *does* carry a decoder of its own for (see
//! `dds_image`), and it is still asked of the codec first. What the codec gives it that
//! no decoder of this side's does is the size: a texture is asked for at the box the
//! layout planned and decoded there, while a decoder of this side's has to decode the
//! whole of it — a four-thousand-square texture is sixty-four megabytes of texels — and
//! scale that down afterwards. What the codec reads of a DDS is a smaller set than what
//! the format holds, which is what `dds_image` is for: BC1, BC2 and BC3 here, and the
//! uncompressed formats, BC4 and BC5 there.

use crate::config::config::frame_bytes_within_budget;
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
use windows::Win32::UI::Shell::SHCreateMemStream;

/// The names a codec is asked for.
///
/// They are entries of the image list as well — that list is what decides a hover is
/// a picture at all, and these are pictures — so this is only which pictures the
/// codec has to be asked about rather than the decoder every other picture goes
/// through. A name is here because no decoder in the binary carries the format: the
/// ones `image` reads (a PNG, a JPEG, a BMP, a TIFF, a TGA, a QOI, an EXR) are asked of
/// it and never of a codec, which is what keeps a picture whose decode is in hand from
/// costing a codec call as well.
///
/// The bitmap, icon and cursor shapes are here beside the codec packages they look
/// unlike, because `image` reads some of them under other names and not at all under
/// these: a `.dib` has no file header to sniff, and a `.cur` is a cursor rather than the
/// icon `image` can open.
const CODEC_EXTENSIONS: &[&str] = &[
    "3fr", "arw", "avcs", "avif", "avifs", "avci", "cr2", "cr3", "crw", "cur", "dcr", "dds", "dib",
    "dng", "erf", "exif", "fff", "heic", "heics", "heif", "heifs", "hif", "ico", "icon", "iiq",
    "jxl", "jxr", "k25", "kdc", "mef", "mos", "mrw", "nef", "nrw", "orf", "pef", "raf", "raw",
    "rle", "rw2", "rwl", "sr2", "srf", "srw", "wdp", "webp", "x3f",
];

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
///
/// What it gets that thread is an apartment of some kind, not necessarily the
/// multi-threaded one being asked for. A concurrency model is fixed once it is chosen, and
/// `main` has already put the tray's own main thread in a single-threaded apartment, so
/// there the request is refused with `RPC_E_CHANGED_MODE` and the refusal is discarded.
/// The factory is made in whatever apartment the thread already had, which is the one its
/// objects belong to either way.
///
/// Treating that refusal as a failure empties the tray's `Codecs` submenu on every machine
/// rather than the odd one: the factory is not made, and a factory that is not made reads
/// as a picture codec that is not installed, so those rows vanish from the menu on
/// hardware that decodes them perfectly well.
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
///
/// It is shared with the media engine's video path rather than made a second time
/// there: one WIC factory is one factory, and a thread that draws a video frame into a
/// WIC bitmap and a thread that decodes a picture into one are asking the same object
/// for the same thing (see `video_player`).
pub(crate) fn factory() -> Option<IWICImagingFactory> {
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
///
/// A `.dib` is the one name here the codec will not open by itself: it is a bitmap with its
/// fourteen-byte file header taken off, and this decoder reads that header rather than the
/// shape. So the header is put back in memory and the codec is handed the file as a stream
/// instead (see `bitmap_resource_bytes`).
fn open_frame(path: &Path) -> Option<IWICBitmapFrameDecode> {
    let factory = factory()?;

    if is_bitmap_resource(path) {
        let stream = unsafe { SHCreateMemStream(Some(&bitmap_resource_bytes(path)?)) }?;

        return unsafe {
            factory
                .CreateDecoderFromStream(&stream, std::ptr::null(), WICDecodeMetadataCacheOnDemand)
                .ok()?
                .GetFrame(0)
                .ok()
        };
    }

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

/// Whether the file is the bitmap the codec will not open by itself — see `open_frame`.
fn is_bitmap_resource(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("dib"))
}

/// A device-independent bitmap as the file the codec does open: its own bytes with the
/// fourteen-byte bitmap-file header put back in front of them.
///
/// The header is written for the one field the codec reads out of it rather than for what it
/// is: where the picture begins, which is the DIB header's own length, the masks a bitfield
/// header carries behind it, and the palette entries the header names — or one for every
/// colour a `bits`-deep palette can hold, where it names none.
///
/// shortcut: the whole bitmap is read into memory, which is what a `.dib` on a hover is; a
/// stream the codec could read in place is what to reach for if one ever turns up large.
fn bitmap_resource_bytes(path: &Path) -> Option<Vec<u8>> {
    let dib = std::fs::read(path).ok()?;

    let header = u32::from_le_bytes(dib.get(0..4)?.try_into().ok()?) as usize;
    let bits = u16::from_le_bytes(dib.get(14..16)?.try_into().ok()?);
    let compression = u32::from_le_bytes(dib.get(16..20)?.try_into().ok()?);
    let colours = u32::from_le_bytes(dib.get(32..36)?.try_into().ok()?);

    // A bitfield header carries its masks behind a 40-byte one, and names them inside a
    // longer one.
    let masks = match compression {
        3 | 6 if header == 40 => {
            if compression == 3 {
                12
            } else {
                16
            }
        }
        _ => 0,
    };
    let palette = if bits <= 8 && bits > 0 {
        match colours {
            0 => 1usize << bits,
            entries => entries as usize,
        }
    } else {
        0
    };

    let mut file = Vec::with_capacity(14 + dib.len());
    file.extend_from_slice(b"BM");
    file.extend_from_slice(&((14 + dib.len()) as u32).to_le_bytes());
    file.extend_from_slice(&[0u8; 4]);
    file.extend_from_slice(&((14 + header + masks + palette * 4) as u32).to_le_bytes());
    file.extend_from_slice(&dib);

    Some(file)
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
            "picture.dds",
            "picture.DDS",
            "picture.HEIC",
            "picture.heif",
            "picture.jxl",
            "picture.webp",
            "picture.WEBP",
            "picture.dib",
            "picture.cur",
            "picture.icon",
            "picture.jxr",
            "picture.wdp",
            "picture.heics",
            "picture.NEF",
            "picture.cr2",
            "picture.dng",
            "picture.raw",
        ] {
            assert!(is_codec_file(Path::new(name)), "{name}");
        }

        for name in [
            "picture.png",
            "picture.jpg",
            "picture.heic.txt",
            "heic",
            "picture.xcf",
            "picture.miff",
        ] {
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

        for (index, pixel) in pixels.as_chunks::<4>().0.iter().enumerate() {
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
        let pixels: Vec<u8> = std::iter::repeat_n(
            [PROBE_RED, PROBE_GREEN, PROBE_BLUE, 255],
            (width * height) as usize,
        )
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

    /// A `.dib` and a `.cur` — the two shapes of a picture in the master list that a codec is
    /// asked about and that has no package of its own to be missing: a bitmap resource and a
    /// cursor. Both are written here by hand rather than carried in the repository, because
    /// what is being tested is that the codec reads what the format says, and a fixture off a
    /// disc would be a second thing to keep in step with it.
    ///
    /// The bitmap is the one every Windows has a codec for, and the one whose file header this
    /// app writes itself, so it is asserted rather than skipped. A cursor is asked of a codec
    /// that a machine could be without — the answer then is a picture with no preview, which is
    /// what this reader gives for a missing package as well — so a machine whose codec answers
    /// nothing to it returns rather than failing (see
    /// `decodes_a_heic_written_by_the_windows_codec`).
    #[test]
    fn decodes_a_bitmap_resource_and_a_cursor_where_the_codec_is_there() {
        initialize_apartment();

        let folder = std::env::temp_dir().join("rust-hover-preview-wic-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let (width, height) = (2u32, 2u32);
        let probes = [
            ("probe.dib", device_independent_bitmap(width, height), true),
            ("probe.cur", cursor(width, height), false),
        ];

        for (name, bytes, everywhere) in probes {
            let path = folder.join(name);
            std::fs::write(&path, bytes).expect("a written picture");

            assert!(
                is_codec_file(&path),
                "`{name}` is a name the codec is asked about"
            );

            let Some(measured) = dimensions(&path) else {
                assert!(
                    !everywhere,
                    "`{name}` is read by the codec every Windows has"
                );
                continue;
            };

            assert_eq!(
                measured,
                (width, height),
                "`{name}` is measured from its own frame"
            );

            let pixels = decode(&path, width, height)
                .unwrap_or_else(|| panic!("`{name}` decodes at the box it was asked for"));
            assert_eq!(
                pixels.len(),
                (width * height * 4) as usize,
                "four bytes to the pixel for `{name}`"
            );
            assert_eq!(
                &pixels[..4],
                &[PROBE_BLUE, PROBE_GREEN, PROBE_RED, 255],
                "`{name}` comes back in the order the frame is composed in"
            );
        }
    }

    /// A device-independent bitmap of one colour: the `BITMAPINFOHEADER` and the rows under it,
    /// which is what a `.bmp` is once its own file header is taken off, and what a `.dib` is.
    fn device_independent_bitmap(width: u32, height: u32) -> Vec<u8> {
        // A row of a 24-bit bitmap is padded to four bytes, and stored bottom-up — the colour
        // here is the same in every pixel, so which end the rows are written from does not
        // matter to what is read back.
        let row = (width as usize * 3).div_ceil(4) * 4;
        let mut bytes = bitmap_info_header(width, height, 24);

        for _ in 0..height {
            let mut pixels = vec![0u8; row];
            for column in 0..width as usize {
                pixels[column * 3..column * 3 + 3].copy_from_slice(&[
                    PROBE_BLUE,
                    PROBE_GREEN,
                    PROBE_RED,
                ]);
            }
            bytes.extend_from_slice(&pixels);
        }

        bytes
    }

    /// A cursor of one colour: the icon directory, the one entry that says where its image is,
    /// and the image itself. A cursor and an icon are the same file with one field of that
    /// directory saying which it is, and the image is a bitmap under a mask — which is why the
    /// header's height is twice the picture's.
    fn cursor(width: u32, height: u32) -> Vec<u8> {
        let image = {
            let mut bytes = bitmap_info_header(width, height * 2, 32);

            for _ in 0..height {
                for _ in 0..width {
                    bytes.extend_from_slice(&[PROBE_BLUE, PROBE_GREEN, PROBE_RED, 255]);
                }
            }

            // The mask, all clear: every pixel is drawn through rather than left out.
            let mask_row = (width as usize).div_ceil(32) * 4;
            bytes.extend(std::iter::repeat_n(0u8, mask_row * height as usize));

            bytes
        };

        // `ICONDIR` — reserved, 2 for a cursor, one image — then the entry: its shape, its
        // palette, the hotspot (zero, the corner) and where its image starts behind the two.
        let mut bytes = Vec::with_capacity(22 + image.len());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&2u16.to_le_bytes());
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&[width as u8, height as u8, 0, 0]);
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&(image.len() as u32).to_le_bytes());
        bytes.extend_from_slice(&22u32.to_le_bytes());
        bytes.extend_from_slice(&image);

        bytes
    }

    /// A `BITMAPINFOHEADER` of a picture that is not compressed, which is the whole of the header
    /// every bitmap resource and every image in an icon opens with.
    fn bitmap_info_header(width: u32, height: u32, bits: u16) -> Vec<u8> {
        let stride = (width as usize * bits as usize).div_ceil(8).div_ceil(4) * 4;
        let image_size = stride * height as usize;

        let mut bytes = Vec::with_capacity(40);
        bytes.extend_from_slice(&40u32.to_le_bytes());
        bytes.extend_from_slice(&(width as i32).to_le_bytes());
        bytes.extend_from_slice(&(height as i32).to_le_bytes());
        bytes.extend_from_slice(&1u16.to_le_bytes());
        bytes.extend_from_slice(&bits.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&(image_size as u32).to_le_bytes());
        bytes.extend_from_slice(&[0u8; 16]);

        bytes
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
