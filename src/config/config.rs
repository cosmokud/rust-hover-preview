use configparser::ini::Ini;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::collections::HashSet;
use std::fs;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::Duration;

use crate::config::theme_files;
use crate::formats::archive_formats::{sanitize_archive_extensions, DEFAULT_ARCHIVE_EXTENSIONS};
use crate::formats::design_formats::{
    sanitize_design_extensions, DEFAULT_DESIGN_EXTENSIONS, DESIGN_EXTENSIONS_BEFORE_AI,
    DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE, DESIGN_EXTENSIONS_WITH_CDR,
};
use crate::formats::font_formats::{sanitize_font_extensions, DEFAULT_FONT_EXTENSIONS};
use crate::formats::image_formats::{
    sanitize_image_extensions, DEFAULT_IMAGE_EXTENSIONS, IMAGE_EXTENSIONS_BEFORE_AVCI,
    IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS, IMAGE_EXTENSIONS_BEFORE_DDS,
    IMAGE_EXTENSIONS_BEFORE_SVG, IMAGE_EXTENSIONS_WITH_SVG,
};
use crate::formats::libre_formats::{
    sanitize_libre_extensions, DEFAULT_LIBRE_EXTENSIONS,
    LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ,
};
use crate::formats::magick_formats::{
    sanitize_magick_extensions, DEFAULT_MAGICK_EXTENSIONS, MAGICK_EXTENSIONS_BEFORE_THE_REST,
};
use crate::formats::office_formats::{sanitize_office_extensions, DEFAULT_OFFICE_EXTENSIONS};
use crate::formats::peazip_formats::{
    sanitize_peazip_extensions, DEFAULT_PEAZIP_EXTENSIONS, PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS,
};
use crate::formats::text_formats::{
    sanitize_extensions, sanitize_names, DEFAULT_TEXT_EXTENSIONS, DEFAULT_TEXT_NAMES,
};
use crate::formats::vector_formats::{
    sanitize_vector_extensions, DEFAULT_VECTOR_EXTENSIONS, VECTOR_EXTENSIONS_BEFORE_SVG,
    VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS,
};
use crate::formats::video_formats::{sanitize_video_extensions, DEFAULT_VIDEO_EXTENSIONS};
use crate::readers::tone_map::Curve;

const CONFIG_SECTION: &str = "settings";
/// The image extension list lives in its own section so the one long value stays
/// easy to find and edit by hand.
const IMAGE_SECTION: &str = "image";
/// The video extension list lives in its own section for the same reason.
const VIDEO_SECTION: &str = "video";
/// The text-preview extension list lives in its own section so the one long
/// value stays easy to find and edit by hand.
const TEXT_SECTION: &str = "text";
/// The archive extension list lives in its own section for the same reason.
const ARCHIVE_SECTION: &str = "archive";
/// The office extension list lives in its own section for the same reason.
const OFFICE_SECTION: &str = "office";
/// The font extension list lives in its own section for the same reason.
const FONT_SECTION: &str = "font";
/// The design extension list lives in its own section for the same reason.
const DESIGN_SECTION: &str = "design";
/// The vector extension list lives in its own section for the same reason.
const VECTOR_SECTION: &str = "vector";
/// The list of documents the render engine is asked about lives in its own section for the
/// same reason: what this app hands to LibreOffice rather than reading itself.
const LIBRE_SECTION: &str = "libre";
/// And the list of pictures the ImageMagick engine is asked about, for the same reason:
/// what this app hands to it rather than reading itself.
const MAGICK_SECTION: &str = "magick";
/// And the list of archives the PeaZip engine is asked about, for the same reason: the
/// formats this app hands to it rather than reading itself.
const PEAZIP_SECTION: &str = "peazip";
pub const DEFAULT_WEBP_PLAYBACK_FPS: u32 = 90;
pub const MAX_WEBP_PLAYBACK_FPS: u32 = 90;
/// The volume a video is played at unless the file says otherwise: silent, so a hover
/// never makes a sound the pointer did not ask for.
pub const DEFAULT_VIDEO_VOLUME: u32 = 0;
pub const DEFAULT_PREVIEW_SCALE_PERCENT: u32 = 100;
pub const MIN_PREVIEW_SCALE_PERCENT: u32 = 1;
pub const MAX_PREVIEW_SCALE_PERCENT: u32 = 1000;
/// The share of its own size a video is drawn at unless asked otherwise.
///
/// A video is measured by a probe and drawn from its first frame, and that frame is a
/// bitmap like a picture — so the share means the same thing here as it does there, a
/// percentage of the size the file asks for, and the setting starts at the same place
/// the picture's does.
pub const DEFAULT_VIDEO_SCALE_PERCENT: u32 = 100;
/// The share of its own size an animated picture is drawn at unless asked otherwise.
///
/// An animation is drawn from the frames it decodes, which are bitmaps like a picture's,
/// so the share means the same thing here as it does for one — a percentage of the size
/// the file asks for — and the setting starts at the same place the picture's does.
pub const DEFAULT_ANIMATED_SCALE_PERCENT: u32 = 100;
/// The share of the display a font specimen is drawn at unless asked otherwise.
///
/// A font has no size it asks to be drawn at — a file holds outlines, and the text they are
/// drawn as is whatever size a caller asks for — so what the preview is sized by is the box
/// `font_preview` measures it at, narrowed to this share of the room the display has: half
/// of the room is where a specimen starts, the same starting point an SVG document has.
pub const DEFAULT_FONT_SCALE_PERCENT: u32 = 50;
/// Which face of a collection a specimen is drawn from, as the menu numbers faces: `1` is
/// the first face a `.ttc` holds, which is what every collection starts at.
pub const DEFAULT_TTC_FACE: u32 = 1;
/// The last face the setting can name.
///
/// A collection a machine carries is one to four faces — a family's regular, bold, italic
/// and bold italic — and the widest ones, the pan-CJK collections, hold ten; the setting
/// stops at the tenth rather than at whatever a crafted file could claim, since what the
/// menu offers is the whole range the setting holds.
pub const MAX_TTC_FACE: u32 = 10;
pub const DEFAULT_TEXT_FONT_SCALE_PERCENT: u32 = 125;
pub const MIN_TEXT_FONT_SCALE_PERCENT: u32 = 1;
pub const MAX_TEXT_FONT_SCALE_PERCENT: u32 = 1000;
pub const DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS: f32 = 40.0;
pub const MAX_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS: f32 = 1000.0;
/// How long the pointer rests on a file before a preview is put up for it, in
/// milliseconds.
///
/// `0` is no wait at all, which is where the app starts: a hover is answered when it
/// happens, and a delay is for the machine that would rather it were not — which the
/// menu offers as a choice rather than making as one.
pub const DEFAULT_HOVER_DELAY_MS: u64 = 0;
/// How long the same file waits before a preview of it is put up again, in milliseconds.
///
/// A pointer crosses the file it has just left again on its way somewhere else as often
/// as it comes back to it, so a preview that has just been taken down is not put back up
/// for this long: a fifth of a second is enough for the hand that was going elsewhere,
/// while a hand that turns back is answered before the wait is over. `0` is a file that
/// previews again the moment it is hovered.
pub const DEFAULT_SAME_FILE_REHOVER_DELAY_MS: u64 = 200;
/// How long the pointer must be still before a preview is put up for what it is on,
/// in milliseconds.
///
/// A pointer crossing a list is on a new file every few dozen milliseconds, and what a
/// hover is about is the file the hand comes to rest on: while it is still moving, the
/// file under the cursor is one it is passing rather than the one being asked about. The
/// wait this names is that one, and `0` is the setting that asks for none of it — a
/// preview is put up for a new file even while the hand is still on its way to it, which
/// is where the app starts — so the machine that would rather see nothing until the
/// pointer has settled is the one that gives this a value. The pointer is the whole of
/// its subject: a keyboard preview is the keyboard's own and is not gated by it.
pub const DEFAULT_SETTLING_DELAY_MS: u64 = 0;
/// How often the app looks at the pointer's world while Explorer has focus, in
/// milliseconds: the loop's tick, and with it how soon a move is answered.
///
/// It is the one number that trades responsiveness against what the app costs while
/// it works: every look is a crossing into Explorer — the cursor is read, the item
/// under it is resolved, and the engines are asked about — and the loop never sleeps
/// for longer than this while a folder window is in front. A system tick is the floor
/// a wait can be honoured at, so the value is meant in whole ones: fifteen is one, and
/// the ladder the menu offers counts up from there. See `MIN_TICK_MS` and `MAX_TICK_MS`.
pub const DEFAULT_TICK_MS: u64 = 15;
/// The least a hand-edited tick is reduced to, in milliseconds: below half a system
/// tick the number asks for a wait the system cannot honour, and the loop would be
/// spinning for nothing.
pub const MIN_TICK_MS: u64 = 8;
/// The most a hand-edited tick is reduced to, in milliseconds: past a second the
/// number says nothing that turning previews off does not say better.
pub const MAX_TICK_MS: u64 = 1000;
/// How long a hover's load may run before the waiting spinner is put up for it.
///
/// The window is hidden while a load runs, so one that finishes inside this has gone
/// straight from nothing to the preview: the delay is what keeps a decode that takes a
/// few milliseconds from flashing a spinner on the way past. A quarter of a second is
/// where a wait starts to be worth showing, and `0` is a spinner that goes up with the
/// load. Every kind of preview is answered by it — a decode, a page Office is rendering,
/// a browser that has to start — because a wait is a wait to the pointer that is on the
/// file.
pub const DEFAULT_SPINNER_DELAY_MS: u64 = 250;
/// The ceiling a hand-edited delay is reduced to: past ten seconds a load this app
/// shows has finished or failed, and a larger number switches the spinner off by
/// arithmetic rather than by a setting of its own.
pub const MAX_SPINNER_DELAY_MS: u64 = 10_000;
/// Memory the decoded-image cache may hold. A preview is decoded at the size the
/// layout asked for, so this is a ceiling on retained pixels rather than on
/// files: how many images fit depends entirely on how large they are shown.
///
/// A small budget is kept by default: a hit skips a full-resolution decode and its
/// resample, which is the most expensive thing a hover does, and what is held is
/// the preview-sized frame rather than the pixels it was decoded from.
pub const DEFAULT_IMAGE_CACHE_MB: u32 = 32;
pub const MAX_IMAGE_CACHE_MB: u32 = 2048;
/// Memory the rendered Office pages may hold. A render is what the preview shows
/// for a document, and producing one costs an Office start and an export, so what
/// has been rendered is kept by default: a page that has been drawn comes back
/// without another render.
pub const DEFAULT_OFFICE_CACHE_MB: u32 = 64;
pub const MAX_OFFICE_CACHE_MB: u32 = 2048;
/// Memory the pages a PDF preview was drawn as may hold. A page is stored as the
/// pixels it was rendered into, so the size the layout asked for is part of what
/// is kept rather than only the file it came from.
///
/// A small budget is kept by default — a hit skips the document load, the raster
/// and the renderer's own encode — and a page's own pixels are what it costs, so a
/// budget holds fewer pages than it would frames of anything smaller.
pub const DEFAULT_PDF_CACHE_MB: u32 = 32;
pub const MAX_PDF_CACHE_MB: u32 = 2048;
/// How much of what the render engine drew may be kept, in megabytes. A document the engine
/// drew is a PDF written under the app's own folder, so what one costs once it has been
/// converted is the size of that file, and what is kept is the working set of the folders a
/// user previews: the oldest are given up first, and a document whose page has been given up
/// is converted again the next time it is hovered. Thirty-two megabytes is eight or ten
/// converted drawings.
pub const DEFAULT_LIBRE_CACHE_MB: u32 = 32;
pub const MAX_LIBRE_CACHE_MB: u32 = 2048;
/// Memory the frames a text preview was painted as may hold. A frame is stored as
/// the pixels it was painted into, so the box it was painted in and the scroll
/// position it starts at are part of what is kept rather than only the file it came
/// from.
///
/// Nothing is held by default: a frame is painted for the hover that asks for it,
/// and keeping the frames of the files that are hovered most is a choice the user
/// makes.
pub const DEFAULT_TEXT_CACHE_MB: u32 = 0;
pub const MAX_TEXT_CACHE_MB: u32 = 2048;
/// What one hover may decode or read for, in gigabytes: the ceiling every reader is
/// handed before it allocates — a picture's decode, a document's bytes, the page
/// Office exported, a theme a preview is painted with.
///
/// A gigabyte by default. A seventy-megapixel illustration decodes in about a quarter
/// of that, so a file someone meant to hover never reaches it, while a file that asks
/// for more memory than a preview could justify — a decompression bomb is the whole of
/// that idea — is answered with no preview instead of with an allocation the allocator
/// could abort on. For that reason there is no value that means *no limit*: a ceiling
/// that is not a positive number falls back to the default rather than opening the app
/// to what this exists to refuse.
pub const DEFAULT_DECODE_BUDGET_GB: f32 = 1.0;
pub const MIN_DECODE_BUDGET_GB: f32 = 0.25;
pub const MAX_DECODE_BUDGET_GB: f32 = 64.0;
/// The curve a picture whose samples are light is brought into eight bits with: an EXR, a
/// Radiance HDR and a float texture hold light rather than levels, so what they need is
/// the transfer function a PNG has already been through and a curve that brings a range
/// wider than the display's into it (see `tone_map`).
///
/// Reinhard by default. It is the curve that leaves a value the display can already show
/// very nearly where it was and brings everything above that down without clipping it, so
/// a render whose lamp is a hundred times white is a picture rather than a white patch;
/// `off` is the bare clamp the app drew before this existed, and `aces` is the filmic
/// answer: darker in the shadows and more saturated.
pub const DEFAULT_HDR_TONE_MAP: Curve = Curve::Reinhard;
/// How many stops those pictures are shifted by before the curve, which is what a file far
/// darker or brighter than a display can be is brought into range with. `0` is the picture
/// as the file holds it.
pub const DEFAULT_HDR_EXPOSURE: f32 = 0.0;
pub const MIN_HDR_EXPOSURE: f32 = -10.0;
pub const MAX_HDR_EXPOSURE: f32 = 10.0;
/// Which engine draws an Office document's page by default: the application that owns the
/// format, with the render engine beside it as the fallback for a family this machine has
/// no application for.
pub const DEFAULT_OFFICE_ENGINE: OfficeEngine = OfficeEngine::MicrosoftOffice;
/// How long the Office engine a family started is kept after that family's last
/// page. Producing a page costs an Office start, and an engine still warm is what
/// makes the next document of that family cheap, so one is kept for a while by
/// default.
pub const DEFAULT_OFFICE_ENGINE_IDLE_SECS: u64 = 600;
/// How long the WebView2 engine is kept warm by default: ten minutes, the same as the
/// Office engine, because both are a process this app would rather not start twice. The
/// browser is what draws every SVG document this app previews, so what this buys is
/// every hover after the first one in a while.
pub const DEFAULT_WEBVIEW_IDLE_SECS: u64 = 600;
/// How long the LibreOffice engine is kept after the last page it drew, by default: ten
/// minutes, the same as the other two engines this app keeps, because the start is the whole
/// of what keeping it saves. What it costs is the application itself — a few hundred
/// megabytes while it is held — which is what the setting is for.
pub const DEFAULT_LIBREOFFICE_IDLE_SECS: u64 = 600;
/// The ceiling a hand-edited number of seconds is reduced to. Past a day there is
/// nothing a number says that `indefinitely` does not say better.
pub const MAX_OFFICE_ENGINE_IDLE_SECS: u64 = 86_400;
/// How long Explorer has to be out of reach before an engine that is not marked
/// `Persistent` is let go, by default: one minute. What an engine costs is a process left
/// running, and a minute is long enough that stepping into another application and
/// straight back out of it is not a start paid for the visit.
pub const DEFAULT_AFK_TIMER_SECS: u64 = 60;
/// The ceiling a hand-edited `afk_timer_seconds` is reduced to, the same day the engine
/// idle times are bounded by — a number past it says nothing the top of the menu does not
/// say better. `0` is left as it is: it is the way to ask for an engine to go the moment
/// Explorer goes out of reach.
pub const MAX_AFK_TIMER_SECS: u64 = 86_400;

pub fn sanitize_webp_playback_fps(value: u32) -> u32 {
    match value {
        0 => DEFAULT_WEBP_PLAYBACK_FPS,
        1..=MAX_WEBP_PLAYBACK_FPS => value,
        _ => MAX_WEBP_PLAYBACK_FPS,
    }
}

/// The decoded-image cache size in megabytes. `0` is a cache that is switched
/// off rather than a size that has to be corrected — it is the way to ask for
/// every preview to be decoded again — and anything past the ceiling is clamped
/// to it.
pub fn sanitize_image_cache_mb(value: u32) -> u32 {
    value.min(MAX_IMAGE_CACHE_MB)
}

/// The tick the loop runs at, in milliseconds: what a hand-edited file is read
/// through, so that a number the system cannot honour — or one that would leave the
/// app looking asleep — is brought to the near end of the two bounds rather than
/// taken as it is (see `DEFAULT_TICK_MS`).
pub fn sanitize_tick_ms(value: u64) -> u64 {
    value.clamp(MIN_TICK_MS, MAX_TICK_MS)
}

/// The rendered-page cache size in megabytes.
///
/// `0` is a cache that holds nothing rather than a render tier that is switched
/// off: a page is rendered for the hover that asks for it and dropped when that
/// hover is over, so the only question a size answers is how much of it is kept
/// between hovers.
pub fn sanitize_office_cache_mb(value: u32) -> u32 {
    value.min(MAX_OFFICE_CACHE_MB)
}

/// The rendered-PDF-page cache size in megabytes.
///
/// `0` is a cache that holds nothing rather than a preview tier that is switched
/// off: a page is rendered for the hover that asks for it either way, so the only
/// question a size answers is how much of what was drawn is kept between hovers.
pub fn sanitize_pdf_cache_mb(value: u32) -> u32 {
    value.min(MAX_PDF_CACHE_MB)
}

/// How much of what the render engine drew is kept, in megabytes.
///
/// `0` is a cache that holds nothing rather than a kind of preview that is switched off: a
/// document is drawn for the hover that asks for it either way, and the only question a size
/// answers is how much of what was drawn is kept between hovers.
pub fn sanitize_libre_cache_mb(value: u32) -> u32 {
    value.min(MAX_LIBRE_CACHE_MB)
}

/// The painted-text-frame cache size in megabytes.
///
/// `0` is a cache that holds nothing rather than a text preview that is switched
/// off: a frame is painted for the hover that asks for it either way, so the only
/// question a size answers is how much of what was painted is kept between hovers.
pub fn sanitize_text_cache_mb(value: u32) -> u32 {
    value.min(MAX_TEXT_CACHE_MB)
}

/// The text preview font scale, where `0` and nonsense land back on the default.
/// Anything from 1% to 1000% is honored: the tray offers a handful of steps, but
/// the value is a percentage either way, so a hand-edited one is not rounded to
/// the nearest menu entry.
pub fn sanitize_text_font_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_TEXT_FONT_SCALE_PERCENT
    } else {
        value.clamp(MIN_TEXT_FONT_SCALE_PERCENT, MAX_TEXT_FONT_SCALE_PERCENT)
    }
}

/// Which face of a collection a specimen is drawn from.
///
/// Faces are numbered from `1`, so a `0` — a key deleted by hand, or one written by an
/// older version of the app that had nothing to write — is the first face rather than a
/// face of its own, and a number past the last face the menu offers is that last face.
pub fn sanitize_ttc_face(value: u32) -> u32 {
    value.clamp(DEFAULT_TTC_FACE, MAX_TTC_FACE)
}

/// How far past the far edge of a text preview the pointer region reaches, in
/// logical pixels at the display's DPI.
///
/// Zero is a distance like any other — the region then ends at the preview, which
/// is what it did before the grace existed — so only a value that is not a number
/// at all falls back to the default.
pub fn sanitize_text_scroll_far_edge_grace_pixels(value: f32) -> f32 {
    if value.is_finite() {
        value.clamp(0.0, MAX_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS)
    } else {
        DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS
    }
}

/// The delay before a hover's load puts the spinner up, in milliseconds.
///
/// `0` is a delay like any other — the spinner then goes up with the load — so only a
/// number past the ceiling is corrected.
pub fn sanitize_spinner_delay_ms(value: u64) -> u64 {
    value.min(MAX_SPINNER_DELAY_MS)
}

/// The decode budget in gigabytes.
///
/// A value that is not a number at all, or one that is not positive, falls back to
/// the default rather than being read as "no limit" — the failure this guards against
/// takes the app with it, so it is not a setting that can be switched off. Anything
/// past the ceiling is clamped to it.
pub fn sanitize_decode_budget_gb(value: f32) -> f32 {
    if !value.is_finite() || value <= 0.0 {
        DEFAULT_DECODE_BUDGET_GB
    } else {
        value.clamp(MIN_DECODE_BUDGET_GB, MAX_DECODE_BUDGET_GB)
    }
}

/// The stops a picture whose samples are light is shifted by. A value that is not a number
/// falls back to the default, and one far outside what any picture would be shifted by is
/// clamped to the range: the shift is taken as a power of two of the number, and what an
/// exponent of a thousand is is an infinity.
pub fn sanitize_hdr_exposure(value: f32) -> f32 {
    if !value.is_finite() {
        return DEFAULT_HDR_EXPOSURE;
    }

    value.clamp(MIN_HDR_EXPOSURE, MAX_HDR_EXPOSURE)
}

/// What a reader may ask the allocator for, in bytes.
///
/// Read from the configuration each time it is asked for rather than captured, like
/// the cache sizes, so an edit applies to the next hover without a restart.
pub fn decode_budget_bytes() -> u64 {
    let gigabytes = crate::CONFIG
        .lock()
        .map(|config| sanitize_decode_budget_gb(config.decode_budget_gb))
        .unwrap_or(DEFAULT_DECODE_BUDGET_GB);

    (f64::from(gigabytes) * 1024.0 * 1024.0 * 1024.0) as u64
}

/// The budget above as the limits an `image` decoder is read under: no limit on a
/// picture's own dimensions, the budget on the memory the decode may ask for.
///
/// Every decoder in the app is handed these — the readers a hover opens a file with,
/// and the ones that draw what a render produced — so no decode reaches an allocator
/// without having asked.
pub fn image_decode_limits() -> image::Limits {
    let mut limits = image::Limits::no_limits();
    limits.max_alloc = Some(decode_budget_bytes());
    limits
}

/// The bytes a frame of this shape is, when they fit what one hover may decode for —
/// the question for the readers that allocate a canvas of their own rather than going
/// through a decoder's limits: the GIF's and the animated WebP's, and the codec
/// Windows is asked for a picture's pixels.
///
/// A picture's shape is itself unbounded: a seven-thousand by ten-thousand
/// illustration is an ordinary thing to hover and is decoded at the size it is, so
/// what a file may ask for is the budget rather than a cap on its dimensions (see
/// `decode_budget_bytes`). The product is taken in `u64`, so a shape whose frame
/// overflows a `u32` cannot wrap into a size that would pass.
pub fn frame_bytes_within_budget(width: u32, height: u32, bytes_per_pixel: u64) -> Option<usize> {
    let bytes = width as u64 * height as u64 * bytes_per_pixel;
    (bytes <= decode_budget_bytes()).then_some(bytes as usize)
}

/// A file's bytes, read whole, when the file is small enough to be read under the
/// budget. `None` is a file that is larger than a hover may ask for, which is answered
/// with no preview.
///
/// The size is asked of the directory entry first, so a file past the budget costs no
/// read at all, and the read itself is bounded as well, because a file can grow between
/// the two questions.
pub fn read_within_budget(path: &Path) -> Option<Vec<u8>> {
    let budget = decode_budget_bytes();

    let size = fs::metadata(path).map_or(u64::MAX, |meta| meta.len());
    if size > budget {
        return None;
    }

    let mut bytes = Vec::new();
    File::open(path)
        .ok()?
        .take(budget + 1)
        .read_to_end(&mut bytes)
        .ok()?;

    if bytes.len() as u64 > budget {
        return None;
    }

    Some(bytes)
}

fn parse_text_font_scale(value: &str) -> Option<u32> {
    let normalized = value.trim().to_ascii_lowercase();
    let normalized = normalized.trim_end_matches('%').trim();
    normalized
        .parse::<u32>()
        .ok()
        .map(sanitize_text_font_scale_percent)
}

fn sanitize_preview_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_PREVIEW_SCALE_PERCENT
    } else {
        value.clamp(MIN_PREVIEW_SCALE_PERCENT, MAX_PREVIEW_SCALE_PERCENT)
    }
}

/// How the preview is sized relative to the media's native pixel dimensions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviewScale {
    /// Scale as large as the available display area allows.
    FitToScreen,
    /// Scale as large as the available display area allows, then reduced to this
    /// share of it.
    ///
    /// This is a scale the app derives rather than one the configuration holds: a
    /// source that is drawn at any size it is asked for — a PDF page, an SVG
    /// document, a page Office rendered — is laid out at fit-to-screen, because the
    /// display's room is free quality there, and a configured percentage below `100`
    /// is answered by reducing that size rather than ignored.
    FitToScreenReduced(u32),
    /// Scale by a percentage of the media's native size.
    Percent(u32),
}

impl PreviewScale {
    pub fn as_str(self) -> String {
        match self {
            Self::FitToScreen => "fit".to_string(),
            // Never written: the reduced fit is derived from the configured scale,
            // and what a file would be read back as is the plain fit it is a share
            // of.
            Self::FitToScreenReduced(_) => "fit".to_string(),
            Self::Percent(percent) => sanitize_preview_scale_percent(percent).to_string(),
        }
    }

    /// Requested scale relative to the native size, or `None` when the preview
    /// should use the largest scale the display area allows.
    pub fn target_scale(self) -> Option<f32> {
        match self {
            Self::FitToScreen | Self::FitToScreenReduced(_) => None,
            Self::Percent(percent) => Some(sanitize_preview_scale_percent(percent) as f32 / 100.0),
        }
    }

    /// The share of the fitted size this scale asks for: `1.0` where the preview
    /// takes the room it is given, less where it is a reduction of that room.
    pub fn fit_share(self) -> f32 {
        match self {
            Self::FitToScreenReduced(percent) => {
                sanitize_preview_scale_percent(percent) as f32 / 100.0
            }
            _ => 1.0,
        }
    }

    /// The scale a `config.ini` value names, or `None` for one that is neither a
    /// percentage nor a word for the fit.
    pub(crate) fn from_str(value: &str) -> Option<Self> {
        let normalized = value.trim().to_ascii_lowercase();
        let normalized = normalized.trim_end_matches('%').trim();

        match normalized {
            "fit" | "fit to screen" | "fit-to-screen" | "fit_to_screen" | "fittoscreen" => {
                Some(Self::FitToScreen)
            }
            _ => normalized
                .parse::<u32>()
                .ok()
                .map(|percent| Self::Percent(sanitize_preview_scale_percent(percent))),
        }
    }
}

/// What a PDF page is drawn at unless the configuration says otherwise: the whole of
/// the room the display has for it, which is the answer that asks for nothing in
/// particular — the page at its own size where the display can hold it, reduced only
/// where it cannot.
pub const DEFAULT_PDF_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// The same for the page an Office document is drawn as.
pub const DEFAULT_OFFICE_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// What a picture, a video and an animated picture are drawn at unless the configuration
/// says otherwise: the size each file asks for, at the share its own setting names.
pub const DEFAULT_PREVIEW_SCALE: PreviewScale =
    PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT);
pub const DEFAULT_VIDEO_SCALE: PreviewScale = PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT);
pub const DEFAULT_ANIMATED_SCALE: PreviewScale =
    PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT);
/// The same for a vector drawing, at the whole of the room: a drawing is drawn at whatever
/// size it is asked for — an SVG document by the browser that rasterizes nothing until it
/// is told the size, a metafile by the drawing layer playing its records again — so the
/// room the display has is free quality rather than an enlargement, and the size that asks
/// for nothing in particular is all of it. A share below it is a size the user picked, and
/// reduces what the room would have given rather than being ignored.
pub const DEFAULT_VECTOR_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// The same for a design document, at the whole of the room: what a preview of one is made
/// of is the picture the file keeps of the whole document, so the question the setting
/// answers is how much of the display to give it, and the answer that asks for nothing in
/// particular is the room the display has.
pub const DEFAULT_DESIGN_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// The same for a document the render engine draws, at the whole of the room: what a preview
/// of one is is a page, drawn by the engine at whatever size the display has, so the question
/// the setting answers is how much of the display to give it. See `libreoffice_render`.
pub const DEFAULT_LIBRE_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// The same for a font specimen, at the share rather than the whole of the room: a specimen
/// is a page of text rather than a document to be studied, and half the display holds the
/// pangram at a size that can be read at a glance.
pub const DEFAULT_FONT_SCALE: PreviewScale = PreviewScale::Percent(DEFAULT_FONT_SCALE_PERCENT);

/// Which engine an Office document's page is asked of, as the tray's
/// `Engine → Select Engine → Office` lists it.
///
/// There are two engines to ask, and the choice between them is one setting: the application
/// that owns the format, which is what a page of one has always been drawn by, and the render
/// engine beside it — an installed LibreOffice — which draws every Office document whether the
/// application is here or not. What the second is for is a machine where the application
/// draws a page badly or not at all, or one whose user would rather every document of the
/// kind came out of the engine they know (see `office_formats::page_engine`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfficeEngine {
    /// Word, Excel or PowerPoint, and the render engine as the fallback for a family this
    /// machine has no application for.
    MicrosoftOffice,
    /// The render engine for every Office document, where one is installed to be asked.
    LibreOffice,
}

impl OfficeEngine {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::MicrosoftOffice => "microsoft_office",
            Self::LibreOffice => "libreoffice",
        }
    }

    /// The engine a `config.ini` value names, or `None` for one that is neither: a value the
    /// app cannot read leaves the setting where it is, and the file is written back with the
    /// value the app is actually using (see `differs`).
    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "microsoft_office" | "microsoft office" | "msoffice" | "office" => {
                Some(Self::MicrosoftOffice)
            }
            "libreoffice" | "libre" | "soffice" => Some(Self::LibreOffice),
            _ => None,
        }
    }
}

/// How long an engine that is kept warm between documents is kept.
///
/// Two settings are one shape: the Office engine, which is the application this app
/// started and would rather not start again, and the WebView2 engine that draws SVG
/// documents, which is a browser process. Both are kept for a while after the
/// last document and then let go, both may be kept for the life of the app, and both
/// take the same words in `config.ini`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EngineIdle {
    /// Kept for this many seconds after the family's last page, `0` included: an
    /// engine let go as soon as it has drawn one.
    Seconds(u64),
    /// Kept for as long as the app runs.
    ///
    /// An engine that is never let go is never left holding what a render did to
    /// it either. The automation settings a render needs are taken and put back
    /// around each render rather than held for the engine's life (see
    /// `office_render`), so what an indefinite setting keeps is a process that is
    /// doing nothing this app's business — which is what makes it safe to keep an
    /// instance that is the user's own Word or Excel.
    Indefinite,
}

impl EngineIdle {
    pub fn as_str(self) -> String {
        match self {
            Self::Seconds(seconds) => sanitize_engine_idle_secs(seconds).to_string(),
            Self::Indefinite => "indefinitely".to_string(),
        }
    }

    /// The idle time a `config.ini` value names, or `None` for one that is
    /// neither a number of seconds nor a word for never letting go.
    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "indefinite" | "indefinitely" | "forever" | "always" => Some(Self::Indefinite),
            other => other
                .parse::<u64>()
                .ok()
                .map(|seconds| Self::Seconds(sanitize_engine_idle_secs(seconds))),
        }
    }

    /// Whether an engine that has gone `idle_for` without drawing a page has been
    /// idle long enough to be let go. An engine that is kept for the life of the
    /// app never has.
    pub fn has_expired(self, idle_for: Duration) -> bool {
        match self {
            Self::Seconds(seconds) => idle_for >= Duration::from_secs(seconds),
            Self::Indefinite => false,
        }
    }

    /// The same, as the wait it is: `None` for an engine that is never let go.
    pub fn as_duration(self) -> Option<Duration> {
        match self {
            Self::Seconds(seconds) => Some(Duration::from_secs(seconds)),
            Self::Indefinite => None,
        }
    }
}

fn sanitize_engine_idle_secs(seconds: u64) -> u64 {
    seconds.min(MAX_OFFICE_ENGINE_IDLE_SECS)
}

/// The away time a hand-edited `afk_timer_seconds` is read through, so that a number past
/// what the menu offers is brought to the ceiling rather than taken as it is. `0` is left
/// alone: the menu does not offer it either, and an engine let go the moment Explorer goes
/// out of reach is a setting someone may mean.
fn sanitize_afk_timer_secs(seconds: u64) -> u64 {
    seconds.min(MAX_AFK_TIMER_SECS)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransparentBackground {
    Transparent,
    Black,
    White,
    Checkerboard,
}

impl TransparentBackground {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Transparent => "transparent",
            Self::Black => "black",
            Self::White => "white",
            Self::Checkerboard => "checkerboard",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "transparent" => Some(Self::Transparent),
            "black" => Some(Self::Black),
            "white" => Some(Self::White),
            "checkerboard" => Some(Self::Checkerboard),
            _ => None,
        }
    }
}

/// What a picture is drawn over unless the configuration says otherwise — and with it
/// every other preview this app draws rather than one a page is handed to it: a PDF page,
/// a painted text frame, a page Office rendered, a design document's own picture.
///
/// A checkerboard, which is how transparency is shown wherever it is shown at all: what a
/// picture's alpha channel leaves unpainted is the pixels under the picture, and the
/// squares are the answer that cannot be taken for part of the picture the way black can.
pub const DEFAULT_IMAGE_BACKGROUND: TransparentBackground = TransparentBackground::Checkerboard;

/// The same for a vector drawing — an SVG document the browser draws, or a metafile the
/// drawing layer replays.
///
/// The squares again, and for the reason above with a document's own twist: a drawing
/// says nothing about the sheet under it, and the page a browser is handed can only be
/// given a colour (see `webview_preview::frame_page`), so a backdrop that reads as
/// *nothing here* is the one that keeps an unpainted region from looking painted black.
pub const DEFAULT_VECTOR_BACKGROUND: TransparentBackground = TransparentBackground::Checkerboard;

/// What a font specimen is drawn over: white, which is the page a specimen is written on
/// rather than a backdrop behind one.
///
/// A specimen is this app's own page with a font's glyphs on it, and the ink is picked for
/// the page it is written on — light on black, dark on white and on the squares (see
/// `webview_preview::font_page`) — so what is behind the text is a page's colour rather
/// than a transparency to be shown.
pub const DEFAULT_FONT_BACKGROUND: TransparentBackground = TransparentBackground::White;

/// What a `.dds` texture is drawn over: white.
///
/// The two backdrops that show what stands behind a preview are not offered for a texture
/// at all, and this is where the setting starts instead: a texture's alpha channel is as
/// often a mask, a height or a roughness as it is transparency (see `dds_image`), so what
/// is drawn behind one is a page to read the channels against rather than a hole to look
/// through.
pub const DEFAULT_DDS_BACKGROUND: TransparentBackground = TransparentBackground::White;

/// What a design document is drawn over: the squares, for the reason a picture's is —
/// what a document is previewed from is the picture the file keeps of the whole thing, and
/// that picture's transparency is the document's own.
pub const DEFAULT_DESIGN_BACKGROUND: TransparentBackground = TransparentBackground::Checkerboard;

/// The backdrop a texture is drawn over, as the setting keeps it: either of the two a
/// texture is offered, and the default for anything else — a `transparent` or a
/// `checkerboard` a file still holds from when the texture's half of the `Background`
/// submenu listed the same four backdrops as every other half.
pub fn sanitize_dds_background(background: TransparentBackground) -> TransparentBackground {
    match background {
        TransparentBackground::Black | TransparentBackground::White => background,
        TransparentBackground::Transparent | TransparentBackground::Checkerboard => {
            DEFAULT_DDS_BACKGROUND
        }
    }
}

/// The marker a theme from the `theme` folder is written after in `config.ini`.
///
/// It is what keeps a file named after a bundled theme apart from the bundled
/// theme of that name: `light.tmTheme` in the folder is `custom:light`, while
/// `light` on its own is Atom One Light however many files the folder holds.
const CUSTOM_THEME_PREFIX: &str = "custom:";

/// Names already interned for [`TextTheme::custom`], so one file name is one
/// `&'static str` however many times a menu or a configuration read hands it over.
static CUSTOM_THEME_NAMES: Lazy<Mutex<HashSet<&'static str>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextTheme {
    /// Atom One Light.
    Light,
    /// One Dark Pro.
    Dark,
    /// A `.tmTheme` file in the `theme` folder, by the name it is listed under.
    Custom(&'static str),
}

impl TextTheme {
    /// How the theme is written in `config.ini`.
    pub fn as_str(self) -> String {
        match self {
            Self::Light => "light".to_string(),
            Self::Dark => "dark".to_string(),
            Self::Custom(name) => format!("{CUSTOM_THEME_PREFIX}{name}"),
        }
    }

    /// The theme a file of that name in the `theme` folder is, interned so that it
    /// can live in a value the rest of the app copies around.
    ///
    /// The folder is listed and `config.ini` is read back whenever either changes,
    /// and both hand a name over again each time; interning is what keeps the same
    /// theme one value rather than a fresh string per reading.
    pub fn custom(name: &str) -> Self {
        let mut names = match CUSTOM_THEME_NAMES.lock() {
            Ok(names) => names,
            Err(poisoned) => poisoned.into_inner(),
        };
        if let Some(name) = names.get(name).copied() {
            return Self::Custom(name);
        }

        let name: &'static str = Box::leak(name.to_string().into_boxed_str());
        names.insert(name);
        Self::Custom(name)
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "light" | "atom one light" | "atom-one-light" => Some(Self::Light),
            "dark" | "one dark pro" | "one-dark-pro" => Some(Self::Dark),
            _ => None,
        }
    }

    /// The theme a `config.ini` value names, if this machine has one.
    ///
    /// Three spellings, in the order they are read: a name the tray wrote for a
    /// file (`custom:atom-one-light`), the name of a bundled theme — `light`,
    /// `dark`, and the long spellings accepted as ever — and a file in the `theme`
    /// folder written by hand without the marker, its extension optional. A value
    /// that is none of those leaves the theme as it was, which is what puts the
    /// default back when a selection stops making sense.
    fn resolve(value: &str) -> Option<Self> {
        let value = value.trim();

        if let Some(name) = value.strip_prefix(CUSTOM_THEME_PREFIX) {
            let name = name.trim();
            if name.is_empty() {
                return None;
            }

            // A file the folder holds under another spelling is still that file,
            // and the menu lists the spelling the folder uses; a name it does not
            // hold is kept as written, so the choice survives a file that is
            // missing for the moment.
            let name = theme_files::find(name).unwrap_or_else(|| name.to_string());
            return Some(Self::custom(&name));
        }

        if let Some(built_in) = Self::from_str(value) {
            return Some(built_in);
        }

        theme_files::find(value).map(|name| Self::custom(&name))
    }
}

/// How a Markdown file is laid out: the document it describes, or the markup
/// itself with the Markdown syntax highlighted like any other source file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MarkdownMode {
    Rendered,
    Source,
}

impl MarkdownMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Rendered => "rendered",
            Self::Source => "source",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "rendered" | "render" | "document" => Some(Self::Rendered),
            "source" | "raw" | "highlighted" | "highlighted source" => Some(Self::Source),
            _ => None,
        }
    }
}

/// A kind of preview, as the tray's `Preview Types` submenu lists them.
///
/// A gate is not a file list: it says whether previews of that kind may be shown
/// at all, and the lists that decide *which* files of that kind are previewed are
/// left untouched by it, so switching a kind off and back on restores exactly
/// what was configured.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviewType {
    Images,
    Videos,
    Text,
    Pdf,
    Archives,
    Office,
    /// Font files, the second kind the browser draws rather than a decoder: a specimen is
    /// a page of this app's own with the font in it, so a machine without the engine has no
    /// font preview either, and a user who wants none of them has this switch.
    Fonts,
    /// Design documents and projects, previewed from the picture their own format keeps
    /// of the whole document rather than from their layers: a user who wants none of
    /// them has this switch, which is not the switch for pictures even though the
    /// preview is one.
    Design,
    /// Vector drawings — SVG documents, Windows' metafiles, and the preview an
    /// encapsulated PostScript file carries — drawn rather than decoded: an SVG by the
    /// browser engine, the rest by the drawing layer, and none of them by a decoder the
    /// way a picture is, which is why a drawing is sharp at any size the display has.
    ///
    /// SVG is the one kind in this list that is also an entry of the image list — a
    /// `.svg` is a picture's name to every list that reads names — and what draws one is
    /// not a decoder but an engine, so its switch is this one rather than the switch for
    /// pictures; see `svg_preview`.
    Vector,
    /// Documents this app hands to a render engine rather than reading — CorelDRAW above
    /// all, and the word processors, spreadsheets, presentations and drawings whose own
    /// formats no reader here has. What draws one is LibreOffice where it is installed, and
    /// what comes back is a page; see `libre_formats` for what is listed and
    /// `libreoffice_render` for how it is drawn.
    Libre,
    /// Pictures this app hands to an installed ImageMagick rather than decoding — the camera
    /// raw formats above all, which nothing else on a Windows machine opens at all. What
    /// comes back is a PNG, and it is drawn as the picture it is: the picture scale and the
    /// picture backdrop, held in the picture cache. See `magick_formats` for what is listed
    /// and `imagemagick_render` for how one is converted.
    Magick,
    /// Archives this app hands to an installed PeaZip rather than reading — the cabinet files,
    /// isos, disk images, installers and single-stream compressors no reader here has. What
    /// comes back is the archive's own table of contents, read into the shape every other
    /// listing is and drawn as the same page, which is why this is a gate of its own rather
    /// than a second list of names: a user who wants their `.cab` files left alone is not
    /// asking for their zips to be left alone. See `peazip_formats` for what is listed and
    /// `peazip_render` for how one is listed.
    Peazip,
}

impl PreviewType {
    /// Whether this kind of preview may be shown under the current configuration.
    pub fn enabled(self) -> bool {
        crate::CONFIG
            .lock()
            .map(|config| self.enabled_in(&config))
            .unwrap_or(true)
    }

    /// Whether this kind of preview is switched on in `config`.
    pub fn enabled_in(self, config: &AppConfig) -> bool {
        match self {
            Self::Images => config.image_preview_enabled,
            Self::Videos => config.video_preview_enabled,
            Self::Text => config.text_preview_enabled,
            Self::Pdf => config.pdf_preview_enabled,
            Self::Archives => config.archive_preview_enabled,
            Self::Office => config.office_preview_enabled,
            Self::Fonts => config.font_preview_enabled,
            Self::Design => config.design_preview_enabled,
            Self::Vector => config.vector_preview_enabled,
            Self::Libre => config.libre_preview_enabled,
            Self::Magick => config.magick_preview_enabled,
            Self::Peazip => config.peazip_preview_enabled,
        }
    }

    /// Switch this kind of preview on or off.
    pub fn set_enabled_in(self, config: &mut AppConfig, enabled: bool) {
        match self {
            Self::Images => config.image_preview_enabled = enabled,
            Self::Videos => config.video_preview_enabled = enabled,
            Self::Text => config.text_preview_enabled = enabled,
            Self::Pdf => config.pdf_preview_enabled = enabled,
            Self::Archives => config.archive_preview_enabled = enabled,
            Self::Office => config.office_preview_enabled = enabled,
            Self::Fonts => config.font_preview_enabled = enabled,
            Self::Design => config.design_preview_enabled = enabled,
            Self::Vector => config.vector_preview_enabled = enabled,
            Self::Libre => config.libre_preview_enabled = enabled,
            Self::Magick => config.magick_preview_enabled = enabled,
            Self::Peazip => config.peazip_preview_enabled = enabled,
        }
    }
}

/// What the trigger key does while it is held.
///
/// A preview appears when a file is hovered, so the trigger key is normally what
/// stops that: hold it and nothing previews while it is down. The reverse suits
/// a machine where previews are the exception rather than the rule — hold the key
/// and they appear, let go and they stop — and it is the same gesture either way,
/// which is why one key covers both.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TriggerKeyMode {
    Disable,
    Enable,
}

impl TriggerKeyMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Disable => "disable",
            Self::Enable => "enable",
        }
    }

    /// Whether a preview may appear while the trigger key is in this state.
    ///
    /// One question, whichever way round the setting is: in `Disable` the key is
    /// what stops previews, so it has to be up; in `Enable` it is the only thing
    /// that starts them, so it has to be down.
    pub fn allows_previews(self, trigger_key_down: bool) -> bool {
        match self {
            Self::Disable => !trigger_key_down,
            Self::Enable => trigger_key_down,
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "disable" | "disables" | "off" | "hold to disable" => Some(Self::Disable),
            "enable" | "enables" | "on" | "hold to enable" => Some(Self::Enable),
            _ => None,
        }
    }
}

/// How far a preview is placed clear of the item it is about.
///
/// A view draws an item's name, and the views that draw their items as rows draw the
/// columns beside it as well — when the row was modified, its type, its size. What a
/// preview has to clear to stay off the name is therefore a choice: nothing, the name
/// where it is drawn, the whole column the name is drawn in, or everything the item
/// draws.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AvoidMode {
    /// A preview is placed where the position alone puts it.
    Off,
    /// A preview is kept clear of the name and nothing else — as far as the name is
    /// drawn, so the rest of the column it sits in may be covered.
    Filename,
    /// A preview is kept clear of the whole column the name is drawn in, as the view
    /// reports it — the `Name` column of a `Details` row, whatever the name takes of
    /// it — so the columns beside it may be covered.
    FilenameColumn,
    /// A preview is kept clear of everything the item draws: its name, and the
    /// columns beside it.
    Details,
}

impl AvoidMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Off => "off",
            Self::Filename => "filename",
            Self::FilenameColumn => "filename_column",
            Self::Details => "details",
        }
    }

    /// The way a `config.ini` value names, or `None` for one that names no way of
    /// keeping a preview off an item.
    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "off" | "none" | "don't avoid" | "dont avoid" => Some(Self::Off),
            "filename" | "name" | "avoid filename" => Some(Self::Filename),
            "filename_column" | "filename column" | "name column" | "avoid filename column" => {
                Some(Self::FilenameColumn)
            }
            "details" | "all" | "avoid details" => Some(Self::Details),
            _ => None,
        }
    }
}

/// How far a preview is kept clear of the item it is about unless the configuration says
/// otherwise: the file's name where it is drawn, which is the one thing a row the pointer
/// is on says about itself.
pub const DEFAULT_AVOID_MODE: AvoidMode = AvoidMode::Filename;

/// Where a preview lands unless the configuration says otherwise: `Best Position` is the
/// app's own answer to where a preview should go — the room the display has rather than
/// wherever the pointer happens to be — and `Follow Cursor` is the other answer.
pub const DEFAULT_FOLLOW_CURSOR: bool = false;

#[derive(Debug, Clone)]
pub struct AppConfig {
    pub is_first_run: bool,
    pub run_at_startup: bool,
    pub hover_delay_ms: u64,
    pub preview_enabled: bool,
    /// The modifier the trigger key watches, by name.
    pub trigger_key: String,
    /// What holding it does: stop previews, or allow them.
    pub trigger_key_mode: TriggerKeyMode,
    /// Whether the trigger key is watched at all. Off means previews behave as if
    /// no key were held, whatever the mode says.
    pub trigger_key_enabled: bool,
    pub confirm_file_type: bool,
    pub follow_cursor: bool,
    /// How far a preview is placed clear of the item it is about, so the file the
    /// pointer is on or the keyboard is focused on stays readable while its preview is
    /// up. `Filename` by default, and it is the name that makes it so: the name of the
    /// file being previewed is part of what the preview is about, and a preview that
    /// covers it hides the one thing the pointer's item says, while the columns a row
    /// writes beside the name are not the file's own. What is kept off is the name
    /// where it is drawn rather than the column it sits in, so a short name leaves the
    /// rest of its column to be covered; `FilenameColumn` keeps previews off the whole
    /// of that column, `Details` off the columns beside it as well, and `Off` puts them
    /// back where the position modes alone would have them.
    pub avoid_mode: AvoidMode,
    /// How long the same file waits before a preview of it is put up again, in
    /// milliseconds: what keeps a pointer that crosses a file it has just left from
    /// putting the preview back up on its way past (see
    /// `DEFAULT_SAME_FILE_REHOVER_DELAY_MS`).
    pub same_file_rehover_delay_ms: u64,
    /// How long the pointer must be still before a preview is put up for what it is on,
    /// in milliseconds: what keeps a pointer crossing a list from answering every file it
    /// passes on its way (see `DEFAULT_SETTLING_DELAY_MS`).
    pub settling_delay_ms: u64,
    /// How often the app looks at the pointer's world while Explorer has focus, in
    /// milliseconds (see `DEFAULT_TICK_MS`).
    ///
    /// The one timing that is about the app rather than about a hover: it is the loop's
    /// own tick, so it sets how soon a move is answered, how soon the file under a
    /// parked pointer is read again, and — with them — what the app costs while a folder
    /// window is in front. Every look is a crossing into Explorer, which is why a tick
    /// of nothing is not offered and a very small one is clamped rather than honoured.
    /// What does not move with it: the keyboard's own probe, which keeps its rate (see
    /// `explorer_hook`), and the folder, display and state checks, which are
    /// amortizations of expensive enumerations with clocks of their own.
    pub tick_ms: u64,
    /// How long a hover's load may run before the waiting spinner is put up for it,
    /// in milliseconds. `0` puts it up with the load.
    ///
    /// The one spinner timing there is: every kind of preview is answered by it — a
    /// decode, a page Office is rendering, a browser that has to start — so what the
    /// number says is the one thing it means, how long a wait is given before it is
    /// shown as one.
    pub spinner_delay_ms: u64,
    pub webp_playback_fps: u32,
    /// Memory the decoded-image cache may hold, in megabytes. `0` switches the
    /// cache off, so every preview is decoded again.
    pub image_cache_mb: u32,
    /// The backdrop a picture is drawn over — and every other preview that is not a
    /// document: a PDF page, a painted text frame, a page Office rendered.
    pub image_background: TransparentBackground,
    /// The backdrop a font specimen is drawn over.
    ///
    /// A setting of its own for the reason the one above is: a specimen is a page with text
    /// on it rather than a picture that carries transparency of its own, and what stands
    /// behind the glyphs is the page's business. The transparent one is the case with a rule
    /// of its own — there is no colour of text that reads over whatever the desktop happens
    /// to be — so the glyphs there are drawn light with a soft dark shadow behind them (see
    /// `webview_preview::font_page`).
    pub font_background: TransparentBackground,
    /// The backdrop a `.dds` texture is drawn over.
    ///
    /// A setting of its own for a reason of the format's rather than of the picture's: a
    /// texture's alpha channel is very often not alpha at all — a mask, a height, the
    /// roughness of a material, or a channel a tool never touched and left at zero — so a
    /// texture is the one kind of picture whose transparency says least about what is
    /// behind it, and the one kind a reader may want composited differently from a
    /// photograph (see `dds_image`). For the same reason the two backdrops that show what
    /// stands behind a preview are not offered for a texture at all — the menu lists the
    /// two that are a page, black and white — and a file that names one of the other two
    /// is read as the one this setting starts at (see `sanitize_dds_background`).
    pub dds_background: TransparentBackground,
    /// The backdrop a design document is drawn over.
    ///
    /// A setting of its own for the reason a texture's is: what a document is previewed
    /// from is the picture the file keeps of the whole thing — a merged image, or the
    /// flattened document a project container holds — and that picture is as often one a
    /// designer saved with its transparency as it is one to be looked at against a page,
    /// so what stands behind it is worth being a setting rather than a guess.
    pub design_background: TransparentBackground,
    /// The backdrop a vector drawing is drawn over.
    ///
    /// A setting of its own because a drawing is made for a page and says nothing about it
    /// — the records of a metafile are what was drawn and not the sheet under it, and an
    /// SVG document's own opacity is not a photograph's — so what stands behind the marks
    /// is this app's answer rather than the file's. For the document half of the kind the
    /// page is the engine's and this is the colour it is given; the checkerboard among the
    /// four is the page's own there, since a browser can be handed a colour and nothing
    /// else (see `webview_preview::frame_page`).
    pub vector_background: TransparentBackground,
    pub video_volume: u32,
    /// How large a picture is drawn, as a share of its own size: `100%` is the size the
    /// file asks for, `50%` half of it, and `fit` the largest size the room the layout
    /// gives it allows.
    pub preview_scale: PreviewScale,
    /// How large a video is drawn, as a share of its own size — the same question, and the
    /// same answers, as the picture scale above it.
    ///
    /// A video is measured by a probe and drawn from its first frame, and what the layout
    /// places is that frame, so a share of it is a share of the size the file asks for the
    /// same way a picture's is. It is a setting of its own because the two are hovered for
    /// different reasons: a picture wants the detail it holds, while a video at a size
    /// large enough to read a frame is a window over the file rather than a photograph on
    /// a desk.
    pub video_scale: PreviewScale,
    /// How large an animated picture — a GIF, an animated WebP or an APNG — is drawn, as a
    /// share of its own size: the same question, and the same answers, as the picture scale
    /// above it.
    ///
    /// An animation is decoded into frames, and a frame is a bitmap like a picture's, so a
    /// share of it is a share of the size the file asks for the same way a picture's is —
    /// and the share is read for one whether it moves or not: a `.gif` holding a single
    /// frame is a still picture and keeps `preview_scale`, which is what this setting is
    /// asked apart from. It is a setting of its own because the two are hovered for
    /// different reasons: a picture is studied at the size it was written, while an
    /// animation at `50%` is half the pixels to decode and draw for every frame of it.
    pub animated_scale: PreviewScale,
    /// How large a PDF page is drawn, as a share of the room the display has for it.
    ///
    /// A page is a vector, so what it is asked for is a size rather than a resample:
    /// the room the display has is free quality there, and a share of that room is what
    /// the setting names — `50%` is half the display, not half of the page. `Fit to
    /// Screen` is the whole of it, which is where the setting starts, and `100%` or
    /// more reads as that fit.
    ///
    /// It is a setting of its own rather than the SVG scale beside it because the two
    /// documents are hovered for different reasons: a page is read at a glance and a
    /// drawing is looked at, so the size one wants is rarely the size the other wants.
    pub pdf_scale: PreviewScale,
    /// How large the page an Office document is drawn as is shown, as a share of the
    /// room the display has — the same question, and the same answers, as the PDF
    /// scale beside it.
    ///
    /// The one source that is not a page is the bitmap a workbook is answered with
    /// where no page can be exported: it is only as good as the pixels it holds, so it
    /// follows the share as a share of its own size and is never enlarged.
    pub office_scale: PreviewScale,
    /// How large a font specimen is drawn, as a share of the room the display has — the same
    /// question, and the same answers, as the document scales above.
    ///
    /// What the share is applied to is the nominal specimen box `font_preview` measures a
    /// font at, because a font has no size of its own to be a percentage of: a file holds
    /// outlines, and `50%` is half the room the display would give the specimen at its
    /// largest.
    pub font_scale: PreviewScale,
    /// How large a design document is drawn, as a share of the room the display has for
    /// it — the same question, and the same answers, as the document scales above.
    ///
    /// It is a setting of its own rather than the picture scale beside them because
    /// what the share is applied to is a whole document rather than a photograph: a
    /// layered document is previewed from the picture its format keeps of the whole
    /// thing, at whatever size that picture is, so the size one wants is a share of the
    /// screen the way a page's is rather than a share of the file's own size.
    pub design_scale: PreviewScale,
    /// How large a document the render engine draws is shown, as a share of the room the
    /// display has. See `libreoffice_render`.
    pub libre_scale: PreviewScale,
    /// How large a vector drawing is drawn, as a share of the room the display has for it —
    /// the same question, and the same answers, as the document scales above.
    ///
    /// A drawing is not a bitmap: an SVG document is drawn by the browser engine at
    /// whatever size it is told, and a metafile's records are played again at whatever size
    /// the preview is, so the room the display has is free quality rather than an
    /// enlargement — `50%` is half the display, not half of the file. It is a setting of
    /// its own because a drawing is asked for a share of the screen the way a page is,
    /// rather than for a share of a size the file asks for.
    pub vector_scale: PreviewScale,
    /// Which face of a collection a specimen is drawn from, as the menu numbers faces: `1`
    /// (the default) is the first face the file holds.
    ///
    /// A `.ttc` is several fonts in one file — a family's regular and bold, a typeface's
    /// several languages — which share their outlines and are found by offsets into them,
    /// and a page can be pointed at none of them but the first. So which face a preview is
    /// of is a choice rather than a property of the file, and this is that choice: the face
    /// is the one written out for the engine to draw, and a collection with fewer faces
    /// than the setting names is drawn from the last one it has. The specimen's heading
    /// says which face came out — `(2 of 4)` — so what was asked for and what was drawn
    /// cannot be mistaken for each other.
    pub ttc_face: u32,
    pub theme: TextTheme,
    pub markdown_mode: MarkdownMode,
    /// Whether image previews may be shown at all.
    pub image_preview_enabled: bool,
    /// Whether video previews may be shown at all.
    pub video_preview_enabled: bool,
    /// Whether text files are previewed at all, ahead of the extension list.
    pub text_preview_enabled: bool,
    /// Whether PDF previews may be shown at all.
    pub pdf_preview_enabled: bool,
    /// Whether archive contents are listed at all, ahead of the extension list.
    pub archive_preview_enabled: bool,
    /// Whether Office documents are previewed at all, ahead of the extension list.
    pub office_preview_enabled: bool,
    /// Whether font files are previewed at all, ahead of the font list their names are
    /// entries of.
    pub font_preview_enabled: bool,
    /// Whether design documents and projects are previewed at all, ahead of the design
    /// list their names are entries of.
    pub design_preview_enabled: bool,
    /// Whether a document the render engine draws may be previewed at all. It is the gate
    /// for `[libre]`, and the switch a user who wants their CorelDRAW files left alone
    /// reaches for.
    pub libre_preview_enabled: bool,
    /// Whether a picture the ImageMagick engine develops may be previewed at all. It is the
    /// gate for `[magick]`, and the switch a user who wants their camera raw files left
    /// alone reaches for; see `magick_formats`.
    pub magick_preview_enabled: bool,
    /// Whether an archive the PeaZip engine lists may be previewed at all. It is the gate for
    /// `[peazip]`, and the switch a user who wants their cabinet files, isos and disk images
    /// left alone reaches for — a switch of its own rather than the archive gate beside it,
    /// because these are the names nothing on this machine but an engine opens; see
    /// `peazip_formats`.
    pub peazip_preview_enabled: bool,
    /// Whether vector drawings are previewed at all, ahead of the vector list their names
    /// are entries of, and of the browser engine a document of that kind is drawn by.
    pub vector_preview_enabled: bool,
    /// Memory the rendered pages may hold, in megabytes, between hovers. A page is
    /// still rendered at `0` — a document has no other source for its preview — it
    /// is simply not kept once the hover it was rendered for is over.
    pub office_cache_mb: u32,
    /// How much of what the render engine drew is kept between hovers, in megabytes. What
    /// is kept is the converted pages themselves, written under the app's own folder; see
    /// `Performance → Cache → Libre` in the tray.
    pub libre_cache_mb: u32,
    /// Which engine draws an Office document's page, which is the tray's
    /// `Engine → Select Engine → Office` setting.
    pub office_engine: OfficeEngine,
    /// How long the Office engine a family started is kept after that family's
    /// last page, which is the tray's `Engine → Microsoft Office TTL` setting.
    pub office_engine_idle: EngineIdle,
    /// How long the WebView2 engine is kept after the last document it drew. Beginning
    /// one is a browser start, and pointing a warm one at another document is a few
    /// milliseconds, so what this buys is every hover after the first.
    pub webview_idle: EngineIdle,
    /// How long the LibreOffice engine is kept after the last page it drew, which is the
    /// tray's `Engine → LibreOffice TTL` setting.
    ///
    /// The engine is kept the way the other two are — a process left running rather than a
    /// launch paid per document — and the idle time is what bounds it: `0 seconds` keeps no
    /// engine at all and is every document launched for itself, while an engine that is
    /// never let go of is a process this app holds for the rest of the run (see
    /// `libreoffice_render`).
    pub libreoffice_idle: EngineIdle,
    /// How long Explorer has to be out of reach before an engine that is not marked
    /// `Persistent` is let go, which is the tray's `Engine → AFK Timer` setting.
    ///
    /// It is the whole of what bounds such an engine: one is kept while an Explorer window
    /// is reachable and let go once none has been for this long, whatever its idle time
    /// says (see `app::afk`).
    pub afk_timer_seconds: u64,
    /// Whether the Office engines are kept whatever the user is doing, which is the
    /// `Persistent` toggle at the top of the tray's `Engine → Microsoft Office TTL`
    /// submenu. On, an engine is kept for its idle time while Explorer is minimized or
    /// behind another application; off, `afk_timer_seconds` is what bounds it.
    pub office_engine_persistent: bool,
    /// The same for the browser engine, under `Engine → WebView2 TTL`.
    pub webview_persistent: bool,
    /// The same for the engine the documents beside Office are drawn by, under
    /// `Engine → LibreOffice TTL`.
    pub libreoffice_persistent: bool,
    /// Memory the pages PDF previews were drawn as may hold, in megabytes, between
    /// hovers. A page is rendered at `0` like at any other size; it is simply not
    /// kept once the hover that asked for it is over.
    pub pdf_cache_mb: u32,
    /// Memory the frames text previews were painted as may hold, in megabytes,
    /// between hovers. A frame is painted at `0` like at any other size; it is
    /// simply not kept once the hover that asked for it is over.
    pub text_cache_mb: u32,
    /// What one hover may decode or read for, in gigabytes.
    ///
    /// The one setting here that bounds a file rather than a cache: a reader is
    /// handed it before it allocates, so a file larger than the budget is answered
    /// with no preview instead of with memory the app may not get.
    pub decode_budget_gb: f32,
    /// The curve a picture whose samples are light is brought into eight bits with — an
    /// EXR, a Radiance HDR, a float texture — as `hdr_tone_map` names it.
    pub hdr_tone_map: Curve,
    /// How many stops those pictures are shifted by before that curve, as `hdr_exposure`.
    pub hdr_exposure: f32,
    /// Whether a text preview is more than something to look at: a preview that
    /// scrolls, that can be selected and copied from, and that a pointer can rest
    /// on without closing it. Off by default, because it changes what a preview
    /// does rather than what it shows.
    pub text_preview_full_mode: bool,
    /// Font scale for text previews, as a percentage of the default size.
    pub text_font_scale_percent: u32,
    /// How far past the far edge of a text preview the pointer region reaches, in
    /// logical pixels at the display's DPI, so a hand that overshoots the edge on
    /// its way to the scrollbar does not take the preview down with it.
    pub text_scroll_far_edge_grace_pixels: f32,
    /// Extensions previewed as images, already normalized for lookup.
    pub image_extensions: Vec<String>,
    /// Extensions previewed as videos, already normalized for lookup.
    pub video_extensions: Vec<String>,
    /// Extensions previewed as text, already normalized for lookup.
    pub text_extensions: Vec<String>,
    /// File names previewed as text — the ones with no extension to match, like
    /// `LICENSE` and `Makefile` — already normalized for lookup.
    pub text_names: Vec<String>,
    /// Extensions previewed as archives, already normalized for lookup.
    pub archive_extensions: Vec<String>,
    /// Extensions previewed as Office documents, already normalized for lookup.
    pub office_extensions: Vec<String>,
    /// Extensions previewed as fonts, already normalized for lookup.
    pub font_extensions: Vec<String>,
    /// Extensions previewed as design documents and projects, already normalized for
    /// lookup.
    pub design_extensions: Vec<String>,
    /// The names of the documents the render engine is asked about, as `[libre]
    /// extensions` in `config.ini`: the formats LibreOffice reads and this app has no
    /// reader of its own for. See `libre_formats`.
    pub libre_extensions: Vec<String>,
    /// The names of the pictures the ImageMagick engine is asked about, as `[magick]
    /// extensions` in `config.ini`: the formats ImageMagick reads and this app has no
    /// reader of its own for — the camera raw formats above all. See `magick_formats`.
    pub magick_extensions: Vec<String>,
    /// The names of the archives the PeaZip engine is asked about, as `[peazip] extensions` in
    /// `config.ini`: the formats its console archiver reads and this app has no reader of its
    /// own for — the cabinet files, isos, disk images and single-stream compressors. See
    /// `peazip_formats`.
    pub peazip_extensions: Vec<String>,
    /// Extensions previewed as vector drawings, already normalized for lookup.
    pub vector_extensions: Vec<String>,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            is_first_run: false,
            run_at_startup: true,
            hover_delay_ms: DEFAULT_HOVER_DELAY_MS,
            preview_enabled: true,
            trigger_key: "alt".to_string(),
            trigger_key_mode: TriggerKeyMode::Disable,
            trigger_key_enabled: true,
            confirm_file_type: true,
            follow_cursor: DEFAULT_FOLLOW_CURSOR,
            avoid_mode: DEFAULT_AVOID_MODE,
            same_file_rehover_delay_ms: DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
            settling_delay_ms: DEFAULT_SETTLING_DELAY_MS,
            tick_ms: DEFAULT_TICK_MS,
            spinner_delay_ms: DEFAULT_SPINNER_DELAY_MS,
            webp_playback_fps: DEFAULT_WEBP_PLAYBACK_FPS,
            image_cache_mb: DEFAULT_IMAGE_CACHE_MB,
            image_background: DEFAULT_IMAGE_BACKGROUND,
            font_background: DEFAULT_FONT_BACKGROUND,
            dds_background: DEFAULT_DDS_BACKGROUND,
            design_background: DEFAULT_DESIGN_BACKGROUND,
            vector_background: DEFAULT_VECTOR_BACKGROUND,
            video_volume: DEFAULT_VIDEO_VOLUME,
            preview_scale: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video_scale: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated_scale: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            pdf_scale: DEFAULT_PDF_SCALE,
            office_scale: DEFAULT_OFFICE_SCALE,
            font_scale: DEFAULT_FONT_SCALE,
            design_scale: DEFAULT_DESIGN_SCALE,
            libre_scale: DEFAULT_LIBRE_SCALE,
            vector_scale: DEFAULT_VECTOR_SCALE,
            ttc_face: DEFAULT_TTC_FACE,
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Rendered,
            image_preview_enabled: true,
            video_preview_enabled: true,
            text_preview_enabled: true,
            pdf_preview_enabled: true,
            archive_preview_enabled: true,
            office_preview_enabled: true,
            font_preview_enabled: true,
            design_preview_enabled: true,
            libre_preview_enabled: true,
            magick_preview_enabled: true,
            peazip_preview_enabled: true,
            vector_preview_enabled: true,
            office_cache_mb: DEFAULT_OFFICE_CACHE_MB,
            libre_cache_mb: DEFAULT_LIBRE_CACHE_MB,
            office_engine: DEFAULT_OFFICE_ENGINE,
            office_engine_idle: EngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS),
            webview_idle: EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS),
            libreoffice_idle: EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS),
            afk_timer_seconds: DEFAULT_AFK_TIMER_SECS,
            office_engine_persistent: false,
            webview_persistent: false,
            libreoffice_persistent: false,
            pdf_cache_mb: DEFAULT_PDF_CACHE_MB,
            text_cache_mb: DEFAULT_TEXT_CACHE_MB,
            decode_budget_gb: DEFAULT_DECODE_BUDGET_GB,
            hdr_tone_map: DEFAULT_HDR_TONE_MAP,
            hdr_exposure: DEFAULT_HDR_EXPOSURE,
            text_preview_full_mode: false,
            text_font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            text_scroll_far_edge_grace_pixels: DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS,
            image_extensions: sanitize_image_extensions(DEFAULT_IMAGE_EXTENSIONS),
            video_extensions: sanitize_video_extensions(DEFAULT_VIDEO_EXTENSIONS),
            text_extensions: sanitize_extensions(DEFAULT_TEXT_EXTENSIONS),
            text_names: sanitize_names(DEFAULT_TEXT_NAMES),
            archive_extensions: sanitize_archive_extensions(DEFAULT_ARCHIVE_EXTENSIONS),
            office_extensions: sanitize_office_extensions(DEFAULT_OFFICE_EXTENSIONS),
            font_extensions: sanitize_font_extensions(DEFAULT_FONT_EXTENSIONS),
            design_extensions: sanitize_design_extensions(DEFAULT_DESIGN_EXTENSIONS),
            libre_extensions: sanitize_libre_extensions(DEFAULT_LIBRE_EXTENSIONS),
            magick_extensions: sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS),
            peazip_extensions: sanitize_peazip_extensions(DEFAULT_PEAZIP_EXTENSIONS),
            vector_extensions: sanitize_vector_extensions(DEFAULT_VECTOR_EXTENSIONS),
        }
    }
}

/// Whether a list holds exactly the entries the built-in list holds, order aside.
fn same_entries(list: &[String], canonical: &[String]) -> bool {
    list.len() == canonical.len() && canonical.iter().all(|entry| list.contains(entry))
}

/// The headings the settings section is written under, in the order the tray lists its
/// menus: the menu a setting is changed from is the menu it is found under, so the file
/// reads the way the tray does rather than as one alphabetical run of fifty keys.
///
/// The table is also the list of what this app writes, key for key — a test holds `save` to it —
/// so a key in it that a file does not have is a setting the file is missing, and a file missing
/// one is written out again with it (see `differs`). A setting added to the app is therefore a
/// setting that reaches the files that already exist, with nothing else to remember.
///
/// The file lists keep sections of their own below it — `[image]`, `[video]` and the rest
/// — because a list of extensions is a collection rather than a setting, and a section is
/// what the format has for a collection. A heading is a comment instead, which is all the
/// grouping needs to be: nothing reads it back, so the shape of the file stays something
/// the app writes rather than something it has to parse. `Advanced` is where the settings
/// the tray has no menu item for are kept — the ones a hand edit reaches and a menu does
/// not — and a setting `save` writes that is missing from the table is written last of
/// all under `; Ungrouped`, which is where a heading that was forgotten shows up.
///
/// Nothing reads a heading back, and a file grouped by an earlier arrangement of the tray
/// holds exactly the settings of one grouped this way — so the headings a file lists, and the
/// order it lists them in, are the one thing that tells the two apart, and the one thing that
/// brings such a file to be written again (see `headings_are_old`).
const SETTING_GROUPS: &[(&str, &[&str])] = &[
    ("General", &["preview_enabled", "run_at_startup"]),
    (
        "Preview Types",
        &[
            "archive_preview_enabled",
            "design_preview_enabled",
            "font_preview_enabled",
            "image_preview_enabled",
            "libre_preview_enabled",
            "magick_preview_enabled",
            "office_preview_enabled",
            "pdf_preview_enabled",
            "peazip_preview_enabled",
            "text_preview_enabled",
            "vector_preview_enabled",
            "video_preview_enabled",
        ],
    ),
    (
        "Text Preview",
        &[
            "markdown_mode",
            "text_font_scale",
            "text_preview_full_mode",
            "theme",
        ],
    ),
    (
        "Timing",
        &[
            "hover_delay_ms",
            "same_file_rehover_delay_ms",
            "settling_delay_ms",
            "trigger_key",
            "trigger_key_enabled",
            "trigger_key_mode",
        ],
    ),
    ("Placement", &["avoid_mode", "follow_cursor"]),
    (
        "Scaling",
        &[
            "animated_scale",
            "design_scale",
            "font_scale",
            "libre_scale",
            "office_scale",
            "pdf_scale",
            "preview_scale",
            "vector_scale",
            "video_scale",
        ],
    ),
    (
        "Background",
        &[
            "dds_background",
            "design_background",
            "font_background",
            "image_background",
            "vector_background",
        ],
    ),
    ("Volume", &["video_volume"]),
    (
        "Performance",
        &[
            "confirm_file_type",
            "decode_budget_gb",
            "image_cache_mb",
            "libre_cache_mb",
            "office_cache_mb",
            "pdf_cache_mb",
            "text_cache_mb",
            "tick_ms",
        ],
    ),
    (
        "Engine",
        &[
            "afk_timer_seconds",
            "libreoffice_idle",
            "libreoffice_persistent",
            "office_engine",
            "office_engine_idle",
            "office_engine_persistent",
            "webview_idle",
            "webview_persistent",
        ],
    ),
    (
        "Advanced",
        &[
            "hdr_exposure",
            "hdr_tone_map",
            "spinner_delay_ms",
            "text_scroll_far_edge_grace_pixels",
            "ttc_face",
            "webp_playback_fps",
        ],
    ),
];

/// One run of keys as the file writes them: the key, the value when it has one, and the
/// line's newline.
fn write_keys(out: &mut String, keys: &[(&str, &Option<String>)]) {
    for (key, value) in keys {
        out.push_str(key);
        if let Some(value) = value {
            out.push('=');
            out.push_str(value);
        }
        out.push('\n');
    }
}

/// The configuration as the text a person reads: the settings section first, under the
/// headings the tray lists its menus under, then the sections that hold the file lists in
/// alphabetical order, and the keys of every group in alphabetical order within it.
///
/// The `Ini` the values are collected in keeps its sections and keys in hash maps, so what
/// it writes by itself is a different order every time — which is what shuffled a
/// hand-edited file under the person editing it. What is returned here is the same text
/// that writer produces, in an order that does not move.
fn ordered_text(ini: &Ini) -> String {
    let map = ini.get_map_ref();

    let mut sections: Vec<&str> = map.keys().map(String::as_str).collect();
    // Settings first, the rest alphabetically: the flag is false only for the
    // settings section, and false sorts ahead of true.
    sections.sort_unstable_by_key(|section| (*section != CONFIG_SECTION, *section));

    let mut out = String::new();
    for section in sections {
        let Some(keys) = map.get(section) else {
            continue;
        };

        // Every section but the first is kept clear of what came before it: the settings
        // section ends with the last heading's last key, and a list section beginning right
        // under it reads as another key of that heading.
        if !out.is_empty() {
            out.push('\n');
        }

        out.push('[');
        out.push_str(section);
        out.push_str("]\n");

        let mut keys: Vec<(&str, &Option<String>)> = keys
            .iter()
            .map(|(key, value)| (key.as_str(), value))
            .collect();

        // A section below the settings one is one list and is written as one run.
        if section != CONFIG_SECTION {
            keys.sort_unstable_by_key(|(key, _)| *key);
            write_keys(&mut out, &keys);
            continue;
        }

        let mut written: Vec<&str> = Vec::new();
        for (heading, heading_keys) in SETTING_GROUPS {
            let mut in_heading: Vec<(&str, &Option<String>)> = keys
                .iter()
                .copied()
                .filter(|(key, _)| heading_keys.contains(key))
                .collect();
            if in_heading.is_empty() {
                continue;
            }
            in_heading.sort_unstable_by_key(|(key, _)| *key);

            // No blank line above the first heading: what it would separate it from is
            // the section's own name.
            if !written.is_empty() {
                out.push('\n');
            }
            out.push_str("; ");
            out.push_str(heading);
            out.push('\n');
            write_keys(&mut out, &in_heading);
            written.extend(in_heading.iter().map(|(key, _)| *key));
        }

        let mut ungrouped: Vec<(&str, &Option<String>)> = keys
            .into_iter()
            .filter(|(key, _)| !written.contains(key))
            .collect();
        if !ungrouped.is_empty() {
            ungrouped.sort_unstable_by_key(|(key, _)| *key);
            out.push_str("\n; Ungrouped\n");
            write_keys(&mut out, &ungrouped);
        }
    }

    out
}

/// One list as the file has it: the entries its key names, or the built-in list when the key is
/// gone — which is a file to write, since what it holds is then not what the app is using (see
/// `differs`).
///
/// What is read here is what the file says: a list anyone has edited keeps its own
/// entries and its own order, and an empty value is a list with nothing in it rather than
/// a missing one. A list this app itself wrote and has since added entries to was dealt
/// with before this ran, by the repair the file is put through as it is read (see
/// `repair_older_lists`), so by the time a list is read here it is either the user's or the
/// list of now.
fn configured_list(
    ini: &Ini,
    section: &str,
    key: &str,
    defaults: &str,
    sanitize: fn(&str) -> Vec<String>,
) -> Vec<String> {
    match ini.get(section, key) {
        Some(value) => sanitize(&value),
        None => sanitize(defaults),
    }
}

/// The built-in lists this app shipped and then changed, brought up to the list of now.
///
/// A list is only ever read out of a file — nothing in the tray edits one — so a list that
/// differs from the built-in one is either this app's own older list, written before an entry
/// was added to it or before its entries were put in alphabetical order, or an edit somebody made
/// by hand. The two are told apart by their entries, and only one of them is rewritten: a list
/// holding exactly the entries of a list this app shipped is this app's own — nobody typed it —
/// so it is replaced with the built-in list, while a list with any one entry added, removed or
/// spelled differently is the user's and is kept exactly as it is. Without this, an entry added
/// to a built-in list would reach a fresh installation only, since every file already written
/// holds the list as it was.
///
/// Telling them apart by their entries costs one thing, and it is worth saying out loud: an entry
/// a user took out can come back, because a list trimmed to exactly the entries this app shipped
/// before that entry existed is this app's own as far as this can tell, and is read as one. What
/// that buys is the other half — the formats added since, which a list nobody had touched would
/// otherwise never be given.
fn repair_older_lists(ini: &mut Ini) -> bool {
    let mut repaired = false;

    for (section, defaults, previous, sanitize) in [
        (
            IMAGE_SECTION,
            DEFAULT_IMAGE_EXTENSIONS,
            &[
                IMAGE_EXTENSIONS_BEFORE_AVCI,
                IMAGE_EXTENSIONS_BEFORE_DDS,
                IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS,
                IMAGE_EXTENSIONS_BEFORE_SVG,
                IMAGE_EXTENSIONS_WITH_SVG,
            ][..],
            sanitize_image_extensions as fn(&str) -> Vec<String>,
        ),
        (
            DESIGN_SECTION,
            DEFAULT_DESIGN_EXTENSIONS,
            &[
                DESIGN_EXTENSIONS_BEFORE_AI,
                DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE,
                DESIGN_EXTENSIONS_WITH_CDR,
            ][..],
            sanitize_design_extensions as fn(&str) -> Vec<String>,
        ),
        (
            LIBRE_SECTION,
            DEFAULT_LIBRE_EXTENSIONS,
            &[LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ][..],
            sanitize_libre_extensions as fn(&str) -> Vec<String>,
        ),
        (
            MAGICK_SECTION,
            DEFAULT_MAGICK_EXTENSIONS,
            &[MAGICK_EXTENSIONS_BEFORE_THE_REST][..],
            sanitize_magick_extensions as fn(&str) -> Vec<String>,
        ),
        (
            PEAZIP_SECTION,
            DEFAULT_PEAZIP_EXTENSIONS,
            &[PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS][..],
            sanitize_peazip_extensions as fn(&str) -> Vec<String>,
        ),
        (
            VECTOR_SECTION,
            DEFAULT_VECTOR_EXTENSIONS,
            &[
                VECTOR_EXTENSIONS_BEFORE_SVG,
                VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS,
            ][..],
            sanitize_vector_extensions as fn(&str) -> Vec<String>,
        ),
    ] {
        let Some(value) = ini.get(section, "extensions") else {
            continue;
        };

        let list = sanitize(&value);
        let canonical = sanitize(defaults);
        let written_by_the_app = same_entries(&list, &canonical)
            || previous
                .iter()
                .any(|older| same_entries(&list, &sanitize(older)));

        // A list that already agrees with the built-in one, entry for entry and in order,
        // is left alone rather than written out again: a write moves the mtime, and the
        // watcher would read the file back for a change that was not one.
        if written_by_the_app && list != canonical {
            ini.set(section, "extensions", Some(canonical.join(",")));
            repaired = true;
        }
    }

    repaired
}

/// Whether a file is one whose settings were grouped the way an earlier build grouped them.
///
/// A heading is a comment, so a file grouped by one arrangement holds exactly the settings of a
/// file grouped by another: nothing is read back from the difference and every key is read the
/// same either way. What the headings a file does list, and the order it lists them in, do say
/// is whether it was written before a heading was added, before a setting moved out of another
/// one, or before the menus themselves were rearranged — and the grouping is the app's to
/// write, so such a file is written again rather than left with the shape it happens to have.
///
/// One write puts every heading back at once, in the order of the table, so a file this has been
/// through lists them as this build lists them and is left alone the next time it is read. That
/// is the whole point of asking the question about the text rather than about the parse: the
/// file is read once a second while it is being watched, and a file the app has just written
/// must not be a file there is something to write. The question is asked of the text for that
/// reason too — a heading is a comment, and a parse keeps everything about a file except the
/// comments (`load`, `reload_from_disk`).
fn headings_are_old(text: &str) -> bool {
    // The headings the file lists, in the order it lists them. Compared against the whole line
    // rather than searched for in the file, so a key named after a heading — or a heading
    // written inside a value — is not read as one, and a line's own trailing space is not part
    // of the comparison: an editor that leaves one is not a reason to write the file again.
    let listed: Vec<&str> = text
        .lines()
        .filter_map(|line| {
            SETTING_GROUPS
                .iter()
                .map(|(heading, _)| *heading)
                .find(|heading| line.trim_end().strip_prefix("; ") == Some(*heading))
        })
        .collect();

    // Every heading the table names, in the order the tray lists them. A file missing one, or
    // listing them in another order, is a file written before the settings were arranged this
    // way — and a heading the table does not name at all, `; Ungrouped` among them, is no part
    // of the question: what is compared is the headings this app writes and the places they
    // are written in.
    let written: Vec<&str> = SETTING_GROUPS.iter().map(|(heading, _)| *heading).collect();

    listed != written
}

impl AppConfig {
    /// The app's own folder under the roaming profile, holding `config.ini` and
    /// the `theme` folder beside it.
    fn folder() -> Option<PathBuf> {
        BaseDirs::new().map(|dirs| dirs.config_dir().join("rust-hover-preview"))
    }

    pub fn config_path() -> Option<PathBuf> {
        Self::folder().map(|folder| folder.join("config.ini"))
    }

    /// The folder the user's `.tmTheme` files live in, beside `config.ini`.
    pub fn theme_dir() -> Option<PathBuf> {
        Self::folder().map(|folder| folder.join("theme"))
    }

    /// The folder rendered pages are kept in, beside `config.ini`: a design document an
    /// installed render engine drew is kept there as the PDF it was drawn as, so a
    /// document is converted once and every hover after that is a read of a file.
    pub fn rendered_dir() -> Option<PathBuf> {
        Self::folder().map(|folder| folder.join("rendered"))
    }

    /// The configuration as the file has it, with whatever the file had wrong or missing put
    /// right.
    ///
    /// The file is read in one of two ways: there is none, so a fresh installation is written —
    /// or there is one, and the lists it holds that are this app's own older ones are brought up
    /// before anything is read from it (`repair_older_lists`), and the settings it holds under
    /// headings an older build wrote are grouped again the way this build groups them
    /// (`headings_are_old`) — since the file is what the user edits and a key left missing would
    /// be repaired again on every load. It is written back where that left something to write,
    /// and where it did not, the file on disk is left exactly as it is: a file this build wrote,
    /// with nothing deleted from it and nothing added to it since, is one there is nothing to say
    /// about.
    pub fn load() -> Self {
        let mut config = Self::default();

        // The folder exists before anything can list it, so that adding a theme is
        // dropping a file into a path the app can name rather than creating one.
        theme_files::ensure();

        let Some(path) = Self::config_path() else {
            return config;
        };

        config.is_first_run = !path.exists();
        if !config.is_first_run {
            let mut ini = Ini::new();

            // A file that cannot be read is left exactly where it is, and what it holds is not
            // guessed at: writing the defaults over it would be a write with nothing to do with
            // the file, which is the one kind of write this app does not make. A file that is
            // missing by then — or was never there — is the fresh installation below.
            //
            // The text is read here rather than by `Ini::load`, which reads and parses the same
            // file: a parse keeps nothing of the comments, and whether the settings sit under the
            // headings this build writes them under is a question about the text (see
            // `headings_are_old`).
            if let Ok(text) = fs::read_to_string(&path) {
                let old_headings = headings_are_old(&text);

                if ini.read(text).is_ok() {
                    let repaired = repair_older_lists(&mut ini) || old_headings;
                    config.apply_ini(&ini);

                    // The file is written again where the repair had a list to bring up or a
                    // heading to put back, and where it does not hold what this app writes: a
                    // setting the file does not have, one whose value is not the value the app
                    // reads it back as — which is every value the app could not read at all —
                    // and any key that is not one the app writes, which is where a name this
                    // app no longer uses goes.
                    if repaired || config.differs(&ini) {
                        config.save();
                    }
                }
            }
        }

        if config.is_first_run {
            config.save();
        }

        config
    }

    /// Read the file again, after the watcher saw it change, with the same repairs a start puts
    /// it through and the same question about whether it says what the app is using.
    ///
    /// Both belong here as much as they do at a start: the file is what the user edits, and an
    /// edit that writes a value the app cannot read, adds a key of their own, or names a setting
    /// the way an older build named it, is one to put right there and then rather than at the
    /// next start — as is a file an older build wrote and this one has not read since. It settles
    /// the same way a start does — what it writes is a file that needs nothing, so the write the
    /// watcher sees after it is a file nothing further is done to.
    pub fn reload_from_disk(&mut self) {
        if let Some(path) = Self::config_path() {
            let mut ini = Ini::new();
            if let Ok(text) = fs::read_to_string(&path) {
                let old_headings = headings_are_old(&text);

                if ini.read(text).is_ok() {
                    let repaired = repair_older_lists(&mut ini) || old_headings;
                    self.apply_ini(&ini);

                    if repaired || self.differs(&ini) {
                        self.save();
                    }
                }
            }
        }
    }

    /// Write the configuration out: `to_ini` says what it holds, and `ordered_text` says what
    /// the text of the file is.
    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }

            let _ = fs::write(&path, ordered_text(&self.to_ini()));
        }
    }

    /// The settings as the file holds them: every key this build writes, with the value it
    /// writes that key as.
    ///
    /// This is what `save` writes to disk, and it is also what a file that has just been read is
    /// held up against, to tell whether it says what the app is using (see `differs`) — which is
    /// why it is a method of its own rather than the body of `save`.
    fn to_ini(&self) -> Ini {
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "run_at_startup",
            Some(self.run_at_startup.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "hover_delay_ms",
            Some(self.hover_delay_ms.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "preview_enabled",
            Some(self.preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "trigger_key",
            Some(self.trigger_key.clone()),
        );
        ini.set(
            CONFIG_SECTION,
            "trigger_key_mode",
            Some(self.trigger_key_mode.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "trigger_key_enabled",
            Some(self.trigger_key_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "confirm_file_type",
            Some(self.confirm_file_type.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "follow_cursor",
            Some(self.follow_cursor.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "avoid_mode",
            Some(self.avoid_mode.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "same_file_rehover_delay_ms",
            Some(self.same_file_rehover_delay_ms.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "settling_delay_ms",
            Some(self.settling_delay_ms.to_string()),
        );
        ini.set(CONFIG_SECTION, "tick_ms", Some(self.tick_ms.to_string()));
        ini.set(
            CONFIG_SECTION,
            "spinner_delay_ms",
            Some(sanitize_spinner_delay_ms(self.spinner_delay_ms).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "webp_playback_fps",
            Some(sanitize_webp_playback_fps(self.webp_playback_fps).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_cache_mb",
            Some(sanitize_image_cache_mb(self.image_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_background",
            Some(self.image_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "font_background",
            Some(self.font_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "dds_background",
            Some(self.dds_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "design_background",
            Some(self.design_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_background",
            Some(self.vector_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_volume",
            Some(self.video_volume.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "preview_scale",
            Some(self.preview_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_scale",
            Some(self.video_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "animated_scale",
            Some(self.animated_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_scale",
            Some(self.vector_scale.as_str()),
        );
        ini.set(CONFIG_SECTION, "pdf_scale", Some(self.pdf_scale.as_str()));
        ini.set(
            CONFIG_SECTION,
            "office_scale",
            Some(self.office_scale.as_str()),
        );
        ini.set(CONFIG_SECTION, "font_scale", Some(self.font_scale.as_str()));
        ini.set(
            CONFIG_SECTION,
            "design_scale",
            Some(self.design_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "libre_scale",
            Some(self.libre_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "ttc_face",
            Some(sanitize_ttc_face(self.ttc_face).to_string()),
        );
        ini.set(CONFIG_SECTION, "theme", Some(self.theme.as_str()));
        ini.set(
            CONFIG_SECTION,
            "markdown_mode",
            Some(self.markdown_mode.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_preview_enabled",
            Some(self.image_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_preview_enabled",
            Some(self.video_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_preview_enabled",
            Some(self.text_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pdf_preview_enabled",
            Some(self.pdf_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "archive_preview_enabled",
            Some(self.archive_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_preview_enabled",
            Some(self.office_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "font_preview_enabled",
            Some(self.font_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "design_preview_enabled",
            Some(self.design_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "libre_preview_enabled",
            Some(self.libre_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "magick_preview_enabled",
            Some(self.magick_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "peazip_preview_enabled",
            Some(self.peazip_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_preview_enabled",
            Some(self.vector_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_cache_mb",
            Some(sanitize_office_cache_mb(self.office_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "libre_cache_mb",
            Some(sanitize_libre_cache_mb(self.libre_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine",
            Some(self.office_engine.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine_idle",
            Some(self.office_engine_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "webview_idle",
            Some(self.webview_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "libreoffice_idle",
            Some(self.libreoffice_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "afk_timer_seconds",
            Some(sanitize_afk_timer_secs(self.afk_timer_seconds).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine_persistent",
            Some(self.office_engine_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "webview_persistent",
            Some(self.webview_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "libreoffice_persistent",
            Some(self.libreoffice_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pdf_cache_mb",
            Some(sanitize_pdf_cache_mb(self.pdf_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_cache_mb",
            Some(sanitize_text_cache_mb(self.text_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "decode_budget_gb",
            Some(sanitize_decode_budget_gb(self.decode_budget_gb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "hdr_tone_map",
            Some(self.hdr_tone_map.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "hdr_exposure",
            Some(sanitize_hdr_exposure(self.hdr_exposure).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_preview_full_mode",
            Some(self.text_preview_full_mode.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_font_scale",
            Some(sanitize_text_font_scale_percent(self.text_font_scale_percent).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_scroll_far_edge_grace_pixels",
            Some(
                sanitize_text_scroll_far_edge_grace_pixels(self.text_scroll_far_edge_grace_pixels)
                    .to_string(),
            ),
        );
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(sanitize_image_extensions(&self.image_extensions.join(",")).join(",")),
        );
        ini.set(
            VIDEO_SECTION,
            "extensions",
            Some(sanitize_video_extensions(&self.video_extensions.join(",")).join(",")),
        );
        ini.set(
            TEXT_SECTION,
            "extensions",
            Some(sanitize_extensions(&self.text_extensions.join(",")).join(",")),
        );
        ini.set(
            TEXT_SECTION,
            "names",
            Some(sanitize_names(&self.text_names.join(",")).join(",")),
        );
        ini.set(
            ARCHIVE_SECTION,
            "extensions",
            Some(sanitize_archive_extensions(&self.archive_extensions.join(",")).join(",")),
        );
        ini.set(
            OFFICE_SECTION,
            "extensions",
            Some(sanitize_office_extensions(&self.office_extensions.join(",")).join(",")),
        );
        ini.set(
            FONT_SECTION,
            "extensions",
            Some(sanitize_font_extensions(&self.font_extensions.join(",")).join(",")),
        );
        ini.set(
            DESIGN_SECTION,
            "extensions",
            Some(sanitize_design_extensions(&self.design_extensions.join(",")).join(",")),
        );
        ini.set(
            LIBRE_SECTION,
            "extensions",
            Some(sanitize_libre_extensions(&self.libre_extensions.join(",")).join(",")),
        );
        ini.set(
            MAGICK_SECTION,
            "extensions",
            Some(sanitize_magick_extensions(&self.magick_extensions.join(",")).join(",")),
        );
        ini.set(
            PEAZIP_SECTION,
            "extensions",
            Some(sanitize_peazip_extensions(&self.peazip_extensions.join(",")).join(",")),
        );
        ini.set(
            VECTOR_SECTION,
            "extensions",
            Some(sanitize_vector_extensions(&self.vector_extensions.join(",")).join(",")),
        );
        ini
    }

    /// Whether the file holds something other than what this app would write for the settings
    /// it has just read, which is what makes it a file to write back. The two are held up against
    /// each other both ways round: every key the app writes has to be in the file, holding what
    /// the app writes for it, and every key in the file has to be one the app writes.
    ///
    /// A key the file does not have is a difference, and so is one it holds another value for. A
    /// value the app could not read at all is that second kind, and so is one a setting reduced to
    /// what it allows — a delay past its ceiling, a face of a collection past the last one the
    /// menu offers, a tone map that is not one of them. A key the app does not write is a
    /// difference too: the file is what this app writes and nothing else, so a key of the user's
    /// own, or one left behind by an older or a newer build, is dropped by the write it asks for.
    ///
    /// What is not compared is what is not a key: a comment, a blank line, the order the keys are
    /// written in — so a comment survives until a write happens for another reason, which is when
    /// `save` writes the file out from the settings it knows and the comment goes with it.
    fn differs(&self, ini: &Ini) -> bool {
        let wanted = self.to_ini();

        for (section, keys) in wanted.get_map_ref() {
            for (key, value) in keys {
                if ini.get(section, key) != *value {
                    return true;
                }
            }
        }

        for (section, keys) in ini.get_map_ref() {
            for key in keys.keys() {
                if wanted.get(section, key).is_none() {
                    return true;
                }
            }
        }

        false
    }

    /// Read the file into the configuration: what the file says, under the names of now, for
    /// every setting it has. A setting it does not have keeps the value it already holds, which
    /// is the default for a configuration that has just been made — and whether the file needs
    /// writing afterwards is not this read's answer to give (see `differs`).
    fn apply_ini(&mut self, ini: &Ini) {
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "run_at_startup") {
            self.run_at_startup = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "hover_delay_ms") {
            self.hover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "preview_enabled") {
            self.preview_enabled = value;
        }
        // The key the trigger watches, by name. A file that names it says what it is; a name an
        // older build wrote it under is not a name this app reads at all, and the line goes when
        // the file is written again (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "trigger_key") {
            let value = value.trim();
            if !value.is_empty() {
                self.trigger_key = value.to_string();
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "trigger_key_mode") {
            if let Some(mode) = TriggerKeyMode::from_str(&value) {
                self.trigger_key_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "trigger_key_enabled") {
            self.trigger_key_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "confirm_file_type") {
            self.confirm_file_type = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "follow_cursor") {
            self.follow_cursor = value;
        }
        // How far a preview is kept off its item, by the name it is written under now. The yes
        // or no question this used to be is not read: an older name is a line the app does not
        // write, and the file is written again without it (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "avoid_mode") {
            if let Some(mode) = AvoidMode::from_str(&value) {
                self.avoid_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "same_file_rehover_delay_ms") {
            self.same_file_rehover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "settling_delay_ms") {
            self.settling_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "tick_ms") {
            self.tick_ms = sanitize_tick_ms(value);
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "spinner_delay_ms") {
            self.spinner_delay_ms = sanitize_spinner_delay_ms(value);
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "webp_playback_fps") {
            if let Ok(value) = u32::try_from(value) {
                self.webp_playback_fps = sanitize_webp_playback_fps(value);
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "image_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.image_cache_mb = sanitize_image_cache_mb(value);
            }
        }
        // A picture's backdrop, by its own name. The one setting every preview was drawn over
        // before each kind had one of its own is not read from a file that still names it: the
        // line is one the app does not write, and it goes with the next write (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "image_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.image_background = background;
            }
        }
        // The backdrop a vector drawing is drawn over, which is one setting for the two halves
        // of the kind — the documents the browser draws and the metafiles the drawing layer
        // replays — and one name for both.
        if let Some(value) = ini.get(CONFIG_SECTION, "vector_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.vector_background = background;
            }
        }
        // A font's backdrop is read from its own name, which it always had: fonts are a kind of
        // its own, so there is no earlier spelling of this one to answer.
        if let Some(value) = ini.get(CONFIG_SECTION, "font_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.font_background = background;
            }
        }
        // A texture's is read the same way, and for the same reason: the DDS kind is one of
        // its own, so a file written before it has nothing under this name and what it wrote
        // about a picture stays what it wrote about a picture. Two of the four backdrops are
        // not offered for a texture — the two that show what stands behind it — so a file
        // that names one of them is read as the backdrop this setting starts at, rather than
        // kept as a value the menu beside it has no item for.
        if let Some(value) = ini.get(CONFIG_SECTION, "dds_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.dds_background = sanitize_dds_background(background);
            }
        }
        // A design document's is read the same way and for the same reason: the kind is
        // one of its own, so a file written before it has nothing under this name, and
        // what it wrote about a picture stays what it wrote about a picture.
        if let Some(value) = ini.get(CONFIG_SECTION, "design_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.design_background = background;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "video_volume") {
            if let Ok(value) = u32::try_from(value) {
                self.video_volume = value;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "preview_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.preview_scale = scale;
            }
        }
        // A video's scale is written the way a picture's is, and read the same way. A file with no
        // key for it leaves the setting where a fresh installation starts: the picture scale is
        // not an answer to this question, and a file that says nothing about it is a file that
        // says nothing about it (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "video_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.video_scale = scale;
            }
        }
        // And an animation's, the same way again.
        if let Some(value) = ini.get(CONFIG_SECTION, "animated_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.animated_scale = scale;
            }
        }
        // A drawing's scale is the one setting an SVG document and a metafile share — they
        // are one kind in the tray — and it is read from its own name. What the number is a
        // percentage of is the room the display has rather than a size the file asks for.
        if let Some(value) = ini.get(CONFIG_SECTION, "vector_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.vector_scale = scale;
            }
        }
        // A page's scale is read the same way and against the same whole: what the
        // number is a percentage of is the room the display has, one setting per kind
        // of document.
        if let Some(value) = ini.get(CONFIG_SECTION, "pdf_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.pdf_scale = scale;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "office_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.office_scale = scale;
            }
        }
        // A specimen's scale is read the same way again, against the share of the display
        // the box `font_preview` measures a font at takes.
        if let Some(value) = ini.get(CONFIG_SECTION, "font_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.font_scale = scale;
            }
        }
        // And a design document's, against the room the display has for the picture the
        // file keeps of the document.
        if let Some(value) = ini.get(CONFIG_SECTION, "design_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.design_scale = scale;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "libre_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.libre_scale = scale;
            }
        }
        // And which face of a collection the specimen is of, in the numbering the tray's
        // `Font Face` submenu offers it in.
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "ttc_face") {
            if let Ok(value) = u32::try_from(value) {
                self.ttc_face = sanitize_ttc_face(value);
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "theme") {
            if let Some(theme) = TextTheme::resolve(&value) {
                self.theme = theme;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "markdown_mode") {
            if let Some(mode) = MarkdownMode::from_str(&value) {
                self.markdown_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "image_preview_enabled") {
            self.image_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_preview_enabled") {
            self.video_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_enabled") {
            self.text_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pdf_preview_enabled") {
            self.pdf_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "archive_preview_enabled") {
            self.archive_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "office_preview_enabled") {
            self.office_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "font_preview_enabled") {
            self.font_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "design_preview_enabled") {
            self.design_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "libre_preview_enabled") {
            self.libre_preview_enabled = value;
        }
        // The ImageMagick kind's switch is read from its own name, and stays where it is
        // where the name is not written at all.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "magick_preview_enabled") {
            self.magick_preview_enabled = value;
        }
        // And the PeaZip kind's, read and left where it is for the same reason.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "peazip_preview_enabled") {
            self.peazip_preview_enabled = value;
        }
        // The vector kind's switch is read from its own name, and stays where it is where the
        // name is not written at all.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "vector_preview_enabled") {
            self.vector_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "office_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.office_cache_mb = sanitize_office_cache_mb(value);
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "libre_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.libre_cache_mb = sanitize_libre_cache_mb(value);
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "office_engine") {
            if let Some(engine) = OfficeEngine::from_str(&value) {
                self.office_engine = engine;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "office_engine_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.office_engine_idle = idle;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "webview_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.webview_idle = idle;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "libreoffice_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.libreoffice_idle = idle;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "afk_timer_seconds") {
            self.afk_timer_seconds = sanitize_afk_timer_secs(value);
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "office_engine_persistent") {
            self.office_engine_persistent = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "webview_persistent") {
            self.webview_persistent = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "libreoffice_persistent") {
            self.libreoffice_persistent = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pdf_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.pdf_cache_mb = sanitize_pdf_cache_mb(value);
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "text_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.text_cache_mb = sanitize_text_cache_mb(value);
            }
        }
        // A budget is written in gigabytes and may be fractional — `0.5` is a small
        // machine's ceiling — so it is read as the number it is rather than as a
        // count of them.
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "decode_budget_gb") {
            self.decode_budget_gb = sanitize_decode_budget_gb(value as f32);
        }
        // A curve is read by name, and a name that is not one of them is answered with the
        // default this field already holds rather than with a curve picked at random: what
        // the setting says has to be a curve for it to be used.
        if let Some(value) = ini.get(CONFIG_SECTION, "hdr_tone_map") {
            if let Some(curve) = Curve::from_str(&value) {
                self.hdr_tone_map = curve;
            }
        }
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "hdr_exposure") {
            self.hdr_exposure = sanitize_hdr_exposure(value as f32);
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_full_mode") {
            self.text_preview_full_mode = value;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "text_font_scale") {
            if let Some(scale) = parse_text_font_scale(&value) {
                self.text_font_scale_percent = scale;
            }
        }
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "text_scroll_far_edge_grace_pixels") {
            self.text_scroll_far_edge_grace_pixels =
                sanitize_text_scroll_far_edge_grace_pixels(value as f32);
        }
        // A list is what the file says it is, and a key that is gone is a list the
        // file no longer has: the built-in entries are put back, and the file is written out
        // again because it does not say what the app is using. An empty value is not the same
        // thing — it is a list the user emptied, and it is kept as written.
        //
        // The image list is read as the file has it. What an older file's list needs —
        // `svg` and `svgz` given up to the vector list, the formats Windows has a codec
        // for, `dds`, and the order the entries are written in — was seen to by
        // `repair_older_lists`, which ran over the file before this did.
        let list = configured_list(
            ini,
            IMAGE_SECTION,
            "extensions",
            DEFAULT_IMAGE_EXTENSIONS,
            sanitize_image_extensions,
        );
        self.image_extensions = list;
        let list = configured_list(
            ini,
            VIDEO_SECTION,
            "extensions",
            DEFAULT_VIDEO_EXTENSIONS,
            sanitize_video_extensions,
        );
        self.video_extensions = list;
        let list = configured_list(
            ini,
            TEXT_SECTION,
            "extensions",
            DEFAULT_TEXT_EXTENSIONS,
            sanitize_extensions,
        );
        self.text_extensions = list;
        let list = configured_list(
            ini,
            TEXT_SECTION,
            "names",
            DEFAULT_TEXT_NAMES,
            sanitize_names,
        );
        self.text_names = list;
        let list = configured_list(
            ini,
            ARCHIVE_SECTION,
            "extensions",
            DEFAULT_ARCHIVE_EXTENSIONS,
            sanitize_archive_extensions,
        );
        self.archive_extensions = list;
        let list = configured_list(
            ini,
            OFFICE_SECTION,
            "extensions",
            DEFAULT_OFFICE_EXTENSIONS,
            sanitize_office_extensions,
        );
        self.office_extensions = list;
        // The font list is one whose built-in entries are new with the kind itself, so an
        // older file simply has no section: the key is gone, the built-in entries come back
        // with it, and the file is written out again with them.
        let list = configured_list(
            ini,
            FONT_SECTION,
            "extensions",
            DEFAULT_FONT_EXTENSIONS,
            sanitize_font_extensions,
        );
        self.font_extensions = list;
        // The design list is new with the kind itself, the way the font list above is: an
        // older file has no section at all, so the key is gone, the built-in entries come
        // back with it, and the file is written out again holding them. Its entries have
        // grown since — `ai` is the one that did — and that is `repair_older_lists`' business
        // rather than this read's.
        let list = configured_list(
            ini,
            DESIGN_SECTION,
            "extensions",
            DEFAULT_DESIGN_EXTENSIONS,
            sanitize_design_extensions,
        );
        self.design_extensions = list;
        // The render engine's list, new with its kind: an older file has no section at all, so
        // the key is gone and the built-in entries come back with it. What a file holds is read
        // back like every other list — a name a user added is asked about from the next read, and
        // one they took out is not — and the entries the engine has no filter for, which an
        // earlier list carried, are taken out of a file by `repair_older_lists` before this runs.
        let list = configured_list(
            ini,
            LIBRE_SECTION,
            "extensions",
            DEFAULT_LIBRE_EXTENSIONS,
            sanitize_libre_extensions,
        );
        self.libre_extensions = list;
        // And the ImageMagick engine's, the same shape once more: the camera raw formats above
        // all, written from the built-in list on the first run and read back from there, so a
        // user can add a format the engine reads and this app does not know.
        let list = configured_list(
            ini,
            MAGICK_SECTION,
            "extensions",
            DEFAULT_MAGICK_EXTENSIONS,
            sanitize_magick_extensions,
        );
        self.magick_extensions = list;
        // And the PeaZip engine's, the same shape once more: the archives its console archiver
        // reads and this app has no reader of its own for — the cabinet files, isos, disk
        // images and single-stream compressors — written from the built-in list on the first
        // run and read back from there, so a user can add a format the engine reads and this app
        // does not know, or take one out.
        let list = configured_list(
            ini,
            PEAZIP_SECTION,
            "extensions",
            DEFAULT_PEAZIP_EXTENSIONS,
            sanitize_peazip_extensions,
        );
        self.peazip_extensions = list;
        // And the vector list, new with its kind: an older file has no section at all, so
        // the key is gone and the built-in entries come back with it. Its entries have
        // grown since — `svg` and `svgz`, which were entries of the image list until the
        // kind they belong to was given them — and that, too, is `repair_older_lists`'
        // business rather than this read's.
        let list = configured_list(
            ini,
            VECTOR_SECTION,
            "extensions",
            DEFAULT_VECTOR_EXTENSIONS,
            sanitize_vector_extensions,
        );
        self.vector_extensions = list;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn avoid_mode_reads_back_what_it_writes() {
        for mode in [
            AvoidMode::Off,
            AvoidMode::Filename,
            AvoidMode::FilenameColumn,
            AvoidMode::Details,
        ] {
            let written = mode.as_str();
            assert_eq!(
                AvoidMode::from_str(written),
                Some(mode),
                "`{written}` read back"
            );
        }

        assert_eq!(AvoidMode::from_str("  DETAILS "), Some(AvoidMode::Details));
        assert_eq!(
            AvoidMode::from_str("every column"),
            None,
            "a value that names no way of avoiding is not one"
        );
    }

    /// The name-only way of avoiding is what the app does unless it is told otherwise,
    /// so a configuration that says nothing about it keeps a preview off the name.
    #[test]
    fn a_fresh_configuration_avoids_the_name_alone() {
        assert_eq!(AppConfig::default().avoid_mode, AvoidMode::Filename);
    }

    #[test]
    fn office_engine_idle_reads_back_what_it_writes() {
        for idle in [
            EngineIdle::Seconds(0),
            EngineIdle::Seconds(600),
            EngineIdle::Seconds(MAX_OFFICE_ENGINE_IDLE_SECS),
            EngineIdle::Indefinite,
        ] {
            let written = idle.as_str();
            assert_eq!(
                EngineIdle::from_str(&written),
                Some(idle),
                "`{written}` read back"
            );
        }
    }

    #[test]
    fn office_engine_idle_takes_the_words_a_person_would_write() {
        for written in ["indefinitely", "Indefinite", " forever ", "ALWAYS"] {
            assert_eq!(
                EngineIdle::from_str(written),
                Some(EngineIdle::Indefinite),
                "`{written}`"
            );
        }

        assert_eq!(
            EngineIdle::from_str(" 900 "),
            Some(EngineIdle::Seconds(900))
        );
        assert_eq!(
            EngineIdle::from_str("999999"),
            Some(EngineIdle::Seconds(MAX_OFFICE_ENGINE_IDLE_SECS)),
            "a number past the ceiling is reduced to it"
        );
        assert_eq!(
            EngineIdle::from_str("soon"),
            None,
            "a value that is neither a time nor a word is not one"
        );
    }

    #[test]
    fn office_engine_idle_expires_on_its_own_clock() {
        let minute = Duration::from_secs(60);

        assert!(
            EngineIdle::Seconds(0).has_expired(Duration::ZERO),
            "an engine let go as soon as it has drawn a page is idle at once"
        );
        assert!(!EngineIdle::Seconds(60).has_expired(minute - Duration::from_secs(1)));
        assert!(EngineIdle::Seconds(60).has_expired(minute));
        assert!(EngineIdle::Seconds(60).has_expired(Duration::from_secs(9_999)));
        assert!(
            !EngineIdle::Indefinite.has_expired(Duration::from_secs(365 * 24 * 60 * 60)),
            "an engine kept for the life of the app never goes idle"
        );
    }

    /// A drawing is drawn at the whole room the display has unless the file says
    /// otherwise — which is what the setting starts at, and what a fresh install
    /// writes.
    #[test]
    fn a_drawings_scale_starts_at_the_whole_room() {
        let config = AppConfig::default();

        assert_eq!(config.vector_scale, PreviewScale::FitToScreen);
        assert_eq!(config.vector_scale.as_str(), "fit");
    }

    /// A page's scale is a setting of its own as well: a PDF, a page Office rendered
    /// and a hand-edited picture scale are three answers to three questions, and one
    /// of them changing leaves the others where they were.
    #[test]
    fn a_pages_scale_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
        ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
        ini.set(CONFIG_SECTION, "pdf_scale", Some("25".to_string()));
        ini.set(CONFIG_SECTION, "office_scale", Some("10".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.pdf_scale, PreviewScale::Percent(25));
        assert_eq!(config.office_scale, PreviewScale::Percent(10));
        assert_eq!(config.vector_scale, PreviewScale::Percent(75));
        assert_eq!(config.preview_scale, PreviewScale::Percent(400));

        // The words a person would write are read for a page key the same way they are
        // for a document's, since it is the same value read against the same whole.
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "pdf_scale",
            Some(" Fit to Screen ".to_string()),
        );
        ini.set(CONFIG_SECTION, "office_scale", Some("50%".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.pdf_scale, PreviewScale::FitToScreen);
        assert_eq!(config.office_scale, PreviewScale::Percent(50));
    }

    /// The same for a video: a video's scale is a setting of its own like the four
    /// document scales beside it, so one key changing leaves the others where they were.
    /// A picture and a video are drawn at the same share by default, which the two keys
    /// keep apart rather than sharing.
    #[test]
    fn a_videos_scale_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
        ini.set(CONFIG_SECTION, "video_scale", Some("50".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.video_scale, PreviewScale::Percent(50));
        assert_eq!(config.preview_scale, PreviewScale::Percent(400));

        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "video_scale",
            Some(" Fit to Screen ".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(config.video_scale, PreviewScale::FitToScreen);
    }

    /// A video's scale is a setting of its own: a file that says nothing about one leaves it
    /// where a fresh installation starts, whatever the file says its pictures are drawn at —
    /// one share is not an answer to the other's question.
    #[test]
    fn a_file_without_a_video_scale_leaves_it_where_it_starts() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("25".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.preview_scale, PreviewScale::Percent(25));
        assert_eq!(config.video_scale, DEFAULT_VIDEO_SCALE);

        let config = AppConfig::default();

        assert_eq!(config.preview_scale, DEFAULT_PREVIEW_SCALE);
        assert_eq!(config.video_scale, DEFAULT_VIDEO_SCALE);
    }

    /// And the same for an animation: a setting of its own beside the video and picture
    /// scales, so one key changing leaves the others where they were — and a file that has
    /// no key for it leaves it where a fresh installation starts, like the video's.
    #[test]
    fn an_animations_scale_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
        ini.set(CONFIG_SECTION, "video_scale", Some("75".to_string()));
        ini.set(CONFIG_SECTION, "animated_scale", Some("50".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.animated_scale, PreviewScale::Percent(50));
        assert_eq!(config.video_scale, PreviewScale::Percent(75));
        assert_eq!(config.preview_scale, PreviewScale::Percent(400));

        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "animated_scale",
            Some(" Fit to Screen ".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(config.animated_scale, PreviewScale::FitToScreen);

        // A file with no key for it leaves the setting where a fresh installation starts,
        // whatever share the file draws its pictures at.
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("25".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.animated_scale, DEFAULT_ANIMATED_SCALE);

        let config = AppConfig::default();

        assert_eq!(config.animated_scale, DEFAULT_ANIMATED_SCALE);
    }

    /// A page is drawn at the whole room the display has unless the file says
    /// otherwise — which is what the two page scales start at, and what a fresh
    /// install writes.
    #[test]
    fn a_pages_scale_starts_at_the_whole_room() {
        let config = AppConfig::default();

        assert_eq!(config.pdf_scale, PreviewScale::FitToScreen);
        assert_eq!(config.pdf_scale.as_str(), "fit");
        assert_eq!(config.office_scale, PreviewScale::FitToScreen);
        assert_eq!(config.office_scale.as_str(), "fit");
    }

    /// The scale is written the way the picture scale is, so the words a person
    /// would write by hand are the words it reads: `fit`, a number, either with a
    /// percent sign or without.
    #[test]
    fn a_drawing_scale_takes_the_words_a_person_would_write() {
        for (written, expected) in [
            ("fit", PreviewScale::FitToScreen),
            ("Fit to Screen", PreviewScale::FitToScreen),
            (" 75 ", PreviewScale::Percent(75)),
            ("25%", PreviewScale::Percent(25)),
            ("10", PreviewScale::Percent(10)),
        ] {
            let mut ini = Ini::new();
            ini.set(CONFIG_SECTION, "vector_scale", Some(written.to_string()));

            let config = read_file(&mut ini);

            assert_eq!(config.vector_scale, expected, "`{written}` read back");
        }
    }

    /// A specimen's scale is the fourth of them and a setting of its own like the three:
    /// one key changing leaves the others where they were.
    #[test]
    fn a_specimens_scale_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
        ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
        ini.set(CONFIG_SECTION, "pdf_scale", Some("25".to_string()));
        ini.set(CONFIG_SECTION, "office_scale", Some("10".to_string()));
        ini.set(
            CONFIG_SECTION,
            "font_scale",
            Some(" Fit to Screen ".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(config.font_scale, PreviewScale::FitToScreen);
        assert_eq!(config.vector_scale, PreviewScale::Percent(75));
        assert_eq!(config.pdf_scale, PreviewScale::Percent(25));
        assert_eq!(config.office_scale, PreviewScale::Percent(10));
        assert_eq!(config.preview_scale, PreviewScale::Percent(400));
    }

    /// A design document's scale is the fifth of them and a setting of its own like the
    /// four: it starts at the whole of the room, and one key changing leaves the others
    /// where they were.
    #[test]
    fn a_designs_scale_is_read_from_its_own_key() {
        assert_eq!(DEFAULT_DESIGN_SCALE, PreviewScale::FitToScreen);
        assert_eq!(AppConfig::default().design_scale, DEFAULT_DESIGN_SCALE);

        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
        ini.set(CONFIG_SECTION, "pdf_scale", Some("25".to_string()));
        ini.set(CONFIG_SECTION, "office_scale", Some("fit".to_string()));
        ini.set(CONFIG_SECTION, "font_scale", Some("50".to_string()));
        ini.set(CONFIG_SECTION, "design_scale", Some(" 10% ".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.design_scale, PreviewScale::Percent(10));
        assert_eq!(config.vector_scale, PreviewScale::Percent(75));
        assert_eq!(config.pdf_scale, PreviewScale::Percent(25));
        assert_eq!(config.office_scale, PreviewScale::FitToScreen);
        assert_eq!(config.font_scale, PreviewScale::Percent(50));
    }

    /// Which face of a collection a specimen is of is a key of its own: a file that has never
    /// named one previews the first face, one that names a face is read at it, and a number
    /// outside the range the menu offers is brought back into it rather than kept.
    #[test]
    fn the_face_of_a_collection_is_read_from_its_own_key() {
        let config = AppConfig::default();
        assert_eq!(config.ttc_face, DEFAULT_TTC_FACE);

        for (written, expected) in [
            ("3", 3),
            ("10", MAX_TTC_FACE),
            ("0", DEFAULT_TTC_FACE),
            ("99", MAX_TTC_FACE),
        ] {
            let mut ini = Ini::new();
            ini.set(CONFIG_SECTION, "ttc_face", Some(written.to_string()));

            let mut config = AppConfig::default();
            config.apply_ini(&ini);

            assert_eq!(config.ttc_face, expected, "`{written}` read back");
        }

        // A face and a specimen's scale are two keys: one does not move the other.
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "font_scale", Some("25".to_string()));
        ini.set(CONFIG_SECTION, "ttc_face", Some("2".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.ttc_face, 2);
        assert_eq!(config.font_scale, PreviewScale::Percent(25));
    }

    /// A specimen is drawn at half the room the display has unless the file says otherwise:
    /// the share an SVG document starts at, which is what a fresh install writes and what
    /// the menu marks as the default.
    #[test]
    fn a_specimens_scale_starts_at_half_the_room() {
        let config = AppConfig::default();

        assert_eq!(config.font_scale, PreviewScale::Percent(50));
        assert_eq!(config.font_scale.as_str(), "50");
        assert_eq!(config.font_scale, DEFAULT_FONT_SCALE);
    }

    /// A specimen's backdrop is a key of its own: fonts are a kind of their own, so the setting
    /// is read from its own name and from nothing else.
    #[test]
    fn a_specimens_backdrop_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "font_background",
            Some("checkerboard".to_string()),
        );

        let config = read_file(&mut ini);
        assert_eq!(config.font_background, TransparentBackground::Checkerboard);

        // A file that says nothing about one leaves the setting where it starts, which for a
        // specimen is the page it is written on.
        let ini = Ini::new();

        let mut config = AppConfig::default();
        config.apply_ini(&ini);
        assert_eq!(config.font_background, DEFAULT_FONT_BACKGROUND);
    }

    /// A design document's backdrop is a key of its own for the reason a specimen's is:
    /// the kind is one of its own, so the key a picture was written under is read for a
    /// picture and leaves this one where it starts.
    #[test]
    fn a_designs_backdrop_is_read_from_its_own_key() {
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "image_background",
            Some("white".to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "design_background",
            Some("checkerboard".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.design_background,
            TransparentBackground::Checkerboard
        );
        assert_eq!(config.image_background, TransparentBackground::White);
    }

    /// The font list is the fifth of them and behaves like the rest: a file that has never
    /// named it is written out with the built-in entries, and one that has been edited keeps
    /// what the user wrote.
    #[test]
    fn the_font_list_is_written_out_and_read_back() {
        let config = AppConfig::default();
        assert_eq!(
            config.font_extensions,
            sanitize_font_extensions(DEFAULT_FONT_EXTENSIONS)
        );

        let mut ini = Ini::new();
        ini.set(FONT_SECTION, "extensions", Some(".OTF,ttf".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(config.font_extensions, vec!["otf", "ttf"]);

        // A file with no section at all is answered with the built-in list, and the file is one
        // to write out again with it, since a key that is gone is not what the app is using.
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "run_at_startup", Some("true".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(
            config.font_extensions,
            sanitize_font_extensions(DEFAULT_FONT_EXTENSIONS)
        );
        assert!(
            config.differs(&ini),
            "and the file, which has no such section, is one to write"
        );
    }

    /// Fonts are their own kind under `Preview Types`: the gate is a switch of its own, and
    /// switching it leaves the list and every other kind where they were.
    #[test]
    fn the_fonts_gate_is_a_switch_of_its_own() {
        let mut config = AppConfig::default();
        assert!(PreviewType::Fonts.enabled_in(&config));

        PreviewType::Fonts.set_enabled_in(&mut config, false);
        assert!(!PreviewType::Fonts.enabled_in(&config));
        assert!(PreviewType::Vector.enabled_in(&config));
        assert!(PreviewType::Images.enabled_in(&config));

        PreviewType::Fonts.set_enabled_in(&mut config, true);
        assert!(PreviewType::Fonts.enabled_in(&config));
    }

    /// The pictures the ImageMagick engine is asked about are a list of their own in a section
    /// of their own — written from the built-in list on first run, normalized on the way in
    /// and out, and read back from the file — and the kind they belong to has a gate of its
    /// own under `Preview Types`.
    ///
    /// What the engine's own work costs is not a setting of its own, and there is nothing to
    /// assert about one: a picture it develops is held in the image cache, under the budget
    /// pictures have always been held under, and a file whose frame has been given up is
    /// developed again rather than kept in a file of the app's own (see `imagemagick_render`).
    #[test]
    fn the_pictures_the_magick_engine_is_asked_about_are_a_list_and_a_gate_of_their_own() {
        let config = AppConfig::default();
        assert_eq!(
            config.magick_extensions,
            sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS)
        );

        let mut ini = Ini::new();
        ini.set(MAGICK_SECTION, "extensions", Some(".NEF, cr3".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(config.magick_extensions, vec!["nef", "cr3"]);

        // A section that is gone is answered with the built-in list, and the file is one to
        // write out again with it, since a key that is gone is not what the app is using.
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "run_at_startup", Some("true".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(
            config.magick_extensions,
            sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS)
        );
        assert!(
            config.differs(&ini),
            "and the file, which has no such section, is one to write"
        );

        // And a file holding the list this app shipped before the names nobody had asked for were
        // added to it is this app's own rather than a user's edit, so it is brought up to the list
        // of now — which is how an installation that already exists is given them.
        let mut ini = Ini::new();
        ini.set(
            MAGICK_SECTION,
            "extensions",
            Some(MAGICK_EXTENSIONS_BEFORE_THE_REST.to_string()),
        );

        let config = read_file(&mut ini);
        assert_eq!(
            config.magick_extensions,
            sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS)
        );
        for name in ["sun", "pict", "rgb", "ase", "fax"] {
            assert!(
                config.magick_extensions.iter().any(|entry| entry == name),
                "`{name}` is one of the names added since"
            );
        }

        // The gate is a switch of its own: switching it leaves every other kind and the list
        // where they were.
        let mut config = AppConfig::default();
        assert!(PreviewType::Magick.enabled_in(&config));

        PreviewType::Magick.set_enabled_in(&mut config, false);
        assert!(!PreviewType::Magick.enabled_in(&config));
        assert!(PreviewType::Libre.enabled_in(&config));
        assert!(PreviewType::Images.enabled_in(&config));
        assert_eq!(
            config.magick_extensions,
            sanitize_magick_extensions(DEFAULT_MAGICK_EXTENSIONS)
        );

        PreviewType::Magick.set_enabled_in(&mut config, true);
        assert!(PreviewType::Magick.enabled_in(&config));
    }

    /// A file as the app reads one: what it has wrong or missing is put right first, and then
    /// what the configuration reads is what the file says.
    fn read_file(ini: &mut Ini) -> AppConfig {
        repair_older_lists(ini);

        let mut config = AppConfig::default();
        config.apply_ini(ini);
        config
    }

    /// A file holding the built-in list of an earlier version has never been edited,
    /// so the entries added to the list since then are put into it. Without this, a
    /// format added to a built-in list would preview on a fresh installation only.
    #[test]
    fn a_list_holding_the_apps_own_older_image_entries_takes_the_ones_added_to_it() {
        let mut ini = Ini::new();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_BEFORE_SVG.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.image_extensions,
            sanitize_image_extensions(DEFAULT_IMAGE_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
    }

    /// And the same for the `[peazip]` list, which is new to this table with the backends beside
    /// the console archiver: a file holding the entries this app shipped before those tools were
    /// driven has never been edited, so it is brought up to the list of now — which is how an
    /// installation that already exists is given the names the archiver's own table never
    /// declared, and is what keeps a `.arc` or a `.br` from previewing on a fresh installation
    /// only.
    #[test]
    fn a_list_holding_the_apps_own_peazip_entries_takes_the_backends_added_to_them() {
        let mut ini = Ini::new();
        ini.set(
            PEAZIP_SECTION,
            "extensions",
            Some(PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.peazip_extensions,
            sanitize_peazip_extensions(DEFAULT_PEAZIP_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
        for name in ["arc", "zpaq", "br", "bcm", "lpaq8"] {
            assert!(
                config.peazip_extensions.iter().any(|entry| entry == name),
                "`{name}` is one of the names added with the backends"
            );
        }
    }

    /// Names taken out of a built-in list leave the files already written with them: the
    /// list the app shipped with those names is this app's own — nobody typed it — so it is
    /// read as the list of now, and a name the engine cannot read stops being asked about
    /// on an installation that has been running since before they were taken out.
    #[test]
    fn a_list_holding_the_names_the_engine_cannot_read_loses_them() {
        let mut ini = Ini::new();
        ini.set(
            LIBRE_SECTION,
            "extensions",
            Some(LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.libre_extensions,
            sanitize_libre_extensions(DEFAULT_LIBRE_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
        for name in ["swf", "epub", "qxp", "pm3", "vssm", "uof"] {
            assert!(
                !config.libre_extensions.iter().any(|entry| entry == name),
                "`{name}` is one of the names that left it"
            );
        }
    }

    /// A name this app no longer writes is a key it does not write, and that is all it is: the
    /// setting the line once named is not read from it — this app has no way to know what the
    /// name meant — and the line goes with every other line that is not the app's, which is what
    /// keeps a file from carrying the history of the names it has been written under.
    #[test]
    fn a_name_the_app_no_longer_writes_is_dropped_with_the_line_it_is_on() {
        let mut ini = written_file();
        ini.set(CONFIG_SECTION, "svg_scale", Some("75".to_string()));

        assert!(
            AppConfig::default().differs(&ini),
            "a line the app does not write is one the file is written again for"
        );

        let mut config = AppConfig::default();
        config.apply_ini(&ini);

        assert_eq!(
            config.vector_scale, DEFAULT_VECTOR_SCALE,
            "and the setting it once named is where a fresh installation starts"
        );
    }

    /// An old file, read as the app reads one: the lists it holds are brought up to the ones this
    /// build ships, which is what gives an installation that already exists the formats added
    /// since — while the settings it holds under names this app no longer writes are not read at
    /// all, so those go back to their defaults and the lines are dropped with the write.
    #[test]
    fn an_old_file_keeps_its_lists_and_loses_the_names_this_app_no_longer_writes() {
        let mut ini = written_file();
        ini.set(CONFIG_SECTION, "off_trigger_key", Some("ctrl".to_string()));
        ini.set(CONFIG_SECTION, "avoid_filename", Some("true".to_string()));
        ini.set(CONFIG_SECTION, "svg_scale", Some("75".to_string()));
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_BEFORE_DDS.to_string()),
        );
        ini.set(
            DESIGN_SECTION,
            "extensions",
            Some(DESIGN_EXTENSIONS_BEFORE_AI.to_string()),
        );

        assert!(
            repair_older_lists(&mut ini),
            "the lists are what the repair has to do with a file like this"
        );

        let mut config = AppConfig::default();
        config.apply_ini(&ini);

        for extension in ["dds", "avif", "heic", "heif", "jxl"] {
            assert!(
                config.image_extensions.contains(&extension.to_string()),
                "`{extension}` is in the list of now"
            );
        }
        assert!(config.design_extensions.contains(&"ai".to_string()));

        assert_eq!(
            config.trigger_key, "alt",
            "a name the app does not write says nothing, however clear it looks"
        );
        assert_eq!(config.avoid_mode, DEFAULT_AVOID_MODE);
        assert_eq!(config.vector_scale, DEFAULT_VECTOR_SCALE);

        assert!(
            config.differs(&ini),
            "and the file, holding lines the app does not write, is one to write again"
        );
    }

    /// A file holding everything this app writes, as `save` would write it: what a file on disk
    /// is once it has been through the app, and what the tests below change one key of.
    fn written_file() -> Ini {
        let mut ini = Ini::new();
        for (section, keys) in AppConfig::default().to_ini().get_map_ref() {
            for (key, value) in keys {
                ini.set(section, key, value.clone());
            }
        }

        ini
    }

    /// A setting the file does not have is one the app has to write: the line was deleted, a
    /// whole section was, or the setting is one this build has and the file was written before
    /// it existed. A file that holds everything this build writes is one there is nothing to do.
    #[test]
    fn a_setting_the_file_does_not_have_is_a_reason_to_write_it() {
        let config = AppConfig::default();

        let mut partial = Ini::new();
        partial.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));

        assert!(
            config.differs(&partial),
            "a file holding one setting of fifty does not say what the app is using"
        );
        assert!(
            !config.differs(&written_file()),
            "a file holding every one of them is one to leave alone"
        );
    }

    /// The file is what this app writes and nothing else: a key of the user's own, or a section
    /// of their own, is a difference like any other, and the write it asks for is what drops it.
    /// The one thing this does not touch is a comment, which is not a key and is not compared.
    #[test]
    fn a_key_the_app_does_not_write_is_a_reason_to_write_it() {
        let mut ini = written_file();
        assert!(
            !AppConfig::default().differs(&ini),
            "a file holding what the app writes is one to leave alone"
        );

        ini.set(CONFIG_SECTION, "cat", Some("yes".to_string()));
        assert!(
            AppConfig::default().differs(&ini),
            "a key of the user's own is not a key the app writes"
        );

        let mut own_section = written_file();
        own_section.set("mine", "key", Some("yes".to_string()));
        assert!(
            AppConfig::default().differs(&own_section),
            "and neither is a section of their own"
        );
    }

    /// A value the app cannot read is one it does not keep, and a file holding one is a file to
    /// write: what was written by hand — a tone map that is not one of them, a delay past the
    /// ceiling, a face of a collection past the last one there is — is replaced by the value the
    /// app is actually using, so the file says what the app does rather than what it could not
    /// read.
    #[test]
    fn a_value_the_app_cannot_read_is_written_back_as_the_one_it_uses() {
        let mut ini = written_file();
        ini.set(
            CONFIG_SECTION,
            "hdr_tone_map",
            Some("reinharzzz".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.hdr_tone_map, DEFAULT_HDR_TONE_MAP,
            "a curve that is not one of them leaves the setting where it starts"
        );
        assert!(
            config.differs(&ini),
            "and the file, which names no curve the app reads, is one to write"
        );

        // And the same for a value a setting reduces: a delay past its ceiling is read as the
        // ceiling, so a file asking for more than that is not one the app would write.
        let mut ini = written_file();
        ini.set(
            CONFIG_SECTION,
            "spinner_delay_ms",
            Some((MAX_SPINNER_DELAY_MS * 10).to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(config.spinner_delay_ms, MAX_SPINNER_DELAY_MS);
        assert!(config.differs(&ini));

        // A tick is reduced too, at both ends: below the system's own clock the number
        // asks for a wait nothing can honour, and past a second it says nothing that
        // turning previews off does not say better.
        let mut ini = written_file();
        ini.set(CONFIG_SECTION, "tick_ms", Some("0".to_string()));

        let config = read_file(&mut ini);

        assert_eq!(
            config.tick_ms, MIN_TICK_MS,
            "a tick of nothing is the floor"
        );
        assert!(config.differs(&ini));
    }

    /// A file the app writes is a file it reads back as itself: reading the text `save` puts on
    /// disk gives the same settings, and writing those out again gives the same text. This is
    /// what the whole check rests on — a file that read back as something else would be written
    /// on every read, the watcher would see that write as a change, and the app would spend the
    /// rest of the run rewriting a file it had just written.
    #[test]
    fn a_file_the_app_wrote_is_one_it_reads_back_as_itself() {
        let config = AppConfig {
            theme: TextTheme::Dark,
            avoid_mode: AvoidMode::Details,
            hdr_tone_map: Curve::Aces,
            hdr_exposure: -2.5,
            decode_budget_gb: 0.5,
            office_engine_idle: EngineIdle::Indefinite,
            webview_idle: EngineIdle::Seconds(60),
            pdf_scale: PreviewScale::Percent(25),
            image_background: TransparentBackground::Transparent,
            text_scroll_far_edge_grace_pixels: 12.5,
            office_cache_mb: 1024,
            tick_ms: 47,
            ..Default::default()
        };

        for config in [AppConfig::default(), config] {
            let written = ordered_text(&config.to_ini());

            let mut ini = Ini::new();
            ini.read(written.clone()).expect("a file this app wrote");

            let mut read_back = AppConfig::default();
            read_back.apply_ini(&ini);

            assert_eq!(
                ordered_text(&read_back.to_ini()),
                written,
                "the file is read back as the file it is"
            );
            assert!(
                !read_back.differs(&ini),
                "and it is a file there is nothing to write"
            );
        }
    }

    /// The repair runs on every read, so a file it has already been through has to come out of
    /// it untouched: a file written every time it is read is a file whose mtime moves under the
    /// watcher, and the app would keep reading its own write back for as long as it ran.
    #[test]
    fn a_file_the_repair_has_been_through_is_left_alone() {
        let mut ini = written_file();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_BEFORE_SVG.to_string()),
        );

        assert!(
            repair_older_lists(&mut ini),
            "the file had a list of the app's own from before it grew"
        );

        let once = ordered_text(&ini);

        assert!(
            !repair_older_lists(&mut ini),
            "the file it made is one there is nothing left to do to"
        );
        assert_eq!(
            ordered_text(&ini),
            once,
            "and it is the file it was left as"
        );
    }

    /// The cost of telling this app's own older lists from a user's by their entries, said out
    /// loud: an entry taken out of a list comes back where what is left is exactly a list this
    /// app shipped. `dds` is the entry to take out — the list without it is the one this app
    /// shipped before it was added — while a list with anything else changed is the user's and
    /// is kept, which the test above this one covers.
    #[test]
    fn an_entry_taken_out_of_a_list_can_come_back() {
        let mut ini = Ini::new();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(sanitize_image_extensions(IMAGE_EXTENSIONS_BEFORE_DDS).join(",")),
        );

        let config = read_file(&mut ini);

        assert!(
            config.image_extensions.contains(&"dds".to_string()),
            "the list is read as this app's own older one, and `dds` is put back into it"
        );
    }

    /// The design list's own version of the same: a file holding a list the app shipped
    /// before two more names were added to it — `ai`, and then `cdr` and `procreate` — is
    /// the app's own older list, so those entries reach an installation that already exists
    /// rather than a fresh one only — and a list anyone has edited is kept exactly as it is.
    #[test]
    fn a_list_holding_the_apps_own_design_entries_takes_the_ones_added_to_them() {
        for shipped in [
            DESIGN_EXTENSIONS_BEFORE_AI,
            DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE,
        ] {
            let mut ini = Ini::new();
            ini.set(DESIGN_SECTION, "extensions", Some(shipped.to_string()));

            let config = read_file(&mut ini);

            assert_eq!(
                config.design_extensions,
                sanitize_design_extensions(DEFAULT_DESIGN_EXTENSIONS),
                "the list the app shipped before (`{shipped}`) is read as the list it ships now"
            );
        }

        let edited = format!("dng,{DESIGN_EXTENSIONS_BEFORE_AI}");
        let mut ini = Ini::new();
        ini.set(DESIGN_SECTION, "extensions", Some(edited.clone()));

        let config = read_file(&mut ini);

        assert_eq!(
            config.design_extensions,
            sanitize_design_extensions(&edited),
            "a list with an entry of its own is the user's and is kept as written"
        );
    }

    /// The same, one list's worth of entries later: a file holding the list the app
    /// shipped before the formats Windows has a codec for were added to it is the
    /// app's own older list, so those four reach an installation that already exists
    /// rather than a fresh one only.
    #[test]
    fn a_list_holding_the_apps_own_image_entries_takes_the_codec_formats_added_to_them() {
        let mut ini = Ini::new();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.image_extensions,
            sanitize_image_extensions(DEFAULT_IMAGE_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
        for extension in ["avif", "heic", "heif", "jxl"] {
            assert!(
                config.image_extensions.contains(&extension.to_string()),
                "`{extension}` was added to the built-in list"
            );
        }
    }

    /// The same list one move later, the other way round: a file holding the list the app
    /// shipped while `svg` and `svgz` were entries of it is the app's own older list, so
    /// those two entries are given up — the kind they belong to names them now — rather
    /// than left in a list of pictures they were never pictures of.
    #[test]
    fn a_list_holding_the_apps_own_image_entries_gives_the_documents_up() {
        let mut ini = Ini::new();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_WITH_SVG.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.image_extensions,
            sanitize_image_extensions(DEFAULT_IMAGE_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
        assert!(!config.image_extensions.contains(&"svg".to_string()));
        assert!(!config.image_extensions.contains(&"svgz".to_string()));
    }

    /// And the vector list's own version of it: a file holding the list the app shipped
    /// before the documents were added to it takes them, so an installation that already
    /// exists keeps previewing an `svg` when the image list gives it up.
    #[test]
    fn a_list_holding_the_apps_own_vector_entries_takes_the_documents_added_to_them() {
        let mut ini = Ini::new();
        ini.set(
            VECTOR_SECTION,
            "extensions",
            Some(VECTOR_EXTENSIONS_BEFORE_SVG.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.vector_extensions,
            sanitize_vector_extensions(DEFAULT_VECTOR_EXTENSIONS),
            "the list the app shipped before is read as the list it ships now"
        );
        for extension in ["svg", "svgz"] {
            assert!(
                config.vector_extensions.contains(&extension.to_string()),
                "`{extension}` is a drawing and belongs to the vector list"
            );
        }
    }

    /// And the spellings of an encapsulated PostScript file, which were added to that list
    /// after the documents were: a file holding the list as it stood before them takes them
    /// too, so a `.epsf` or an `.ept` is a drawing an installation already exists previews.
    #[test]
    fn a_list_holding_the_apps_own_vector_entries_takes_the_eps_spellings_added_to_them() {
        let mut ini = Ini::new();
        ini.set(
            VECTOR_SECTION,
            "extensions",
            Some(VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.vector_extensions,
            sanitize_vector_extensions(DEFAULT_VECTOR_EXTENSIONS)
        );
        for extension in ["epsf", "epi", "ept", "ept2", "ept3"] {
            assert!(
                config.vector_extensions.contains(&extension.to_string()),
                "`{extension}` is another spelling of the same drawing"
            );
        }
    }

    /// And the image list's own version: a file holding the list the app shipped before the
    /// codec's AVC still was added to it takes the name, so a `.avci` is a picture an
    /// installation that already exists previews.
    #[test]
    fn a_list_holding_the_apps_own_image_entries_takes_the_avc_still_added_to_them() {
        let mut ini = Ini::new();
        ini.set(
            IMAGE_SECTION,
            "extensions",
            Some(IMAGE_EXTENSIONS_BEFORE_AVCI.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.image_extensions,
            sanitize_image_extensions(DEFAULT_IMAGE_EXTENSIONS)
        );
        assert!(config.image_extensions.contains(&"avci".to_string()));
    }

    /// The same list with one entry of the user's own in it is the user's list, not
    /// the app's: the entries added since are left out of it.
    #[test]
    fn an_image_list_anyone_has_edited_is_kept_as_written() {
        let written = format!("dng,{IMAGE_EXTENSIONS_BEFORE_SVG}");

        let mut ini = Ini::new();
        ini.set(IMAGE_SECTION, "extensions", Some(written.clone()));

        let config = read_file(&mut ini);

        assert_eq!(
            config.image_extensions,
            sanitize_image_extensions(&written),
            "what the file says is what the list is"
        );
        assert!(!config.image_extensions.contains(&"svg".to_string()));
    }

    /// The delay a hover's load is given before the spinner goes up is a setting of its
    /// own, read from its own key in milliseconds: `0` is a delay like any other, and a
    /// number past the ceiling is reduced to it.
    #[test]
    fn the_spinner_delay_is_read_from_its_own_key() {
        assert_eq!(
            AppConfig::default().spinner_delay_ms,
            DEFAULT_SPINNER_DELAY_MS,
            "a load is given a quarter of a second unless the file says otherwise"
        );

        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "spinner_delay_ms", Some("900".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(config.spinner_delay_ms, 900);

        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "spinner_delay_ms", Some("0".to_string()));

        let config = read_file(&mut ini);
        assert_eq!(config.spinner_delay_ms, 0, "`0` is a delay like any other");

        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "spinner_delay_ms",
            Some((MAX_SPINNER_DELAY_MS + 1).to_string()),
        );

        let config = read_file(&mut ini);
        assert_eq!(config.spinner_delay_ms, MAX_SPINNER_DELAY_MS);
    }

    /// The settings section is written under the headings the tray lists its menus under, in
    /// the order the tray lists them, with the keys of a heading in alphabetical order and the
    /// file lists left to the sections below — which is the shape a person reads, rather than
    /// one alphabetical run of every key the app knows.
    #[test]
    fn the_settings_are_written_under_the_headings_the_tray_lists_them_under() {
        let mut ini = Ini::new();
        // Set out of order, and out of the order the headings run in, so what is tested is the
        // writer's order rather than the order they were handed over in.
        ini.set(CONFIG_SECTION, "video_volume", Some("50".to_string()));
        ini.set(CONFIG_SECTION, "avoid_mode", Some("details".to_string()));
        ini.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));
        ini.set(
            CONFIG_SECTION,
            "image_background",
            Some("white".to_string()),
        );
        ini.set(IMAGE_SECTION, "extensions", Some("png,jpg".to_string()));

        assert_eq!(
            ordered_text(&ini),
            "\
[settings]
; General
preview_enabled=true

; Placement
avoid_mode=details

; Background
image_background=white

; Volume
video_volume=50

[image]
extensions=png,jpg
"
        );
    }

    /// The engine settings are written under a heading of their own, below the one the caches
    /// and the budget are under, which is where the tray lists them.
    #[test]
    fn the_engine_settings_are_written_under_a_heading_of_their_own() {
        let written = ordered_text(&AppConfig::default().to_ini());

        let performance = written
            .find("; Performance\n")
            .expect("the settings are grouped under a heading for what the app costs");
        let engine = written
            .find("; Engine\n")
            .expect("and under one for the engines themselves");
        let advanced = written
            .find("; Advanced\n")
            .expect("and under one for the settings the tray has no item for");

        assert!(
            performance < engine,
            "the caches are listed above the engines"
        );
        assert!(engine < advanced, "and the engines above the rest");

        let under_engine = &written[engine..advanced];
        for key in [
            "libreoffice_idle",
            "office_engine",
            "office_engine_idle",
            "webview_idle",
        ] {
            assert!(under_engine.contains(key), "`{key}` is under `Engine`");
        }
        for key in ["confirm_file_type", "decode_budget_gb", "image_cache_mb"] {
            assert!(!under_engine.contains(key), "`{key}` is not under `Engine`");
        }

        // A file this build wrote is a file there is nothing to write again, which is what keeps
        // the watcher from writing the file it has just read back.
        assert!(
            !headings_are_old(&written),
            "a file this build wrote needs nothing done to it"
        );
    }

    /// A file grouped the way an older build grouped it is written again under the headings of
    /// this one. A heading is a comment, so what such a file holds is read exactly as any other
    /// file is: the grouping is the one thing about it that is not what this build writes, and
    /// the one thing that a write puts right.
    #[test]
    fn a_file_written_before_the_settings_were_regrouped_is_written_again() {
        let written = ordered_text(&AppConfig::default().to_ini());

        // The same file as an older build wrote it: the engines had no heading of their own, so
        // the settings they are named by sat among the caches and the budget.
        let older = written.replace("; Engine\n", "");
        assert!(
            headings_are_old(&older),
            "a file with no heading for the engines is one to write again"
        );

        // A file with every heading, listed in the order the build before this one listed them —
        // the engines above the caches — is one to write again as well: the menus were
        // rearranged, and the headings say so.
        let (head, rest) = written
            .split_once("; Performance\n")
            .expect("a heading for the caches");
        let (performance, rest) = rest
            .split_once("; Engine\n")
            .expect("a heading for the engines");
        let (engine, tail) = rest
            .split_once("; Advanced\n")
            .expect("a heading for the rest");
        let rearranged =
            format!("{head}; Engine\n{engine}; Performance\n{performance}; Advanced\n{tail}");
        assert!(
            headings_are_old(&rearranged),
            "a file listing the headings in another order is one to write again"
        );

        // An editor that saves the file with the other line ending leaves one there is nothing
        // to write either: the heading is still the line it was, and what an editor adds to the
        // end of it is not the app's business.
        assert!(
            !headings_are_old(&written.replace('\n', "\r\n")),
            "a heading is not read by the line ending it happens to have"
        );

        // And the keys are read the same either way, which is why the grouping can be put right
        // without anything being migrated: what moved is the comment, not the setting.
        let mut read_back = Ini::new();
        assert!(read_back.read(older).is_ok());
        assert_eq!(
            read_back.get(CONFIG_SECTION, "office_engine"),
            Some("microsoft_office".to_string())
        );
    }

    /// Every setting the app writes is a setting the heading table names. The table is what the
    /// headings of a file are written from, and it is what the repair reads to tell whether a
    /// file is missing a setting, so a key `save` writes that the table does not name would be
    /// written under `; Ungrouped` — and its absence from a file would never be noticed.
    #[test]
    fn every_setting_the_app_writes_is_one_the_table_names() {
        let written = AppConfig::default().to_ini();
        let keys = written
            .get_map_ref()
            .get(CONFIG_SECTION)
            .expect("a settings section");

        for key in keys.keys() {
            assert!(
                SETTING_GROUPS
                    .iter()
                    .any(|(_, group)| group.contains(&key.as_str())),
                "`{key}` is written by `save` and named by no heading"
            );
        }

        assert_eq!(
            keys.len(),
            SETTING_GROUPS
                .iter()
                .map(|(_, group)| group.len())
                .sum::<usize>(),
            "the table names as many settings as `save` writes"
        );
    }

    /// A heading is a comment, so the file the app writes is one it can read back: the reader
    /// steps over the headings and the blank lines they are kept apart by, and every value
    /// comes back as it was written.
    #[test]
    fn a_file_written_under_the_headings_reads_back_as_it_was() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "avoid_mode", Some("details".to_string()));
        ini.set(CONFIG_SECTION, "hover_delay_ms", Some("200".to_string()));
        ini.set(VECTOR_SECTION, "extensions", Some("svg,svgz".to_string()));

        let mut read_back = Ini::new();
        read_back
            .read(ordered_text(&ini))
            .expect("a file this app wrote is one it can read");

        assert_eq!(
            read_back.get(CONFIG_SECTION, "avoid_mode"),
            Some("details".to_string())
        );
        assert_eq!(
            read_back.get(CONFIG_SECTION, "hover_delay_ms"),
            Some("200".to_string())
        );
        assert_eq!(
            read_back.get(VECTOR_SECTION, "extensions"),
            Some("svg,svgz".to_string())
        );
    }

    /// A setting the table has not been told about is written last, under a heading of its
    /// own: a key added to `save` and forgotten there lands at the bottom of the file where it
    /// can be seen, rather than inside a heading it has nothing to do with or nowhere at all.
    #[test]
    fn a_setting_the_table_does_not_know_is_written_under_its_own_heading() {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "something_new", Some("1".to_string()));
        ini.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));

        assert_eq!(
            ordered_text(&ini),
            "\
[settings]
; General
preview_enabled=true

; Ungrouped
something_new=1
"
        );
    }

    /// One setting belongs to one heading. A key listed under two of them is written under
    /// both, and since what is read back is the key rather than the heading, the second one is
    /// the one the setting takes.
    #[test]
    fn every_setting_belongs_to_one_heading() {
        let mut seen: Vec<&str> = Vec::new();

        for (heading, keys) in SETTING_GROUPS {
            assert!(
                !keys.is_empty(),
                "the heading `{heading}` lists no settings"
            );

            for key in *keys {
                assert!(
                    !seen.contains(key),
                    "`{key}` is listed under more than one heading"
                );
                seen.push(key);
            }
        }
    }

    /// The defaults the tray marks are the ones the app starts at: a picture, a drawing and a
    /// design document are drawn over the squares, a specimen and a texture over a white page,
    /// and the placement, the two delays and the volume start where the menu says.
    #[test]
    fn the_defaults_the_tray_marks_are_the_ones_the_app_starts_at() {
        let config = AppConfig::default();

        assert_eq!(config.image_background, DEFAULT_IMAGE_BACKGROUND);
        assert_eq!(config.vector_background, DEFAULT_VECTOR_BACKGROUND);
        assert_eq!(config.design_background, DEFAULT_DESIGN_BACKGROUND);
        assert_eq!(config.image_background, TransparentBackground::Checkerboard);

        assert_eq!(config.font_background, DEFAULT_FONT_BACKGROUND);
        assert_eq!(config.font_background, TransparentBackground::White);
        assert_eq!(config.dds_background, DEFAULT_DDS_BACKGROUND);
        assert_eq!(config.dds_background, TransparentBackground::White);

        assert_eq!(config.avoid_mode, DEFAULT_AVOID_MODE);
        assert_eq!(config.avoid_mode, AvoidMode::Filename);
        assert_eq!(config.follow_cursor, DEFAULT_FOLLOW_CURSOR);
        assert!(
            !config.follow_cursor,
            "a preview is placed at its best position"
        );

        assert_eq!(config.hover_delay_ms, DEFAULT_HOVER_DELAY_MS);
        assert_eq!(config.hover_delay_ms, 0);
        assert_eq!(
            config.same_file_rehover_delay_ms,
            DEFAULT_SAME_FILE_REHOVER_DELAY_MS
        );
        assert_eq!(config.same_file_rehover_delay_ms, 200);
        assert_eq!(config.tick_ms, DEFAULT_TICK_MS);
        assert_eq!(config.tick_ms, 15, "the loop looks once a system tick");
        assert_eq!(config.video_volume, DEFAULT_VIDEO_VOLUME);
        assert_eq!(config.video_volume, 0, "a hover never makes a sound");
    }

    /// A texture is offered two backdrops of the four, which is what `dds_image` is written
    /// against: the two that show what stands behind a preview are read as the white page the
    /// setting starts at, so a file that still names one is not left holding a value the menu
    /// beside it has no item for.
    #[test]
    fn a_textures_backdrop_is_one_of_the_two_it_is_offered() {
        for kept in [TransparentBackground::Black, TransparentBackground::White] {
            assert_eq!(sanitize_dds_background(kept), kept, "`{kept:?}` is offered");
        }

        for dropped in [
            TransparentBackground::Transparent,
            TransparentBackground::Checkerboard,
        ] {
            assert_eq!(
                sanitize_dds_background(dropped),
                DEFAULT_DDS_BACKGROUND,
                "`{dropped:?}` is not a backdrop a texture is offered"
            );
        }

        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "dds_background",
            Some("checkerboard".to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(config.dds_background, DEFAULT_DDS_BACKGROUND);
    }
}
