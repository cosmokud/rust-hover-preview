//! What every setting starts at, and the reduction a hand-edited `config.ini` is read
//! through: a bound, a default, or a clamp. Everything here answers "what is this setting
//! allowed to be" rather than "what does it hold", so it is the half of the configuration
//! that knows nothing about the file it is written to.

use std::fs;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use crate::readers::tone_map::Curve;

use super::setting_types::{OfficeEngine, VideoEngine};

pub(super) const CONFIG_SECTION: &str = "settings";
pub const DEFAULT_WEBP_PLAYBACK_FPS: u32 = 90;
pub const MAX_WEBP_PLAYBACK_FPS: u32 = 90;
/// The volume a video is played at unless the file says otherwise: silent, so a hover
/// never makes a sound the pointer did not ask for.
pub const DEFAULT_VIDEO_VOLUME: u32 = 0;
/// The volume a sound file is played at unless the file says otherwise.
///
/// It is not the video's zero, and the difference is deliberate: a video is hovered for its
/// picture, and its soundtrack is as likely to be a distraction as anything — while a sound
/// file *is* the sound, and a preview of one at zero is a card with nothing behind it. Quiet
/// rather than loud, so a pointer crossing a folder of music is a few seconds of something
/// half-heard rather than a jukebox.
pub const DEFAULT_AUDIO_VOLUME: u32 = 10;
/// Whether a sound's measured loudness is brought to one level before it is played, so that a
/// folder of files is heard at one level rather than at each file's own.
///
/// It is on where the app starts, and FFmpeg is what makes it possible at all: the loudness is
/// measured by FFmpeg's own scanner and the gain is applied by FFmpeg's own player, where the
/// engine Windows has can only quieten a file — a level is attenuation there, and full volume
/// is as loud as it goes (see `codecs::normalize_available`).
pub const DEFAULT_NORMALIZE_VOLUME: bool = true;
/// Whether a video's soundtrack is brought to the same level before it is played, on the same
/// terms and by the same measurement as a sound file's own (see above).
///
/// It is on where the app starts, on the same terms as the sound's own: a folder of films is
/// heard at one level rather than at each film's own, and the cost of the measurement — a
/// decode of the film — is the same one the sound's already pays (see above).
pub const DEFAULT_NORMALIZE_VIDEO_VOLUME: bool = true;
/// Whether a sound's previewed level is kept between hovers, or whether the level the setting
/// names is the level every sound is previewed at.
///
/// It is on where the app starts: a level that is kept is what a person who turned it up once
/// meant for the rest of the folder, and a level that is not is a quiet app that plays every
/// file the same way until it is asked otherwise (see `remember_audio_volume`).
pub const DEFAULT_REMEMBER_AUDIO_VOLUME: bool = true;
/// And the video's own, which is the same question asked about a soundtrack and is on for the
/// same reason the sound's own is: a level that is kept is what a person who turned it up once
/// meant for the rest of the folder (see `remember_video_volume`).
pub const DEFAULT_REMEMBER_VIDEO_VOLUME: bool = true;
/// Whether a video is decoded on the graphics card, which is the `Video` toggle under
/// `Performance → Hardware Acceleration` in the tray.
///
/// It is on where the app starts, because the answer is a machine's rather than a person's: a
/// film previewed in software decoding costs a core for as long as the hover lasts, and every
/// machine this app runs on has a Direct3D 11 device FFmpeg can decode on. It is a setting
/// because the one machine that cannot is a machine where the fallback is the whole answer — and
/// FFmpeg falls back by itself rather than being asked to (see `video_hw_accel_device`).
pub const DEFAULT_VIDEO_HW_ACCEL: bool = true;
/// The levels either volume is offered at: silence, the one step above it, and the decades
/// between — smallest first here, and listed the other way round in the tray, loudest first.
///
/// One table for both menus rather than one apiece, because the question a user asks of either
/// is the same — how loud is this — and the two are the same setting applied to two kinds of
/// preview. `1%` is in the list because a hover that makes a sound at all is sometimes wanted
/// without its being audible across a room, and the list stops at `100%` because past it is an
/// amplification this app has no business applying to someone else's file.
pub const VOLUME_CHOICES: [u32; 10] = [0, 1, 5, 10, 20, 35, 50, 65, 80, 100];

/// A volume reduced to what the app allows.
///
/// A hand-edited `config.ini` is the only way past the menu's own list, and what it can ask for
/// is an amplification (`volume=10.0` on FFmpeg's filter is ten times the file) or a number the
/// players read differently from a percentage. What is kept is anything from silence to the
/// whole of it, so a level the menu does not offer — `37` — is honoured as written rather than
/// rounded to the nearest, the same way a delay the menu does not offer is.
pub fn sanitize_volume(volume: u32) -> u32 {
    volume.min(MAX_VOLUME)
}

/// The loudest a preview may play at.
pub const MAX_VOLUME: u32 = 100;
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
/// The share of the display a sound's card is drawn at unless the
/// configuration says otherwise.
///
/// A sound's card holds no bitmap the file asks to be drawn at a size — what it
/// holds is the sound, a name, a seek bar and the facts about the file, laid
/// out over whatever room it is given — so the share is of the display rather
/// than of the file, as it is for the drawings and documents. It starts at a
/// tenth of the display: the size the card was measured at across the files it
/// was tried on, which is the size a sound's card wants to be.
pub const DEFAULT_AUDIO_SCALE_PERCENT: u32 = 10;
pub const MIN_AUDIO_SCALE_PERCENT: u32 = 1;
pub const MAX_AUDIO_SCALE_PERCENT: u32 = 100;
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
/// Whether a page of HTML is drawn by the browser engine rather than shown as its markup.
///
/// It is off where the app starts: a `.htm` and a `.html` are text files, and what the text
/// lists claim is a page of text drawn at a fixed font size. The setting is what a file of
/// those two names is previewed by instead — the same page the browser would show, in the
/// engine's own window — and it is answered by whether the runtime is on the machine, so a
/// machine without it keeps the text preview however the setting is written (see
/// `webview_preview::draws`).
pub const DEFAULT_RENDER_HTML: bool = false;
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
/// Sixty-four megabytes by default, which is one full-screen frame at 4K: a hit skips a
/// full-resolution decode and its resample, which is the most expensive thing a hover does,
/// and what is held is the preview-sized frame in BGRA — a picture filling a large display
/// is about thirty-three megabytes of it. A budget below that holds nothing at all for such
/// a picture, because a frame larger than the whole budget is not trimmed to fit but left
/// out of the cache, so every hover of it paid the decode again; sixty-four holds that frame
/// and the one after it.
pub const DEFAULT_IMAGE_CACHE_MB: u32 = 64;
pub const MAX_IMAGE_CACHE_MB: u32 = 2048;
/// The folder a document's page is kept in may hold, in megabytes. What a page costs is the
/// size of the file it is kept as — a PDF the engine that drew it exported, a slide's PNG, a
/// workbook's picture — and what is kept is the working set of the documents a user previews,
/// least recently used first: a page that has been given up is drawn again the next time the
/// document is hovered, and drawing one costs an Office start or a conversion rather than a
/// read, which is what makes holding them worth a folder of their own (see `document_cache`).
///
/// Two hundred and fifty-six megabytes by default, because the folder carries pages of three
/// different weights at once: a text page is a couple of hundred kilobytes, an ordinary
/// document a couple of megabytes, and the illustrated documents the engines that convert a
/// whole file leave — a LibreOffice drawing, a comic, and above all a scanned book — five to
/// thirty megabytes apiece, with a book of plates reaching a hundred. A budget of a hundred
/// and twenty-eight was one fat page away from full, and a page the trim gives up is the
/// Office start or the conversion that drew it paid again to bring it back. What holding
/// twice as much costs is disk rather than memory: nothing is read from the volume while it
/// is not in use.
pub const DEFAULT_DOCUMENT_CACHE_MB: u32 = 256;
pub const MAX_DOCUMENT_CACHE_MB: u32 = 2048;
/// The folder a developed picture is kept in may hold, in megabytes. What a picture costs is the
/// size of the PNG the image engine wrote — a few hundred kilobytes to a few megabytes at the
/// size a preview is shown — and what is kept is the pictures a user hovers, least recently used
/// first, so that a second hover and the next run are a read rather than a conversion.
///
/// Five hundred and twelve megabytes by default, because this budget covers the one thing the
/// image cache beside it cannot: a picture the engine developed for a preview larger than
/// `image_cache_mb` is not held in memory at all, so without a page here every hover of it pays
/// the launch again — a fraction of a second to a second, per hover, for a file that never
/// changed. It is disk rather than memory: nothing is read from the folder while it is not in
/// use. See `document_cache`.
pub const DEFAULT_IMAGE_DISK_CACHE_MB: u32 = 512;
pub const MAX_IMAGE_DISK_CACHE_MB: u32 = 2048;
/// The folder a film's extracted subtitle files are kept in may hold, in megabytes. What it
/// holds are the small files this app copied a film's own subtitle tracks into, with the fonts
/// that came out of the container dumped beside them (see `preview_window::subtitle_files`),
/// and what the folder buys is the whole point of that extraction: a hover whose subtitles are
/// embedded in the film otherwise streams the entire container before the first frame is
/// drawn — measured on this machine, a cold 1.4 GB MKV read 1 423 MB and took 14 904 ms to
/// show a frame with the embedded filter, against 492 ms and 41 MB for the same film drawn
/// from a 30 KB extracted `.ass`.
///
/// A hundred and twenty-eight megabytes by default, which is thousands of films: the measured
/// extraction of a 1.4 GB anime episode is a 30 KB `.ass`, and even a film whose tracks are
/// PGS bitmaps writes a few megabytes. It is a disk cache like the two beside it — nothing is
/// read from the folder while it is not in use — and the budget is what bounds a library's
/// worth of films hovered across months rather than anything a single film needs.
pub const DEFAULT_GENERAL_DISK_CACHE_MB: u32 = 128;
pub const MAX_GENERAL_DISK_CACHE_MB: u32 = 2048;
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
/// Which engine plays a video unless the configuration says otherwise: the best this machine has,
/// which is `VideoEngine::Hybrid` where FFmpeg is installed — the media engine for a small film
/// and FFmpeg's player for a large one — and the media engine where it is not.
pub const DEFAULT_VIDEO_ENGINE: VideoEngine = VideoEngine::Best;
/// Whether an explicitly chosen engine that cannot play a file falls through to the others:
/// on, so a film no media-engine decoder reaches is still played rather than shown as nothing.
pub const DEFAULT_VIDEO_ENGINE_FALLBACK: bool = true;
/// The resolution above which `VideoEngine::Hybrid` leaves a film to FFmpeg's player: 3.2 million
/// pixels, which is between 1080p and QHD, so a 1440p film is handed over and a 1080p one is drawn
/// here. It is total pixels and not an axis, because what it stands for is how much there is to
/// draw rather than how wide the picture is (see `video_hw::resolve_video_engine`).
pub const VIDEO_FFMPEG_ABOVE_PIXELS: u64 = 3_200_000;
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

/// The page cache's size in megabytes, which is the budget of the `Document` kind.
///
/// `0` is a cache that holds nothing rather than a preview tier that is switched off: a
/// document is drawn for the hover that asks for it either way, and the only question a size
/// answers is how much of what was drawn is kept between hovers.
pub fn sanitize_document_cache_mb(value: u32) -> u32 {
    value.min(MAX_DOCUMENT_CACHE_MB)
}

/// The budget of the folder a developed picture is kept in, which is the image converter's own
/// cache rather than the pages the document engines write.
///
/// `0` is a cache that holds nothing between hovers, the same as the one above: a picture is
/// developed for the hover that asks for it either way, and the only question a size answers is
/// whether the hover after it reads one back or pays for it again.
pub fn sanitize_image_disk_cache_mb(value: u32) -> u32 {
    value.min(MAX_IMAGE_DISK_CACHE_MB)
}

/// The budget of the folder a film's extracted subtitle files are kept in, which is the derived
/// subtitle files' own cache rather than the pages the document engines write or the pictures
/// an image converter develops.
///
/// `0` is a cache that keeps nothing, the same as the two above: at that size the whole-film
/// read an extraction pays for would be given straight back, so nothing is extracted at all
/// and a film whose subtitles are embedded is drawn without them (see `subtitle_files`).
pub fn sanitize_general_disk_cache_mb(value: u32) -> u32 {
    value.min(MAX_GENERAL_DISK_CACHE_MB)
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

pub(super) fn parse_text_font_scale(value: &str) -> Option<u32> {
    let normalized = value.trim().to_ascii_lowercase();
    let normalized = normalized.trim_end_matches('%').trim();
    normalized
        .parse::<u32>()
        .ok()
        .map(sanitize_text_font_scale_percent)
}

pub(super) fn sanitize_preview_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_PREVIEW_SCALE_PERCENT
    } else {
        value.clamp(MIN_PREVIEW_SCALE_PERCENT, MAX_PREVIEW_SCALE_PERCENT)
    }
}

/// The share of the display a sound's card is laid out over, which a
/// hand-edited `config.ini` is read through: a share of nothing names no share
/// at all, so it is the tenth the setting starts at, and a share past the whole
/// display is the whole display — a card cannot be given more room than there
/// is. Anything between is honored as written, the same way a delay the menu
/// does not offer is.
pub(super) fn sanitize_audio_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_AUDIO_SCALE_PERCENT
    } else {
        value.clamp(MIN_AUDIO_SCALE_PERCENT, MAX_AUDIO_SCALE_PERCENT)
    }
}
