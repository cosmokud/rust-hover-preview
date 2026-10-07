//! The shapes a setting's value can take, and the word each one is written as in
//! `config.ini`. Every enum here carries its own spelling both ways, so a setting of this
//! kind is written and read without the file's shape knowing anything about it.

use std::collections::HashSet;
use std::sync::Mutex;
use std::time::Duration;

use once_cell::sync::Lazy;

use crate::config::theme_files;
use crate::formats::codecs;

use super::app_config::AppConfig;
use super::defaults::{
    sanitize_audio_scale_percent, sanitize_preview_scale_percent, DEFAULT_ANIMATED_SCALE_PERCENT,
    DEFAULT_AUDIO_SCALE_PERCENT, DEFAULT_FONT_SCALE_PERCENT, DEFAULT_PREVIEW_SCALE_PERCENT,
    DEFAULT_VIDEO_SCALE_PERCENT, MAX_AFK_TIMER_SECS, MAX_OFFICE_ENGINE_IDLE_SECS,
};

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

    /// The scale a sound's `config.ini` value names, or `None` for one that is
    /// neither a percentage nor a word for the fit.
    ///
    /// The same words as a picture's scale are read, through the sound's own
    /// bounds rather than the picture's: the setting is a share of the display
    /// rather than of the file, so a hand-edited percentage past the whole
    /// display is the whole display, and a `0` — a share of nothing — is the
    /// share a fresh installation starts at. A value that names nothing is
    /// `None`, which leaves the setting where the configuration already holds
    /// it rather than guessing.
    pub(crate) fn from_audio_str(value: &str) -> Option<Self> {
        let normalized = value.trim().to_ascii_lowercase();
        let normalized = normalized.trim_end_matches('%').trim();

        match normalized {
            "fit" | "fit to screen" | "fit-to-screen" | "fit_to_screen" | "fittoscreen" => {
                Some(Self::FitToScreen)
            }
            _ => normalized
                .parse::<u32>()
                .ok()
                .map(|percent| Self::Percent(sanitize_audio_scale_percent(percent))),
        }
    }
}

/// What a PDF page — the `Ebook` kind — is drawn at unless the configuration says
/// otherwise: the whole of the room the display has for it, which is the answer that asks
/// for nothing in particular — the page at its own size where the display can hold it,
/// reduced only where it cannot.
pub const DEFAULT_EBOOK_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// The same for a page of the `Document` kind, drawn by the application that owns the
/// format or by the render engine beside it.
pub const DEFAULT_DOCUMENT_SCALE: PreviewScale = PreviewScale::FitToScreen;
/// What a picture, a video and an animated picture are drawn at unless the configuration
/// says otherwise: the size each file asks for, at the share its own setting names.
pub const DEFAULT_PREVIEW_SCALE: PreviewScale =
    PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT);
pub const DEFAULT_VIDEO_SCALE: PreviewScale = PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT);
pub const DEFAULT_ANIMATED_SCALE: PreviewScale =
    PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT);
/// What a sound's card is drawn at unless the configuration says otherwise: a
/// share of the room the display has for it, at the tenth the card was measured
/// at across the files it was tried on. A card holds no bitmap of the file's own
/// to take a share of — what it holds is laid out over the room the setting
/// names, its height kept by the font it is set in — so the room is the whole
/// question, as it is for the drawings and documents beside it.
pub const DEFAULT_AUDIO_SCALE: PreviewScale = PreviewScale::Percent(DEFAULT_AUDIO_SCALE_PERCENT);
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
/// The same for a font specimen, at the share rather than the whole of the room: a specimen
/// is a page of text rather than a document to be studied, and half the display holds the
/// pangram at a size that can be read at a glance.
pub const DEFAULT_FONT_SCALE: PreviewScale = PreviewScale::Percent(DEFAULT_FONT_SCALE_PERCENT);
/// The share of the display a text page is given before it is measured: the whole of it, like
/// the drawing, the document and the page beside it. A text page has no size of its own to take
/// a share of, so the setting is the room it may be measured in.
pub const DEFAULT_TEXT_SCALE: PreviewScale = PreviewScale::FitToScreen;

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

/// Which engine plays a video, as the tray's `Engine -> Select Engine -> Video` lists it.
///
/// `Best` is the machine's own answer and the default, and it is `Hybrid` where FFmpeg is
/// installed and `Native` where it is not (see `video_hw::resolve_video_engine`). The other three
/// name one engine each, or — `Hybrid` — a rule between two of them, whatever the machine would
/// have preferred.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum VideoEngine {
    /// The best engine this machine has for the file.
    Best,
    /// The media engine Windows has, drawn into this window's own frame.
    Native,
    /// FFmpeg's `ffplay`, in a window of its own.
    Ffmpeg,
    /// The media engine for a film small enough to draw here and FFmpeg's player for a larger one:
    /// drawing a big film through this window costs more than handing it to a player, and drawing
    /// a small one costs less. What divides the two is `VIDEO_FFMPEG_ABOVE_PIXELS` total pixels,
    /// and a film nobody measured is handed over as well, because it is the probe that would have
    /// weighed it that could not read it (see `video_hw::resolve_video_engine`).
    Hybrid,
}

impl VideoEngine {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Best => "best",
            Self::Native => "native",
            Self::Ffmpeg => "ffmpeg",
            Self::Hybrid => "hybrid",
        }
    }

    /// Whether this machine has the engine this choice names, which is what a tray row is greyed on
    /// and what a choice is put back to `Best` for at a start. `Best` is always here — it is the
    /// app's own answer rather than a named player — and `Native` is the engine Windows ships;
    /// both `Ffmpeg` and `Hybrid` need `ffplay` (see `video_hw::resolve_video_engine`).
    pub(crate) fn installed(self) -> bool {
        match self {
            Self::Best | Self::Native => true,
            Self::Ffmpeg | Self::Hybrid => codecs::ffplay_available(),
        }
    }

    /// The engine a `config.ini` value names, or `None` for one that names no player: a value
    /// the app cannot read leaves the setting where it is (see `OfficeEngine::from_str`).
    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "best" | "auto" | "default" => Some(Self::Best),
            "native" | "media_engine" | "media engine" | "media" | "windows" => Some(Self::Native),
            "ffmpeg" | "ffplay" | "ff" => Some(Self::Ffmpeg),
            "hybrid" | "native_ffmpeg" | "native_above" | "native then ffmpeg" => {
                Some(Self::Hybrid)
            }
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
pub(super) fn sanitize_afk_timer_secs(seconds: u64) -> u64 {
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

    pub(super) fn from_str(value: &str) -> Option<Self> {
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

/// What a page of HTML is drawn over: white, which is the page a page is written on rather
/// than a backdrop behind one.
///
/// A document the browser is handed is a page already — it brings its own markup, its own
/// stylesheet and its own idea of what a heading looks like — so what is behind it is
/// something to read it against, and white is where the setting starts. The two backdrops
/// that show what stands behind a preview are the two a page has no use for, so rather than
/// offering transparency the menu offers the page black, the page white and the squares the
/// picture half keeps.
pub const DEFAULT_HTML_BACKGROUND: TransparentBackground = TransparentBackground::White;

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

/// The backdrop a page is drawn over, as the setting keeps it: the two pages a preview can
/// be read against and the squares beside them, and the default for the one that is not
/// offered — a `transparent` a hand-edited `config.ini` still holds from when a page was
/// drawn over the vector half's setting, and which is not something a page is offered.
pub fn sanitize_html_background(background: TransparentBackground) -> TransparentBackground {
    match background {
        TransparentBackground::Black
        | TransparentBackground::White
        | TransparentBackground::Checkerboard => background,
        TransparentBackground::Transparent => DEFAULT_HTML_BACKGROUND,
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

    pub(super) fn from_str(value: &str) -> Option<Self> {
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
    pub(super) fn resolve(value: &str) -> Option<Self> {
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

    pub(super) fn from_str(value: &str) -> Option<Self> {
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
///
/// The last four kinds are not rows of that submenu, because what they name is a rendering
/// path rather than something a user asks for: a document an installed engine draws, a
/// picture an installed converter develops, an archive an installed listing engine reads, and a
/// book an installed ebook engine converts. Each is the same preview as the kind it belongs to — a
/// page, a picture, a page of contents, a book — and is switched by that kind's gate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PreviewType {
    Images,
    Videos,
    /// Sound files: what this app plays rather than draws, previewed as a card of what the
    /// file holds with the sound behind it — played by the media engine Windows has where its
    /// decoders reach the format and by FFmpeg's player where they do not, both of which are
    /// the machine's answer rather than the user's. See `audio_formats` for what is listed,
    /// `audio_track` for the probe, and `audio_preview` for the card.
    Audio,
    Text,
    /// PDF pages: the one kind this app draws as a page with a reader of its own rather than
    /// by handing the file to an engine, a decoder or the drawing layer. See `pdf_preview`.
    Ebook,
    Archives,
    /// Documents previewed as a page — the words of an Office document, and the drawings and
    /// older formats an installed render engine draws. Which of the two draws a file is
    /// answered by the machine rather than by the user; see `office_formats::page_engine`.
    Document,
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
    ///
    /// It is the `Document` kind's second half rather than a switch of its own: what a user
    /// turns off is documents.
    Libre,
    /// Pictures this app hands to an installed ImageMagick rather than decoding — the camera
    /// raw formats above all, which nothing else on a Windows machine opens at all. What
    /// comes back is a PNG, and it is drawn as the picture it is: the picture scale and the
    /// picture backdrop, held in the picture cache. See `magick_formats` for what is listed
    /// and `imagemagick_render` for how one is converted.
    ///
    /// It is the `Images` kind's second half rather than a switch of its own: what a user
    /// turns off is pictures.
    Magick,
    /// Archives this app hands to an installed PeaZip rather than reading — the cabinet files,
    /// isos, disk images, installers and single-stream compressors no reader here has. What
    /// comes back is the archive's own table of contents, read into the shape every other
    /// listing is and drawn as the same page. See `peazip_formats` for what is listed and
    /// `peazip_render` for how one is listed.
    ///
    /// It is the `Archives` kind's second half rather than a switch of its own: what a user
    /// turns off is archives.
    Peazip,
    /// Books this app hands to an installed Calibre rather than reading — the Kindle and
    /// Mobipocket formats above all, the open EPUB, the FictionBook and the scanned book, none of
    /// which the PDF reader here opens. What comes back is a PDF of the book, drawn as a PDF page
    /// is: at the book kind's scale, over the book kind's backdrop, held in the page cache. See
    /// `calibre_formats` for what is listed and `calibre_render` for how one is converted.
    ///
    /// It is the `Ebook` kind's second half rather than a kind of its own: what a user turns off
    /// is books, and what is on screen for either is a page of one. The switch is a row of its own
    /// all the same — the way the picture a converter develops and the page a render engine draws
    /// have theirs — because an engine a user did not install is a setting worth being able to
    /// reach, and because `[calibre]` and the PDF reader answer for different files entirely.
    Calibre,
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
    ///
    /// Four of the kinds share a gate with the kind they belong to — a document an engine
    /// drew with the documents, a picture a converter developed with the pictures, an archive
    /// a listing engine read with the archives, and a book an ebook engine converted with the
    /// books — so there is one switch for each pair rather than one for the reader and one for
    /// the engine.
    pub fn enabled_in(self, config: &AppConfig) -> bool {
        match self {
            Self::Images | Self::Magick => config.image_preview_enabled,
            Self::Videos => config.video_preview_enabled,
            Self::Audio => config.audio_preview_enabled,
            Self::Text => config.text_preview_enabled,
            Self::Ebook | Self::Calibre => config.ebook_preview_enabled,
            Self::Archives | Self::Peazip => config.archive_preview_enabled,
            Self::Document | Self::Libre => config.document_preview_enabled,
            Self::Fonts => config.font_preview_enabled,
            Self::Design => config.design_preview_enabled,
            Self::Vector => config.vector_preview_enabled,
        }
    }

    /// Switch this kind of preview on or off, which for the four engine-drawn kinds is the
    /// switch of the kind they belong to — see `enabled_in`.
    pub fn set_enabled_in(self, config: &mut AppConfig, enabled: bool) {
        match self {
            Self::Images | Self::Magick => config.image_preview_enabled = enabled,
            Self::Videos => config.video_preview_enabled = enabled,
            Self::Audio => config.audio_preview_enabled = enabled,
            Self::Text => config.text_preview_enabled = enabled,
            Self::Ebook | Self::Calibre => config.ebook_preview_enabled = enabled,
            Self::Archives | Self::Peazip => config.archive_preview_enabled = enabled,
            Self::Document | Self::Libre => config.document_preview_enabled = enabled,
            Self::Fonts => config.font_preview_enabled = enabled,
            Self::Design => config.design_preview_enabled = enabled,
            Self::Vector => config.vector_preview_enabled = enabled,
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

    pub(super) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "disable" | "disables" | "off" | "hold to disable" => Some(Self::Disable),
            "enable" | "enables" | "on" | "hold to enable" => Some(Self::Enable),
            _ => None,
        }
    }
}

/// Where in a file a sound starts playing.
///
/// A sound is the one preview that is heard rather than looked at, and a file being listened
/// to is as often a file being listened to *again* as one met for the first time — so where a
/// hover drops the needle is a question of its own, apart from the kind's own switch and from
/// how loud it plays. The four answers are where a pointer crossing a folder of music should
/// land in it: where this app last left it, at the beginning, in the middle, or anywhere at
/// all.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AudioSeek {
    /// Where the sound was left the last time it was hovered.
    ///
    /// The one answer that is a memory rather than a rule: what it names is a position this
    /// app wrote down for the file, held in memory for the run and under the temp folder
    /// across runs (see `audio_seek`). A file nothing is remembered about — one hovered for
    /// the first time — starts at the beginning, which is what every file did before there
    /// was a setting.
    Remember,
    /// At the beginning, whatever the file is and whatever was heard of it before.
    Start,
    /// Half way in, for a file whose worth is somewhere past its opening.
    Middle,
    /// Anywhere in the file at all, so that a folder of sounds is a different few seconds of
    /// each one every time it is crossed.
    Random,
}

impl AudioSeek {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Remember => "remember",
            Self::Start => "start",
            Self::Middle => "middle",
            Self::Random => "random",
        }
    }

    /// The way a `config.ini` value names, or `None` for one that names no way of starting a
    /// sound.
    pub(super) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "remember" | "resume" | "last" | "last position" => Some(Self::Remember),
            "start" | "beginning" | "from start" | "from the start" => Some(Self::Start),
            "middle" | "half" | "halfway" | "from middle" | "from the middle" => Some(Self::Middle),
            "random" | "shuffle" | "anywhere" => Some(Self::Random),
            _ => None,
        }
    }
}

/// Where a sound starts unless the configuration says otherwise: where it was left the last
/// time it was hovered, which is the answer that makes a folder of music behave the way a
/// player does — the file picked up again rather than begun again.
pub const DEFAULT_AUDIO_SEEK: AudioSeek = AudioSeek::Remember;

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
    pub(super) fn from_str(value: &str) -> Option<Self> {
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

/// Whether the trigger key reaches a pinned preview unless the configuration says otherwise:
/// off, which is where the app starts — a pin is a window the user put there, and the key that
/// stops hovers is not read while one is up — and on, the key is what brings a pin down with the
/// previews it stops. It is `Hold to Disable Preview` the setting speaks for; the reverse mode is
/// left as it is (see the `Timing → Trigger Key` submenu).
pub const DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE: bool = false;

/// Whether a pin is shown another file while it is up unless the configuration says otherwise:
/// a file the pointer clicks, or one the keyboard selects, becomes what the pin is showing.
///
/// On, which is what makes a pin the thing a Quick Look window is — a window that follows the
/// listing rather than one that has to be closed and taken up again on every file read.
pub const DEFAULT_PIN_UPDATE_ENABLED: bool = true;

/// And whether the pointer's own hover is one of the ways it follows, which the configuration
/// does not ask for by default: off, a pin moves when the user clicks or presses a key, and a
/// pointer crossed over a listing — or parked over another file — leaves it where it is.
pub const DEFAULT_PIN_UPDATE_ON_HOVER: bool = false;

/// Whether a pin collapsed into its bubble holds the video it is playing where it is, which it
/// does unless the configuration says otherwise: a bubble is a pin put away, and a film playing
/// on behind one is heard from a window nobody can see. What is held is put back by the pin
/// coming up again, at the second it was stopped at (see the `Pin Mode` submenu).
pub const DEFAULT_PIN_PAUSE_VIDEO: bool = true;

/// And whether it holds a sound, which is the same question asked of the other kind and
/// answered the same way: a card is not read from a bubble either, and a sound is the half of
/// a preview that is heard where nothing of it is seen (see `DEFAULT_PIN_PAUSE_VIDEO`).
pub const DEFAULT_PIN_PAUSE_AUDIO: bool = true;

/// Which files a pinned window's own previous/next buttons step through.
///
/// The pin walks its folder rather than its own kind, so the two answers are what the walk is
/// made of. `All` is every file this build could preview, so a `.mp4` sits beside a `.mp3`
/// and the buttons are a way through a folder rather than a way through one kind of it.
/// `Category` narrows the walk to what the pinned file is — a picture, a sound, a document —
/// which is what a folder of mixed work wants, where stepping from a photograph into a video
/// is a different gesture from stepping to the next photograph.
///
/// The two are categories rather than kinds because the kinds are claims: a camera raw and a
/// JPEG are both pictures to anyone walking a folder, and whether the picture was developed
/// by a converter on this machine is not a question the buttons on a caption should ask (see
/// `formats::routing::nav_category`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PinNavFileTypes {
    /// Every file this build can preview, whatever kind it is.
    All,
    /// Only the files of the pinned file's own category.
    Category,
}

impl PinNavFileTypes {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::All => "all",
            Self::Category => "category",
        }
    }

    /// The way a `config.ini` value names, or `None` for one that names no set of files.
    pub(super) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "all" | "every" | "all files" | "all file types" => Some(Self::All),
            "category" | "same" | "same category" | "same type" => Some(Self::Category),
            _ => None,
        }
    }
}

/// What the pin's own buttons step through unless the configuration says otherwise: every
/// file this build can preview, which is the answer that makes them a way through a folder.
pub const DEFAULT_PIN_NAV_FILE_TYPES: PinNavFileTypes = PinNavFileTypes::All;
