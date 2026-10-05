//! The configuration itself: every field a `config.ini` can hold, and the values a
//! configuration that has just been made starts at.

use crate::formats::lists;
use crate::readers::tone_map::Curve;

use super::defaults::{
    DEFAULT_AFK_TIMER_SECS, DEFAULT_ANIMATED_SCALE_PERCENT, DEFAULT_AUDIO_VOLUME,
    DEFAULT_DECODE_BUDGET_GB, DEFAULT_DOCUMENT_CACHE_MB, DEFAULT_GENERAL_DISK_CACHE_MB,
    DEFAULT_HDR_EXPOSURE, DEFAULT_HDR_TONE_MAP, DEFAULT_HOVER_DELAY_MS, DEFAULT_IMAGE_CACHE_MB,
    DEFAULT_IMAGE_DISK_CACHE_MB, DEFAULT_LIBREOFFICE_IDLE_SECS, DEFAULT_NORMALIZE_VIDEO_VOLUME,
    DEFAULT_NORMALIZE_VOLUME, DEFAULT_OFFICE_ENGINE, DEFAULT_OFFICE_ENGINE_IDLE_SECS,
    DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_REMEMBER_AUDIO_VOLUME, DEFAULT_REMEMBER_VIDEO_VOLUME,
    DEFAULT_RENDER_HTML, DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS,
    DEFAULT_SPINNER_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_TICK_MS, DEFAULT_TTC_FACE,
    DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK, DEFAULT_VIDEO_HW_ACCEL,
    DEFAULT_VIDEO_SCALE_PERCENT, DEFAULT_VIDEO_VOLUME, DEFAULT_WEBP_PLAYBACK_FPS,
    DEFAULT_WEBVIEW_IDLE_SECS,
};
use super::setting_types::{
    AudioSeek, AvoidMode, EngineIdle, MarkdownMode, OfficeEngine, PinNavFileTypes, PreviewScale,
    TextTheme, TransparentBackground, TriggerKeyMode, VideoEngine, DEFAULT_AUDIO_SEEK,
    DEFAULT_AVOID_MODE, DEFAULT_DDS_BACKGROUND, DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE,
    DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE, DEFAULT_FOLLOW_CURSOR, DEFAULT_FONT_BACKGROUND,
    DEFAULT_FONT_SCALE, DEFAULT_HTML_BACKGROUND, DEFAULT_IMAGE_BACKGROUND,
    DEFAULT_PIN_NAV_FILE_TYPES, DEFAULT_PIN_PAUSE_AUDIO, DEFAULT_PIN_PAUSE_VIDEO,
    DEFAULT_PIN_UPDATE_ENABLED, DEFAULT_PIN_UPDATE_ON_HOVER, DEFAULT_TEXT_SCALE,
    DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE, DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE,
};

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
    /// Whether the trigger key reaches a pinned preview: on, holding the key brings the pin
    /// down with the previews it stops, and off — where the app starts — the key is not read
    /// while a preview is pinned, up or collapsed into its bubble, since a pin is a window the
    /// user put there rather than a hover for the key to hold back. It is the `Hold to Disable
    /// Preview` mode the setting speaks for; the reverse mode is left as it is (see the
    /// `Timing → Trigger Key` submenu).
    pub trigger_key_affect_pin_mode: bool,
    /// Whether the pin key is watched: a preview on screen when it is pressed becomes
    /// a window of its own — captioned, movable, always on top — and previews stay
    /// quiet until that window is closed (see `pin_key`).
    pub pin_enabled: bool,
    /// The key that pins the preview on screen, by name — the same spellings the
    /// trigger key accepts, since both are read by the same table.
    pub pin_key: String,
    /// Whether a pin collapsed into its bubble holds the video it is playing where it is:
    /// on, a film is paused the moment the window becomes a bubble and started again at the
    /// second it was stopped at when the pin comes back up, and off, it plays on behind the
    /// bubble as it always did (see the `Pin Mode` submenu).
    pub pin_pause_video: bool,
    /// The same question about a sound, which is a switch of its own because the two are two
    /// different things to want quiet — a film a user wants to hear while the bubble is up,
    /// and a podcast they want left alone while the pin beside it is restored (see
    /// `pin_pause_video`).
    pub pin_pause_audio: bool,
    /// Whether a pin is shown another file while it is up: a file the pointer clicks, or one
    /// the keyboard selects, becomes the pin's own — shown in the box the pin already has
    /// rather than as a second preview beside it (see the `Pin Mode` submenu).
    ///
    /// It is a setting of its own rather than part of the pin, because a pin is also a window
    /// to read or work in — a text frame, a video being watched — and one that swapped its file
    /// out from under the hand every time a key crossed the folder would be a window nobody
    /// could read. Off, the pin keeps the file it was taken up on until it is closed.
    pub pin_update_enabled: bool,
    /// Whether that following includes the pointer's own hover, or only what the user asks for
    /// with a click or a key: on, the pin is shown whatever the pointer settles on, one file
    /// after another, the way a Quick Look window does; off, moving the pointer across a
    /// listing leaves the pin alone (see `pin_update_enabled`).
    pub pin_update_on_hover: bool,
    /// Which files the pin's own previous/next buttons step through: every file this build
    /// could preview, or only the ones of the pinned file's own category (see
    /// `DEFAULT_PIN_NAV_FILE_TYPES`).
    ///
    /// It is a setting of its own because the two are two different folders: a folder of mixed
    /// work is a list the buttons walk end to end under `All` and only part of under
    /// `Category`, and neither is the other. The walk itself is over the folder the pin was
    /// taken up in and no other, and its order is the order the listing is showing.
    pub pin_nav_file_types: PinNavFileTypes,
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
    /// Whether the keyboard driving Explorer holds a parked pointer back rather than
    /// sharing the screen with it: with this on, the file under a pointer that has not
    /// been moved since a key press does not preview of its own — a key pressed onto an
    /// item with no preview to give reads exactly as one pressed onto an item that has a
    /// preview — until the pointer takes its turn back with a move or a wheel, or a
    /// folder change hands it over. It is on where the app starts; switched off, the
    /// pointer's own hover raises a preview whenever a file is under it, whatever the
    /// keyboard is doing, and a key pressed onto an item with no preview to give leaves
    /// that hover standing (see `explorer_hook`).
    pub prioritize_keyboard: bool,
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
    /// The backdrop a page of HTML is drawn over.
    ///
    /// A setting of its own for the reason a specimen's is: a page is a page, and what is
    /// behind it is the page to read it against rather than a transparency to be shown
    /// through — the markup brings its own colours and a background behind it is only ever
    /// the page's. Three of the four backdrops are offered, all but the transparent one, so
    /// a file that names transparency — which is what a page's preview was held at before
    /// the kind had a setting of its own — is read as the one this setting starts at rather
    /// than kept as a value the menu beside it has no item for (see
    /// `sanitize_html_background`).
    pub html_background: TransparentBackground,
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
    /// The volume a sound file is previewed at, as a percentage of what the file holds.
    ///
    /// A setting of its own rather than the video's above it, because the two are hovered for
    /// different things: a video's soundtrack is played while its picture is looked at, and a
    /// sound file is played instead of being drawn at all. Which of the two a file answers to
    /// is the kind the router gave it, so a video whose streams hold no picture answers to
    /// this one — the same answer the card beside it is drawn from.
    pub audio_volume: u32,
    /// Where in the file a sound starts playing — the same question the volume above it
    /// answers, and a setting of its own for the same reason: what a hover on a sound is
    /// worth depends on how much of it is heard, and a file being listened to again is
    /// usually wanted from where it was left rather than from the top (see `AudioSeek`).
    pub audio_seek: AudioSeek,
    /// Whether a sound's loudness is measured and brought to one level before it is played —
    /// the one thing that asks a file to be as loud as the next rather than as loud as it was
    /// recorded.
    ///
    /// A setting of its own rather than a level among the ones above it, because it is a
    /// question about the file rather than about the hover: a level says how loud this app
    /// should be, and this says where a file's own loudness is counted from. It is on where the app
    /// starts, and on a machine without FFmpeg it does nothing at all — what measures the loudness
    /// and what applies the gain are both FFmpeg's (see `codecs::normalize_available`), which is
    /// also why the tray greys the row where FFmpeg is not installed.
    pub normalize_volume: bool,
    /// Whether a video's soundtrack is measured and brought to the same level before it is played,
    /// on the same terms as the sound above it and by the same measurement.
    ///
    /// It is a setting of its own rather than one shared with the sound's, because the two are
    /// hovered for different things: a sound file is the sound, and a film's soundtrack is heard
    /// beside a picture that was asked for. It is off where the app starts, the way the video's
    /// own level starts at silence — and like the sound's it does nothing without FFmpeg, which is
    /// why the tray greys the row where FFmpeg is not installed.
    pub normalize_video_volume: bool,
    /// Whether a sound previewed from the pinned window's own knob is previewed at that level next
    /// time as well, rather than at whatever `audio_volume` says.
    ///
    /// It is asked of the knob rather than of the menu, because that is where the level changes:
    /// with it on, letting go of the knob writes the level to the file, and the next hover plays
    /// at it (see `current_audio_volume`). With it off the knob is the pin's own — a window's level
    /// rather than the app's, which is why nothing at all is written when it moves.
    pub remember_audio_volume: bool,
    /// And the same question asked about a video's soundtrack, on its own switch for its own
    /// reasons: a film is looked at, and its level is more often a fault to get past than a
    /// preference to keep (see `remember_video_volume`).
    pub remember_video_volume: bool,
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
    /// How large a PDF page — the `Ebook` kind — is drawn, as a share of the room the
    /// display has for it.
    ///
    /// A page is a vector, so what it is asked for is a size rather than a resample:
    /// the room the display has is free quality there, and a share of that room is what
    /// the setting names — `50%` is half the display, not half of the page. `Fit to
    /// Screen` is the whole of it, which is where the setting starts, and `100%` or
    /// more reads as that fit.
    ///
    /// It is a setting of its own rather than the vector scale beside it because the two
    /// documents are hovered for different reasons: a page is read at a glance and a
    /// drawing is looked at, so the size one wants is rarely the size the other wants.
    pub ebook_scale: PreviewScale,
    /// How large a page of the `Document` kind is shown, as a share of the room the
    /// display has — the same question, and the same answers, as the Ebook scale beside it.
    ///
    /// One setting covers both halves of the kind — the page an Office document's own
    /// application exports, and the page the render engine draws for a document no
    /// application here has — because it is one question asked of one shape of preview: how
    /// much of the display a page is given. See `effective_preview_scale`.
    ///
    /// The one source that is not a page is the bitmap a workbook is answered with
    /// where no page can be exported: it is only as good as the pixels it holds, so it
    /// follows the share as a share of its own size and is never enlarged.
    pub document_scale: PreviewScale,
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
    /// How large a page of text is measured for, as a share of the room the display has —
    /// the same question, and the same answers, as the document scales above.
    ///
    /// What the share is applied to is the room rather than a size the file asks for, because
    /// a page of text has no size of its own: it is measured at the font size the display
    /// gives it, and the answer to how large it is drawn is how much display it is allowed to
    /// be measured in. A page that takes less room than the share allows keeps the room it
    /// takes, so this is a ceiling on the measure rather than a zoom over it — and the whole
    /// of the display is where it starts, which is the room this kind was given before the
    /// setting existed.
    pub text_scale: PreviewScale,
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
    /// Whether a page of HTML is drawn by the browser engine instead of shown as its
    /// markup.
    ///
    /// A `.htm` and a `.html` are in the text lists, so what they are previewed by is the
    /// text preview: the markup, at a fixed font size, in a page of this app's own. This is
    /// what asks for the page rather than the markup — the engine that draws an SVG document
    /// draws the page in a window of its own, and the box it is given is the share of the
    /// display `document_scale` names, which is the document question and not the text one.
    ///
    /// It is a switch over two names and nothing else: `.xhtml` is a page of text, as is
    /// every other file the text lists claim, and a machine with no WebView2 runtime keeps
    /// the text preview whatever this says.
    pub render_html: bool,
    /// Whether pictures are previewed at all, ahead of the image list their names are
    /// entries of.
    ///
    /// It is the gate for the pictures this app decodes and for the ones an installed
    /// ImageMagick develops — the camera raw formats above all, which is `[magick]` — since
    /// both are pictures: what the engine hands back is a PNG, drawn at the picture scale,
    /// over the picture backdrop, held in the picture cache. Which of the two a file is, is
    /// the list its name is in rather than a switch of its own.
    pub image_preview_enabled: bool,
    /// Whether video previews may be shown at all.
    pub video_preview_enabled: bool,
    /// Whether sound files are previewed at all, ahead of the sound list their names are
    /// entries of, and of the card a preview of one is.
    ///
    /// It is the switch over both engines that play a file — the media engine Windows has and
    /// FFmpeg's player — since which of the two a file is, is the machine's answer rather than
    /// the user's, exactly as it is for the two halves of the video kind.
    pub audio_preview_enabled: bool,
    /// Whether text files are previewed at all, ahead of the extension list.
    pub text_preview_enabled: bool,
    /// Whether a PDF — the `Ebook` kind — is previewed at all.
    pub ebook_preview_enabled: bool,
    /// Whether archive contents are listed at all, ahead of the extension list.
    ///
    /// It is the gate for the archives this app reads and for the ones an installed PeaZip
    /// lists — the cabinet files, isos, disk images and installers that are `[peazip]` —
    /// since both are answered with the same page of contents. Which of the two a file is,
    /// is the list its name is in rather than a switch of its own.
    pub archive_preview_enabled: bool,
    /// Whether documents are previewed at all, ahead of the lists their names are entries of.
    ///
    /// It is the gate for both halves of the `Document` kind: the documents whose own
    /// application exports a page, and the ones an installed render engine draws — the word
    /// processors, spreadsheets, presentations and drawings of `[libre]`, CorelDRAW above
    /// all — since what either is previewed as is a page, drawn at the same share of the
    /// display. Which of the two draws it is a question about the machine rather than a
    /// switch; see `office_formats::page_engine`.
    pub document_preview_enabled: bool,
    /// Whether font files are previewed at all, ahead of the font list their names are
    /// entries of.
    pub font_preview_enabled: bool,
    /// Whether design documents and projects are previewed at all, ahead of the design
    /// list their names are entries of.
    pub design_preview_enabled: bool,
    /// Whether vector drawings are previewed at all, ahead of the vector list their names
    /// are entries of, and of the browser engine a document of that kind is drawn by.
    pub vector_preview_enabled: bool,
    /// How much of what the engines drew is kept between hovers, in megabytes: the pages the
    /// render tier exported and the pages the engine beside it converted, kept as files under
    /// the temp folder and given up least recently used first. See
    /// `Performance → Cache → Document` in the tray.
    pub document_cache_mb: u32,
    /// How much of what an image-developing engine developed is kept between hovers, in
    /// megabytes: the PNG the engine wrote for a file, kept as a page under the temp folder
    /// beside the documents' pages and given up least recently used first. See
    /// `Performance → Cache → Image (Disk)` in the tray.
    ///
    /// It is the one cache a developed picture has of its own: the frame it was drawn as is
    /// `image_cache_mb`, and a page of a *document* is `document_cache_mb` whatever it is drawn
    /// as — a slide's PNG and a workbook's picture included.
    pub image_disk_cache_mb: u32,
    /// How much of what a film's own subtitle tracks were copied into is kept between hovers, in
    /// megabytes: the small subtitle files this app's extraction wrote under the general folder,
    /// with the fonts that came out of the container dumped beside them, given up least recently
    /// used first. See `Performance → Cache → General (Disk)` in the tray.
    ///
    /// It is the one cache a hover's subtitles have, and it is what makes a hover a read of a few
    /// dozen kilobytes instead of a stream of the whole film: the filter that draws an embedded
    /// track opens the film and reads it to its first subtitle — measured at 14 904 ms and
    /// 1 423 MB read on a cold 1.4 GB MKV, against 492 ms and 41 MB for the same film drawn from
    /// its extracted `.ass` (see `subtitle_files`).
    pub general_disk_cache_mb: u32,
    /// Which engine draws an Office document's page, which is the tray's
    /// `Engine → Select Engine → Office` setting.
    pub office_engine: OfficeEngine,
    /// Which engine plays a video, which is the tray's `Engine -> Select Engine -> Video`
    /// setting. `Best` is the machine's own answer and what the app starts at; the other two
    /// name one engine each — see `video_hw::resolve_video_engine`.
    pub video_engine: VideoEngine,
    /// Whether an explicitly chosen engine that cannot play a given file falls through to the
    /// others, which is the `Fallback` switch at the top of the same submenu. It is read only
    /// where the choice names an engine rather than `Best`, since `Best` is a walk of the same
    /// list already.
    pub video_engine_fallback: bool,
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
    /// What one hover may decode or read for, in gigabytes.
    ///
    /// The one setting here that bounds a file rather than a cache: a reader is
    /// handed it before it allocates, so a file larger than the budget is answered
    /// with no preview instead of with memory the app may not get.
    pub decode_budget_gb: f32,
    /// Whether a video is decoded on the graphics card rather than on a core, which is the
    /// `Video` toggle under `Performance → Hardware Acceleration` in the tray.
    ///
    /// It is a setting about *this build's* FFmpeg rather than about a kind of file, which is
    /// why it is asked of the player that is launched and not of the route: a file this app
    /// hands to FFmpeg is decoded by whatever FFmpeg is on the machine, and whether that is a
    /// card or a core is the only thing this names (see `video_hw_accel_device`).
    pub video_hw_accel: bool,
    /// The curve a picture whose samples are light is brought into eight bits with — an
    /// EXR, a Radiance HDR, a float texture — as `hdr_tone_map` names it.
    pub hdr_tone_map: Curve,
    /// How many stops those pictures are shifted by before that curve, as `hdr_exposure`.
    pub hdr_exposure: f32,
    /// Font scale for text previews, as a percentage of the default size.
    pub text_font_scale_percent: u32,
    /// How far past the far edge of a text preview the pointer region reaches, in
    /// logical pixels at the display's DPI, so a hand that overshoots the edge on
    /// its way to the scrollbar does not take the preview down with it.
    pub text_scroll_far_edge_grace_pixels: f32,
    /// Extensions previewed as images, already normalized for lookup.
    pub image_extensions: Vec<String>,
    /// The video names the media engine Windows has is asked to play, as `[video] extensions`
    /// in `config.ini`: the containers and streams the codecs Windows ships demux and decode,
    /// already normalized for lookup.
    ///
    /// A file of one of these names is played by the engine, in this app's own window, so its
    /// frames are this app's to draw — which is what a pinned window of one is resized,
    /// maximized and dragged by. Whether the machine in hand really decodes *this* file is asked
    /// of the engine itself, once per file, and a file it turns down is played by FFmpeg's
    /// player where one is installed. See `video_formats`.
    pub video_extensions: Vec<String>,
    /// The video names only FFmpeg's player reads, as `[ffmpeg] extensions` in `config.ini`:
    /// what the two engines' own coverage leaves to it, already normalized for lookup.
    ///
    /// The engine is never asked about a name in this list, and a machine with no FFmpeg shows
    /// nothing for one — moving a name up into `[video]` is the whole of asking the engine to
    /// play it instead. See `video_formats`.
    pub ffmpeg_extensions: Vec<String>,
    /// The names of the sounds this app plays, as `[audio] extensions` in `config.ini`: the
    /// formats the media engine Windows has decoders for and the ones only an installed FFmpeg
    /// reads, in one list, because which engine plays a file is the machine's answer rather
    /// than a setting. A user who wants their music left alone takes names out of it, or
    /// switches the kind off. See `audio_formats`.
    pub audio_extensions: Vec<String>,
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
    /// `config.ini`: the formats its console archiver reads and this app has no reader of its own
    /// for — the cabinet files, isos, disk images and single-stream compressors. See
    /// `peazip_formats`.
    pub peazip_extensions: Vec<String>,
    /// The names of the books the Calibre engine is asked about, as `[calibre] extensions` in
    /// `config.ini`: the ebook formats its converter reads and this app has no reader of its own
    /// for — the Kindle and Mobipocket families, the open EPUB, the FictionBook, the scanned book
    /// and the containers of the dedicated readers. See `calibre_formats`.
    pub calibre_extensions: Vec<String>,
    /// The names of the pages this app draws itself, as `[ebook] extensions` in `config.ini`: the
    /// PDF's three spellings, which its own reader draws, and the three comic containers, whose
    /// first plate is a picture inside the box. See `ebook_formats`.
    ///
    /// It is the one list that holds two readers' worth of names, because a user asks one question
    /// of them — what is a book — rather than which reader is the right one: what the PDF reader is
    /// asked about is a name, and what the comic reader answers for is decided by the file.
    pub ebook_extensions: Vec<String>,
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
            trigger_key_affect_pin_mode: DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
            pin_enabled: true,
            pin_key: "space".to_string(),
            pin_pause_video: DEFAULT_PIN_PAUSE_VIDEO,
            pin_pause_audio: DEFAULT_PIN_PAUSE_AUDIO,
            pin_update_enabled: DEFAULT_PIN_UPDATE_ENABLED,
            pin_update_on_hover: DEFAULT_PIN_UPDATE_ON_HOVER,
            pin_nav_file_types: DEFAULT_PIN_NAV_FILE_TYPES,
            follow_cursor: DEFAULT_FOLLOW_CURSOR,
            avoid_mode: DEFAULT_AVOID_MODE,
            same_file_rehover_delay_ms: DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
            settling_delay_ms: DEFAULT_SETTLING_DELAY_MS,
            prioritize_keyboard: true,
            tick_ms: DEFAULT_TICK_MS,
            spinner_delay_ms: DEFAULT_SPINNER_DELAY_MS,
            webp_playback_fps: DEFAULT_WEBP_PLAYBACK_FPS,
            image_cache_mb: DEFAULT_IMAGE_CACHE_MB,
            image_background: DEFAULT_IMAGE_BACKGROUND,
            font_background: DEFAULT_FONT_BACKGROUND,
            dds_background: DEFAULT_DDS_BACKGROUND,
            design_background: DEFAULT_DESIGN_BACKGROUND,
            html_background: DEFAULT_HTML_BACKGROUND,
            vector_background: DEFAULT_VECTOR_BACKGROUND,
            video_volume: DEFAULT_VIDEO_VOLUME,
            audio_volume: DEFAULT_AUDIO_VOLUME,
            audio_seek: DEFAULT_AUDIO_SEEK,
            normalize_volume: DEFAULT_NORMALIZE_VOLUME,
            normalize_video_volume: DEFAULT_NORMALIZE_VIDEO_VOLUME,
            remember_audio_volume: DEFAULT_REMEMBER_AUDIO_VOLUME,
            remember_video_volume: DEFAULT_REMEMBER_VIDEO_VOLUME,
            preview_scale: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video_scale: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated_scale: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            ebook_scale: DEFAULT_EBOOK_SCALE,
            document_scale: DEFAULT_DOCUMENT_SCALE,
            font_scale: DEFAULT_FONT_SCALE,
            design_scale: DEFAULT_DESIGN_SCALE,
            vector_scale: DEFAULT_VECTOR_SCALE,
            text_scale: DEFAULT_TEXT_SCALE,
            ttc_face: DEFAULT_TTC_FACE,
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Rendered,
            render_html: DEFAULT_RENDER_HTML,
            image_preview_enabled: true,
            video_preview_enabled: true,
            audio_preview_enabled: true,
            text_preview_enabled: true,
            ebook_preview_enabled: true,
            archive_preview_enabled: true,
            document_preview_enabled: true,
            font_preview_enabled: true,
            design_preview_enabled: true,
            vector_preview_enabled: true,
            document_cache_mb: DEFAULT_DOCUMENT_CACHE_MB,
            image_disk_cache_mb: DEFAULT_IMAGE_DISK_CACHE_MB,
            general_disk_cache_mb: DEFAULT_GENERAL_DISK_CACHE_MB,
            office_engine: DEFAULT_OFFICE_ENGINE,
            video_engine: DEFAULT_VIDEO_ENGINE,
            video_engine_fallback: DEFAULT_VIDEO_ENGINE_FALLBACK,
            office_engine_idle: EngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS),
            webview_idle: EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS),
            libreoffice_idle: EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS),
            afk_timer_seconds: DEFAULT_AFK_TIMER_SECS,
            office_engine_persistent: false,
            webview_persistent: false,
            libreoffice_persistent: false,
            decode_budget_gb: DEFAULT_DECODE_BUDGET_GB,
            video_hw_accel: DEFAULT_VIDEO_HW_ACCEL,
            hdr_tone_map: DEFAULT_HDR_TONE_MAP,
            hdr_exposure: DEFAULT_HDR_EXPOSURE,
            text_font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            text_scroll_far_edge_grace_pixels: DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS,
            image_extensions: Vec::new(),
            video_extensions: Vec::new(),
            ffmpeg_extensions: Vec::new(),
            audio_extensions: Vec::new(),
            text_extensions: Vec::new(),
            text_names: Vec::new(),
            archive_extensions: Vec::new(),
            office_extensions: Vec::new(),
            font_extensions: Vec::new(),
            design_extensions: Vec::new(),
            libre_extensions: Vec::new(),
            magick_extensions: Vec::new(),
            peazip_extensions: Vec::new(),
            calibre_extensions: Vec::new(),
            ebook_extensions: Vec::new(),
            vector_extensions: Vec::new(),
        }
        .with_built_in_lists()
    }
}

impl AppConfig {
    /// This configuration with every extension list at what this build ships it with.
    ///
    /// The lists are not written out here one field each, which is what they used to be: what
    /// each one holds is a row of the table in `formats::lists`, and a kind added to the app is a
    /// row added to that. So the sixteen empty lists above are the only mention of them in this
    /// file, and it is this line that says what they are for.
    fn with_built_in_lists(mut self) -> Self {
        lists::reset_built_in(&mut self);
        self
    }
}
