//! The command ids the tray hands out, and the tables each menu lists.
//!
//! Every row of this tray is a command id, and a click is answered by comparing against
//! them: the ids are what the window proc reads a `WM_COMMAND` against, and what the menus
//! are built with. Which is why they live here rather than beside the menu that hands them
//! out - a range that ran into another submenu's is a click answered as the wrong setting
//! (see `tray::tests`).
//!
//! What sits here is one of three things: a single row's id, a range a submenu hands out
//! (a base plus the position the choice was listed at), or the table of choices the range is
//! as wide as. The statics at the bottom are the tray window's own: its class name, its
//! handle, the message Explorer sends when it restarts, and the themes the menu last listed.

use crate::config::config::{
    AudioSeek, AvoidMode, EngineIdle, PreviewScale, TextTheme, TransparentBackground, VideoEngine,
};
use once_cell::sync::Lazy;
use std::sync::Mutex;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::WM_USER;

pub(super) const WM_TRAYICON: u32 = WM_USER + 1;
pub(super) const ID_TRAY_EXIT: u16 = 1001;
pub(super) const ID_TRAY_STARTUP: u16 = 1002;
pub(super) const ID_TRAY_ENABLE: u16 = 1003;
pub(super) const ID_TRAY_PIN: u16 = 1004; // Whether a preview can be pinned with a key
/// The two rows of the `Pin Mode → Update Preview` submenu: whether a pin is shown another
/// file while it is up — the one the pointer clicks or the keyboard selects — and whether the
/// pointer's own hover is one of the ways it is told about one.
///
/// They sit in the slack between the pin's own row and the trigger key's, which is the gap
/// this block has left: a row of the submenu must never be read as a click on the row beside
/// it, and the two questions belong together under the one row they hang from.
pub(super) const ID_TRAY_PIN_UPDATE: u16 = 1016;
pub(super) const ID_TRAY_PIN_UPDATE_HOVER: u16 = 1017;
/// And the two rows of the `Pin Mode → Pause in Bubble` submenu: whether a pin collapsed into its
/// bubble holds the sound it is playing where it is, and whether it holds a video. They sit in
/// the slack the same block left beside the update pair, for the same reason — both are the
/// pin's own business and neither is a click on the row beside it.
pub(super) const ID_TRAY_PIN_PAUSE_AUDIO: u16 = 1018;
pub(super) const ID_TRAY_PIN_PAUSE_VIDEO: u16 = 1019;
/// The two rows of the `Pin Mode → Navigation Files` submenu: whether a pin's own previous/next
/// buttons walk every file this build can preview or only those of the pinned file's own kind
/// of thing (see `PinNavFileTypes`).
///
/// They sit in the slack above the pin's own rows, which begin at
/// 1016 — the one gap this block has left between the HTML
/// backdrop's half, which ends at 1012, and the pin's own rows:
/// a walk's two answers belong under the pin's own row rather
/// than in a range of their own, since a setting of two ways is
/// two ids and not a list.
pub(super) const ID_TRAY_PIN_NAV_ALL: u16 = 1014;
pub(super) const ID_TRAY_PIN_NAV_CATEGORY: u16 = 1015;
/// The two rows of the `Pin Mode → Minimize` submenu: where a pin the minimize button put away
/// goes — the round bubble left on the desktop, or the app's own button in the taskbar.
///
/// They are a question of two answers rather than a range, and they do not sit beside each other
/// on the number line: 1013 is the one id the `Background → HTML` half leaves before the walk's
/// pair at 1014/1015, and 1022 is the one the `Placement` pair leaves before the image backdrop's
/// range at 1023. Neither is inside a range a click is read against, which is what keeps a click
/// on one from being answered as a row of a submenu of another's.
pub(super) const ID_TRAY_PIN_MINIMIZE_BUBBLE: u16 = 1013;
pub(super) const ID_TRAY_PIN_MINIMIZE_TASKBAR: u16 = 1022;
pub(super) const ID_TRAY_TRIGGER_DISABLE: u16 = 1005; // Hold the trigger key to stop previews
pub(super) const ID_TRAY_TRIGGER_ENABLE: u16 = 1006; // Hold the trigger key to allow previews
pub(super) const ID_TRAY_TRIGGER_ENABLED: u16 = 1068; // Whether the trigger key is watched at all
/// The trigger key's second switch: whether a held key reaches a pinned preview. It sits in the
/// one id left between the `Audio` gate at 1098 and the `theme` folder's items at 1100.
pub(super) const ID_TRAY_TRIGGER_AFFECT_PIN: u16 = 1099;
/// The `Engine → Office Engine` pair: which engine an Office document's page is
/// asked of — the application that owns the format, or the render engine beside it. Two ids
/// rather than a range, the way the trigger key's mode has two, and they sit in the gap the
/// update row at 1007 leaves before the page's backdrop range that begins at 1010.
pub(super) const ID_TRAY_ENGINE_OFFICE_MS: u16 = 1008;
pub(super) const ID_TRAY_ENGINE_OFFICE_LIBRE: u16 = 1009;
/// The `Engine → Video Engine` submenu: the `Fallback` toggle at the top, and one
/// row per engine below it. The toggle is a switch and the engines are a radio group, so the
/// toggle has an id of its own and the engines are a base plus the position each was listed at
/// (see `VIDEO_ENGINE_CHOICES`). They sit in the slack the LibreOffice idle range leaves before
/// the config row at 1040.
pub(super) const ID_TRAY_VIDEO_ENGINE_FALLBACK: u16 = 1034;
pub(super) const ID_TRAY_VIDEO_ENGINE_BASE: u16 = 1035;

/// The engines the `Video` submenu lists, in the order it lists them, with `Best` at the top and
/// the hybrid — `Native (FFmpeg above 3.2MP)` — below the two plain ones.
pub(super) const VIDEO_ENGINE_CHOICES: [VideoEngine; 4] = [
    VideoEngine::Best,
    VideoEngine::Native,
    VideoEngine::Ffmpeg,
    VideoEngine::Hybrid,
];
/// The `Background` submenu: one command per backdrop it offers, in the order it
/// lists them, for each of the six kinds of preview it keeps apart — a picture's
/// backdrop, a vector drawing's, a page of HTML's, a font specimen's, a texture's, and a
/// design document's. Each half is a base plus the position the choice was listed at, so
/// one table and one builder serve all of them, and each range is exactly as wide as the
/// choices its half lists, which is what keeps an item of one from being read as a choice
/// of another.
pub(super) const ID_TRAY_IMAGE_BACKGROUND_BASE: u16 = 1023;
/// The second half, for a vector drawing — an SVG document the browser draws, or a
/// metafile the drawing layer replays: one page's backdrop for both halves of that kind,
/// since what stands behind a drawing is the same question either way.
pub(super) const ID_TRAY_VECTOR_BACKGROUND_BASE: u16 = 1054;
/// The third half of the `Background` submenu, for a font specimen — which is a page of its
/// own and so has a backdrop of its own, the same way a document does.
///
/// It sits beside the texture's half rather than in the block the first two are in: every
/// id four wide in that block belongs to something else — 1058 to 1061 is where the text
/// preview's `Full Mode` item is, which is where this range was to begin with, so the last
/// two backdrops of a specimen were answered as full mode being switched on.
pub(super) const ID_TRAY_FONT_BACKGROUND_BASE: u16 = 1204;
/// The fourth, for a `.dds` texture: a texture's alpha channel is as often a mask or a
/// channel nobody filled in as it is transparency, so what is drawn behind one is a
/// question of its own.
pub(super) const ID_TRAY_DDS_BACKGROUND_BASE: u16 = 1200;
/// The fifth, for a design document: what its preview is made of is the picture the file
/// keeps of the whole document — a merged image, or the flattened document a project
/// container holds — and that picture's transparency is the document's own, so what stands
/// behind one is a question of its own.
pub(super) const ID_TRAY_DESIGN_BACKGROUND_BASE: u16 = 1208;
/// The sixth, for a page of HTML: a document the browser is handed is a page already, so
/// what stands behind it is a page to read it against rather than a transparency to look
/// through — the reason it did not go on drawing over the vector half's backdrop.
///
/// It sits in the gap at 1010 to 1015, which is where the six single ids the `Volume`
/// submenu's two halves replaced used to be: a backdrop is a base plus a position in a list
/// rather than a name of its own, so it wanted the same stretch the volume halves gave up.
pub(super) const ID_TRAY_HTML_BACKGROUND_BASE: u16 = 1010;
/// The backdrops a half of the `Background` submenu offers, in the order it lists
/// them, with `Transparent` at the top: the whole range the setting holds, so nothing
/// a hand-edited `config.ini` can ask for is left unmarked.
pub(super) const BACKGROUND_CHOICES: [TransparentBackground; 4] = [
    TransparentBackground::Transparent,
    TransparentBackground::Black,
    TransparentBackground::White,
    TransparentBackground::Checkerboard,
];
/// The backdrops the `HTML` half offers, which are three of the four rather than
/// all of them: a page is drawn over a page rather than over what stands behind one, so
/// transparency is the one backdrop left off, and the white page the setting starts at is
/// listed first, as every other half lists its own. A file that names transparency anyway,
/// from when a page was drawn over the vector half's setting, is read as the backdrop the
/// setting starts at rather than kept as a value the menu has no item for (see
/// `sanitize_html_background`).
pub(super) const HTML_BACKGROUND_CHOICES: [TransparentBackground; 3] = [
    TransparentBackground::White,
    TransparentBackground::Black,
    TransparentBackground::Checkerboard,
];
/// The backdrops the `DDS` half offers, which are two of the four rather than all
/// of them: a texture's alpha channel is as often a mask, a height or a roughness as it is
/// transparency (see `dds_image`), so what is drawn behind one is a page to read the channels
/// against rather than a hole to look through — and the two backdrops that show what stands
/// behind a preview are the two that page has no use for. A file that names one of them
/// anyway, from when this half listed all four, is read as the backdrop the setting starts at
/// rather than kept as a value the menu has no item for (see `sanitize_dds_background`).
pub(super) const DDS_BACKGROUND_CHOICES: [TransparentBackground; 2] =
    [TransparentBackground::Black, TransparentBackground::White];
/// The two halves of the `Volume` submenu: one command per level it offers, in the order it
/// lists them, the video's range first and the sound's beside it. Each half is a base plus the
/// position a level was listed at, so one table and one builder serve both — the arrangement
/// the `Background`, `Scale` and `Timing` submenus already have — and the two ranges are ten
/// wide and apart from each other and from everything else, which is what keeps a level of one
/// from being read as a level of the other.
///
/// They sit in the stretch between the disk-cache range and the decode budget's, which is the
/// one gap this block had left. The six single ids they replace (1010 to 1015) are gone with
/// them: a volume is a level of one of two lists now rather than a name of its own. The bottom
/// three of that stretch have since gone to a page's backdrop, which is a list too (see
/// `ID_TRAY_HTML_BACKGROUND_BASE`).
pub(super) const ID_TRAY_VIDEO_VOLUME_BASE: u16 = 1360;
pub(super) const ID_TRAY_AUDIO_VOLUME_BASE: u16 = 1370;
/// The `Volume → Audio Seek` submenu: one command per way a sound can be started, in the order
/// it lists them. It is the `Volume` submenu's third item, below the two halves above it,
/// because the two questions belong together — a sound is heard at a level and from a place,
/// and both are answered the moment a hover starts its player rather than while one is playing.
///
/// Its range sits in the slack the `Decode Budget` ceilings leave before the `Vector`
/// half begins, and it is four wide because there are four ways a sound can be started. That it
/// is past the volume halves on the number line rather than beside them is the whole of what
/// keeps a level of either half from being read as a way of starting a sound — see
/// `the_two_volume_submenus_carry_a_range_apiece`, which holds all three away from the sounds
/// gate and from each other.
pub(super) const ID_TRAY_AUDIO_SEEK_BASE: u16 = 1390;
/// The `Volume → Pin Mode Audio Seek` submenu: one command per
/// way a *pinned* sound can be started, in the order it lists
/// them — the same four ways as the hover's own submenu above
/// it, asked of the pin's setting rather than the hover's (see
/// `pin_mode_audio_seek`). Its range sits directly below the
/// hover's, which is where the submenu sits below it, and is
/// four wide for the same reason the hover's is: a way of
/// starting a sound is one of four, and a click on one of the
/// pin's is never a click on one of the hover's.
pub(super) const ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE: u16 = 1394;
/// The `Volume → Audio` submenu's first row: whether a sound's measured loudness is brought to one
/// level before it is played (see `Normalize`).
///
/// It is a switch of its own rather than a level of the list under it, so it takes an id of its
/// own — out past the tick menu, where the tray's other lone switches sit, rather than in the
/// stretch the levels are read from.
pub(super) const ID_TRAY_NORMALIZE_VOLUME: u16 = 1525;
/// The video half's own row of the same name, which is the same switch asked about a film's
/// soundtrack: a setting of its own, and one that starts switched on along with its sound half
/// (see `normalize_video_volume`). It carries the id beside its sound half's, so neither row is
/// ever read as a level of a list under it.
pub(super) const ID_TRAY_NORMALIZE_VIDEO_VOLUME: u16 = 1526;
/// The row under a sound's `Normalize`: whether a level turned on a pinned window's own knob is the
/// level the next sound is previewed at, rather than the level the list above it names.
///
/// It is a switch about the level rather than a level itself, so it takes an id of its own — out
/// past the normalize rows, beside its video half's, where no level can be read from it (see
/// `remember_audio_volume`).
pub(super) const ID_TRAY_REMEMBER_VOLUME: u16 = 1527;
/// And the video's own row of the same name, which is the same switch asked about a soundtrack and
/// starts switched on along with it (see `remember_video_volume`).
pub(super) const ID_TRAY_REMEMBER_VIDEO_VOLUME: u16 = 1528;
/// The ways a sound can be started, in the order the submenu lists them: where it was left the
/// last time it was hovered, which is where the setting starts, then its beginning, its middle,
/// and anywhere in it at all (see `AudioSeek`).
pub(super) const AUDIO_SEEK_CHOICES: [AudioSeek; 4] = [
    AudioSeek::Remember,
    AudioSeek::Start,
    AudioSeek::Middle,
    AudioSeek::Random,
];
pub(super) const ID_TRAY_POSITION_FOLLOW: u16 = 1020; // Follow cursor
pub(super) const ID_TRAY_POSITION_BEST: u16 = 1021; // Best position
/// The `Placement → Avoid` submenu: one command per way a preview is kept off the
/// item it is about, in the order it lists them. The ids sit in a range of their own
/// in the slack the scaling ranges leave, so a way of avoiding is never read as a
/// share of a preview's size.
pub(super) const ID_TRAY_AVOID_BASE: u16 = 1410;
/// The ways the `Avoid` submenu offers, in the order it lists them: nothing avoided,
/// the name where it is drawn, the column the name is drawn in, and every column the
/// item draws.
pub(super) const AVOID_CHOICES: [AvoidMode; 4] = [
    AvoidMode::Off,
    AvoidMode::Filename,
    AvoidMode::FilenameColumn,
    AvoidMode::Details,
];
/// The `Timing` submenus that list a delay each — `Delay`, `Rehover Delay` and
/// `Settling Delay`, in the order the menu lists them — one range apiece, each as wide
/// as the delays it offers. They sit past every other range the app hands out, so a
/// delay is never read as a size, a share or a backdrop.
pub(super) const ID_TRAY_DELAY_BASE: u16 = 1455;
pub(super) const ID_TRAY_REHOVER_DELAY_BASE: u16 = 1470;
pub(super) const ID_TRAY_SETTLING_DELAY_BASE: u16 = 1485;
/// The one row of `Timing` that is a switch rather than a delay or a key: whether the
/// keyboard driving Explorer holds a parked pointer back instead of sharing the screen
/// with it. It sits past every range the app hands out up to the tick menu, so it is never
/// read as one of their items.
pub(super) const ID_TRAY_PRIORITIZE_KEYBOARD: u16 = 1515;
/// The delays those three submenus offer, in the order they list them: no wait at all
/// at the top, where the app starts, and a whole second at the bottom. A delay a
/// hand-edited `config.ini` holds that is not one of these is shown with nothing checked
/// rather than rounded to the nearest.
pub(super) const TIMING_DELAY_CHOICES_MS: [u64; 15] = [
    0, 25, 50, 100, 150, 200, 250, 300, 400, 500, 600, 700, 800, 900, 1000,
];
/// The `Performance → Explorer Poll` submenu: one command per rate the loop may run at, in the
/// order it lists them. It is the app's own rate rather than a delay a hover waits out,
/// and its range sits past every other range the app hands out so that a tick is never
/// read as a delay, a size or a share.
pub(super) const ID_TRAY_TICK_BASE: u16 = 1500;
/// The ticks the `Tick` submenu offers, in the order it lists them: one system tick at
/// the top, where the app starts, and five of them at the bottom.
///
/// Whole system ticks rather than round numbers, because that is what a wait is
/// honoured in — the loop wakes on the system's clock, so a number that falls between
/// two of its ticks spends the same time as the one below it and reads as a step that
/// changed nothing. A tick a hand-edited `config.ini` holds that is not one of these is
/// shown with nothing checked rather than rounded to the nearest.
pub(super) const TICK_CHOICES_MS: [u64; 5] = [15, 31, 47, 63, 78];
pub(super) const ID_TRAY_OPEN_CONFIG: u16 = 1040;
/// The two rows inside the `System` submenu, beside the one that opens the config file:
/// the first puts every setting back at what this build recommends and leaves the extension
/// lists alone, the second puts the lists back and leaves every other setting alone.
///
/// They sit between every range the submenus share and the `Codecs` commands above them, so
/// neither can be read as a click on one of those.
pub(super) const ID_TRAY_RESET_SETTINGS: u16 = 1520;
pub(super) const ID_TRAY_RESET_LISTS: u16 = 1521;
/// The `System → Check for Updates` row: a check asked for now, past the
/// once-an-hour one an opening of the menu makes, with what the check found
/// said in a dialog of its own where there is nothing to put on.
///
/// It is a switch of a question rather than a level of a list, so it takes an
/// id of its own — the one the reset rows beside it leave free, where a row of
/// the same submenu belongs and no range the app hands out reaches.
pub(super) const ID_TRAY_CHECK_UPDATES: u16 = 1522;
/// The rows of the `Codecs` submenu, numbered as one list across its three groups: only the
/// rows this machine is missing and has a page for are given an id at all, and this is where
/// those ids begin. The range is wider than the list is long, so that a row added to any of
/// the three groups is still inside it.
pub(super) const ID_TRAY_CODEC_BASE: u16 = 1600;
pub(super) const CODEC_COMMANDS: u16 = 64;
/// The row above `Run at Startup`, which is in the menu only while a newer release is
/// waiting: it puts the installer `updates` fetched on, and the app ends itself as the
/// installer takes over rather than being the copy that has to be terminated.
pub(super) const ID_TRAY_UPDATE: u16 = 1007;
/// The `Scaling → Image` submenu: one command per share a picture is
/// drawn at, in the order it lists them — the first of the three bitmap
/// submenus, each with a range of its own so a click on one is never read
/// as a click on the other. The shares are of two bases, the display's
/// fitted size and the file's own, and the range is in the stretch past
/// the `Audio` one, which ends at 1534 — the first run of twelve the
/// stretch has, with the reset rows and the toggles owning the ids below
/// it — so a click on a share is never read as a row of another
/// submenu's.
pub(super) const ID_TRAY_SCALE_BASE: u16 = 1536;
/// `Vector`, the second of them, in the range the `Scaling` submenus share: a
/// drawing is asked for a share of the display rather than for a share of a size the file
/// asks for, and both halves of the kind — a document the browser draws and a metafile the
/// drawing layer replays — are asked with this one setting.
pub(super) const ID_TRAY_VECTOR_SCALE_BASE: u16 = 1400;
/// The `Ebook` and `Document` submenus, each listing the same shares:
/// they sit in the slack the `Avoid` items leave, so a share of the display is never
/// read as a way of avoiding the item a preview is about.
pub(super) const ID_TRAY_EBOOK_SCALE_BASE: u16 = 1405;
pub(super) const ID_TRAY_DOCUMENT_SCALE_BASE: u16 = 1415;
/// `Font`, in the range after the Document one: a specimen is drawn at a share of
/// the display the same way a document is.
pub(super) const ID_TRAY_FONT_SCALE_BASE: u16 = 1420;
/// `Video`, the `Image` submenu's twin below it: it lists the same
/// shares, so the two share a table and a builder, and it is a range of
/// its own because a click on a video's scale is never a click on a
/// picture's. It sits in the stretch past the image's own range, the
/// second run of twelve the `Audio` scale's end at 1534 leaves.
pub(super) const ID_TRAY_VIDEO_SCALE_BASE: u16 = 1550;
/// `Animated`, the third of them, in the range after the video
/// one: an animated picture is a bitmap like the two above it, so it
/// lists the same shares through the same builder, and the range is its
/// own because what moves has a size apart from what does not.
pub(super) const ID_TRAY_ANIMATED_SCALE_BASE: u16 = 1564;
/// `Design`, in the range after the font one: a design document is previewed
/// from a picture the file keeps of the whole of itself, so it is asked for a share of the
/// display the way a page is rather than for a share of its own size.
pub(super) const ID_TRAY_DESIGN_SCALE_BASE: u16 = 1445;
/// `Text`, in the range after the design one: a text page is measured against a
/// share of the display rather than drawn at a share of a size of its own, so it is asked the
/// same question a page is — and the share is its own because what a page of text is given
/// and what a drawing is given are not the same answer.
pub(super) const ID_TRAY_TEXT_SCALE_BASE: u16 = 1450;
/// The shares of the display every display-share submenu offers, in the order it lists
/// them: the whole room a document can be given at the top, then the shares of it a
/// document is asked for below. What differs between the settings is where they start —
/// `50`, half the display, for a font specimen, and `Fit to Screen` for a drawing, a page
/// and a design document.
pub(super) const DOCUMENT_SCALE_CHOICES: [PreviewScale; 5] = [
    PreviewScale::FitToScreen,
    PreviewScale::Percent(75),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
    PreviewScale::Percent(10),
];
/// The shares the `Image`, `Video` and `Animated` submenus offer, in
/// the order they list them: the shares of the display's fitted size at the
/// top — the whole room a fit takes, then the room reduced to a share of
/// it — and the shares of a bitmap's own size below, which are the ones
/// that mean something for a picture a video's first frame and a moving
/// frame are. Nothing is marked as the default here — the default is passed
/// to the labels rather than written into the table, because which share a
/// setting starts at is the setting's own business. The two groups are one
/// table because they are one question — how large a bitmap is drawn — with
/// two bases, and the submenu says which is which around the items
/// themselves (see `append_bitmap_scale_menu`).
pub(super) const BITMAP_SCALE_CHOICES: [PreviewScale; 12] = [
    PreviewScale::FitToScreen,
    PreviewScale::FitToScreenReduced(75),
    PreviewScale::FitToScreenReduced(50),
    PreviewScale::FitToScreenReduced(25),
    PreviewScale::FitToScreenReduced(10),
    PreviewScale::Percent(400),
    PreviewScale::Percent(300),
    PreviewScale::Percent(200),
    PreviewScale::Percent(150),
    PreviewScale::Percent(100),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
];
/// The shares the `Audio` submenu offers, in the order it lists
/// them: a sound's card is laid out over a share of the display rather than
/// drawn at a share of a size of its own, so the percentages are the ones
/// that mean something for one. Nothing is marked as the default here — the
/// default is passed to the labels rather than written into the table,
/// because which share a setting starts at is the setting's own business.
pub(super) const AUDIO_SCALE_CHOICES: [PreviewScale; 6] = [
    PreviewScale::Percent(25),
    PreviewScale::Percent(20),
    PreviewScale::Percent(15),
    PreviewScale::Percent(10),
    PreviewScale::Percent(7),
    PreviewScale::Percent(5),
];
pub(super) const ID_TRAY_THEME_LIGHT: u16 = 1050; // Atom One Light
pub(super) const ID_TRAY_THEME_DARK: u16 = 1051; // One Dark Pro
pub(super) const ID_TRAY_MARKDOWN_RENDERED: u16 = 1052; // Rendered document
pub(super) const ID_TRAY_MARKDOWN_SOURCE: u16 = 1053; // Highlighted Markdown source
/// The `Text Preview → Render HTML` row: whether a page of HTML is drawn by the browser
/// engine rather than shown as its markup.
///
/// The id is the one the text preview's `Full Mode` item carried, which is a row of this
/// submenu that has gone (see `ID_TRAY_FONT_BACKGROUND_BASE` for why that block was moved):
/// the setting is the same question `Full Mode` was asked and one this app still asks.
pub(super) const ID_TRAY_RENDER_HTML: u16 = 1058;
/// The `Preview Types` submenu, one command per kind of preview.
pub(super) const ID_TRAY_TYPE_IMAGES: u16 = 1062;
pub(super) const ID_TRAY_TYPE_VIDEOS: u16 = 1063;
pub(super) const ID_TRAY_TYPE_TEXT: u16 = 1064;
/// The `Ebook` gate: the pages this app reads and draws itself, which is every PDF.
pub(super) const ID_TRAY_TYPE_EBOOK: u16 = 1065;
pub(super) const ID_TRAY_TYPE_ARCHIVES: u16 = 1066;
/// The `Document` gate: the pages drawn for a document, whether the application that owns
/// the format drew one or an installed render engine did — the one switch both halves of the
/// kind answer to (see `PreviewType`).
pub(super) const ID_TRAY_TYPE_DOCUMENT: u16 = 1067;
/// The `Fonts` gate beside it, under the same `Preview Types` submenu.
pub(super) const ID_TRAY_TYPE_FONTS: u16 = 1070;
/// The `Design` gate beside those, under the same submenu.
pub(super) const ID_TRAY_TYPE_DESIGN: u16 = 1071;
/// The `Vector` gate, for the drawings that are not pictures: SVG documents, which are the
/// kind SVG documents have always had — the id is the one this gate carried under that name
/// — and the metafiles and encapsulated PostScript files the same kind grew to hold.
pub(super) const ID_TRAY_TYPE_VECTOR: u16 = 1069;
/// The `Audio` gate: the sounds this app plays, which are the one kind of preview that is
/// heard rather than looked at — and the reason `Volume` has two halves (see
/// `ID_TRAY_AUDIO_VOLUME_BASE`).
///
/// Its id sits in the slack the font sizes leave rather than beside the other gates: 1072 is
/// where the text preview's own `100%` begins, and the block of gates there is two sizes wide.
pub(super) const ID_TRAY_TYPE_AUDIO: u16 = 1098;
/// The `Cache` submenu: one command per size it offers, in the order it lists
/// them, for each of the four caches it sizes. They start past the range the `theme`
/// folder's own items occupy (see `ID_TRAY_THEME_CUSTOM_BASE`).
pub(super) const ID_TRAY_IMAGE_CACHE_BASE: u16 = 1300;
/// The `Cache → Document` sizes: how much of what an engine drew is kept between hovers. The
/// pages are files under the temp folder rather than memory, which is what makes it one of the
/// two caches a size is measured in bytes of something on disk.
pub(super) const ID_TRAY_DOCUMENT_CACHE_BASE: u16 = 1320;
/// The `Cache → Image (Disk)` sizes: how much of what the image converter developed is kept
/// between hovers. The other cache whose size is bytes on disk — pictures of its own, in a
/// folder beside the documents' pages rather than a share of them (see `document_cache`).
pub(super) const ID_TRAY_IMAGE_DISK_CACHE_BASE: u16 = 1340;
/// The `Cache → General (Disk)` sizes: how much of what a preview asked FFmpeg to write is kept
/// between hovers. It holds the small files a film's own subtitle tracks were copied into, with
/// the container's fonts dumped beside them (see `subtitle_files`), which is what a hover of an
/// embedded-subtitle film draws from instead of streaming the whole film before its first frame.
pub(super) const ID_TRAY_GENERAL_DISK_CACHE_BASE: u16 = 1360;
/// The `Performance → Decode Budget` submenu: one command per ceiling it offers, in
/// the order it lists them. It sits in the slack between the `Cache` sizes and the
/// document scale's own range.
pub(super) const ID_TRAY_DECODE_BUDGET_BASE: u16 = 1380;
/// The sizes the `Cache` submenu offers, in megabytes, largest first — `2 GB` at
/// the top and a cache that holds nothing at the bottom — and the whole range the
/// settings allow, so a size a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing checked rather than rounded to one of them.
pub(super) const CACHE_SIZE_CHOICES_MB: [u32; 9] = [2048, 1024, 512, 256, 128, 64, 32, 16, 0];
/// The ceilings the `Decode Budget` submenu offers, in gigabytes, largest first. It is
/// what one hover may decode or read for rather than what is kept, so the range starts
/// far above any file someone meant to hover and ends at the smallest ceiling a large
/// picture still fits in. A value a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing marked rather than rounded to one of them.
pub(super) const DECODE_BUDGET_CHOICES_GB: [f32; 6] = [16.0, 8.0, 4.0, 2.0, 1.0, 0.5];
pub(super) const ID_TRAY_FONT_100: u16 = 1072;
pub(super) const ID_TRAY_FONT_125: u16 = 1073;
pub(super) const ID_TRAY_FONT_150: u16 = 1074;
pub(super) const ID_TRAY_FONT_175: u16 = 1075;
pub(super) const ID_TRAY_FONT_200: u16 = 1076;
pub(super) const ID_TRAY_FONT_250: u16 = 1077;
pub(super) const ID_TRAY_FONT_300: u16 = 1078;
pub(super) const ID_TRAY_FONT_400: u16 = 1079;
pub(super) const ID_TRAY_FONT_90: u16 = 1080;
pub(super) const ID_TRAY_FONT_80: u16 = 1081;
pub(super) const ID_TRAY_FONT_70: u16 = 1082;
/// `110%` is listed between `125%` and `100%` and carries the one id the font sizes
/// have left: the engine-idle ranges take 1083 up to 1096, and the `theme` folder's own
/// items begin at 1100.
pub(super) const ID_TRAY_FONT_110: u16 = 1097;
/// The sizes the `Text Preview → Font Size` submenu offers, in the order it lists them —
/// largest first — with the id each size carries. A size a hand-edited `config.ini` asks
/// for that is not one of these is shown with nothing marked rather than rounded to the
/// nearest.
pub(super) const FONT_SIZE_CHOICES: [(u32, u16); 12] = [
    (400, ID_TRAY_FONT_400),
    (300, ID_TRAY_FONT_300),
    (250, ID_TRAY_FONT_250),
    (200, ID_TRAY_FONT_200),
    (175, ID_TRAY_FONT_175),
    (150, ID_TRAY_FONT_150),
    (125, ID_TRAY_FONT_125),
    (110, ID_TRAY_FONT_110),
    (100, ID_TRAY_FONT_100),
    (90, ID_TRAY_FONT_90),
    (80, ID_TRAY_FONT_80),
    (70, ID_TRAY_FONT_70),
];
/// The `Engine → Microsoft Office TTL` submenu: one command per idle time it offers, in
/// the order it lists them. The IDs the app used before this ended at 1082 and the
/// `theme` folder's items start at 1100, so this range is the slack between the two.
pub(super) const ID_TRAY_ENGINE_IDLE_BASE: u16 = 1083;
/// The `Engine → WebView2 TTL` submenu, the same shape as the Microsoft Office one and in
/// the range after it.
pub(super) const ID_TRAY_WEBVIEW_IDLE_BASE: u16 = 1090;
/// The `Engine → LibreOffice TTL` submenu, the third of them. It sits in the slack the
/// backdrop halves leave rather than in the block the other two share — 1083 to 1096 is full,
/// one font size at 1097, and the `theme` folder's items begin at 1100 — so its range is the
/// widest run left between the image backdrop's four ids at 1023 and the config row at 1040.
pub(super) const ID_TRAY_LIBREOFFICE_IDLE_BASE: u16 = 1027;
/// The idle times the three `… TTL` submenus offer, longest first — the
/// order the menus list them in, so an engine that is never let go is the topmost
/// item and one that is let go as soon as it has drawn a page is the bottom one. A
/// value a hand-edited `config.ini` asks for that is not one of these is shown with
/// nothing marked rather than rounded to the nearest.
pub(super) const ENGINE_IDLE_CHOICES: [EngineIdle; 7] = [
    EngineIdle::Indefinite,
    EngineIdle::Seconds(3600),
    EngineIdle::Seconds(1800),
    EngineIdle::Seconds(600),
    EngineIdle::Seconds(300),
    EngineIdle::Seconds(60),
    EngineIdle::Seconds(0),
];
/// The `Engine → Away Timer` submenu: one command per away time it offers, in the order it
/// lists them. Both it and the `Persistent` range below sit past every other range the app
/// hands out — the `Explorer Poll` range is the last of those and ends at 1504 — so a time is
/// never read as a poll and a toggle is never read as either.
pub(super) const ID_TRAY_AFK_TIMER_BASE: u16 = 1505;
/// The `Persistent` toggle at the top of each `… TTL` submenu, one command apiece, in the
/// order those submenus are listed: `Microsoft Office TTL`, then `LibreOffice TTL`, then
/// `WebView2 TTL`.
pub(super) const ID_TRAY_ENGINE_PERSISTENT_BASE: u16 = 1512;
/// The `Performance → Hardware Acceleration` submenu's one row: whether a video is decoded on
/// the graphics card. It sits in the slack past the three `Persistent` toggles, which end at
/// 1515, so a click on it is never read as a toggle belonging to an engine's TTL submenu.
pub(super) const ID_TRAY_VIDEO_HW_ACCEL: u16 = 1516;
/// The `Audio` submenu under `Scaling`: one command per share of the display a
/// sound's card is laid out over, in the order it lists them. It sits in
/// the stretch between the volume toggles, which end at 1528, and the
/// `Codecs` commands, which begin at 1600 — the first run in it six ids
/// wide, the reset rows and the toggles owning the ids below it — so a
/// click on a share is never read as a row of another submenu's.
pub(super) const ID_TRAY_AUDIO_SCALE_BASE: u16 = 1529;
/// The away times the `Away Timer` submenu offers, in the order it lists them: an hour at the
/// top and a quarter of a minute at the bottom, with the one that bounds an engine by
/// default in the middle. There is no `Indefinitely` here — a time that never comes round is
/// what the `Persistent` toggle beside it is for — and a value a hand-edited `config.ini`
/// asks for that is not one of these is shown with nothing marked rather than rounded to the
/// nearest, the way every other menu of this shape reads one. See `app::afk` for what the
/// time is counted against.
pub(super) const AFK_TIMER_CHOICES_SECS: [u64; 7] = [3600, 1800, 600, 300, 60, 30, 15];

/// The two command ranges one `… TTL` submenu hands out: the idle times it lists, and the
/// `Persistent` toggle above them. They are one value rather than two because they belong to
/// the same submenu and are always handed over together — a submenu is built for one engine,
/// and both of its ranges are that engine's.
pub(super) struct EngineIdleIds {
    pub(super) times: u16,
    pub(super) persistent: u16,
}

/// Where the `theme` folder's own items start: one command ID each, in the order
/// the submenu listed them. The IDs the app uses end at the `Office Engine TTL`
/// range above, so these collide with nothing.
pub(super) const ID_TRAY_THEME_CUSTOM_BASE: u16 = 1100;
/// How many files the theme submenu will list. A menu that long is unusable well
/// before this, and the cap is what keeps a folder of thousands of files from
/// running off the end of the command IDs.
pub(super) const MAX_TRAY_CUSTOM_THEMES: usize = 200;

pub(super) const TRAY_CLASS: PCWSTR = w!("RustHoverPreviewTrayClass");

pub(super) static mut TRAY_HWND: HWND = HWND(std::ptr::null_mut());
pub(super) static mut TASKBAR_CREATED: u32 = 0;

/// The custom themes the `Theme` submenu last listed, in the order it
/// listed them: a command ID carries a position, and this is what it was a
/// position in. The menu can outlive a change to the folder, so a click has to
/// select the file the item it landed on named rather than whatever is in that
/// place now.
pub(super) static TRAY_CUSTOM_THEMES: Lazy<Mutex<Vec<TextTheme>>> =
    Lazy::new(|| Mutex::new(Vec::new()));
