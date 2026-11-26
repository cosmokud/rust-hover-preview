//! Custom previews: one external program per file extension, written as `[custom:<name>]`
//! sections of `config.ini`.
//!
//! A preset is the user's own answer to "what should preview this file", and it outranks every
//! list this app ships: a `.rme` a preset names is previewed by that preset's program whatever
//! `[archive]` or `[magick]` would have said about the name (see `routing::custom_preset_for`).
//! The section is the preset's own name — `[custom:Stress Test]` — so the file reads as the
//! window lists it, and the keys under it are the five things the window edits:
//!
//! ```ini
//! [custom:Stress Test]
//! enabled=true
//! exe=C:\Program Files\AMD\RyzenMaster\bin\Stress.exe
//! extensions=rmc, rmp
//! window_mode=normal
//! scaling=fit
//! volume=35
//! ```
//!
//! **Why the section is `custom:` and not a list's name.** Every other section is one of the
//! sixteen lists, keyed by the row that reads it (`formats::lists`); a preset is not a row and
//! is not one per kind — it is a program, and there may be a hundred of them. The prefix is what
//! keeps the two apart in a file both are written to, and it is what makes a preset recognisable
//! to a hand reading the file.
//!
//! **What a preset is not.** It is not a renderer: nothing here draws anything, and the app does
//! not decode a file a preset claims. The program is started with the file, its window is what
//! the preview shows, and a pin opens that program properly rather than pinning a picture of it
//! (see `ui::preview_window::custom_preview`). So `scaling` is how large the preview is drawn in
//! the hover — a percentage of the room the display leaves beside the item, and the fit for the
//! whole of that room, which is what the tray's own menu offers a document: `100%` is the whole of
//! that room, so a 1024x1024 picture on a 1440-tall display is a 1440x1440 preview whatever program
//! draws it. The file gives the shape of the box — a picture's own, and the program's window's where
//! the file declares none — `background` is what stands behind it where the window is smaller than
//! the band, and `volume` is the level a preview of it is played at — the same three settings the
//! tray's own rows are, read from the preset rather than from the configuration.

use std::path::Path;

use configparser::ini::Ini;

use crate::formats::lists::sanitize_extension_list;

use super::app_config::AppConfig;
use super::defaults::{
    sanitize_preview_scale_percent, sanitize_volume, CONFIG_SECTION, DEFAULT_AUDIO_VOLUME,
};
use super::setting_types::PreviewScale;

/// The prefix every preset's section carries: `[custom:<name>]`.
pub const SECTION_PREFIX: &str = "custom:";

/// How many presets the file may hold, which is the window's own ceiling on its list.
pub const MAX_PRESETS: usize = 100;

/// The longest name a preset may have, in characters — a tab has to fit it, and the section
/// name is written from it.
pub const MAX_NAME_CHARS: usize = 48;

/// The settings key holding the order the window lists the presets in: their names, top first.
///
/// Their order is what decides which preset wins an extension two of them name
/// (see `preset_for`), so it is written as a key of its own — the file's sections are a map's
/// order, alphabetical, and a file read back from it would hand `.jpg` to the wrong program.
pub(crate) const PRESET_ORDER_KEY: &str = "custom_preset_order";

/// How a custom preview's target window is captured and framed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomWindowMode {
    /// Leave the target window's chrome/styles intact and capture the full window.
    #[default]
    Normal,
    /// Strip standard frame chrome and capture client area only.
    Borderless,
}

impl CustomWindowMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Borderless => "borderless",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "normal" => Some(Self::Normal),
            "borderless" => Some(Self::Borderless),
            _ => None,
        }
    }
}

/// How a custom preview captures and displays the target application's window.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomCaptureTechnique {
    /// On-monitor `WS_EX_LAYERED` + `LWA_ALPHA = 1` + DWM thumbnail (auto-falls back when unsupported).
    InvisibleLayered,
    /// On-monitor behind opaque `host` at `HWND_BOTTOM` + DWM thumbnail.
    BehindHost,
    /// Parked off-screen + DWM thumbnail.
    OffScreen,
    /// Target window placed directly at preview rect (`HWND_TOPMOST`, click-through, no host/thumbnail).
    DirectWindow,
    /// Staged at `(-10000, -10000)` until settled, then revealed on-monitor with zero flash.
    OffscreenStaged,
    /// In-process `WH_CBT` + `WH_CALLWNDPROC` hook with `DWMWA_CLOAK` + synchronous rect lock.
    #[default]
    InProcessHook,
}

impl CustomCaptureTechnique {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::InvisibleLayered => "invisible_layered",
            Self::BehindHost => "behind_host",
            Self::OffScreen => "off_screen",
            Self::DirectWindow => "direct_window",
            Self::OffscreenStaged => "offscreen_staged",
            Self::InProcessHook => "in_process_hook",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "invisible_layered" | "invisible-layered" => Some(Self::InvisibleLayered),
            "behind_host" | "behind-host" => Some(Self::BehindHost),
            "off_screen" | "off-screen" | "offscreen" => Some(Self::OffScreen),
            "direct_window" | "direct-window" => Some(Self::DirectWindow),
            "offscreen_staged" | "offscreen-staged" | "off_screen_staged" | "off-screen-staged" => {
                Some(Self::OffscreenStaged)
            }
            "in_process_hook" | "in-process-hook" | "inprocess_hook" | "inprocess-hook" => {
                Some(Self::InProcessHook)
            }
            _ => None,
        }
    }
}

/// How pinning / opening a custom preview promotes the target application.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomPromotionTechnique {
    /// Detach and promote the live preview window in-place (falls back to spawn when cold).
    #[default]
    DirectPromotion,
    /// End the preview and launch a fresh, untouched detached process for the file.
    OpenAsNewProcess,
}

impl CustomPromotionTechnique {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::DirectPromotion => "direct_promotion",
            Self::OpenAsNewProcess => "open_as_new_process",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "direct_promotion" | "direct-promotion" => Some(Self::DirectPromotion),
            "open_as_new_process" | "open-as-new-process" => Some(Self::OpenAsNewProcess),
            _ => None,
        }
    }
}

/// How far past the launched pid a custom preview may look for the window it shows.
///
/// **A program that hands its file to a process of its own is a program whose window is not the
/// one the launched id owns** — a launcher and the window it puts up, a viewer and its renderer,
/// the stub an installer writes and the program it starts. Walking the tree is what finds those
/// windows, and it is also what lets a preview hide, and on teardown end, processes the user never
/// named — so how far it may walk is the user's to say, per preset, and it is off unless asked for.
///
/// The walk itself is bounded (see `process_tree`); these are the three shapes it is asked in.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomChildScope {
    /// The launched pid alone: no tree walk, no job pids, no brokered pid — the one process this
    /// app began and nothing beside it.
    #[default]
    Disabled,
    /// The launched pid and the processes it started, one level deep.
    Children,
    /// The launched pid and everything below it, as deep as the machine's own parent links go.
    ChildrenAndGrandchildren,
}

impl CustomChildScope {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Disabled => "disabled",
            Self::Children => "children",
            Self::ChildrenAndGrandchildren => "children_and_grandchildren",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "disabled" | "off" | "none" => Some(Self::Disabled),
            "children" | "child" | "direct" => Some(Self::Children),
            "children_and_grandchildren"
            | "children-and-grandchildren"
            | "grandchildren"
            | "all"
            | "tree" => Some(Self::ChildrenAndGrandchildren),
            _ => None,
        }
    }

    /// How many `children_of` levels past the launched pid may be walked: none, one, or as deep as
    /// the tree goes.
    pub fn depth(self) -> CustomChildDepth {
        match self {
            Self::Disabled => CustomChildDepth::LaunchedOnly,
            Self::Children => CustomChildDepth::OneLevel,
            Self::ChildrenAndGrandchildren => CustomChildDepth::WholeTree,
        }
    }
}

/// The walk a scope stands for, in the terms `process_tree` and `discovery_processes` ask it in.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CustomChildDepth {
    /// `[launch.pid]` and nothing else.
    LaunchedOnly,
    /// `launch.pid` plus the pids it started directly.
    OneLevel,
    /// `launch.pid` plus every descendant, as far as the machine's parent links go.
    WholeTree,
}

/// Which of the program's windows is the picture, when it made more than one of them a picture could
/// be.
///
/// **Not the same question as [`CustomSiblingWindows`].** That row is about the windows *around* the
/// picture — the splash, the About box — and this one is about which window the picture *is*. They
/// are separate because a program can be wrong about the first and right about the second, and
/// because the wrong answer here is the one the user sees as an empty preview.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomPictureWindow {
    /// The largest picture-shaped window in the best tier, and nothing else considered. What every
    /// preset did before this row existed, and what a program whose document window is the only one
    /// of its size class gets anyway (see the settle below).
    Largest,
    /// Of the windows in the best tier whose area is within [`COMPARABLE_AREA_PERCENT`] of the
    /// largest, the one the program made **last** — once it has stopped making windows, so a helper
    /// born first cannot be mistaken for the picture.
    ///
    /// **MPV Player is why this exists, and why it is the default.** Its `mpv.exe` is built with
    /// `win32-smtc`, and that feature puts up a top-level `mpv-smtc` window titled `mpv smtc` about
    /// 200 ms before the video window — and because this app's launch hands `STARTF_USESIZE` to the
    /// program, USER32 gives *that* helper the preview's own box. A titled, unowned, non-tool window
    /// at 1976x1112 against the video window's 1936x1119 is the same tier, the same visibility and
    /// 1.4 % more area, so it won as the picture: a white, frozen rectangle where the film should be.
    /// SVP's `mpv.exe` is the same version without the feature, which is why one build previewed and
    /// the other did not.
    ///
    /// **What it costs a program that is not MPV Player is nothing that was measured.** A program
    /// with one picture-shaped window — SVP's `mpv.exe`, Windows 11's Paint, PhotoCraft, Dragon —
    /// has already answered the question, so it is adopted without waiting: the settle is paid only
    /// where two windows are of the same size class and the newer one is what this row asks for.
    /// VLC's three windows were measured too: `largest` and this pick the same one.
    #[default]
    NewestComparable,
}

impl CustomPictureWindow {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Largest => "largest",
            Self::NewestComparable => "newest_comparable",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "largest" | "biggest" | "area" => Some(Self::Largest),
            "newest_comparable" | "newest-comparable" | "newest" | "latest" => {
                Some(Self::NewestComparable)
            }
            _ => None,
        }
    }

    /// Whether the choice prefers the window made last; the ranking asks this (see
    /// `select_best_candidate_by`).
    pub fn prefers_the_newest(self) -> bool {
        self == Self::NewestComparable
    }
}

/// How much of the largest candidate's area a window must cover to count as a picture of the same
/// document rather than a different thing altogether.
///
/// The margin exists so that the row cannot turn "the document window owns the box" into "a small
/// toast owns the box": only windows that are the same *size class* as the largest compete on
/// creation order. Measured against the case that made the row (MPV Player's 1976x1112 helper
/// against its 1936x1119 film: 98.6 %), and left clear of the furniture those programs also make —
/// a splash or a tooltip is a few per cent of the box, not sixty.
pub const COMPARABLE_AREA_PERCENT: i64 = 60;

/// What a custom preview does with the program's *other* windows — the splash, the About box, the
/// second document — while one of them is the picture.
///
/// **Not a child-process question.** Those windows belong to the same process as the one being
/// shown, so nothing here reads the process tree; it is the suppression half of keeping a
/// program's furniture off the listing, and it is a setting of its own rather than a step of
/// [`CustomChildScope`] for exactly that reason.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomSiblingWindows {
    /// Neither concealed nor hidden: the preview shows its window and leaves the rest alone.
    Disabled,
    /// `DWMWA_CLOAK` on every candidate before the winner is known, so nothing of the program is
    /// composited while discovery runs. Nothing is hidden.
    Cloak,
    /// Cloaked during discovery and then hidden once the winner is adopted: a visible companion is
    /// `SW_HIDE`d, and a modal `#32770` dialog is closed with `WM_CLOSE` instead of hidden.
    #[default]
    CloakAndHide,
}

impl CustomSiblingWindows {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Disabled => "disabled",
            Self::Cloak => "cloak",
            Self::CloakAndHide => "cloak_and_hide",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "disabled" | "off" | "none" => Some(Self::Disabled),
            "cloak" | "cloak_only" | "cloak-only" => Some(Self::Cloak),
            "cloak_and_hide" | "cloak-and-hide" | "hide" | "on" => Some(Self::CloakAndHide),
            _ => None,
        }
    }

    /// Whether candidates are cloaked before the winner is known.
    pub fn cloaks(self) -> bool {
        self != Self::Disabled
    }

    /// Whether the program's other visible windows are hidden once the winner is adopted.
    pub fn hides(self) -> bool {
        self == Self::CloakAndHide
    }
}

/// How a custom preview positions the target window relative to `Tray -> Placement`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomPlacement {
    /// Default: trust the program to honor `Tray -> Placement` coordinates.
    #[default]
    FollowTray,
    /// Aggressively enforce `Tray -> Placement` across initial show, size constraints, and drift.
    Forced,
    /// Inject `custom_preview_hook.dll` (and dynamic attach fallback) to lock creation & tracking rects.
    Hook,
}

impl CustomPlacement {
    /// What a preset that has never had the row set reads as on this machine.
    ///
    /// **`Tray -> Placement + Hook` where the hook can be had, the plain placement where it
    /// cannot.** The hook is the placement that holds a program's window where the hover put it
    /// whatever the program itself tries — a resize, a move, a maximize, a foreground grab — and the
    /// DLL it needs ships beside this executable, so on a machine that has it there is no reason for
    /// a new preset to be the weaker of the two. Without the DLL the row reads `(.dll is missing)`
    /// and cannot be picked at all (see `placement_selectable`), which is why the fallback is the
    /// plain placement rather than an enforcement nothing would answer.
    ///
    /// **This is the default, not a migration.** A preset whose file says `placement=tray` was
    /// chosen that way by whoever wrote it, and it keeps it: only a preset with no such key at all
    /// — `Add`, and a section written before the row existed — is answered here.
    pub fn default_on_this_machine() -> Self {
        if crate::paths::hook_dll().is_some() {
            Self::Hook
        } else {
            Self::FollowTray
        }
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Self::FollowTray => "tray",
            Self::Forced => "tray_forced",
            Self::Hook => "tray_hook",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "tray" | "follow_tray" | "follow-tray" => Some(Self::FollowTray),
            "tray_forced" | "tray-forced" | "forced" => Some(Self::Forced),
            "tray_hook" | "tray-hook" | "hook" => Some(Self::Hook),
            _ => None,
        }
    }
}

/// How a custom preview is sized: either following `Tray -> Scaling` for the hovered file's kind,
/// or an explicit per-preset override.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomScaling {
    #[default]
    FollowTray,
    Override(PreviewScale),
}

impl CustomScaling {
    pub fn as_str(self) -> String {
        match self {
            Self::FollowTray => "tray".to_string(),
            Self::Override(scale) => scale.as_str(),
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "tray" | "follow tray" | "follow_tray" | "follow-tray" | "default" | "auto" => {
                Some(Self::FollowTray)
            }
            _ => PreviewScale::from_str(value).map(Self::Override),
        }
    }

    pub fn sanitized(self) -> Self {
        match self {
            Self::FollowTray => Self::FollowTray,
            Self::Override(scale) => Self::Override(match scale {
                PreviewScale::FitToScreen => PreviewScale::FitToScreen,
                PreviewScale::FitToScreenReduced(percent) => {
                    PreviewScale::FitToScreenReduced(sanitize_preview_scale_percent(percent))
                }
                PreviewScale::Percent(percent) => {
                    PreviewScale::Percent(sanitize_preview_scale_percent(percent))
                }
            }),
        }
    }
}

/// How the hovered file's path is spelled on the command line.
///
/// **A path and a URL are different words to a program**, and which one it wants is the program's
/// own business: `mpv` takes either, `ffplay` takes a path, and a player handed `file://?/E%3A/…`
/// answers that it cannot open the stream. So the spelling is a setting of the preset rather than
/// something this app guesses, and the three placeholders in `arguments` read it — `{file}` in
/// whichever spelling is picked here, `{path}` whichever-preset for the plain path and `{url}` for
/// the URL — so one command line may carry both.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomPathMethod {
    /// `E:\Videos\a b.mkv` — the path as Windows writes it, and what every preset did before this
    /// row existed.
    #[default]
    Plain,
    /// `file:///E:/Videos/a%20b.mkv` — a URL, separators turned round and the characters a URL may
    /// not carry escaped. `file://server/share/a.mkv` for a share.
    FileUrl,
    /// `E:/Videos/a b.mkv` — separators turned round and nothing escaped, for a program that reads
    /// `/` but treats `\` as an escape.
    ForwardSlashes,
    /// `E:\VIDEO~1\A~1.MKV` — the volume's own 8.3 name, for a program that splits on spaces or
    /// chokes on Unicode. Reads as `Plain` on a volume that keeps no short names.
    ShortPath,
    /// `file:///E:/Videos/a b.mkv` — the URL with nothing escaped, for the program that takes a URL
    /// and then hands `%20` on to the filesystem as those four characters.
    FileUrlLiteral,
}

impl CustomPathMethod {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Plain => "plain",
            Self::FileUrl => "file_url",
            Self::ForwardSlashes => "forward_slashes",
            Self::ShortPath => "short_path",
            Self::FileUrlLiteral => "file_url_literal",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "plain" | "path" | "windows" | "native" | "" => Some(Self::Plain),
            "file_url" | "file-url" | "url" | "file_uri" | "uri" => Some(Self::FileUrl),
            "forward_slashes" | "forward-slashes" | "slashes" | "unix" => {
                Some(Self::ForwardSlashes)
            }
            "short_path" | "short-path" | "short" | "8.3" | "dos" => Some(Self::ShortPath),
            "file_url_literal" | "file-url-literal" | "url_literal" | "file_url_raw" => {
                Some(Self::FileUrlLiteral)
            }
            _ => None,
        }
    }

    /// The file in this method's spelling: what `{file}` stands for, and what is appended where no
    /// placeholder is written at all.
    ///
    /// **Every method reads the plain path, the verbatim prefix off.** The hovered path arrives from
    /// the Shell as `\\?\E:\…`, and that form is not a name any program resolves: VLC handed it
    /// answered `VLC is unable to open the MRL 'file://?/E%3A/…'`, a URL it reconstructed out of a
    /// prefix that was never part of one — so a preset with no `Path Method` of its own was broken
    /// by the spelling this app used to hand over unread.
    pub fn spell(self, path: &Path) -> String {
        match self {
            Self::Plain => crate::paths::plain_path(path),
            Self::ForwardSlashes => crate::paths::forward_slashes(path),
            Self::ShortPath => {
                crate::paths::short_path(path).unwrap_or_else(|| crate::paths::plain_path(path))
            }
            Self::FileUrl => {
                crate::paths::file_url(path).unwrap_or_else(|| crate::paths::plain_path(path))
            }
            Self::FileUrlLiteral => crate::paths::file_url_literal(path)
                .unwrap_or_else(|| crate::paths::plain_path(path)),
        }
    }
}

/// The three words a token may be written as, and what each one stands for here.
///
/// **A placeholder is this app's, not the program's**, so it is spelled in braces: `%s` is a
/// printf field to one program and a substitution of its own to another (`vlc`'s `%` expansion is
/// why `%S` cannot be used), `%1` is a file association's own parameter, and `$F`/`%f` belong to
/// editors and scripts. Nothing in this app's own command line ever carries a brace.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PathSpellings {
    /// `{file}` — the preset's own `Path Method`.
    pub file: String,
    /// `{path}` — the plain Windows path, whatever the method is.
    pub path: String,
    /// `{url}` — a `file://` URL, whatever the method is.
    pub url: String,
}

impl PathSpellings {
    pub fn of(path: &Path, method: CustomPathMethod) -> Self {
        Self {
            file: method.spell(path),
            path: crate::paths::plain_path(path),
            url: crate::paths::file_url(path).unwrap_or_else(|| crate::paths::plain_path(path)),
        }
    }

    /// Which of the three a given placeholder is, if it is one at all.
    fn of_token(token: &str) -> Option<&str> {
        match token {
            "{file}" => Some("{file}"),
            "{path}" => Some("{path}"),
            "{url}" => Some("{url}"),
            _ => None,
        }
    }

    fn stand_in(&self, token: &str) -> &str {
        match token {
            "{path}" => &self.path,
            "{url}" => &self.url,
            _ => &self.file,
        }
    }
}

/// One token with every placeholder it carries replaced **in a single pass**, and whether it
/// carried one at all.
///
/// One pass rather than a `replace` per word: a file named `{url}.mkv` is a file, and a path that
/// was substituted into and then read again would be a second substitution of somebody's name.
fn expand_placeholders(token: &str, spellings: &PathSpellings) -> (String, bool) {
    let mut out = String::with_capacity(token.len());
    let mut rest = token;
    let mut substituted = false;
    while let Some(open) = rest.find('{') {
        let Some(close) = rest[open..].find('}').map(|close| open + close) else {
            break;
        };
        let candidate = &rest[open..=close];
        let Some(known) = PathSpellings::of_token(candidate) else {
            // Not one of ours: kept as it is, and the scan goes on past it rather than stopping,
            // so a token holding one of ours after a brace it does not know still gets it.
            out.push_str(&rest[..=close]);
            rest = &rest[close + 1..];
            continue;
        };
        out.push_str(&rest[..open]);
        out.push_str(spellings.stand_in(known));
        substituted = true;
        rest = &rest[close + 1..];
    }
    out.push_str(rest);
    (out, substituted)
}

/// One custom preview: the program, the names it claims, and the settings a preview of it
/// is drawn with.
///
/// It is a plain value rather than something with a lock of its own: the configuration holds the
/// list, a read clones the one preset it found, and every reader asks `AppConfig` for it — so a
/// preview on a worker thread never holds the configuration's lock while a program starts.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomPreset {
    /// The name the window lists it by, and the half of its section name after the prefix.
    pub name: String,
    /// Whether this preset is active; disabled presets are skipped by `preset_for`.
    pub enabled: bool,
    /// The program, as the user typed or browsed to it. Empty is a preset with nothing to run,
    /// which is a preset that claims nothing until one is given.
    pub exe: String,
    /// Optional command-line arguments passed before (or around `{file}` instead of trailing) the hovered file path.
    pub arguments: String,
    /// How the path is spelled where `arguments` stands it: `{file}` and the appended file read
    /// this, `{path}` and `{url}` do not (see [`CustomPathMethod`]).
    pub path_method: CustomPathMethod,
    /// The extensions it claims, each a bare extension without its dot, lowercased — the same
    /// shape `formats::lists` keeps its own entries in.
    pub extensions: Vec<String>,
    /// Whether the target window keeps its normal frame or is captured borderless.
    pub window_mode: CustomWindowMode,
    /// Capture/display technique used for this preset's target window.
    pub capture_technique: CustomCaptureTechnique,
    /// How far past the launched pid discovery may look for that window.
    pub child_scope: CustomChildScope,
    /// Whether a program begun by an activator or a stub — a packaged app, a Squirrel/NSIS
    /// launcher — may be found by its image name or its install folder after the launched tree
    /// has shown nothing.
    pub broker_fallback: bool,
    /// What happens to the program's other windows while one of them is the picture.
    pub sibling_windows: CustomSiblingWindows,
    /// Which of the program's windows is the picture: the largest, or the last one it made of the
    /// same size class (see [`CustomPictureWindow`]).
    pub picture_window: CustomPictureWindow,
    /// Promotion technique used when pinning / opening this preset's file.
    pub promotion_technique: CustomPromotionTechnique,
    /// Placement strategy used for this preset's preview window.
    pub placement: CustomPlacement,
    /// How large the preview is drawn in the hover: following `Tray -> Scaling` by default, or an
    /// explicit per-preset share of the room.
    pub scaling: CustomScaling,
    /// Whether to override per-process audio volume via WASAPI (`false` leaves app/Windows Volume Mixer in control).
    pub volume_enabled: bool,
    /// The level a preview of it is played at, a percentage.
    pub volume: u32,
}

impl Default for CustomPreset {
    fn default() -> Self {
        Self {
            name: String::new(),
            enabled: true,
            exe: String::new(),
            arguments: String::new(),
            path_method: CustomPathMethod::default(),
            extensions: Vec::new(),
            window_mode: CustomWindowMode::default(),
            capture_technique: CustomCaptureTechnique::default(),
            child_scope: CustomChildScope::default(),
            broker_fallback: false,
            sibling_windows: CustomSiblingWindows::default(),
            picture_window: CustomPictureWindow::default(),
            promotion_technique: CustomPromotionTechnique::default(),
            placement: CustomPlacement::default(),
            scaling: CustomScaling::default(),
            volume_enabled: false,
            volume: DEFAULT_AUDIO_VOLUME,
        }
    }
}

impl CustomPreset {
    /// A preset of this name and nothing else: what the window's own `Add` makes.
    pub fn named(name: &str) -> Self {
        Self {
            name: sanitize_name(name).unwrap_or_default(),
            ..Self::default()
        }
    }

    /// The section this preset is written under.
    pub fn section(&self) -> String {
        section_of(&self.name)
    }

    /// The extensions as the window's one input shows them: `jpg, png, gif`.
    pub fn extensions_text(&self) -> String {
        self.extensions.join(", ")
    }

    /// The program, where the preset has one that could be started.
    pub fn exe_path(&self) -> Option<&str> {
        let exe = self.exe.trim();
        (!exe.is_empty()).then_some(exe)
    }

    /// Whether this preset is one a preview can be made of: a program, and at least one name it
    /// claims. A half-filled preset is one the window keeps and nothing previews.
    pub fn is_usable(&self) -> bool {
        self.exe_path().is_some() && !self.extensions.is_empty() && !self.name.is_empty()
    }

    /// Whether this preset claims `path`, by the extension the file ends in.
    ///
    /// The extension alone, lowercased, because that is what the window's input takes and what
    /// the file lists are read by: a preset naming `psd` claims `art.PSD` and does not claim
    /// `art.psd.bak`. A name with no extension is never claimed.
    pub fn claims(&self, path: &Path) -> bool {
        let Some(extension) = path.extension().and_then(|extension| extension.to_str()) else {
            return false;
        };
        let extension = extension.to_ascii_lowercase();
        self.extensions.contains(&extension)
    }

    /// The extensions as the window's input would sanitize them, which is what a save puts back
    /// in the field: a comma-separated list of anything is reduced to the names it may hold.
    pub fn set_extensions_text(&mut self, text: &str) {
        self.extensions = sanitize_extension_list(text);
    }

    /// The preset reduced to what the file and the preview may hold: a name that can be a
    /// section, a program path that is one line, the extensions read through the list's own
    /// sanitiser, and the three settings inside their bounds.
    pub fn sanitized(mut self) -> Option<Self> {
        self.name = sanitize_name(&self.name)?;
        self.exe = sanitize_exe(&self.exe);
        self.arguments = sanitize_arguments(&self.arguments);
        self.extensions = sanitize_extension_list(&self.extensions.join(","));
        self.scaling = self.scaling.sanitized();
        self.volume = sanitize_volume(self.volume);
        Some(self)
    }
}

/// The section name a preset's name is written under.
///
/// **Lowercased, because the file's reader lowercases every section it parses** — so a section
/// written as `[custom:Stress Test]` is read back as `custom:stress test`, and a file this app
/// wrote would not be the file it reads. The name as the user typed it is the `name` key inside
/// the section, which is what the window shows and what a save writes back.
pub fn section_of(name: &str) -> String {
    format!("{SECTION_PREFIX}{}", name.to_lowercase())
}

/// A preset name reduced to what may be a section of the file and a tab of the window: letters,
/// digits, spaces, `-`, `_` and `.`, trimmed, and no longer than [`MAX_NAME_CHARS`].
///
/// `None` for a name that holds nothing once those are gone, which is a name the window refuses
/// rather than writes as an empty section.
pub fn sanitize_name(raw: &str) -> Option<String> {
    let cleaned: String = raw
        .trim()
        .chars()
        .filter(|character| {
            character.is_alphanumeric() || matches!(character, ' ' | '-' | '_' | '.')
        })
        .take(MAX_NAME_CHARS)
        .collect();
    let cleaned = cleaned.trim().trim_matches('.').trim().to_string();
    (!cleaned.is_empty()).then_some(cleaned)
}

/// A program path reduced to what one `key=value` line may hold: trimmed, without the quotes a
/// user copies around a path with spaces in it, and without a newline a paste may bring.
fn sanitize_exe(raw: &str) -> String {
    sanitize_program_input(raw)
}

/// The program field as a paste hands it over: **Explorer's `Copy as path` (Ctrl+Shift+C) wraps the
/// path in `"`**, so `"C:\Program Files\IrfanView\i_view64.exe"` is the program `C:\Program
/// Files\IrfanView\i_view64.exe`, and a multi-line copy's line break is nobody's path.
pub fn sanitize_program_input(raw: &str) -> String {
    raw.replace(['\r', '\n'], "")
        .trim()
        .trim_matches('"')
        .trim()
        .to_string()
}

/// Strip line breaks from the Arguments field while preserving spaces, quotes, flags, and `=`.
pub fn sanitize_arguments(raw: &str) -> String {
    raw.replace(['\r', '\n'], "")
}

/// Reversible INI value escaping so `configparser` 1.0 never truncates at `;`/`#` or chokes on `[`/`]`.
pub(crate) fn encode_ini_arguments(raw: &str) -> String {
    let mut out = String::with_capacity(raw.len());
    for ch in raw.chars() {
        match ch {
            '%' => out.push_str("%25"),
            '#' => out.push_str("%23"),
            ';' => out.push_str("%3B"),
            '[' => out.push_str("%5B"),
            ']' => out.push_str("%5D"),
            other => out.push(other),
        }
    }
    out
}

/// Inverse of [`encode_ini_arguments`], leaving unencoded literal `%` (such as hand-edited `100%`) untouched.
pub(crate) fn decode_ini_arguments(raw: &str) -> String {
    let mut out = String::with_capacity(raw.len());
    let mut rest = raw;
    while let Some(ch) = rest.chars().next() {
        if let Some((decoded, tail)) = rest
            .strip_prefix("%25")
            .map(|t| ('%', t))
            .or_else(|| rest.strip_prefix("%23").map(|t| ('#', t)))
            .or_else(|| rest.strip_prefix("%3B").map(|t| (';', t)))
            .or_else(|| rest.strip_prefix("%5B").map(|t| ('[', t)))
            .or_else(|| rest.strip_prefix("%5D").map(|t| (']', t)))
        {
            out.push(decoded);
            rest = tail;
        } else {
            out.push(ch);
            rest = &rest[ch.len_utf8()..];
        }
    }
    out
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ArgQuote {
    None,
    Double,
    Single,
}

fn is_single_outer_quoted_span(s: &str) -> bool {
    let bytes = s.as_bytes();
    if bytes.len() < 2 {
        return false;
    }
    let outer = bytes[0];
    if (outer != b'"' && outer != b'\'') || bytes[bytes.len() - 1] != outer {
        return false;
    }
    if outer == b'\'' {
        return !bytes[1..bytes.len() - 1].contains(&b'\'');
    }

    let mut i = 1usize;
    while i < bytes.len() {
        if bytes[i] == b'\\' {
            let mut slashes = 0usize;
            while i < bytes.len() && bytes[i] == b'\\' {
                slashes += 1;
                i += 1;
            }
            if i < bytes.len() && bytes[i] == b'"' {
                if slashes.is_multiple_of(2) {
                    return i == bytes.len() - 1;
                }
                i += 1;
            }
        } else if bytes[i] == b'"' {
            return i == bytes.len() - 1;
        } else {
            i += 1;
        }
    }
    false
}

/// Parse an Arguments string into individual argv tokens following Windows `CommandLineToArgvW`
/// backslash-before-quote rules plus single-quote grouping outside double quotes, unwrapping an
/// accidental single outer wrapper quote around multiple flags/args while preserving real quoted values.
///
/// **A quoted span is one argument unless what is inside it opens as a flag.** The unwrap exists for
/// a run of flags copied from a docs page — `"--fullscreen --silent"` — and it is the one thing the
/// CRT does not do. Taken one step too far it is a *path* torn into its words: VLC handed
/// `"file:///E:/Videos/…/2026 - 02 - Batch/Sample Show/82.mkv"` asked for six media files
/// (`…/2026`, `-`, `02`, …) and played none of them, which is the failure this predicate is written
/// against — the old one unwrapped a span on *any* `-`-led word inside it, and a file name holding
/// ` - ` has one. So the unwrap is taken only where the first word inside the span is a flag in its
/// own right, or one of this app's three placeholders; every later word comes along with it.
pub fn parse_arguments(raw: &str) -> Vec<String> {
    let trimmed = raw.trim();
    if is_single_outer_quoted_span(trimmed) {
        let raw_tokens = parse_arguments_raw(trimmed);
        if raw_tokens.len() == 1 {
            let inner_tokens = parse_arguments_raw(&raw_tokens[0]);
            if looks_like_a_run_of_flags(&inner_tokens) {
                return inner_tokens;
            }
        }
    }
    parse_arguments_raw(trimmed)
}

/// Whether a quoted span holds an argument *list* rather than one argument.
///
/// The first word decides: a flag opens a list, and so does a placeholder, while a value — a path,
/// a file name, a URL — opens one argument that happens to have spaces in it (see
/// `parse_arguments`).
fn looks_like_a_run_of_flags(tokens: &[String]) -> bool {
    let Some(first) = tokens.first() else {
        return false;
    };
    tokens.len() >= 2
        && (first.starts_with("{file}")
            || first.starts_with("{path}")
            || first.starts_with("{url}")
            || is_a_flag(first))
}

/// A word that is a flag rather than the head of a path: `-`/`/`-led, carrying no `=` (a `--k=v`
/// pair is one argument of its own) and no separator after the lead, which is what keeps `E:/…` and
/// `/mnt/x` out — neither of them opens a run of flags.
fn is_a_flag(token: &str) -> bool {
    if token.contains('=') {
        return false;
    }
    let Some(rest) = token.strip_prefix('-').or_else(|| token.strip_prefix('/')) else {
        return false;
    };
    !rest.is_empty() && !rest.contains(['/', '\\', ':'])
}

fn parse_arguments_raw(raw: &str) -> Vec<String> {
    let mut args = Vec::new();
    let mut current = String::new();
    let mut in_token = false;
    let mut quote = ArgQuote::None;
    let mut chars = raw.chars().peekable();

    while let Some(ch) = chars.next() {
        if quote == ArgQuote::Single {
            if ch == '\'' {
                quote = ArgQuote::None;
            } else {
                current.push(ch);
            }
            continue;
        }

        if ch == '\\' {
            in_token = true;
            let mut slashes = 1usize;
            while chars.peek() == Some(&'\\') {
                chars.next();
                slashes += 1;
            }
            if chars.peek() == Some(&'"') {
                chars.next();
                for _ in 0..(slashes / 2) {
                    current.push('\\');
                }
                if slashes % 2 == 1 {
                    current.push('"');
                } else {
                    quote = match quote {
                        ArgQuote::None => ArgQuote::Double,
                        ArgQuote::Double => ArgQuote::None,
                        ArgQuote::Single => ArgQuote::Single,
                    };
                }
            } else {
                for _ in 0..slashes {
                    current.push('\\');
                }
            }
        } else if ch == '"' {
            in_token = true;
            quote = match quote {
                ArgQuote::None => ArgQuote::Double,
                ArgQuote::Double => ArgQuote::None,
                ArgQuote::Single => ArgQuote::Single,
            };
        } else if ch == '\'' && quote == ArgQuote::None {
            in_token = true;
            quote = ArgQuote::Single;
        } else if quote == ArgQuote::None && matches!(ch, ' ' | '\t' | '\r' | '\n') {
            if in_token {
                args.push(std::mem::take(&mut current));
                in_token = false;
            }
        } else {
            in_token = true;
            current.push(ch);
        }
    }

    if in_token {
        args.push(current);
    }
    args
}

/// Build the argv list (excluding argv[0]) for launching a custom preview with `arguments` and
/// `file_path`.
///
/// Three placeholders stand where the file goes: `{file}` in the preset's own [`CustomPathMethod`]
/// spelling, `{path}` for the plain Windows path and `{url}` for a `file://` URL, so a preset that
/// needs both a path and a URL in one command line has both. Any of them may be named more than
/// once, and any of them may sit inside a value (`--input={file}`). Where no placeholder appears at
/// all the file is appended last, in the method's own spelling — which is what a preset with no
/// `arguments` gets.
#[cfg_attr(not(test), allow(dead_code))] // The plain spelling's own name: every caller in the app asks `build_launch_args_with`, and the tests read this one as what "plain" means.
pub fn build_launch_args(arguments: &str, file_path: &Path) -> Vec<String> {
    build_launch_args_with(arguments, file_path, CustomPathMethod::default())
}

/// [`build_launch_args`] in the spelling `method` names.
pub fn build_launch_args_with(
    arguments: &str,
    file_path: &Path,
    method: CustomPathMethod,
) -> Vec<String> {
    let spellings = PathSpellings::of(file_path, method);
    let mut args = parse_arguments(arguments);
    let mut substituted = false;
    for arg in &mut args {
        let (expanded, carried) = expand_placeholders(arg, &spellings);
        *arg = expanded;
        substituted |= carried;
    }
    if !substituted {
        args.push(spellings.file);
    }
    args
}

fn parse_enabled(value: &str) -> Option<bool> {
    match value.trim().to_ascii_lowercase().as_str() {
        "true" | "1" | "yes" | "on" => Some(true),
        "false" | "0" | "no" | "off" => Some(false),
        _ => None,
    }
}

/// Every preset written as its own section, in the order the window lists them.
///
/// The order is the list's own rather than the map's, which is what makes a file the app writes
/// and a file it reads back name the same preset where two claim the same extension.
pub fn write_all(config: &AppConfig, ini: &mut Ini) {
    for preset in config.custom_presets.iter().take(MAX_PRESETS) {
        let Some(preset) = preset.clone().sanitized() else {
            continue;
        };
        let section = preset.section();
        ini.set(&section, "name", Some(preset.name.clone()));
        ini.set(
            &section,
            "enabled",
            Some(if preset.enabled { "true" } else { "false" }.to_string()),
        );
        ini.set(&section, "exe", Some(preset.exe.clone()));
        ini.set(
            &section,
            "arguments",
            Some(encode_ini_arguments(&preset.arguments)),
        );
        ini.set(
            &section,
            "path_method",
            Some(preset.path_method.as_str().to_string()),
        );
        ini.set(&section, "extensions", Some(preset.extensions_text()));
        ini.set(
            &section,
            "window_mode",
            Some(preset.window_mode.as_str().to_string()),
        );
        ini.set(
            &section,
            "capture_technique",
            Some(preset.capture_technique.as_str().to_string()),
        );
        ini.set(
            &section,
            "child_scope",
            Some(preset.child_scope.as_str().to_string()),
        );
        ini.set(
            &section,
            "broker_fallback",
            Some(
                if preset.broker_fallback {
                    "true"
                } else {
                    "false"
                }
                .to_string(),
            ),
        );
        ini.set(
            &section,
            "sibling_windows",
            Some(preset.sibling_windows.as_str().to_string()),
        );
        ini.set(
            &section,
            "picture_window",
            Some(preset.picture_window.as_str().to_string()),
        );
        ini.set(
            &section,
            "promotion_technique",
            Some(preset.promotion_technique.as_str().to_string()),
        );
        ini.set(
            &section,
            "placement",
            Some(preset.placement.as_str().to_string()),
        );
        ini.set(&section, "scaling", Some(preset.scaling.as_str()));
        ini.set(
            &section,
            "volume_enabled",
            Some(
                if preset.volume_enabled {
                    "true"
                } else {
                    "false"
                }
                .to_string(),
            ),
        );
        ini.set(&section, "volume", Some(preset.volume.to_string()));
    }

    // The order the window lists them in, which is what the top-to-bottom precedence is
    // (`preset_for`). Read back from the settings section rather than from the order of the
    // sections in the file, which is a map's order: alphabetical, and not the list's. Written
    // even with nothing to name, so a file this build wrote is one that has the key.
    let names: Vec<String> = config
        .custom_presets
        .iter()
        .filter_map(|preset| sanitize_name(&preset.name))
        .collect();
    ini.set(CONFIG_SECTION, PRESET_ORDER_KEY, Some(names.join(", ")));
}

/// Every `[custom:<name>]` section read into the configuration, replacing what it held.
///
/// A section that names nothing, that repeats a name already read, or that is the hundred and
/// first is dropped: the file is the user's to edit, and what the window can show is bounded.
/// A key that is missing or unreadable leaves the preset at where a fresh one starts rather than
/// dropping the preset — the name and the program are the two things worth keeping.
pub fn read_all(ini: &Ini, config: &mut AppConfig) {
    let mut presets: Vec<CustomPreset> = Vec::new();

    let mut sections: Vec<&String> = ini
        .get_map_ref()
        .keys()
        .filter(|section| section.starts_with(SECTION_PREFIX))
        .collect();
    sections.sort_unstable();

    for section in sections {
        if presets.len() >= MAX_PRESETS {
            break;
        }
        // The name as the user typed it, which the section cannot hold: a parsed section is
        // lowercased, so the `name` key is what preserves the case (see `section_of`).
        let named = ini
            .get(section, "name")
            .and_then(|name| sanitize_name(&name))
            .or_else(|| sanitize_name(section.trim_start_matches(SECTION_PREFIX)));
        let Some(name) = named else {
            continue;
        };
        if presets
            .iter()
            .any(|held| held.name.eq_ignore_ascii_case(&name))
        {
            continue;
        }

        let mut preset = CustomPreset::named(&name);
        preset.placement = CustomPlacement::default_on_this_machine();
        if let Some(enabled) = ini.get(section, "enabled") {
            if let Some(val) = parse_enabled(&enabled) {
                preset.enabled = val;
            }
        }
        if let Some(exe) = ini.get(section, "exe") {
            preset.exe = sanitize_exe(&exe);
        }
        if let Some(arguments) = ini.get(section, "arguments") {
            preset.arguments = sanitize_arguments(&decode_ini_arguments(&arguments));
        }
        if let Some(method) = ini.get(section, "path_method") {
            if let Some(parsed) = CustomPathMethod::from_str(&method) {
                preset.path_method = parsed;
            }
        }
        if let Some(extensions) = ini.get(section, "extensions") {
            preset.set_extensions_text(&extensions);
        }
        if let Some(window_mode) = ini.get(section, "window_mode") {
            if let Some(mode) = CustomWindowMode::from_str(&window_mode) {
                preset.window_mode = mode;
            }
        }
        if let Some(technique) = ini.get(section, "capture_technique") {
            if let Some(parsed) = CustomCaptureTechnique::from_str(&technique) {
                preset.capture_technique = parsed;
            }
        }
        if let Some(scope) = ini.get(section, "child_scope") {
            if let Some(parsed) = CustomChildScope::from_str(&scope) {
                preset.child_scope = parsed;
            }
        }
        if let Some(enabled) = ini.get(section, "broker_fallback") {
            if let Some(parsed) = parse_enabled(&enabled) {
                preset.broker_fallback = parsed;
            }
        }
        if let Some(siblings) = ini.get(section, "sibling_windows") {
            if let Some(parsed) = CustomSiblingWindows::from_str(&siblings) {
                preset.sibling_windows = parsed;
            }
        }
        if let Some(picture) = ini.get(section, "picture_window") {
            if let Some(parsed) = CustomPictureWindow::from_str(&picture) {
                preset.picture_window = parsed;
            }
        }
        if let Some(technique) = ini.get(section, "promotion_technique") {
            if let Some(parsed) = CustomPromotionTechnique::from_str(&technique) {
                preset.promotion_technique = parsed;
            }
        }
        if let Some(placement) = ini.get(section, "placement") {
            if let Some(parsed) = CustomPlacement::from_str(&placement) {
                preset.placement = parsed;
            }
        }
        if let Some(scaling) = ini.get(section, "scaling") {
            if let Some(scale) = CustomScaling::from_str(&scaling) {
                preset.scaling = scale;
            }
        }
        // Legacy `background=` keys in existing config.ini files are silently ignored.
        if let Some(ve) = ini.get(section, "volume_enabled") {
            if let Some(val) = parse_enabled(&ve) {
                preset.volume_enabled = val;
            }
        }
        if let Some(volume) = ini.get(section, "volume") {
            if let Ok(volume) = volume.trim().parse::<u32>() {
                preset.volume = sanitize_volume(volume);
            }
        }
        presets.push(preset);
    }

    // The order the window listed them in, which the sections themselves do not keep: a name the
    // order names but the file does not hold is dropped, and one it does not name — a preset
    // written by hand — keeps the place the sections gave it, at the end.
    if let Some(order) = ini.get(CONFIG_SECTION, PRESET_ORDER_KEY) {
        let names: Vec<String> = order.split(',').filter_map(sanitize_name).collect();
        presets.sort_by_key(|preset| {
            names
                .iter()
                .position(|name| name.eq_ignore_ascii_case(&preset.name))
                .unwrap_or(usize::MAX)
        });
    }

    config.custom_presets = presets;
}

/// The preset that claims `path`: **the first usable, enabled one in the order the window lists
/// them, top to bottom**, which is what lets a preset above another win the extensions both name.
///
/// It is the question every side asks — the hook, the loader, the layout, the pin — and it is a
/// plain read of the configuration the caller already holds rather than a lock of its own, so a
/// hook callback may ask it (see `routing::custom_preset_for`). A preset with no program or no
/// extension claims nothing: a half-filled tab is one the window keeps and no file is routed by.
pub fn preset_for<'a>(config: &'a AppConfig, path: &Path) -> Option<&'a CustomPreset> {
    if !config.custom_preview_enabled {
        return None;
    }
    config
        .custom_presets
        .iter()
        .find(|preset| preset.enabled && preset.is_usable() && preset.claims(path))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn config_with(presets: Vec<CustomPreset>) -> AppConfig {
        AppConfig {
            custom_presets: presets,
            ..AppConfig::default()
        }
    }

    #[test]
    fn a_name_is_reduced_to_what_a_section_and_a_tab_may_hold() {
        assert_eq!(sanitize_name("  Stress Test  "), Some("Stress Test".into()));
        assert_eq!(sanitize_name("a[b]c"), Some("abc".into()));
        assert_eq!(sanitize_name("   "), None);
        assert_eq!(sanitize_name("..."), None);
        assert_eq!(
            sanitize_name(&"x".repeat(80)).unwrap().chars().count(),
            MAX_NAME_CHARS
        );
    }

    #[test]
    fn extensions_are_read_through_the_list_sanitiser() {
        let mut preset = CustomPreset::named("Videos");
        preset.set_extensions_text(" jpg, .PNG, , gif, not an extension!, jpg ");
        assert_eq!(preset.extensions, vec!["jpg", "png", "gif"]);
        assert_eq!(preset.extensions_text(), "jpg, png, gif");
    }

    #[test]
    fn a_preset_claims_the_extension_it_names_and_nothing_else() {
        let mut preset = CustomPreset::named("Art");
        preset.set_extensions_text("psd");
        assert!(preset.claims(Path::new("C:/work/art.PSD")));
        assert!(!preset.claims(Path::new("C:/work/art.psd.bak")));
        assert!(!preset.claims(Path::new("C:/work/psd")));
    }

    #[test]
    fn the_top_preset_wins_an_extension_two_of_them_name() {
        let mut paint = CustomPreset::named("Painted");
        paint.exe = "C:/Windows/System32/mspaint.exe".into();
        paint.set_extensions_text("jpg, png");

        let mut art = CustomPreset::named("Art");
        art.exe = "C:/Art/Art.exe".into();
        art.set_extensions_text("jpg, psd");

        let config = config_with(vec![paint, art]);
        assert_eq!(
            preset_for(&config, Path::new("holiday.JPG")).map(|p| p.name.as_str()),
            Some("Painted")
        );
        assert_eq!(
            preset_for(&config, Path::new("art.psd")).map(|p| p.name.as_str()),
            Some("Art")
        );
    }

    #[test]
    fn the_window_order_is_written_and_read_back() {
        let mut first = CustomPreset::named("Painted");
        first.exe = "C:/mspaint.exe".into();
        first.set_extensions_text("jpg");
        let mut second = CustomPreset::named("Art");
        second.exe = "C:/Art.exe".into();
        second.set_extensions_text("jpg");

        // "Art" sorts above "Painted" as a section; the written order is what puts Painted top.
        let config = config_with(vec![first, second]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(
            read.custom_presets
                .iter()
                .map(|preset| preset.name.as_str())
                .collect::<Vec<_>>(),
            vec!["Painted", "Art"]
        );
        assert_eq!(
            preset_for(&read, Path::new("holiday.jpg")).map(|p| p.name.as_str()),
            Some("Painted")
        );
    }

    #[test]
    fn a_half_filled_preset_claims_nothing() {
        let mut preset = CustomPreset::named("Empty");
        preset.set_extensions_text("psd");
        assert!(!preset.is_usable());
        preset.exe = "C:/tools/art.exe".into();
        assert!(preset.is_usable());

        let config = config_with(vec![preset]);
        assert!(preset_for(&config, Path::new("art.psd")).is_some());
        assert!(preset_for(&config, Path::new("art.png")).is_none());
    }

    #[test]
    fn a_copied_as_path_program_loses_its_quotes() {
        assert_eq!(
            sanitize_program_input("\"C:\\Program Files\\IrfanView\\i_view64.exe\""),
            "C:\\Program Files\\IrfanView\\i_view64.exe"
        );
        assert_eq!(
            sanitize_program_input("  \"C:\\tools\\art.exe\"\r\n"),
            "C:\\tools\\art.exe"
        );
        assert_eq!(sanitize_program_input("  "), "");
    }

    #[test]
    fn the_file_writes_a_section_per_preset_and_reads_it_back() {
        let mut preset = CustomPreset::named("Stress Test");
        preset.enabled = false;
        preset.exe = "\"tools/Stress.exe\"".into();
        preset.set_extensions_text("rmc, rmp");
        preset.window_mode = CustomWindowMode::Borderless;
        preset.scaling = CustomScaling::Override(PreviewScale::FitToScreenReduced(40));
        preset.volume = 35;

        let config = config_with(vec![preset.clone()]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        assert_eq!(
            ini.get("custom:stress test", "name").as_deref(),
            Some("Stress Test")
        );
        assert_eq!(
            ini.get("custom:stress test", "enabled").as_deref(),
            Some("false")
        );
        assert_eq!(
            ini.get("custom:stress test", "exe").as_deref(),
            Some("tools/Stress.exe")
        );
        assert_eq!(
            ini.get("custom:stress test", "extensions").as_deref(),
            Some("rmc, rmp")
        );
        assert_eq!(
            ini.get("custom:stress test", "window_mode").as_deref(),
            Some("borderless")
        );
        assert_eq!(
            ini.get("custom:stress test", "scaling").as_deref(),
            Some("screen 40")
        );
        assert!(ini.get("custom:stress test", "background").is_none());

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(read.custom_presets.len(), 1);
        assert_eq!(read.custom_presets[0].name, "Stress Test");
        assert!(!read.custom_presets[0].enabled);
        assert_eq!(read.custom_presets[0].extensions, vec!["rmc", "rmp"]);
        assert_eq!(
            read.custom_presets[0].window_mode,
            CustomWindowMode::Borderless
        );
        assert_eq!(
            read.custom_presets[0].scaling,
            CustomScaling::Override(PreviewScale::FitToScreenReduced(40))
        );
        assert_eq!(read.custom_presets[0].volume, 35);
    }

    #[test]
    fn enabled_defaults_true_skips_disabled_and_ignores_legacy_background() {
        let mut ini = Ini::new();
        ini.set("custom:legacy", "name", Some("Legacy".into()));
        ini.set("custom:legacy", "exe", Some("viewer.exe".into()));
        ini.set("custom:legacy", "extensions", Some("psd".into()));
        ini.set("custom:legacy", "background", Some("black".into()));

        let mut cfg = AppConfig::default();
        read_all(&ini, &mut cfg);
        assert_eq!(cfg.custom_presets.len(), 1);
        assert!(cfg.custom_presets[0].enabled);

        let mut off = CustomPreset::named("First");
        off.enabled = false;
        off.exe = "first.exe".into();
        off.set_extensions_text("psd");

        let mut on = CustomPreset::named("Second");
        on.enabled = true;
        on.exe = "second.exe".into();
        on.set_extensions_text("psd");

        let cfg = config_with(vec![off.clone(), on]);
        assert_eq!(
            preset_for(&cfg, Path::new("art.psd")).map(|p| p.name.as_str()),
            Some("Second")
        );

        let cfg_only_off = config_with(vec![off]);
        assert!(preset_for(&cfg_only_off, Path::new("art.psd")).is_none());
    }

    #[test]
    fn volume_enabled_defaults_false_and_round_trips() {
        assert!(!CustomPreset::default().volume_enabled);
        assert!(!CustomPreset::named("Fresh").volume_enabled);

        // Missing key in existing config.ini defaults to false.
        let mut ini = Ini::new();
        ini.set("custom:legacy", "name", Some("Legacy".into()));
        ini.set("custom:legacy", "exe", Some("viewer.exe".into()));
        ini.set("custom:legacy", "extensions", Some("mp4".into()));
        ini.set("custom:legacy", "volume", Some("40".into()));
        let mut cfg = AppConfig::default();
        read_all(&ini, &mut cfg);
        assert!(!cfg.custom_presets[0].volume_enabled);
        assert_eq!(cfg.custom_presets[0].volume, 40);

        // Explicit true and false round-trip through write_all / read_all.
        let mut on = CustomPreset::named("Loud");
        on.exe = "player.exe".into();
        on.set_extensions_text("mp4");
        on.volume_enabled = true;
        on.volume = 65;

        let mut off = CustomPreset::named("Default");
        off.exe = "viewer.exe".into();
        off.set_extensions_text("mkv");
        off.volume_enabled = false;
        off.volume = 25;

        let mut out = Ini::new();
        write_all(&config_with(vec![on, off]), &mut out);
        assert_eq!(
            out.get("custom:loud", "volume_enabled").as_deref(),
            Some("true")
        );
        assert_eq!(
            out.get("custom:default", "volume_enabled").as_deref(),
            Some("false")
        );

        let mut read = AppConfig::default();
        read_all(&out, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        assert!(read.custom_presets[0].volume_enabled);
        assert_eq!(read.custom_presets[0].volume, 65);
        assert!(!read.custom_presets[1].volume_enabled);
        assert_eq!(read.custom_presets[1].volume, 25);
    }

    #[test]
    fn follow_tray_and_normal_window_mode_are_defaults_and_backward_compatible() {
        let mut preset = CustomPreset::named("Viewer");
        preset.exe = "viewer.exe".into();
        preset.set_extensions_text("png");
        assert_eq!(preset.window_mode, CustomWindowMode::Normal);
        assert_eq!(preset.scaling, CustomScaling::FollowTray);

        let config = config_with(vec![preset]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);
        assert_eq!(
            ini.get("custom:viewer", "window_mode").as_deref(),
            Some("normal")
        );
        assert_eq!(ini.get("custom:viewer", "scaling").as_deref(), Some("tray"));

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(read.custom_presets[0].window_mode, CustomWindowMode::Normal);
        assert_eq!(read.custom_presets[0].scaling, CustomScaling::FollowTray);

        // Legacy ini without `window_mode` and with `fit` / numeric `scaling` still loads cleanly.
        let mut legacy = Ini::new();
        legacy.set("custom:old", "name", Some("Old".into()));
        legacy.set("custom:old", "exe", Some("old.exe".into()));
        legacy.set("custom:old", "extensions", Some("jpg".into()));
        legacy.set("custom:old", "scaling", Some("fit".into()));
        let mut read_legacy = AppConfig::default();
        read_all(&legacy, &mut read_legacy);
        assert_eq!(
            read_legacy.custom_presets[0].window_mode,
            CustomWindowMode::Normal
        );
        assert_eq!(
            read_legacy.custom_presets[0].scaling,
            CustomScaling::Override(PreviewScale::FitToScreen)
        );
    }

    #[test]
    fn a_file_that_names_nothing_readable_leaves_a_fresh_preset() {
        let mut ini = Ini::new();
        ini.set("custom:Loose", "volume", Some("900".into()));
        ini.set("custom:Loose", "scaling", Some("enormous".into()));
        ini.set("custom:Loose", "window_mode", Some("unknown".into()));
        ini.set("custom:", "exe", Some("x.exe".into()));

        let mut config = AppConfig::default();
        read_all(&ini, &mut config);

        assert_eq!(config.custom_presets.len(), 1);
        // A section with no `name` key is read by the name the parser gives it, which is the
        // lowercased one (see `section_of`).
        assert_eq!(config.custom_presets[0].name, "loose");
        assert_eq!(config.custom_presets[0].volume, 100);
        assert_eq!(
            config.custom_presets[0].window_mode,
            CustomWindowMode::Normal
        );
        assert_eq!(config.custom_presets[0].scaling, CustomScaling::FollowTray);
        assert!(!config.custom_presets[0].is_usable());
    }

    #[test]
    fn capture_technique_round_trips_and_defaults_to_in_process_hook() {
        assert_eq!(
            CustomCaptureTechnique::default(),
            CustomCaptureTechnique::InProcessHook
        );
        assert_eq!(
            CustomPreset::named("Fresh").capture_technique,
            CustomCaptureTechnique::InProcessHook
        );
        let cases = [
            (CustomCaptureTechnique::InProcessHook, "in_process_hook"),
            (
                CustomCaptureTechnique::InvisibleLayered,
                "invisible_layered",
            ),
            (CustomCaptureTechnique::BehindHost, "behind_host"),
            (CustomCaptureTechnique::OffScreen, "off_screen"),
            (CustomCaptureTechnique::DirectWindow, "direct_window"),
            (CustomCaptureTechnique::OffscreenStaged, "offscreen_staged"),
        ];
        for (tech, expected_str) in cases {
            let mut preset = CustomPreset::named("TechTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.capture_technique = tech;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:techtest", "capture_technique").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].capture_technique, tech);
        }

        // Missing or unknown value falls back to InProcessHook.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "capture_technique", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(p.capture_technique, CustomCaptureTechnique::InProcessHook);
        }
    }

    #[test]
    fn child_scope_round_trips_and_defaults_to_disabled() {
        // Off unless asked for: a walk that can hide, and end, a process the user never named is
        // not a default.
        assert_eq!(CustomChildScope::default(), CustomChildScope::Disabled);
        assert_eq!(
            CustomPreset::named("Fresh").child_scope,
            CustomChildScope::Disabled
        );
        let cases = [
            (CustomChildScope::Disabled, "disabled"),
            (CustomChildScope::Children, "children"),
            (
                CustomChildScope::ChildrenAndGrandchildren,
                "children_and_grandchildren",
            ),
        ];
        for (scope, expected_str) in cases {
            let mut preset = CustomPreset::named("ScopeTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.child_scope = scope;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:scopetest", "child_scope").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].child_scope, scope);
        }

        // Missing or unknown value falls back to Disabled, which is what every config.ini written
        // before this key existed holds.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "child_scope", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(p.child_scope, CustomChildScope::Disabled);
        }

        // A new preset from the window's `+` starts where an unasked preset starts.
        assert_eq!(
            CustomPreset::named("Fresh").child_scope,
            CustomChildScope::Disabled
        );

        // The scope is the walk, in the terms discovery asks it in.
        assert_eq!(
            CustomChildScope::Disabled.depth(),
            CustomChildDepth::LaunchedOnly
        );
        assert_eq!(
            CustomChildScope::Children.depth(),
            CustomChildDepth::OneLevel
        );
        assert_eq!(
            CustomChildScope::ChildrenAndGrandchildren.depth(),
            CustomChildDepth::WholeTree
        );
    }

    #[test]
    fn broker_fallback_defaults_to_off_and_round_trips() {
        assert!(!CustomPreset::named("Fresh").broker_fallback);
        for want in [true, false] {
            let mut preset = CustomPreset::named("BrokerTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.broker_fallback = want;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:brokertest", "broker_fallback").as_deref(),
                Some(if want { "true" } else { "false" })
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].broker_fallback, want);
        }

        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:junk", "name", Some("Junk".into()));
        legacy.set("custom:junk", "exe", Some("b.exe".into()));
        legacy.set("custom:junk", "extensions", Some("png".into()));
        legacy.set("custom:junk", "broker_fallback", Some("perhaps".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert!(!p.broker_fallback);
        }
    }

    #[test]
    fn sibling_windows_round_trips_and_defaults_to_cloak_and_hide() {
        // Today's behaviour is the default, so a file written before this key existed previews
        // exactly as it did.
        assert_eq!(
            CustomSiblingWindows::default(),
            CustomSiblingWindows::CloakAndHide
        );
        assert_eq!(
            CustomPreset::named("Fresh").sibling_windows,
            CustomSiblingWindows::CloakAndHide
        );
        let cases = [
            (CustomSiblingWindows::Disabled, "disabled"),
            (CustomSiblingWindows::Cloak, "cloak"),
            (CustomSiblingWindows::CloakAndHide, "cloak_and_hide"),
        ];
        for (siblings, expected_str) in cases {
            let mut preset = CustomPreset::named("SiblingTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.sibling_windows = siblings;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:siblingtest", "sibling_windows").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].sibling_windows, siblings);
        }

        assert!(!CustomSiblingWindows::Disabled.cloaks());
        assert!(!CustomSiblingWindows::Disabled.hides());
        assert!(CustomSiblingWindows::Cloak.cloaks());
        assert!(!CustomSiblingWindows::Cloak.hides());
        assert!(CustomSiblingWindows::CloakAndHide.cloaks());
        assert!(CustomSiblingWindows::CloakAndHide.hides());

        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:junk", "name", Some("Junk".into()));
        legacy.set("custom:junk", "exe", Some("b.exe".into()));
        legacy.set("custom:junk", "extensions", Some("png".into()));
        legacy.set("custom:junk", "sibling_windows", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(p.sibling_windows, CustomSiblingWindows::CloakAndHide);
        }
    }

    #[test]
    fn picture_window_round_trips_and_defaults_to_newest_comparable() {
        // **The window made last is the default**, which is the one MPV Player's helper window
        // needed; a program with a single picture-shaped window is adopted without waiting for it
        // (see `select_the_picture`), so nothing else pays for the choice.
        assert_eq!(
            CustomPictureWindow::default(),
            CustomPictureWindow::NewestComparable
        );
        assert_eq!(
            CustomPreset::named("Fresh").picture_window,
            CustomPictureWindow::NewestComparable
        );
        assert!(!CustomPictureWindow::Largest.prefers_the_newest());
        assert!(CustomPictureWindow::NewestComparable.prefers_the_newest());

        let cases = [
            (CustomPictureWindow::Largest, "largest"),
            (CustomPictureWindow::NewestComparable, "newest_comparable"),
        ];
        for (picture, expected_str) in cases {
            let mut preset = CustomPreset::named("PictureTest");
            preset.exe = "player.exe".into();
            preset.set_extensions_text("mkv");
            preset.picture_window = picture;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:picturetest", "picture_window").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].picture_window, picture);
        }

        // A key this build does not know, and a section that never had one, both read as the
        // default rather than as a guess.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("mkv".into()));
        legacy.set("custom:junk", "name", Some("Junk".into()));
        legacy.set("custom:junk", "exe", Some("b.exe".into()));
        legacy.set("custom:junk", "extensions", Some("mkv".into()));
        legacy.set("custom:junk", "picture_window", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(p.picture_window, CustomPictureWindow::NewestComparable);
        }
    }

    #[test]
    fn promotion_technique_round_trips_and_defaults_to_direct_promotion() {
        assert_eq!(
            CustomPromotionTechnique::default(),
            CustomPromotionTechnique::DirectPromotion
        );
        assert_eq!(
            CustomPreset::named("Fresh").promotion_technique,
            CustomPromotionTechnique::DirectPromotion
        );
        let cases = [
            (
                CustomPromotionTechnique::DirectPromotion,
                "direct_promotion",
            ),
            (
                CustomPromotionTechnique::OpenAsNewProcess,
                "open_as_new_process",
            ),
        ];
        for (tech, expected_str) in cases {
            let mut preset = CustomPreset::named("PromoTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.promotion_technique = tech;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:promotest", "promotion_technique")
                    .as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].promotion_technique, tech);
        }

        // Missing or unknown value falls back to DirectPromotion.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set(
            "custom:unknown",
            "promotion_technique",
            Some("bogus".into()),
        );

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(
                p.promotion_technique,
                CustomPromotionTechnique::DirectPromotion
            );
        }
    }

    #[test]
    fn placement_round_trips_and_defaults_to_the_machines_own_placement() {
        // The struct's own default is the plain placement, and it is the same on every machine: what
        // varies is the *row's* default, which is asked of the machine where a preset is made (see
        // `CustomPlacement::default_on_this_machine`). A table that read the disk is a table a test
        // could not assert.
        assert_eq!(CustomPlacement::default(), CustomPlacement::FollowTray);
        assert_eq!(
            CustomPreset::named("Fresh").placement,
            CustomPlacement::FollowTray
        );
        let expected_default = CustomPlacement::default_on_this_machine();
        assert_eq!(
            expected_default,
            if crate::paths::hook_dll().is_some() {
                CustomPlacement::Hook
            } else {
                CustomPlacement::FollowTray
            },
            "the hook where the DLL is there to hold the window, the plain placement where it is not"
        );
        let cases = [
            (CustomPlacement::FollowTray, "tray"),
            (CustomPlacement::Forced, "tray_forced"),
            (CustomPlacement::Hook, "tray_hook"),
        ];
        for (placement, expected_str) in cases {
            let mut preset = CustomPreset::named("PlaceTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.placement = placement;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:placetest", "placement").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].placement, placement);
        }

        // **A preset with no `placement` key at all is a new one, and reads the machine's default.**
        // A key that says `tray` was chosen by whoever wrote the file, and it is left alone; an
        // unknown value is not a choice either and takes the same default as a missing key.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:chosen", "name", Some("Chosen".into()));
        legacy.set("custom:chosen", "exe", Some("c.exe".into()));
        legacy.set("custom:chosen", "extensions", Some("png".into()));
        legacy.set("custom:chosen", "placement", Some("tray".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "placement", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 3);
        let by_name = |name: &str| {
            read.custom_presets
                .iter()
                .find(|p| p.name == name)
                .map(|p| p.placement)
                .expect("the preset was read")
        };
        assert_eq!(by_name("Missing"), expected_default);
        assert_eq!(by_name("Unknown"), expected_default);
        assert_eq!(
            by_name("Chosen"),
            CustomPlacement::FollowTray,
            "a preset whose file names the plain placement keeps it"
        );
    }

    #[test]
    fn master_toggle_off_bypasses_all_presets_and_routing() {
        let preset = CustomPreset {
            name: "Viewer".into(),
            exe: "viewer.exe".into(),
            extensions: vec!["png".into()],
            ..CustomPreset::default()
        };
        let mut cfg = config_with(vec![preset]);
        let path = Path::new("photo.png");

        assert!(preset_for(&cfg, path).is_some());
        assert!(crate::formats::routing::is_custom_preview(path, &cfg));

        cfg.custom_preview_enabled = false;
        assert!(preset_for(&cfg, path).is_none());
        assert!(!crate::formats::routing::is_custom_preview(path, &cfg));
        assert!(crate::formats::routing::custom_preset_for(path, &cfg).is_none());
    }

    #[test]
    fn parse_arguments_handles_quotes_escapes_backslashes_and_edge_cases() {
        assert_eq!(parse_arguments(""), Vec::<String>::new());
        assert_eq!(parse_arguments("   \t \r\n "), Vec::<String>::new());
        assert_eq!(
            parse_arguments("--fullscreen --silent"),
            vec!["--fullscreen", "--silent"]
        );
        assert_eq!(parse_arguments("--key=value"), vec!["--key=value"]);
        assert_eq!(
            parse_arguments(r#"--title="hello world""#),
            vec!["--title=hello world"]
        );
        assert_eq!(
            parse_arguments("--title='single quoted' 'has \"double\" inside'"),
            vec!["--title=single quoted", "has \"double\" inside"]
        );
        assert_eq!(
            parse_arguments(r#""it's fine" "\"escaped\"""#),
            vec!["it's fine", "\"escaped\""]
        );
        assert_eq!(
            parse_arguments(r#"--dir=sub\dir\file "trailing\dir\\""#),
            vec![r"--dir=sub\dir\file", r"trailing\dir\"]
        );
        assert_eq!(parse_arguments(r#"a"b c"d"#), vec!["ab cd"]);
        assert_eq!(parse_arguments(r#""""#), vec![""]);
        assert_eq!(
            parse_arguments(r#"--title="unterminated quote"#),
            vec!["--title=unterminated quote"]
        );
        assert_eq!(
            parse_arguments("--title='unterminated single"),
            vec!["--title=unterminated single"]
        );
        assert_eq!(
            sanitize_arguments("  --flag=1 \r\n --second=\"a b\" "),
            "  --flag=1  --second=\"a b\" "
        );
    }

    #[test]
    fn parse_arguments_unwraps_accidental_outer_quotes_only_for_arg_lists() {
        assert_eq!(
            parse_arguments(r#""-fullscreen -borderless""#),
            vec!["-fullscreen", "-borderless"]
        );
        assert_eq!(
            parse_arguments(r#""--fullscreen --silent""#),
            vec!["--fullscreen", "--silent"]
        );
        assert_eq!(
            parse_arguments("'--fullscreen --silent'"),
            vec!["--fullscreen", "--silent"]
        );
        assert_eq!(
            parse_arguments(r#""-fullscreen --path \"dir/sub/app.exe\"""#),
            vec!["-fullscreen", "--path", "dir/sub/app.exe"]
        );
        assert_eq!(
            parse_arguments(r#"--path "dir/sub/app.exe""#),
            vec!["--path", "dir/sub/app.exe"]
        );
        assert_eq!(
            parse_arguments(r#""--path" "dir/sub/app.exe""#),
            vec!["--path", "dir/sub/app.exe"]
        );
        assert_eq!(
            parse_arguments(r#""--title=My Video""#),
            vec!["--title=My Video"]
        );
        assert_eq!(
            parse_arguments(r#""my single file.txt""#),
            vec!["my single file.txt"]
        );
        assert_eq!(
            parse_arguments(r#""{file} --fullscreen""#),
            vec!["{file}", "--fullscreen"]
        );
    }

    /// **A quoted path is one argument, whatever is inside its name.** The unwrap that reads a docs
    /// page's `"--fullscreen --silent"` as two flags used to fire on any `-`-led word inside the
    /// span, and a media file named `2026 - 02 - Batch/Sample Show/…` has one: VLC was handed six media
    /// files (`…/2026`, `-`, `02`, …), asked for each of them, and played none.
    #[test]
    fn a_quoted_path_is_never_read_as_a_run_of_flags() {
        let real = r#""file:///E:/Videos/2026 - 02 - Batch/Sample Show/Sample Show - 82 (1080p).mkv""#;
        assert_eq!(
            parse_arguments(real),
            vec![
                r"file:///E:/Videos/2026 - 02 - Batch/Sample Show/Sample Show - 82 (1080p).mkv"
            ]
        );

        // A plain path with a space in it is one argument as well, and so is a name that happens to
        // open with a dash.
        assert_eq!(
            parse_arguments(r#""E:\Videos\2026 - 02 - Batch\Sample Show\82.mkv""#),
            vec![r"E:\Videos\2026 - 02 - Batch\Sample Show\82.mkv"]
        );
        assert_eq!(
            parse_arguments(r#""- a name that opens with a dash.mkv""#),
            vec!["- a name that opens with a dash.mkv"]
        );
        // A drive letter is not a flag: `E:/…` is not a `/`-led word with something after it.
        assert_eq!(
            parse_arguments(r#""E:/Videos/a b.mkv""#),
            vec!["E:/Videos/a b.mkv"]
        );

        // What the unwrap is for still unwraps: a pasted run of flags, and a placeholder first.
        assert_eq!(
            parse_arguments(r#""--fullscreen --silent""#),
            vec!["--fullscreen", "--silent"]
        );
        assert_eq!(
            parse_arguments(r#""{file} --fullscreen""#),
            vec!["{file}", "--fullscreen"]
        );
        assert_eq!(
            parse_arguments(r#""{url} --network-caching=1000""#),
            vec!["{url}", "--network-caching=1000"]
        );
        // One word, however it reads, is one word.
        assert_eq!(
            parse_arguments(r#""--title=My Video""#),
            vec!["--title=My Video"]
        );
    }

    #[test]
    fn build_launch_args_handles_file_placeholder_and_default_trailing_path() {
        let sample = Path::new(r"samples\my video.mp4");
        assert_eq!(build_launch_args("", sample), vec![r"samples\my video.mp4"]);
        assert_eq!(
            build_launch_args("--fullscreen --silent", sample),
            vec!["--fullscreen", "--silent", r"samples\my video.mp4"]
        );
        assert_eq!(
            build_launch_args("{file} --loop", sample),
            vec![r"samples\my video.mp4", "--loop"]
        );
        assert_eq!(
            build_launch_args("--input={file} --silent", sample),
            vec![r"--input=samples\my video.mp4", "--silent"]
        );
        assert_eq!(
            build_launch_args("--first={file} --second={file}", sample),
            vec![
                r"--first=samples\my video.mp4",
                r"--second=samples\my video.mp4"
            ]
        );
    }

    /// **Every method reads the plain path, the verbatim prefix off.** The Shell hands a hovered
    /// file over as `\\?\E:\…`, and that form is a name no program resolves: VLC answered
    /// `VLC is unable to open the MRL 'file://?/E%3A/…'` — a URL it reconstructed out of a prefix that
    /// was never part of one — which is what a preset with no `Path Method` of its own was handed.
    #[test]
    fn no_method_hands_a_program_the_shells_verbatim_prefix() {
        let verbatim = Path::new(r"\\?\E:\Videos\82.mkv");
        for method in [
            CustomPathMethod::Plain,
            CustomPathMethod::FileUrl,
            CustomPathMethod::ForwardSlashes,
            CustomPathMethod::ShortPath,
            CustomPathMethod::FileUrlLiteral,
        ] {
            let spelled = method.spell(verbatim);
            assert!(
                !spelled.starts_with(r"\\?\"),
                "{method:?} handed over {spelled}"
            );
        }
        assert_eq!(
            CustomPathMethod::Plain.spell(verbatim),
            r"E:\Videos\82.mkv"
        );
        assert_eq!(
            CustomPathMethod::ForwardSlashes.spell(verbatim),
            "E:/Videos/82.mkv"
        );
        assert_eq!(
            CustomPathMethod::FileUrl.spell(verbatim),
            "file:///E:/Videos/82.mkv"
        );
        // And the verbatim form of a share is the share, not a `file:///?` anything.
        assert_eq!(
            CustomPathMethod::FileUrl.spell(Path::new(r"\\?\UNC\server\share\a.mkv")),
            "file://server/share/a.mkv"
        );
    }

    /// A path in a quoted `Arguments` field is **one argument**, and it arrives that way.
    #[test]
    fn a_quoted_path_reaches_the_program_as_one_argument() {
        let file = Path::new(r"\\?\E:\Videos\2026 - 02 - Batch\Sample Show\82.mkv");
        assert_eq!(
            build_launch_args_with(r#""{file}""#, file, CustomPathMethod::Plain),
            vec![r"E:\Videos\2026 - 02 - Batch\Sample Show\82.mkv"]
        );

        // The shape a user pastes after reading a docs page still unwraps, and the file still goes
        // on the end.
        assert_eq!(
            build_launch_args_with(r#""--fullscreen --silent""#, file, CustomPathMethod::Plain),
            vec![
                "--fullscreen",
                "--silent",
                r"E:\Videos\2026 - 02 - Batch\Sample Show\82.mkv"
            ]
        );
    }

    #[test]
    fn path_method_spells_the_file_and_the_two_fixed_words_do_not_move() {
        let file = Path::new(r"E:\Videos\a b.mkv");

        // A method is what `{file}` and the appended file read, and what a preset with no
        // `arguments` at all gets.
        assert_eq!(
            build_launch_args_with("", file, CustomPathMethod::FileUrl),
            vec!["file:///E:/Videos/a%20b.mkv"]
        );
        assert_eq!(
            build_launch_args_with("{file}", file, CustomPathMethod::FileUrl),
            vec!["file:///E:/Videos/a%20b.mkv"]
        );
        assert_eq!(
            build_launch_args_with("--input={file}", file, CustomPathMethod::ForwardSlashes),
            vec!["--input=E:/Videos/a b.mkv"]
        );

        // `{path}` and `{url}` say what they are whatever the row is set to, so one command line
        // may carry both spellings at once.
        assert_eq!(
            build_launch_args_with("--a={path} --b={url}", file, CustomPathMethod::ShortPath),
            vec![r"--a=E:\Videos\a b.mkv", "--b=file:///E:/Videos/a%20b.mkv"]
        );

        // A placeholder may be named twice, once in a value and once on its own.
        assert_eq!(
            build_launch_args_with("--pair={url}:{url}", file, CustomPathMethod::Plain),
            vec!["--pair=file:///E:/Videos/a%20b.mkv:file:///E:/Videos/a%20b.mkv"]
        );

        // A brace that is nobody's is left as it was written, and does not stop the scan: a
        // program whose own syntax has one still gets its file.
        assert_eq!(
            build_launch_args_with("--x={nope} --y={file}", file, CustomPathMethod::Plain),
            vec![r"--x={nope}", r"--y=E:\Videos\a b.mkv"]
        );

        // A file whose own name holds a placeholder is a file, not a second substitution: the
        // expansion is one pass, and what it wrote is never read again.
        let named = Path::new(r"E:\Videos\{url}.mkv");
        assert_eq!(
            build_launch_args_with("{file}", named, CustomPathMethod::Plain),
            vec![r"E:\Videos\{url}.mkv"]
        );
    }

    #[test]
    fn path_method_round_trips_and_an_unknown_one_reads_as_plain() {
        assert_eq!(CustomPathMethod::default(), CustomPathMethod::Plain);
        assert_eq!(
            CustomPreset::named("Fresh").path_method,
            CustomPathMethod::Plain
        );

        for method in [
            CustomPathMethod::Plain,
            CustomPathMethod::FileUrl,
            CustomPathMethod::ForwardSlashes,
            CustomPathMethod::ShortPath,
            CustomPathMethod::FileUrlLiteral,
        ] {
            let mut preset = CustomPreset::named("Paths");
            preset.exe = "vlc.exe".into();
            preset.set_extensions_text("mkv");
            preset.path_method = method;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:paths", "path_method").as_deref(),
                Some(method.as_str())
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].path_method, method);
        }

        // A preset written before the key existed, and one naming something this build does not
        // know, are both the plain path: that is what every preset did before the row was there.
        let mut legacy = Ini::new();
        legacy.set("custom:old", "name", Some("Old".into()));
        legacy.set("custom:old", "exe", Some("a.exe".into()));
        legacy.set("custom:old", "extensions", Some("png".into()));
        legacy.set("custom:bogus", "name", Some("Bogus".into()));
        legacy.set("custom:bogus", "exe", Some("b.exe".into()));
        legacy.set("custom:bogus", "extensions", Some("png".into()));
        legacy.set("custom:bogus", "path_method", Some("quantum".into()));
        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for preset in &read.custom_presets {
            assert_eq!(preset.path_method, CustomPathMethod::Plain);
        }
    }

    #[test]
    fn arguments_ini_write_serialize_load_round_trips_special_chars_and_defaults_empty() {
        assert_eq!(CustomPreset::default().arguments, "");

        let tricky = [
            "--fullscreen --silent",
            r#"--title="a=b c""#,
            "--background-color=#111111",
            "--vf=lavfi=[scale=1280:-1]",
            "--opts=a=1;b=2",
            "100%",
            "--literal=%23 --combo=%25#;[ ]",
        ];

        for &args in &tricky {
            let mut preset = CustomPreset::named("ArgsTest");
            preset.exe = "viewer.exe".into();
            preset.arguments = args.into();
            preset.set_extensions_text("mp4");

            let cfg = config_with(vec![preset]);
            let mut written = Ini::new();
            write_all(&cfg, &mut written);

            let serialized = written.writes();
            if args == "--fullscreen --silent" {
                assert!(
                    serialized.contains("arguments=--fullscreen --silent"),
                    "normal args must remain human-readable in INI:\n{serialized}"
                );
            }

            let mut loaded = Ini::new();
            loaded
                .read(serialized)
                .expect("serialized INI with escaped arguments must parse cleanly");

            let mut read = AppConfig::default();
            read_all(&loaded, &mut read);
            assert_eq!(read.custom_presets.len(), 1);
            assert_eq!(
                read.custom_presets[0].arguments, args,
                "failed round-trip for `{args}`"
            );
        }

        // Hand-edited literal `%` (e.g. `100%`) without `%25` encoding still reads cleanly,
        // and missing `arguments` key defaults to empty string `""`.
        let raw_ini = "\
[custom:hand]
name=Hand
exe=viewer.exe
extensions=png
arguments=--zoom=100% --silent

[custom:legacy]
name=Legacy
exe=old.exe
extensions=bmp
";
        let mut parsed = Ini::new();
        parsed
            .read(raw_ini.to_string())
            .expect("hand INI must parse");
        let mut read = AppConfig::default();
        read_all(&parsed, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        let hand = read
            .custom_presets
            .iter()
            .find(|p| p.name == "Hand")
            .expect("Hand preset");
        let legacy = read
            .custom_presets
            .iter()
            .find(|p| p.name == "Legacy")
            .expect("Legacy preset");
        assert_eq!(hand.arguments, "--zoom=100% --silent");
        assert_eq!(legacy.arguments, "");
    }
}
