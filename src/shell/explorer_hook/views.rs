//! What a look at Explorer comes back in: the view that answered, the item the pointer
//! is over, and the file that item names. Nothing here reads the shell — these are
//! the answers the walks beside them produce, and the comparisons every caller
//! would otherwise write for itself (is this the same item, does this one draw in
//! columns, where does its name sit in the view).

use super::*;

/// What the walk over Explorer's own windows found: how many there are, how many are
/// showing, and how many of those the region the window in front covers does not
/// hold — the ones a hover can still reach.
pub(super) struct ExplorerWindowCounts {
    pub(super) total: usize,
    pub(super) visible: usize,
    pub(super) reachable: usize,
    /// The region the window in front hides what is behind, where that window is one
    /// that hides anything at all — see `foreground_cover_rect`.
    pub(super) cover: Option<RECT>,
}

/// What the desktop looks like from here, which is what decides whether everything
/// the hook remembers about a view — the window the pointer is in, the folder it has
/// open, the item an observation was made of — still describes anything.
///
/// Each display is in it, rather than the union of them: a display rescaled from 100%
/// to 150% rearranges everything drawn on that display and leaves the union exactly
/// where it was, and a display becoming the primary one moves where new windows open
/// and what the system's own metrics are measured against, without moving the union
/// either.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(super) struct DisplaySignature {
    pub(super) displays: Vec<DisplayEntry>,
}

/// One display as the signature sees it: where it is, what it is scaled to, and
/// whether it is the primary one.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(super) struct DisplayEntry {
    pub(super) rect: (i32, i32, i32, i32),
    pub(super) dpi: u32,
    pub(super) primary: bool,
}

/// The view under a point, as the shell describes it: the window it is drawn in,
/// the URL it was opened with, and the folder it has open.
pub(super) struct ActiveShellViewContext {
    pub(super) shell_view_hwnd: isize,
    pub(super) location_url: Option<String>,
    pub(super) folder_path: Option<String>,
}

#[derive(Clone, Default)]
pub(super) struct HoverResolverHints {
    pub(super) current_folder: Option<String>,
    pub(super) location_url: Option<String>,
    pub(super) is_search_view: bool,
    pub(super) search_root: Option<String>,
    pub(super) shell_view_hwnd: Option<isize>,
}

/// What both paths need to resolve an item to the file it stands for: one UI
/// Automation client whose property reads are batched into a single round trip per
/// element, Explorer's own item-position property, and the view that last answered
/// for a window.
///
/// What it holds is as telling as what it does not: there is no folder index, no
/// view index and no search root here. A file is found from the item that stands
/// for it rather than from its name, so nothing has to be walked, remembered or
/// kept warm for either path to have an answer.
#[derive(Default)]
pub(super) struct ItemResolver {
    pub(super) automation: Option<IUIAutomation>,
    /// Whether every call the client above makes is bounded. A resolver holding one
    /// that is not asks for a bounded client again on a slow cadence, because a
    /// probe through an unbounded one is a wait with nothing watching it — see
    /// `rebuild_automation`.
    pub(super) automation_bounded: bool,
    /// The batched property request every element is read with, so an element
    /// costs one crossing into the view's provider rather than one per property.
    pub(super) cache: Option<IUIAutomationCacheRequest>,
    pub(super) walker: Option<IUIAutomationTreeWalker>,
    /// Explorer's own `ItemIndex` property, registered once per process. `None`
    /// when the registrar refuses it, which leaves the identity route skipped and
    /// the item's own value to answer.
    pub(super) item_index_property: Option<UIA_PROPERTY_ID>,
    /// The Shell window collection, created once and kept: building it is the one
    /// call every lookup would otherwise repeat.
    pub(super) shell_windows: Option<IShellWindows>,
    /// The views of the window the last look was in, kept while they still describe
    /// it: reading them is one crossing into the shell per Shell window the desktop
    /// holds, and a window that holds tabs is one registration per tab (see
    /// `frame_views`).
    pub(super) window_views: Option<WindowViews>,
    /// The answer the last look at the item under the pointer produced, kept while the
    /// pointer stays inside that item — see `AnsweredItem`.
    pub(super) item: Option<AnsweredItem>,
    pub(super) probe: Option<ProbeMemo>,
}

/// One view of a window, as the Shell hands it over.
pub(super) struct AnsweredView {
    /// The view's own identity, which is what tells one view of a window from another —
    /// a window that holds tabs has one per tab — and is the same object across a
    /// navigation: what changes then is the folder the view holds, not the view.
    pub(super) view_identity: *mut core::ffi::c_void,
    /// The window the view is drawn in, which is what says which of a frame's views the
    /// item under the pointer belongs to (see `ItemWindow`). Nothing where the view will
    /// not name one, which leaves it to answer with the rest.
    pub(super) view_hwnd: isize,
    /// The browser object the view was found through, which is what knows the URL the view
    /// was opened with: the place a probe remembers, and the one fact of a place the Shell
    /// always answers (see `ActiveShellViewContext`).
    pub(super) browser: IWebBrowser2,
    /// The view itself, which is what the folder it has open is walked out of. It is kept
    /// rather than asked for again because asking for it is a crossing into the shell, and
    /// the walk that found this view has already made it (see `folder_views_for_window`).
    pub(super) shell_view: IShellView,
    pub(super) folder_view: IFolderView2,
}

impl AnsweredView {
    /// What this view is showing: the URL it was opened with, and — where the caller wants
    /// it — the folder it has open, which is what a probe remembers about the place.
    ///
    /// It is the question a folder probe asks of the view it is about — the URL the view
    /// was opened with and the folder it has open — asked of a view that was found once for
    /// a window rather than once for every probe: the same two facts of the same objects,
    /// paid for once per window instead of once per tick (see `WindowViews`).
    pub(super) fn describe(&self, want_folder: bool) -> ActiveShellViewContext {
        unsafe {
            let location_url = self.browser.LocationURL().ok().map(|url| url.to_string());
            let folder_path = if want_folder {
                get_shell_view_folder_path(&self.shell_view)
            } else {
                None
            };

            ActiveShellViewContext {
                shell_view_hwnd: self.view_hwnd,
                location_url,
                folder_path,
            }
        }
    }
}

/// Every view one window holds, as a walk of the Shell window collection found them.
///
/// What is kept is the whole set and not the one view that answered, because which of
/// them is showing what is under the pointer is a question about the pointer rather than
/// about the set: a tab switched is the same set with a different window inside it (see
/// `ItemWindow`), and a set read again for that would be the walk this exists to save.
/// So the set is read again only where it cannot describe the window any more, and that
/// is a short list: the window is gone, hidden or minimized, or the desktop holds a
/// different number of Shell windows than it did when the set was read — one cheap
/// number that moves when a tab or a window is opened or closed, and the only thing that
/// ever adds a view to a window or takes one away.
pub(super) struct WindowViews {
    pub(super) frame: isize,
    /// How many Shell windows the desktop held when this set was read.
    pub(super) registrations: i32,
    /// Which of the views a probe last found the pointer inside, where one did.
    ///
    /// It is what a probe that is in *none* of them is answered by: a pointer over the
    /// navigation pane, the toolbar or the details pane is in no tab, so no view of the
    /// frame can be told from another by the pointer, and what the frame is showing is a
    /// question this set cannot answer. Answering with the view the hand was last inside
    /// is answering with the tab the user was last working in, and — which matters more —
    /// it is an answer that does not move while the pointer does not: a place read twice
    /// is one place rather than two (see `anchored_view_context`).
    ///
    /// It is an index into `views` and it goes with the set: a set read again is a set
    /// whose views were found again, and the anchor is dropped with the old one.
    pub(super) anchor: Option<usize>,
    pub(super) views: Vec<AnsweredView>,
}

impl WindowViews {
    /// Whether the window this set was read for is still one an item can be drawn in.
    /// A window that is gone, hidden or minimized holds nothing to answer about — the
    /// same question `folder_views_for_window` asks of a window before walking it.
    pub(super) fn is_live(&self) -> bool {
        let frame = HWND(self.frame as *mut core::ffi::c_void);
        !frame.is_invalid()
            && unsafe { IsWindowVisible(frame).as_bool() && !IsIconic(frame).as_bool() }
    }
}

/// The window an item is resolved in: the frame whose views can be holding it, and the
/// window the item is drawn in.
///
/// The second is what tells a frame's views apart, and it is the fact this path did
/// without for as long as they could not be told apart at all. A window that holds tabs
/// registers one Shell window per tab, every one of them answering with the frame's own
/// window, so the frame names a *set* of views rather than one — and the tabs that are
/// not showing cannot be told from the showing one by their own windows, which are
/// visible either way. What the item is drawn in is inside a window of exactly one of
/// them: the tab that is showing. A view whose window the item's own window descends
/// from is therefore the view the item was drawn by, and it is asked alone. It is the
/// same test the folder probe makes to find the view a point is in (see
/// `anchored_view_context`), which is why a pointer that is in none of them —
/// over the navigation pane, the toolbar, the details pane, none of which belongs to a
/// tab — names no view, and the frame's views are told apart by what they answer, as
/// they were before this window was read.
#[derive(Clone, Copy)]
pub(super) struct ItemWindow {
    /// The frame the item path resolves against: the root window under the pointer, or
    /// the frame the focused item is drawn in.
    pub(super) frame: isize,
    /// The window the item is drawn in — the window under the pointer for the pointer's
    /// item, the window the item's own provider reports for the keyboard's — or nothing
    /// where neither is known.
    pub(super) drawn_in: isize,
}

impl ItemWindow {
    /// The one view of a frame that drew the item, where the frame's views can be told
    /// apart by the window the item is in.
    ///
    /// Two views claiming that window is a frame this cannot tell apart, and it is
    /// answered the way it was before the window was read: every view is asked and what
    /// they answer has to agree. A window is on exactly one chain from the desktop down
    /// to what is under the pointer, so two views holding it is a reading that cannot be
    /// trusted rather than one of several, and nothing is answered on it.
    pub(super) fn view_holding(&self, views: &[AnsweredView]) -> Option<usize> {
        let mut found = None;
        for (position, view) in views.iter().enumerate() {
            if !self.draws_inside(view.view_hwnd) {
                continue;
            }
            if found.is_some() {
                return None;
            }
            found = Some(position);
        }

        found
    }

    /// Whether a view's own window is one the item is drawn inside.
    pub(super) fn draws_inside(&self, view_hwnd: isize) -> bool {
        self.drawn_in != 0
            && view_hwnd != 0
            && hwnd_is_same_or_ancestor(
                HWND(self.drawn_in as *mut core::ffi::c_void),
                HWND(view_hwnd as *mut core::ffi::c_void),
            )
    }
}

/// What a view says about the item at a position: whether it holds it, and whether there
/// is a file to preview if it does.
///
/// The two questions are asked in order and are not the same question (see
/// `item_file_path`), and they are asked of one item object rather than two: the item at
/// that position is fetched once and asked for the name it is shown under and for the
/// path it stands for. Fetching it is the crossing the two questions used to pay twice —
/// and a frame whose every view is asked paid it twice for each of them.
pub(super) enum ViewItem {
    /// The view could not be asked for the item at all. A view is not something this app
    /// can see the end of: a window that has closed takes its views with it and the
    /// proxies left behind answer nothing, so this is what a set that may be describing a
    /// window that has moved on looks like, and the frame's views are read again.
    Unaskable,
    /// The view does not hold the item at that position.
    NotHeld,
    /// The view holds it, and it is not a file this app previews.
    NoFile,
    /// The view holds it, and this is the file it stands for.
    File(PathBuf),
}

/// The answer one look at the item under the pointer produced, kept while the pointer
/// stays inside the item it was read from.
///
/// An answer is kept against the item rather than the point, because a point is what a
/// hand sweeping a list has a new one of every tick and the item is not: a `Details` row
/// is as wide as the view, so a pointer moving along one crosses many points and never
/// leaves the item the first of them was answered for. What ends it is the pointer
/// leaving that box — or the window the box was read in, which a box on its own cannot
/// stand in for: a tab switched under a parked pointer is another view drawing another
/// folder at the same place, and the box that held the item of one is the box of the
/// item of the other. Everything else that makes the item under a parked pointer a new
/// question drops the answer where it happens (see `ItemResolver::forget_item`).
pub(super) struct AnsweredItem {
    /// The window the item under the pointer was drawn in — see `ItemWindow`. The
    /// pointer has to be in the same window for the answer to still be about it.
    pub(super) drawn_in: isize,
    /// The box the view draws the item in, which the pointer has to stay inside.
    pub(super) bounds: (i32, i32, i32, i32),
    /// The file the item stands for, or nothing where the view holds an item with no
    /// file to preview — a folder, an application. That is an answer as well: the item
    /// under the pointer is one this app has nothing to show for, which is not a reason
    /// to ask the shell about it again on every tick.
    pub(super) path: Option<PathBuf>,
}

/// What one look at the pointer found.
pub(super) struct PointerLook {
    /// The file the pointer is on, where what is under it is a file this app previews.
    pub(super) path: Option<PathBuf>,
    /// The box the view draws the item the look read in, or nothing where the look found
    /// no item at all. It is published for the preview thread and the pointer-left check
    /// as it always was, and it is what the answer is kept against — see `AnsweredItem`.
    pub(super) item_bounds: Option<(i32, i32, i32, i32)>,
    /// The window the look read the item in, or nothing where the pointer was in no
    /// window this app could name.
    pub(super) drawn_in: isize,
}

/// The answer one point produced, kept for the rest of the loop tick.
pub(super) struct ProbeMemo {
    pub(super) point: POINT,
    pub(super) answer: Option<PathBuf>,
}

/// The text an item draws inside its own box, as the item's own children report it.
///
/// One walk reads the pieces of it for every way of avoiding that asks for one, so it
/// answers with both boxes rather than being asked twice: the whole of what the item
/// draws, and the piece its name is drawn in.
#[derive(Clone, Copy)]
pub(super) struct ItemText {
    /// Every piece of the item's text taken as one box — a row's name with the columns
    /// beside it, the label under an icon — which is the region `Avoid Details` keeps a
    /// preview off. Its right edge is where the item's content stops.
    pub(super) all: RECT,
    /// The piece the name is drawn in: the leftmost of them, which in the views that
    /// draw their items as rows is the `Name` column of `Details` — the name above the
    /// path of `Content` — and the label itself under an icon. It is the room the name
    /// is *given*: the region `Avoid Filename Column` keeps a preview off, and the box
    /// `Avoid Filename` narrows to the width the name itself is drawn at — see
    /// [`HoveredItem::name_box`].
    pub(super) name: RECT,
    /// Whether anything is drawn beside the piece the name is drawn in. A view that
    /// draws its items as rows writes the columns beside the name there — a `Details`
    /// row's type, date and size, the path and details of a `Content` row — where a
    /// label under an icon, a tile's stacked lines and a name alone draw nothing
    /// beside it. It is what says the item is a row of its view rather than a box: a
    /// row's text is written into the left end of a box as wide as the view, so the
    /// columns are the room a keyboard preview may take and the box's own right edge
    /// is not — see [`HoveredItem::draws_columns`].
    pub(super) columns: bool,
}

/// The item a file is resolved from, as the view's accessibility provider reports
/// it — read at the cursor for the pointer and at the focused item for the
/// keyboard, so both paths get their answer from the same facts.
pub(super) struct HoveredItem {
    /// The item's position in the view, one-based as Explorer's own `ItemIndex`
    /// reports it — the one fact about a search result that a shared name cannot
    /// take away, because two results may share a name and only one of them is
    /// at this position.
    pub(super) index: Option<i32>,
    /// The name the item goes by in the view, which may hide the extension.
    pub(super) name: String,
    /// The item's legacy accessible value: for a file a search has surfaced this
    /// is normally the file's own path, which is the second way an item is
    /// answered when it reports no position.
    pub(super) value: Option<String>,
    /// The box the item occupies on screen, which is what says the pointer is
    /// inside it and where a keyboard preview is placed.
    pub(super) bounds: RECT,
    /// The text the item draws — its name, and the columns a view that draws its
    /// items as rows writes beside it — or `None` when the view reported no text, or
    /// the walk that would have read one was not asked for it, which is what a walk
    /// with nothing to keep a preview off asks. What it holds is the region the
    /// `Avoid` setting places a preview from, the keyboard's as much as the pointer's,
    /// and whether the item draws a row of columns at all — see
    /// [`HoveredItem::draws_columns`].
    pub(super) text: Option<ItemText>,
    /// The window the item is drawn in, whose frame is the window the item's view
    /// belongs to.
    pub(super) native_window: isize,
}

impl HoveredItem {
    /// Whether a second look found the same item. What the view says about an item
    /// is only true of the item — a list can move under a parked pointer — so an
    /// answer is only taken when both looks agree.
    pub(super) fn same_item(&self, other: &HoveredItem) -> bool {
        self.index == other.index && self.name == other.name
    }

    /// The region a preview of this item is kept off, as the `Avoid` setting has it:
    /// the name where it is drawn at `Filename`, the box the view gives the name at
    /// `FilenameColumn`, that box with the columns a row writes beside it at `Details`,
    /// and nothing at all at `Off` — or the item's own box at any of the first three
    /// for a view that reports no text, the name being drawn inside that box whatever
    /// the view says about it, so an item whose text cannot be measured is avoided as
    /// the whole of itself.
    ///
    /// The region comes with whether it is a *column* of the view, which is the one
    /// thing a placement cannot read off the box itself: a view that draws its items
    /// as rows ([`Self::draws_columns`]) puts every row's text in the same columns, so
    /// a preview that overlaps the `Name` column — or the columns beside it, at
    /// `Details` — covers the rows next to the item however it is placed vertically,
    /// and the ways out of it are the two to its sides (see
    /// `preview_window::avoiding_text`). The name alone, a label under an icon and the
    /// item's own box are regions of their own kind however the setting reads, and a
    /// preview is stepped off them in either axis.
    ///
    /// It is the pointer's region. A keyboard preview is placed from a region of its
    /// own, which is this one except at `Off` — see [`Self::keyboard_avoid_box`].
    pub(super) fn avoid_box(&self) -> Option<((i32, i32, i32, i32), bool)> {
        let (region, column) = match avoid_mode() {
            AvoidMode::Off => return None,
            AvoidMode::Filename => (self.name_box(), false),
            AvoidMode::FilenameColumn => (self.text.map(|text| text.name), true),
            AvoidMode::Details => (self.text.map(|text| text.all), true),
        };

        let region = region.unwrap_or(self.bounds);

        Some((
            (region.left, region.top, region.right, region.bottom),
            column && self.draws_columns(),
        ))
    }

    /// The region a *keyboard* preview of this item is kept off: what [`Self::avoid_box`]
    /// names for every way of avoiding but one — `Avoid Nothing` is read as
    /// `Avoid Filename`.
    ///
    /// A pointer's `Avoid Nothing` places a preview by the position mode alone because
    /// the cursor is what such a placement is read from. A keyboard preview has no
    /// cursor, and the item it has instead is a box as wide as the view for a row — so a
    /// placement by the position mode alone is one anchored at that box's middle, over
    /// the very file it describes, which is the placement the region exists to forbid.
    /// The name is the one thing the item says about itself and the one line a preview
    /// can be put beside, so a setting that keeps nothing off keeps the name off, and a
    /// keyboard preview is placed the way `Avoid Filename` places it.
    pub(super) fn keyboard_avoid_box(&self) -> Option<(i32, i32, i32, i32)> {
        let region = match avoid_mode() {
            AvoidMode::Off | AvoidMode::Filename => self.name_box(),
            AvoidMode::FilenameColumn => self.text.map(|text| text.name),
            AvoidMode::Details => self.text.map(|text| text.all),
        }
        .unwrap_or(self.bounds);

        Some((region.left, region.top, region.right, region.bottom))
    }

    /// Whether the item draws its text as a row of its view: a name with the columns
    /// of a `Details` or `Content` row written beside it, rather than the label under
    /// an icon or a name on its own. See [`ItemText::columns`].
    ///
    /// It is read off the item's own text for the keyboard rather than guessed from
    /// the item's box, because a row's box is as wide as the *view* it is drawn in and
    /// not as wide as the display: a `Details` row of a window that is a quarter of the
    /// display across is a row all the same, and at `Avoid Filename` its preview
    /// belongs past the name — over the columns the setting lets it cover — and not
    /// past the row's own right edge, which is what the placement would take were the
    /// row read as a box. An item whose text cannot be measured answers `false`, which
    /// leaves it placed by its box the way it always was.
    pub(super) fn draws_columns(&self) -> bool {
        self.text.is_some_and(|text| text.columns)
    }

    /// The box the item's name is drawn in, cut to the width the name itself takes:
    /// the region `Avoid Filename` keeps a preview off, where the box as the view
    /// reported it is the one `Avoid Filename Column` keeps it off.
    ///
    /// A view reports the room an item's name is *given* — the `Name` column of a
    /// `Details` row is one width for every file in it, a long name and a short one
    /// alike — so the width the name is drawn at, at the larger of the two sizes a view
    /// may draw it at, is measured and the box is narrowed to it. It is only ever
    /// narrowed: a name that fills the room it was given, or is drawn truncated to it,
    /// is left as the view reported it.
    fn name_box(&self) -> Option<RECT> {
        let text = self.text?;

        let Some(width) = drawn_name_width(&self.name, region_display_dpi(&text.name)) else {
            return Some(text.name);
        };

        Some(RECT {
            right: (text.name.left + width).min(text.name.right),
            ..text.name
        })
    }
}

/// The display a region is drawn on, as the DPI its pixels are in.
///
/// The window an item is reported through is not a display that can always be asked:
/// the shell's own item provider answers with no window at all, and a display taken
/// from the system's instead is the wrong one for an item on a scaled display — a name
/// drawn at 200% would be measured half its width. Where the region is, is what says
/// which display it is drawn on, so the middle of it is what is looked up.
pub(super) fn region_display_dpi(region: &RECT) -> u32 {
    monitor_dpi_from_point(
        (region.left + region.right) / 2,
        (region.top + region.bottom) / 2,
    )
}

/// The share of the icon font the shell draws an item's name at in the views that give
/// the name a line of its own: `Content` draws the name at 125% of the icon font, where
/// a `Details` row is drawn at the font itself. A name is measured at the larger of the
/// two so that one region clears the name in either view — which costs a `Details` row
/// a quarter more room than its name takes, and is what keeps a `Content` row's name,
/// extension and all, from being covered.
///
/// The share was read off the pixels: at 100% scaling, `aa.txt`,
/// `mid-length-name.txt` and `a-very-long-file-name-here.txt` are drawn 36, 136 and 202
/// pixels wide in `Content` against 28, 110 and 162 in `Details`.
pub(super) const NAME_FONT_SCALE: (i64, i64) = (5, 4);

/// The width a name is drawn at, in the pixels of the display the item is on, or
/// `None` when it cannot be measured — which leaves the name as the view reported it.
///
/// The shell draws an item's name in the icon font, the one folder views are given, at
/// the size the view gives it — see `NAME_FONT_SCALE` — so that is what a name is
/// measured with: a font and a memory DC are made for the one call and let go again. It
/// is asked beside the walk that read the item — once per preview, not once per probe —
/// and the font is the shell's own, written for the display the system is at, so it is
/// scaled here to the display the item is drawn on — `dpi`, the one the region is in
/// (see [`region_display_dpi`]) — which is the scale the drawn name is in.
///
/// A display that is not known is not measured at another one's scale: the name is left
/// as the view reported it, the answer a name with nothing in it gets, because a region
/// that clears too much room is one a preview is placed too far away, while a region cut
/// at the wrong scale is one that covers the name it was cut from.
pub(super) fn drawn_name_width(name: &str, dpi: u32) -> Option<i32> {
    let wide: Vec<u16> = name.encode_utf16().collect();
    if wide.is_empty() || dpi == 0 {
        return None;
    }

    unsafe {
        let mut logfont = LOGFONTW::default();
        SystemParametersInfoW(
            SPI_GETICONTITLELOGFONT,
            0,
            Some(&mut logfont as *mut LOGFONTW as *mut core::ffi::c_void),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .ok()?;

        let system_dpi = GetDpiForSystem();
        if system_dpi != 0 && dpi != system_dpi {
            let height = logfont.lfHeight as i64 * dpi as i64;
            logfont.lfHeight = (height / system_dpi as i64) as i32;
        }

        // The name is drawn at one of two sizes and the region has to clear whichever
        // it is, so the larger is what is measured — see `NAME_FONT_SCALE`.
        let height = logfont.lfHeight as i64 * NAME_FONT_SCALE.0;
        logfont.lfHeight = (height / NAME_FONT_SCALE.1) as i32;

        let font = CreateFontIndirectW(&logfont);
        if font.0.is_null() {
            return None;
        }

        let dc = CreateCompatibleDC(None);
        if dc.0.is_null() {
            let _ = DeleteObject(font);
            return None;
        }

        let previous = SelectObject(dc, font);
        let mut extent = SIZE::default();
        let measured = GetTextExtentPoint32W(dc, &wide, &mut extent).as_bool();
        let _ = SelectObject(dc, previous);
        let _ = DeleteObject(font);
        let _ = DeleteDC(dc);

        (measured && extent.cx > 0).then_some(extent.cx)
    }
}
