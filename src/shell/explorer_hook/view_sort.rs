//! How long the hook waits between looks, and the order a folder's listing is in.
//!
//! The sort is the shell's own answer to `GetSortColumns`, remembered per folder,
//! because the pin's previous and next walk the files as the listing shows them: a
//! listing walked in another order is a listing the pin walks wrongly (see
//! `shell::pin_navigation`).

use super::*;

pub(super) const EXPLORER_WINDOW_CACHE_TTL_MS: u64 = 1000;
pub(super) const EXPLORER_REAL_FOLDER_CACHE_MAX_ENTRIES: usize = 256;
pub(super) const FOLDER_PROBE_MS: u64 = 200;
pub(super) const IDLE_FOLDER_PROBE_MS: u64 = 750;
pub(super) const FOLDER_PROBE_TRIGGER_MS: u64 = 400;
pub(super) const DISPLAY_CHANGE_BACKOFF_MS: u64 = 1500;
/// How often the desktop is walked and compared with what it was: the displays a
/// machine has, each one's scale, and which of them is the primary one (see
/// `current_display_signature`). A display change is acted on within this of it
/// happening, which is a fifth of a second against the backoff the rebuild waits out
/// anyway — and walking every display on every tick would be the most expensive thing
/// in a loop that runs two or three dozen times a second.
pub(super) const DISPLAY_CHECK_MS: u64 = 200;
pub(super) const KEYBOARD_FOCUS_INPUT_GRACE_MS: u64 = 500;
pub(super) const HOVER_RESOLVER_INPUT_GRACE_MS: u64 = 1500;
/// How long a click a pinned preview never got to look at is held for before it is let go.
///
/// The press bit that says a click happened can only be read once, and the tick that reads it is
/// the only tick it is on: a click that takes the focus away from the pinned window, and a click
/// that gives it back to Explorer, are both asked about before UI Automation has caught up with
/// the listing under the hand, and what is under the pointer is answered on a tick the hand has
/// already moved on (see `PinUpdateWatch::follow`). Half a second is the same window the
/// keyboard's own input is given to move something in, and a hand that clicks and then leaves has
/// not asked for anything to be looked up on its behalf after that.
pub(super) const PIN_CLICK_RETRY_MS: u64 = 500;
/// How often the file the listing under the pointer has selected is read while a pin follows
/// picks rather than hovers: the selection is answered out of live shell objects on every
/// read, so it is polled on its own cadence rather than on every tick of a loop that runs
/// several dozen times a second. A click is still answered on its own tick — the press is
/// what schedules a read outside this cadence — so what this bounds is the keyboard's lag
/// behind a parked hand, and what a folder opened under one costs (see
/// `PinUpdateWatch::follow_selection`).
pub(super) const PIN_SELECTION_POLL_MS: u64 = 120;
/// How often the item the keyboard is on is looked at, while a key is being pressed or while a
/// preview is pinned and the pin follows the keyboard: the focus is read through UI Automation,
/// which is a crossing into Explorer, so it is asked no faster than this — fast enough that a list
/// walked with the arrow keys is read as being followed rather than as catching up.
pub(super) const KEYBOARD_FOCUS_PROBE_MS: u64 = 30;

/// The sleep each Explorer state is answered with, and how often the state is read again while it
/// is being answered with (see `ExplorerState` and `explorer_pace`).
///
/// The active row is the `tick_ms` setting rather than a constant of its own — the one number that
/// trades how soon a move is answered against what the app costs while it works — so a state is
/// asked for its row rather than carrying its numbers.
pub(super) const DEEP_SLEEP_MS: u64 = 1000; // No Explorer windows - check once per second
pub(super) const LONG_SLEEP_MS: u64 = 500; // All minimized or hidden - check twice per second
pub(super) const MEDIUM_SLEEP_MS: u64 = 150; // Visible but not focused - moderate checking
pub(super) const STATE_RECHECK_DEEP_MS: u64 = 2000; // When no Explorer windows
pub(super) const STATE_RECHECK_LONG_MS: u64 = 1000; // When minimized/hidden
pub(super) const STATE_RECHECK_MEDIUM_MS: u64 = 300; // When visible but not focused
pub(super) const STATE_RECHECK_ACTIVE_MS: u64 = 100; // When active

/// How long the preview loop may go without ticking before the engines it is holding
/// are ended from here.
///
/// The engines an idle tier keeps warm are the preview loop's to end, and a loop that
/// has stopped answering cannot end anything: a document nothing is waiting on, held
/// by a loop that is not running, is the leftover process this app exists not to leave
/// behind (see `engine_processes`). What counts as stopped is the age of the loop's
/// last tick, which it notes as it runs (`preview_window::preview_stall_ms`) — its own
/// waits are a sixteenth of a second with a preview up and half a second idle, so
/// three seconds of quiet is not a wait but work that has not come back, or a loop
/// that is gone.
pub(super) const PREVIEW_STALL_MS: u64 = 3000;
/// How far the pointer has to move before it counts as moved at all, in logical
/// pixels — and the wider distance that counts while the keyboard owns the screen, so
/// a pointer resting on a desk cannot cancel a keyboard preview. Logical distances,
/// scaled by the display the pointer is on: a hand moves the same distance whatever
/// the display it is over is scaled to.
pub(super) const MOUSE_MOVE_PIXELS: f32 = 5.0;
pub(super) const KEYBOARD_POINTER_MOVE_TOLERANCE_PIXELS: f32 = 20.0;
pub(super) const KEYBOARD_PREVIEW_BOX_WATCH_MS: u64 = 2500;
pub(super) const VK_BACK_CODE: i32 = 0x08;
pub(super) const VK_CONTROL_CODE: i32 = 0x11;
pub(super) const VK_MENU_CODE: i32 = 0x12;
pub(super) const VK_T_CODE: i32 = 0x54;
pub(super) const VK_LWIN_CODE: i32 = 0x5B;
pub(super) const VK_RWIN_CODE: i32 = 0x5C;
/// How far up from the element under the pointer the item that holds it is looked
/// for. The item is the nearest list row or data item; what lies between it and
/// the element under the pointer is the view's own chrome — an icon, a label, a
/// row's text — so the walk is short by nature, and it is bounded here anyway.
pub(super) const POINTER_ITEM_ANCESTOR_LIMIT: usize = 8;
/// The most Shell windows that will be asked which one the pointer is in. A
/// collection that reports more than this is not one to walk for every probe.
pub(super) const SHELL_WINDOW_LIMIT: i32 = 64;

pub(super) static EXPLORER_LAST_REAL_FOLDERS: Lazy<Mutex<HashMap<isize, String>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
pub(super) static EXPLORER_WINDOW_CACHE: Lazy<Mutex<HashMap<isize, (bool, Instant)>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// What each folder was last seen sorted by, as the view that is drawing it said so.
///
/// It is here rather than asked of the view on the far side of a click because a click on a
/// caption button is the one moment nothing is hovering: the listing is not being read for a
/// file, so the sort is whatever was read the last time a pointer was over something in it.
/// A header click drops what the resolver holds of the item it was on (`forget_item`), and a
/// sort is read per item, so what is here cannot outlive the sort it describes by more than
/// the folder the pointer is not in.
///
/// The map is keyed by folder rather than by window because the window is not what the
/// question is about: two windows on one folder are one listing's order, and a window with
/// several folders open is several. A pin in a folder nothing has been hovered asks nothing
/// and is walked in name order, which is the order a listing nobody has touched is in.
pub(super) static VIEW_SORT: Lazy<Mutex<HashMap<PathBuf, ViewSort>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How a view's own `Name` column compares two names, and whether a folder's sort can be
/// read at all: a column the system defines and a view that draws items wherever they were
/// dropped.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) struct ViewSort {
    pub(crate) key: SortKey,
    pub(crate) descending: bool,
}

/// The four columns a folder is ordered by that this app can reproduce, which are the four
/// the pin's own buttons can walk in the order the listing shows.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum SortKey {
    Name,
    DateModified,
    Size,
    FileType,
}

/// The sort a folder was last seen in, if it was seen at all and if the view could name one.
///
/// Nothing is asked of a view that cannot be asked: no sort at all is what a `Details` or
/// `List` view in an icon mode answers, and it is a normal answer rather than a failure, so
/// the walk falls back to name order rather than to a guess.
pub(crate) fn view_sort_of(folder: &Path) -> Option<ViewSort> {
    VIEW_SORT.lock().ok()?.get(folder).copied()
}

/// What a folder's sort is remembered as, and how many folders are remembered at once. A pin
/// walks one folder, so the cap is generous rather than tight; it exists so a machine that
/// sweeps the pointer across a drive cannot grow this without end.
pub(super) const VIEW_SORT_LIMIT: usize = 256;

pub(super) fn remember_view_sort(folder: PathBuf, sort: ViewSort) {
    let Ok(mut sorts) = VIEW_SORT.lock() else {
        return;
    };

    // A header click is a new sort of a folder already in here, so it is written in place
    // rather than added beside itself; and a folder is given up when the map is full, which
    // is a list of folders a hand swept across rather than one it is working in. Which of
    // them goes is whatever the map offers: a sort read again is a couple of crossings, and
    // a listing walked again is a folder read, so neither is worth an order kept beside it.
    sorts.retain(|held, _| *held != folder);
    if sorts.len() >= VIEW_SORT_LIMIT {
        if let Some(dropped) = sorts.keys().next().cloned() {
            sorts.remove(&dropped);
        }
    }
    sorts.insert(folder, sort);
}

/// The order a folder is in, asked of the view that is drawing it: how many sort columns it has,
/// the first of them, and whether items are still in the order the sort puts them. Every one of
/// those can say "no" without anything being wrong, and each "no" is an order there is nothing
/// here to reproduce.
pub(super) fn read_view_sort(view: &IFolderView2) -> Option<ViewSort> {
    unsafe {
        let count = view.GetSortColumnCount().ok()?;
        let flags = view.GetCurrentFolderFlags().ok()?;

        let mut columns = [SORTCOLUMN::default()];
        let column = view
            .GetSortColumns(&mut columns)
            .ok()
            .map(|()| (columns[0].propkey, columns[0].direction));

        sort_from_columns(count, flags, column)
    }
}

/// What a view's own answers mean, with the columns read separately so the two questions
/// that can be asked without a view at all are the two that are tested without one.
///
/// * **No columns** is what an icon, tile, list or medium-icon view says, and it is a normal
///   answer: there is nothing to reproduce, and the walk falls back to name order.
/// * **Auto-arrange off** is a folder whose items are where the user put them. Its order is
///   the order they were dragged into, which is not a column and not anything this app could
///   work back out from a file's name, its date or its size.
pub(super) fn sort_from_columns(
    count: i32,
    flags: u32,
    column: Option<(PROPERTYKEY, SORTDIRECTION)>,
) -> Option<ViewSort> {
    if count <= 0 {
        return None;
    }

    if flags & FWF_AUTOARRANGE.0 as u32 == 0 {
        return None;
    }

    let (key, direction) = column?;

    Some(ViewSort {
        key: sort_key_of(&key)?,
        descending: direction == SORT_DESCENDING,
    })
}

/// The four columns a listing can be ordered by that this app walks, written out by their
/// canonical keys rather than imported.
///
/// Every one of them is a property of the `System` property set, and the crate binds that
/// set's keys only behind a further feature this app has no other use for — so the four this
/// walk needs are written here, each with the property id Windows documents for it. The
/// `Type` key is the one the crate does not bind at all.
///
/// They are compared as whole keys rather than matched on a name a locale prints, so a
/// column this app cannot walk is recognised as one it cannot walk in every language.
pub(super) const PKEY_ITEM_NAME_DISPLAY: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xb725f130_47ef_101a_a5f1_02608c9eebac),
    pid: 10,
};
pub(super) const PKEY_SIZE: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xb725f130_47ef_101a_a5f1_02608c9eebac),
    pid: 12,
};
pub(super) const PKEY_DATE_MODIFIED: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xb725f130_47ef_101a_a5f1_02608c9eebac),
    pid: 14,
};
pub(super) const PKEY_FILE_TYPE: PROPERTYKEY = PROPERTYKEY {
    fmtid: GUID::from_u128(0xb725f130_47ef_101a_a5f1_02608c9eebac),
    pid: 26,
};

/// The column a property key names, of the four this app walks, and nothing for any other:
/// a folder sorted by its tags is a folder whose order there is no way to reproduce, and the
/// name order is a better answer than a half-right one.
pub(super) fn sort_key_of(key: &PROPERTYKEY) -> Option<SortKey> {
    if *key == PKEY_ITEM_NAME_DISPLAY {
        Some(SortKey::Name)
    } else if *key == PKEY_DATE_MODIFIED {
        Some(SortKey::DateModified)
    } else if *key == PKEY_SIZE {
        Some(SortKey::Size)
    } else if *key == PKEY_FILE_TYPE {
        Some(SortKey::FileType)
    } else {
        None
    }
}

/// What a view's folder is sorted by, remembered against the folder — or deliberately not,
/// for the two places where a position in the listing means something other than an order.
///
/// The first is a search. A `search-ms:` view holds the results of a query across as many
/// folders as the query reached, and its order is how relevant each result is to what was
/// typed. There is no column behind that, so what a folder's own sort is says nothing about
/// the order the user is looking at, and remembering one would put the pin's buttons in an
/// order no window on the desktop is showing.
///
/// The second is a view whose sort could not be read at all: no columns, or items where they
/// were dropped. Whatever was remembered for the folder before is a sort of a listing that is
/// not this one, so it is taken out rather than left to stand.
pub(super) fn note_view_sort(view: &AnsweredView) {
    let folder = view_folder_path(view);

    match (read_view_sort(&view.folder_view), folder.clone()) {
        (Some(sort), Some(folder)) => remember_view_sort(folder, sort),
        (_, Some(folder)) => {
            if let Ok(mut sorts) = VIEW_SORT.lock() {
                sorts.remove(&folder);
            }
        }
        _ => {}
    }
}

/// The folder a view is showing, which is what a sort is remembered against — and nothing at
/// all for a view opened on a search, since a search has results rather than a folder.
pub(super) fn view_folder_path(view: &AnsweredView) -> Option<PathBuf> {
    unsafe {
        let url = view.browser.LocationURL().ok()?.to_string();
        if is_search_ms_url(&url) {
            return None;
        }

        get_shell_view_folder_path(&view.shell_view).map(PathBuf::from)
    }
}

/// Explorer restarts seen, counted rather than flagged: the hook loop is the one
/// that has to notice one, since it is the thread holding what a restart
/// invalidates, and a count cannot be missed by a reset that would have landed
/// between two of its ticks. Bumped by the tray thread from the `TaskbarCreated`
/// broadcast the shell sends when its taskbar is built again.
pub(super) static EXPLORER_RESTARTS: AtomicU64 = AtomicU64::new(0);

/// How long a Shell collection that could not be built is left alone before it is
/// built again. Building it fails while the shell is still coming up — at logon,
/// and in the moment after a restart — and a collection left missing answers every
/// later lookup the way a dead one does.
pub(super) const SHELL_COLLECTION_RETRY_MS: u64 = 1000;

/// How long previews wait after Explorer has restarted: every window the pointer
/// could be over is the new shell's, and it is still putting them up.
pub(super) const EXPLORER_RESTART_BACKOFF_MS: u64 = 1500;

/// Note that Explorer is not the process it was. Every Shell object this app holds
/// is served by explorer.exe, and the ones taken before a restart are proxies into
/// the process that is gone. Called on the `TaskbarCreated` broadcast, which
/// reaches every top-level window when the shell's taskbar is created again.
pub fn note_explorer_restart() {
    EXPLORER_RESTARTS.fetch_add(1, Ordering::SeqCst);
}

/// Explorer restarts seen so far; the tray thread is the only writer.
pub(super) fn explorer_restart_count() -> u64 {
    EXPLORER_RESTARTS.load(Ordering::SeqCst)
}
