//! Where a file is: an Explorer URL turned into a path, the text a shell item answers
//! with, the folder a view has open, and the place a window or a click describes.
//!
//! The lookups are kept per window because a search view answers a different URL each
//! time it is asked, and a path on a share or on a slow disk answers none of them. The
//! second walk for an answer already had is the cost this cache exists not to pay.

use super::*;

/// How far a preview is placed clear of the item it is about, as the tray's `Avoid`
/// submenu and `config.ini` have it.
pub(super) fn avoid_mode() -> AvoidMode {
    CONFIG
        .lock()
        .map(|config| config.avoid_mode)
        .unwrap_or(AvoidMode::Off)
}

/// Whether two paths name the same file.
///
/// What "the same" means is the shell's rather than `Path`'s: Windows paths are not
/// case-sensitive, and a listing that answers with `D:\Pictures\One.PNG` for a pin showing
/// `D:\Pictures\one.png` is answering with the file that is already on screen.
pub(super) fn same_path(a: &Path, b: &Path) -> bool {
    a == b
        || a.as_os_str()
            .encode_wide()
            .map(ascii_lower)
            .eq(b.as_os_str().encode_wide().map(ascii_lower))
}

/// Fold one UTF-16 unit the way `eq_ignore_ascii_case` folds a character, so a
/// path can be compared without being turned into a string first.
pub(super) fn ascii_lower(unit: u16) -> u16 {
    if (b'A' as u16..=b'Z' as u16).contains(&unit) {
        unit + 32
    } else {
        unit
    }
}

pub(super) fn urlencoding_decode(s: &str) -> String {
    let mut bytes = Vec::with_capacity(s.len());
    let mut chars = s.as_bytes().iter().copied().peekable();

    while let Some(byte) = chars.next() {
        if byte == b'%' {
            let hi = chars.next();
            let lo = chars.next();
            if let (Some(hi), Some(lo)) = (hi, lo) {
                let hex = [hi, lo];
                if let Ok(hex_str) = std::str::from_utf8(&hex) {
                    if let Ok(decoded) = u8::from_str_radix(hex_str, 16) {
                        bytes.push(decoded);
                        continue;
                    }
                }
                bytes.push(b'%');
                bytes.push(hi);
                bytes.push(lo);
            } else {
                bytes.push(b'%');
                if let Some(hi) = hi {
                    bytes.push(hi);
                }
            }
        } else if byte == b'+' {
            bytes.push(b' ');
        } else {
            bytes.push(byte);
        }
    }

    String::from_utf8_lossy(&bytes).into_owned()
}

pub(super) fn urlencoding_decode_repeated(s: &str) -> String {
    let mut current = s.to_string();
    for _ in 0..3 {
        let decoded = urlencoding_decode(&current);
        if decoded == current {
            break;
        }
        current = decoded;
    }
    current
}

pub(super) fn is_search_ms_url(url_str: &str) -> bool {
    url_str
        .trim_start()
        .to_ascii_lowercase()
        .starts_with("search-ms:")
}

pub(super) fn normalize_file_url_path(url_str: &str) -> Option<String> {
    let path = if let Some(path) = url_str.strip_prefix("file:///") {
        path.replace('/', "\\")
    } else {
        let path = url_str.strip_prefix("file://")?;
        format!("\\\\{}", path.replace('/', "\\"))
    };

    Some(urlencoding_decode(&path))
}

pub(super) fn normalize_search_location(location: &str) -> Option<String> {
    let decoded = urlencoding_decode_repeated(location);
    let location = decoded.trim();
    if location.is_empty() {
        return None;
    }

    let path = normalize_file_url_path(location).unwrap_or_else(|| location.replace('/', "\\"));
    let path = path.trim().trim_matches('"').to_string();
    if path.is_empty() {
        None
    } else {
        Some(path)
    }
}

pub(super) fn search_ms_location_from_url(url_str: &str) -> Option<String> {
    if !is_search_ms_url(url_str) {
        return None;
    }

    let decoded_url = urlencoding_decode_repeated(url_str);
    for part in decoded_url.split('&') {
        let decoded_part = urlencoding_decode_repeated(part.trim());
        let part_lower = decoded_part.to_ascii_lowercase();

        for prefix in ["crumb=location:", "crumb=folder:"] {
            if let Some(index) = part_lower.find(prefix) {
                let location_start = index + prefix.len();
                return normalize_search_location(&decoded_part[location_start..]);
            }
        }
    }

    None
}

pub(super) fn is_usable_folder_path(path: &str) -> bool {
    let path = path.trim();
    !path.is_empty() && PathBuf::from(path).is_dir()
}

pub(super) fn resolve_media_path_candidate(text: &str) -> Option<PathBuf> {
    let candidate = text.trim().trim_matches(|c| c == '"' || c == '\'');
    if candidate.is_empty() {
        return None;
    }

    let normalized = normalize_file_url_path(candidate).unwrap_or_else(|| candidate.to_string());
    if !is_valid_file_path(&normalized) {
        return None;
    }

    let path = PathBuf::from(normalized);
    // The file's own entry is read once, and it is what says whether there is one to ask
    // about at all (see `crate::formats::head::Facts`).
    if is_media_file(&path) {
        Some(path)
    } else {
        None
    }
}

pub(super) fn resolve_media_path_from_text(text: &str) -> Option<PathBuf> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }

    if let Some(path) = resolve_media_path_candidate(text) {
        return Some(path);
    }

    let chars: Vec<(usize, char)> = text.char_indices().collect();
    for window in chars.windows(3) {
        let [(start, drive), (_, colon), (_, slash)] = window else {
            continue;
        };
        if !drive.is_ascii_alphabetic() || *colon != ':' || (*slash != '\\' && *slash != '/') {
            continue;
        }

        let candidate = &text[*start..];
        if let Some(path) = resolve_media_path_candidate(candidate) {
            return Some(path);
        }

        let mut ends: Vec<usize> = candidate.char_indices().map(|(idx, _)| idx).collect();
        ends.push(candidate.len());
        for end in ends.into_iter().rev() {
            if end <= 3 {
                continue;
            }
            if let Some(path) = resolve_media_path_candidate(&candidate[..end]) {
                return Some(path);
            }
        }
    }

    None
}

pub(super) fn cache_explorer_real_folder(hwnd: isize, folder: &str) {
    if !is_usable_folder_path(folder) {
        return;
    }

    if let Ok(mut cache) = EXPLORER_LAST_REAL_FOLDERS.lock() {
        if !cache.contains_key(&hwnd) && cache.len() >= EXPLORER_REAL_FOLDER_CACHE_MAX_ENTRIES {
            if let Some(old_key) = cache.keys().next().copied() {
                cache.remove(&old_key);
            }
        }
        cache.insert(hwnd, folder.to_string());
    }
}

pub(super) fn get_cached_explorer_real_folder(hwnd: isize) -> Option<String> {
    EXPLORER_LAST_REAL_FOLDERS
        .lock()
        .ok()
        .and_then(|cache| cache.get(&hwnd).cloned())
        .filter(|folder| is_usable_folder_path(folder))
}

pub(super) fn resolve_explorer_location_folder(hwnd: isize, url_str: &str) -> Option<String> {
    if let Some(path) = normalize_file_url_path(url_str) {
        if is_usable_folder_path(&path) {
            cache_explorer_real_folder(hwnd, &path);
            return Some(path);
        }
    }

    if is_search_ms_url(url_str) {
        if let Some(path) = search_ms_location_from_url(url_str) {
            if is_usable_folder_path(&path) {
                cache_explorer_real_folder(hwnd, &path);
                return Some(path);
            }
        }

        // Win11 search can stop exposing a parseable root. Use the last normal
        // folder seen for this exact Explorer window, which matches the second
        // same-folder-window workaround without needing a real second window.
        return get_cached_explorer_real_folder(hwnd);
    }

    None
}

pub(super) fn resolve_search_root_from_context(context: &ActiveShellViewContext) -> Option<String> {
    if let Some(url) = context.location_url.as_deref() {
        if let Some(root) = resolve_explorer_location_folder(context.shell_view_hwnd, url) {
            return Some(root);
        }
    }

    if let Some(folder) = context
        .folder_path
        .as_deref()
        .filter(|folder| is_usable_folder_path(folder))
    {
        cache_explorer_real_folder(context.shell_view_hwnd, folder);
        return Some(folder.to_string());
    }

    get_cached_explorer_real_folder(context.shell_view_hwnd)
}

pub(super) fn is_probable_search_view_context(context: &ActiveShellViewContext) -> bool {
    context
        .location_url
        .as_deref()
        .map(is_search_ms_url)
        .unwrap_or(false)
}

/// The place an item is in, out of the views the window it is drawn in belongs to.
///
/// The window is handed in rather than read here, because the two paths that ask are about
/// two different items: the pointer's own, whose window is what the tick has under it, and the
/// one the keyboard is on, whose window is the one its provider reports (see `ItemWindow` and
/// `FocusedItemInfo::item_window`). Both are the same question once a window is named.
///
/// It stands where a walk of every Shell window the desktop has registered once stood, and
/// what it asks instead is the question the item path already settles a point with: the
/// window given is a window, or a child of one, of exactly one of the views a window holds,
/// and with tabs that is the tab that is showing — so that view, and nothing else, is asked
/// what it is showing (see `ItemWindow` and `frame_views`). The cost is one frame's worth of
/// work rather than the desktop's: the set is in hand for all but the first probe of a
/// window, and what is left is the shell being asked to describe the one view the item is in.
///
/// A window in none of the frame's views is a pointer over the navigation pane, the
/// toolbar or the details pane, none of which belongs to a tab: which tab the frame is
/// showing is then not something it can say, and what is answered for is the view the item
/// was last read inside of that frame — the tab the hand was last working in — or the
/// frame's first view where it has never been inside one (see `WindowViews::anchor`).
/// Any of a frame's views is a guess at that point; what this one has over a view picked at
/// random is that it does not change while the item does not, so one place is read as
/// one place.
///
/// What a probe that cannot be answered at all is left with is nothing, which is what it
/// was left with before: a look that answered no fact is a question to ask again rather
/// than a place that changed (see `HoverLocation`).
pub(super) fn anchored_view_context(
    resolver: &mut ItemResolver,
    window: ItemWindow,
    want_folder: bool,
) -> Option<ActiveShellViewContext> {
    // Read before the frame's views are, because reading them borrows the resolver: what
    // the item was last read inside of this frame is what a probe that is inside none of them
    // is answered by, and it belongs to the set the borrow is about to be taken of.
    let remembered = resolver.remembered_view(window.frame);
    let registrations = shell_window_count(resolver);

    // Timed at the whole of it the way the walk it stands in for was: what the shell is
    // being held for is describing the view, and a set already in hand costs nothing.
    let started = Instant::now();
    let views = frame_views(resolver, window.frame, registrations);
    let anchor = window.view_holding(views);
    let context = anchor
        .or_else(|| remembered.filter(|index| *index < views.len()))
        .and_then(|index| views.get(index))
        .map(|view| view.describe(want_folder));
    note_probe_ms(&PROBE_VIEW_SLOWEST_MS, started.elapsed());

    if let Some(anchor) = anchor {
        note_probe(&PROBE_VIEW_ANCHORED);
        // Remembered for the probes that are in none of the frame's views: it is the tab
        // the hand was last working in, and what such a probe is answered by.
        resolver.remember_view(window.frame, anchor);
    }

    context
}

/// Whether the pointer is inside a window that is the view itself or one of its children.
pub(super) fn hwnd_is_same_or_ancestor(child: HWND, ancestor: HWND) -> bool {
    if child.is_invalid() || ancestor.is_invalid() {
        return false;
    }

    let mut current = child;
    for _ in 0..32 {
        if current == ancestor {
            return true;
        }

        match unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(current) } {
            Ok(parent) if !parent.is_invalid() && parent != current => {
                current = parent;
            }
            _ => break,
        }
    }

    false
}

/// Whether the view under the pointer is a search's results. The hints carry the
/// same answer, but they are read at the folder probe's cadence: a search opened a
/// moment ago is seen here before it is seen there.
///
/// It is the same view the hints are read from — the one the pointer is in — so the two
/// cannot be about different views: what differs between them is when they were read (see
/// `anchored_view_context`).
pub(super) fn is_current_search_view_legacy(
    resolver: &mut ItemResolver,
    pointer: &PointerTick,
) -> bool {
    // Only the URL is read, so the folder is not asked for: it is a walk through the
    // view's own objects, and this check discards everything but the location.
    let Some(window) = item_window_of(pointer.window) else {
        return false;
    };

    anchored_view_context(resolver, window, false)
        .and_then(|context| context.location_url)
        .is_some_and(|url| is_search_ms_url(&url))
}

/// The facts one view's own description holds, as the probes that watch for a change of place
/// read them: the folder it has open, the URL it was opened with, and whether it is a search —
/// whose root is the folder behind the results rather than the view's own.
///
/// The folder is the one fact here that is not free, and it is kept only as a fallback for a
/// search whose root the URL does not name (see `resolve_search_root_from_context`). Everywhere
/// else it is read once and then only compared, and the comparison works on the URL and the
/// view's own handle, because those name the same place for a folder view (see
/// `place_of_window`).
pub(super) fn view_resolver_hints(context: &ActiveShellViewContext) -> HoverResolverHints {
    let is_search_view = is_probable_search_view_context(context);
    let search_root = if is_search_view {
        resolve_search_root_from_context(context)
    } else {
        None
    };

    HoverResolverHints {
        // The search root answers for the folder where the view has none to give: a results
        // view holds no folder of its own, and what is behind it is where its files are.
        current_folder: context.folder_path.clone().or_else(|| search_root.clone()),
        location_url: context.location_url.clone(),
        is_search_view,
        search_root,
        shell_view_hwnd: (context.shell_view_hwnd != 0).then_some(context.shell_view_hwnd),
    }
}

/// What the view under the pointer is showing, for the probes that need to know a
/// location has changed.
///
/// It is answered by the view the pointer is in — the folder it has open and the URL it was
/// opened with — and by nothing else: the resolution of a file does not depend on it, so a
/// view the shell does not describe leaves the hints empty rather than sending the hook
/// looking for another witness.
///
/// The folder is only asked for where it is the fact that decides something. Reading it is
/// five crossings into the shell plus a canonicalize and two `stat`s, paid at this probe's
/// rate, and of the four facts here it is the only one that is not free — the URL names the
/// same place as the folder does for a folder view (see `place_of_window`), and the view's
/// own handle says which of two tabs. So a folder view is described by the URL alone, and
/// only a search asks for the folder: there it is what the root behind the results is
/// resolved out of, and nothing else supplies one.
pub(super) fn get_current_hover_resolver_hints(
    resolver: &mut ItemResolver,
    pointer: &PointerTick,
) -> HoverResolverHints {
    let Some(window) = item_window_of(pointer.window) else {
        return HoverResolverHints::default();
    };

    let Some(context) = anchored_view_context(resolver, window, false) else {
        return HoverResolverHints::default();
    };

    // A search is the one view whose folder is not a place of its own but the root of another,
    // so it is the one view that cannot be described without it. The cheap look has already
    // said which kind of view this is by the time the folder is asked for.
    let context = if is_probable_search_view_context(&context) {
        anchored_view_context(resolver, window, true).unwrap_or(context)
    } else {
        context
    };

    view_resolver_hints(&context)
}

/// The place a window's own view is showing, in the form places are compared in: the view that
/// drew the item, and what that view was showing (see `HoverLocation`). The window is half of the
/// place for the reason two tabs of one window are two places: two tabs can be showing one folder,
/// and a tab switched between them is a move the folder alone cannot see (see
/// `hover_location_changed`).
///
/// The folder a view *has open* is deliberately not asked for. It is a walk out through the
/// shell's own objects to a filesystem path and a `stat` of what comes back, while the URL the view
/// was opened with is answered every time by the browser object the view was found through — and it
/// names the same place as the folder for a folder view (a search answers with its own query, and
/// its root is resolved out of that). The place is read once per item the focus lands on, and a key
/// being held lands it on one every few dozen milliseconds.
pub(super) fn place_of_window(
    resolver: &mut ItemResolver,
    window: ItemWindow,
) -> Option<HoverLocation> {
    let context = anchored_view_context(resolver, window, false)?;

    Some(HoverLocation::of(&view_resolver_hints(&context)))
}

/// The place a click was made in, as the view under the pointer describes itself: the listing a
/// click held for a retry is measured against (see `PendingClick::place`).
///
/// A shell that described nothing is `None` rather than an empty place: a click whose listing cannot
/// be named is left to its own look at the point, and a comparison against a place nobody answered
/// is what the place rule exists to avoid (see `hover_location_changed`).
pub(super) fn click_place(
    resolver: &mut ItemResolver,
    pointer: &PointerTick,
) -> Option<HoverLocation> {
    let place = place_of_window(resolver, item_window_of(pointer.window)?)?;

    place.was_answered().then_some(place)
}

/// Whether a point is inside a box, read the way a window reads one: the right and
/// bottom edges are outside it. It is the reading the published item box is compared
/// with the pointer through (see `pointer_item_holds`), asked here of a box the hook is
/// holding itself — the item one answer was read from.
pub(super) fn point_in_box(point: POINT, bounds: (i32, i32, i32, i32)) -> bool {
    let (left, top, right, bottom) = bounds;
    point.x >= left && point.x < right && point.y >= top && point.y < bottom
}

pub(super) fn normalize_existing_path(path: PathBuf) -> Option<PathBuf> {
    if !path.exists() {
        return None;
    }

    std::fs::canonicalize(&path).ok().or(Some(path))
}

pub(super) fn normalize_media_path(path: PathBuf) -> Option<PathBuf> {
    // One reading of the file's own entry answers every question this asks about it: that it is
    // there, that it is a file, what version it is at, and whether its content is on this
    // machine. What follows is the form the rest of the app works in — and the file is not
    // asked about again for any of it (see `crate::formats::head::Facts`).
    let facts = crate::formats::head::Facts::read(&path)?;

    if !facts.is_file() || facts.needs_download() || !is_media_file_with_facts(&path, &facts) {
        return None;
    }

    std::fs::canonicalize(&path).ok().or(Some(path))
}

/// The path a Shell item stands for, in the form the rest of the app works in.
pub(super) fn shell_item_to_path(shell_item: &IShellItem) -> Option<PathBuf> {
    unsafe {
        let display_name = shell_item
            .GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING)
            .ok()?;
        let path_string = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));
        let path_string = path_string?;

        normalize_existing_path(PathBuf::from(path_string))
    }
}

pub(super) fn get_shell_view_folder_path(shell_view: &IShellView) -> Option<String> {
    let folder_view = shell_view.cast::<IFolderView>().ok()?;

    unsafe {
        let persist_folder = folder_view.GetFolder::<IPersistFolder2>().ok()?;
        let pidl = persist_folder.GetCurFolder().ok()?;
        let shell_item = SHCreateItemFromIDList::<IShellItem>(pidl).ok();
        CoTaskMemFree(Some(pidl as *const core::ffi::c_void));
        shell_item
            .and_then(|shell_item| shell_item_to_path(&shell_item))
            .filter(|path| path.is_dir())
            .map(|path| path.to_string_lossy().into_owned())
    }
}
