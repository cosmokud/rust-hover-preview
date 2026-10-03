//! The Shell side of the same question: the views a window holds, the item at an index
//! in one of them, and the path that item names.
//!
//! Two ways to one file, because the UI Automation answer is the one that answers for
//! a listing and the Shell one is the one that answers for everything else — a search
//! result, a virtual folder, a pane drawn by something that is not a list at all.

use super::*;

/// The window an item is resolved in, from the window under the pointer.
///
/// Both handles come out of the one walk: the window under the pointer, and the root it
/// belongs to. The root is the frame whose views can be holding the item — a Shell view
/// has to belong to the window the pointer is over for the items it draws to be the ones
/// under it — and the window under the pointer is what says which of those views is
/// drawing it (see `ItemWindow`).
///
/// The window is handed in rather than read here, because it is the tick's own reading of
/// what the pointer is over: the item a tick resolves and the place it describes are about
/// that one window, and a second reading of it would be a second answer to one question
/// (see `PointerTick`).
pub(super) fn item_window_of(window: HWND) -> Option<ItemWindow> {
    unsafe {
        if window.is_invalid() {
            return None;
        }

        let frame = GetAncestor(window, GA_ROOT);
        if frame.is_invalid() {
            return None;
        }

        Some(ItemWindow {
            frame: frame.0 as isize,
            drawn_in: window.0 as isize,
        })
    }
}

/// Every Shell view registered for a window, deduplicated by the view's own
/// identity.
///
/// A window that holds several tabs registers one Shell window per tab, and every
/// one of them answers with the frame's own window — so the frame names a set of
/// views rather than one, and something else has to say which of them is showing
/// what the item is: the window the item is drawn in (see `ItemWindow`). Nothing
/// here skips a tab that is not showing. This walk does not ask a view anything, it
/// collects them, and which of them answers is settled afterwards.
///
/// What it costs is the reason it is kept rather than made per probe: every
/// registration is crossed into for the view it holds and for the window that view
/// is drawn in, so a window holding tabs costs a few calls per tab — and a window
/// that holds eight of them is not one to walk three dozen times a second. The
/// caller that keeps the set is `frame_views`.
pub(super) fn folder_views_for_window(
    resolver: &ItemResolver,
    frame: isize,
    registrations: Option<i32>,
) -> Vec<AnsweredView> {
    let mut candidates: Vec<AnsweredView> = Vec::new();
    let Some(shell_windows) = resolver.shell_windows.as_ref() else {
        return candidates;
    };
    note_probe(&PROBE_VIEW_WALKS);

    unsafe {
        // The count the caller is already holding is the one used: asking the
        // collection for it again is another crossing into the shell for a number
        // that is already in hand. It is read here only for a caller that had none,
        // and a caller in that position is a collection that could not answer.
        let count = match registrations {
            Some(registrations) => registrations,
            None => match shell_windows.Count() {
                Ok(count) => count,
                Err(_) => return candidates,
            },
        };

        for index in 0..count.min(SHELL_WINDOW_LIMIT) {
            note_probe(&PROBE_VIEW_WINDOWS);
            let Ok(dispatch) = shell_windows.Item(&VARIANT::from(index)) else {
                continue;
            };
            let Ok(browser) = dispatch.cast::<IWebBrowser2>() else {
                continue;
            };
            let Ok(handle) = browser.HWND() else {
                continue;
            };
            let browser_window = HWND(handle.0 as *mut core::ffi::c_void);
            if browser_window.0 as isize != frame {
                continue;
            }
            if !IsWindowVisible(browser_window).as_bool() || IsIconic(browser_window).as_bool() {
                continue;
            }

            let Ok(service_provider) = browser.cast::<IServiceProvider>() else {
                continue;
            };
            let shell_browser: IShellBrowser =
                match service_provider.QueryService(&SID_STopLevelBrowser) {
                    Ok(shell_browser) => shell_browser,
                    Err(_) => continue,
                };
            let shell_view = match shell_browser.QueryActiveShellView() {
                Ok(shell_view) => shell_view,
                Err(_) => continue,
            };
            let Ok(identity) = shell_view.cast::<IUnknown>() else {
                continue;
            };
            let view_identity = Interface::as_raw(&identity);
            if candidates
                .iter()
                .any(|candidate| candidate.view_identity == view_identity)
            {
                continue;
            }

            // The window the view is drawn in, which is what says which of a frame's
            // views an item is drawn by. A view that will not name one is kept with no
            // window rather than dropped: it can still answer the item question, which
            // is all this walk asked it for before the window was read.
            let view_hwnd = shell_view
                .GetWindow()
                .map(|hwnd| hwnd.0 as isize)
                .unwrap_or_default();

            let Ok(folder_view) = shell_view.cast::<IFolderView2>() else {
                continue;
            };

            candidates.push(AnsweredView {
                view_identity,
                view_hwnd,
                browser,
                shell_view,
                folder_view,
            });
        }
    }

    candidates
}

/// How many Shell windows are registered, which is what says whether a window
/// could be holding tabs.
///
/// The number is read once and carried: the walk that looks a window's views up
/// needs the same count, and asking the collection for it a second time would be a
/// second crossing into the shell for it — see `folder_views_for_window`.
pub(super) fn shell_window_count(resolver: &ItemResolver) -> Option<i32> {
    unsafe { resolver.shell_windows.as_ref()?.Count().ok() }
}

/// The views a window holds, from the set kept for it where that set still describes
/// the window, and read from the Shell window collection where it does not.
///
/// Keeping the set is what turns this path's cost from one paid per probe into one paid
/// per window: reading it is a crossing into the shell for every Shell window the
/// desktop holds — a few of them each, and one more for every tab a window is holding —
/// while asking one of the views it holds for an item is a call or two. What is read
/// again is read again for one of two reasons, and they are the only two that can leave
/// the set describing a window that is no longer there: the window the set was read for
/// is gone, hidden or minimized, or the desktop holds a different number of Shell
/// windows than it did — which is the one cheap number that moves when a tab or a window
/// is opened or closed, and the only thing that ever adds a view to a window or takes
/// one away. Everything else a view does leaves the set alone: navigating is the same
/// view holding another folder, and which of a frame's views is showing what the item
/// is, is a question about the window the item is drawn in rather than about the set
/// (see `ItemWindow`).
///
/// What is *not* kept is a walk that found no view at all. A window showing a folder has
/// a view in it, so no views means the shell was met between two of them — a window that
/// has just opened, or a folder change with the old view let go of and the new one not
/// up yet — and a set of no views is not an answer, it is the absence of one: it can be
/// asked for no item and it describes no place, and kept it would go on answering that
/// for as long as the window was up, since neither of the two reasons above follows from
/// it (the count has not moved, and the window is as live as it ever was). What that
/// costs is a walk per caller while the shell is between views, which is what the walk
/// is there for; what it saves is every preview in that window until the next tab, the
/// next window, or the window being minimized (see `item_file_path`).
pub(super) fn frame_views(
    resolver: &mut ItemResolver,
    frame: isize,
    registrations: Option<i32>,
) -> &[AnsweredView] {
    let kept = resolver.window_views.as_ref().is_some_and(|set| {
        set.frame == frame
            && registrations.is_none_or(|count| count == set.registrations)
            && set.is_live()
    });

    if kept {
        note_probe(&PROBE_VIEW_SETS_KEPT);
    } else {
        // Timed at the walk rather than around the probe: what the shell is being held
        // for is this, and an answer already in hand costs nothing worth timing.
        let started = Instant::now();
        let views = folder_views_for_window(resolver, frame, registrations);
        note_probe_ms(&PROBE_VIEW_SLOWEST_MS, started.elapsed());

        // A walk that found none is a shell between two views rather than a set to keep,
        // and what is left in hand is nothing so that the next caller reads again (see
        // the note on an empty walk above).
        resolver.window_views = (!views.is_empty()).then(|| WindowViews {
            frame,
            // A collection that will not say how many windows it holds leaves the number
            // of views it produced, which the next probe's count will disagree with — the
            // set is read again then, which is what a count that could not be read is
            // worth.
            registrations: registrations.unwrap_or(views.len() as i32),
            // The views were found again, so the one the pointer was last inside of the
            // old set says nothing about this one.
            anchor: None,
            views,
        });
    }

    match resolver.window_views.as_ref() {
        Some(set) => set.views.as_slice(),
        None => &[],
    }
}

/// The file an item stands for, asked of the view that is drawing it.
///
/// The view belongs to a window — the frame the pointer is over, or the one the focused
/// item is drawn in — and a window that holds tabs registers one Shell window per tab,
/// all of them answering with the frame's own window, so the frame names a set of views
/// and not one. Which of them it is, is settled in two steps that must not be confused
/// with each other.
///
/// First the view: the window the item is drawn in is a window of exactly one of the
/// frame's views — the tab that is showing — and that view, and nothing else, is asked
/// (see `ItemWindow`). A view answers about its own folder's items whether it is showing
/// or not, so a tab the item is not in is one whose answer is about something else: it
/// is not asked, and its agreement is not waited for. What that leaves is a pointer in
/// none of them — over the navigation pane, the toolbar, the details pane, none of which
/// belongs to a tab — and there the views are told apart the way all of them were before
/// the item's own window was read: each is asked, and what they answer has to agree.
///
/// A set that cannot answer for the window the item is drawn in — no view of it claims that
/// window, or the one that does says the item at that position is another name — is read
/// again once before the item is answered with nothing. The item is drawn in this window, so
/// some view of it holds the item, and what answers otherwise is a set read while the window
/// was between the two views of a folder change: the view the folder was left from, which
/// goes on answering for its own folder and which nothing else can tell from a right one (see
/// `frame_views`).
///
/// Then the item: a candidate view is asked whether the item at that position is the
/// item we are on, by name and nothing else. That is an identity question, and a folder
/// answers it exactly as a file does — what the item *is* says nothing about which view
/// holds it. Then, and only for the views that claimed the item, the file: the path the
/// Shell hands over, gated to a file this app previews. A view that holds the item but
/// has no file to show it (a folder, an archive, a document) is a *match* with nothing
/// to preview, not a view that failed to match — treating it as the latter is how
/// another tab's file gets shown while a folder is hovered. What several matches do has
/// to agree: two tabs showing the same folder are one answer, while tabs that disagree —
/// about the file, or about whether there is one at all — are a question the item cannot
/// settle, and no answer is better than the wrong tab's file.
pub(super) fn item_file_path(
    resolver: &mut ItemResolver,
    window: &ItemWindow,
    item: &HoveredItem,
) -> Option<PathBuf> {
    let index = item.index? - 1;
    let registrations = shell_window_count(resolver);

    // An item with no name cannot be told from another tab's item, so only a window
    // showing a single view can be answered without one.
    if item.name.is_empty() && registrations != Some(1) {
        return None;
    }

    // The view the item is drawn in answers alone. A set that cannot answer for the window
    // the item is drawn in is read again, once, before what it answered is taken: no view
    // of it claims that window at all, or the one that does says the item at that position
    // is another name. Neither is an answer the item can be given — the item is drawn in
    // this window, so some view of it holds the item — and both are what a set read while
    // the window was between two views answers with: the walk met the shell as a folder was
    // being changed, and what it collected was the view the folder was left from. Nothing
    // else tells that set from a right one — a folder probe asks its place of the same
    // view, so the view it describes as the place is the one that has been left — which is
    // why the re-read is here rather than in the walk: it is the item that says the set is
    // wrong, and every look at the item makes the question askable again (see `frame_views`).
    // What the re-read costs is one walk of a collection this is holding for exactly that.
    for attempt in 0..2 {
        let (asked, anchored, unclaimed) = {
            let views = frame_views(resolver, window.frame, registrations);
            match window.view_holding(views) {
                Some(position) => {
                    note_probe(&PROBE_VIEW_ANCHORED);
                    note_view_sort(&views[position]);
                    (
                        Some(view_item(&views[position].folder_view, index, &item.name)),
                        Some(position),
                        false,
                    )
                }
                // No view of the set claims the window the item is drawn in. Whether that is
                // the set having been read wrong — which is what the re-read below is for — or
                // a set whose views cannot be told apart by their windows at all: a set that
                // names no window can be matched against the item's by nothing, and one whose
                // windows the item is inside of more than once is a reading that is not
                // trusted rather than one of several (see `ItemWindow`). Both answer through
                // the questions below, so neither is read again.
                None => (
                    None,
                    None,
                    views.iter().any(|view| view.view_hwnd != 0)
                        && !views.iter().any(|view| window.draws_inside(view.view_hwnd)),
                ),
            }
        };

        // The view the item was drawn in is the view the pointer is in, which is what a
        // probe that is in none of them is answered by: it is remembered here as well as by
        // the hints, because a hand sweeping a list is this path many times a second and a
        // folder probe once every few hundred milliseconds (see `WindowViews::anchor`).
        if let Some(anchored) = anchored {
            resolver.remember_view(window.frame, anchored);
        }

        match asked {
            Some(ViewItem::File(path)) => return Some(path),
            // A view that holds the item and has no file to show for it answered what the
            // item is: a folder, an application, a name no kind claims. That is an answer
            // and not a reason to ask a tab the pointer is not in.
            Some(ViewItem::NoFile) => return None,
            // A view that will not answer, and a view that answers that the item at that
            // position is another name, are the two answers a set that does not describe
            // this window gives (see above): both are read again, once.
            Some(ViewItem::Unaskable | ViewItem::NotHeld) if attempt == 1 => return None,
            Some(ViewItem::Unaskable | ViewItem::NotHeld) => {
                // The item an answer was read from goes with the set: an answer read
                // through the views that have been left is about the item of the folder
                // that has been left, which the pointer may be standing on all the same.
                resolver.forget_window_views();
                resolver.forget_item();
            }
            None if attempt == 0 && unclaimed => {
                resolver.forget_window_views();
                resolver.forget_item();
            }
            None => break,
        }
    }

    let mut matches = 0usize;
    let mut matched_without_file = 0usize;
    let mut answer: Option<PathBuf> = None;
    let mut disagreed = false;

    {
        let views = frame_views(resolver, window.frame, registrations);
        for candidate in views {
            match view_item(&candidate.folder_view, index, &item.name) {
                // A view that could not be asked answered nothing, which is how it was
                // counted before the views could be told apart: another view may still
                // answer, and if none does the item is not one of theirs.
                ViewItem::Unaskable | ViewItem::NotHeld => continue,
                ViewItem::NoFile => {
                    matches += 1;
                    matched_without_file += 1;
                }
                ViewItem::File(path) => {
                    matches += 1;
                    match &answer {
                        None => answer = Some(path),
                        Some(existing) if !same_path(existing, &path) => disagreed = true,
                        // Another tab showing the same folder is the same answer.
                        Some(_) => {}
                    }
                }
            }
        }
    }

    if matches == 0 {
        return None;
    }

    // Tabs that do not agree about what the item is — one holding a file, another
    // holding a folder, a name that is in both at that position — cannot be told
    // apart, and either answer could be the wrong tab's.
    if matched_without_file > 0 && (matches > 1 || disagreed) {
        return None;
    }
    if disagreed {
        return None;
    }

    answer
}

/// The file system path the Shell holds for an item — the path the item *is*,
/// rather than a name it is shown under. An item that stands for no file on disk
/// (a library, a drive, a search root) has none, which is the Shell's own answer
/// that there is nothing here to preview.
pub(super) fn shell_item_filesystem_path(item: &IShellItem) -> Option<PathBuf> {
    unsafe {
        let display_name = item.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
        let path_string = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));
        let path_string = path_string?;
        let path = PathBuf::from(path_string);

        (!path.as_os_str().is_empty()).then_some(path)
    }
}

/// Whether the name a view shows an item under is the name that was asked about.
/// Explorer labels a file with the name its view displays — the extension hidden
/// when the user has chosen to hide it — so the item's own display name is what
/// the asked-about name is compared with.
pub(super) fn item_display_name_matches(item: &IShellItem, expected_name: &str) -> bool {
    unsafe {
        let Ok(display_name) = item.GetDisplayName(SIGDN_NORMALDISPLAY) else {
            return false;
        };
        let name = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));

        name.map(|name| name.trim().eq_ignore_ascii_case(expected_name.trim()))
            .unwrap_or(false)
    }
}

/// What a view answers about the item at a position: whether it holds it, and what the
/// item is if it does.
///
/// The identity question comes first and is nothing else: the name the view shows the
/// item under against the name the accessibility tree reports for it. What the item *is*
/// is not part of it — a folder goes by its name exactly as a file does, and a view that
/// holds a folder has to be seen as holding the item, or the tab that owns it abstains
/// and another tab's file answers in its place. An item with no name to ask about is
/// taken as held, which the caller has already established can only be asked of a window
/// showing one view.
///
/// Then the file, and only for the views that claimed the item: the path the Shell holds
/// for the item at that position (`SIGDN_FILESYSPATH`, which for a search result is the
/// real file wherever it lives), gated to a file this app previews. A folder reaches
/// this point and stops here — held, with nothing to preview — which is what keeps the
/// preview of a folder from being another tab's file.
///
/// Both answers are read off one item object, which is the whole of why the two
/// questions are one function: the item at a position is what each of them is about, and
/// fetching it twice was a crossing into the shell paid twice for every view asked.
pub(super) fn view_item(folder_view: &IFolderView2, index: i32, expected_name: &str) -> ViewItem {
    // A position that is not one is the answer the caller has already established it can
    // be given: an item with no name is only asked about of a window showing one view,
    // and what such a view is asked for is the file at that position.
    if index < 0 {
        return ViewItem::NoFile;
    }

    unsafe {
        let item = match folder_view.GetItem::<IShellItem>(index) {
            Ok(item) => item,
            Err(_) => return ViewItem::Unaskable,
        };

        if !expected_name.is_empty() && !item_display_name_matches(&item, expected_name) {
            return ViewItem::NotHeld;
        }

        match shell_item_filesystem_path(&item).and_then(normalize_media_path) {
            Some(path) => ViewItem::File(path),
            None => ViewItem::NoFile,
        }
    }
}

/// The file the pointer is over, resolved once per point and kept while the pointer
/// stays inside the item it was read from.
///
/// The pointer asks one question — what is under me — and the view under it
/// answers by identity: the item the accessibility provider says the point is
/// inside, the position that item holds in the view, and the file that position
/// stands for. Nothing is looked up by name for the pointer, because a search
/// across folders is full of names that belong to more than one file and a name
/// is the one thing the view does not need. What follows the identity route is the
/// same answer asked of the item itself: the accessible value it carries, when
/// that value is a whole path.
///
/// What the answer is kept against is the item rather than the point, and that is what a
/// hand sweeping a list is answered from: a `Details` row is as wide as the view, so a
/// pointer moving along one is a new point on every tick and the same item on every one
/// of them, and a file it has already been answered for is not asked about again while
/// the pointer stays in the item — and in the window — that answer was read in. Everything
/// that makes the item under a parked pointer a new question drops the answer with it
/// (`forget_item`).
pub(super) fn get_file_under_cursor(
    resolver: &mut ItemResolver,
    pointer: &PointerTick,
) -> Option<PathBuf> {
    let point = pointer.point;

    // A tick asks about the same point more than once — the move path asks for
    // the file it latched and then for the one on screen — and the answer is the
    // same both times. Nothing is carried past the tick: the list under a parked
    // pointer may have moved on by the next one.
    if let Some(answer) = resolver.probed_at(point) {
        return answer;
    }

    // The window the pointer is in is read before the answer is, because it is half of
    // what that answer is kept against: a tab switched under a parked pointer is another
    // view drawing another folder at the same place, and the box alone cannot say the
    // item under the pointer is the one this answer was read from. It is the window the
    // tick already has, and not a second reading of a pointer that may have moved since
    // (see `PointerTick`).
    let window = item_window_of(pointer.window);
    let drawn_in = window
        .as_ref()
        .map(|window| window.drawn_in)
        .unwrap_or_default();

    if let Some(answer) = resolver.item_under(point, drawn_in) {
        note_probe(&PROBE_ITEM_MEMO_HITS);
        return answer;
    }

    let look = resolve_file_under_cursor(resolver, point, window.as_ref());
    resolver.remember_probe(point, look.path.clone());
    resolver.remember_item(&look);

    look.path
}

/// The file the pointer is over.
///
/// Two witnesses and no third: the item the pointer is on, turned into a file by
/// the view that is showing it, and the item's own accessible value when that
/// value is a whole path. Nothing is looked up by name, nothing is walked, and a
/// name that nothing can vouch for is left unanswered rather than guessed at — a
/// search across folders is full of names that belong to more than one file.
///
/// What it answers with is the file *and* the item it was read from, because the two
/// leave together: the window and the box are what the pointer stays inside for the file
/// to still be the one under it (see `AnsweredItem`), and a look that found no item at all
/// answers for neither. The window is handed in rather than read here, because the caller
/// reads it before the answer it is holding is asked about.
pub(super) fn resolve_file_under_cursor(
    resolver: &mut ItemResolver,
    point: POINT,
    window: Option<&ItemWindow>,
) -> PointerLook {
    let drawn_in = window.map(|window| window.drawn_in).unwrap_or_default();

    note_probe(&PROBE_POINTER_RESOLUTIONS);
    let Some(item) = uia_item_from_point(resolver, point, false) else {
        return PointerLook {
            path: None,
            item_bounds: None,
            drawn_in,
        };
    };

    // The item the pointer is on, as the box the view draws it in, published for the
    // preview thread to hold a reveal to and for this loop to read a move off: a
    // pointer outside that box has left the item, whatever a distance says. It is
    // published with every look at the item under the pointer — this one and the one
    // the walk below makes — so what it holds is where the pointer is now rather than
    // where the hover on screen was resolved from (see
    // `preview_window::publish_pointer_item_box`).
    //
    // What is published is the part of that box the view actually shows: the item's box
    // as it comes back from the walk is already the part of it that exists on screen
    // (see `view_bounds`), so a pointer over the toolbar above a clipped item, or off the
    // window below one, is outside it and has left the item — where the box the provider
    // drew carries on saying it has not.
    let bounds = (
        item.bounds.left,
        item.bounds.top,
        item.bounds.right,
        item.bounds.bottom,
    );
    publish_pointer_item_box(bounds);

    // The view is asked twice: a wheel turns the list under a parked pointer, and an
    // item that is no longer at the point the first answer described is not what that
    // answer is about. The second look is made only where there is an answer to confirm:
    // what a view could not answer for is not a file this loop has to take back.
    let path = window
        .and_then(|window| item_file_path(resolver, window, &item))
        .filter(|_| {
            uia_item_from_point(resolver, point, false)
                .map(|again| again.same_item(&item))
                .unwrap_or(false)
        });

    // What the item says about itself: a search result carries the file's own path
    // in its accessible value, and a name that is a whole path was come by the same
    // way.
    let path = path
        .or_else(|| item.value.as_deref().and_then(resolve_media_path_from_text))
        .or_else(|| resolve_media_path_from_text(&item.name));

    PointerLook {
        path,
        item_bounds: Some(bounds),
        drawn_in,
    }
}

/// The region a preview of the file under the pointer is kept off, as the `Avoid`
/// setting has it for the item that file is — with whether that region is a column of
/// the view, which is what says a placement steps off it to the side rather than over
/// or under the item (see [`HoveredItem::avoid_box`]).
///
/// Asked when a preview is about to be shown rather than with every probe. A probe
/// answers which file the pointer is on, and that answer is what the rest of the loop
/// runs on; where that file's name is drawn is a question only a preview asks, and it
/// is asked here so that a pointer merely sweeping over a list pays nothing for it. A
/// walk that finds no item at all leaves the preview placed as it always was.
pub(super) fn avoid_box_under_cursor(
    resolver: &ItemResolver,
    point: POINT,
) -> Option<((i32, i32, i32, i32), bool)> {
    // Read before the walk as well as inside `avoid_box`: with the setting off there
    // is no region to be had, so the item is never asked for one.
    if avoid_mode() == AvoidMode::Off {
        return None;
    }

    let item = uia_item_from_point(resolver, point, true)?;
    item.avoid_box()
}
