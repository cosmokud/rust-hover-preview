//! The two facts a Common Item Dialog answers that no Explorer window has to: the folder
//! it is showing, and the file an item of it stands for.
//!
//! **Nothing else resolves a dialog's item.** It is a `#32770` in the process of the program
//! that asked for it, so it is nobody's Shell window: `Shell.Windows()` has no entry for its
//! frame, `CDM_GETFOLDERPATH`/`CDM_GETFILEPATH`/`CDM_GETSPEC` answer nothing for it at all,
//! `CDM_GETFOLDERIDLIST` hands back a stale folder, and no Shell view is registered for the
//! frame — so `item_file_path`'s route, which is a position in a view, has no position to ask
//! about. The items themselves are the shell's own drawing: `ElementFromPoint` finds a
//! `UIItem` `ListItem` cross-process exactly as it does in a folder view.
//!
//! What is missing is the folder. An item reports its **display name only** — no path in its
//! name, its value, or any other property — so the answer is the folder the dialog is showing
//! followed by the name the dialog shows, and that is only taken where the file is there to
//! vouch for it (see `dialog_item_path`). A name nothing can vouch for stays unanswered: a
//! search results pane, a virtual folder or a name the dialog truncated is not a file.

use windows::Win32::UI::WindowsAndMessaging::{GetDlgItem, GetWindowTextW};

use super::*;

/// The automation id of the address breadcrumb every Common Item Dialog draws:
/// `ReBarWindow32` > `Breadcrumb Parent` > `ToolbarWindow32`, whose `Name` is
/// `"Address: <full folder path>"`.
///
/// It is read through a property condition rather than out of a cached property, because the
/// batched request every element is read with does not carry `AutomationId` — asking the
/// dialog for the descendants and filtering on that one property would cost a crossing per
/// element, and the id is what says which of them is the address.
pub(super) const THE_BREADCRUMB_AUTOMATION_ID: &str = "1001";

/// What the breadcrumb's `Name` carries in front of the folder it names. The prefix is
/// **required** rather than stripped where it happens to be there: `1001` is not the
/// breadcrumb's own id alone — a `Save As` reuses it for its name edit — and a name that
/// does not say `Address:` is not an address (see `breadcrumb_folder`).
pub(super) const THE_ADDRESS_PREFIX: &str = "Address: ";

/// How many descendants are read out of the dialog before the breadcrumb is given up on:
/// the branch is one toolbar deep, and a dialog that reports more than this is one whose
/// automation tree is not the shell's.
pub(super) const THE_BREADCRUMB_CANDIDATES: i32 = 16;

/// The folder a Common Item Dialog is showing, memoized per frame for `FOLDER_PROBE_MS`.
///
/// It is asked twice for one hover — once for the item's path and once for the place the
/// tick describes — and each ask is a read into a window of another process, so the answer
/// is kept: a pointer resting on a row is that question many times a second, and the folder
/// only moves when the user navigates inside the dialog, which is a place change the TTL
/// and `forget_window_views` both cover (the view caches are dropped on a place change, and
/// the dialog's own memo with them).
///
/// The order is the dialog's own: the name field first, because once a file has been chosen
/// or typed it holds that file's whole path and is the most recent thing the user said; then
/// the breadcrumb, which is the folder the list is actually showing. A dialog that answers
/// neither has no folder, and the hover of its items is unanswered rather than guessed at.
pub(super) fn dialog_folder(resolver: &mut ItemResolver, frame: isize) -> Option<PathBuf> {
    let frame_window = HWND(frame as *mut core::ffi::c_void);
    if frame_window.is_invalid() {
        return None;
    }

    let kept = resolver
        .dialog_folder
        .as_ref()
        .filter(|memo| {
            memo.frame == frame && memo.read_at.elapsed() < Duration::from_millis(FOLDER_PROBE_MS)
        })
        .map(|memo| memo.answer.clone());
    if let Some(answer) = kept {
        return answer;
    }

    let answer = read_dialog_folder(resolver, frame_window);
    resolver.dialog_folder = Some(DialogFolderMemo {
        frame,
        answer: answer.clone(),
        read_at: Instant::now(),
    });

    answer
}

/// The folder one dialog is showing, read the way `dialog_folder` describes.
fn read_dialog_folder(resolver: &ItemResolver, frame: HWND) -> Option<PathBuf> {
    let answer = folder_from_the_name_edit(frame)
        .map(|folder| ("the name field", folder))
        .or_else(|| breadcrumb_folder(resolver, frame).map(|folder| ("the breadcrumb", folder)));

    // Which witness answered, because the two are one question asked of two places and a dialog
    // whose folder came from the wrong one resolves a hover of a name nobody pointed at.
    note_pin_click!(
        "dialog folder  witness {}",
        answer
            .as_ref()
            .map(|(witness, folder)| format!("{witness}  in {}", folder.display()))
            .unwrap_or_else(|| "none".to_string()),
    );

    answer.map(|(_, folder)| folder)
}

/// The folder the dialog's name field names, where it names one.
///
/// The field holds what the user typed or chose, which after an `Open` is the chosen file's
/// own path — so the folder it stands for is that file's parent. It is read from the
/// dialog's own control rather than from its automation tree, because `GetDlgItem` on the
/// id is one call into the dialog's thread and the tree is a walk.
fn folder_from_the_name_edit(frame: HWND) -> Option<PathBuf> {
    unsafe {
        let edit = GetDlgItem(
            frame,
            crate::shell::window_classes::THE_NAME_EDIT_CONTROL_ID,
        )
        .ok()?;
        if edit.is_invalid() {
            return None;
        }

        let mut text = [0u16; 1024];
        let len = GetWindowTextW(edit, &mut text).max(0) as usize;
        let typed = String::from_utf16_lossy(&text[..len.min(text.len())]);
        let typed = typed.trim().trim_matches('"');
        if typed.is_empty() || !is_valid_file_path(typed) {
            return None;
        }

        let path = normalize_existing_path(PathBuf::from(typed))?;
        Some(if path.is_dir() {
            path
        } else {
            path.parent()?.to_path_buf()
        })
    }
}

/// The folder the dialog's address breadcrumb names, where it names one that is there.
///
/// The breadcrumb is found by automation id, because that is the only fact about it the
/// dialog states without the folder: its `Name` is the address, and the toolbar it is drawn
/// in is one element among the buttons, the list and the navigation pane. `Save As` reuses
/// the id for its name edit, so the candidate that *is* a directory is preferred and a
/// candidate that is not is only taken where nothing better was found.
fn breadcrumb_folder(resolver: &ItemResolver, frame: HWND) -> Option<PathBuf> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;

    unsafe {
        let root = automation.ElementFromHandle(frame).ok()?;
        let condition = automation
            .CreatePropertyCondition(
                UIA_AutomationIdPropertyId,
                &VARIANT::from(THE_BREADCRUMB_AUTOMATION_ID),
            )
            .ok()?;
        let found = root
            .FindAllBuildCache(TreeScope_Subtree, &condition, cache)
            .ok()?;
        let count = found.Length().unwrap_or(0);

        let mut folder_that_is_not_there = None;
        for index in 0..count.min(THE_BREADCRUMB_CANDIDATES) {
            let Ok(element) = found.GetElement(index) else {
                continue;
            };
            let Some(name) = element_name(&element) else {
                continue;
            };
            let Some(candidate) = address_from_breadcrumb(&name)
                .and_then(normalize_existing_path)
                .filter(|path| path.is_absolute())
            else {
                continue;
            };

            if candidate.is_dir() {
                return Some(candidate);
            }
            folder_that_is_not_there.get_or_insert(candidate);
        }

        folder_that_is_not_there
    }
}

/// The folder a breadcrumb's `Name` names, as it reads it — the parse, with no `stat` of
/// what comes back.
///
/// The prefix is required rather than optional, and that is the whole of what this is for:
/// `1001` names a toolbar in one dialog and a name field in another, and a path that
/// happens to be typed into the name field is not the dialog's address. A name without the
/// prefix is not an address at all, and nothing about it is answered.
pub(super) fn address_from_breadcrumb(name: &str) -> Option<PathBuf> {
    let name = name.trim();
    let head = name.get(..THE_ADDRESS_PREFIX.len())?;
    if !head.eq_ignore_ascii_case(THE_ADDRESS_PREFIX) {
        return None;
    }

    let path = name.get(THE_ADDRESS_PREFIX.len()..)?.trim();
    if !is_valid_file_path(path) {
        return None;
    }

    Some(PathBuf::from(path))
}

/// The file an item of a Common Item Dialog stands for, or nothing where nothing vouches
/// for it.
///
/// Three answers, and each is a different witness:
///
/// - the name *is* a path — the dialog's own name field is not an item, but a name that
///   carries a whole path was come by the way a search result's value is, and that is
///   taken as given;
/// - the folder the dialog is showing, joined with the name it shows the item under — the
///   ordinary answer, and the one gated to a file this app previews (`normalize_media_path`
///   reads the file's own entry, so the path is only taken where the file is there, is a
///   file, and is a kind this build shows);
/// - a name with no extension that the folder answers with exactly one media file of that
///   stem, which is the shell's hidden-extension listing: an item shown as `holiday` is
///   `holiday.mp4` on disk, and a second `holiday.*` means the name belongs to no one file.
///
/// Anything else answers nothing. The dialog shows a display name and hides extensions,
/// so a name is a hint at a file rather than the file's own name — and a hover answered
/// with the wrong file is worse than no preview at all.
pub(super) fn dialog_item_path(
    resolver: &mut ItemResolver,
    window: &ItemWindow,
    item: &HoveredItem,
) -> Option<PathBuf> {
    // The switch is read here as well as where the window was classified, because a tick that
    // classified a dialog while it was on must answer nothing after it is turned off: `Preview Settings
    // → General → Enable in Browse Dialogs` off means the dialog is no listing, and a preview made in it
    // is the whole of what is refused (see `app::seam::a_dialog_is_previewed`).
    if !crate::app::seam::browse_dialog_previews() {
        return None;
    }
    if let Some(path) = resolve_media_path_from_text(&item.name) {
        return Some(path);
    }

    let name = item.name.trim();
    if name.is_empty() || is_container_name(name) {
        return None;
    }

    let folder = dialog_folder(resolver, window.frame)?;
    let answer = item_path_in_folder(&folder, name);

    // The name the dialog shows against the file it was vouched for as: an unanswered row is the
    // ordinary case (a folder, a name with no extension that belongs to no file) and reads here as
    // `-`, so a run that hovers nothing says which row it was asked about.
    note_pin_click!(
        "dialog item  {}  in {}  -> {}",
        name,
        folder.display(),
        answer
            .as_deref()
            .map_or_else(|| "-".to_string(), |path| path.display().to_string()),
    );

    answer
}

/// The file a name the dialog shows stands for in the folder it is showing.
///
/// Split from `dialog_item_path` for the part of it that is a question about a folder rather
/// than about a window: the folder is read once (see `dialog_folder`) and this is what it is
/// asked, which is the half that can be answered without a dialog in front.
pub(super) fn item_path_in_folder(folder: &Path, name: &str) -> Option<PathBuf> {
    // A folder drawn in the list is a name in the dialog and nothing to preview: its name
    // is not a file's, and a stem search below would answer with a sibling file of the same
    // name — a file nobody pointed at.
    if folder.join(name).is_dir() {
        return None;
    }

    if let Some(path) = normalize_media_path(folder.join(name)) {
        return Some(path);
    }

    // The name has an extension of its own, so a stem search is a search for another file:
    // a dialog shown `holiday.mp4` is `holiday.mp4` or nothing.
    if Path::new(name).extension().is_some() {
        return None;
    }

    sole_media_file_of_stem(folder, name)
}

/// The one media file in `folder` whose stem is `name`, case-insensitively — or nothing
/// where there is none, and nothing where there is a second: two files of one stem are a
/// name that belongs to neither.
fn sole_media_file_of_stem(folder: &Path, name: &str) -> Option<PathBuf> {
    let mut found: Option<PathBuf> = None;

    for entry in std::fs::read_dir(folder).ok()? {
        let Ok(entry) = entry else {
            continue;
        };
        let file_name = entry.file_name();
        let Some(entry_name) = file_name.to_str() else {
            continue;
        };
        let stem = entry_name
            .rsplit_once('.')
            .map_or(entry_name, |(stem, _)| stem);
        if !stem.eq_ignore_ascii_case(name) {
            continue;
        }

        let Some(path) = normalize_media_path(entry.path()) else {
            continue;
        };
        // A second candidate answers nothing at all (see the note above).
        if found.is_some() {
            return None;
        }
        found = Some(path);
    }

    found
}
