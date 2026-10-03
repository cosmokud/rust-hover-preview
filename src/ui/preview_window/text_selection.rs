//! A selection made in a text preview: the press and drag that mark it, what it covers, and
//! the two ways it is taken out again - put on the clipboard, or all of it at once.

use super::*;

/// Start selecting from a point in the preview. Answers whether there was anything
/// to select, which is what tells a press on a text preview from a press on one of
/// the other formats.
pub(super) fn begin_text_selection(x: i32, y: i32) -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
        return false;
    };
    if state.lines.is_empty() {
        return false;
    }

    // A press with no drag behind it is an empty selection, which is what clears
    // whatever was selected before.
    let at = text_preview::position_in(&state.lines, x, y);
    state.selection = Some(text_preview::Selection {
        anchor: at,
        caret: at,
    });
    state.selecting = true;
    true
}

/// Move the end of a selection to a point. Answers whether a drag is running, so
/// the caller knows whether it is the one holding the capture.
pub(super) fn extend_text_selection(x: i32, y: i32) -> Option<bool> {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return None;
    };

    let state = media.as_mut()?.text_state.as_mut()?;
    if !state.selecting {
        return Some(false);
    }

    let at = text_preview::position_in(&state.lines, x, y);
    let changed = state
        .selection
        .map(|selection| selection.caret != at)
        .unwrap_or(false);

    if changed {
        if let Some(selection) = state.selection.as_mut() {
            selection.caret = at;
        }
    } else {
        return Some(false);
    }

    Some(true)
}

/// End a selection drag. Answers whether one was running.
pub(super) fn end_text_selection() -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| std::mem::replace(&mut state.selecting, false))
        .unwrap_or(false)
}

/// Whether the text preview on screen has something selected, which is what makes
/// a Ctrl+C the preview's to answer.
pub(super) fn has_text_selection() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .and_then(|state| state.selection)
                .map(|selection| !selection.is_empty())
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

/// Select everything the frame shows: the whole of a document that fits on one
/// page, and the screenful a longer one is showing.
///
/// The range is the one a Copy with nothing selected already takes, so what the
/// highlight covers is what that Copy puts on the clipboard.
pub(super) unsafe fn select_all_text_preview(hwnd: HWND) {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return;
    };

    let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
        return;
    };
    if state.lines.is_empty() {
        return;
    }

    state.selection = Some(text_preview::Selection {
        anchor: (0, 0),
        caret: (usize::MAX, usize::MAX),
    });

    // The repaint reads the same state through its own lock, so the guard goes
    // back before it is asked to draw.
    drop(media);

    repaint_text_preview(hwnd);
}

/// What a Copy takes from the preview on screen: what is selected, or — with
/// nothing selected — everything the frame shows.
pub(super) fn text_preview_clipboard_text() -> Option<String> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let state = media.text_state.as_ref()?;

        Some(match state.selection {
            Some(selection) if !selection.is_empty() => {
                text_preview::text_in(&state.lines, selection)
            }
            _ => text_preview::frame_text(&state.lines),
        })
    })
}

/// Copy the text preview to the clipboard, answering whether there was anything to
/// copy.
///
/// The clipboard is owned by one window at a time, so this opens it, hands over a
/// movable block of UTF-16, and closes it again; a failure at any step leaves the
/// previous clipboard contents alone.
pub(super) unsafe fn copy_text_preview(hwnd: HWND) -> bool {
    let Some(text) = text_preview_clipboard_text().filter(|text| !text.is_empty()) else {
        return false;
    };

    if OpenClipboard(hwnd).is_err() {
        return false;
    }

    let _ = EmptyClipboard();

    let units: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let mut copied = false;

    if let Ok(block) = GlobalAlloc(GMEM_MOVEABLE, units.len() * std::mem::size_of::<u16>()) {
        let target = GlobalLock(block) as *mut u16;
        if target.is_null() {
            let _ = GlobalFree(block);
        } else {
            std::ptr::copy_nonoverlapping(units.as_ptr(), target, units.len());
            let _ = GlobalUnlock(block);

            if SetClipboardData(CF_UNICODETEXT.0 as u32, HANDLE(block.0)).is_err() {
                let _ = GlobalFree(block);
            } else {
                // The clipboard owns the block now.
                copied = true;
            }
        }
    }

    let _ = CloseClipboard();
    copied
}

/// Ask for Ctrl+C to copy the preview's selection.
///
/// The preview never takes focus, so a keystroke never arrives as a message: it is
/// read the way the rest of the app reads the keys that belong to it, by asking
/// whether Ctrl is down and C has been pressed since the last time this was asked.
/// Nothing is read at all unless there is a selection, so a Ctrl+C meant for
/// something else is left alone.
pub(super) fn text_preview_copy_requested() -> bool {
    if !has_text_selection() {
        return false;
    }

    unsafe {
        let control = GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0;
        let just_pressed = GetAsyncKeyState(VK_C.0 as i32) & 1 != 0;
        control && just_pressed
    }
}

/// Whether Select All has just been asked for with the keyboard, over a text preview that is
/// pinned.
///
/// The key is polled rather than waited for, for the reason Ctrl+C's is: a preview the user
/// has not pressed is a window nobody is in, and a window nobody is in is never sent a
/// keystroke. It is answered for a pin alone — a pinned text preview is a window the user put
/// there to work in, and full mode gave it the selection this makes use of, while a hover is a
/// preview being read and not typed at.
pub(super) fn pinned_select_all_requested() -> bool {
    if !pinned() || !has_text_document() {
        return false;
    }

    unsafe {
        let control = GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0;
        let just_pressed = GetAsyncKeyState(VK_A.0 as i32) & 1 != 0;
        control && just_pressed
    }
}

/// Whether the media on screen is a text preview with a document in it: what a selection needs
/// something to be made against, whatever it is currently holding (see `text_state`).
pub(super) fn has_text_document() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .is_some()
        })
        .unwrap_or(false)
}

/// What the preview's own menu offers: the whole frame selected, and what is
/// selected put on the clipboard.
pub(super) const ID_TEXT_PREVIEW_SELECT_ALL: usize = 1;
pub(super) const ID_TEXT_PREVIEW_COPY: usize = 2;

/// Show the preview's context menu at a point in window coordinates and act on
/// what it returns.
///
/// The menu is asked for its command rather than posting one, so it needs no
/// message loop of its own and the preview never has to take focus to be worked
/// with.
pub(super) unsafe fn show_text_preview_menu(hwnd: HWND, x: i32, y: i32) {
    let Ok(menu) = CreatePopupMenu() else {
        return;
    };

    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_TEXT_PREVIEW_SELECT_ALL,
        w!("Select All"),
    );
    let _ = AppendMenuW(menu, MF_STRING, ID_TEXT_PREVIEW_COPY, w!("Copy"));

    let mut point = POINT { x, y };
    let _ = ClientToScreen(hwnd, &mut point);

    // A menu belongs to the window in front of it, and this one is never in front:
    // asking for the foreground is what lets the menu see the click that chooses
    // from it.
    let _ = SetForegroundWindow(hwnd);

    let command = TrackPopupMenu(
        menu,
        TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN,
        point.x,
        point.y,
        0,
        hwnd,
        None,
    );

    let _ = DestroyMenu(menu);

    match command.0 as usize {
        ID_TEXT_PREVIEW_SELECT_ALL => select_all_text_preview(hwnd),
        ID_TEXT_PREVIEW_COPY => {
            copy_text_preview(hwnd);
        }
        _ => {}
    }
}
