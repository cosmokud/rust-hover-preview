//! A selection made in a text preview: the press, drag and double press that mark it, and the two
//! ways it is taken out again - put on the clipboard, or all of it at once.

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

    // A second press inside the time and box the machine reads a double click in takes the word
    // under the hand; a press with no drag behind it is an empty selection, which is what clears
    // whatever was selected before; and a press with Shift held moves the far end of what is
    // already selected instead, so a range can be grown a click at a time.
    let now = Instant::now();
    let double = state
        .last_press
        .is_some_and(|(then, at)| within_double_click(then, at, now, (x, y)));
    state.last_press = Some((now, (x, y)));
    state.drag_pointer = Some((x, y));

    let at = text_preview::position_in(&state.lines, x, y);
    state.selection = Some(if double {
        text_preview::word_in(&state.lines, x, y)
    } else {
        text_preview::Selection {
            anchor: anchor_for_press(state.selection, shift_is_down(), at),
            caret: at,
        }
    });
    // What a double click took is the selection the hand asked for rather than the start of a
    // drag: arming one would let the first small move of the hand cut the word down to the
    // character it had reached.
    state.selecting = !double;
    true
}

/// Where a press on a text preview anchors its selection: a plain press anchors where it landed,
/// and a press with Shift held keeps the anchor the selection already has, which is the editor's
/// own way of growing a range without dragging.
pub(super) fn anchor_for_press(
    selection: Option<text_preview::Selection>,
    shift: bool,
    at: (usize, usize),
) -> (usize, usize) {
    match selection {
        Some(selection) if shift => selection.anchor,
        _ => at,
    }
}

/// Whether either Shift is held, asked the way Ctrl+C and Ctrl+A are rather than from a message:
/// a pinned preview is never sent its own keystrokes, so a modifier is a question about the
/// keyboard at the moment of the press (see `text_preview_copy_requested`).
fn shift_is_down() -> bool {
    unsafe { GetAsyncKeyState(VK_SHIFT.0 as i32) as u16 & 0x8000 != 0 }
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

    // Where the pointer is is remembered on every move of a drag, whether or not the caret moves
    // with it: a drag carried past the edge of the box is one the page has to follow, and a page
    // that asks the pointer again a tick later would be asking about a hand that has moved on (see
    // `text_drag_step`).
    state.drag_pointer = Some((x, y));

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
        .map(|state| {
            state.drag_pointer = None;
            std::mem::replace(&mut state.selecting, false)
        })
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

/// Select everything: the whole of the document, which is what a Copy of it takes.
///
/// The frame draws only the part of it that fits, so what the highlight covers is what the box
/// shows of the selection — the selection itself reaches past it, and a copy of it does too (see
/// `text_preview::text_of`).
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
/// nothing selected — the whole of the page's document.
pub(super) fn text_preview_clipboard_text() -> Option<String> {
    let (path, options, selection) = CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let state = media.text_state.as_ref()?;

        Some((state.path.clone(), state.options, state.selection))
    })?;

    // Nothing selected is the whole document: a copy with nothing selected is a copy of the page
    // the user is looking at — the same thing `Select All` reaches — and it is the one thing a
    // page whose window is a window nobody may type in can be asked for by a click alone (see
    // `copy_text_preview`).
    let selection = selection
        .filter(|selection| !selection.is_empty())
        .unwrap_or(text_preview::Selection {
            anchor: (0, 0),
            caret: (usize::MAX, usize::MAX),
        });

    text_preview::text_of(&path, options, selection)
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
/// something else is left alone — and a pinned preview that is not the window the
/// user is in is skipped for the same reason, since a chord read from the keyboard is
/// otherwise a chord pressed in whatever window is in front (see `pin_is_focused`).
pub(super) fn text_preview_copy_requested() -> bool {
    if !has_text_selection() {
        return false;
    }

    if pinned() && !pin_is_focused() {
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
/// preview being read and not typed at — and only while that pin is the window the user is in:
/// the chord is read from the keyboard's own state rather than received as a message, so a
/// Ctrl+A pressed in another window is otherwise read as one meant for a pin standing behind it
/// (see `pin_is_focused`).
pub(super) fn pinned_select_all_requested() -> bool {
    if !pin_is_focused() || !has_text_document() {
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

/// Lines a found occurrence is put below the top of a page that was not already showing it: enough
/// of the page above it that a reader arrives at what they looked for rather than at a frame that
/// begins with it (see `find_next_in_text_preview`).
const FIND_LEAD_LINES: usize = 2;

/// Take the pinned page to the next place its own selection appears, and select it: `F3`, and the
/// page's own `Find Next` row, are the same command (see `find_next`).
///
/// A find with nothing to look for, or with nowhere left to look, leaves the page exactly as it is.
/// The reader asked for the occurrence after this one, and a page that jumped back to the top of the
/// document for nothing would be a find that lost their place.
pub(super) unsafe fn find_next_in_text_preview(hwnd: HWND) {
    let Some((path, options, selection, first_line, visible, max_first_line)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media.text_state.as_ref().map(|state| {
                    (
                        state.path.clone(),
                        state.options,
                        state.selection,
                        state.first_line,
                        state.visible_lines,
                        state.max_first_line(),
                    )
                })
            })
        })
    else {
        return;
    };

    let Some(selection) = selection.filter(|selection| !selection.is_empty()) else {
        return;
    };
    let Some(found) = text_preview::find_next(&path, options, selection) else {
        return;
    };

    // Written before the page is asked to draw itself, because what a draw draws is the state as it
    // stands: the frame that comes back carries the occurrence already (see `scroll_text_preview`).
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.selection = Some(found);
        }
    }

    // A page that is already showing the line stays where it is: the page does not move for a find
    // that landed on it, and what moved is the selection.
    if (first_line..first_line + visible).contains(&found.anchor.0) {
        repaint_text_preview(hwnd);
    } else {
        // The lead is what a reader arrives at what they looked for with rather than at a frame that
        // begins with it, and it is given up for a box too short to hold it as well as the line.
        let lead = FIND_LEAD_LINES.min(visible.saturating_sub(1));
        let line = found.anchor.0.saturating_sub(lead).min(max_first_line);
        scroll_text_preview(hwnd, line);
    }
}

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
