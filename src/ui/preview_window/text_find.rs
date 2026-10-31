//! The find bar over a pinned page of text: the query it holds, the places that query was found,
//! and the road a key or a press takes to reach them.
//!
//! **The bar is state of the page rather than a window of its own.** A page of text is the one media
//! a pin's own keyboard already reaches, and the bar floats over the page the way the volume popup
//! floats over a film — a card drawn into the pin's own pixels, with no window of its own to place
//! and no caret to lift out of the page (see `pin_chrome::paint_find_bar` and
//! `pin_chrome_strips`). So the page keeps the keyboard, the pin keeps the work, and the bar is one
//! more thing the page's state says.
//!
//! **A key is queued rather than searched for where it lands.** The message that brings a character
//! is delivered in the window procedure, where the input queue is waiting on it, and a search styles
//! the document from its first line; so the bar's edits are written down in the page's state and run
//! by the tick that may take the lock (see `settle`). The keyboard hook's commands travel the same
//! road for a different reason of their own: a low-level callback may not take a lock at all.
//!
//! Nothing here paints the bar: what the frame draws of the matches that are on screen is
//! `text_preview`'s, and the card over it is the chrome's (see `FindHighlights` and
//! `pin_chrome::paint_find_bar`).

use super::*;

/// Whether the bar is up over the page on screen.
pub(super) fn is_open() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.find.open)
        })
        .unwrap_or(false)
}

/// Open the bar over the pinned page on screen, or ask the same search of a bar that is already
/// open: `Ctrl`+`F` over a bar that holds a query refocuses it rather than throwing the query away,
/// because the reader who reaches for the key again is reaching for the bar they were just in.
///
/// The caret has nowhere to move to — the bar's query is its only field, so being up is being
/// focused — and the query is searched again from the line the reader is on either way, because a
/// bar closed and opened again comes back at the query it was left with rather than at its matches
/// (see `FindState`).
pub(super) fn open_text_find() {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
        else {
            return;
        };

        find.open = true;
        find.pending.push(FindEdit::SearchAgain);
    }
}

/// Write down one thing the bar was told, for the tick that can run the search it asks for.
pub(super) fn note(edit: FindEdit) {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
        {
            // Only a bar that is up takes an edit: a key delivered after it went away is a key for
            // the page, and the page's own road answers it (see `window_proc`).
            if find.open {
                find.pending.push(edit);
            }
        }
    }
}

/// Whether `Ctrl`+`F` has just been asked of a pin, which is the bar's key the way `Ctrl`+`A` is
/// Select All's.
///
/// The key is polled rather than waited for, for the reason `Ctrl`+`A`'s is: a pin the user has not
/// pressed is a window nobody is in, and a window nobody is in is never sent a keystroke. It is
/// answered only while the pin is the window the user is in, because the chord is read from the
/// keyboard's own state rather than received as a message — a `Ctrl`+`F` pressed in another window
/// would otherwise open a bar in a pin standing behind it (see `pin_is_focused`).
pub(super) fn pinned_find_requested() -> bool {
    if !pin_is_focused() || !text_selection::has_text_document() {
        return false;
    }

    unsafe {
        let control = GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0;
        let just_pressed = GetAsyncKeyState(VK_F.0 as i32) & 1 != 0;
        control && just_pressed
    }
}

/// The character a `WM_CHAR` is, where the bar is to type it: `wparam` is one UTF-16 code unit, so a
/// surrogate is half of a character this bar does not hold — BMP only — and it is dropped rather
/// than typed as half of one. A control character is not a printable one, which is what keeps a
/// backspace, a return and a `Ctrl` chord from arriving twice: they are answered as keys, and
/// `TranslateMessage` has already turned each into a character by the time the bar reads them.
pub(super) fn typed_character(wparam: usize) -> Option<char> {
    char::from_u32(wparam as u32).filter(|character| !character.is_control())
}

/// What one key pressed on the pin's own window asks of a bar that is up, or nothing for a key the
/// bar only swallows (see `window_proc`).
///
/// `Enter` and the down arrow step on and `Shift`+`Enter` and the up arrow step back, which is the
/// order every find bar a reader has used before this one has. The page's own keys — the two page
/// keys, its two ends and a Space — ask nothing here and are swallowed above, and `F3` asks for the
/// same step its own key always has, over the bar's query rather than over the page's selection.
pub(super) fn key_edit(vk: i32) -> Option<FindEdit> {
    if vk == VK_ESCAPE.0 as i32 {
        return Some(FindEdit::Close);
    }

    if vk == VK_RETURN.0 as i32 {
        return Some(FindEdit::Step(if text_selection::shift_is_down() {
            -1
        } else {
            1
        }));
    }

    if vk == VK_BACK.0 as i32 {
        return Some(FindEdit::Deleted);
    }

    if vk == VK_DOWN.0 as i32 || vk == VK_F3.0 as i32 {
        return Some(FindEdit::Step(1));
    }

    if vk == VK_UP.0 as i32 {
        return Some(FindEdit::Step(-1));
    }

    None
}

/// Answer one command the keyboard hook took while the page had the keys, where a bar is up.
/// Answers whether the bar took it.
///
/// The bar owns **every** key while it is up, so a command it does not name is still a command it
/// swallows: the pin's walk, its two page keys, its two ends, a Space and a subtitle step are all
/// inert under a query, and letting one of them through would turn the page or walk the folder out
/// from under a hand that is typing.
///
/// The two arrows are read as the page keys the pin's table calls them — `PageBack` and `PageOn` —
/// because that table is the hook's and is deliberately not changed for a bar: `VK_UP` and `VK_PRIOR`
/// are one command there, and so are `VK_DOWN` and `VK_NEXT`, so a page key steps the matches as the
/// arrow of the same direction does. What the bar will not do is turn the page, which is the page
/// key's own work.
pub(super) fn take_hooked_key(command: PinCommand) -> bool {
    if !is_open() {
        return false;
    }

    match command {
        PinCommand::PageBack => note(FindEdit::Step(-1)),
        PinCommand::PageOn | PinCommand::FindNext => note(FindEdit::Step(1)),
        PinCommand::Close => note(FindEdit::Close),
        _ => {}
    }

    true
}

/// What the bar has been asked since the tick last looked, taken out of the page in one go so that a
/// character arriving while the search runs is queued for the next tick rather than half-applied in
/// this one.
struct Settled {
    open: bool,
    closed: bool,
    search: bool,
    step: i32,
    path: PathBuf,
    options: TextPreviewOptions,
    query: String,
    case_sensitive: bool,
    whole_word: bool,
}

/// Run everything the bar was told since the last tick: type or delete, flip a switch, step through
/// the matches, close — and then search the document again where the query or a switch changed,
/// take the page to the match the reader is on, and paint what changed.
///
/// A new query starts the reader at the first match at or after the line they were reading, and
/// wraps to the first of all where the document has none below it: a search that put them back at
/// the top of the file with every keystroke would take them away from the paragraph they are in. A
/// step keeps the place it has reached and wraps at either end (see `stepped`).
///
/// The search runs outside the lock and the page is written only where it is still the same document
/// when it comes back: a walk onto another file while a query was being searched is a page whose
/// matches are not this one's, and its state went with the file (see `keep_text_place`).
pub(super) unsafe fn settle(hwnd: HWND) {
    let Some(settled) = take_settled() else {
        return;
    };

    if settled.closed {
        // The page does not move: a close is the highlights going, and what was found is kept by
        // nothing once the bar is not up (see `FindState`).
        render_layered_preview(hwnd);
        return;
    }

    if !settled.open {
        return;
    }

    if settled.search {
        let matches = text_preview::find_all(
            &settled.path,
            settled.options,
            &settled.query,
            settled.case_sensitive,
            settled.whole_word,
        );

        let Ok(mut media) = CURRENT_MEDIA.lock() else {
            return;
        };
        let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
            return;
        };
        if state.path != settled.path {
            return;
        }

        state.find.matches = matches;
        state.find.current = current_at_or_after(&state.find.matches, state.first_line);
    }

    if settled.step != 0 {
        if let Ok(mut media) = CURRENT_MEDIA.lock() {
            if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
                let current = state.find.current;
                let total = state.find.matches.len();
                state.find.current = stepped(current, settled.step, total);
            }
        }
    }

    // The page is taken to the match when it is not already showing it, and the frame that comes
    // back carries the highlights either way: a step inside the box is a repaint, one outside it is a
    // scroll (see `show_found_in_text_preview`).
    let at = CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| {
                state
                    .find
                    .current
                    .and_then(|current| state.find.matches.get(current))
                    .copied()
            })
    });

    let painted = at.is_some_and(|at| text_selection::show_found_in_text_preview(hwnd, at));
    if !painted {
        // Nothing was found to go to — an empty query, or one the document does not hold — so the
        // page stays where it is and only the card over it changed (see `pin_chrome::paint_find_bar`).
        render_layered_preview(hwnd);
    }
}

/// Take the bar's edits out of the page and apply the ones that need no document: a typed character,
/// a backspace and the two switches change the query the search will read, and a step is kept for
/// after the search has replaced the matches it counts through.
fn take_settled() -> Option<Settled> {
    let mut media = CURRENT_MEDIA.lock().ok()?;
    let state = media.as_mut()?.text_state.as_mut()?;
    let find = &mut state.find;

    if find.pending.is_empty() {
        return None;
    }

    let mut settled = Settled {
        open: find.open,
        closed: false,
        search: false,
        step: 0,
        path: state.path.clone(),
        options: state.options,
        query: String::new(),
        case_sensitive: false,
        whole_word: false,
    };

    for edit in std::mem::take(&mut find.pending) {
        match edit {
            FindEdit::Typed(character) => {
                find.query.push(character);
                settled.search = true;
            }
            FindEdit::Deleted => {
                find.query.pop();
                settled.search = true;
            }
            FindEdit::ToggleCase => {
                find.case_sensitive = !find.case_sensitive;
                settled.search = true;
            }
            FindEdit::ToggleWholeWord => {
                find.whole_word = !find.whole_word;
                settled.search = true;
            }
            FindEdit::SearchAgain => settled.search = true,
            FindEdit::Step(delta) => settled.step += delta,
            FindEdit::Close => {
                find.open = false;
                find.matches.clear();
                find.current = None;
                settled.closed = true;
                settled.open = false;
                // Nothing after a close is read: a key pressed after the bar went away is a key for
                // the page, and the next tick's window procedure answers it there.
                break;
            }
        }
    }

    settled.query = find.query.clone();
    settled.case_sensitive = find.case_sensitive;
    settled.whole_word = find.whole_word;
    settled.open = find.open;

    Some(settled)
}

/// The match a new query starts the reader on: the first one at or after the line they were reading,
/// and the first of all where the document has none below it (see `settle`).
fn current_at_or_after(matches: &[text_preview::Selection], line: usize) -> Option<usize> {
    if matches.is_empty() {
        return None;
    }

    Some(
        matches
            .iter()
            .position(|found| found.anchor.0 >= line)
            .unwrap_or(0),
    )
}

/// The match a step from `current` lands on, wrapping at either end, or nothing where the document
/// holds no match at all (see `settle`).
fn stepped(current: Option<usize>, delta: i32, total: usize) -> Option<usize> {
    if total == 0 {
        return None;
    }

    Some(match current {
        Some(at) => (at as i64 + delta as i64).rem_euclid(total as i64) as usize,
        None if delta < 0 => total - 1,
        None => 0,
    })
}
