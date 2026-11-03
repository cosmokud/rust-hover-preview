//! The caret a pinned page of text holds, and the keys that move it.
//!
//! **The page is a document, not a scroll box.** A reader's `↑`/`↓` move a caret a line at a time
//! and keep the column they were in over the short lines in between, `←`/`→` cross a line end the way
//! an editor's do, `Ctrl` and an arrow take a word by the page's own run rule, `Home`/`End` are the
//! line's two ends and `Ctrl`+`Home`/`Ctrl`+`End` the document's, and `PgUp`/`PgDn` are a screenful
//! (see `text_preview::moved`). `Shift` with any of them grows a selection instead of moving the
//! caret, and a bare move collapses one. The page scrolls **only to keep the caret visible**, which
//! is what tells this from the page keys the same arrows used to be.
//!
//! **The file on disk is never touched, so nothing here may change the text.** The page is a view of
//! a file: a typed character, a `Backspace`, a `Delete`, a `Ctrl`+`X` and a `Ctrl`+`V` all change
//! nothing — deliberately, because a cut that cannot cut is worse than a chord that does nothing,
//! and a page that accepted a keystroke would be a page that lost it at the next hover. `Ctrl`+`A`
//! and `Ctrl`+`C` are the page's (they read, they do not write), and a press on the page, a drag, a
//! double press and `Ctrl`+`F` are untouched (see `text_selection` and `text_find`).
//!
//! **The page answers the caret's keys only once it has the reader's focus.** A pin is drawn
//! unfocused, with no caret and no key of the page's: every walk key — `←`/`→`, `↑`/`↓`,
//! `PgUp`/`PgDn`, `Home`/`End` — steps the folder, exactly as it does over a picture. A press
//! inside the page's body focuses it, which is the press that already put the caret down or began a
//! selection; a press on the pin's caption takes the focus away again — its walk, its buttons and the
//! empty drag area, and nothing else, so the scrollbar, the transport bar, the find bar, an edge, the
//! menu and the wheel all leave the page focused. **Focus is not the caret**: taking it away clears no
//! caret, no column and no selection, and only stops the bar being drawn and the keys being answered,
//! so a page pressed into again comes back where it was left (see `page_has_focus`).
//!
//! **A move is queued rather than made where it lands.** The keys arrive in the window procedure, on
//! the message road, where a document may not be read and a page may not be painted — so the move is
//! written down and the tick that may take the lock runs it (see `settle`). That is the same road the
//! find bar's edits take, and for the same reason.

use super::*;

/// One move a key asked the caret to make, with whether `Shift` was held when it was asked.
///
/// The flag travels with the move rather than being read when the move runs, because the key is
/// pressed at one moment and the tick that answers it is another: a `Shift` let go of in between
/// would otherwise turn an extending move into a collapsing one.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct CaretMove {
    pub(super) motion: text_preview::Motion,
    pub(super) extend: bool,
}

/// What a command pressed on a pinned page asks its caret for, or `None` for one the caret does not
/// own — a pin that holds no page, a page nobody has pressed into, or a command that is not a move.
/// Answering `Some(())` is the page taking the key, which is what tells the page's caret from the
/// folder's walk (see `pin_command_request` and `page_has_focus`).
///
/// The table is the reader's own: the four arrows, the two page keys, the two ends, and the two
/// document ends under `Ctrl`. `Home`/`End` are the *line's* ends and `Ctrl`+`Home`/`Ctrl`+`End` the
/// document's — the opposite of a browser's page keys and the same as every editor's, because the page
/// is a document. `PgUp`/`PgDn` are a screenful, and they keep the caret's column, so a page turned
/// through short lines comes out where a `↓` held down would (see `Motion::ScreenDown`).
///
/// **The command is what the key became, not what the key was.** The pin's keyboard hook swallows
/// these keys on key-down and hands the loop a command, so `VK_LEFT` arrives here as `Previous` and
/// `VK_PRIOR` as `PageBack`; the modifiers are the only thing left to ask the keyboard for, and they
/// are read here rather than at the press because a command carries none (see `Modifiers::polled`
/// and `CaretMove`).
///
/// **Nothing that writes is here.** A letter, a `Backspace`, a `Delete` and the two clipboard chords
/// that would cut and paste are all absent, and `window_proc` swallows them rather than passing them
/// on (see the module documentation).
pub(super) fn command_move(command: PinCommand) -> Option<()> {
    // A page nobody has pressed into has no keys of its own: every one of them is the folder's, which
    // is what `None` here becomes at each of `pin_command_request`'s arms.
    if !holds_a_page() || !page_has_focus() {
        return None;
    }

    let text_find::Modifiers { shift, control } = text_find::Modifiers::polled();
    let screen = visible_lines()?;
    let motion = match command {
        PinCommand::CaretUp => text_preview::Motion::Up,
        PinCommand::CaretDown => text_preview::Motion::Down,
        PinCommand::Previous => {
            if control {
                text_preview::Motion::WordLeft
            } else {
                text_preview::Motion::Left
            }
        }
        PinCommand::Next => {
            if control {
                text_preview::Motion::WordRight
            } else {
                text_preview::Motion::Right
            }
        }
        PinCommand::PageTop => {
            if control {
                text_preview::Motion::DocumentStart
            } else {
                text_preview::Motion::LineStart
            }
        }
        PinCommand::PageBottom => {
            if control {
                text_preview::Motion::DocumentEnd
            } else {
                text_preview::Motion::LineEnd
            }
        }
        PinCommand::PageBack => text_preview::Motion::ScreenUp(screen),
        PinCommand::PageOn => text_preview::Motion::ScreenDown(screen),
        _ => return None,
    };

    note(CaretMove {
        motion,
        extend: shift,
    });

    Some(())
}

/// The lines the box the page is drawn in holds, which is what a screenful is.
fn visible_lines() -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .map(|state| state.visible_lines)
    })
}

/// Write a move down for the tick to make, the way a find bar's edit is written down (see
/// `text_find::note`). Nothing is read of the document here: the window procedure may not.
pub(super) fn note(movement: CaretMove) {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.pending_caret.push(movement);
        }
    }
}

/// Whether the pin on screen is holding a page of text, which is what tells a key that is the page's
/// from one the folder's: a pin holding a page has somewhere for the caret's keys to go, and one
/// holding anything else has not.
///
/// **The page's focus is a second question and this is not it.** A page can be held and unfocused —
/// every page is, until it is pressed into — and until then its keys are the folder's (see
/// `page_has_focus`).
///
/// **The page on the screen is asked first and the media second, and in that order.** What the last
/// paint published is the answer to whether the preview is a page (see `TEXT_PREVIEW_HOLDING`), and a
/// pin whose page is being replaced this tick is still a pin holding one — while a page of text left
/// in the media by the file before it is a page that is no longer standing, and reading that one
/// would turn a walk of the folder into a caret move over a document nobody is looking at.
pub(super) fn holds_a_page() -> bool {
    if !TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
        return false;
    }

    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| !state.lines.is_empty())
        })
        .unwrap_or(false)
}

/// Whether the page on screen has the reader's focus, which is what the caret and the caret's keys
/// are the page's by.
///
/// **Asked beside `holds_a_page` rather than folded into it**, because a page can be held and
/// unfocused: what a key means is the two questions together, and a pin holding no page at all
/// answers this one whatever it says.
///
/// Focus is taken by a press inside the page's body — the same press that puts the caret down or
/// begins a selection (see `text_selection::begin_text_selection`) — and given up by a press on the
/// pin's caption (see `unfocus_page`). Nothing else on the pin takes or gives it.
pub(super) fn page_has_focus() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.focused)
        })
        .unwrap_or(false)
}

/// Give up the page's focus, which is what a press on the pin's caption asks for.
///
/// **Nothing of where the caret stands is cleared**: not the place, not the column a vertical move
/// aims at, not the selection. Only the drawing of the bar and the answering of the keys stop, so a
/// page pressed into again comes back exactly where it was left rather than homed.
pub(super) fn unfocus_page() {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.focused = false;
        }
    }
}

/// Run every move the page was told since the last tick, and draw what changed.
///
/// The moves are run in the order they were pressed, and each one starts where the one before it
/// left the caret — so a hand that pressed `↓` twice in one tick arrives two lines down rather than
/// twice on the first line. A move that changes nothing draws nothing.
///
/// The document is read once for the whole batch rather than once per move: a held arrow on a tick
/// that has been slow is several moves, and styling the document for each of them would be several
/// reads of a file for one paint.
pub(super) unsafe fn settle(hwnd: HWND) {
    let Some((path, options, moves, visible, first_line, selection)) =
        CURRENT_MEDIA.lock().ok().and_then(|mut media| {
            let state = media.as_mut()?.text_state.as_mut()?;
            if state.pending_caret.is_empty() {
                return None;
            }

            Some((
                state.path.clone(),
                state.options,
                std::mem::take(&mut state.pending_caret),
                state.visible_lines,
                state.first_line,
                state.selection,
            ))
        })
    else {
        return;
    };

    // Held: the page is drawn from a document already read, and a move is a question about that
    // page rather than about whether the file has changed since (see `drawn_lines_held`).
    let lines = text_preview::drawn_lines_held(&path, options);
    if lines.is_empty() {
        return;
    }

    // The moves are worked out with the lock let go of, because a page is styled here and the lock is
    // the one a paint takes (see `scroll_text_preview`).
    let mut caret = caret_now();
    let mut goal = caret_goal_now();
    let mut next = None;

    for movement in moves {
        let (place, kept) = text_preview::moved(
            &lines,
            caret.unwrap_or((first_line, 0)),
            movement.motion,
            goal,
        );
        let extended = if movement.extend {
            let anchor = selection.map_or(place, |selection| selection.anchor);
            Some(text_preview::Selection {
                anchor,
                caret: place,
            })
        } else {
            None
        };

        caret = Some(place);
        goal = kept;
        next = Some((place, extended, movement.extend));
    }

    let Some((place, extended, _)) = next else {
        return;
    };

    // What the page owes after the batch: the caret where the moves left it, the selection grown or
    // collapsed, and the page taken far enough that the caret is on it.
    let scrolled = {
        let Ok(mut media) = CURRENT_MEDIA.lock() else {
            return;
        };
        let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
            return;
        };

        state.caret = Some(place);
        state.caret_goal = goal;
        state.selection = extended;
        // Solid from every move, so the reader sees the bar they just moved rather than a blink that
        // happened to be off (see `CaretBlink::restart`).
        state
            .caret_blink
            .restart(Instant::now(), text_find::caret_blink());

        state
            .scroll_to_show(place)
            .filter(|line| *line != state.first_line)
    };

    match scrolled {
        Some(line) => scroll_text_preview(hwnd, line),
        None => {
            // A caret that is already on screen is a mark over the frame, so the frame itself has not
            // changed — and a repaint of the band is what draws the bar (see `paint_text_caret`).
            let _ = visible;
            render_layered_preview(hwnd);
        }
    }
}

/// The caret the page is holding, if it has one.
pub(super) fn caret_now() -> Option<text_preview::Place> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| state.caret)
    })
}

/// The column the page's next vertical move aims at.
fn caret_goal_now() -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| state.caret_goal)
    })
}

/// The caret this page draws this frame, or nothing where the page holds none or the bar is in its
/// dark half of the blink.
///
/// **A selection draws the bar solid at its focus end**, and a page with a caret draws it blinking:
/// a bar that blinked inside a selection would be a bar the reader could not find, and one that
/// blinked over nothing would be a mark with no text to stand in.
pub(super) fn caret_mark() -> Option<text_preview::TextCaret> {
    let media = CURRENT_MEDIA.lock().ok()?;
    let state = media.as_ref()?.text_state.as_ref()?;

    // Nothing is drawn for a page the reader has not pressed into, whether or not it holds a caret and
    // a selection from the last time it did: the bar is the page's focus made visible, and the wash a
    // selection lays is the frame's own and stays drawn (see `page_has_focus`).
    if !state.focused {
        return None;
    }

    let selection = state.selection.filter(|selection| !selection.is_empty());
    let at = match selection {
        Some(selection) => selection.caret,
        None => state.caret?,
    };
    let lit = selection.is_some() || state.caret_blink.lit();

    Some(text_preview::TextCaret { at, lit })
}

/// Whether the caret's blink has turned over since the last time this was asked, which is the one
/// thing about a page that changes with nothing else having happened (see `CaretBlink::advance`).
pub(super) fn the_caret_turned() -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };
    let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
        return false;
    };
    // A selection holds the bar solid, and a page with no caret has no bar at all: neither is a bar
    // that blinks, so neither asks for a repaint twice a second. Nor is one on a page nobody has
    // pressed into, which draws no bar to turn over (see `caret_mark`).
    if !state.focused
        || state.caret.is_none()
        || state
            .selection
            .is_some_and(|selection| !selection.is_empty())
    {
        return false;
    }

    state
        .caret_blink
        .advance(Instant::now(), text_find::caret_blink())
}
