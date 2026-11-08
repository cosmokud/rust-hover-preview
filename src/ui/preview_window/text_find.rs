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
//! by the tick that may take the lock (see `settle`).
//!
//! **A message is the only road a key has while the bar is up.** The keyboard hook stands down from
//! the pin's keys for as long as a bar is up (see `key_input::publish_find_bar_up`), because a key it
//! swallowed would never become the `WM_CHAR` a text field types from — and `key_edit_with` below is
//! the one table that reads what a key that does arrive means. A second table keyed on the hook's
//! commands is what let a `Space` and a `T` be eaten before they could become characters.
//!
//! Nothing here paints the bar: what the frame draws of the matches that are on screen is
//! `text_preview`'s, and the card over it is the chrome's (see `FindHighlights` and
//! `pin_chrome::paint_find_bar`).

use super::*;

/// How long a chevron is held before it starts stepping on its own, and how often it steps after that:
/// the delay and the rate a key held at Windows' own default repeat is given, because a hand kept on a
/// chevron is asking for what a hand kept on a key asks for.
pub(super) const HOLD_DELAY_MS: u64 = 350;
pub(super) const HOLD_REPEAT_MS: u64 = 90;
/// How many steps one tick may make up after a stall. The repeat is a cadence and not a budget: a tick
/// that arrived a second late owes a hand the step it missed, not a walk through half the page.
const HOLD_CATCH_UP: usize = 3;

/// How many queries the run keeps to step back through, and the text one step puts in the field.
///
/// The list is what `↑` and `↓` walk, an editor's find widget's own history: the last fifty queries
/// this run searched, most recent first, with a query that is already in it moved to the front rather
/// than repeated (see `remember_the_query` and `step_the_history`).
const HISTORY_LIMIT: usize = 50;

/// The queries this run has searched, most recent first — the history `↑` and `↓` walk in a find
/// bar's field (see `step_the_history`).
///
/// **The run's rather than the page's.** A page's state goes with the file it belongs to, so a reader
/// searching the same word in the next document would find an empty list there; an editor's find
/// widget keeps one history for the session, which is what this is.
static SEARCHED: Lazy<Mutex<Vec<String>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Forget every query this run has searched, which a test of the history does before it asks: the list
/// is the run's rather than a page's, so every test that searched into it shares one (see `SEARCHED`).
#[cfg(test)]
pub(super) fn forget_the_history() {
    if let Ok(mut searched) = SEARCHED.lock() {
        searched.clear();
    }
}

/// Remember a query the reader has finished asking: what a step through the matches, an emptied field
/// and a closed bar commit — the moments an editor's find widget writes its history (see
/// `step_the_history`).
///
/// Nothing is written for an empty query, and a query already in the list is moved to the front rather
/// than added again: the list is what a hand reaches back through, and the same word twice in a row is
/// one entry.
fn remember_the_query(query: &str) {
    let query = query.trim();
    if query.is_empty() {
        return;
    }

    let Ok(mut searched) = SEARCHED.lock() else {
        return;
    };
    searched.retain(|kept| kept != query);
    searched.insert(0, query.to_string());
    searched.truncate(HISTORY_LIMIT);
}

/// Step the field back through the queries this run has searched, or forward again to the one the
/// reader was typing.
///
/// **`↑` from a query being typed keeps it.** The first step back stores what the field held as the
/// draft, so `↓` returns to it — a history that swallowed a half-typed query would be a search the
/// reader lost for pressing the wrong arrow. `↑` at the oldest entry stays on it, and `↓` with the
/// field already showing the draft does nothing, which is the two ends of the list.
///
/// The entry is put in the field the way a seed is — whole, one line, all of it selected — so the next
/// character typed is the new query rather than an append to the one stepped to (see
/// `text_find_field::Field::seed`).
fn step_the_history(find: &mut FindState, delta: i32) {
    let history = SEARCHED
        .lock()
        .map(|searched| searched.clone())
        .unwrap_or_default();
    if history.is_empty() {
        return;
    }

    let Some(at) = find.history_at else {
        // A step forward with nothing stepped back is the end of the list the reader is already at.
        if delta > 0 {
            return;
        }

        find.history_draft = Some(find.field.text().to_string());
        if !find.field.seed(&history[0]) {
            return;
        }
        find.history_at = Some(0);
        return;
    };

    // `-1` is a step back, and the list is most recent first: a step back is a step *up* the list.
    let next = at as i32 - delta;
    if next < 0 {
        // Back past the newest entry is the query the reader was typing, which is where the list ends.
        let draft = find.history_draft.take().unwrap_or_default();
        if draft.is_empty() {
            find.field.clear();
        } else {
            find.field.seed(&draft);
            find.field.select_all();
        }
        find.history_at = None;
        return;
    }

    let Some(entry) = history.get(next as usize) else {
        return;
    };
    if find.field.seed(entry) {
        find.history_at = Some(next as usize);
    }
}

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

/// Where the bar's card is and what it says: the card against the page's top-right corner, under the
/// caption, and the query, count and switches it is drawn with.
///
/// It is asked of the pin and the page rather than of the paint, and by both the paint and a press, which
/// is what makes the card that is drawn the card a hand answers — the same road the volume popup's
/// panel takes (see `pinned_volume_geometry`).
pub(super) struct FindBarPlacement {
    pub(super) card: pin_chrome::FindBar,
    pub(super) state: pin_chrome::FindBarPaint,
    /// The scale the card is drawn at: what a press is turned into a place in the query by, because
    /// the query's own columns are measured in the same pixels its glyphs are drawn in.
    pub(super) dpi: u32,
}

/// Where the bar is and what it says, or nothing while it is down or the pin holds something other
/// than a page (see `FindBarPlacement`).
pub(super) fn find_bar() -> Option<FindBarPlacement> {
    let pinned = pin_state()?;
    find_bar_over(pinned.pin()?)
}

/// The card over a pin already in hand, asked of the pin itself rather than through `pin_state`: the
/// tick holds that lock while it decides what the pin's chrome owes, and the same lock asked twice is a
/// window that stops answering (see `refresh_pin_chrome`).
fn find_bar_over(pin: &PinnedPreview) -> Option<FindBarPlacement> {
    let (page_top, dpi) = {
        if pin.collapsed {
            return None;
        }

        let height = pin.window_size().1;
        let (top, _) = pinned_band_rows(
            height,
            pin.caption,
            pinned_transport_height(pin.dpi, pin.transport_bar),
            pin.overlay,
        );
        (top, pin.dpi)
    };

    let media = CURRENT_MEDIA.lock().ok()?;
    let state = media.as_ref()?.text_state.as_ref()?;
    let find = &state.find;
    if !find.open {
        return None;
    }

    Some(FindBarPlacement {
        // The count's own column is laid out from the count itself: `1210/19000` is wider than the
        // room a four-figure count was given, and a count cut at its last digits is a count that is
        // wrong (see `pin_chrome::find_bar_layout`).
        card: pin_chrome::find_bar_layout(
            state.width as i32,
            state.height as i32,
            page_top,
            dpi,
            &pin_chrome::find_bar_count_text(find.current, find.matches.len()),
        ),
        state: pin_chrome::FindBarPaint {
            query: find.field.text().to_string(),
            caret: find.field.caret(),
            anchor: find.field.anchor(),
            caret_lit: find.blink.lit(),
            current: find.current,
            total: find.matches.len(),
            case_sensitive: find.case_sensitive,
            whole_word: find.whole_word,
            hovered: find.hovered,
            pressed: find.pressed,
        },
        dpi,
    })
}

/// A press on the find bar over the page, answering whether it landed on the card, and what the card
/// was asked of it.
///
/// A press on the card's own face — the row the query is drawn in, the gaps between its controls —
/// is the bar's as well as one on a control: the card floats over the text, so nothing under it is
/// what the hand meant. A press in the query places the caret, which is the whole of what makes the
/// field a field rather than a picture of one (see `press_the_field`); a press on the rest of the
/// face asks nothing. What a control asks for is queued like a key's, for the same reason: the
/// search styles the document and the tick is what may take that work (see `settle`).
///
/// A press is also a hand on a button, which the card draws and, on the two chevrons, steps on while
/// it lasts (see `take_the_press` and `step_the_hold`), and a press in the query is a drag until
/// the release comes (see `drag`).
pub(super) fn press(x: i32, y: i32) -> bool {
    let Some(placed) = find_bar() else {
        return false;
    };
    if !placed.card.holds(x, y) {
        return false;
    }

    let part = placed.card.part_at(x, y);
    take_the_press(part);

    match part {
        Some(pin_chrome::FindBarPart::Case) => note(FindEdit::ToggleCase),
        Some(pin_chrome::FindBarPart::WholeWord) => note(FindEdit::ToggleWholeWord),
        Some(pin_chrome::FindBarPart::Previous) => note(FindEdit::Step(-1)),
        Some(pin_chrome::FindBarPart::Next) => note(FindEdit::Step(1)),
        Some(pin_chrome::FindBarPart::Close) => note(FindEdit::Close),
        None if placed.card.holds_query(x, y) => press_the_field(&placed, x, y),
        None => {}
    }

    true
}

/// A press in the query: the caret at the character under the pointer, the word under it where the
/// presses have run together into a double click, or the selection grown from its old anchor where
/// `Shift` is held.
///
/// **The place is measured from the column the query is really drawn at.** A query wider than its
/// field is scrolled to keep its caret in view, so a press turned into a place from the field's left
/// edge would land away from the character the hand is on (see `pin_chrome::query_index_at`).
///
/// The double click is read rather than delivered: the pin's window class is not registered with
/// `CS_DBLCLKS`, so a second press is a second `WM_LBUTTONDOWN` and two presses are one gesture by
/// the time and the box Windows counts them in — the same road a double press on the page takes
/// (see `begin_text_selection`).
fn press_the_field(placed: &FindBarPlacement, x: i32, y: i32) {
    let now = Instant::now();
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
        .filter(|find| find.open)
    else {
        return;
    };

    // A hand on the query is the field focused, and a `Ctrl`+`C` after it copies the query (see
    // `FindFocus`).
    find.focused = FindFocus::Field;

    let at = pin_chrome::query_index_at(
        &placed.card,
        find.field.text(),
        find.field.caret(),
        x,
        placed.dpi,
    );
    let run = find
        .click_run
        .filter(|(then, was, _)| within_double_click(*then, *was, now, (x, y)))
        .map_or(1, |(_, _, presses)| presses.saturating_add(1));
    find.click_run = Some((now, (x, y), run));

    match (run, text_selection::shift_is_down()) {
        // A second press takes the word the pointer is on, by the page's own run rule: letters,
        // digits and the underscore are one run, and a mark is its own (see
        // `text_preview::run_at`). A third press is the same word again rather than a whole line —
        // the query is one line, so there is no line above it to take.
        (2.., _) => {
            let (start, end) = find.field.word_at(at);
            find.field.place_caret(start, false);
            find.field.place_caret(end, true);
        }
        (_, true) => find.field.place_caret(at, true),
        (_, false) => find.field.place_caret(at, false),
    }

    find.dragging = true;
    find.blink.restart(now, caret_blink());
}

/// Carry a press in the field on: the selection grows to the place under the pointer, and answers
/// whether the card changed, which is what a move repaints on.
///
/// Nothing is repainted for a move that stays inside the character the selection already ends on: a
/// drag is many messages a second and most of them do not move the caret at all (see
/// `pinned_mouse_move`, which repaints on the same rule).
pub(super) fn drag(x: i32) -> bool {
    let Some(placed) = find_bar() else {
        return false;
    };

    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
        .filter(|find| find.open && find.dragging)
    else {
        return false;
    };

    let at = pin_chrome::query_index_at(
        &placed.card,
        find.field.text(),
        find.field.caret(),
        x,
        placed.dpi,
    );
    let changed = find.field.place_caret_changed(at, true);
    if changed {
        find.blink.restart(Instant::now(), caret_blink());
    }
    changed
}

/// How long the caret is drawn before it blinks the other way: the system's own caret rate where it
/// has one, so that the field blinks at the pace the reader's own text fields do.
///
/// `GetCaretBlinkTime` answers zero where a caret is not to blink at all and `u32::MAX` where it
/// could not be asked; both are a caret this field still draws, at the ordinary Windows pace.
pub(super) fn caret_blink() -> Duration {
    const ORDINARY_BLINK_MS: u64 = 530;

    let asked = unsafe { GetCaretBlinkTime() };
    let millis = u64::from(asked);
    let fallback = if millis == 0 || asked == u32::MAX {
        ORDINARY_BLINK_MS
    } else {
        millis
    };

    Duration::from_millis(fallback.max(1))
}

/// Where the pointer is on the card, asked on every move a pin is sent: the control under it is drawn
/// lit, and answers whether the card changed.
///
/// **Only a control coming lit or going out is worth a repaint.** A repaint of this pin is a band that
/// copies the whole page frame, so a move that stays on the one control — or that is anywhere off the
/// card at all — changes nothing that is drawn and repaints nothing. The answer is what
/// `pinned_mouse_move` and the tick's chrome refresh repaint on (see `pinned_mouse_move`).
pub(super) fn hover(x: i32, y: i32) -> bool {
    let Some(pinned) = pin_state() else {
        return false;
    };
    match pinned.pin() {
        Some(pin) => hover_over(pin, x, y),
        None => false,
    }
}

/// The same question of a pin already in hand, for the tick that holds the pin's lock (see
/// `find_bar_over`).
pub(super) fn hover_over(pin: &PinnedPreview, x: i32, y: i32) -> bool {
    let Some(placed) = find_bar_over(pin) else {
        return false;
    };
    let part = placed.card.part_at(x, y);

    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
    else {
        return false;
    };
    if find.hovered == part {
        return false;
    }

    find.hovered = part;
    true
}

/// Whether the card is holding a press — a control under a hand, or the field being dragged — which
/// is what says the pointer is to be kept until the release comes back to this window rather than
/// to whatever the hand has slid onto (see `window_proc`).
pub(super) fn a_control_is_held() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.find.pressed.is_some() || state.find.dragging)
        })
        .unwrap_or(false)
}

/// A release of a press the card took: the control it lit is let go of, a drag in the field stops
/// where it got to, and a chevron held down stops stepping. Answers whether the card was holding a
/// press, which is what makes the release the card's rather than the page's (see `window_proc`).
pub(super) fn release() -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
    else {
        return false;
    };

    let held = find.pressed.take().is_some();
    let dragging = std::mem::replace(&mut find.dragging, false);
    let stepping = find.hold.take().is_some();
    held || dragging || stepping
}

/// Take a press on one of the card's controls: the part is drawn under the hand until the release
/// comes back, and the two chevrons arm the repeat a hand expects of a button it keeps down. `Aa`, `W`
/// and the cross are one-shot and arm nothing.
///
/// The press has already written its own step down (see `press`), so what is armed here is only the
/// next one: the hold is the tick's question, and the wake is so that the loop is not asleep on a wait
/// longer than the delay (see `wake_preview_loop` and `step_the_hold`).
fn take_the_press(part: Option<pin_chrome::FindBarPart>) {
    let mut repeating = false;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
            .filter(|find| find.open)
        {
            find.pressed = part;
            find.hold = match part {
                Some(
                    held @ (pin_chrome::FindBarPart::Previous | pin_chrome::FindBarPart::Next),
                ) => {
                    repeating = true;
                    Some(Hold {
                        delta: if held == pin_chrome::FindBarPart::Previous {
                            -1
                        } else {
                            1
                        },
                        due: Instant::now() + Duration::from_millis(HOLD_DELAY_MS),
                    })
                }
                _ => None,
            };
        }
    }

    if repeating {
        wake_preview_loop();
    }
}

/// Step a chevron that is still held: nothing before the delay, and once per interval after it.
///
/// The tick asks this before it settles what the bar was told, so a repeat is the same `FindEdit::Step`
/// a key or the press wrote down, and it takes the same road to the page. `now` is passed rather than
/// read so that the tests can hold a chevron on a clock of their own; the only caller is the tick, and
/// the loop's own wait is the granularity of the repeat rather than a rate of its own — a step that
/// comes due is a step the next wake makes, so a pin of a static page is left waking at
/// `STATIC_PIN_WAIT_MS` (see `settle`).
pub(super) fn step_the_hold(now: Instant) {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
    else {
        return;
    };
    let Some(hold) = find.hold.as_ref() else {
        return;
    };

    let (delta, mut due) = (hold.delta, hold.due);
    let mut steps = 0;
    while due <= now && steps < HOLD_CATCH_UP {
        find.pending.push(FindEdit::Step(delta));
        due += Duration::from_millis(HOLD_REPEAT_MS);
        steps += 1;
    }
    // A tick that comes round after a stall catches up on the steps it owes and no more: the cadence
    // starts again at the next interval rather than bursting through the ones that were missed.
    if due <= now {
        due = now + Duration::from_millis(HOLD_REPEAT_MS);
    }

    if let Some(hold) = find.hold.as_mut() {
        hold.due = due;
    }
}

/// Open the bar over the pinned page on screen, or ask the same search of a bar that is already
/// open: `Ctrl`+`F` over a bar that holds a query refocuses it rather than throwing the query away,
/// because the reader who reaches for the key again is reaching for the bar they were just in — and
/// it selects the whole query, the way the browser they are used to does, so that the next character
/// typed replaces it rather than appending to it (see `text_find_field::Field::select_all`).
///
/// The caret has nowhere to move to — the bar's query is its only field, so being up is being
/// focused — and the query is searched again from the line the reader is on when the bar was away,
/// because a bar closed and opened again comes back at the query it was left with rather than at its
/// matches (see `FindState`). A bar that was already up has nothing to search again: its query has
/// not changed, and the reader asked for the field rather than for the document — unless they have
/// picked something else out of the page since, which is the seed below.
///
/// **A page with something selected seeds the query with it**, which is what an editor's own find box
/// does: a word double pressed in the page, or a run of it dragged over, is the thing the reader is
/// asking to find, and reaching for the key is the whole of what they asked. It replaces the query
/// the bar was left holding rather than joining it, and the caret is left after it — the reader who
/// wants the rest of a word typed appends to it, and the one who wants a different query presses
/// `Ctrl`+`F` again over a page with nothing selected, or `Ctrl`+`A` in the field.
///
/// **That holds over a bar that is already up as well.** A reader who picks another word out of the
/// page and reaches for the key is asking to find that one, not the query the bar is showing, so the
/// seed replaces it and the document is searched again — the same box a second press of a shortcut in
/// an editor is (see `the_pages_selection`, which is what tells that reader from the one who is still
/// in the query).
pub(super) fn open_text_find() {
    // Asked before the bar is written to, and outside the lock the bar is opened under, because it
    // reads the document rather than the bar's own state (see `the_pages_selection`).
    let seed = the_pages_selection();

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
        else {
            return;
        };

        find.focused = FindFocus::Field;
        // A query put in the bar is a new search rather than a step through the run's history: what the
        // reader stepped to is behind them (see `step_the_history`).
        find.history_at = None;
        find.history_draft = None;

        if find.open {
            // A bar that is up takes a seed and searches it; with nothing seeded — the reader who is
            // still in the query rather than picking words out of the page — it selects the query it
            // holds, so that the next character typed replaces it.
            if seed.is_some_and(|text| find.field.seed(&text)) {
                find.pending.push(FindEdit::SearchAgain);
            } else {
                find.field.select_all();
            }
        } else {
            find.open = true;
            // A bar opened over a selection comes back at that and not at the query it was left
            // holding; opened over nothing it comes back at the query it was left with, selected so
            // that a typed character replaces it (see `Field::seed` and `Field::select_all`).
            if !seed.is_some_and(|text| find.field.seed(&text)) {
                find.field.select_all();
            }
            find.pending.push(FindEdit::SearchAgain);
        }
        find.blink.restart(Instant::now(), caret_blink());
    }
}

/// What the reader selected in the page, or nothing while the page holds no selection or the reader
/// is still in the query: the seed `Ctrl`+`F` puts in the query (see `open_text_find`).
///
/// **A bar that is up and focused says no.** A selection left in the page under a bar the reader is
/// typing in is where they were before, not what they are asking to find: the chord pressed from
/// inside the query asks for the field, and a query replaced by a word the reader stopped looking at
/// would be a search they never asked for. A page they have just pressed into is the other way round —
/// `FindFocus` follows the hand — so that is the reader picking out a new word to find.
///
/// It is read of the page rather than of the frame that shows it, so a selection reaching past the
/// box is seeded whole, and it is read before the bar's own lock is taken because it reads the
/// document.
fn the_pages_selection() -> Option<String> {
    let (path, options, selection) = CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let state = media.text_state.as_ref()?;
        if state.find.open && state.find.focused != FindFocus::Page {
            return None;
        }

        state
            .selection
            .filter(|selection| !selection.is_empty())
            .map(|selection| (state.path.clone(), state.options, selection))
    })?;

    text_preview::text_of(&path, options, selection)
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
///
/// **The press bit alone is not enough to catch it.** That bit is spent by whoever reads the key
/// first, and the Explorer hook's own pass reads every letter's on the ticks where neither `Ctrl` nor
/// `Alt` is down (see `explorer_hook::input::type_ahead_keys`) — so a chord whose `Ctrl` is let go
/// before this pass reads it is a chord this poll never sees. The key being *down* where it was not
/// on the last pass is the other half, and a hand that holds the chord for either a tap or a moment is
/// caught by one of the two (see `CHORD_HELD`).
pub(super) fn pinned_find_requested() -> bool {
    if !pin_is_focused() || !text_selection::has_text_document() {
        return false;
    }

    unsafe {
        let control = GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0;
        if !control {
            CHORD_HELD.store(false, Ordering::Relaxed);
            return false;
        }

        // One read of the key, for both of the things it says: whether it has been pressed since the
        // last read of it, and whether it is down now. **Read only under `Ctrl`**, because a read is
        // what spends that press bit: a pass that found the modifier already let go would otherwise
        // spend a press it had no chord for, and the hand that asked for the bar would have to ask a
        // second time.
        let f = GetAsyncKeyState(VK_F.0 as i32) as u16;
        let held = f & 0x8000 != 0;
        let was_held = CHORD_HELD.swap(held, Ordering::Relaxed);

        f & 0x0001 != 0 || (held && !was_held)
    }
}

/// Whether the last pass that could answer the chord found `Ctrl`+`F` held down, which is the half of
/// the chord that no other reader of the key can spend (see `pinned_find_requested`).
static CHORD_HELD: AtomicBool = AtomicBool::new(false);

/// The character a `WM_CHAR` is, where the bar is to type it: `wparam` is one UTF-16 code unit, so a
/// surrogate is half of a character this bar does not hold — BMP only — and it is dropped rather
/// than typed as half of one. A control character is not a printable one, which is what keeps a
/// backspace, a return and a `Ctrl` chord from arriving twice: they are answered as keys, and
/// `TranslateMessage` has already turned each into a character by the time the bar reads them.
pub(super) fn typed_character(wparam: usize) -> Option<char> {
    char::from_u32(wparam as u32).filter(|character| !character.is_control())
}

/// The modifiers a key was pressed under, named by the caller rather than read here so that a test can
/// ask about `Ctrl`+`←` without holding `Ctrl` down.
///
/// **A `WM_KEYDOWN` carries none of them, and the message road is where they are asked.** A
/// `WH_KEYBOARD_LL` callback may not ask the keyboard — it runs in the OS input path, where every
/// keystroke on the machine waits for it to return — so the press is read where it is delivered, at
/// the window procedure's message, and the edit it means is queued for the tick from there (see
/// `key_edit_with`). That is the same reading `Ctrl`+`A` and a `Shift`+click have always been given
/// (see `text_selection::shift_is_down`).
#[derive(Clone, Copy, Default, Debug)]
pub(super) struct Modifiers {
    pub(super) shift: bool,
    pub(super) control: bool,
}

impl Modifiers {
    /// The two keys as they stand now, asked of the same two polls `text_selection` has always made
    /// for a `Ctrl` chord and a `Shift`+click.
    pub(super) fn polled() -> Self {
        Self {
            shift: text_selection::shift_is_down(),
            control: text_selection::control_is_down(),
        }
    }
}

/// What one key pressed on the pin's own window asks of a bar that is up, or nothing for a key the
/// bar only swallows (see `window_proc`). It is the window's entry to the table below, with the
/// modifiers read off the keyboard of now, which is the only place that can answer them: a message
/// carries none (see `Modifiers`).
///
/// **A key that would type a character is not here.** A letter reaches the field as `WM_CHAR`, where
/// `TranslateMessage` has already turned it into one, and the same is true of every control the
/// keyboard delivers as a character: a `Ctrl` chord arrives both as a key and as a control character,
/// and it is answered here so that `typed_character` can drop the character (see `typed_character`).
///
/// **`Ctrl`+`U` is deliberately not a key here.** It clears a line in a browser's view source and
/// underlines in Office, so a reader who presses it has asked for nothing this field owns; emptying
/// the query is `Esc`'s (see `FindEdit::Close`).
pub(super) fn key_edit(vk: i32) -> Option<FindEdit> {
    key_edit_with(vk, Modifiers::polled())
}

/// What one key means to a bar that is up: `←`/`→` move the caret (`Ctrl` by a word, `Shift`
/// extending), `Home`/`End` take the ends of the query (`Shift` extending), `↑` and `↓` step back and
/// forward through the queries this run has searched, `Enter`/`Shift`+`Enter` step the matches on and
/// back, `F3`/`Shift`+`F3` step them the same way, `Esc` clears the query or closes the bar,
/// `Backspace`/`Delete` delete (`Ctrl` by a word), and `Tab` is inert.
///
/// **This is the only table that reads a key, and a message is its only road.** The keyboard hook
/// stands down for the whole of a bar's life (see `key_input::publish_find_bar_up`), so every one of
/// these arrives at the pin's window and is read here; the table keyed on the hook's own commands
/// that used to shadow it is what let a `Space` and a `T` rot unnoticed — the two keys a query
/// cannot be written without, and the two the hook ate before a character could exist.
///
/// **The order is an editor's find widget's, not a browser's.** `↑` and `↓` are the run's search
/// history rather than the matches, `Enter` and `Shift`+`Enter` are the two steps through the matches
/// and `F3` repeats `Enter`; the page keys are nothing here at all, because a find widget's up and
/// down belong to its own field and a hard page is a scroll the reader has not asked for. A key the
/// field does not name is swallowed before it can reach the page's own table (see `window_proc`).
pub(super) fn key_edit_with(vk: i32, held: Modifiers) -> Option<FindEdit> {
    let Modifiers { shift, control } = held;

    if vk == VK_ESCAPE.0 as i32 {
        return Some(FindEdit::ClearOrClose);
    }

    if vk == VK_RETURN.0 as i32 {
        return Some(FindEdit::Step(if shift { -1 } else { 1 }));
    }

    // `F3` repeats `Enter`, and `Shift` makes it the step back — the pair a reader's hand knows from
    // every find a document window has ever had. The two arrows are the field's own history, so they
    // are answered whether or not `Ctrl` is down: a hand resting on the modifier is not a hand asking
    // for something else (see `step_the_history`).
    if vk == VK_F3.0 as i32 {
        return Some(FindEdit::Step(if shift { -1 } else { 1 }));
    }

    if vk == VK_UP.0 as i32 {
        return Some(FindEdit::History(-1));
    }

    if vk == VK_DOWN.0 as i32 {
        return Some(FindEdit::History(1));
    }

    if control {
        return match vk {
            v if v == VK_A.0 as i32 => Some(FindEdit::SelectAll),
            v if v == VK_C.0 as i32 => Some(FindEdit::Copy),
            v if v == VK_X.0 as i32 => Some(FindEdit::Cut),
            v if v == VK_V.0 as i32 => Some(FindEdit::Paste),
            // The two undo chords: `Ctrl`+`Z` puts the last edit back, `Ctrl`+`Y` and
            // `Ctrl`+`Shift`+`Z` put back what an undo took away — the pair Windows' own text
            // fields answer.
            v if v == VK_Z.0 as i32 => Some(if shift {
                FindEdit::Redo
            } else {
                FindEdit::Undo
            }),
            v if v == VK_Y.0 as i32 => Some(FindEdit::Redo),
            v if v == VK_LEFT.0 as i32 => {
                Some(FindEdit::Move(text_find_field::Motion::WordLeft, shift))
            }
            v if v == VK_RIGHT.0 as i32 => {
                Some(FindEdit::Move(text_find_field::Motion::WordRight, shift))
            }
            v if v == VK_HOME.0 as i32 => {
                Some(FindEdit::Move(text_find_field::Motion::Home, shift))
            }
            v if v == VK_END.0 as i32 => Some(FindEdit::Move(text_find_field::Motion::End, shift)),
            v if v == VK_BACK.0 as i32 => Some(FindEdit::DeleteWordLeft),
            v if v == VK_DELETE.0 as i32 => Some(FindEdit::DeleteWordRight),
            _ => None,
        };
    }

    match vk {
        v if v == VK_BACK.0 as i32 => Some(FindEdit::Backspace),
        v if v == VK_DELETE.0 as i32 => Some(FindEdit::DeleteForward),
        v if v == VK_LEFT.0 as i32 => Some(FindEdit::Move(
            text_find_field::Motion::CharacterLeft,
            shift,
        )),
        v if v == VK_RIGHT.0 as i32 => Some(FindEdit::Move(
            text_find_field::Motion::CharacterRight,
            shift,
        )),
        v if v == VK_HOME.0 as i32 => Some(FindEdit::Move(text_find_field::Motion::Home, shift)),
        v if v == VK_END.0 as i32 => Some(FindEdit::Move(text_find_field::Motion::End, shift)),
        _ => None,
    }
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
    /// Text to put on the clipboard once the page's lock is back: a copy's selection, or what a cut
    /// took out. The clipboard is opened on the tick's window rather than under this lock, because
    /// the writer asks this module's own lock before it writes (see `text_selection::copy_text_preview`).
    copy: Option<String>,
    /// Whether the page's own selection is to go on the clipboard: a `Ctrl`+`C` over a bar whose query
    /// has nothing selected is the reader copying what they selected in the page, which is the page's
    /// copy and not the field's (see the `FindEdit::Copy` arm of `settle`).
    copy_page: bool,
    /// Whether the clipboard is to be read and pasted into the field: the read happens on the tick
    /// too, for the same reason, and the field is written once its text is in hand (see `settle`).
    paste: bool,
}

/// What the card over the page draws of the bar, as one value: the query, where the caret and the
/// other end of its selection stand, the two switches, and how many places were found and which of
/// them the reader is on.
///
/// **A batch of edits that leaves every one of these as it was is a batch with nothing to draw**, and
/// that is what keeps a key that moved nothing from repainting: a band repaint copies the whole page
/// frame under the card (see `settle`). The caret's blink is not in here, because it is the tick's
/// own question rather than an edit's (see `the_caret_turned`).
#[derive(PartialEq, Eq)]
struct Drawn {
    open: bool,
    query: String,
    caret: usize,
    anchor: usize,
    case_sensitive: bool,
    whole_word: bool,
    count: usize,
    current: Option<usize>,
}

/// The bar as the card draws it now, or nothing while the pin holds no page at all (see `Drawn`).
fn as_drawn() -> Option<Drawn> {
    let media = CURRENT_MEDIA.lock().ok()?;
    let find = &media.as_ref()?.text_state.as_ref()?.find;

    Some(Drawn {
        open: find.open,
        query: find.field.text().to_string(),
        caret: find.field.caret(),
        anchor: find.field.anchor(),
        case_sensitive: find.case_sensitive,
        whole_word: find.whole_word,
        count: find.matches.len(),
        current: find.current,
    })
}

/// How many settled batches have asked for the band to be drawn again.
///
/// It is kept for the one thing about a repaint a window cannot tell a test: that a key which moved
/// nothing drew nothing. A count rather than a flag because a test asks what one batch added, and a
/// batch that repaints is a batch that asked at all — whichever of the two painters answered it (see
/// `settle`).
pub(super) static BAR_PAINTS: AtomicU64 = AtomicU64::new(0);

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
    // A chevron held down steps on this tick's clock before what it steps is settled: the wait is the
    // granularity of the repeat, so a step that has come due is a step the wake that brought this tick
    // makes, and the loop keeps the cadence it has (see `step_the_hold`).
    step_the_hold(Instant::now());

    // What the card draws before the batch is applied, so that the tail of this can tell a batch that
    // drew something from one that only queued an edit (see `Drawn`).
    let before = as_drawn();
    let Some(mut settled) = take_settled() else {
        // Nothing was said to the bar, so the only thing of its own that can have changed is the
        // caret's blink: a bar that is up and holding text asks for a repaint twice a blink, and one
        // that is down asks for nothing at all (see `CaretBlink`).
        if the_caret_turned() {
            render_layered_preview(hwnd);
        }
        return;
    };

    // A copy is the clipboard's whether or not the bar survives the same batch of keys: a cut
    // followed by `Esc` puts the query on the clipboard before the bar it was cut in goes away.
    if let Some(text) = settled.copy.take() {
        text_selection::write_clipboard_text(hwnd, &text);
    }
    if settled.copy_page {
        // The page's own copy, asked of the page: it takes what is selected in the document and puts
        // that on the clipboard, which is the same road `Ctrl`+`C` takes with no bar up.
        copy_text_preview(hwnd);
    }

    if settled.closed {
        // The page does not move: a close is the highlights going, and what was found is kept by
        // nothing once the bar is not up (see `FindState`).
        render_layered_preview(hwnd);
        return;
    }

    if !settled.open {
        return;
    }

    // The clipboard is read and written here rather than under the page's lock, and before the
    // search below, because a paste is a change to the query the search is about to read.
    if settled.paste {
        if let Some(pasted) = text_selection::clipboard_text(hwnd) {
            if let Ok(mut media) = CURRENT_MEDIA.lock() {
                if let Some(find) = media
                    .as_mut()
                    .and_then(|media| media.text_state.as_mut())
                    .map(|state| &mut state.find)
                    .filter(|find| find.open)
                {
                    let before = find.field.text().to_string();
                    find.field.paste(&pasted);
                    if find.field.text() != before {
                        // A pasted query is the reader's own, not a step through the history they had
                        // walked to (see `step_the_history`).
                        find.history_at = None;
                        find.history_draft = None;
                        find.blink.restart(Instant::now(), caret_blink());
                        settled.query = find.field.text().to_string();
                        settled.search = true;
                    }
                }
            }
        }
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

    // **A batch that changed nothing the card draws owes no paint.** A `←` at the start of a query, a
    // `Shift`+`←` with the selection already at it, a `Backspace` with nothing before the caret: each
    // is an edit the road queues like any other, and a band repaint copies the whole page frame under
    // the card. Asked after the search and the step rather than before, so a step that landed on the
    // match it started from is the same answer — and asked of the caret's blink never, because that is
    // the tick's own question where no edit is in hand (see `the_caret_turned`).
    //
    // **A search is exempt, and that is what makes a bar that has just come up visible.** A bar opened
    // over a query its page does not hold finds nothing, and the answer it draws of that — `0/0` and no
    // place to be — is the answer the bar already stood in when the search began, because opening it
    // is what put the query there: compared on what is drawn, that batch would paint nothing, and the
    // bar would be up on the state's side and nowhere in the window's pixels until something else
    // happened to repaint them. So a batch that searched paints.
    if !settled.search && as_drawn() == before {
        return;
    }

    BAR_PAINTS.fetch_add(1, Ordering::Relaxed);

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

/// Whether the caret's blink turned over on this tick. Asked only when the bar said nothing, so that
/// a tick with an edit in hand answers the edit's solid caret rather than a blink that arrives in the
/// same turn as it.
fn the_caret_turned() -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };
    let Some(find) = media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| &mut state.find)
        .filter(|find| find.open)
    else {
        return false;
    };

    find.blink.advance(Instant::now(), caret_blink())
}

/// What a `Ctrl`+`C` under the bar copies: one surface's selection or the other's.
#[derive(Clone, PartialEq, Eq, Debug)]
pub(super) enum CopyTarget {
    /// The query's own selection, which is the text as it stands in the field.
    Field(String),
    /// The page's selection, copied the way `Ctrl`+`C` copies it with no bar up (see
    /// `copy_text_preview`).
    Page,
}

/// The surface a `Ctrl`+`C` under the bar copies from: the one that was focused last, if it has
/// something selected.
///
/// **A bar being up is not a claim on the clipboard, and neither is a caret in it.** The field is
/// focused by opening the bar and by a press in the query; the page is focused by a press or a drag
/// in it, which is a selection made in the page whether the bar is up or not. So the reader who
/// selected in the page and then clicked into the query to search it copies the query, and the one
/// who typed a query and then clicked into the page to pick the match by hand copies the page — even
/// though both surfaces have a selection at that moment.
///
/// A surface with nothing selected is not a claim either: the copy falls back to the other one's
/// selection, which is what the chord has always done with the query's selection empty.
pub(super) fn copy_target(
    focused: FindFocus,
    field_text: String,
    page_has_a_selection: bool,
) -> CopyTarget {
    match focused {
        FindFocus::Field if field_text.is_empty() => CopyTarget::Page,
        FindFocus::Field => CopyTarget::Field(field_text),
        FindFocus::Page if !page_has_a_selection => CopyTarget::Field(field_text),
        FindFocus::Page => CopyTarget::Page,
    }
}

/// What the bar's own road decided a batch of keys gives the clipboard, and without touching it: the
/// text `settle` would write and whether it would be the page's own copy instead.
///
/// A test of which surface a `Ctrl`+`C` copies may not write the machine's clipboard, and the
/// decision is the whole of what is being asked (see `copy_target` and the `FindEdit::Copy` arm of
/// `take_settled`).
#[cfg(test)]
pub(super) fn the_copy_a_batch_decided_for_a_test() -> Option<(Option<String>, bool)> {
    let settled = take_settled()?;
    Some((settled.copy, settled.copy_page))
}

/// Take the bar's edits out of the page and apply the ones that need no document: the field's own
/// edits, its selection and its undo stack, the two switches, and a close — while a paste and a copy
/// are held for the tick, which is the one that may take the clipboard (see `settle`).
fn take_settled() -> Option<Settled> {
    let mut media = CURRENT_MEDIA.lock().ok()?;
    let state = media.as_mut()?.text_state.as_mut()?;
    // Read before the bar's own state is borrowed: which of the two selections a copy is the
    // reader's is a question about both (see `copy_target`).
    let page_has_a_selection = state
        .selection
        .is_some_and(|selection| !selection.is_empty());
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
        copy: None,
        copy_page: false,
        paste: false,
    };
    let before = find.field.text().to_string();
    let ends_before = (find.field.caret(), find.field.anchor());
    // The query as a step through the run's history left it, or nothing while this batch stepped
    // through none: what tells a field the reader typed over from one the history put there (see
    // `step_the_history`).
    let mut after_the_history: Option<String> = None;

    for edit in std::mem::take(&mut find.pending) {
        match edit {
            FindEdit::Typed(character) => find.field.insert(character),
            FindEdit::Backspace => find.field.backspace(),
            FindEdit::DeleteForward => find.field.delete(),
            FindEdit::DeleteWordLeft => find.field.delete_word_left(),
            FindEdit::DeleteWordRight => find.field.delete_word_right(),
            FindEdit::Move(motion, extend) => {
                find.field.move_caret(motion, extend);
            }
            FindEdit::SelectAll => find.field.select_all(),
            FindEdit::Copy => {
                // The surface the reader was last in takes the chord, and one with nothing selected
                // gives it back to the other — a bar being up is not a claim on the clipboard (see
                // `copy_target`).
                match copy_target(
                    find.focused,
                    find.field.selected_text(),
                    page_has_a_selection,
                ) {
                    CopyTarget::Field(text) => {
                        if !text.is_empty() {
                            settled.copy = Some(text);
                        }
                    }
                    CopyTarget::Page => settled.copy_page = true,
                }
            }
            FindEdit::Cut => {
                let text = find.field.cut();
                if !text.is_empty() {
                    settled.copy = Some(text);
                }
            }
            FindEdit::Paste => settled.paste = true,
            FindEdit::Undo => {
                find.field.undo();
            }
            FindEdit::Redo => {
                find.field.redo();
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
            FindEdit::Step(delta) => {
                settled.step += delta;
                // A query stepped through is a query the reader is done asking, and that is what the
                // run's history is written from: `Enter`, `F3` and the two chevrons commit it (see
                // `remember_the_query`).
                remember_the_query(find.field.text());
            }
            FindEdit::History(delta) => {
                step_the_history(find, delta);
                after_the_history = Some(find.field.text().to_string());
            }
            FindEdit::ClearOrClose => {
                // `Esc` is the field's clear first and the bar's close second: a query with text is
                // emptied and left standing, and a query already empty is the bar going away. One
                // key does both in turn rather than two keys doing one each, which is what makes it
                // the clearing key a reader reaches for without closing what they were in.
                if !find.field.text().is_empty() {
                    // A query emptied is a query finished with, so it goes into the history the
                    // reader can step back to — and the history it was stepped from is left behind
                    // (see `remember_the_query`).
                    remember_the_query(find.field.text());
                    find.field.clear();
                    find.history_at = None;
                    find.history_draft = None;
                    continue;
                }

                find.open = false;
                find.matches.clear();
                find.current = None;
                find.pressed = None;
                find.hold = None;
                find.dragging = false;
                find.click_run = None;
                settled.closed = true;
                settled.open = false;
                break;
            }
            FindEdit::Close => {
                // A bar put away is a search finished with: what it was asking goes into the run's
                // history, where `Ctrl`+`F` and `↑` can bring it back (see `remember_the_query`).
                remember_the_query(find.field.text());
                find.open = false;
                find.matches.clear();
                find.current = None;
                // A bar that is away has no button under a hand, no chevron still stepping and no
                // drag still growing a selection: a close from a key leaves none of them, and the
                // page under a pointer that is still down is the page's from here (see
                // `take_the_press` and `release`).
                find.pressed = None;
                find.hold = None;
                find.dragging = false;
                find.click_run = None;
                settled.closed = true;
                settled.open = false;
                // Nothing after a close is read: a key pressed after the bar went away is a key for
                // the page, and the next tick's window procedure answers it there.
                break;
            }
        }
    }

    // Every edit that changed the field leaves the caret solid for a blink, which is what a reader
    // typing looks at (see `CaretBlink::restart`). An edit that changed nothing leaves the blink where
    // it was: a restart is drawing a solid caret, and a batch that draws nothing must not leave the
    // field's blink lit while the screen still shows the phase it was in (see `settle`).
    if (find.field.caret(), find.field.anchor()) != ends_before || find.field.text() != before {
        find.blink.restart(Instant::now(), caret_blink());
    }

    // A query the reader is typing is the draft that `↓` comes back to, so the history they had stepped
    // to is behind them the moment the field changes by hand: what they type is their own search, not a
    // correction of the one they stepped to (see `step_the_history`). Compared against the query the
    // history left rather than against the one the batch began with, so a character typed *after* a step
    // is a change of the reader's own and a step after a typed one is not.
    if find.history_at.is_some() && after_the_history.as_deref() != Some(find.field.text()) {
        find.history_at = None;
        find.history_draft = None;
    }

    settled.query = find.field.text().to_string();
    // A change to the query is the search running again, however it was made: a typed character, a
    // deleted selection, an undo, a paste the tick applies next to this. A caret move and a select
    // all change nothing the document is searched for.
    if settled.query != before {
        settled.search = true;
    }
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
