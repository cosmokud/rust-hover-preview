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
        card: pin_chrome::find_bar_layout(state.width as i32, state.height as i32, page_top, dpi),
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
/// field is drawn from the right, so a press turned into a place from the field's left edge would
/// land a field's width away from the character the hand is on (see `pin_chrome::query_index_at`).
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

    let at = pin_chrome::query_index_at(&placed.card, find.field.text(), x, placed.dpi);
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

    let at = pin_chrome::query_index_at(&placed.card, find.field.text(), x, placed.dpi);
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
/// not changed, and the reader asked for the field rather than for the document.
pub(super) fn open_text_find() {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
        else {
            return;
        };

        if find.open {
            find.field.select_all();
        } else {
            find.open = true;
            find.field.select_all();
            find.pending.push(FindEdit::SearchAgain);
        }
        find.blink.restart(Instant::now(), caret_blink());
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
/// extending), `Home`/`End` take the ends of the query (`Shift` extending), `↑`/`PgUp` and
/// `↓`/`PgDn` step the matches, `Enter`/`Shift`+`Enter` step on and back, `F3` steps on, `Esc`
/// clears the query or closes the bar, `Backspace`/`Delete` delete (`Ctrl` by a word), and `Tab` is
/// inert.
///
/// **This is the only table that reads a key, and a message is its only road.** The keyboard hook
/// stands down for the whole of a bar's life (see `key_input::publish_find_bar_up`), so every one of
/// these arrives at the pin's window and is read here; the table keyed on the hook's own commands
/// that used to shadow it is what let a `Space` and a `T` rot unnoticed — the two keys a query
/// cannot be written without, and the two the hook ate before a character could exist.
///
/// The order is the one every find bar a reader has used before this one has: `Enter` and the down
/// arrow step on, `Shift`+`Enter` and the up arrow step back, and both wrap at the ends. The page
/// keys step the matches rather than turning the page, which is what a find bar's up and down are
/// for; a key the field does not name is swallowed before it can reach the page's own table (see
/// `window_proc`).
///
/// `key_edit_with` is named for its modifiers so a test can ask about `Ctrl`+`←` without holding
/// `Ctrl` down; `key_edit` is the window's call, and it polls them (see `Modifiers`).
pub(super) fn key_edit_with(vk: i32, held: Modifiers) -> Option<FindEdit> {
    let Modifiers { shift, control } = held;

    if vk == VK_ESCAPE.0 as i32 {
        return Some(FindEdit::ClearOrClose);
    }

    if vk == VK_RETURN.0 as i32 {
        return Some(FindEdit::Step(if shift { -1 } else { 1 }));
    }

    // The two arrows, the two page keys and `F3` step the matches whether or not `Ctrl` is down: they
    // are the find's own keys, and the bar must not stop stepping the moment a hand rests on `Ctrl`.
    if vk == VK_DOWN.0 as i32 || vk == VK_NEXT.0 as i32 || vk == VK_F3.0 as i32 {
        return Some(FindEdit::Step(1));
    }

    if vk == VK_UP.0 as i32 || vk == VK_PRIOR.0 as i32 {
        return Some(FindEdit::Step(-1));
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
    // the card. Asked after the search and the step rather than before, so a re-search that found the
    // same places and a step that landed on the match it started from are the same answer — and asked
    // of the caret's blink never, because that is the tick's own question where no edit is in hand
    // (see `the_caret_turned`).
    if as_drawn() == before {
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

/// Take the bar's edits out of the page and apply the ones that need no document: the field's own
/// edits, its selection and its undo stack, the two switches, and a close — while a paste and a copy
/// are held for the tick, which is the one that may take the clipboard (see `settle`).
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
        copy: None,
        paste: false,
    };
    let before = find.field.text().to_string();
    let ends_before = (find.field.caret(), find.field.anchor());

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
                let text = find.field.selected_text();
                if !text.is_empty() {
                    settled.copy = Some(text);
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
            FindEdit::Step(delta) => settled.step += delta,
            FindEdit::ClearOrClose => {
                // `Esc` is the field's clear first and the bar's close second: a query with text is
                // emptied and left standing, and a query already empty is the bar going away. One
                // key does both in turn rather than two keys doing one each, which is what makes it
                // the clearing key a reader reaches for without closing what they were in.
                if !find.field.text().is_empty() {
                    find.field.clear();
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
