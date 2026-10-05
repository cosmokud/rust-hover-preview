//! What the hook watches while a preview is pinned, so that the tray's
//! `Pin Mode … Update Preview` can show the pin the file the user picks next.
//!
//! It is state of the same shape as the hover machinery beside it, and is kept apart
//! from it rather than shared: a pin is not a hover. The latch that holds a re-hover
//! back, the gate a folder change raises and the file a preview is "about" all describe
//! a preview that is not on screen, and none of them is read unless a pin is up and the
//! setting asks for one to follow.

use super::*;

/// Whether a pin that is up is shown the file the user picks next, and whether the pointer's own
/// hover is one of the ways it is told about one, as the configuration has them (see the tray's
/// `Pin Mode … Update Preview`).
///
/// They are read here rather than kept in the tick's own snapshot of the configuration, and they
/// are read only while a pin is up: the two switches are the hook's answer to a question the rest
/// of its state has nothing to do with, and a switch thrown in the tray is honoured on the tick
/// after it is thrown rather than at the next snapshot.
pub(super) fn pin_update_settings() -> (bool, bool) {
    CONFIG
        .lock()
        .map(|config| (config.pin_update_enabled, config.pin_update_on_hover))
        .unwrap_or((DEFAULT_PIN_UPDATE_ENABLED, DEFAULT_PIN_UPDATE_ON_HOVER))
}

/// What the hook watches while a preview is pinned, so that the tray's `Pin Mode … Update Preview`
/// can show the pin the file the user picks next.
///
/// It is state of the shape the hover machinery beside it keeps, and it is kept apart from it rather
/// than shared: a pin is not a hover, so the latch that holds a re-hover back, the gate a folder
/// change raises and the file a preview is "about" all describe a preview that is not on screen.
/// Nothing here is read unless a pin is up and the setting asks for one to follow, and nothing of
/// the hover machinery is written by it — what a pin does with an answer is the preview loop's
/// business, and the file it is showing is read back from there (see `pinned_path`).
///
/// One argument runs through all of it, and the three places it is asked about carry it in three
/// shapes (`place`, `pending.place`, `sel_place`): a *listing* changing is not a *file* being
/// picked. A folder, a tab and a window the user moves to each put the focus on a new item without
/// a key having walked to it, and the item a fresh listing puts under a hand nobody moved is drawn
/// exactly where the click landed — so a place is what tells a move from a landing, and a witness
/// (a key, or a click standing as one) is what tells a landing from a pick. Where the shell
/// describes no place, a look that answered nothing cannot answer either way, and each baseline is
/// left as that read left it: the listing's own selection keeps its last answer, since a miss
/// rather than a file means nothing has been seen to replace it, while the keyboard's item's place
/// is dropped rather than kept, because the item has moved and the only place in hand is the one it
/// moved out of (see `PinUpdateWatch::note_place`).
#[derive(Default)]
pub(super) struct PinUpdateWatch {
    /// The file the pin was last seen showing: a take-up is a watch beginning, a swap is not
    /// (see `PinUpdateWatch::note_shown`).
    pub(super) showing: Option<PathBuf>,
    /// The pointer as the last tick read it, or nothing on the first tick: what has been
    /// measured since is how the hand has come.
    pub(super) pointer: Option<POINT>,
    /// The window under the pointer as the tick before this one read it, or nothing on the
    /// first tick. What a press is refused by when that window is gone or is a popup covering
    /// the listing — a menu the press just dismissed, read on the tick it was up or on the tick
    /// after it closed (see `press_is_a_listing`).
    pub(super) window: Option<HWND>,
    /// When the pointer last moved, which is what a hover of what is under it is measured from.
    pub(super) settled_at: Option<Instant>,
    /// Whether the file under a settled pointer has already been resolved for that settle.
    pub(super) probed: bool,
    /// Whether the hand has moved at all since this watch began. A pointer parked on another
    /// file while the pin goes up has not hovered anything.
    pub(super) arrived: bool,
    /// The item the keyboard was last seen on. What is on it when a watch begins is a
    /// baseline and not a choice.
    pub(super) focused: Option<FocusedItemKey>,
    /// The place the keyboard's item was read in: the baseline `note_place` measures a
    /// landing against.
    pub(super) place: Option<HoverLocation>,
    /// When a key that walks a listing was last seen, or nothing where something that moves the
    /// focus by other means has been seen since. A time rather than a flag, so a witness
    /// nothing made good on cannot outlive the press that gave it — and a key walking the
    /// selection and nothing else that sets it, so an arrow pressed before an Enter does not
    /// stand as the witness for the item that Enter lands the focus on.
    pub(super) walk_at: Option<Instant>,
    /// A click the tick it arrived on could not resolve, held for the short while the shell
    /// takes to catch up rather than dropped with the press bit that is the only evidence it
    /// happened. `None` wherever the click resolved, or the pin was given the answer.
    pub(super) pending: Option<PendingClick>,
    /// The listing the selection below was last read in: the baseline `note_selection`
    /// measures a pick against.
    pub(super) sel_place: Option<HoverLocation>,
    /// The file the listing above last had selected, as the view's own selection pattern
    /// reported it, whoever has the foreground.
    pub(super) sel_selected: Option<PathBuf>,
    /// When the selection above was last polled: the poll answers out of live shell objects
    /// on every read, so it runs on its own cadence rather than on every tick.
    pub(super) sel_polled_at: Option<Instant>,
    /// Where the pointer was when the last press was read, and when it was read, or nothing
    /// where this watch has not seen one. It is what tells a press of its own from a later press
    /// inside the same gesture — one gesture is two presses at one spot in one window of time,
    /// not two choices (see `PinUpdateWatch::press_is_a_pick`).
    ///
    /// The spot and its clock are one record rather than two because one of them alone decides
    /// nothing: a spot with no clock beside it is a spot that is never past, so every press
    /// after it at that pixel is read as part of a gesture that ended long ago.
    pub(super) press_point: Option<(POINT, Instant)>,
}

/// A click in hand: the three things that are always said of it together — when it was read,
/// where the pointer was standing, and the place it was made in. All three or none, because a
/// click held without the spot it landed on, or without the listing it was made in, can neither
/// be retried nor told from a listing that has been replaced.
pub(super) struct PendingClick {
    /// When the press was read, which is the hold's own clock and is set by the click that
    /// started it and by nothing else.
    pub(super) at: Instant,
    /// Where the pointer stood when the press was read: the spot the retry looks the file up
    /// under, and what a hand that has since travelled is told from one that has not.
    pub(super) point: POINT,
    /// The place the click was made in, as the view under the pointer described itself then.
    /// `None` where the shell described none, which leaves the click to its own look at the
    /// point.
    pub(super) place: Option<HoverLocation>,
}

impl PinUpdateWatch {
    /// Follow a pin for one tick: whether the file the user has picked while it is up has changed,
    /// and whether the reason is one the settings count.
    ///
    /// Three things are watched, and each of them is a way a file is picked in Explorer: a click,
    /// which is the pointer acting on the view; the item the keyboard is on, which is what a key
    /// the user presses moves; and — where `On Hover` asks for it — the file the pointer
    /// settles on, at the delay and the settling the hover behind the pin would have been given.
    /// A pin is not told about a file the pointer merely crosses, and not about one that was under
    /// a pointer nobody moved.
    ///
    /// The first item a watch sees is a baseline rather than a pick — it is the item the keyboard
    /// was already on when the watch began — and a swap does not begin a watch again, so the item
    /// the pin was shown another file over is still the one the next key is measured against. That
    /// is what leaves the first selection a user makes after pressing the pin to give it the
    /// keyboard and then clicking back into Explorer a pick rather than a baseline, with the one
    /// after it not the first that counts (see `PinUpdateWatch::note_shown`).
    ///
    /// A click is held for a moment rather than answered once: the press bit it comes from is
    /// spent by the read, and the tick it is spent on is the only tick it is on — which is exactly
    /// the tick a click that lands as a pin takes the focus away from Explorer, or as Explorer takes
    /// it back, is asked about, with the shell not yet answering for the listing under it. Such a
    /// click is held and asked again while the hand stays where it left it (see `pending`).
    pub(super) fn follow(
        &mut self,
        resolver: &mut ItemResolver,
        on_hover: bool,
        hover_delay_ms: u64,
        last_focus_probe: &mut Instant,
        focus_move: FocusMoveInput,
    ) {
        let now = Instant::now();

        // What the keyboard and the pointer did is noted before anything can return: one of the two
        // takes the witness away, and a tick that takes it away has to be a tick that keeps it taken
        // (see `walk_at`).
        self.note_focus_move(focus_move, now);
        let Some(showing) = pinned_path() else {
            // Nothing is pinned, so there is nothing to show anybody: the watch starts again from
            // what the next pin is showing (see `showing`).
            *self = Self::default();
            return;
        };

        if self.showing.as_ref() != Some(&showing) {
            // A pin taken up, and a pin shown another file — by this watch, or by a key the preview
            // loop answered in between. What a swap from here is measured against is the file the
            // pin has now, and the tick is not abandoned over it: the press bits this tick was
            // given have already been spent by the read above, so a watch that returned here would
            // drop the very pick they are the answer to — which is a click on the file the pin is
            // being swapped to, arriving on the tick the swap is noticed on.
            self.note_shown(&showing);
        }

        let Some(pointer) = read_pointer() else {
            return;
        };

        // What is under the pointer is asked about only where a listing can be: the pinned window
        // is this app's own and stands over the listing the file was picked in, and neither the
        // desktop nor a player of this app's own is a file anybody picked.
        let over_explorer = is_cursor_over_explorer_full(pointer.window);

        // Whether this app's own window is what the pointer is on. A pinned window stands over
        // the listing it came from, and it stands over it *where the files are* — a click that
        // lands on it is a click on the listing underneath, and a listing is what is being read
        // for it. Left out, every such click answers "there is no listing here" and is asked
        // again against the same answer until its hold runs out, which is the whole of why the
        // first click after a pin was focused did nothing and the next one worked: whether the
        // spot the hand was on happened to be under the window or beside it (see
        // `click_is_over_a_listing`).
        let over_our_own = is_our_own_window(pointer.window);

        // How far the hand has come since the last tick of the watch, by the tolerance the hover
        // machinery measures a move with. The first tick of a watch has nothing to measure against,
        // so it is that move: the pointer is taken as it is and the hand has arrived at nothing.
        let threshold = KeyboardPointerPause::default().move_threshold_px(false, pointer.dpi);
        let first = self.pointer.is_none();
        let moved = match self.pointer {
            Some(last) => {
                (pointer.point.x - last.x).abs() > threshold
                    || (pointer.point.y - last.y).abs() > threshold
            }
            None => true,
        };
        let previous_window = self.window;
        self.window = Some(pointer.window);
        self.pointer = Some(pointer.point);

        if moved {
            if !first {
                self.arrived = true;
            }
            self.settled_at = Some(Instant::now());
            self.probed = false;
        }

        // Whether this tick looks at what the pointer is on, and why. A click is the pointer acting
        // on the view — what it does in Explorer is select, and the file it selected is the pin's to
        // show — and it counts whatever else is on: it is the one reason the setting leaves on when
        // a hover is not wanted. A hover is the file a hand has come to rest on, at the delay the
        // previews behind the pin are given one, and it is only asked for where the setting asks.
        let settled = self
            .settled_at
            .map(|at| at.elapsed() >= Duration::from_millis(hover_delay_ms))
            .unwrap_or(false);
        let hovered = on_hover && self.arrived && !self.probed && settled;
        // Read by the tick that hands the watch its input rather than here, because the press bit is
        // the one thing about a click that can only be read once a tick (see `focus_move_input`).
        //
        // A press the hand has not moved for, and that landed inside the machine's own
        // double-click window of the press before it, is another press of a gesture it has
        // already read, and is not a pick of its own. That is what a double-click is: one
        // gesture at one point in one window of time. Where the first press opened a folder, the
        // second is answered out of the listing that first one put under a pointer nobody moved —
        // a file the user never chose, and the one the new folder happens to have drawn there.
        // Both doors a press comes through are closed by this one answer: the file under the
        // point, and the listing's own selection (see `PinUpdateWatch::press_is_a_pick`).
        //
        // The window is what leaves the click *after* that folder change alone, though: it shares
        // its pixel with the press that opened the folder and shares the listing under it, and
        // only the time tells them apart. Without it the pin refuses every click at that spot for
        // as long as the hand stays on it, and a pin standing over that spot stops following the
        // listing — which is the fault, and the one this rule's own half above is written for.
        //
        // Where the press landed is the other half of what makes it a pick: a press is read off
        // the key rather than off a window, so the window it landed on is the only thing that
        // says a listing is under it (see `press_is_a_listing`). It is asked *after* the
        // double-click rule rather than before, so a press refused for having landed somewhere
        // else still books its point and cannot make the next press its own second half.
        //
        // Both ticks are asked about chrome, and the tick before is kept for it because a press
        // on a menu is read on whichever tick the shell spends the press bit on: while the popup
        // is still up, or on the poll after it has closed and the point reads as the listing it
        // was covering. The hand does not move for either, and the file under it is not one it
        // picked (see `window_is_menu_popup_chrome`).
        let clicked = focus_move.clicked;
        let current_is_chrome = window_is_menu_popup_chrome(pointer.window);
        let previous_was_chrome = previous_window.is_some_and(window_is_menu_popup_chrome);
        let pick = clicked
            && self.press_is_a_pick(pointer.point, threshold, now)
            && press_is_a_listing(
                over_explorer,
                over_our_own,
                window_is_gone(previous_window),
                current_is_chrome,
                previous_was_chrome,
            );

        // Whether the press was read at all, before anything is decided about it, and where
        // the pointer was when it was. This is the one reading that says a click was lost
        // before any of the rest had a chance to answer it, and it is not in any of the arms
        // below because every one of them is reached only where something else already held.
        if clicked {
            note_pin_click!(
                "press read  fg {}  at {},{}  win {}  prev win {}  gone {}  chrome {}/{}",
                is_foreground_explorer() as u8,
                pointer.point.x,
                pointer.point.y,
                window_class_of(pointer.window),
                previous_window.map_or_else(|| "none".to_string(), window_class_of),
                window_is_gone(previous_window) as u8,
                current_is_chrome as u8,
                previous_was_chrome as u8,
            );
        }

        // The drop the loop makes on the focus transition cannot cover a click: the press bit is
        // spent the moment the button goes down, so the tick this drop is made on is a tick *after*
        // the click that needed it, and a stale view set answers that click with the file the pin
        // is already showing. The click carries the drop itself, once, here, for every shape it
        // takes below (see `PinUpdateWatch::answer_click`).
        if clicked {
            resolver.forget_item();
            resolver.forget_window_views();
        }

        // A pick is already a press that landed on a listing — that is what `press_is_a_listing`
        // asked above — so this is the two ways a listing under the pointer is read, and neither
        // of them asks again whether there is one. A hover is not a pick and is asked for
        // separately, and only over Explorer's own window rather than over this app's.
        if pick || (hovered && over_explorer) {
            self.probed = true;

            if let Some(path) = get_file_under_cursor(resolver, &pointer) {
                // What was resolved is answered whatever it turned out to be — a file nothing can
                // be shown for is one of them — so there is nothing here left to ask about again.
                let offered = self.offer(&path, &showing);
                self.answer_click(&path, &showing, pick, now, pointer.point);
                trace_click(
                    &pointer,
                    &showing,
                    over_explorer,
                    pick,
                    Some(&path),
                    offered,
                    self.pending.as_ref().map(|held| now - held.at),
                );
            } else if pick {
                // The lookup answered nothing, and the click is the only evidence it happened: the
                // press bit it came from has been spent by the read that gave this tick its input
                // and is not on any later tick. Explorer is commonly still coming up when a click
                // lands on it — a click that takes the focus out of the pinned window and a click
                // back into Explorer are both read before the shell has caught up with either — so
                // it is held and asked again below rather than spent for nothing.
                self.pending = Some(PendingClick {
                    at: now,
                    point: pointer.point,
                    place: click_place(resolver, &pointer),
                });
                trace_click(
                    &pointer,
                    &showing,
                    over_explorer,
                    true,
                    None,
                    false,
                    Some(Duration::ZERO),
                );
            }
        }

        // The click above, asked again now that the shell may have caught up. It runs only where
        // something is held, and only something a click put there, so a tick that never saw a
        // pick offers nothing here — the held click is a user's pick and never a hover.
        if !pick {
            // The click's own facts, copied out before anything below is allowed to let it go.
            let held = self.pending.as_ref().map(|held| (held.at, held.point));

            if let Some((at, point)) = held {
                // Three ways a held click is let go, and one release for all of them. The place
                // is only asked where the other two have not already answered it, because
                // asking it is a walk out through the shell.
                let expired = now - at > Duration::from_millis(PIN_CLICK_RETRY_MS);
                // The hand has moved on since the click: it is asking about whatever is under
                // it now, and resolving the file it clicked on a moment ago would be a file
                // the user has not picked.
                let moved = (pointer.point.x - point.x).abs() > threshold
                    || (pointer.point.y - point.y).abs() > threshold;
                // The listing the click was made in is not the one under the pointer any more —
                // a folder, a tab or a window moved to — so the click is let go rather than
                // answered: what the point holds now is a file that arrived under a hand nobody
                // has moved, which is not a file anybody picked.
                let relisted = !expired && !moved && self.click_place_changed(resolver, &pointer);

                if expired || moved || relisted {
                    // Nothing has come of it by now, and what a click held for half a second is
                    // not worth a stale file being shown beside it (see `PIN_CLICK_RETRY_MS`).
                    self.pending = None;
                    trace_click(
                        &pointer,
                        &showing,
                        over_explorer,
                        false,
                        None,
                        false,
                        Some(now - at),
                    );
                } else if click_is_over_a_listing(over_explorer, over_our_own) {
                    // The hand is where it clicked and Explorer is answering for it. The lookup is
                    // kept only while it answers nothing: a file it does resolve is offered and
                    // stops the retry, whatever the offer then does with it.
                    //
                    // What is on top of the point is not asked again here, because it is not
                    // what decides the question: the shell is asked what file is at the point,
                    // and it will answer while this app's own window stands over the listing as
                    // readily as while Explorer's does. Gating this on Explorer's own window
                    // made a click that landed on a pinned window unanswerable for as long as
                    // the hold lasted — the retry could not run, so the shell was never asked,
                    // so nothing was ever resolved.
                    //
                    // Both caches go before the ask rather than only the item's answer: a retry is
                    // a question about a listing that has *since* caught up, and what the click was
                    // answered out of is the listing as it was when the click landed. Keeping the
                    // item's own memo would hand the retry back the very answer it exists to
                    // replace — a look that finds an item and cannot name a file for it is kept
                    // against that item exactly as it keeps the answer that an item is a folder,
                    // and the two are told apart nowhere — and keeping the view set would ask the
                    // stale walk the same question again. One walk a tick for as long as the hold
                    // lasts, and only on a click that is not yet answered, is what a retry is.
                    resolver.forget_item();
                    resolver.forget_window_views();

                    match get_file_under_cursor(resolver, &pointer) {
                        Some(path) => {
                            let offered = self.offer(&path, &showing);
                            self.answer_click(&path, &showing, clicked, now, pointer.point);
                            trace_click(
                                &pointer,
                                &showing,
                                over_explorer,
                                false,
                                Some(&path),
                                offered,
                                self.pending.as_ref().map(|held| now - held.at),
                            );
                        }
                        None => trace_click(
                            &pointer,
                            &showing,
                            over_explorer,
                            false,
                            None,
                            false,
                            Some(now - at),
                        ),
                    }
                }
            }
        }

        // The keyboard's own answer: the item the focus is on, which is what a key the user presses
        // moves. It is probed on the terms the hover path probes it — while Explorer has the
        // foreground, and no more often than the hover path asks — and a focus that has moved onto
        // another file is a file the user picked as surely as one a click selected: another folder,
        // another tab and another window move the focus onto an item as well, and those are moves
        // this setting does not follow.
        if is_foreground_explorer()
            && last_focus_probe.elapsed() >= Duration::from_millis(KEYBOARD_FOCUS_PROBE_MS)
        {
            *last_focus_probe = Instant::now();

            if let Some(focused) = get_focused_explorer_item(resolver) {
                let key = FocusedItemKey::new(focused.item.name.clone(), &focused.item.bounds);
                let changed = self.focused.as_ref() != Some(&key);
                let known = self.focused.is_some();
                self.focused = Some(key);

                if changed {
                    // Why the focus has moved, which is what tells a key the user pressed from a
                    // folder, a tab or a window the user moved to. The place is read for the watch's
                    // first item as much as for a move — a baseline is taken in a place too.
                    let place = focused_item_location(resolver, &focused);

                    // And a place is not the whole of it, which is why the keyboard's own keys are
                    // asked as well: the place a focused item is read in is the view the *pointer*
                    // was last seen working in rather than the one the item is drawn in — an item's
                    // provider reports no window for a frame's views to be told apart by — so a tab
                    // switched or a folder opened under the keyboard reads as the place the watch was
                    // already watching. A move both facts call the keyboard's is a file the user
                    // picked; one either of them calls somebody else's is not (see `focus_pick`).
                    //
                    // A click is the other thing that moves the focus onto a file, and it is read
                    // here because it is the same change: what a click selects, the view reports as
                    // the focus it moved — from its own side, where the watch's own look at the point
                    // answers for whatever is standing on top. It carries its own witness, and a
                    // click that is not this one is not a pick (see `click_picked_item`).
                    let by_key = self.focus_pick(place.clone(), known, now);
                    let bounds = focused.item.bounds;
                    let by_click = self.click_picked_item(
                        (bounds.left, bounds.top, bounds.right, bounds.bottom),
                        place.as_ref(),
                        now,
                    );

                    if by_key || by_click {
                        if let Some(path) = resolve_focused_item_to_path(resolver, &focused) {
                            let offered = self.offer(&path, &showing);

                            // Which of the two readings answered is the one thing a click lost to the
                            // wrong reading cannot say for itself, so it is written down where a
                            // trace is being written.
                            note_pin_click!(
                                "focus pick  by {}  at {},{}  file {}  offered {}",
                                if by_click { "click" } else { "key" },
                                pointer.point.x,
                                pointer.point.y,
                                trace_name(&path),
                                offered as u8,
                            );
                        }
                    }
                }
            }
        }

        // The selection's own answer, where the setting follows picks rather than hovers: the
        // file the listing under the pointer holds selected, whoever has the keyboard. A hover
        // needs the hand to settle on a file; a pick only needs the listing to hold one, so
        // this runs on every such tick rather than only where the press or the focus above
        // answered (see `PinUpdateWatch::follow_selection`).
        if !on_hover {
            self.follow_selection(resolver, &pointer, &showing, over_explorer, pick);
        }
    }

    /// Whether the press that was just read is a pick of its own, rather than a later press in
    /// the same gesture — one gesture at one point rather than two choices. The press is noted on
    /// the way out, so the next one is measured against this one.
    ///
    /// This is the whole of the rule that keeps a folder change from being answered as a pick.
    /// A press is read against whatever is under the pointer at the tick it lands, and a
    /// double-click that opened a folder leaves the new listing drawn exactly where the hand
    /// already was: the second press was answered with a file that arrived under a pointer
    /// nobody moved. The place cannot tell those apart, because by the time the second press
    /// lands the folder has already changed and "where the press was made" and "where the
    /// pointer is" are one place. The hand can, because a pick in a new listing is a file the
    /// hand travelled to.
    ///
    /// A press within the move tolerance of the last one *and* inside the machine's own
    /// double-click window is therefore the other half of a gesture already read, and is not
    /// offered — through either door, the file under the point and the listing's own selection.
    /// It is still noted, so the third press of a triple-click is refused the same way rather
    /// than answering on the strength of the first.
    ///
    /// The window is the other half of what makes a double-click a double-click, and it is what
    /// this rule was reading without. Measured on the spot alone, the second press of the
    /// double-click that opened a folder is the same thing as the click a user makes on the file
    /// that folder then drew under a hand they never moved: the same pixel, and the button went
    /// down in both cases. So the folder opens and every click at that pixel after it is refused
    /// for good — which is the whole of the fault, because the pin is standing over that pixel
    /// and stops following the listing it is standing in. Read in time as well, the press that
    /// opened the folder is still inside the gesture's own window while the click after it is
    /// not, and a click past that window is a file the user chose however little the hand has
    /// moved since.
    ///
    /// What this does not cost: a click on a different file is a different row, and the hand
    /// crosses a row to reach it. A click on the file the pin already shows is declined in
    /// silence by `offer` whatever this says, and a second click on one file — what a slow
    /// double-click on a file is — has nothing new to ask for either.
    pub(super) fn press_is_a_pick(&mut self, point: POINT, threshold: i32, now: Instant) -> bool {
        let pick = self.press_point.is_none_or(|(last, at)| {
            let past_the_gesture =
                now.saturating_duration_since(at) > Duration::from_millis(double_click_ms());
            let moved =
                (point.x - last.x).abs() > threshold || (point.y - last.y).abs() > threshold;

            past_the_gesture || moved
        });
        self.press_point = Some((point, now));
        pick
    }

    /// Begin again from the file a pin is showing now, and tell apart the two reasons the file can
    /// be one the watch has not seen: a pin taken up, and a pin that has been shown another file
    /// since the last tick.
    ///
    /// The pointer's own reading is taken again either way, because all of it is about what is
    /// under the hand rather than about the pin: the settle and the probe were measured against a
    /// window that has since moved, and the file under a pointer nobody moved is not a hover.
    ///
    /// A swap is not a new watch, though, and what a swap keeps — the item the keyboard is on, the
    /// place it was read in, the selection, whether the hand has arrived — belongs to the listing
    /// rather than to the pin. Forgetting it is what made the first selection a user makes after
    /// pressing the pin read as a baseline rather than as a pick. A pin taken up is the watch
    /// beginning, and it begins from nothing: what is on the keyboard when a pin comes up is a
    /// baseline and not a choice, whatever the listing had on it.
    pub(super) fn note_shown(&mut self, showing: &Path) {
        let mut next = Self {
            showing: Some(showing.to_path_buf()),
            ..Self::default()
        };

        if self.showing.is_some() {
            next.focused = self.focused.clone();
            next.place = self.place.clone();
            next.arrived = self.arrived;
            // The selection the listing holds goes with the listing, for the same reason the
            // keyboard's item does: a swap is the window being shown another file, and what the
            // listing has selected is where it was — so the first pick after one is still a
            // change, and still a pick (see `PinUpdateWatch::follow_selection`).
            next.sel_place = self.sel_place.clone();
            next.sel_selected = self.sel_selected.clone();
        }

        *self = next;
    }

    /// Note what the keyboard and the pointer did on one tick, as the witness a focus moved by the
    /// keyboard is made of: a key walked a listing, or something moved the focus by other means.
    /// Everything else takes that reading away from whatever key press came before it, which is
    /// what a folder, a tab and a window are reached by (see `FocusMoveInput` and `walk_at`).
    pub(super) fn note_focus_move(&mut self, input: FocusMoveInput, now: Instant) {
        if input.walked_by_key {
            self.walk_at = Some(now);
        }

        if input.moved_otherwise || input.clicked {
            // Noted after the key, so that a chord — a Ctrl+PageDown, say — is the command it is
            // rather than the walk its keys look like.
            self.walk_at = None;
        }
    }

    /// Whether a focus that has moved onto another item is a file the user picked with the
    /// keyboard, out of the two facts that tell a key's move from a listing's own: the place the
    /// item was read in is the one this watch is already watching, and a key that walks a listing
    /// is what put the focus there (see `PinUpdateWatch::focus_moved_by_key`).
    ///
    /// The place is read and noted here rather than beside it, because **a place the shell cannot
    /// answer for is not a place this watch is watching**. The item has moved and the read of where
    /// it moved to has failed, so the place in hand describes the listing it came out of — a
    /// listing that is no longer on screen. Leaving it standing is what made the first selection a
    /// user makes after a folder change do nothing: the shell cannot describe a view while it is
    /// navigating it, so the landing on the new folder's first item is read with no place at all,
    /// and the *first* read that succeeds is the user's own key press — compared against the folder
    /// that has been left and read as a move, so the pick it is was offered to nobody. The second
    /// key press found the place already in hand and worked, which is the whole of what was seen:
    /// one selection lost per folder change, every one after it answered.
    ///
    /// Dropping it costs nothing: with no place in hand the next read establishes one, and a move
    /// the keyboard did not make is not offered either way — `focus_moved_by_key` is what separates
    /// them, not the place.
    pub(super) fn focus_pick(
        &mut self,
        place: Option<HoverLocation>,
        known: bool,
        now: Instant,
    ) -> bool {
        let landed = self.note_place(place);

        known && !landed && self.focus_moved_by_key(now)
    }

    /// Whether the keyboard's own keys are what put the focus where it is: a key that walks a
    /// listing was seen within the window a key press is given to have moved something, and nothing
    /// that moves the focus by other means has been seen since.
    ///
    /// It is asked of a focus that has moved beside the place it moved in, because neither answers
    /// alone. The place a focused item is read in is the view the *pointer* was last seen working in
    /// rather than the one the item is drawn in — an item's provider reports no window for a frame's
    /// views to be told apart by — which leaves a tab switched, a folder opened with Enter and a
    /// window moved to reading as the place the watch was already watching.
    pub(super) fn focus_moved_by_key(&self, now: Instant) -> bool {
        recent_elapsed_within(
            self.walk_at.map(|at| now.saturating_duration_since(at)),
            KEYBOARD_FOCUS_INPUT_GRACE_MS,
        )
    }

    /// Whether the click in hand is what moved the focus onto this item: the click's own pick, read
    /// out of the item the view has the focus on rather than out of the point under the hand.
    ///
    /// It is the pick the keyboard's own rule reads — an item the focus has moved to, in a place the
    /// watch was already watching — asked of a click instead of a key, because a click *is* a focus
    /// change the view reports and this app never sees: Explorer gives the item a click selects the
    /// keyboard focus, and what the view says about it is answered from the view's own side, where
    /// the hit test this watch's own look uses answers for whatever is standing on top.
    ///
    /// Three facts make a focus change this click's, and all three are needed: a click is in hand
    /// and its hold has not run out; the item is drawn where the click landed, which is the box the
    /// view gives it; and the item is read in the place the click was *made in*. The place is
    /// asked of the click rather than of the watch because the watch's own place is where the
    /// *keyboard* was last read, and the click is the one thing here that says which listing the
    /// pointer was working in.
    pub(super) fn click_picked_item(
        &self,
        bounds: (i32, i32, i32, i32),
        place: Option<&HoverLocation>,
        now: Instant,
    ) -> bool {
        let Some(pending) = &self.pending else {
            return false;
        };

        if now.saturating_duration_since(pending.at) > Duration::from_millis(PIN_CLICK_RETRY_MS) {
            return false;
        }

        if !point_in_box(pending.point, bounds) {
            return false;
        }

        let (Some(clicked_in), Some(here)) = (pending.place.as_ref(), place) else {
            return false;
        };

        !hover_location_changed(clicked_in, here)
    }

    /// Whether the listing the click in hand was made in is no longer the one under the pointer.
    ///
    /// The place is what a retry cannot say for itself: read without it, a listing that replaced
    /// the one the click was made in answers the retry about the *point* with whatever the fresh
    /// listing happens to have under a hand nobody moved. A click whose place was never read is not
    /// answered as changed — a shell that did not describe the view is not a listing that moved,
    /// and the click is left to its own look at the point (see `PendingClick::place`).
    fn click_place_changed(&self, resolver: &mut ItemResolver, pointer: &PointerTick) -> bool {
        let Some(clicked_in) = self.pending.as_ref().and_then(|held| held.place.as_ref()) else {
            return false;
        };

        click_place(resolver, pointer).is_some_and(|here| hover_location_changed(clicked_in, &here))
    }

    /// Note the place the item the keyboard has landed on was read in, and answer whether the focus
    /// has landed somewhere the watch was not watching: another folder, another tab of one window,
    /// another window. What such a move lands on is the baseline the next key is measured against
    /// rather than a file the user picked — the three things this setting follows are a click, a
    /// hover and the keyboard, and a listing that changed under the keyboard is none of them.
    ///
    /// A look that answered nothing is `false` and takes the baseline with it, because the item has
    /// moved and the only place in hand is the one it moved out of: a view cannot be described while
    /// it is navigating, so the landing on a new folder's first item is read with no place at all,
    /// and keeping the old one made the user's next key press — the first read since the listing was
    /// replaced — read as the folder change rather than as the pick it is (see
    /// `PinUpdateWatch::focus_pick`). With no place in hand the next read establishes one.
    pub(super) fn note_place(&mut self, place: Option<HoverLocation>) -> bool {
        let Some(place) = place else {
            self.place = None;
            return false;
        };

        take_place(&mut self.place, place)
    }

    /// What an answer to a click does to the hold on it: an answer the pin is already showing
    /// holds it, anything else ends it.
    ///
    /// The file the pin is already showing is not an answer, it is the *absence* of one, and
    /// letting it end the hold is what made the first click after a pin was focused the one that
    /// was lost: `offer` declines such a file in silence, and the decline was read as "resolved",
    /// so nothing was offered, nothing waited, and the next click on another file — read a tick
    /// later, against caches the click's own drop had by then cleared — was the first one answered.
    /// Every time, on any file, the first click only.
    ///
    /// So a click is held until the shell names a file that is not the one on screen, or until its
    /// own window runs out (`PIN_CLICK_RETRY_MS`). A hand that clicks the file already on screen
    /// pays for the hold and nothing else: the answer it keeps getting is the answer, and the pin
    /// is left exactly as it was, which is what clicking the file you are already looking at is
    /// supposed to do.
    ///
    /// A click in hand is one read on this tick or one already being held for a retry; a hover
    /// is neither and arms nothing — a hand that settles on the file already on screen has asked
    /// for nothing, so there is nothing to hold on its behalf.
    ///
    /// The hold's own clock is set by the click that started it and by nothing else. Re-arming it
    /// on a retry would slide the window along with every tick told "not the file on screen", and
    /// a click that is being spent rather than answered would then be held for as long as the
    /// hand stayed still over it.
    pub(super) fn answer_click(
        &mut self,
        path: &Path,
        showing: &Path,
        clicked: bool,
        now: Instant,
        point: POINT,
    ) {
        if same_path(path, showing) && (clicked || self.pending.is_some()) {
            if clicked {
                // The place is carried forward rather than re-read. A second click on the file
                // the pin already shows re-arms the hold on a *later* tick, and the listing it
                // was made in is the same one: the press bit is spent, so this tick is a tick
                // after the click that needed it, and by then the view under the pointer is the
                // answer to the retry rather than a new place. Reading it fresh here would
                // compare the retry against the listing that click itself produced, which never
                // differs, and the click would never be released.
                let place = self.pending.take().and_then(|held| held.place);
                self.pending = Some(PendingClick {
                    at: now,
                    point,
                    place,
                });
            }

            return;
        }

        self.pending = None;
    }

    /// Offer a file to the pin that is up: the one thing this watch does, and the only
    /// thing in it that asks for anything.
    ///
    /// Nothing is offered where the pin is already showing the file, and nothing where the file is
    /// not one this app previews at all — a name no kind claims is a file Explorer selects and the
    /// pin has nothing to show for, and a pin that went blank or took a hover of its own for one
    /// would be worse than a pin that stayed as it was (see `is_media_file`). Which of the two it
    /// was is returned, because the caller has to tell them apart: the first is the silence that a
    /// stale listing answers with (see `PinUpdateWatch::answer_click`) and the second is an answer
    /// (see `PinUpdateWatch::answer_click` again).
    fn offer(&self, path: &Path, showing: &Path) -> bool {
        if same_path(path, showing) {
            return false;
        }

        if !is_media_file(path) {
            return true;
        }

        update_pinned_preview(path);
        true
    }

    /// Follow the listing's own selection for one tick: the file the view under the pointer
    /// holds selected, offered wherever it is another pick in the listing already watched.
    ///
    /// This is what a pin follows when `Update Preview` is on and `On Hover` is off, and it is
    /// the selection's answer to the hover's reading: the first click after the pin takes the
    /// focus is exactly the pick the press and the focus both miss, because the press is spent
    /// before the shell has caught up with it and the focus is not read until Explorer is in
    /// front again.
    ///
    /// Nothing is asked where the pointer is not over Explorer at all: the selection of a
    /// listing the hand is not in is not a file anybody just picked. Nothing is asked twice
    /// a dozen times a second either: the poll runs on its own cadence, and a click schedules
    /// a read outside it (see `PIN_SELECTION_POLL_MS`).
    fn follow_selection(
        &mut self,
        resolver: &mut ItemResolver,
        pointer: &PointerTick,
        showing: &Path,
        over_explorer: bool,
        clicked: bool,
    ) {
        if !over_explorer {
            return;
        }

        let now = Instant::now();
        let due = self
            .sel_polled_at
            .map(|at| {
                now.saturating_duration_since(at) >= Duration::from_millis(PIN_SELECTION_POLL_MS)
            })
            .unwrap_or(true);
        if !clicked && !due {
            return;
        }
        self.sel_polled_at = Some(now);

        // The place is what a pick in this listing is told apart from a listing that replaced
        // it by, and it is read without the folder walk the hover's own place pays for: the
        // URL a view was opened with is the folder for a folder view and the query for a
        // search, which is the comparison this needs (see `click_place`). A shell that
        // describes nothing leaves the last baseline standing rather than answering either
        // way.
        let Some(place) = click_place(resolver, pointer) else {
            return;
        };
        let selected = selected_file_in_view(resolver, pointer);

        if let Some(path) = self.note_selection(place, selected, clicked) {
            self.offer(&path, showing);
        }
    }

    /// Note what the listing under the pointer has selected, and answer which file the pin is
    /// owed, if any: the file itself where this sighting is a pick, and nothing where it is a
    /// baseline or no change at all.
    ///
    /// A sighting is a pick by two facts, and only two. The listing is the one already
    /// watched — a folder, a tab or a window moved to is a fresh baseline, offered nothing,
    /// the way a landing is (see `PinUpdateWatch::note_place`) — and, where the listing is a
    /// new one, only a click in hand offers out of it: the click selects before the shell
    /// reports it, so the selection a click has just made in a listing this watch has not
    /// seen is still that click's pick. The very first sighting is a baseline too: which file
    /// a listing holds before the hand does anything in it is not a file anybody picked.
    ///
    /// A click whose selection the shell has not caught up with offers nothing either — the
    /// sighting keeps the old selection, so the change answers on the tick the shell reports
    /// it on. What the pin is already showing is answered the same way by the caller: it is
    /// offered and declined, which is what clicking the file on screen is supposed to do.
    pub(super) fn note_selection(
        &mut self,
        place: HoverLocation,
        selected: Option<PathBuf>,
        clicked: bool,
    ) -> Option<PathBuf> {
        let known = self.sel_place.is_some();
        let changed = take_place(&mut self.sel_place, place);

        if clicked {
            // The click selects what it is on. Where the shell has caught up the selection
            // is the pick; where it has not, `selected` is the old file or nothing, and the
            // baseline above keeps the old answer standing for the change below.
            if let Some(path) = selected {
                self.sel_selected = Some(path.clone());
                return Some(path);
            }
            return None;
        }

        if !known || changed {
            self.sel_selected = selected;
            return None;
        }

        let Some(path) = selected else {
            // A poll that names no file in the listing already watched is a miss, not a
            // deselection: keep the baseline standing so the next sighting of the same
            // file is not read as a pick. A miss is what a pointer over a gap, a UIA
            // element that names no item, or a transient shell failure answers with,
            // and moving the mouse makes each of them likelier on every poll.
            return None;
        };

        if self.sel_selected.is_none() {
            // The baseline above was a miss (the first sighting named no file): the
            // first file actually seen is the baseline, not a pick.
            self.sel_selected = Some(path);
            return None;
        }

        if Some(&path) != self.sel_selected.as_ref() {
            self.sel_selected = Some(path.clone());
            return Some(path);
        }

        None
    }
}
