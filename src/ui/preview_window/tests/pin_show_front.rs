use super::*;

/// A pin that is a window of its own, off the screen, with `Pin Mode → Always On Top` off: the only
/// case that owes the topmost band and back to reach the front.
///
/// The fault this is a regression test for: a pin taken up with SPACE appeared *behind* Explorer. The
/// show asked for the front with `HWND_TOP`, and `HWND_TOP` does nothing for a process that does not
/// hold the foreground — the call reports success and the window stays where it was. A pin taken up
/// with a click did not have the fault, because the click had already made its window the foreground
/// one; a pin taken up with a key had not, and its window went where a hidden window sits: the bottom
/// of the z-order, behind everything opened over it since (see `show_pinned_window`).
///
/// The three facts are what the decision is made of, so they are the three things asked here.
#[test]
fn only_a_pin_that_is_its_own_window_and_off_the_screen_owes_the_band_to_reach_the_front() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let before_always_on_top = pin_always_on_top();
    let before_is_its_own = pin_is_its_own_window();

    // A pin that is up, because "the pin stands in a window of its own" is a fact about the pin on
    // the screen and not about the flag alone (see `pin_is_its_own_window`).
    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));

    let set = |its_own: bool, always_on_top: bool| {
        pin_stands_in_its_own_window(its_own);
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_always_on_top = always_on_top;
        }
    };

    set(true, false);
    assert!(
        showing_owes_the_band(false),
        "a picture pin that is not Always On Top and is not on the screen owes the band: it is the \
         one way a window that did not come with the foreground is put in front"
    );

    set(true, false);
    assert!(
        !showing_owes_the_band(true),
        "a pin already on the screen is left where the user's own clicks put it: a showing that \
         raised it every tick would be a pin that could not be put behind anything"
    );

    set(true, true);
    assert!(
        !showing_owes_the_band(false),
        "a pin that is Always On Top is going to the band above every other window anyway, so the \
         band it is already going to is not taken and given back first"
    );

    set(false, false);
    assert!(
        !showing_owes_the_band(false),
        "the kinds still standing in the preview's window keep the plain raise they have always \
         had: the press lands on that window itself, so the press is what has already brought it \
         forward"
    );

    stand_pin(previous_pin);
    pin_stands_in_its_own_window(before_is_its_own);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before_always_on_top;
    }
}

/// A raise owes the band unless the window asking for it is already the foreground one — and only a
/// pin that is a window of its own can have answered it that way.
///
/// The fault this is a regression test for: with `Pin Mode → Always On Top` off, clicking a file or a
/// preview update left the pin where it was instead of in front of the listing, because the raise was
/// skipped for a pin that is its own window on the grounds that Windows raises such a window — true of
/// a *press*, which activates it, and false of everything else: a pick arrives while the listing is
/// the foreground window (see `put_pin_at_the_front`, `raise_pinned_preview`).
#[test]
fn a_raise_is_answered_only_when_the_pin_is_a_window_of_its_own_and_holds_the_foreground() {
    let pin = HWND(1 as *mut _);
    let listing = HWND(2 as *mut _);

    assert!(
        raise_is_already_answered(true, pin, pin),
        "a pin that is a window of its own and is the foreground one is in front already"
    );
    assert!(
        !raise_is_already_answered(true, listing, pin),
        "a pin that is a window of its own with the listing in front owes the band: that is exactly          the pick and the preview update, which arrive while the user is clicking in the listing"
    );
    assert!(
        !raise_is_already_answered(false, pin, pin),
        "a pin standing in the preview's window keeps the band-and-back its press has always asked          for, even with its own window holding the foreground"
    );
    assert!(
        !raise_is_already_answered(false, listing, pin),
        "and owes it all the same with the listing in front"
    );
}
