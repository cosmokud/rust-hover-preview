use super::*;

/// **`Enable Preview` off is not `Pin Mode` off.** The two settings are different rows, and what the tick
/// does when no preview may be raised is the housekeeping a hover would have had — with the pin left
/// standing, because that window is the user's, and with the file under the pointer still resolved, because
/// the pin key names it (see `previews_are_held_back` and `publish_the_pointer_file`).
///
/// What ends a pin is `Pin Mode → Enable` off, its own close button, or its media going away — and nothing
/// here may end it for the other row.
#[test]
fn a_run_with_no_previews_still_resolves_a_file_and_still_keeps_a_pin() {
    assert!(
        previews_are_held_back(false, true),
        "the trigger key holding hovers back is the same no-hovers the row is"
    );
    assert!(
        previews_are_held_back(true, false),
        "and the `Enable Preview` row off is the other"
    );
    assert!(
        !previews_are_held_back(true, true),
        "with previews allowed the tick is the hover's, exactly as it was"
    );

    assert!(
        !a_pin_may_stand(false),
        "`Pin Mode → Enable` off refuses the press and takes the pin down"
    );
    assert!(
        a_pin_may_stand(true),
        "and a pin enabled is a pin the tick resolves a file for, whatever the row above says"
    );
}
