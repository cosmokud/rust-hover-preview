use super::*;

/// A pin up, with nothing being dragged and nothing on the keyboard.
fn a_pin_awaiting_a_press() {
    stand_pin(Some(PinnedPreview::for_test()));
}

/// A press that lands on a pin takes the pointer for the drag it begins, and a drag it
/// could not begin lets the pointer go instead.
///
/// The two halves are one decision and only one of them used to be reachable. Taking the
/// pointer is what a press on a pinned window is for — a drag answers nothing until the hand
/// lets go, so without the capture the window stops following the hand — but the capture was
/// taken unconditionally, and a pin taken down between the press being read and the drag being
/// installed left the window holding the pointer for the whole desktop with nothing that
/// would ever release it. Every mouse message then went to that window rather than to
/// whatever the pointer was aimed at, and clicking elsewhere appeared to bring it back.
///
/// The order is load-bearing as well as the two halves: the pointer is asked for *before* the
/// window's own box, so a machine that will not say where the pointer is never gets asked
/// where its window stands — there is nothing to measure a drag against, and a drag begun
/// from a box alone is a drag from the wrong origin.
#[test]
fn a_press_takes_the_pointer_for_its_drag_and_a_drag_it_cannot_begin_lets_it_go() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let window = a_window_at((300, 200, 700, 600));
    let hwnd = HWND(0x1000 as *mut _);
    a_pin_awaiting_a_press();
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);

    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::Pointer(Some((300, 200))),
            PinWindowCall::WindowBox(Some((300, 200, 700, 600))),
            PinWindowCall::Capture,
        ],
        "a press that begins a drag takes the pointer, and asks for the pointer before the \
             box because the box alone is a drag from the wrong origin"
    );

    // No pin to put the drag in, and the pointer is let go rather than left taken. This is
    // the branch that used to be unreachable: the capture was taken on the way in and only
    // released by a release that a window with no drag never sees.
    let window = a_window_at((300, 200, 700, 600));
    stand_pin(None);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);

    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::Pointer(Some((300, 200))),
            PinWindowCall::WindowBox(Some((300, 200, 700, 600))),
            PinWindowCall::ReleaseCapture,
        ],
        "a drag that could not be installed lets the pointer go, because a window holding it \
             with nothing to release it takes every mouse message on the desktop"
    );
}

/// A press on a window that cannot say where it is takes nothing at all.
///
/// Both of the refusals are real answers a real machine gives — `GetCursorPos` and
/// `GetWindowRect` can both be refused — and the second one is the dangerous case: the drag
/// was already installed by the time the window's own box is asked for, so a refusal here
/// leaves a drag in the pin with no origin to measure against and no capture to have been
/// taken, which is a drag the loop carries on from a window that is not where it was.
///
/// So the order is what answers it: the window's box is asked for *before* the drag is
/// installed, and a refusal returns with nothing installed and nothing taken.
#[test]
fn a_press_that_cannot_be_measured_takes_nothing_at_all() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    for window in [
        RecordedPinWindow::with(0x1000, None, Some((300, 200, 700, 600))),
        RecordedPinWindow::with(0x1000, Some((300, 200)), None),
    ] {
        a_pin_awaiting_a_press();
        begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);

        assert!(
            !window
                .calls()
                .iter()
                .any(|call| matches!(call, PinWindowCall::Capture)),
            "{:?}: a drag that cannot be measured is not begun, and nothing is taken for it",
            window.calls()
        );
        assert!(
            carried_drag().is_none(),
            "{:?}: and no drag is left installed with no box to measure against",
            window.calls()
        );
    }
}

/// The two ends of a capture are one list, and they are asked of the same window.
///
/// A drag begun by a message is ended by the message that releases it and a drag begun out of
/// the hook's published button state is ended by the tick instead, but both are the same
/// work: the pointer this window took for the drag is let go of. They were two hand-written
/// halves — a `SetCapture` on the way in and a `release_pin_capture` on the way out — and
/// nothing said they had to agree about which window.
///
/// The player's window is **not** on this list any more, and the reason is the second of the two
/// flashes a drag used to have. A move has no relaunch behind it, so the release used to put the
/// picture back itself — the flag down, the band transparent, and the compositor's first look at a
/// window that has been hidden for the length of a drag still filling in. The band is now filled
/// with the frame the drag was holding until there is a player to see through it, and that is the
/// settle's question on the loop's tick rather than a window procedure's (see
/// `settle_pinned_park`). What this asserts is the negative of it: the release puts the pointer
/// back and repaints, and touches nothing of somebody else's.
///
/// The band is moved by the hand while the drag runs, which is what a move *is*, and the player's
/// window has not followed it: the park is a hide and nothing else, so the rect the window is
/// still standing at is the one the drag started from. The swap puts it back at the band as it
/// stands rather than the box the drag began at, which is the same place it did before — a tick
/// later.
#[test]
fn a_drag_lets_go_of_the_pointer_it_took() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let window = a_window_at((300, 200, 700, 600));
    let hwnd = HWND(0x1000 as *mut _);
    a_pin_awaiting_a_press();
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);
    with_pin(|pin| pin.content = (900, 700, 1300, 1100));
    assert!(
        finish_pin_drag(hwnd, &window),
        "there was a drag to let go of"
    );

    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::Pointer(Some((300, 200))),
            PinWindowCall::WindowBox(Some((300, 200, 700, 600))),
            PinWindowCall::Capture,
            PinWindowCall::ReleaseCapture,
            PinWindowCall::Repaint,
        ],
        "the pointer is taken for the drag and given back when the drag is over, and the window \
             is drawn at where the hand left it — the pointer first, because a window still \
             holding it after the drag has gone eats every mouse message on the desktop"
    );

    // And a second end has nothing to *repaint*: the drag is taken out of the pin by the first,
    // so the second road finds nothing rather than drawing for a drag that has already been let go
    // of. It does still hand the pointer back, because that is the one thing a road cannot know it
    // has no business doing — a capture taken from under a drag is Windows' to report and nobody
    // else's to reason about, so every end lets go of it (see `release_the_pointer`). On the
    // machine that release finds no capture to give back, which is the answer the recorder records
    // rather than the absence of one.
    let after_the_first_end = window.calls().len();
    assert!(
        !finish_pin_drag(hwnd, &window),
        "a drag that is over is not ended twice"
    );
    assert_eq!(
        window.calls()[after_the_first_end..],
        vec![PinWindowCall::ReleaseCapture],
        "and the second end asks the window for nothing but the pointer — a repaint for a drag \
             that has already been let go of would draw a window nobody is carrying"
    );
}

#[test]
fn a_pinned_windows_bands_are_the_caption_above_it_and_the_bar_below_it() {
    // A window with no transport bar is its media with a caption on top; one that plays
    // carries the bar as well, which is a band the media gives up at the bottom.
    let window = (100, 100, 500, 600);
    assert_eq!(
        content_box_of(window, 96, false, false, pinned_caption_height(96, None)),
        (100, 130, 500, 600)
    );
    assert_eq!(
        content_box_of(window, 96, true, false, pinned_caption_height(96, None)),
        (100, 130, 500, 570)
    );

    // And a kind whose chrome is drawn over its media has no bands at all: its window is its
    // media, and the caption and the bar are strips *of* it rather than room beside it.
    assert_eq!(
        content_box_of(window, 96, false, true, pinned_caption_height(96, None)),
        window
    );
    assert_eq!(
        content_box_of(window, 96, true, true, pinned_caption_height(96, None)),
        window
    );
}

/// A key pressed on a pin that is the window the user is in walks the pin's own folder, and a
/// Space thrown at the same window holds a sound or lets it go.
///
/// The gate is the window itself, not a reading of the keyboard: these keys arrive here
/// because Windows routed them to the pin, and that is a thing Windows does and not a
/// thing this app polls for. So the mapping is the whole of what decides which keys a pin
/// answers, and a key it does not answer is a key the listing behind it is sent as usual.
#[test]
fn a_key_reaches_a_pin_by_being_the_one_the_keyboard_is_in() {
    for (vk, expected, name) in [
        (VK_LEFT.0 as i32, Some(PinCommand::Previous), "left"),
        (VK_UP.0 as i32, Some(PinCommand::Previous), "up"),
        (VK_RIGHT.0 as i32, Some(PinCommand::Next), "right"),
        (VK_DOWN.0 as i32, Some(PinCommand::Next), "down"),
        (VK_ESCAPE.0 as i32, Some(PinCommand::Close), "escape"),
        (VK_SPACE.0 as i32, Some(PinCommand::TogglePlayback), "space"),
    ] {
        assert_eq!(
            pinned_key_command(vk),
            expected,
            "{name} is the key a pin answers with {expected:?}"
        );
    }

    // Nor is any other key the pin's to answer. An unmodified letter belongs to whatever
    // the user is typing into, and a modified one is a command of some other program: both
    // are keys this window was never asked about, and answering them would be a preview
    // acting on the keyboard it is merely standing in front of.
    for vk in [
        VK_A.0 as i32,
        VK_C.0 as i32,
        0x41,
        0x0D,
        0x09,
        0x2E,
        0x21,
        0x22,
    ] {
        assert_eq!(
            pinned_key_command(vk),
            None,
            "{vk:#x} is nobody's to answer with"
        );
    }
}

/// A key held down is one press and not a run of them. Windows repeats the key it is holding,
/// and a play/pause answered on every repeat is a sound flickering between playing and held
/// for as long as the hand is on the key; a walk is a thing a hand can sensibly hold down, so
/// the arrows still repeat.
#[test]
fn a_key_this_app_is_still_holding_is_pressed_once() {
    let held = KEY_REPEAT;

    assert_eq!(
        pinned_key_down_command(VK_SPACE.0 as i32, 0),
        Some(PinCommand::TogglePlayback),
        "a Space pressed is a play/pause"
    );
    assert_eq!(
        pinned_key_down_command(VK_SPACE.0 as i32, held),
        None,
        "and a Space held is one press, not a sound flickering for as long as it is down"
    );
    assert_eq!(
        pinned_key_down_command(VK_LEFT.0 as i32, held),
        Some(PinCommand::Previous),
        "while an arrow held still walks, which is what holding one is for"
    );
}

/// A Space in a pin is the pause of the player the file on screen is played by, and that is
/// the file's answer rather than the key's: a video is played by the pin's own transport and
/// a sound by the clock behind its card, so the one key is two different actions.
///
/// What a failure here means is one of the two swallowed. A video read as a sound's is a
/// video that cannot be held with the keyboard at all — a picture that runs to its end while
/// the bar underneath it holds what a press on it holds — and a sound read as a video's is a
/// card with no clock moved under it, drawn at a second nothing is playing to.
#[test]
fn a_space_is_the_pause_of_the_player_the_file_is_played_by() {
    for kind in [MediaType::NativeVideo, MediaType::Video] {
        assert_eq!(
            pin_toggle_target(Some(kind)),
            PinToggle::Video,
            "{kind:?} is played by the pin's own transport, which is what a Space in one holds"
        );
    }

    assert_eq!(
        pin_toggle_target(Some(MediaType::Audio)),
        PinToggle::Audio,
        "a sound is played by the clock behind its card, and that clock is what a Space moves"
    );
}

/// A Space in a pin of anything that is not being played does nothing at all, and doing
/// nothing is the whole of the answer.
///
/// A picture and a page have no player behind them that a key could hold, and a pin whose
/// media is not there has nothing a key is about. Guessing for them would be worse than
/// refusing: a key answered with the wrong player's pause is a pin acting on a file that has
/// no playback — the exact swallowing this is, arrived at from the other side.
#[test]
fn a_space_pressed_on_something_that_is_not_playing_does_nothing() {
    for kind in [
        MediaType::StaticImage,
        MediaType::AnimatedGif,
        MediaType::Pdf,
        MediaType::Text,
        MediaType::Archive,
    ] {
        assert_eq!(
            pin_toggle_target(Some(kind)),
            PinToggle::None,
            "{kind:?} is drawn rather than played, so there is no player for a Space to hold"
        );
    }

    assert_eq!(
        pin_toggle_target(None),
        PinToggle::None,
        "and a pin with no media up has no player at all for a key to be about"
    );
}

/// A seek taken of a sound that is playing leaves it playing, and that is the whole of what a
/// second seek depends on: a sound left recorded as held reads as a pause to the next press,
/// so the card stands at one second over a sound that is still going — and no further seek
/// does anything at all, because a hold is what a seek moves rather than one it follows,
/// until a key is pressed and lifted again.
#[test]
fn a_seek_of_a_playing_sound_leaves_it_playing() {
    let (started, from, held) = pinned_audio_after_seek(true, true, 90.0);
    assert!(started.is_some(), "a player came up for the second named");
    assert_eq!(from, 90.0, "and the card's clock is counted from it");
    assert_eq!(
        held, None,
        "while a sound that is playing is not left recorded as held"
    );

    // So the next press is a second seek rather than a hold being moved, and the sound goes on
    // being played while the card follows it to the new second.
    let (started, from, held) = pinned_audio_after_seek(held.is_none(), true, 30.0);
    assert!(
        started.is_some(),
        "and it is a real seek, with a player of its own"
    );
    assert_eq!(from, 30.0, "begun at the second the hand named");
    assert_eq!(held, None, "and still nothing held over it");

    // A player that did not come up is a card with no clock rather than a hold over a sound
    // that is not playing: a decoder that would not have the file gives that answer at any level.
    assert_eq!(
        pinned_audio_after_seek(true, false, 90.0),
        (None, 90.0, None),
        "a sound with no player behind it has neither a clock nor a pause"
    );

    // And a sound a key is holding is held at the second rather than begun again by it, so
    // the key is what sets it going, from where the hand put it.
    assert_eq!(
        pinned_audio_after_seek(false, false, 90.0),
        (None, 90.0, Some(90.0)),
        "a seek of a held sound moves the hold rather than starting it"
    );
}

/// Two files of the *same* loudness were played two decibels apart, because the meter was
/// reading a peak and not the loudness: one file's loudest sample happened to sit at `0 dBFS`
/// and the other's at `-2 dBFS`, and a gain that brings a peak to full scale is therefore a
/// gain that says how hard the file was limited rather than how loud it is. The two reports
/// below are the whole of that, verbatim: `Bayangan Di Cermin3.mp3` and `Bayangan Di Cermin3
/// (1).mp3`, which measure `-13.3 LUFS` each and whose loudest samples are 2 dB apart.
///
/// What is read is the meter that measures loudness — FFmpeg's `ebur128`, whose `Summary` is
/// ITU-R BS.1770 integrated loudness — and what comes of it is the same gain for both.
#[test]
fn a_files_loudness_and_not_its_peak_is_what_normalize_measures() {
    const LIMITED_MASTER: &str = "\
[Parsed_ebur128_0 @ 00000204639ecd80] t: 183.499979 TARGET:-23 LUFS    M:-94.7 S:-94.6     I: -13.3 LUFS       LRA:   4.6 LU  FTPK: -87.9 -87.6 dBFS  TPK:   0.0   0.0 dBFS
[Parsed_ebur128_0 @ 00000204639ecd80] Summary:

  Integrated loudness:
    I:         -13.3 LUFS
    Threshold: -23.4 LUFS

  Loudness range:
    LRA:         4.6 LU
    Threshold: -33.4 LUFS
    LRA low:   -16.4 LUFS
    LRA high:  -11.8 LUFS

  True peak:
    Peak:        0.0 dBFS
";
    const LESS_LIMITED_MASTER: &str = "\
[Parsed_ebur128_0 @ 00000193854524c0] Summary:

  Integrated loudness:
    I:         -13.3 LUFS
    Threshold: -23.4 LUFS

  Loudness range:
    LRA:         4.3 LU
    Threshold: -33.4 LUFS
    LRA low:   -16.2 LUFS
    LRA high:  -11.9 LUFS

  True peak:
    Peak:       -1.9 dBFS
";

    let limited =
        audio_gain_from_report(LIMITED_MASTER).expect("a loudness report is a gain to apply");
    let less_limited =
        audio_gain_from_report(LESS_LIMITED_MASTER).expect("a loudness report is a gain to apply");

    // `-13.3 LUFS` against a target of `-14` is `-0.7 dB` for both files and the same
    // `-0.7 dB` for the pair, which is the whole of what went wrong: two files within a tenth
    // of a dB of each other are now within a tenth of a dB of each other.
    assert!(
        (limited - less_limited).abs() < 0.01,
        "two files of the same loudness were played {} dB apart",
        20.0 * (limited / less_limited).abs().log10()
    );
    assert!(
        (limited - 0.9226).abs() < 0.001,
        "a file at -13.3 LUFS is played at the target, not at its own peak: {} is not -0.7 dB",
        limited
    );

    // The whole of what the toggle is for: the file that is quiet by a *measurement* rather
    // than by a limitation is brought up to the target, and where that much gain would carry
    // the file's own peaks past the ceiling it is held there instead — `-6 dBFS` of true peak
    // has 5 dB of room under `-1 dBFS`, and 5 dB is what it is given of the 16 it asks for.
    let quiet = audio_gain_from_report(
        "\
[Parsed_ebur128_0 @ 00000193854524c0] Summary:

  Integrated loudness:
    I:         -30.0 LUFS

  True peak:
    Peak:       -6.0 dBFS
",
    )
    .expect("a quiet file is measured all the same");
    assert!(
        (quiet - 1.7783).abs() < 0.001,
        "a quiet file is lifted as far as its own headroom allows, not clipped past it: {}",
        quiet
    );

    // And a file of silence has nothing to measure and nothing to clip: its true peak is
    // `-inf`, which is the same answer a meter that failed gives, and it is played as it holds.
    assert_eq!(
        audio_gain_from_report(
            "\
[Parsed_ebur128_0 @ 0000010db76efe00] Summary:

  Integrated loudness:
    I:         -70.0 LUFS

  True peak:
    Peak:       -inf dBFS
"
        ),
        None,
        "a file with nothing to hear is left as it is rather than lifted off its floor"
    );
}

/// A file the probe says the engine can decode is not necessarily played by it: `Normalize`
/// puts a gain on a file, and a gain is an FFmpeg filter the engine cannot be handed — so such
/// a file is played by FFmpeg while every question asked while it plays used to branch on the
/// probe and speak to an engine with no session. That is a bar press that seeks nothing, on
/// exactly the quiet files a machine full of them has.
#[test]
fn a_gain_puts_a_probed_native_sound_on_ffmpegs_player() {
    // A gain of one is a gain like any other: it is the answer for a file that measures at the
    // target already, and it is played through the same filter as every other gain rather than
    // being read as a stand-in for "nothing was measured".
    assert_eq!(
        player_for_gain(Player::Native, true, Some(1.0)),
        Player::Ffmpeg,
        "a measured gain of one is still FFmpeg's filter, so FFmpeg is what plays the file"
    );

    // And one nothing has measured, which is the hover before the scan answers.
    assert_eq!(
        player_for_gain(Player::Native, true, None),
        Player::Native,
        "a file nothing has measured is played by the engine that was probed for it"
    );

    // And the case that broke: the engine has a decoder for the file, and a gain sends the
    // sound to FFmpeg regardless.
    assert_eq!(
        player_for_gain(Player::Native, true, Some(1.38)),
        Player::Ffmpeg,
        "a measured gain is FFmpeg's filter, so FFmpeg is what plays the file"
    );

    // With `Normalize` off there is no gain to apply whatever was measured, so the probe is
    // the whole of the answer again.
    assert_eq!(
        player_for_gain(Player::Native, false, Some(1.38)),
        Player::Native,
        "a gain measured while Normalize was off is not applied, so nothing moves the file"
    );

    // And a file the engine cannot play is FFmpeg's whatever its gain is — the answer is only
    // ever moved towards FFmpeg, never away from it.
    assert_eq!(
        player_for_gain(Player::Ffmpeg, true, Some(1.0)),
        Player::Ffmpeg,
        "a file the engine cannot decode stays with FFmpeg"
    );
}

/// A card is drawn at the second a key held a sound at rather than with no clock at all: a    /// player of this app's is ended to pause one, so the clock that measured it is gone, and a
/// card that has forgotten where the file was left says nothing about the pause — which is
/// the one thing a pause has to say.
#[test]
fn a_sound_a_key_held_is_drawn_where_it_was_held() {
    let folder = std::env::temp_dir().join("rust-hover-preview-pin-held-sound");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("song.mp3");
    std::fs::write(&path, b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00")
        .expect("a written file");
    audio_track::remember(
        &path,
        audio_track::Probed::Track(audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        }),
    );

    // A sound this app plays is measured by its own clock over the moment its player was
    // started, counted from the second it was started at — so a file dropped into the middle
    // of itself draws its bar where it is rather than at the beginning.
    let started = Instant::now() - Duration::from_secs(10);
    let (playing, length) = audio_clock(&path, Some(started), 30.0, None);
    assert!(
        playing.is_some_and(|at| (at - 40.0).abs() < 0.5),
        "thirty seconds into a file plus ten of playing is forty: {playing:?}"
    );
    assert_eq!(length, Some(180.0), "and the whole is the file's own");

    // A sound held is drawn where it was held, and at nothing else: the clock that measured
    // it is gone, so this is the only thing that can say where in the file it was left.
    assert_eq!(
        audio_clock(&path, Some(started), 30.0, Some(67.0)).0,
        Some(67.0),
        "a held sound stands at the second it stopped at"
    );
    assert_eq!(
        audio_clock(&path, None, 0.0, Some(67.0)).1,
        Some(180.0),
        "and the whole it is measured against is still the file's own"
    );

    // While a card with no player behind it says nothing about where the sound is at all,
    // which is the answer a file nothing will play gives.
    assert_eq!(
        audio_clock(&path, None, 0.0, None).0,
        None,
        "a card with nothing playing it has no clock to draw"
    );

    // And a sound still inside its file is never drawn past the end of it, whatever the
    // header said the length was. A container's length is a header's reading of itself and
    // is a little out for some formats, so wrapping this clock on it would put the card back
    // at 0:00 — with a sub-second position, which is a pixel or two of bar and a clock still
    // reading zero — a moment before the sound was really over, and the bar would then step
    // backwards as the pass turned over for real (see `wrap_audio_player`).
    let at = Instant::now() - Duration::from_secs(180);
    assert_eq!(
        audio_clock(&path, Some(at), 0.0, None).0,
        Some(180.0),
        "a sound at the whole of the file is drawn at the whole of the file"
    );

    let past = Instant::now() - Duration::from_secs(183);
    assert_eq!(
        audio_clock(&path, Some(past), 0.0, None).0,
        Some(180.0),
        "and a sound a moment past what the header said is still at the end of it, not at the start"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// The note that the keyboard was taken is dropped on every road out of a pin, and kept
/// everywhere else.
///
/// The note is what makes a handover happen at all. A window hidden while it still holds the
/// focus leaves Windows to pick what to activate next, and a `WS_EX_TOOLWINDOW` popup is not
/// reliably followed by the Explorer window that was in front a moment ago — so the claim is
/// what a teardown asks before it goes to the trouble of putting the keyboard back, and
/// leaving it standing on a road out is what strands a caret on a window that is not there.
///
/// Which window the user is in is not asked here, because it is not this claim's job: that is
/// `GetFocus` (see `pin_is_focused`), which is the one answer that cannot be wrong.
#[test]
fn the_keyboard_is_dropped_on_every_road_out_of_a_pin() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    install(PinnedPreview::for_test());
    assert!(
        !pin_holds_a_keyboard(),
        "a pin that has just come up has taken no keyboard, so owes nobody a handover"
    );

    // Windows taking the focus away is the user clicking into something else: the window now
    // in front holds the keyboard, so there is nothing to hand over and the claim is dropped.
    take_keyboard(0x2000, true);
    pin_release_focus();
    assert!(
        !pin_holds_a_keyboard(),
        "losing the focus drops the claim with it"
    );

    // A pin ending is the road that has to do the work: it drops the claim and hands the
    // keyboard back, and the handover is what that claim is asked for.
    take_keyboard(0x2000, true);
    end_pin(Reason::Closed, &Win32PinWindow);
    assert!(
        !pin_holds_a_keyboard(),
        "a pin that is over holds no keyboard, and remembers no window to hand it back to"
    );
}

/// The pin key brings a bubble back, and it does not take a window down any more.
///
/// Taking a window down is gone rather than narrowed, which is the point: a pin the user has
/// not pressed is a pin they are not in, and a key thrown at a window nobody is in is not a
/// request to close it.
#[test]
fn the_pin_key_only_brings_a_bubble_back_and_never_hides_a_window() {
    // The one thing it still does: a bubble is a window the user put away, and the key that
    // put it away brings it back on the file picked behind it.
    assert!(pin_key_restores_bubble(true, true));
    assert!(
        !pin_key_restores_bubble(true, false),
        "a bubble behind another program is not brought back by a key that program's"
    );

    // A window that is up is never brought back and never taken down, wherever the keyboard
    // is: a window already up and a key pressed at nothing are the two halves of a pin that
    // did nothing at all, which is the answer a Space now always gets.
    assert!(!pin_key_restores_bubble(false, true));
    assert!(!pin_key_restores_bubble(false, false));
}

/// The pin key is answered by what is *on screen*, and for the three kinds the engine
/// draws that is the engine's own window rather than any media of this app's — the slot
/// behind one is empty by design, or holds the spinner that stood in for the page until it
/// landed. Reading the slot alone is what made the key do nothing at all over an `.html`
/// with **Render HTML** on, an SVG document and a font specimen, while the very same key
/// worked over everything else.
///
/// What is not settled stays unpinnable: a spinner over a file of this app's own is still a
/// promise, and the engine saying it is showing a file is the only thing that settles the
/// three kinds it draws.
#[test]
fn what_is_on_screen_is_settled_by_the_engine_as_well_as_by_the_media() {
    // Everything this app draws is settled by its own media, and a spinner never is.
    assert!(pin_screen_is_settled(Some(MediaType::StaticImage), false));
    assert!(pin_screen_is_settled(Some(MediaType::Text), false));
    assert!(
        !pin_screen_is_settled(Some(MediaType::Loading), false),
        "a spinner is a promise rather than a preview"
    );
    assert!(
        !pin_screen_is_settled(None, false),
        "and nothing on screen is not a preview either"
    );

    // A document, a specimen and a page of HTML: the engine has put the file up, so there
    // is a preview on screen — with no media of ours behind it at all, which is the whole
    // of what used to stop the key.
    assert!(
        pin_screen_is_settled(None, true),
        "a page the engine is showing is a preview, whatever the media slot says"
    );
    assert!(
        pin_screen_is_settled(Some(MediaType::Loading), true),
        "the spinner left behind by a page that has landed does not un-settle it"
    );
}
