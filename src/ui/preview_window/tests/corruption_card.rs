//! The card a pinned window shows when the file it is showing may be damaged, and the two things
//! that keep it honest: it belongs to one file, and it is never drawn on a hover.

use super::pin_film_host_window::StandingFilm;
use super::*;

/// Where a window of the suite stands on the screen.
unsafe fn window_rect(hwnd: HWND) -> ScreenRegion {
    let mut rect = RECT::default();
    let _ = unsafe { GetWindowRect(hwnd, &mut rect) };
    (rect.left, rect.top, rect.right, rect.bottom)
}

/// A folder of this test's own under the machine's temp directory, emptied first.
fn a_folder(name: &str) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-corruption-card")
        .join(name);
    let _ = std::fs::remove_dir_all(&folder);
    let _ = std::fs::create_dir_all(&folder);
    folder
}

/// A pin taken up on `path` while `pid` is the player the app is showing it through.
fn a_pin_of(path: &Path, pid: u32) -> PinnedPreview {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    take_up_pinned_window(path, (100, 100, 400, 3_000))
}

#[test]
fn a_damaged_file_does_not_hand_its_card_to_the_file_stepped_to() {
    let _room = the_room_to_itself();
    let folder = a_folder("step");
    let damaged = folder.join("damaged.mkv");
    let healthy = folder.join("healthy.mkv");

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.load(Ordering::SeqCst);

    // The player showing the damaged file has said so in its own log.
    let first_watch = player_logs::a_watch();
    first_watch.observe(
        "[mkv] Corrupt file detected. Trying to resync starting from position 20511866...",
    );
    player_logs::remember(1_111, &damaged, &first_watch);

    let first = a_pin_of(&damaged, 1_111);
    assert!(
        first.corruption.tripped(),
        "the file the player accused carries the card"
    );
    assert!(
        first.corruption.coverage_at(1_000).is_some(),
        "and the card is drawn on it"
    );

    // The step: the next file is shown by a player of its own, which has said nothing.
    let second_watch = player_logs::a_watch();
    player_logs::remember(1_115, &healthy, &second_watch);
    let second = a_pin_of(&healthy, 1_115);

    assert!(
        !second.corruption.tripped(),
        "a step to another file is not a step to its accusation"
    );
    assert_eq!(
        second.corruption.coverage_at(1_000),
        None,
        "and no card is drawn on it"
    );

    // **And it is the file the player was begun for, not the id alone.** A step whose launch has
    // not landed yet still answers the old player's id, and a card taken up on the answer would be
    // the last file's card standing over this one.
    let stale = a_pin_of(&healthy, 1_111);
    assert_eq!(
        stale.corruption.coverage_at(1_000),
        None,
        "the old player's accusation is not this file's"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    let _ = std::fs::remove_dir_all(&folder);
}

#[test]
fn the_card_is_asked_of_the_file_it_was_taken_up_for_and_not_of_the_one_before_it() {
    let _room = the_room_to_itself();
    let folder = a_folder("taken-up");
    let first = folder.join("first.mkv");
    let second = folder.join("second.mkv");

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.load(Ordering::SeqCst);

    let watch = player_logs::a_watch();
    watch.observe("Failed to recognize file format.");
    player_logs::remember(1_201, &first, &watch);

    let pin = a_pin_of(&first, 1_201);
    let held = pin
        .corruption
        .watch()
        .expect("the pin holds the watch of the player it was taken up on");
    assert!(
        Arc::ptr_eq(held, &watch),
        "the watch is the one created for that player and no copy of it"
    );

    // And a pin taken up on the file beside it holds nothing at all.
    player_logs::remember(1_205, &second, &player_logs::a_watch());
    let beside = a_pin_of(&second, 1_205);
    assert!(!beside.corruption.tripped());

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    let _ = std::fs::remove_dir_all(&folder);
}

#[test]
fn the_card_of_a_hosted_film_stands_over_it() {
    // **The film composites over the pin's own surface**, so a card painted into that surface is a
    // card under the picture. What carries it over is a strip of its own — the road the caption and
    // the volume popup take — and this is the pin standing with a hosted film, which is the
    // arrangement the card has to survive (see `pin_chrome_strips`).
    let _room = the_room_to_itself();
    let film = StandingFilm::new();

    // The film's own player has said the file is damaged.
    let watch = player_logs::a_watch();
    watch.observe(
        "[mkv] Corrupt file detected. Trying to resync starting from position 20511866...",
    );
    with_pin(|pin| pin.corruption = CorruptionCard::watching(Some(Arc::clone(&watch))));

    unsafe { show_pinned_window(film.pin_hwnd) };
    let card = pin_chrome_strips::card_strip().expect("the card's strip of its own");
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };

    assert!(
        unsafe { IsWindowVisible(card) }.as_bool(),
        "a card up over a hosted film is a card the strip carries"
    );
    assert_eq!(
        unsafe { window_rect(card) },
        film.box_,
        "and it stands over the whole band the picture is seen through: the scrim dims it"
    );

    // The card's one run over, the strip is off the screen again: there is no state left over for a
    // pin to carry into the next file.
    with_pin(|pin| pin.corruption = CorruptionCard::none());
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        !unsafe { IsWindowVisible(card) }.as_bool(),
        "and no card is a strip off the screen"
    );

    // The pin's window is the stand's own, made by the road a pin's window is made by rather than by
    // the suite's maker, and the stand puts it away as it drops (see `StandingFilm::drop`): putting it
    // away here as well would be the ledger's own tripwire firing on a handle that was never in it.
}

#[test]
fn a_hover_preview_is_never_drawn_a_card() {
    // A hover is painted by `paint::render_layered_preview_at` and a pin by the function beside it,
    // so what makes the card pin-only is that the hover's painter never reaches it. Read out of the
    // source rather than asserted about a window: the two painters are one file apart from each
    // other, and the whole of the guarantee is which of them names the card at all.
    let painter = include_str!("../paint.rs");
    let hover = painter
        .split_once("pub(super) unsafe fn render_layered_preview_at")
        .expect("the hover's painter")
        .1;
    let hover = hover
        .split_once("pub(super) unsafe fn render_pinned_preview_at")
        .expect("the pin's painter, which follows it")
        .0;

    assert!(
        !hover.contains("corruption"),
        "a hover preview draws no card, whatever its player has said"
    );
}
