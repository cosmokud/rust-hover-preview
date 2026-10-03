use super::*;

/// A frame of `bytes` pixels, for a cache to hold.
fn a_frame_of(bytes: usize) -> Arc<ImageFrame> {
    Arc::new(ImageFrame::new(vec![0u8; bytes], 1, 1, 0))
}

/// A cache key for a file that is not there, named so a test can say which frame it is
/// looking at.
fn a_cache_key_for(name: &str) -> ImageCacheKey {
    ImageCacheKey {
        path: PathBuf::from(format!(r"C:\pictures\{name}")),
        version: FileVersion {
            modified: None,
            len: 0,
        },
        width: 1,
        height: 1,
    }
}

/// A cache holding one frame per name, each `bytes` long, stored in that order so that
/// `first` is the least recently used and `last` the most.
///
/// A cache of its own rather than the process-wide one, because the trim takes one as an
/// argument for exactly this reason: it is a decision about a set of frames and a number,
/// and the number comes from the configuration at every call so that an edit in the tray
/// applies without a restart (see `image_cache_limit_bytes`).
fn a_cache_of(names: &[(&str, usize)]) -> ImageCache {
    let mut cache = ImageCache::default();
    for (name, bytes) in names {
        cache.tick += 1;
        cache.entries.insert(
            a_cache_key_for(name),
            ImageCacheEntry {
                frame: a_frame_of(*bytes),
                bytes: *bytes,
                last_used: cache.tick,
            },
        );
        cache.bytes += bytes;
    }
    cache
}

/// The names a cache still holds, least recently used first.
///
/// Read in the order the trim decides in rather than in whatever order a hash map iterates,
/// because the order *is* the policy: a test that could not see it would pass against a
/// cache that dropped frames by some other rule entirely.
fn held_by_a_cache(cache: &ImageCache) -> Vec<String> {
    let mut held: Vec<(&ImageCacheKey, u64)> = cache
        .entries
        .iter()
        .map(|(key, entry)| (key, entry.last_used))
        .collect();
    held.sort_by_key(|(_, last_used)| *last_used);
    held.into_iter()
        .map(|(key, _)| {
            key.path
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .into_owned()
        })
        .collect()
}

/// The frame least recently used is the one dropped, and the rest are kept whatever their
/// sizes are.
///
/// This is the whole policy, and it is the opposite of the two rules it is not: a cache of
/// decoded frames that dropped the largest keeps re-decoding the file the user has just
/// moved on from — a 4K photograph is evicted by a 32-pixel icon, so the preview the user
/// is looking at is decoded again on every re-hover — and a cache that dropped by insertion
/// order evicts a frame that is being looked at right now in favour of one nobody has asked
/// for since the file was first hovered.
///
/// The budget is bytes rather than entries for the same reason: what this app runs out of is
/// memory, and a frame at the size of a display is 33 MB while the icon is a few kilobytes.
#[test]
fn the_frame_least_recently_used_is_the_one_dropped() {
    let mut cache = a_cache_of(&[
        ("oldest.png", 4_000),
        ("middle.png", 100),
        ("newest.png", 4_000),
    ]);

    // A budget that only two of the three fit inside, and the big ones are the ones that do
    // not fit — so a size-based rule would drop them and this rule does not.
    image_cache_trim(&mut cache, 5_000);

    assert_eq!(
        held_by_a_cache(&cache),
        vec!["middle.png", "newest.png"],
        "the frame nobody has asked for since it was first decoded is the one that goes, \
             whatever its size"
    );
    assert_eq!(
        cache.bytes, 4_100,
        "and the accounting is what is left, not what was held"
    );
}

/// A budget of zero means "hold nothing", and every other budget means what it says.
///
/// Zero is the one value a loop could spin on: `image_cache_trim` drops until the cache fits
/// inside the limit, and a cache that cannot be emptied would leave it dropping nothing with
/// the budget still unmet. That is why the loop breaks on an empty cache rather than asking
/// again, and it is the half of the trim that is a decision rather than arithmetic.
#[test]
fn a_budget_of_zero_empties_the_cache_rather_than_looping() {
    let mut cache = a_cache_of(&[("a.png", 4_000), ("b.png", 4_000), ("c.png", 4_000)]);

    image_cache_trim(&mut cache, 0);

    assert!(
        cache.entries.is_empty(),
        "`image_cache_mb = 0` means hold nothing, rather than hold everything until \
             something else is stored"
    );
    assert_eq!(
        cache.bytes, 0,
        "and nothing is still counted against the budget"
    );

    // And a budget the cache already fits inside drops nothing, which is the other direction
    // the loop has to get right.
    let mut cache = a_cache_of(&[("a.png", 100)]);
    image_cache_trim(&mut cache, 1_000_000);
    assert_eq!(held_by_a_cache(&cache), vec!["a.png"]);
    assert_eq!(cache.bytes, 100);
}

/// A budget that lands between two frames drops exactly as many as it takes, and stops.
///
/// The eviction is by a whole frame rather than a whole byte, so a budget that cannot be hit
/// exactly is overshot downwards rather than approximated: the alternative is dropping the
/// frame *after* the one that crossed the line, which frees memory in units of tens of
/// megabytes and frees far more than the budget was short by.
#[test]
fn the_budget_is_hit_by_whole_frames_and_no_further() {
    let mut cache = a_cache_of(&[("a.png", 1_000), ("b.png", 1_000), ("c.png", 1_000)]);

    // Room for two of three, and the third is dropped: exactly as many as it takes.
    image_cache_trim(&mut cache, 2_500);
    assert_eq!(held_by_a_cache(&cache), vec!["b.png", "c.png"]);

    // Room for one and a half, which no whole frame fits into twice over: one is dropped,
    // and dropping a second would free 1000 bytes the budget did not ask for.
    image_cache_trim(&mut cache, 1_500);
    assert_eq!(held_by_a_cache(&cache), vec!["c.png"]);
    assert_eq!(cache.bytes, 1_000);
}

/// A loop with nothing on screen does not wake for a caption that is not there, and one with a
/// caption that is does not wait as long.
///
/// The four bands and what each is for, in one list, because the numbers were tuned twice by
/// two commits that disagreed with each other and there was nowhere to ask whether a change
/// had broken the band next to it. The ordering is the claim: each band is strictly longer
/// than the one before it except where the caption says it should not be, and a band that
/// moved past its neighbour without anybody noticing is exactly what two tunings disagree
/// about looks like from the outside.
///
/// Nothing on screen is the slowest band and it ignores the other two entirely: there is no
/// picture to advance and no button to press, so the only thing the interval bounds is how
/// long a window message waits to be noticed.
#[test]
fn the_loop_s_own_cadence_is_the_slowest_thing_it_does() {
    assert_eq!(
        wait_before_the_next_tick(true, false, false),
        IDLE_WAIT_MS,
        "nothing on screen waits the longest, and is not a function of whether a pin is up"
    );
    assert_eq!(
        wait_before_the_next_tick(true, true, true),
        IDLE_WAIT_MS,
        "a pin and a wait do not make an idle loop turn: there is nothing to draw either"
    );

    assert_eq!(
        wait_before_the_next_tick(false, true, false),
        FRAME_WAIT_MS,
        "something moving turns at the frame rate, whatever is behind it"
    );
    assert_eq!(
        wait_before_the_next_tick(false, true, true),
        FRAME_WAIT_MS,
        "and a pin does not slow that down — a pinned video is still a video"
    );

    assert!(
        wait_before_the_next_tick(false, false, true)
            < wait_before_the_next_tick(false, false, false),
        "a pin's static tick is shorter than a hover's, because a pin's caption carries the \
             buttons that close it and a static preview's carries nothing: {} against {}",
        wait_before_the_next_tick(false, false, true),
        wait_before_the_next_tick(false, false, false)
    );
    assert!(
        wait_before_the_next_tick(false, false, false) < IDLE_WAIT_MS,
        "and anything on screen is more work than nothing on screen"
    );
}

/// A static tick skips the media lock, and the four things that stop it are each a way the
/// hint can be wrong.
///
/// The backstop is the one that matters and the one that costs. It exists because a file's
/// kind can change without a swap — streaming frames land after the fact, and this app
/// discovers a stream is a video by looking at it — and without it those frames are noticed
/// on the next swap rather than within half a second: a preview that stays a spinner over a
/// stream that is already playing. It is also the arm most likely to be argued away as a
/// cost on a machine where nothing streams, which is why it is a named argument rather than
/// a bare `elapsed()` four hundred lines from where it is decided.
#[test]
fn a_static_tick_looks_at_the_media_only_when_the_hint_cannot_be_trusted() {
    let just_looked = Duration::ZERO;
    let long_untouched = Duration::from_millis(STATIC_MEDIA_REFRESH_MS);

    assert!(
        !needs_a_full_media_look(false, false, false, just_looked),
        "a static tick that just looked, with nothing new and nothing waited for, looks again \
             not at all — this is the arm that made the loop cheap enough to leave running"
    );

    for (dynamic, a_wait, generation_moved, what) in [
        (true, false, false, "the hint says something moved"),
        (
            false,
            true,
            false,
            "a load, a walk or a player is outstanding",
        ),
        (false, false, true, "a swap has installed a different hover"),
    ] {
        assert!(
            needs_a_full_media_look(dynamic, a_wait, generation_moved, just_looked),
            "{what}, so the hint is not worth trusting this tick"
        );
    }

    // And the backstop, which is the only arm that fires with nothing new at all.
    assert!(
        needs_a_full_media_look(false, false, false, long_untouched),
        "and after half a second of nothing the media is looked at anyway, so a kind that \
             changed without a swap — a stream that turned out to be a video — is noticed within \
             half a second rather than on the next hover"
    );
}

/// A pin's video keeps the place at the top and a hover's does not, because a hover's
/// competes with nothing and re-asserting topmost for a tooltip is five DWM reorders a
/// second spent on a window that was in front when nobody clicked anything.
#[test]
fn only_a_pinned_players_window_is_put_back_in_front_often() {
    assert_eq!(topmost_cadence_ms(true), PIN_TOPMOST_REASSERT_MS);
    assert_eq!(topmost_cadence_ms(false), HOVER_TOPMOST_REASSERT_MS);
    assert!(
        topmost_cadence_ms(true) < topmost_cadence_ms(false),
        "a pin competes with the pin's own window for the top and so keeps the tighter band"
    );
}

/// the floor an animation frame's own delay is lifted to, and which /// frames it is lifted for.
#[test]
fn an_animation_floor_holds_back_nothing_a_file_itself_timed() {
    // A frame's delay as the APNG decoder hands it over: milliseconds, kept as the
    // exact ratio the file wrote so that a hundredth of a second is not lost to
    // integer division before the floor ever sees it.
    let frame_of = |numerator: u32, denominator: u32| {
        image::Frame::from_parts(
            image::RgbaImage::new(1, 1),
            0,
            0,
            image::Delay::from_numer_denom_ms(numerator, denominator),
        )
    };

    assert_eq!(
        apng_frame_delay_ms(&frame_of(50, 3)),
        16,
        "a frame timed at sixty a second — sixteen and two thirds milliseconds, which is \
             what a hundredth of a second a frame is written as — is played at the file's own \
             sixty and not held to thirty"
    );
    assert_eq!(
        apng_frame_delay_ms(&frame_of(33, 1)),
        33,
        "and a file slower than the floor keeps the delay it was authored at"
    );
    assert_eq!(
        apng_frame_delay_ms(&frame_of(0, 1)),
        16,
        "while a file saying nothing at all is lifted to the floor rather than spun as \
             fast as the message pump allows"
    );
    assert_eq!(
        SPINNER_TURN_MS, 33,
        "and the two floors stay apart: an arc turns at thirty because it is a mark, and \
             a mark is not a picture to be run at a picture's speed"
    );
}
