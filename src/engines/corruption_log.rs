//! What a player's own log lines say about a file that is damaged, and the small machine that
//! decides when one has said it.
//!
//! **Nothing here reads a file, opens a pipe or draws anything.** A player's log lines are matched
//! against a fixed table of lowercased substrings (see `SIGNATURES`), a per-file latch is tripped
//! by the ones that are evidence (see `CorruptionWatch`), and a pinned window draws a card once for
//! the file whose latch is tripped (see `CorruptionCard`). The pipe those lines are drained from is
//! `ui::preview_window::player_logs`; the card itself is `ui::preview_window::paint`.
//!
//! **Why a table of substrings and not a parser.** The lines are two programs' own prose, they
//! change between builds, and this is an informational overlay: what it may not do is guess from
//! anything else. It is a list read once per line, case-insensitively and without an allocation —
//! not a regular expression and not a scan of the whole file.
//!
//! Every signature below is a spelling measured out of `mpv` 0.41 on Windows, with the file that
//! earned it in the comment beside it (see the notes on the table). A line is matched, never
//! printed, which matters: several of these carry the file's own path.

use std::sync::atomic::{AtomicBool, AtomicU64, AtomicU8, Ordering};
use std::sync::{Arc, LazyLock};
use std::time::{Duration, Instant};

/// Whether a line is evidence enough that the file is damaged on its own, or only in company.
///
/// The split is what keeps the card from being raised by a single frame: a container that will not
/// open or a codec that will not start is a broken file, and one mangled frame is a film playing.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Weight {
    /// Damaged in a way no healthy file shows: it trips the latch alone.
    Hard,
    /// A hiccup a healthy file can also have: it takes `SOFTS_BEFORE_TRIP` of them.
    Soft,
}

/// How many `Soft` lines, of any mix, are enough to say the file is damaged.
pub(crate) const SOFTS_BEFORE_TRIP: u8 = 2;

/// The signatures, lowercased, and which of the two weights each carries.
///
/// The `Hard` half is a container that will not open or a stream that will not parse — the answers
/// mpv gave a file of random bytes (`Failed to recognize file format.`), a truncated one
/// (`[mkv] Unexpected end of file - partial or corrupt file?`, `[mkv] goes past end of containing
/// element`, `[mkv] Error parsing element Attachments`) and a missing one (`[file] Cannot open file
/// '…': Invalid argument`) on this machine. The `Soft` half is the per-frame prose: a film whose
/// clusters were scribbled over printed `[ffmpeg/video] h264: error while decoding MB 53 19,
/// bytestream 20509` — one per mangled frame, which is why one of them is not enough.
const SIGNATURES: &[(&str, Weight)] = &[
    // Hard: the file is not a container this player can open, or its stream cannot be parsed.
    ("failed to recognize file format", Weight::Hard),
    ("cannot open file", Weight::Hard),
    ("failed to open", Weight::Hard),
    ("could not open codec", Weight::Hard),
    ("invalid data found when processing input", Weight::Hard),
    ("corrupt file detected", Weight::Hard),
    ("invalid ebml", Weight::Hard),
    ("ebml stream parsing failed", Weight::Hard),
    ("unexpected end of file", Weight::Hard),
    ("goes past end of containing element", Weight::Hard),
    ("error parsing element", Weight::Hard),
    // Soft: a frame, a packet or a stamp the player recovered from.
    ("error while decoding", Weight::Soft),
    ("corrupt decoded frame", Weight::Soft),
    ("invalid nal unit size", Weight::Soft),
    ("non-monotonous dts", Weight::Soft),
    ("non-monotonic dts", Weight::Soft),
    ("invalid video timestamp", Weight::Soft),
    ("invalid video pts", Weight::Soft),
];

/// The weight `line` carries, or nothing where it says nothing about the file's health.
///
/// A lowercased substring match, without the lowercased copy: the line is compared where it lies,
/// so a log that arrives at a few hundred bytes a second costs no allocation at all.
pub(crate) fn weight_of(line: &str) -> Option<Weight> {
    let haystack = line.as_bytes();

    SIGNATURES
        .iter()
        .find_map(|(signature, weight)| holds(haystack, signature.as_bytes()).then_some(*weight))
}

/// Whether `haystack` contains `needle`, ignoring ASCII case.
fn holds(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() || haystack.len() < needle.len() {
        return false;
    }

    haystack
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle))
}

/// The clock the latch and the card are both measured on: one base for the process, so a millisecond
/// written by a drain thread is the same millisecond the paint reads.
static CLOCK: LazyLock<Instant> = LazyLock::new(Instant::now);

/// Milliseconds since the process's own base, which is the unit both `CorruptionWatch` and
/// `CorruptionCard` store.
fn clock_ms() -> u64 {
    CLOCK.elapsed().as_millis() as u64
}

/// One file's latch, written by the thread draining its player's log and read by the paint.
///
/// **It is monotonic by construction**: there is no way to clear it, so a file that has tripped
/// stays tripped, and the answer cannot flicker between two ticks. What makes a *different* file
/// unaccused is that it is a different watch — the pin is taken up holding the watch of the player
/// it was taken up on (see `PinnedPreview::corruption` and `player_logs::watch_of`), so there is no
/// reset anywhere to forget.
pub(crate) struct CorruptionWatch {
    /// Milliseconds on `CLOCK` at which it tripped, plus one, or zero while it has not.
    tripped_at_ms: AtomicU64,
    /// How many `Soft` lines have been seen, saturating at `SOFTS_BEFORE_TRIP`.
    softs: AtomicU8,
}

impl CorruptionWatch {
    pub(crate) fn new() -> Self {
        Self {
            tripped_at_ms: AtomicU64::new(0),
            softs: AtomicU8::new(0),
        }
    }

    /// Read one log line, and answer what it was read as.
    ///
    /// A `Hard` line trips the latch at once; a `Soft` one trips it when its kind has been seen
    /// `SOFTS_BEFORE_TRIP` times, whichever soft lines they were. A line saying nothing about the
    /// file is answered with nothing and changes no state.
    pub(crate) fn observe(&self, line: &str) -> Option<Weight> {
        let weight = weight_of(line)?;

        match weight {
            Weight::Hard => self.trip(),
            Weight::Soft => {
                let seen = self.softs.fetch_add(1, Ordering::AcqRel).saturating_add(1);
                if seen >= SOFTS_BEFORE_TRIP {
                    self.trip();
                }
            }
        }

        Some(weight)
    }

    fn trip(&self) {
        // Written once and never cleared; the first line to trip it is the moment the card is
        // measured from, so a later one does not move it.
        let _ = self.tripped_at_ms.compare_exchange(
            0,
            clock_ms().saturating_add(1),
            Ordering::AcqRel,
            Ordering::Acquire,
        );
    }

    /// Whether this file has been accused, and it stays so.
    pub(crate) fn tripped(&self) -> bool {
        self.tripped_at_ms.load(Ordering::Acquire) != 0
    }
}

/// How long the card takes to fade in, how long it stands, and how long it takes to go.
pub(crate) const CARD_FADE_IN: Duration = Duration::from_millis(150);
pub(crate) const CARD_HOLD: Duration = Duration::from_millis(3_000);
pub(crate) const CARD_FADE_OUT: Duration = Duration::from_millis(500);

/// The whole of the card's one run, from the first pixel of its fade to the last.
pub(crate) const CARD_TOTAL: Duration = Duration::from_millis(3_650);

/// How much of the card is drawn this far into its run: from zero at the first frame, up over
/// `CARD_FADE_IN`, whole for `CARD_HOLD`, and away again over `CARD_FADE_OUT`.
///
/// A function rather than a field so the curve can be asked without a window, a clock or a pin: the
/// painting, the tick that decides whether another frame is owed and the tests all read this one
/// answer.
pub(crate) fn card_coverage(elapsed: Duration) -> f32 {
    if elapsed >= CARD_TOTAL {
        return 0.0;
    }

    if elapsed < CARD_FADE_IN {
        return elapsed.as_secs_f32() / CARD_FADE_IN.as_secs_f32();
    }

    let held = CARD_FADE_IN + CARD_HOLD;
    if elapsed < held {
        return 1.0;
    }

    let out = elapsed - held;
    (1.0 - out.as_secs_f32() / CARD_FADE_OUT.as_secs_f32()).clamp(0.0, 1.0)
}

/// The card one pinned file may show: the watch of the player showing *that* file, and the clock of
/// the one run the card has.
///
/// **One per file, made with the pin.** A pin is rebuilt whole for every file it is shown (see
/// `pin_install::take_up_pinned_window`), so this value is too, and a step to another file is a card
/// that has never been drawn rather than one that has to be put back. Nothing resets it, and a card
/// of a file the pin has left is dropped with the pin that held it.
pub(crate) struct CorruptionCard {
    /// The latch of the player showing this file, or nothing where no player of this app's began
    /// one for it — a file shown by the media engine, or one nothing of its output was read from.
    watch: Option<Arc<CorruptionWatch>>,
    /// Milliseconds on `CLOCK` at which the card was first painted, or zero while it has not been.
    started_ms: AtomicU64,
    /// Whether its one run is over. A card is drawn once per file and never again.
    done: AtomicBool,
}

impl CorruptionCard {
    /// The card of a file nothing is watching, which is never drawn.
    #[cfg(test)]
    pub(crate) fn none() -> Self {
        Self::watching(None)
    }

    /// The card of a file whose player's log is being drained into `watch`.
    pub(crate) fn watching(watch: Option<Arc<CorruptionWatch>>) -> Self {
        Self {
            watch,
            started_ms: AtomicU64::new(0),
            done: AtomicBool::new(false),
        }
    }

    /// The latch this card is drawn for, where there is one: what a test reads to know which
    /// player's log a pin was taken up against.
    #[cfg(test)]
    pub(crate) fn watch(&self) -> Option<&Arc<CorruptionWatch>> {
        self.watch.as_ref()
    }

    /// Whether the file this card belongs to has been accused, which is not yet a card on screen.
    pub(crate) fn tripped(&self) -> bool {
        self.watch.as_ref().is_some_and(|watch| watch.tripped())
    }

    /// How much of the card is on screen now, or nothing where it is not to be drawn — the player
    /// has not accused the file, or the card's one run is over.
    ///
    /// It is the paint's own question and it starts the clock: the run is timed from the first frame
    /// that has something to draw, so a file tripped before its pin was taken up gets the whole of
    /// the card rather than the tail of a run that began on a hover.
    pub(crate) fn coverage(&self) -> Option<f32> {
        self.coverage_at(clock_ms())
    }

    /// `coverage` at a given millisecond on the process's clock, which is what a test drives.
    pub(crate) fn coverage_at(&self, now_ms: u64) -> Option<f32> {
        if !self.tripped() || self.done.load(Ordering::Acquire) {
            return None;
        }

        let started = match self.started_ms.load(Ordering::Acquire) {
            0 => {
                let started = now_ms.max(1);
                let _ = self.started_ms.compare_exchange(
                    0,
                    started,
                    Ordering::AcqRel,
                    Ordering::Acquire,
                );
                started
            }
            started => started,
        };

        let elapsed = Duration::from_millis(now_ms.saturating_sub(started));
        if elapsed >= CARD_TOTAL {
            self.done.store(true, Ordering::Release);
            return None;
        }

        Some(card_coverage(elapsed))
    }

    /// Whether another frame is owed to the card: the one question the loop's tick asks.
    ///
    /// It goes false by itself when the run is over, which is what keeps the card from being a tick
    /// that never stops (see `event_loop::run_preview_window`).
    pub(crate) fn is_animating(&self) -> bool {
        self.coverage().is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A watch that has been handed one line.
    fn a_watch_that_saw(line: &str) -> CorruptionWatch {
        let watch = CorruptionWatch::new();
        watch.observe(line);
        watch
    }

    #[test]
    fn one_hard_line_trips_the_latch_on_its_own() {
        for line in [
            "Failed to recognize file format.",
            "[file] Cannot open file 'C:\\films\\a.mkv': Invalid argument",
            "[ffmpeg] Could not open codec",
            "[mkv] Corrupt file detected. Trying to resync starting from position 20511866...",
            "[mkv] Invalid EBML length at position 20511866",
            "[ffmpeg] Invalid data found when processing input",
        ] {
            let watch = a_watch_that_saw(line);
            assert!(watch.tripped(), "{line:?} is a hard line");
        }
    }

    #[test]
    fn one_soft_line_does_not_trip_the_latch() {
        let watch = a_watch_that_saw(
            "[ffmpeg/video] h264: error while decoding MB 53 19, bytestream 20509",
        );
        assert!(!watch.tripped());
    }

    #[test]
    fn two_soft_lines_trip_the_latch_whichever_two_they_are() {
        let repeated = CorruptionWatch::new();
        repeated.observe("error while decoding MB 53 19");
        repeated.observe("error while decoding MB 61 22");
        assert!(repeated.tripped());

        // Two of different kinds: the count is of soft lines, not of one spelling of them.
        let mixed = CorruptionWatch::new();
        mixed.observe("Corrupt decoded frame in stream 0");
        assert!(!mixed.tripped());
        mixed.observe("Invalid video timestamp");
        assert!(mixed.tripped());
    }

    #[test]
    fn a_line_about_something_else_never_trips_the_latch() {
        for line in [
            "",
            "[cplayer] Starting playback...",
            "[vd] Using software decoding.",
            "Audio/Video desynchronisation detected! Possible reasons include too slow",
            "[af] [convert] 44100Hz stereo 2ch floatp",
            "[sub/ass] Setting up fonts...",
            // A path is not evidence about the file it names, however it is spelled.
            "[file] Open C:\\films\\a corrupt film.mkv",
        ] {
            let watch = CorruptionWatch::new();
            assert_eq!(
                watch.observe(line),
                None,
                "{line:?} says nothing about health"
            );
            assert!(!watch.tripped());
        }
    }

    #[test]
    fn a_line_is_matched_whatever_its_case_and_without_a_lowercased_copy() {
        assert_eq!(
            weight_of("[FFmpeg/Video] H264: ERROR WHILE DECODING MB 1 2"),
            Some(Weight::Soft)
        );
        assert_eq!(
            weight_of("[mkv] unexpected END of file"),
            Some(Weight::Hard)
        );
    }

    #[test]
    fn a_tripped_latch_stays_tripped_however_many_lines_come_after_it() {
        let watch = CorruptionWatch::new();
        assert!(!watch.tripped());

        watch.observe("Failed to recognize file format.");
        assert!(watch.tripped());

        // Several more lines, hard and soft, and it is the same one latch: nothing clears it.
        watch.observe("error while decoding MB 1 1");
        watch.observe("Corrupt file detected. Trying to resync");
        assert!(watch.tripped());
    }

    #[test]
    fn the_card_fades_in_holds_and_fades_away_once() {
        // The curve itself, asked of whole milliseconds, with the three bands named rather than
        // worked out: a change to any of them is a change to what the user sees.
        assert_eq!(card_coverage(Duration::ZERO), 0.0);
        assert!((card_coverage(Duration::from_millis(75)) - 0.5).abs() < 0.01);
        assert_eq!(card_coverage(CARD_FADE_IN), 1.0);
        assert_eq!(card_coverage(CARD_FADE_IN + CARD_HOLD), 1.0);
        assert!(
            (card_coverage(CARD_FADE_IN + CARD_HOLD + Duration::from_millis(250)) - 0.5).abs()
                < 0.01
        );
        assert_eq!(card_coverage(CARD_TOTAL), 0.0);
        assert_eq!(card_coverage(CARD_TOTAL + Duration::from_secs(1)), 0.0);
    }

    #[test]
    fn nothing_is_drawn_for_a_file_no_player_has_accused() {
        let card = CorruptionCard::none();
        assert!(!card.tripped());
        assert_eq!(card.coverage_at(1_000), None);
        assert!(!card.is_animating());
    }

    #[test]
    fn the_card_is_drawn_for_one_run_and_then_never_again() {
        let watch = Arc::new(a_watch_that_saw("Failed to recognize file format."));
        let card = CorruptionCard::watching(Some(Arc::clone(&watch)));

        // The first paint starts the clock, and the run begins at nothing.
        assert_eq!(card.coverage_at(10_000), Some(0.0));
        assert!((card.coverage_at(10_075).unwrap_or_default() - 0.5).abs() < 0.01);
        assert_eq!(card.coverage_at(10_200), Some(1.0));

        // Over the whole run it is finished, and stays finished however long it is looked at.
        let end = 10_000 + CARD_TOTAL.as_millis() as u64;
        assert_eq!(card.coverage_at(end), None);
        assert_eq!(card.coverage_at(end + 60_000), None);
        assert!(!card.is_animating());
    }

    #[test]
    fn a_file_accused_before_its_pin_was_taken_up_gets_the_whole_of_the_card() {
        let watch = Arc::new(a_watch_that_saw("Failed to recognize file format."));
        // The line arrived long before the pin did; the run is timed from the first paint.
        let card = CorruptionCard::watching(Some(watch));
        assert_eq!(card.coverage_at(900_000), Some(0.0));
        assert!((card.coverage_at(900_075).unwrap_or_default() - 0.5).abs() < 0.01);
    }
}
