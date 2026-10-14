//! The walk a pin takes through a folder: the next and previous file, the planner that is handed
//! one, and the failure a file that cannot be shown ends in.

use super::*;

use crate::formats::lists;
use crate::formats::routing::NavCategory;
use std::hash::{Hash, Hasher};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

/// Put a playlist in the order it is walked in: the files after the first shuffled, so the
/// permutation is drawn once — when the playlist is spawned or respawned — and every step
/// after it is the next file of the order rather than a fresh draw (see `PinStep::order`).
fn shuffle_rest(rest: &mut [PathBuf]) {
    for i in (1..rest.len()).rev() {
        rest.swap(i, roll() % (i + 1));
    }
}

/// The files of a walk's list that are of the anchor's own kind, as `claims` answers for
/// them: the whole of what a shuffled step picks from, and nothing of any other kind — a
/// shuffle is a walk through one kind of file, whichever of the three it is.
fn kind_of_the_list(list: &[PathBuf], claims: impl Fn(&Path) -> bool) -> Vec<PathBuf> {
    list.iter().filter(|path| claims(path)).cloned().collect()
}

/// A roll for the playlist: a number drawn from the clock mixed
/// with a count of how many have been drawn in this run, so that two
/// steps a moment apart do not land on the same file the way two
/// reads of a clock that barely moves would. Nothing here needs a
/// generator anyone could reproduce, and nothing is seeded from the
/// file (see `audio_seek::anywhere`, which draws its rolls the same
/// way).
fn roll() -> usize {
    static DRAWN: AtomicU64 = AtomicU64::new(0);

    let clock = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|since| since.as_nanos() as u64)
        .unwrap_or(0);
    let throw = clock
        ^ DRAWN
            .fetch_add(1, Ordering::Relaxed)
            .wrapping_mul(0x9E37_79B9_7F4A_7C15);

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    throw.hash(&mut hasher);

    hasher.finish() as usize
}

/// A step the pin's own walk has taken, and the playlist it steps through.
///
/// The walk is a playlist and a pin can only be shown one file of it at a time: a file that
/// is not there, a file nothing here can read, a file whose decode comes back with nothing at
/// all. Those are not where a walk stops. A caption button is one press and one gesture, and
/// the gesture is *the next file I can look at* — so the file that cannot be shown is stepped
/// over and the walk carries on from it, which is the one place the file the pin has to stop
/// being what the walk steps from: a file that was not shown leaves the pin showing what it
/// was showing, so the pin's own path is not where the walk is.
#[derive(Clone)]
pub(crate) struct PinStep {
    /// The file the walk last landed on, which is where the next step is taken from.
    pub(super) at: PathBuf,
    /// The file the walk was *asked* from, which is what an answer is matched against.
    ///
    /// It is not `at` and cannot become it: a walk's first step is taken on the planner's
    /// thread, so by the time the loop holds this walk, `at` is already the file the walk
    /// landed on rather than the one the press was made on. A walk that lands after the
    /// pin has been given another file is an answer to a press about a file the pin is no
    /// longer showing, and this is what says so (see `PreviewMessage::PinAnswered`).
    pub(super) from: PathBuf,
    /// Which way the walk goes: `1` along the playlist, `-1` against it.
    pub(super) step: i32,
    /// The playlist the walk steps through: the folder's own order where the pin walks
    /// linearly, or the file the pin was standing on first and the anchor's own kind of the
    /// folder shuffled after it where the pin shuffles. Spawned when the pin is triggered and
    /// reconciled on every step after it, so every file is stepped onto once before any
    /// repeats; a step back is the previous file of the order and the step forward after it
    /// the file it was interrupted on — the order itself, never a draw.
    pub(super) order: Vec<PathBuf>,
    /// How many files in a row were stepped over without being shown: a folder of files
    /// nothing can show is a walk that ends rather than one that goes round for ever, so each
    /// file is tried once and then the mark is what the pin is left standing over (see
    /// `step_over`). A fresh walk starts at none; a step is navigation rather than a miss.
    pub(super) missed: usize,
    /// Whether `order` is a shuffled playlist: the pin's own shuffle — the audio one for a
    /// walk anchored on a sound, the video one for a film, the image one for a picture — and
    /// the folder's own order everywhere else.
    pub(super) shuffle: bool,
    /// The category of the first file pinned, which is what `order` was filtered by and
    /// what the next press reuses instead of recomputing from the file the pin is
    /// standing on — or nothing where the walk was asked with no session behind it, in
    /// which case the next press anchors on the file it stands on, as a first press does.
    pub(super) anchor: Option<NavCategory>,
}

impl PinStep {
    /// The file this walk was asked from, which is what a late answer is matched against.
    pub(super) fn from(&self) -> &Path {
        &self.from
    }

    /// The file the next step of this walk lands on, and where the walk stands after it, or
    /// nothing where the walk has nowhere to step to.
    ///
    /// What a step *is* is decided here and nowhere else, because every road that asks for
    /// one is the same question: a caption's **Next**/**Previous**, a film's end of file and
    /// a sound's all ask the walk for its next file. A step is a move along the playlist: back
    /// is the previous file of the order and forward the next one — the file itself, never a
    /// draw, which is what makes **Previous** take you back exactly where you were and a
    /// **Next** after it land on the file it was interrupted on.
    pub(super) fn step(&mut self) -> Option<PathBuf> {
        self.advance(self.step)
    }

    /// The same step forward over a file that could not be shown (see `pin_step_off`):
    /// navigation moves forward only, so a walk that fell back would be walked back to where
    /// it came from. Nothing where every other file of the order has been tried, and nothing
    /// rather than the file itself where there is nowhere to step to.
    pub(super) fn step_over(&mut self) -> Option<PathBuf> {
        if self.missed >= self.order.len().saturating_sub(1) {
            return None;
        }
        self.missed += 1;
        let leaving = self.at.clone();
        let next = self.advance(1)?;
        if pin_navigation::same_file(&next, &leaving) {
            return None;
        }
        Some(next)
    }

    /// Move along the playlist without writing anything down about the file left: the half of
    /// a step both roads share — the navigation, and the step-over of a file that could not
    /// be shown (see `step_over`).
    ///
    /// A step re-anchors to the file it is standing on in the order in hand, so a press made
    /// while the file the last step landed on is still being read steps the same playlist
    /// rather than starting another. A shuffled walk stepping past its last file respawns the
    /// playlist from the order in hand — the planner respawns from the folder instead, so
    /// files added since join on the next press — and a linear one wraps, the way a folder is
    /// a ring. A file no longer on the order steps onto the end it points at.
    fn advance(&mut self, direction: i32) -> Option<PathBuf> {
        if self.order.is_empty() {
            return None;
        }
        // A shuffled walk stepping past its last file respawns the playlist rather than
        // wrapping it; a single-file shuffle has nowhere to respawn to.
        if direction >= 0
            && self.shuffle
            && self
                .order
                .last()
                .is_some_and(|last| pin_navigation::same_file(last, &self.at))
        {
            let at = self.at.clone();
            let mut rest: Vec<PathBuf> = self
                .order
                .iter()
                .filter(|file| !pin_navigation::same_file(file, &at))
                .cloned()
                .collect();
            if rest.is_empty() {
                return None;
            }
            shuffle_rest(&mut rest);
            self.order = std::iter::once(at).chain(rest).collect();
            self.at = self.order[1].clone();
            return Some(self.at.clone());
        }
        // Every other step is along the order — back the previous file, forward the next
        // one — which wraps at either end, and a file no longer on it steps onto the end it
        // points at.
        self.at = pin_navigation::step_to(&self.at, &self.order, direction)?;
        Some(self.at.clone())
    }
}

/// A walk the planner is reading a folder for, and the wait it is.
///
/// It is the pin's spinner over a wait that is not a load: no frame is being decoded, so
/// there is nothing to be in flight on the media side, but the file the pin is about to be
/// shown is not known yet and the user is owed an answer either way. Painted on the arc the
/// pin already turns, on the same delay, because it is the same question to the person
/// looking at it (see `PinLoad` and `paint_pin_spinner`).
pub(super) type PinWait = PinArc;

/// The file a step along the folder's walk takes the pin to, and the walk that step began,
/// or nothing where the walk has nowhere to step to.
///
/// The walk itself is the folder's own, and reading a folder is a felt read: a whole
/// `read_dir`, and a registry `ProgID` read for every distinct kind in it. So it is not read
/// here. This asks the planner for it and answers nothing this tick, and the wait is painted
/// as the arc the pin already knows how to paint (see `PinPlanner` and `PinWait`).
///
/// Nothing is lost by the round trip: a walk that lands is taken up by the answer, and a
/// walk that does not is a walk with nowhere to step to, which is what this already
/// answered with.
pub(super) fn step_pinned_file(
    step: i32,
    wait: &mut Option<PinWait>,
    anchor: Option<NavCategory>,
) -> Option<PinStep> {
    let at = pinned_path()?;

    if !ask_pin_walk(at, step, anchor) {
        return None;
    }

    // The wait starts when the question is asked, not when the planner gets to it, so the
    // arc is put up on the same delay a load's is however long the planner is busy.
    *wait = Some(PinWait::new());

    None
}

/// The anchor a fresh press asks with: the category of the first file pinned, read off the
/// walk the pin is standing on — or nothing where the pin holds no walk, which is a first
/// press or a press after a listing pick cleared it, and anchors on the file the pin is
/// standing on.
///
/// The walk is the pin's own and lives with the pin (see `PinUp::walk`), so there is nothing
/// for the loop to clear when a pin comes down and nothing for it to carry: a pin that is
/// over holds no walk, and the one taken up after it anchors on the file it is shown.
pub(super) fn pin_walk_anchor(walk: &Option<PinStep>) -> Option<NavCategory> {
    pin_state()
        .as_ref()
        .and_then(|state| state.walk())
        .or(walk.as_ref())
        .and_then(|walk| walk.anchor)
}

/// A question the pin asks that only a thread of its own may answer.
///
/// Each arm is work whose cost is not known in advance and not bounded by this app: a
/// folder is read from the disk, a file is opened through the system's own codecs and its
/// Shell, a registry is consulted. None of it can run on the preview thread, because that
/// thread is the one that pumps this window's messages — see `PinPlanner` for what that
/// means to a caption's buttons.
pub(super) enum PinJob {
    /// The folder's list, for a step of the pin's own walk.
    ///
    /// The configuration is carried by the walk rather than read off a global where it is
    /// wanted, and it is the one field here big enough to be worth boxing: a job slot is
    /// one of them at a time, and a walk that made every queued question as large as the
    /// configuration would be paying for the second question's sake too.
    Walk {
        at: PathBuf,
        step: i32,
        /// The category of the first file pinned, or nothing where the press had no walk
        /// behind it — a first press, or one after a listing pick cleared the walks — in
        /// which case the planner anchors on `at` itself. Carried rather than recomputed
        /// so a press standing on a sound keeps the first file's walk (see `list_for`).
        anchor: Option<NavCategory>,
        config: Box<AppConfig>,
        /// The walk this press is a step of — the pin's own playlist — so that the read the
        /// planner does behind it continues one walk rather than starting another. Nothing
        /// where the press had no walk behind it, which is the same first press and the
        /// same case `anchor` describes.
        ///
        /// Carried rather than read off the pin where it is wanted, the way the
        /// configuration is: the planner is another thread, and a walk it read off the pin's
        /// own state would be the pin's lock taken across a folder read.
        session: Option<Box<PinStep>>,
    },
    /// The name of the program this file would open with, for the hand-off button.
    OpenWith { path: PathBuf },
}

/// The job the planner is working on, and the one it will work on next, coalesced.
///
/// Newest wins, and that is the whole of the design: a user holding the next button is
/// asking for one walk, not one per file, and a queue would spend the planner reading
/// folders nobody is waiting for.
pub(super) type PinJobSlot = (Mutex<Option<PinJob>>, Condvar);
pub(super) static PIN_JOBS: Lazy<PinJobSlot> = Lazy::new(|| (Mutex::new(None), Condvar::new()));

/// The give-up on a job that has not answered.
///
/// It is the bound that turns an unbounded wait into a failed one: a folder on a network
/// share that never answers is a walk that ends rather than a window that never comes back.
/// The pin keeps the file it is showing, which is the only other answer a window that is
/// already up has (see `player_wait`, which bounds the same way for the same reason).
pub(super) const PIN_JOB_GIVEUP: Duration = Duration::from_secs(15);

/// The pin's own walk — the playlist it is standing on — whatever file it is standing on.
///
/// The walk the pin holds is its session and not one ask's (see `PinUp::walk`). It is carried
/// without matching the file the question was asked about: a step re-anchors to that file in
/// the order, so a press made while the file the last step landed on is still being read steps
/// the same playlist rather than starting another (see `advance`).
fn session_walk() -> Option<PinStep> {
    pin_state().as_ref().and_then(|state| state.walk()).cloned()
}

/// The pin's own walk, taken, and the pin left holding none: the walk the pin is standing on,
/// for the step-off that carries it past a file that could not be shown (see `pin_step_off`).
pub(super) fn take_pin_walk() -> Option<PinStep> {
    pin_state()
        .as_mut()
        .and_then(|state| state.walk_mut())
        .and_then(|held| held.take())
}

/// The playlist the walk was answered with goes onto the pin's own walk, without the pin
/// moving.
///
/// Where the order is decided is a load apart from where the file it lands on is installed:
/// a press made while that load is still reading re-anchors to the file on screen in this
/// order, so it steps the same playlist instead of starting another. Only the order goes —
/// the pin has not moved, so where it is standing is still where the next step is taken from
/// (see `PinUp::walk`).
pub(super) fn pin_walk_synced_its_playlist(walk: &PinStep) {
    if let Some(mut state) = pin_state() {
        if let Some(Some(held)) = state.walk_mut() {
            held.order = walk.order.clone();
            held.shuffle = walk.shuffle;
            held.anchor = walk.anchor;
        }
    }
}

/// Ask the planner for a walk, answering whether the question was asked at all.
///
/// The configuration is taken by copy and the lock let go before the job is queued: the
/// walk reads a folder behind it, and a lock held across a folder read is a lock every
/// other thread of the app waits on for as long as the disk takes.
///
/// The pin's own walk goes with it, for the same reason: the press is a step of the playlist
/// the pin is standing on, and the order that walk holds is what makes the next file a
/// continuation of the same playlist rather than a fresh draw (see `answer_pin_job`).
pub(super) fn ask_pin_walk(at: PathBuf, step: i32, anchor: Option<NavCategory>) -> bool {
    let Some(config) = CONFIG.lock().ok().map(|config| config.clone()) else {
        return false;
    };

    let session = session_walk().map(Box::new);

    queue_pin_job(PinJob::Walk {
        at,
        step,
        anchor,
        config: Box::new(config),
        session,
    })
}

/// Ask the planner which program this file would open with.
pub(super) fn ask_pin_open_with(path: PathBuf) {
    queue_pin_job(PinJob::OpenWith { path });
}

/// Leave a job for the planner, answering whether it was left.
///
/// A job that cannot be queued is a question that was never asked, which is a different
/// thing from one whose answer is nothing: the caller starts a wait only for the former, so
/// an arc is never turned for a walk that was not asked.
pub(super) fn queue_pin_job(job: PinJob) -> bool {
    let (lock, cvar) = &*PIN_JOBS;
    if let Ok(mut pending) = lock.lock() {
        *pending = Some(job);
        cvar.notify_one();
        return true;
    }

    false
}

/// Run the planner: one thread, one job at a time, newest first.
///
/// This is the reason a pinned window's caption keeps answering its buttons while the file
/// behind it is slow. Everything the pin wants to know that could be slow goes through here,
/// and the preview thread's tick is left holding only the work that cannot be moved — a
/// repaint, a hit-test, a key, and the engines this app's own windows are drawn by. That
/// separation is what the window needs: a caption is painted by the thread that drains its
/// messages, so anything slow run on that thread is a window whose buttons do nothing for
/// as long as it takes.
///
/// Answers come back as `PreviewMessage::PinAnswered` on the same channel every other
/// answer arrives on, and each is acted on only where it still applies: a walk is dropped
/// unless the pin is standing where the walk was asked from, and a name is dropped unless
/// the pin is standing on the file it names (see `PinStep::from`).
pub(super) fn spawn_pin_planner() {
    std::thread::spawn(|| {
        // A folder is read through the Shell's own directory entry and a file is opened
        // through the system's codecs, so this thread needs an apartment before the first
        // job asks for either (see `spawn_load_worker`, which takes the same two).
        pdf_preview::initialize_apartment();
        wic_image::initialize_apartment();

        while RUNNING.load(Ordering::Acquire) {
            let job = {
                let (lock, cvar) = &*PIN_JOBS;
                let mut pending = match lock.lock() {
                    Ok(guard) => guard,
                    Err(_) => return,
                };

                while pending.is_none() && RUNNING.load(Ordering::Acquire) {
                    pending = match cvar.wait_timeout(pending, Duration::from_millis(200)) {
                        Ok((guard, _)) => guard,
                        Err(_) => return,
                    };
                }

                if !RUNNING.load(Ordering::Acquire) {
                    return;
                }

                // The job is taken rather than left for the next pass, which is what makes
                // this one-at-a-time rather than a queue.
                pending.take()
            };

            let Some(job) = job else {
                continue;
            };

            // A job that came apart is answered as a question that had nothing to say, and
            // it *is* answered rather than dropped: `answer_pin_job` is what turns a folder
            // that could not be read into a walk that is over, so a read that unwound past
            // the take above arrives as that walk — the pin's arc comes down and it keeps
            // the file it is showing. An unwound thread past the take would instead leave
            // the question with nothing to end it at all.
            let answer =
                std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| answer_pin_job(&job)))
                    .unwrap_or_default();

            if let Some(answer) = answer {
                send_preview(PreviewMessage::PinAnswered(answer));
            }
        }
    });
}

/// Answer one job, on the planner's own thread.
///
/// A walk is *always* answered, and that is the whole of what this returns `Some` for: a
/// folder that cannot be read, and a folder with nothing in it this app can step onto, are
/// both ordinary answers rather than the absence of one. A walk answered with nothing would
/// leave the pin's arc turning until `PIN_JOB_GIVEUP` for what is really a press that had
/// nowhere to go — a one-file folder, a folder whose only walkable file is the pinned one,
/// or a share that is not answering. `left: 0` is how that says itself: the walk is spent,
/// so the loop has nothing to ask of it and the pin keeps what it is showing (see `PinStep`).
pub(super) fn answer_pin_job(job: &PinJob) -> Option<PinPlanned> {
    match job {
        PinJob::Walk {
            at,
            step,
            anchor,
            config,
            session,
        } => {
            // A folder that cannot be walked is walked as a folder with nothing in it, so
            // that the answer below is the same shape whichever of the two it was.
            let empty = Vec::new();
            // Anchored on the first file pinned: `anchor` is `None` only where the press
            // had no walk behind it, in which case the walk anchors on `at` itself. The
            // sounds below come out of this list, so they follow the same anchor.
            let anchor = anchor.or_else(|| pin_navigation::wanted_category(at, config));
            // The playlist is reconciled against the folder as it is rather than as it was
            // when the pin went up, so the cached scan is dropped before the list is read:
            // a file copied in joins the playlist and one moved or deleted leaves it.
            if let Some(folder) = at.parent() {
                pin_navigation::forget_folder(folder);
            }
            let list = pin_navigation::list_for(at, anchor, config).unwrap_or_else(|| empty.clone());

            // Whether steps along this walk follow a shuffled playlist or the folder's own
            // order: the pin's own shuffle says so — the audio one for a walk anchored on a
            // sound, the video one for a film, the image one for a picture. The pool is the
            // anchor's kind of the list, read where the list is: which files those are is a
            // question of the anchor's own list of the `config.ini` the walk was asked with,
            // so it is answered here once rather than on every step.
            let (shuffle, pool): (bool, Vec<PathBuf>) = match anchor {
                Some(NavCategory::Audio) if config.pin_mode_audio_shuffle => {
                    (true, kind_of_the_list(&list, |path| lists::AUDIO.claims(path, config)))
                }
                Some(NavCategory::Video) if config.pin_mode_video_shuffle => {
                    (true, kind_of_the_list(&list, |path| lists::VIDEO.claims(path, config)))
                }
                Some(NavCategory::Images) if config.pin_mode_image_shuffle => {
                    (true, kind_of_the_list(&list, |path| lists::IMAGE.claims(path, config)))
                }
                _ => (false, Vec::new()),
            };
            // The playlist: the folder's own order where the walk is linear, so an addition
            // lands in its sorted place and a deletion is gone with the fresh list — or the
            // walk's own order reconciled where it shuffles, so a file gone since is dropped
            // and a file added since goes on the end — or a fresh one where no walk is
            // carried: a shuffle thrown either way, a new anchor, or the trigger that took
            // the pin up at all. A playlist is spawned once per trigger and stepped along
            // after it, so every file is stepped onto once before any repeats.
            let mut order: Vec<PathBuf> = match session {
                Some(session) if session.anchor == anchor && session.shuffle == shuffle => {
                    if shuffle {
                        let mut kept: Vec<PathBuf> = session
                            .order
                            .iter()
                            .filter(|file| {
                                pin_navigation::same_file(file, at)
                                    || pool
                                        .iter()
                                        .any(|pooled| pin_navigation::same_file(pooled, file))
                            })
                            .cloned()
                            .collect();
                        let mut additions: Vec<PathBuf> = pool
                            .iter()
                            .filter(|file| {
                                !kept
                                    .iter()
                                    .any(|kept| pin_navigation::same_file(kept, file))
                            })
                            .cloned()
                            .collect();
                        kept.append(&mut additions);
                        kept
                    } else {
                        list.clone()
                    }
                }
                _ => {
                    if shuffle {
                        let mut rest: Vec<PathBuf> = pool
                            .iter()
                            .filter(|file| !pin_navigation::same_file(file, at))
                            .cloned()
                            .collect();
                        shuffle_rest(&mut rest);
                        std::iter::once(at.clone()).chain(rest).collect()
                    } else {
                        list.clone()
                    }
                }
            };
            // A shuffled walk stepping past its last file respawns the playlist from the
            // folder rather than from the order in hand, so files added since join the new
            // permutation instead of waiting for the step after it.
            if shuffle
                && *step > 0
                && order
                    .last()
                    .is_some_and(|last| pin_navigation::same_file(last, at))
            {
                let mut rest: Vec<PathBuf> = pool
                    .iter()
                    .filter(|file| !pin_navigation::same_file(file, at))
                    .cloned()
                    .collect();
                shuffle_rest(&mut rest);
                order = std::iter::once(at.clone()).chain(rest).collect();
            }
            let mut walk = PinStep {
                at: at.clone(),
                // The walk is asked from the file the pin is standing on, and steps from it
                // along the playlist.
                from: at.clone(),
                step: *step,
                order,
                missed: 0,
                shuffle,
                // Carried so the next press reuses it: the walk the pin lands on is kept
                // as the walk of the file on screen, whose anchor this is (see
                // `step_pinned_file`).
                anchor,
            };

            // A step that lands is the walk; one that does not leaves it standing where it
            // was asked from, which the loop reads as a walk with nowhere to step to.
            walk.step();

            Some(PinPlanned::Walk(walk))
        }
        PinJob::OpenWith { path } => Some(PinPlanned::OpenWith {
            path: path.clone(),
            name: default_app_name(path),
        }),
    }
}

/// An answer to a question the planner was asked, matched back to the walk or the pin
/// that is waiting for it.
///
/// It is a message body, so it is cloned with the message it travels in — which is why
/// the walk it carries is a whole list rather than a borrowed one. It is not the load's
/// answer of the same shape above: that one is a frame a pinned window is installed from,
/// and this one is what the pin asked a question of.
#[derive(Clone)]
pub(crate) enum PinPlanned {
    /// A walk, ready to be handed the file it landed on.
    Walk(PinStep),
    /// The program this file would open with, for the hand-off button's name.
    OpenWith { path: PathBuf, name: String },
}

/// Put the pin on its next file after one it could not be shown, and say where that is.
///
/// A walk is put back where the loop holds it, with the next file of the playlist standing
/// under it, and that is the whole of what makes a walk more than one file: the file it lands
/// on next is picked up on a tick of its own, and a walk dropped after the first file it could
/// not show is a walk of one — which is the navigation stopping on a corrupted file with the
/// button still pressed. The playlist wraps and a spent shuffle respawns it, so a walk put back
/// is a walk that ends only where there is nowhere to step to (see `PinStep::advance`).
///
/// Where the walk has run out — or where the failure had no walk to begin with, which is what a
/// file a pick in the listing failed at is — nothing is queued. Navigation moves forward only, and
/// that is the whole of the rule: a walk that had fallen back to the file before it would have
/// been walked back to where it came from, so a folder with one corrupted file in it could not be
/// got past by pressing **Next** at all. A file that cannot be shown is stepped over; a file that
/// was working is not something to step back onto.
///
/// Which is what leaves the placeholder: with nothing queued there is nowhere to step to, and a
/// window with nothing to show is a window that comes apart. So the file that failed is shown as
/// the cross instead — the window stays up, says so about the file it could not draw, and the
/// walk is free to start again from wherever the user presses next (see `show_pin_failure`).
pub(super) fn pin_step_off(walk: Option<PinStep>, held: &mut Option<PinStep>) {
    if let Some(mut walk) = walk {
        if walk.step_over().is_some() {
            *held = Some(walk);
        }
    }
}

/// The side a failed file is shown at, in pixels: the shape the cross is drawn in and the box a
/// pin is given for it.
///
/// It is its own size rather than the shape of the file that failed, for two reasons. The file's
/// own shape is not knowable — a corrupted container has whatever the header claims, and a header
/// that survived is exactly what a file nobody can play usually has — and a window the size of a
/// 4K frame with a cross in the middle of it is a window that has to be dragged to be read. And a
/// square is a shape the pin can be given whatever the pin's bound is, so installing it never
/// depends on measuring the file, which is the measurement that cannot be trusted here.
pub(super) const PIN_FAILURE_SIDE: u32 = 128;

/// The mark a failed file is shown as, as a frame of a preview.
///
/// It is built rather than loaded, and that is the whole of what it is for: there is no picture
/// of a file that could not be shown, so anything that came out of a decoder would be a
/// fabrication. Which also settles the cache question — a frame handed in here never goes near
/// the image cache, so the next visit to a file that has since been fixed re-renders the file
/// rather than painting the cross over a picture that is now there (see `load_static_image`).
pub(super) fn unplayable_media(size: (u32, u32)) -> MediaData {
    let (width, height) = (size.0.max(1), size.1.max(1));
    let mut pixels = vec![0u8; width as usize * height as usize * 4];

    // No palette is a mark with no ink, which is the same answer the chrome gives everywhere else
    // it is asked for one and cannot read a theme: nothing is painted, and the frame is the
    // tray's own backdrop showing through it. The bundled themes make this unreachable in
    // practice, and a wrong guess of a colour would be worse than none (see `ChromePalette`).
    if let Some(palette) = pin_chrome::ChromePalette::current() {
        pin_chrome::paint_failure_mark(&mut pixels, width, height, &palette);
    }

    static_image_media(
        Arc::new(ImageFrame::new(pixels, width, height, 0)),
        MediaType::Unplayable,
    )
}

/// The plan a pin is given the mark for a file it could not draw: the square of
/// `PIN_FAILURE_SIDE`, fitted into the pin's own room and centred on the box the pin has now.
///
/// It is the same room every other swap is fitted into and at `Percent(100)` rather than
/// `FitToScreen`, which is the one place the two differ and matters: fitted to the screen, a
/// mark this small would be blown up to the size of the display, and a cross across the whole
/// screen is not a thing anyone is meant to read. A room smaller than the mark clamps it, which
/// is the ordinary rule (see `pin_swap_room` and `pinned_media_box`).
pub(super) fn pin_failure_plan() -> Option<PinUpdate> {
    let (space, volume, collapsed) = {
        let state = pin_state()?;
        let pin = state.pin()?;
        (pin_swap_space(pin), pin.volume.level, pin.collapsed)
    };

    // A bubble has no window to show the mark in, and a pin that is collapsed is one about to
    // be restored on a file the user picked (see `pin_bubble_pick`).
    if collapsed {
        return None;
    }

    let dpi = dpi_at(space.current.0, space.current.1);
    let bounds = work_area_at(space.current.0, space.current.1);

    Some(PinUpdate {
        content: pin_update_box(
            pin_swap_room(space, bounds, dpi),
            (PIN_FAILURE_SIDE, PIN_FAILURE_SIDE),
            PreviewScale::Percent(100),
        ),
        dpi,
        volume,
    })
}

/// Hand the pin the mark for a file it could not draw, in the slot every other swap's file is
/// handed in, and say whether it was.
///
/// It is a load rather than an installation, which is the whole of the design: everything a swap
/// has to do to the window — end the player the dead file had, take a browser down, work out the
/// chrome, put the window up at a new box, write down the file on screen — is the same work for
/// a cross as for a picture, and re-doing it here would be a second copy of the take-up that
/// drifts from the first (see `take_pin_load`). The load is answered before it is handed over,
/// because there is nothing to read and nothing to decode, so it is taken up on the very tick
/// that noticed the failure rather than a tick later — which matters, because the tick that
/// notices is the tick `pin_command_request` asks on the same breath whether the pin is still
/// there, and it is asked while the media is still the dead file's own.
pub(super) fn show_pin_failure(path: &Path, pin_load: &mut Option<PinLoad>) -> bool {
    let Some(update) = pin_failure_plan() else {
        return false;
    };

    *pin_load = Some(PinLoad::answered(path, update));
    true
}

/// Whether the pin key brings a bubble back, from the two facts that make the question askable.
///
/// It never hides anything, which is the whole of the change: the key used to answer a window
/// that was up by taking it down, and that read as a pin that got in the way of a key thrown at a
/// window the user was not in — a Space in particular, which is Explorer's own and is as likely to
/// be finishing a name as anything else.
///
/// What is left is the one thing the key is still for: a bubble is a window the user put away on
/// purpose, so the key that put it away is the key that brings it back. It is answered only where
/// the keyboard is in Explorer, because behind a bubble is a listing the user is working in and the
/// file picked there is what the window should come back up on — anywhere else the key is that
/// program's (see `pin_bubble_pick`).
pub(super) fn pin_key_restores_bubble(collapsed: bool, explorer_focused: bool) -> bool {
    collapsed && explorer_focused
}

/// Hand a file to whatever the machine has filed it under - the program the user chose for
/// this kind of file, or the one Windows picked — which is the one way out of a pin into the
/// program that owns the format. A pin shows a file rather than opening it, so this is the
/// only button that gives it away.
///
/// Nothing is asked of the file here and nothing is waited for: the Shell hands the file to
/// the program and returns, and what that program does with it is the program's own business
/// and never the pin's — a pin that came back up afterwards would be a second window of a
/// file this app is already showing. The caller asks for the pin's end as it returns, so that
/// the program it starts is not opening behind a topmost window of the same file, and it asks
/// for it whatever the Shell answered — a format with nothing filed against it takes the pin
/// down too, which is the trade taken here and written down at `pinned_release`: the answer
/// says whether a program started, not whether the user still wanted the preview (see
/// `pinned_release`).
///
/// The program is opened where the pin is standing rather than where the Shell would put it
/// on its own: the pin's own window goes in as the parent and its monitor as `SEE_MASK_HMONITOR`,
/// which is the same context Explorer hands its own launches, so a preview on a second display
/// hands the file to a program on that display rather than on the primary one.
///
/// shortcut: a program that puts its window back where it was last closed ignores the hint and
/// reopens there, and only a program asking the system where to go follows it — nothing here can
/// move another program's window for it, so those keep opening where they were.
///
/// # Safety
///
/// `ShellExecuteExW` is a plain `extern "system"` call with no caller-supplied pointers kept past
/// it: the verb and the file are this function's own null-terminated buffers, and the window handle
/// handed it is the live preview window the press was released on, so there is nothing for the caller
/// to keep alive across the call and nothing it can invalidate underneath it. The file itself belongs to
/// whichever program the Shell starts, and this side has no handle on it to release.
pub(super) unsafe fn open_path_with_default_app(path: &Path, hwnd: HWND) {
    shell_open_on(hwnd, path);
}

/// Ask the Shell to open a file on the monitor the pin is standing on, through `ShellExecuteExW`
/// with the pin's own window as the parent and its monitor as `SEE_MASK_HMONITOR`.
///
/// `SEE_MASK_ASYNCOK` is what keeps this from blocking the window procedure it is pressed on:
/// the pin's end is asked for as this returns, and a call that waited would hold the loop that
/// ends it behind the program it starts.
///
/// The answer is read and deliberately not asserted on, and not acted on either: a format with
/// nothing filed against it is the user's machine rather than a fault in this one, and a check that
/// panicked on it would be an alarm that goes off for an ordinary machine (see `pinned_release`,
/// which takes the pin down whatever came back rather than leaving a topmost window over the program
/// the button exists to start).
///
/// # Safety
///
/// The verb and the file are this function's own null-terminated buffers living past the call, and
/// the window handle is the live preview window the press was released on; the monitor is a number
/// read off that window and never dereferenced. Nothing is kept past the call and nothing borrowed
/// outlives it.
unsafe fn shell_open_on(hwnd: HWND, path: &Path) {
    let file: Vec<u16> = std::ffi::OsStr::new(&plain_path(path))
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    let monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);

    let mut info = SHELLEXECUTEINFOW {
        cbSize: std::mem::size_of::<SHELLEXECUTEINFOW>() as u32,
        fMask: SEE_MASK_HMONITOR | SEE_MASK_ASYNCOK,
        hwnd,
        lpVerb: w!("open"),
        lpFile: PCWSTR(file.as_ptr()),
        nShow: SW_SHOWNORMAL.0,
        Anonymous: SHELLEXECUTEINFOW_0 {
            hMonitor: HANDLE(monitor.0),
        },
        ..Default::default()
    };

    let _ = ShellExecuteExW(&mut info);
}

/// Show the Shell's own "How do you want to open this?" dialog for a file: the list of
/// programs installed on this machine that could open it, which is the only way into a second
/// program when the default is the wrong one. The button beside this one can only ever reach
/// the program the Shell has already chosen.
///
/// It is the system's dialog rather than a list this app drew, and that is the whole of the
/// decision: the answer to "what can open this?" is the set of associations the user has
/// installed and chosen, it changes under them, and a list built here would be a copy of it
/// that is wrong the moment they install anything. Windows draws it in the user's own theme,
/// in their own language, with their own accessibility, and it is the dialog they have already
/// seen a thousand times from Explorer's own context menu.
///
/// It is raised by handing `shell32`'s `OpenAs_RunDLL` export the file, which is the route
/// Explorer's own "Open with" item takes and the one a caller without a COM apartment can
/// still reach. It is *not* the Shell's own `openas` verb: performed through `ShellExecuteExW`
/// that verb answers files with an association but comes back with "no app associated" on a
/// file that has one, the way a plain `open` answers both. The obvious alternative,
/// `SHOpenWithDialog`, cannot be used either: since Windows 11 22H2 that entry point opens
/// nothing at all and says so, telling the caller to go and change the default in Settings
/// instead of offering the list. A button whose only job is to offer the list cannot be built
/// on an entry point that has stopped offering it, so this is the other way in and not a
/// preference between two that both work.
///
/// The list is then carried onto the pin's display: once it is up it is this app's window to
/// find — a top-level window of the system's `OpenWith.exe` — and it is moved to the pin's
/// box before the pin comes down from over it (see `carry_open_with_to`).
///
/// The pin is not up while this is: the caller asks for its end as the dialog is put up, so
/// that the list is not behind a topmost window of the file it is listing ways to open, and
/// so that the program the user goes on to choose is not opening behind that window too. A
/// pin that waited for the dialog to go instead would be a pin that has to be right about
/// when the dialog has gone — a question this app cannot answer, because the dialog is not
/// its window, its class, or on any thread it owns, and on Windows 11 the Shell
/// may hand the list to a process that has already exited by the time anyone looks. Asking
/// nothing and closing is the one answer that cannot be wrong, and the cost of being wrong
/// the other way is the whole of it: a cancelled dialog costs the user a hover to get the
/// pin back.
///
/// Show the Shell's own "How do you want to open this?" dialog for a file: the list of
/// programs installed on this machine that could open it, which is the only way into a second
/// program when the default is the wrong one. The button beside this one can only ever reach
/// the program the Shell has already chosen.
///
/// # Safety
///
/// Nothing here is a borrowed pointer: the command line is a `String` this function owns,
/// handed to the standard library as one argument, and no handle is involved at all. The
/// file belongs to whichever program the user goes on to pick, and this side has no handle
/// on it to release.
pub(super) fn show_open_with_dialog(path: &Path) {
    // The one argument `rundll32` is given, in the spelling it wants.
    //
    // `rundll32` does not read its command line as a list of arguments. It takes everything up
    // to the first space as the "<library>,<entry>" pair, and then hands the function the whole
    // of the rest of the line as one opaque tail — whatever is in it, verbatim, quotes and all.
    // So the tail is passed with `raw_arg`, which puts it on the wire exactly as written and
    // skips the quoting `arg` would otherwise add to any argument holding a space.
    //
    // Two spellings that look reasonable both put a filename into that tail that cannot exist,
    // and neither fails loudly — `rundll32` simply exits at once, and a button that does that
    // is a button that does nothing. Passing the path as a second argument reaches the tail as
    // `"C:\...\my file.md"`, quotes included, because the library was already split off before
    // the path was reached; and quoting it inside a single argument nests a second pair in the
    // same place. Either way the Shell is handed a name with quote characters in it, finds no
    // such file, and shows nothing. The path goes in bare, and `plain_path` above is what
    // guarantees the only space in the tail is the one separating the entry point from it.
    let tail = format!("shell32.dll,OpenAs_RunDLL {}", plain_path(path));

    let mut command = engine_processes::hidden_command("rundll32.exe");
    command.raw_arg(&tail);

    // The child is dropped on the floor rather than kept: nothing here is waiting on it and
    // nothing comes after it, since the pin this was raised from is on its way down. Waiting
    // for it would be waiting on this thread's message loop from inside itself.
    let _ = command
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn();
}

/// The windows of the list already up, as numbers a thread can carry: only a window that appears
/// after this snapshot is this press's to move (see `carry_open_with_to`).
pub(super) fn open_with_snapshot() -> Vec<isize> {
    open_with_dialogs()
        .into_iter()
        .map(|hwnd| hwnd.0 as isize)
        .collect()
}

/// The system's "How do you want to open this?" dialogs now up, found by class rather than by
/// process or by `EnumWindows`: on this build that call does not list the dialog at all — a
/// visible, top-level, unowned window the desktop's own `FindWindowExW` names (measured) — and
/// its process is no identity either, since the id is already in a snapshot before its window
/// exists. So this is the chain off the desktop for the class itself, newest last.
fn open_with_dialogs() -> Vec<HWND> {
    let mut dialogs = Vec::new();
    unsafe {
        let mut after = HWND::default();
        loop {
            match FindWindowExW(HWND::default(), after, w!("Open With"), PCWSTR::null()) {
                Ok(found) if !found.is_invalid() => {
                    dialogs.push(found);
                    after = found;
                }
                _ => break,
            }
            if dialogs.len() > 16 {
                break;
            }
        }
    }
    dialogs
}

/// Whether a window is the list's own: the class the chain above is walked for, asked of one
/// window — the one in front — where the chain cannot say which dialog is the live one (see
/// `carry_open_with_to`).
fn is_open_with_dialog(hwnd: HWND) -> bool {
    let mut class = [0u16; 256];
    let read = unsafe { GetClassNameW(hwnd, &mut class) };
    read > 0 && String::from_utf16_lossy(&class[..read as usize]) == "Open With"
}

/// How long the carry waits for the list to put its window up, which bounds a wait on a window
/// this app does not own: the dialog usually arrives in well under a second, and a list that has
/// not appeared by then is a launch that failed rather than one still starting (see
/// `show_open_with_dialog`).
pub(super) const OPEN_WITH_WAIT: Duration = Duration::from_secs(10);

/// How long the carry waits for a window of this press's before reading the one in front
/// instead: every press raises a window of its own, so a press with nothing new after this is
/// one whose dialog the system is showing in a window it already had (see `carry_open_with_to`).
const OPEN_WITH_APPEAR_WAIT: Duration = Duration::from_secs(2);

/// How long after its last move the carry keeps watching: the dialog centres itself after it
/// comes up, so a window put onto the pin's box that drifts back where the system puts it is
/// carried again — and after this, whatever the window does is the user's (see
/// `carry_open_with_to`).
const OPEN_WITH_SETTLE: Duration = Duration::from_millis(1500);

/// How often the carry looks for the list's window while it waits for it.
pub(super) const OPEN_WITH_POLL: Duration = Duration::from_millis(25);

/// Carry the "How do you want to open this?" list onto the display the pin is standing on.
///
/// The dialog opens wherever the system puts it — on this build the primary display's centre —
/// so once it is up it is moved to the pin's box, which is the display the hand pressed on.
/// `at` is the pin's own screen box, read by the caller holding the release, because the pin
/// comes down as the dialog goes up and there is no window left to measure afterwards.
/// `before` holds the dialogs already up when this press was made, so a list left over from
/// an earlier press is not mistaken for this one's: only a window this press raised is this
/// press's to carry.
///
/// Only a window still sitting where the system put it is touched: one the user has already
/// carried somewhere is theirs, and a second press reaching for it again must not yank it
/// back. Which is what the centre says: a list the system has just put up is still centred on
/// the primary display, and a list the user has touched is not. A press with no window up yet
/// waits for it; a press whose window never comes is a launch that failed, and there is
/// nothing to carry.
///
/// The carry ends settled: a window put onto the pin's box is watched a moment longer and
/// carried again if it drifts back where the system puts it — the dialog centres itself after
/// it comes up — and after that, whatever the window does is the user's. A dialog that went
/// again — picked or cancelled faster than the carry — ends the wait rather than running it
/// out. No window of this press's after a while is the system showing the dialog in a window
/// it already had: the live dialog is the one in front, and only while it too is still where
/// the system put it.
pub(super) fn carry_open_with_to(at: ScreenRegion, before: Vec<isize>) {
    let start = Instant::now();
    let deadline = start + OPEN_WITH_WAIT;
    // Windows this carry has put onto the pin's box, watched until the settle below runs out.
    let mut watch: Vec<HWND> = Vec::new();
    let mut seen = false;
    let mut last_move: Option<Instant> = None;

    loop {
        let now = Instant::now();
        if now >= deadline {
            return;
        }

        watch.retain(|hwnd| unsafe { IsWindow(*hwnd) }.as_bool());
        let mut fresh: Vec<HWND> = Vec::new();
        for hwnd in open_with_dialogs() {
            if before.contains(&(hwnd.0 as isize)) {
                continue;
            }
            fresh.push(hwnd);
        }
        seen |= !fresh.is_empty();
        let fresh_empty = fresh.is_empty();

        // The watched ones ride along: a window put onto the pin's box that drifts back
        // where the system puts it is carried again while the settle below runs.
        let mut moving: Vec<HWND> = fresh;
        moving.extend(watch.iter().copied());
        for hwnd in moving {
            if carry_open_with_window(hwnd, at) {
                last_move = Some(now);
                if !watch.iter().any(|held| held.0 as isize == hwnd.0 as isize) {
                    watch.push(hwnd);
                }
            }
        }

        // No window of this press's after a while, and none watched: the dialog in front is
        // the live one, if it is still where the system put it.
        if fresh_empty && watch.is_empty() && now.duration_since(start) >= OPEN_WITH_APPEAR_WAIT
        {
            if let Some(hwnd) = foreground_open_with_dialog() {
                if carry_open_with_window(hwnd, at) {
                    last_move = Some(now);
                    seen = true;
                    watch.push(hwnd);
                }
            }
        }

        // Settled: something was moved and nothing has needed moving for a while, so a window
        // the user carries after this stays carried.
        if last_move.is_some_and(|moved| now.duration_since(moved) >= OPEN_WITH_SETTLE) {
            return;
        }
        // Gone again after being seen, with nothing moved: picked or cancelled before the
        // carry got to it, and there is nothing left to move.
        if seen && fresh_empty && last_move.is_none() {
            return;
        }

        std::thread::sleep(OPEN_WITH_POLL);
    }
}

/// The live dialog where the chain cannot say which one it is: the window in front, and only
/// while it is the list's own — a press from another program has no pin box to carry anything
/// onto, and this carry only ever runs for one of this app's presses (see `carry_open_with_to`).
fn foreground_open_with_dialog() -> Option<HWND> {
    // Safety: the handle is read off the live desktop and only read from; a window gone since
    // is refused by the calls made on it rather than acted on.
    let foreground = unsafe { GetForegroundWindow() };
    if foreground.is_invalid() || !is_open_with_dialog(foreground) {
        return None;
    }
    Some(foreground)
}

/// Move one of the list's windows onto the pin's box, answering whether it was carried.
///
/// A window is only moved while its centre is still on the primary display: that is where the
/// system centres the dialog, so a window whose middle is there is one nobody has touched yet,
/// and a window the user has carried elsewhere is left alone (see `carry_open_with_to`). The
/// move keeps the window's own size and z-order and never activates it — the pin is coming
/// down and the dialog is the hand's next stop, but stealing the keyboard to say so would
/// take it from wherever the user has since put it.
fn carry_open_with_window(hwnd: HWND, at: ScreenRegion) -> bool {
    let mut rect = RECT::default();
    // Safety: the handle was just named by the desktop's own chain, and only its box is read;
    // a window gone since is refused rather than acted on.
    if unsafe { GetWindowRect(hwnd, &mut rect) }.is_err() {
        return false;
    }
    if rect.right <= rect.left || rect.bottom <= rect.top {
        return false;
    }
    let middle = (
        rect.left.saturating_add(rect.right) / 2,
        rect.top.saturating_add(rect.bottom) / 2,
    );
    if !on_primary_display(middle.0, middle.1) {
        return false;
    }

    let (width, height) = (rect.right - rect.left, rect.bottom - rect.top);
    let (left, top) = open_with_spot(at, (width, height));
    // Safety: the handle is the dialog's own live window and only its position is
    // touched — size, order and activation all left as they are (see above).
    unsafe {
        let _ = SetWindowPos(
            hwnd,
            HWND(std::ptr::null_mut()),
            left,
            top,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
        );
    }
    true
}

/// Where the list goes on the pin's display: centred on the pin's own box, clamped onto the
/// display the pin is standing on so a pin half off a display's edge still leaves a list that
/// is wholly on it.
fn open_with_spot(at: ScreenRegion, size: (i32, i32)) -> (i32, i32) {
    let middle = ((at.0.saturating_add(at.2)) / 2, (at.1.saturating_add(at.3)) / 2);
    open_with_spot_in(work_area_at(middle.0, middle.1), middle, size)
}

/// The arithmetic of `open_with_spot` over an explicit room, which is what the test below holds:
/// the middle centred and then clamped, so a box smaller than the room is centred and a box
/// the room cannot hold at full size is pushed wholly onto it rather than hung off its edge.
pub(super) fn open_with_spot_in(bounds: ScreenBounds, middle: (i32, i32), size: (i32, i32)) -> (i32, i32) {
    let left = middle.0.saturating_sub(size.0 / 2).clamp(bounds.left, bounds.right.saturating_sub(size.0).max(bounds.left));
    let top = middle.1.saturating_sub(size.1 / 2).clamp(bounds.top, bounds.bottom.saturating_sub(size.1).max(bounds.top));
    (left, top)
}

/// Whether a point is on the primary display, whose room is the fallback `primary` names: the
/// system centres the list there, so a list whose middle answers yes is one nobody has moved
/// yet (see `carry_open_with_window`).
fn on_primary_display(x: i32, y: i32) -> bool {
    centre_on_display((x, y), DESKTOPS.primary())
}

/// Whether a point is inside a display's room, which is what `on_primary_display` asks of the
/// primary one — and what the test below holds without a screen attached.
pub(super) fn centre_on_display(point: (i32, i32), bounds: ScreenBounds) -> bool {
    point.0 >= bounds.left && point.0 < bounds.right && point.1 >= bounds.top && point.1 < bounds.bottom
}


