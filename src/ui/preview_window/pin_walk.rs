//! The walk a pin takes through a folder: the next and previous file, the planner that is handed
//! one, and the failure a file that cannot be shown ends in.

use super::*;

/// Every file of a list but the one the pin is showing, which is how many files a walk of that
/// list may be asked for.
///
/// The file the pin is on is not one of them: the walk stepped from it, so it is where the walk
/// came rather than a file to be shown, and offering it again would be a step onto the file
/// already on screen. This is the whole of what bounds a walk, so that a folder of nothing this
/// app can read is a walk that ends rather than one that goes for ever (see `PinStep`).
pub(super) fn walk_budget(list: &[PathBuf]) -> usize {
    list.len().saturating_sub(1)
}

/// A step the pin's own walk has taken, and what is left of the walk to ask for.
///
/// The walk is a list of files and a pin can only be shown some of them: a file that is not
/// there, a file nothing here can read, a file whose decode comes back with nothing at all.
/// Those are not where a walk stops. A caption button is one press and one gesture, and the
/// gesture is *the next file I can look at* — so the file that cannot be shown is stepped over
/// and the walk carries on from it, which is the one place the file the pin has to stop
/// being what the walk steps from: a file that was not shown leaves the pin showing what it
/// was showing, so the pin's own path is not where the walk is.
///
/// Every other file of the list is offered at most once, so a folder of nothing this app can
/// read is a walk that ends rather than one that goes for ever (see `walk_budget`).
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
    /// Which way the walk goes: `1` along the listing, `-1` against it.
    pub(super) step: i32,
    /// How many more files the walk may be asked for (see `walk_budget`).
    pub(super) left: usize,
    /// The folder's list, read once by the planner and carried here so that a step after
    /// the first is a walk through what is already in hand rather than a second read of
    /// the folder (see `PinPlanner`).
    pub(super) list: Vec<PathBuf>,
}

impl PinStep {
    /// The file this walk was asked from, which is what a late answer is matched against.
    pub(super) fn from(&self) -> &Path {
        &self.from
    }

    /// The file the next step of this walk lands on, and where the walk stands after it, or
    /// nothing where the walk has nothing left to offer.
    pub(super) fn step(&mut self) -> Option<PathBuf> {
        if self.left == 0 {
            return None;
        }
        self.left -= 1;

        // The list is already in hand — the planner read it off the preview thread — so a
        // step costs a position in a vector rather than a walk of the folder behind it.
        let path = pin_navigation::step_to(&self.at, &self.list, self.step)?;

        self.at = path.clone();

        Some(path)
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
pub(super) fn step_pinned_file(step: i32, wait: &mut Option<PinWait>) -> Option<PinStep> {
    let at = pinned_path()?;

    if !ask_pin_walk(at, step) {
        return None;
    }

    // The wait starts when the question is asked, not when the planner gets to it, so the
    // arc is put up on the same delay a load's is however long the planner is busy.
    *wait = Some(PinWait::new());

    None
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
        config: Box<AppConfig>,
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

/// Ask the planner for a walk, answering whether the question was asked at all.
///
/// The configuration is taken by copy and the lock let go before the job is queued: the
/// walk reads a folder behind it, and a lock held across a folder read is a lock every
/// other thread of the app waits on for as long as the disk takes.
pub(super) fn ask_pin_walk(at: PathBuf, step: i32) -> bool {
    let Some(config) = CONFIG.lock().ok().map(|config| config.clone()) else {
        return false;
    };

    queue_pin_job(PinJob::Walk {
        at,
        step,
        config: Box::new(config),
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
        PinJob::Walk { at, step, config } => {
            // A folder that cannot be walked is walked as a folder with nothing in it, so
            // that the answer below is the same shape whichever of the two it was.
            let empty = Vec::new();
            let list = pin_navigation::list_for(at, config).unwrap_or_else(|| empty.clone());

            // The walk is the list asked one file at a time, so it is bounded by the list:
            // every file of it but the one the pin is showing is one the walk may still be
            // offered, and this step is the first of them.
            let left = walk_budget(&list);
            let mut walk = PinStep {
                at: at.clone(),
                // The walk is asked from the file the pin is standing on, and steps from it
                // onto the first of the list that is not that file.
                from: at.clone(),
                step: *step,
                left,
                list,
            };

            // A step that lands is the walk; one that does not leaves it with nothing left
            // to offer, which the loop reads as a walk that is over rather than one that is
            // still out.
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
/// A walk is put back where the loop holds it, with the next file of the list standing under it,
/// and that is the whole of what makes a walk more than one file: the file it lands on next is
/// picked up on a tick of its own, and a walk dropped after the first file it could not show is a
/// walk of one — which is the navigation stopping on a corrupted file with the button still
/// pressed. The list wraps (see `pin_navigation::step_to`) and the walk's budget bounds it, so a
/// walk put back is a walk that ends rather than one that goes round for ever (see
/// `walk_budget`).
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
        if walk.step().is_some() {
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
/// # Safety
///
/// `ShellExecuteW` is a plain `extern "system"` call with no caller-supplied pointers: the
/// verb and the file are this function's own null-terminated buffers, and the window handle
/// handed it is null rather than borrowed, so there is nothing for the caller to keep alive
/// across the call and nothing it can invalidate underneath it. The file itself belongs to
/// whichever program the Shell starts, and this side has no handle on it to release.
pub(super) unsafe fn open_path_with_default_app(path: &Path) {
    let wide: Vec<u16> = std::ffi::OsStr::new(&plain_path(path))
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    let started = ShellExecuteW(
        HWND(std::ptr::null_mut()),
        w!("open"),
        PCWSTR(wide.as_ptr()),
        PCWSTR::null(),
        PCWSTR::null(),
        SW_SHOWNORMAL,
    );

    // The handle is read and deliberately not asserted on, and no longer acted on either. It
    // is the documented test of whether the Shell started something — a handle above 32, and
    // one of its own error codes at or below 32 for anything it would not start — and which
    // of those it was is not asked, because a format with nothing filed against it is the
    // user's machine rather than a fault in this one, and a check that panicked on it would
    // be an alarm that goes off for an ordinary machine. It used to be acted on: the pin came
    // down only on a real launch, so a hand-off that never happened left the window standing.
    // It does not any more, because the answer says whether a program started and not whether
    // the user still wanted the preview, and the mistake of leaving a topmost window over the
    // program a button exists to start is the one worth not making (see `pinned_release`).
    let _ = started;
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
/// still reach. The obvious alternative, `SHOpenWithDialog`, cannot be used: since Windows 11
/// 22H2 that entry point opens nothing at all and says so, telling the caller to go and change
/// the default in Settings instead of offering the list. A button whose only job is to offer
/// the list cannot be built on an entry point that has stopped offering it, so this is the
/// other way in and not a preference between two that both work.
///
/// The pin is not up while this is: the caller asks for its end as the dialog is put up, so
/// that the list is not behind a topmost window of the file it is listing ways to open, and
/// so that the program the user goes on to choose is not opening behind that window too. A
/// pin that waited for the dialog to go instead would be a pin that has to be right about
/// when the dialog has gone — a question this app cannot answer, because the dialog is not
/// its window, its class, or on any thread it owns, and on Windows 11 the entry point above
/// may hand the list to a process that has already exited by the time anyone looks. Asking
/// nothing and closing is the one answer that cannot be wrong, and the cost of being wrong
/// the other way is the whole of it: a cancelled dialog costs the user a hover to get the
/// pin back.
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
