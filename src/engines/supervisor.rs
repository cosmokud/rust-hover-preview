//! The worker an out-of-process adapter is run on, and the record of what is on it.
//!
//! Four of the engines here are converters rather than applications: `imagemagick_render`,
//! `peazip_render`, `calibre_render` and `libreoffice_render` each start a program, wait for it and
//! read what it wrote. Each of them needs a thread of its own, because the caller is the preview
//! loop and a loop held inside a launch is a hover that does not come up, a tray that does not
//! answer and a pointer that cannot leave the file it is on. Each of them needs to answer at once
//! rather than wait, and that is what a one-slot queue is: a request put down by a hover, taken up
//! by the thread, and answered where the hover is already watching for it — the folder a page
//! lands in, or a message on the preview channel.
//!
//! Each of them also needs a bound on how long one run may take, because a file an engine cannot
//! read does not fail, it spins: a Flash file, measured, one core at a hundred percent and no page
//! past every bound. The bound is a *stop* and not a trim (see the glossary's **Give-up**), and
//! what it ends is the process, through `engine_processes` and by nothing else.
//!
//! Having been four of everything is how this came to exist. The newest-wins rule, the bounded
//! wait, the record of what was in flight and the `Running { source, pid, started }` behind it were
//! each written four times, and a change to any of them was four changes with the fourth the one
//! that got missed. What is here is the one of each, keyed by which adapter is asking.
//!
//! What is *not* here is anything the four do not share. A request's own shape is that adapter's
//! — a path, a path and the box to develop into, a path and the hover that asked for it — so a
//! request is handed down as the work to be done rather than as a payload the supervisor would
//! have to know the shape of. The give-up is each adapter's own number, for the reason each of them
//! gives. And `office_render` is not one of the four and cannot be made one: it is COM inside this
//! process, its worker is signalled with a window message rather than a condvar, and a worker
//! inside a COM call cannot be ended by anything a caller can reach — it is abandoned, and the
//! thread is left where it is. Forcing it in here would mean pretending a thread is a process (see
//! ADR 5).
//!
//! The seat is not shared either, and it is worth saying why, because one worker for all four is
//! the obvious next step and is wrong: a run that has outrun its bound would be four engines'
//! business rather than one, so the Flash file above is a thirty-second stall of the image
//! converter. Each adapter therefore has a `Worker` of its own and this module owns the kind
//! rather than the instance.

use crate::app::engine_processes;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::{Child, ExitStatus};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Condvar, Mutex};
use std::time::{Duration, Instant};

/// The four out-of-process adapters, named so that a request cannot be put down for one and
/// answered for another. `office_render` is not among them and is not named here, for the reason
/// the module docs give.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum Adapter {
    LibreOffice,
    ImageMagick,
    PeaZip,
    Calibre,
}

/// How long a worker waits on its own slot before looking again.
///
/// The wait is on the slot itself, so a request is taken up the moment it is made and this is only
/// a ceiling on how long anything else goes unnoticed: a run asked to be given up on, a kept
/// engine whose idle time has run out, a run that is ending.
const IDLE_TICK: Duration = Duration::from_secs(1);

/// One piece of work for a worker: whatever the adapter's own `request` put down, and nothing more,
/// because the shape of a request is that adapter's own (see the module docs).
type Work = Box<dyn FnOnce() + Send + 'static>;

/// One adapter's worker: the work waiting to be done, the signal that there is some, and whether
/// the thread that does it has been begun.
///
/// One per adapter rather than one for the four, for the reason the module docs give.
pub struct Worker {
    slot: Mutex<Option<Work>>,
    ready: Condvar,
    started: AtomicBool,
    /// What the thread does when its slot is empty and there is nothing to wait for. `None` for
    /// the three that have nothing between runs; the engine that is kept between documents has its
    /// kept engine let go of here rather than by anybody else, so a thread that is inside a
    /// conversion is a thread that is not looking at the idle time.
    ///
    /// It runs where the slot is found empty and not where a wait runs out, which is what keeps it
    /// off a thread a hover is waiting on: a work in hand is a hover waiting on it.
    on_idle: Option<fn()>,
}

impl Worker {
    /// A worker for one adapter, with what that adapter has to do between its runs.
    pub fn new(on_idle: Option<fn()>) -> Self {
        Self {
            slot: Mutex::new(None),
            ready: Condvar::new(),
            started: AtomicBool::new(false),
            on_idle,
        }
    }

    /// Put one piece of work down, replacing whatever was waiting, and make sure the thread that
    /// does the work is up.
    ///
    /// Replacing rather than queueing is the point of a slot and not a queue: the pointer is on one
    /// file, so the newest hover is the one worth answering and a file whose hover has gone is one
    /// nobody is waiting for. That is the loader's own slot and the same rule.
    pub fn request(&'static self, work: impl FnOnce() + Send + 'static) {
        if let Ok(mut slot) = self.slot.lock() {
            *slot = Some(Box::new(work));
        }

        self.wake();
    }

    /// Make sure the thread is up and looking, without putting anything down for it: what a
    /// request that is a flag rather than a piece of work needs, and what a request put down after
    /// the thread is already up needs.
    ///
    /// It is one of the app's threads rather than one per file: what it does between runs is wait
    /// on its own slot, which costs nothing, and what it is asked for is one file at a time
    /// because what it is holding is one process. The signal goes first and the thread second, as
    /// it did in each adapter this came out of — a thread that is already waiting has to be woken
    /// by it, and one that is not yet waiting looks at its own slot and its adapter's flag before
    /// it ever waits, so nothing is lost either way.
    pub fn wake(&'static self) {
        self.ready.notify_all();

        if self.started.swap(true, Ordering::AcqRel) {
            return;
        }

        std::thread::spawn(move || self.run());
    }

    /// The thread's whole life: one piece of work at a time, and no end of its own — a request
    /// that is never made is a wait, which costs nothing.
    ///
    /// A panic is contained here for the reason the loader contains one: one file's failure is that
    /// file's, and the thread goes on to the next hover — a thread that died on one file would take
    /// every file after it with it, in silence. An adapter whose work has a hover waiting on an
    /// answer contains its own panic as well, and for a different reason: that one is about the
    /// answer reaching the hover rather than about the thread surviving.
    fn run(&'static self) {
        while let Some(work) = self.next() {
            let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(work));
        }
    }

    /// The next piece of work, waiting for one.
    fn next(&self) -> Option<Work> {
        loop {
            // The slot is looked at and let go of again rather than held across the idle work: what
            // this thread does between two looks is a second of work, and the lock it would hold
            // through it is the one a hover puts its request down through.
            if let Some(work) = self.take() {
                return Some(work);
            }

            if let Some(on_idle) = self.on_idle {
                on_idle();
            }

            self.wait();
        }
    }

    fn take(&self) -> Option<Work> {
        match self.slot.lock() {
            Ok(mut slot) => slot.take(),
            // A lock poisoned by a panic on another thread is still the same slot, and a queue of
            // one is not worth standing down over.
            Err(poisoned) => poisoned.into_inner().take(),
        }
    }

    /// The bound on the wait above, and what waits on it. Nothing is done with the answer: what a
    /// bound that runs out is for is a request that may have been made in the meantime, and the
    /// adapter's own work between two runs.
    ///
    /// It waits on the slot rather than for the tick: a request put down between this thread's last
    /// look at its slot and this wait — which is where the idle hook above runs — signalled a
    /// thread that was not yet waiting, so nothing was woken and a request a hover was already
    /// waiting on sat in the slot for the whole of the tick. The condition is checked under the
    /// lock, so a slot that is not empty is answered on the spot rather than after a second.
    fn wait(&self) {
        let slot = self
            .slot
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let _ = self
            .ready
            .wait_timeout_while(slot, IDLE_TICK, |slot| slot.is_none());
    }

    /// Take whatever is waiting and answer whether there was any: what a test asks when it has to
    /// know that a name an adapter's own list does not hold put nothing down.
    #[cfg(test)]
    pub(crate) fn take_queued(&'static self) -> bool {
        self.take().is_some()
    }
}

// ---------------------------------------------------------------- the run in flight

/// What one adapter is running now: the file, the process running it, and since when. It is what
/// tells a busy engine from one that has stopped answering, and it is the id a run that has to be
/// ended is ended by.
///
/// One table for the four rather than one each, keyed by which adapter. This is not a second
/// record of the processes this app started and not a second way to end one: verifying the
/// identity and doing the ending belong to `engine_processes` and to nothing else (ADR 5), and
/// what is here is only the question the ending is asked by, which each of the four used to
/// answer for itself.
struct InFlight {
    source: PathBuf,
    pid: u32,
    started: Instant,
}

static IN_FLIGHT: Lazy<Mutex<HashMap<Adapter, InFlight>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Say what `adapter` is running, and which process is running it, for the threads that may decide
/// it has stopped answering while the one that started it waits.
pub fn begin(adapter: Adapter, source: &Path, pid: u32) {
    if let Ok(mut in_flight) = IN_FLIGHT.lock() {
        in_flight.insert(
            adapter,
            InFlight {
                source: source.to_path_buf(),
                pid,
                started: Instant::now(),
            },
        );
    }
}

/// Say that what `adapter` was running has stopped of its own accord.
///
/// Nothing is ended here: the give-up below is what ends a run, and a run that ended by itself has
/// nothing left to end.
pub fn stop(adapter: Adapter) {
    if let Ok(mut in_flight) = IN_FLIGHT.lock() {
        in_flight.remove(&adapter);
    }
}

/// The file `adapter` is working on now, where it is working on one: a request for the file
/// already in flight is that request, and a second run beside it is a launch spent on a picture
/// nobody would look at.
pub fn running_source(adapter: Adapter) -> Option<PathBuf> {
    IN_FLIGHT
        .lock()
        .ok()?
        .get(&adapter)
        .map(|run| run.source.clone())
}

/// The process of a run that has outrun `give_up`, and nothing for a run that has not.
///
/// The engine that is kept between documents has something to be told before the process is
/// touched, so it asks for the id rather than for the ending (see `libreoffice_render`).
pub fn hung(adapter: Adapter, give_up: Duration) -> Option<u32> {
    let in_flight = IN_FLIGHT.lock().ok()?;
    let run = in_flight.get(&adapter)?;

    outran(run.started, give_up).then_some(run.pid)
}

/// End a run that has outrun `give_up`.
///
/// What is left to the caller is a run whose answer is not believed — a hover answered with a
/// refusal rather than one left spinning — and a slot free for the file behind it.
pub fn end_hung(adapter: Adapter, give_up: Duration) {
    if let Some(pid) = hung(adapter, give_up) {
        // Verified by name and start time before anything is ended, like every other process this
        // app holds a record of.
        engine_processes::terminate_owned(pid);
    }
}

/// Whether a run that began at `started` has had its chance. A run that has just started is a file
/// being read and is left to it; one that has been inside it longer than any of them takes is one
/// the engine is not going to finish.
fn outran(started: Instant, give_up: Duration) -> bool {
    started.elapsed() >= give_up
}

/// Wait for a process, ending it rather than waiting past `give_up`: its own status where it ended
/// inside the bound, and nothing where it was ended here.
///
/// The two are not the same answer, which is the whole of the distinction: a run this app gave up
/// on is not a run whose own answer may be believed. An adapter that wants only the two answers
/// asks for `.is_some()`; one that needs the status — a listing that reported a table of contents
/// and then ended badly, which is a listing that stops where the tool stopped — asks for it.
///
/// The bound and how often the wait looks are each adapter's own numbers, for the reasons each of
/// them gives: they are a question about a run rather than about a file, and the four do not
/// measure the same run.
pub fn wait(child: &mut Child, give_up: Duration, poll: Duration) -> Option<ExitStatus> {
    let deadline = Instant::now() + give_up;

    loop {
        match child.try_wait() {
            Ok(Some(status)) => return Some(status),
            Ok(None) if Instant::now() < deadline => std::thread::sleep(poll),
            _ => {
                let _ = child.kill();
                let _ = child.wait();
                return None;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::AtomicUsize;
    use std::sync::{mpsc, OnceLock};

    /// A worker of a test's own, rather than a static the tests would share: the newest-wins rule
    /// is about a slot, and a slot two tests are putting into is a slot that answers for neither.
    fn worker() -> &'static Worker {
        Box::leak(Box::new(Worker::new(None)))
    }

    /// What a request is answered by, and the whole of what it answers: the work put down is run,
    /// it is run on the worker's thread rather than the caller's — a caller here is the preview
    /// loop, and a loop held inside a run is a hover that does not come up — and the work waiting
    /// when a second request arrives is the one that is dropped, because the newest hover is the
    /// one the pointer is on.
    ///
    /// The first work is what makes the other two answers certain rather than lucky: it is held
    /// open until the two that follow have been put down, so the thread is inside it and cannot
    /// have looked at its slot. What that is testing is a queue of one with a replacement, which
    /// is the rule every one of the four had written down separately.
    #[test]
    fn takes_the_newest_request_and_leaves_nothing_of_the_one_before_it() {
        let worker = worker();
        let (ran, seen) = mpsc::channel::<&'static str>();
        let (replaced, replaced_ran) = (ran.clone(), ran.clone());

        // Occupies the thread, and is only let go of once both later requests are down.
        let (holding, held) = mpsc::channel::<()>();
        worker.request(move || {
            let _ = ran.send("holding");
            let _ = held.recv();
        });
        assert_eq!(
            seen.recv().expect("the first request is taken up"),
            "holding",
            "the request that was put down is run"
        );

        worker.request(move || {
            let _ = replaced_ran.send("replaced");
        });
        worker.request(move || {
            let _ = replaced.send("newest");
        });

        let _ = holding.send(());

        // A timeout rather than a plain `recv`: what is being said here is that nothing *else*
        // arrives, and a plain `recv` would pass whether the second request ran or not.
        assert_eq!(
            seen.recv_timeout(Duration::from_secs(5))
                .expect("the newest request"),
            "newest",
            "what was waiting is taken up in place of the one before it"
        );
        assert!(
            seen.recv_timeout(Duration::from_millis(200)).is_err(),
            "and the request that was replaced is not run afterwards"
        );
    }

    /// A panic in one piece of work is that piece's, and the thread goes on: a thread that died on
    /// one file would take every file after it with it, in silence — which is a silent app that
    /// still answers every question about a file it has already seen.
    #[test]
    fn takes_the_thread_on_a_panic_and_keeps_going() {
        let worker = worker();
        let (ran, seen) = mpsc::channel::<&'static str>();

        worker.request(|| panic!("a run that failed"));
        worker.request(move || {
            let _ = ran.send("after");
        });

        assert_eq!(
            seen.recv_timeout(Duration::from_secs(5))
                .expect("the next request"),
            "after",
            "a work that panicked does not take the thread with it"
        );
    }

    /// The adapter's own work between runs happens with an empty slot and not around a request,
    /// which is what keeps it off a thread a hover is waiting on and what stops a wait from
    /// delaying the request that woke it. The thread is woken rather than given something to do,
    /// which is the whole of what a request that is a flag and not a piece of work needs.
    #[test]
    fn does_an_adapters_own_tick_only_while_its_slot_is_empty() {
        let worker: &'static Worker = Box::leak(Box::new(Worker::new(Some(tick))));
        worker.wake();

        while TICKS.load(Ordering::Acquire) == 0 {
            std::thread::sleep(Duration::from_millis(10));
        }

        let (holding, held) = mpsc::channel::<()>();
        let at_start = TICKS.load(Ordering::Acquire);
        worker.request(move || {
            let _ = held.recv();
        });

        // Well inside the tick the thread would reach on its own, so what is being said is that the
        // thread is inside the work and not looking, rather than that it is slow.
        std::thread::sleep(Duration::from_millis(200));
        assert_eq!(
            TICKS.load(Ordering::Acquire),
            at_start,
            "nothing is looked at between a request being taken up and the work being done"
        );

        let _ = holding.send(());
    }

    /// What an adapter's own tick counts for the test above. A tick is not handed anything — it
    /// is a plain function — so what it counts in is a static, and a test's own worker is the
    /// only thing that could be counting into it.
    static TICKS: AtomicUsize = AtomicUsize::new(0);

    fn tick() {
        TICKS.fetch_add(1, Ordering::AcqRel);
    }

    /// How long a worker may take to take up a request that is already waiting for it, well
    /// inside the tick it waits on for itself: a request answered inside this was not left for
    /// the tick to run out.
    const ANSWERED: Duration = Duration::from_millis(300);

    /// The worker the idle hook of the test below puts a request down for. A hook is handed
    /// nothing, so this is where the worker that owns it is, and only this test's own worker is
    /// ever put into it.
    static IDLED_FOR: OnceLock<&'static Worker> = OnceLock::new();

    /// How many times that hook has run, which is also how the hook knows it is the first time: a
    /// hook that put a request down every time would leave the thread no idle to be measured in.
    static IDLES: AtomicUsize = AtomicUsize::new(0);

    /// What the work of the test below has reported. It is a record rather than a signal because
    /// the first of those works is put down by a plain function, which cannot hold a sender.
    static REPORTED: Mutex<Vec<&'static str>> = Mutex::new(Vec::new());

    /// The idle hook, putting one request down for the worker it belongs to on the first call and
    /// never again. That first call is inside the gap on purpose: the thread runs it between
    /// letting its slot go and taking it again, so the request it puts down signalled a thread
    /// that was not yet waiting.
    fn put_down_the_first_time() {
        if IDLES.fetch_add(1, Ordering::AcqRel) != 0 {
            return;
        }

        if let Some(worker) = IDLED_FOR.get() {
            worker.request(|| report("in the gap"));
        }
    }

    /// A request put down where the thread cannot hear it was taken up without waiting out the
    /// tick, and the thread is still the same afterwards: it takes what comes next, and it still
    /// waits on a slot with nothing in it rather than looking at it again and again.
    ///
    /// The gap is hit rather than raced into, because the idle hook runs exactly inside it — the
    /// hook is what a gap of this kind is wide open for — so a request put down from there missed
    /// the signal on every machine rather than on a slow one. What is left to measure is the wait,
    /// and a wait is a number.
    #[test]
    fn takes_a_request_put_down_before_it_is_waiting_without_waiting_out_the_tick() {
        let worker: &'static Worker =
            Box::leak(Box::new(Worker::new(Some(put_down_the_first_time))));
        assert!(
            IDLED_FOR.set(worker).is_ok(),
            "the hook is put into one worker of a test's own, as every other worker here is"
        );
        worker.wake();

        assert_ran_within(ANSWERED, "in the gap");

        worker.request(|| report("after the gap"));
        assert_ran_within(ANSWERED, "after the gap");

        // The same hook shows the wait is still a wait. An empty slot is found empty once and then
        // waited on, which is one call here rather than the hundreds a thread that never slept
        // would make of it.
        let idles = IDLES.load(Ordering::Acquire);
        std::thread::sleep(ANSWERED);
        assert!(
            IDLES.load(Ordering::Acquire) <= idles + 1,
            "a slot with nothing in it is waited on rather than looked at again and again"
        );
    }

    /// The whole of what a work can say to the test's thread: that it ran.
    fn report(name: &'static str) {
        REPORTED
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .push(name);
    }

    /// That `name` ran inside `ceiling`, and not only once the tick has run out. It is looked for
    /// well past the ceiling so that a failure says how long it really took rather than how long
    /// the ceiling was, and the time it took is in the failure because what is being looked for is
    /// a stall, and a stall is a number.
    fn assert_ran_within(ceiling: Duration, name: &str) {
        let start = Instant::now();
        let watched = ceiling + IDLE_TICK;

        while !reported(name) && start.elapsed() < watched {
            std::thread::sleep(Duration::from_millis(5));
        }

        let took = start.elapsed();
        assert!(
            took < ceiling && reported(name),
            "{name:?} is taken up rather than left for the {IDLE_TICK:?} tick to run out, and \
             took {took:?} of the {ceiling:?} it was given"
        );
    }

    fn reported(name: &str) -> bool {
        REPORTED
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .contains(&name)
    }

    /// The whole of what the in-flight record is for, in one: what is being run is said to be what
    /// it is, it is only the adapter that said it that is told, and a run that has stopped of its
    /// own accord is not left standing as one that is still going.
    #[test]
    fn a_run_is_said_to_be_what_it_is_and_forgotten_when_it_stops() {
        let source = PathBuf::from("book.epub");
        let pid = std::process::id();

        assert!(
            running_source(Adapter::Calibre).is_none(),
            "nothing is running before anything has been begun"
        );

        begin(Adapter::Calibre, &source, pid);
        assert_eq!(
            running_source(Adapter::Calibre).as_deref(),
            Some(source.as_path()),
            "what is running is the file it was begun for"
        );
        assert!(
            hung(Adapter::Calibre, Duration::from_secs(30)).is_none(),
            "a run that has just begun is a file being read and is left to it"
        );
        assert!(
            running_source(Adapter::ImageMagick).is_none(),
            "and it is not another adapter's run"
        );

        stop(Adapter::Calibre);
        assert!(
            running_source(Adapter::Calibre).is_none(),
            "a run that has stopped is not one that is still going"
        );
        assert!(
            hung(Adapter::Calibre, Duration::ZERO).is_none(),
            "and there is nothing to end, however long a bound it is given"
        );
    }

    /// The one reason this is a keyed table and not one record: four adapters each have a run of
    /// their own in flight at the same time, and one adapter's run is not another adapter's to
    /// answer for. This is what four private records could not do.
    #[test]
    fn one_adapters_run_is_not_anothers() {
        begin(Adapter::PeaZip, Path::new("backup.cab"), std::process::id());
        begin(
            Adapter::ImageMagick,
            Path::new("shot.nef"),
            std::process::id(),
        );

        assert_eq!(
            running_source(Adapter::PeaZip).as_deref(),
            Some(Path::new("backup.cab"))
        );
        assert_eq!(
            running_source(Adapter::ImageMagick).as_deref(),
            Some(Path::new("shot.nef")),
            "each adapter is told about its own run and no other"
        );
        assert!(
            running_source(Adapter::LibreOffice).is_none(),
            "and an adapter with nothing running is told so"
        );

        stop(Adapter::PeaZip);
        assert!(
            running_source(Adapter::ImageMagick).is_some(),
            "ending one adapter's run leaves the other's alone"
        );

        stop(Adapter::ImageMagick);
    }

    /// What the give-up is decided from, at the boundary rather than near it: a run that has just
    /// begun is not hung, a run that has not quite reached the bound is not hung, and one that has
    /// reached it or gone past it is.
    #[test]
    fn a_run_is_hung_only_once_it_has_outrun_the_give_up() {
        let began = Duration::from_secs(60);
        let give_up = Duration::from_secs(30);

        assert!(!outran(Instant::now(), give_up));
        assert!(!outran(
            Instant::now() - (give_up - Duration::from_secs(1)),
            give_up
        ));
        assert!(outran(Instant::now() - give_up, give_up));
        assert!(outran(
            Instant::now() - (give_up + Duration::from_secs(30)),
            give_up
        ));
        assert!(outran(Instant::now() - began, give_up));
    }

    /// The give-up is a stop and not a decision, and the whole of that is two answers: a run that
    /// has had its chance is ended, and a run that has not is left to finish. A process that stays
    /// up stands in for the engine — a test is not going to make one of the four spin on a file —
    /// and it is recorded by image name, which is the check that keeps an id from being acted on
    /// by itself.
    #[test]
    fn a_run_that_has_outrun_its_give_up_is_ended_and_one_that_has_not_is_not() {
        let _stand_in = engine_processes::STAND_IN
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();

        assert!(
            engine_processes::processes_named("ping.exe").contains(&pid),
            "the stand-in runs the image the record below names it by"
        );
        engine_processes::record("ping.exe", pid);

        begin(Adapter::LibreOffice, Path::new("drawing.cdr"), pid);
        end_hung(Adapter::LibreOffice, Duration::from_secs(30));
        assert!(
            engine_processes::is_running(pid),
            "a run that has just begun is a file being read, and is left to it"
        );

        // A bound of nothing is the shortest one a run can be outrun by, and says the same thing
        // about the decision without waiting half a minute to say it.
        end_hung(Adapter::LibreOffice, Duration::ZERO);
        assert!(
            !engine_processes::is_running(pid),
            "the engine a run has outrun is ended"
        );

        stop(Adapter::LibreOffice);
        let _ = engine.wait();
    }

    /// The bound the engine thread holds over its own run: a process that ends inside it answers
    /// with its own status, and one that outlasts it is ended there rather than waited on, so that
    /// the bound is a bound on the run rather than only on a file that happens to be asked for
    /// after it.
    #[test]
    fn ends_a_run_rather_than_waiting_past_its_bound() {
        let _stand_in = engine_processes::STAND_IN
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let give_up = Duration::from_secs(30);
        let poll = Duration::from_millis(50);

        let mut quick = std::process::Command::new("ping")
            .args(["-n", "1", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process that ends on its own");
        let status = wait(&mut quick, give_up, poll).expect("a status");
        assert!(
            status.success(),
            "a run that finishes inside its bound is answered as it always was"
        );

        let mut engine = std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a process to stand in for the engine");
        let pid = engine.id();
        let started = Instant::now();

        assert!(
            wait(&mut engine, Duration::from_millis(300), poll).is_none(),
            "a run that has outrun its bound is ended, not waited on"
        );
        assert!(
            started.elapsed() < Duration::from_secs(5),
            "and the wait ends at the bound it was given rather than when the process would have"
        );
        assert!(
            !engine_processes::is_running(pid),
            "the engine is gone with it"
        );
        let _ = engine.wait();
    }
}
