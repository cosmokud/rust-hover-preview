use std::sync::atomic::Ordering;
use std::sync::mpsc::{self, Receiver, RecvTimeoutError, Sender};
use std::time::{Duration, Instant};

use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{PeekMessageW, MSG, PM_NOREMOVE, WM_APP};

use super::api::{
    drop_placement, is_showing, note_engine_up, note_failure, renders_html, take_placement, wanted,
    ENGINE_THREAD, SHOWING,
};
use super::environment::pump_messages;
use super::host::Host;

use crate::config::config::{EngineIdle, PreviewType, DEFAULT_WEBVIEW_IDLE_SECS};
use crate::CONFIG;

/// What the preview thread asks the engine's thread to do.
pub(super) enum Command {
    /// Draw what is wanted, as the want made under this generation: the document itself
    /// travels in `WANTED`, and the generation is what says whether this ask is still the
    /// newest one by the time the engine's thread takes it up (see `WANTED`).
    Show {
        generation: u64,
    },
    /// Move the window of a document the engine is already holding, to the newest box
    /// published in `PLACED`.
    ///
    /// The box is not carried in the command and that is the whole of it: a drag asks for a
    /// box on every pointer move, so carrying one per command would queue a move per move for
    /// a thread that takes one command per pass. The box travels in the cell instead, where
    /// asking again replaces what was there, so a drag of any length is answered by one move
    /// to where the hand ended up rather than by a backlog of where it was (see `PLACED`).
    Place,
    Hide,
    Shutdown,
}

pub(super) struct Engine {
    pub(super) sender: Sender<Command>,
    pub(super) thread: std::thread::JoinHandle<()>,
}

impl Engine {
    /// Start the engine's thread. Nothing is created until the first document is asked
    /// for: a machine that never hovers one never starts a browser.
    pub(super) fn start() -> Self {
        let (sender, receiver) = mpsc::channel();
        let thread = std::thread::spawn(move || engine_thread(receiver));

        Self { sender, thread }
    }
}

/// How long the engine is kept after its last document. It is read from the
/// configuration each time rather than captured, so an edit applies to the engine that
/// is already warm.
fn idle_timeout() -> Option<Duration> {
    CONFIG
        .lock()
        .map(|config| config.webview_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS))
        .as_duration()
}

/// Whether the browser is kept whatever the user is doing, which is the `Persistent` toggle
/// at the top of the same submenu.
///
/// Read rather than captured, for the reason the idle time is: it is asked while the engine
/// is up, so a click applies to the browser that is already warm.
fn engine_persistent() -> bool {
    CONFIG
        .lock()
        .map(|config| config.webview_persistent)
        .unwrap_or(false)
}

/// Write one line to the trace file when `RHP_WEBVIEW_TRACE` is set, for the same
/// reason the probes exist: an engine that does not come up says nothing on its own,
/// and this is what it says. Nothing is written when the variable is not set, so an
/// ordinary run leaves no file anywhere.
pub(super) fn trace(message: &str) {
    if std::env::var_os("RHP_WEBVIEW_TRACE").is_none() {
        return;
    }

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-webview-trace.log"))
    {
        use std::io::Write;
        let _ = writeln!(file, "{message}");
    }
}

fn engine_thread(commands: Receiver<Command>) {
    // The engine's own apartment, and its own thread: WebView2 must be created on a
    // thread that is pumping messages, and this is that thread.
    let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };

    // The thread's message queue, made before anything can be posted to it: a wakeup posted
    // to a thread that has none is dropped, and a want published while this thread is still
    // bringing its browser up would be the first thing that had to be waited for.
    {
        let mut message = MSG::default();
        unsafe {
            let _ = PeekMessageW(&mut message, None, WM_APP, WM_APP, PM_NOREMOVE);
        }
    }
    ENGINE_THREAD.store(unsafe { GetCurrentThreadId() }, Ordering::Release);

    // No placement is owed by a thread that has only just begun: whatever was published before
    // it was aimed at a window this engine has not made, and a flag left set by a previous
    // engine's thread would stop the next box from ever being asked for (see `PLACE_ASKED`).
    drop_placement();

    let mut host: Option<Host> = None;
    let mut idle_since = Instant::now();

    loop {
        // While there is nothing on screen the wait is a poll of the idle clock; while
        // a document is up it is a poll of the message queue, because a browser needs
        // the thread that made it to keep retrieving messages. With no engine at all
        // there is nothing to watch for, so the wait is long enough that only a document
        // wakes the thread, and an app left alone is one thread asleep on a channel.
        let wait = if SHOWING.load(Ordering::Acquire) {
            Duration::from_millis(5)
        } else if host.is_some() {
            Duration::from_millis(250)
        } else {
            Duration::from_secs(60 * 60)
        };

        match commands.recv_timeout(wait) {
            Ok(Command::Shutdown) => break,
            Ok(Command::Hide) => {
                if let Some(host) = host.as_mut() {
                    host.hide();
                }
                idle_since = Instant::now();
            }
            Ok(Command::Place) => carry_out_placement(&mut host),
            Ok(Command::Show { generation }) => {
                // What was asked for here may have been asked for after: the pointer moves
                // while a document is on its way, and the file this is about is then one
                // it has left. Such a want is not taken up at all — nothing is navigated
                // to, and nothing is shown — which is what keeps a folder of documents
                // from being drawn one after another at the speed of the hand crossing it.
                let ask = wanted().filter(|wanted| wanted.generation == generation);

                if ask.is_none() {
                    trace(&format!(
                        "engine: dropped a show the loop has moved on from (generation {generation})"
                    ));
                }

                if let Some(ask) = ask {
                    trace(&format!("engine: show {}", ask.path.display()));

                    if host.is_none() {
                        host = Host::create();

                        // An engine that could not be had is noted, so that hovers stop
                        // opening a box nothing will be drawn into until the folder it
                        // could not have is free again.
                        if host.is_some() {
                            note_engine_up();
                        } else {
                            note_failure();
                        }
                        trace(&format!("engine: host created: {}", host.is_some()));
                    }

                    if let Some(host) = host.as_mut() {
                        host.show(&ask);
                        trace(&format!("engine: shown: {}", is_showing()));
                    }

                    // A navigation the browser never answered is not a document to
                    // hand another one to: it stops answering rather than failing, and
                    // what it would put up next is the file before this one — so it is
                    // let go of here and begun again by the next document, which is
                    // what the app does with an engine that has stopped answering
                    // wherever else one is kept.
                    if host.as_ref().is_some_and(Host::is_hung) {
                        trace("engine: let go, the browser stopped answering");

                        if let Some(mut host) = host.take() {
                            host.close();
                        }
                    }
                    idle_since = Instant::now();
                }
            }
            Err(RecvTimeoutError::Timeout) => {}
            Err(RecvTimeoutError::Disconnected) => break,
        }

        pump_messages();

        // The engine draws three kinds, and every gate that could reach one switched off is a
        // browser held for nothing: while none can be shown no hover can be answered with a
        // document, a specimen or a page at all, so there is nothing warm to keep. It is read
        // here rather than being told because a gate can be closed either way — in the tray,
        // or in `config.ini` for the watcher to reload — and because the thread that would be
        // told is this one, parked on its channel. The browser's own children are its
        // business: ending it ends them.
        //
        // The text gate is one of them, but only while it asks for pages: a page of HTML is
        // a text file, so the kind's gate is what switches the browser off — and with the
        // kind on by default and the switch for pages off by default, the gate alone would
        // hold a browser for everyone who never previews a page (see `draws`).
        if host.is_some()
            && !PreviewType::Vector.enabled()
            && !PreviewType::Fonts.enabled()
            && !(PreviewType::Text.enabled() && renders_html())
        {
            trace("engine: let go, the kinds are switched off");

            if let Some(mut host) = host.take() {
                host.close();
            }
            idle_since = Instant::now();
        }

        // A document that has been off screen for as long as the engine is kept for is one
        // the engine is let go of, browser process and all.
        //
        // Which question is asked is the `Persistent` toggle at the top of the TTL submenu.
        // Persistent, it is the idle time, exactly as it was before the toggle existed; not
        // persistent — the way this starts — it is the AFK timer, and the idle time is not
        // consulted at all. What an idle time is for is the memory an engine holds while the
        // user is elsewhere, which is the question the AFK timer asks directly.
        //
        // The thread is not let go of with it, and that is the whole of it: the channel
        // it holds is the one a hover sends into, and a thread that ended here would
        // leave every document after the first idle timeout with a message nobody reads.
        // What the app would show is no document at all — the engine's window is the
        // whole of an SVG preview — for the rest of the run.
        let expired = match host.as_ref() {
            None => false,
            // A document on screen is the engine earning its keep: the window is the
            // preview, and the browser behind it is what is drawing it.
            Some(_) if SHOWING.load(Ordering::Acquire) => false,
            Some(_) if engine_persistent() => {
                idle_timeout().is_some_and(|limit| idle_since.elapsed() >= limit)
            }
            Some(_) => crate::app::afk::expired(),
        };

        if expired {
            trace("engine: let go after idle");

            if let Some(mut host) = host.take() {
                host.close();
            }
        }
    }

    if let Some(mut host) = host {
        host.close();
    }

    // The thread is going away, so nothing is posted to it any more: a want published after
    // this is one the next engine's thread takes up (see `ENGINE_THREAD`).
    ENGINE_THREAD.store(0, Ordering::Release);

    unsafe {
        windows::Win32::System::Com::CoUninitialize();
    }
}

/// Carry out the placement the engine's thread has been asked for, if it still belongs to a
/// window this engine has.
///
/// The box in hand is the newest one rather than the one the command was sent for: a drag
/// publishes a box per pointer move and only the last of them is a place worth putting the
/// window in (see `PLACED`).
fn carry_out_placement(host: &mut Option<Host>) {
    let Some(placement) = take_placement() else {
        return;
    };

    // A placement for a document the engine no longer holds is a drag of a
    // window that has been swapped or closed since, and moving that window
    // would put somebody else's document where the hand let go.
    if !host
        .as_ref()
        .is_some_and(|host| host.holds(&placement.path))
    {
        trace(&format!(
            "engine: dropped a placement for {} it does not hold",
            placement.path.display()
        ));
        return;
    }

    let Some(ask) = wanted().filter(|ask| ask.path == placement.path) else {
        return;
    };

    // A backdrop is the page's as well as the controller's, so a box
    // published under a backdrop the page was not written for is a page to
    // write again rather than a window to move — which is what `show` is
    // for, and it is asked for once rather than per pointer move, because
    // every box of a drag carries the backdrop that is on record.
    if host
        .as_ref()
        .is_some_and(|host| host.needs_page(&placement))
    {
        if let Some(host) = host.as_mut() {
            host.show(&ask);
        }

        return;
    }

    // And a move: nothing navigates, nothing is woken and nothing is
    // re-shown, all of which describe a document arriving rather than a
    // window being carried (see `Host::place`).
    if let Some(host) = host.as_mut() {
        host.place(&placement);
    }
}
