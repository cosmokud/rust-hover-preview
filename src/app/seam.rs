//! The facts the shell side and the preview-window side exchange, in one module.
//!
//! The tray, the explorer hook, the wheel watcher and the keyboard watcher are the main thread and the
//! threads Windows made for the two low-level hooks; the preview loop is a thread of its own drawing
//! layered windows. What crosses between them is small and deliberate, and it is of three kinds.
//! Values: a channel a `PreviewMessage` is sent on and an event a wait is woken by; the wheel notches a
//! hook counts and the loop drains; a pin's two window handles, the boxes a notch belongs to, whether a
//! pin is up, whether its find bar is, and whether the preview is the surface the pointer stands on.
//! Notices: the same, the other way — a setting the pin's own chrome asks the tray to turn, a shortcut
//! pressed at a preview, a line for the hook's trace — on a channel of its own.
//!
//! And the third kind, which is the reason this module is a doorway as much as a shelf: **every call one
//! side makes of the other is named from here**, even the ones that are calls today. A value can be
//! published where a caller cannot be, so the first two kinds are the whole of what a second process is
//! told; a call is not, and a call named from here is one whose body can become a message, a published
//! fact or a round trip without a call site to follow it. That is what the list of names at the end of
//! this module is for, and it is the one thing here that is not a fact either side owns: the folder walk
//! a pin steps through, the sort the listing is showing, the sizes a page can be drawn at.
//!
//! They used to live wherever the code that first needed one happened to be — the channel and its
//! event in `tick`, the wheel counts in `wheel_input`, the pin's published half in `key_input` — so
//! each side reached into the other side's module for a fact neither owned: `ui::preview_window` into
//! `shell` for a notch it did not count, `shell` into `ui` for the message it posts. That a module
//! line is not a thread line is exactly what read wrong there, and it is why a hook callback could be
//! found holding a type the drawing thread had defined.
//!
//! Nothing here is new. It is the same atomics, the same orderings and the same timing, and no reader
//! and no writer moves off the thread it was on: the hook still counts a notch, the loop still swaps
//! it, the loop still publishes a handle, the hook still reads it. **No fact here is read through a
//! lock**: the three `Mutex`es are the two channels' own — a sender and two ends — taken to send a
//! notice and to drain one on the thread that already owns the drain, and never inside a low-level
//! callback, which is the one place a lock is forbidden. Each accessor stays beside its static for the
//! reason the static is here at all: what is read and written together is one fact, and a count kept in
//! one file with its reader in another is a fact with two homes.
//!
//! The facts are one `#[repr(C)]` struct rather than a scatter of statics, and *where* that struct is
//! kept is settled once, at start-up: in a shared section this run opens, or in this process's own
//! copy. What is in here is the whole of what a second process is told, which is what makes the split
//! a change of one module.

pub(crate) use crate::ui::preview_window::PreviewCursorHover;
use crate::ui::preview_window::{AvoidRegion, PinPlanned, ScreenRegion};
use once_cell::sync::Lazy;
use std::cell::UnsafeCell;
use std::ffi::OsString;
use std::mem::size_of;
use std::ops::Deref;
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::path::{Path, PathBuf};
use std::sync::atomic::{
    AtomicBool, AtomicI32, AtomicIsize, AtomicPtr, AtomicU32, AtomicU64, AtomicUsize, Ordering,
};
use std::sync::mpsc::{Receiver, Sender};
use std::sync::Mutex;
use std::time::{SystemTime, UNIX_EPOCH};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{HANDLE, HWND, INVALID_HANDLE_VALUE};
use windows::Win32::System::Memory::{
    CreateFileMappingW, MapViewOfFile, OpenFileMappingW, FILE_MAP_ALL_ACCESS, PAGE_READWRITE,
};
use windows::Win32::System::Threading::{
    CreateEventW, GetCurrentProcessId, ResetEvent, SetEvent, WaitForSingleObject,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetWindowThreadProcessId, MsgWaitForMultipleObjectsEx, PeekMessageW, MSG, MWMO_INPUTAVAILABLE,
    PM_NOREMOVE, QS_ALLINPUT,
};

/// What one viewer of this run knows about its own window: the block it stands in.
///
/// **A run's facts are not one record but a run's and a window's.** Two viewers stand at once once a
/// pin has been detached — the one serving the hovers and the one standing as a window of its own — and
/// everything either published used to land in the same field: a notch turned over the detached pin was
/// claimed by the hover's viewer, and the window under the pointer, the boxes a notch belongs to, the
/// engine's window and the pin's own handles were whichever viewer wrote last. So what one viewer
/// publishes about its own window is a block of its own (see `SeamFacts`), and a process reaches its
/// own block with `the_facts` and another's with `facts_of`, by pid.
///
/// **One type, one home per block.** A reader below names a field through a `Fact` rather than naming a
/// static, so nothing outside this module knows which block it is inside, and the code that counts a
/// notch inside a low-level hook is the same code whichever block it lands in.
///
/// The layout has to be the same bytes in both processes, which is why this is `#[repr(C)]` with an
/// atomic and nothing else in every field: no padding the compiler invented between one process's copy
/// and the other's, and nothing in it that is not a number. `AtomicU64` and `AtomicIsize` are
/// word-sized and word-aligned and `AtomicBool` is one byte, so two processes agree about where a field
/// starts.
///
/// What each field is and why it exists is said beside its `Fact` in the rest of this module.
pub(crate) struct ViewerFacts {
    /// The pid of the process standing in this block, or 0 while it is free: claimed once, at start-up,
    /// by the viewer that draws in it (see `claim_a_block`).
    owner: AtomicU32,
    text_scroll_delta: AtomicI32,
    pin_zoom_delta: AtomicI32,
    pin_volume_delta: AtomicI32,
    pin_text_size_delta: AtomicI32,
    pin_playlist_delta: AtomicI32,
    preview_under_pointer: AtomicU32,
    pin_keys_owner: AtomicU64,
    pin_keys_video: AtomicU64,
    pin_keys_topmost: AtomicBool,
    pin_follows_the_listing: AtomicBool,
    /// Pin-key commands seen since this was last called, one per press, in the order pressed.
    ///
    /// Drained by the preview loop on the same tick that drains the pin key's own counter, so a
    /// command travels the same path a press of the pin's own caption button does: it is queued, and
    /// the loop acts on it (see `ask_pin`). One count per command of `PIN_KEY_COMMANDS`, in the order
    /// that table lists them.
    pin_key_presses: [AtomicU32; PIN_COMMAND_COUNT],
    /// Win chords the hook took for this block's pin, one count per chord, in the order
    /// `PIN_WIN_CHORDS` lists them (see `take_pin_win_chords`).
    ///
    /// Apart from the table above because a Win chord is not a key of the pin's: it is the Shell's own
    /// arrangement of a window, answered by this app only because the box is the pin's (see
    /// `key_input::pin_win_chord`).
    pin_win_chords: [AtomicU32; PIN_WIN_CHORD_COUNT],
    pin_find_bar_up: AtomicBool,
    pin_up: AtomicBool,
    pin_end_requested: AtomicBool,
    pin_menu_dismiss_requested: AtomicBool,
    pin_menu_pressed: AtomicBool,
    pin_resumed: AtomicBool,
    pin_handed_off: AtomicBool,
    pin_bubble_hwnd: AtomicIsize,
    pin_menu_hwnd: AtomicIsize,
    pin_file_in_flight: AtomicU64,
    /// The window a pin is to be put behind, or zero where nothing is asked (see
    /// `ask_the_pin_to_go_behind`).
    pin_behind_ask: AtomicIsize,
    /// The path of the file a pin is showing, one UTF-16 unit to a word, and how many of them are the path
    /// (see `publish_the_pinned_path`).
    pin_path: [AtomicU32; SEAM_PIN_PATH_UNITS],
    pin_path_len: AtomicU32,
    /// A count of the writes to `pin_path`, odd in the middle of one, so that a reader can tell a name it
    /// read whole from one it read being written.
    pin_path_writing: AtomicU32,
    /// The five boxes, each packed the way `PublishedRegion` packs one (see there).
    pin_volume_region: AtomicU64,
    text_scroll_region: AtomicU64,
    text_size_region: AtomicU64,
    pin_zoom_region: AtomicU64,
    models_band: AtomicU64,
    pin_playlist_band: AtomicU64,
    /// The box the pinned window stands in (see `PIN_WINDOW_BOX`).
    pin_window_box: AtomicU64,
    models_notches: AtomicI32,
    models_notch_point: AtomicU64,
    pin_presses: AtomicU32,
    /// Which of this app's own surfaces the pointer stands on, one bit a surface (see
    /// `cursor_preview_hover`): the preview's picture, the player's window, the engine's, or none.
    preview_hover: AtomicU32,
    /// The box the preview surface on screen stands in (see `preview_screen_rect`).
    preview_rect: AtomicU64,
    /// Whether the page the engine is drawing is one that runs, and whether a drag that began on it is
    /// still down — the two halves of the page rule that are the viewer's to know (see
    /// `ENGINE_PAGE_RUNS`).
    engine_page_runs: AtomicBool,
    engine_page_drag: AtomicBool,
    /// The window the engine draws in while it has one, and whether it is showing a document: the two
    /// answers the tray's own key and press hooks ask about a preview whose window is the viewer's (see
    /// `ENGINE_HWND` and `ENGINE_SHOWING`).
    engine_hwnd: AtomicIsize,
    engine_showing: AtomicBool,
    /// Whether a text preview in full mode is on screen, and whether a spinner is standing in for a
    /// page that is on its way. Both hold the pointer; what they hold it *through* differs (see
    /// `TEXT_PREVIEW_HOLDING` and the two boxes below).
    text_preview_holding: AtomicBool,
    waiting_preview_holding: AtomicBool,
    /// The two boxes a text preview or a pin holds the pointer through, packed the way
    /// `PublishedRegion` packs one (see `HOLD_REGION_JOURNEY`).
    hold_region_journey: AtomicU64,
    hold_region_preview: AtomicU64,
    /// The box the view drew the item under the pointer in (see `POINTER_ITEM_BOX`).
    pointer_item_box: AtomicU64,
    /// The left button's state and how many presses it has taken, published by the one place in this
    /// app that reads the buttons (see `PIN_MEDIA_LEFT_DOWN`).
    pin_media_left_down: AtomicBool,
    pin_media_left_presses: AtomicU64,
    /// The wall clock of the preview loop's last tick, in milliseconds since the epoch: a number both
    /// processes can read, which is the whole of why it is a wall clock rather than an `Instant` (see
    /// `preview_stall_ms`).
    preview_alive_unix_ms: AtomicU64,
    /// The boxes a standing pin's own buttons answer a press in, each packed whole, and `NO_REGION`
    /// for the ones this pin has no button in (see `pin_button_covers`).
    pin_buttons: [AtomicU64; PIN_BUTTONS],
}

// Safety: every field is an atomic, and nothing here is a reference into this process, which is what
// makes the same type usable as a section two processes map.
unsafe impl Sync for ViewerFacts {}

impl ViewerFacts {
    /// Every fact of a window at rest: no box published, no pin up, nothing counted.
    const fn new() -> Self {
        Self {
            owner: AtomicU32::new(0),
            text_scroll_delta: AtomicI32::new(0),
            pin_zoom_delta: AtomicI32::new(0),
            pin_volume_delta: AtomicI32::new(0),
            pin_text_size_delta: AtomicI32::new(0),
            pin_playlist_delta: AtomicI32::new(0),
            preview_under_pointer: AtomicU32::new(0),
            pin_keys_owner: AtomicU64::new(0),
            pin_keys_video: AtomicU64::new(0),
            pin_keys_topmost: AtomicBool::new(false),
            pin_follows_the_listing: AtomicBool::new(false),
            pin_key_presses: [const { AtomicU32::new(0) }; PIN_COMMAND_COUNT],
            pin_win_chords: [const { AtomicU32::new(0) }; PIN_WIN_CHORD_COUNT],
            pin_find_bar_up: AtomicBool::new(false),
            pin_up: AtomicBool::new(false),
            pin_end_requested: AtomicBool::new(false),
            pin_menu_dismiss_requested: AtomicBool::new(false),
            pin_menu_pressed: AtomicBool::new(false),
            pin_resumed: AtomicBool::new(false),
            pin_handed_off: AtomicBool::new(false),
            pin_bubble_hwnd: AtomicIsize::new(0),
            pin_menu_hwnd: AtomicIsize::new(0),
            pin_file_in_flight: AtomicU64::new(0),
            pin_behind_ask: AtomicIsize::new(0),
            pin_path: [const { AtomicU32::new(0) }; SEAM_PIN_PATH_UNITS],
            pin_path_len: AtomicU32::new(0),
            pin_path_writing: AtomicU32::new(0),
            pin_volume_region: AtomicU64::new(NO_REGION),
            text_scroll_region: AtomicU64::new(NO_REGION),
            text_size_region: AtomicU64::new(NO_REGION),
            pin_zoom_region: AtomicU64::new(NO_REGION),
            models_band: AtomicU64::new(NO_REGION),
            pin_playlist_band: AtomicU64::new(NO_REGION),
            pin_window_box: AtomicU64::new(NO_REGION),
            models_notches: AtomicI32::new(0),
            models_notch_point: AtomicU64::new(0),
            pin_presses: AtomicU32::new(0),
            preview_hover: AtomicU32::new(0),
            preview_rect: AtomicU64::new(NO_REGION),
            engine_page_runs: AtomicBool::new(false),
            engine_page_drag: AtomicBool::new(false),
            engine_hwnd: AtomicIsize::new(0),
            engine_showing: AtomicBool::new(false),
            text_preview_holding: AtomicBool::new(false),
            waiting_preview_holding: AtomicBool::new(false),
            hold_region_journey: AtomicU64::new(NO_REGION),
            hold_region_preview: AtomicU64::new(NO_REGION),
            pointer_item_box: AtomicU64::new(NO_REGION),
            pin_media_left_down: AtomicBool::new(false),
            pin_media_left_presses: AtomicU64::new(0),
            preview_alive_unix_ms: AtomicU64::new(0),
            pin_buttons: [const { AtomicU64::new(NO_REGION) }; PIN_BUTTONS],
        }
    }
}

/// How many viewers one run may have standing at once, which is how many blocks its section holds:
/// the tray's own, the viewer serving the hovers, and one per pin detached into a window of its own.
///
/// **Sized against the windows a run may leave standing beside the hovers**: `DETACHED_VIEWERS` of them,
/// each a viewer that keeps the block it claimed for as long as its window stands, plus the tray's block,
/// plus the serving viewer's, plus two replacements that may be beginning before the notice retiring the
/// viewer before them is read. A block whose viewer has gone is claimed again by the next one (see
/// `claim_a_block`), so what this has to hold is the windows a run is standing, not every viewer it ever
/// began.
const SEAM_VIEWERS: usize = DETACHED_VIEWERS + 4;

/// The facts of a run: the blocks its viewers stand in, and the two things that are the run's own.
///
/// **Where a window's facts end and the run's begin** is the whole of what this split is. A block holds
/// what one viewer published about the window it draws and the counts turned at it; what is here holds
/// the two rings the messages cross on, and the two counts a hook of the tray's keeps for the run rather
/// than for a window — the pin key's held chords, which are one hand, and the preview shortcut's
/// presses, which turn one setting.
///
/// **The run keeps one block of its own and no window in it.** A block is a place, and the tray stands in
/// the first of them for the reason every process of a run does: `the_facts` answers with it, and a
/// process with no block of its own would be answered with a viewer's. What it publishes is never
/// written there — the tray draws no window, so every fact it owns goes into the block the hovers are
/// served from (see `the_serving_facts`) or into every block standing (see `note_a_pin_press`).
///
/// The section a run opens is one of these, which is what makes a viewer's block the same address in
/// both processes (see `a_section_of_this_run`).
pub(crate) struct SeamFacts {
    viewers: [ViewerFacts; SEAM_VIEWERS],
    /// Messages the tray side sends the preview side, and the other way, each with an event of its own
    /// to wake the reader (see `SeamRing`).
    preview_ring: SeamRing,
    tray_ring: SeamRing,
    /// The pin key's main keys that are held, as a bit per chord slot.
    ///
    /// The run's rather than a window's because a hand is one hand: the hook that counts an edge reads
    /// it twice a message to tell a press from a repeat, and every viewer's band-hold reads the same
    /// answer (see `PIN_MAIN_DOWN`).
    pin_main_down: AtomicU32,
    /// Preview-shortcut presses seen over Explorer since the last drain: the setting they turn is the
    /// run's, so a press is one flip whichever preview loop reads it first (see `TOGGLE_PRESSES`).
    toggle_presses: AtomicU32,
    /// Wheel notches the run's one wheel hook has seen (see `WHEEL_TICKS`).
    wheel_ticks: AtomicU64,
}

// Safety: every field is an atomic or a block of them, and the two rings are written through their own
// claims (see `SeamRing::push`). Nothing here is a reference into this process, which is what makes the
// same type usable as a section.
unsafe impl Sync for SeamFacts {}

impl SeamFacts {
    /// Every fact of the run at rest: no block stood in, no pin up, nothing counted.
    const fn new() -> Self {
        Self {
            viewers: [const { ViewerFacts::new() }; SEAM_VIEWERS],
            preview_ring: SeamRing::new(),
            tray_ring: SeamRing::new(),
            pin_main_down: AtomicU32::new(0),
            toggle_presses: AtomicU32::new(0),
            wheel_ticks: AtomicU64::new(0),
        }
    }

    /// The block the viewer `pid` stands in, or `None` where no viewer of that pid is standing.
    fn facts_of(&'static self, pid: u32) -> Option<&'static ViewerFacts> {
        if pid == 0 {
            return None;
        }

        self.viewers
            .iter()
            .find(|block| block.owner.load(Ordering::Acquire) == pid)
    }

    /// The blocks a question about the run's pin is asked of, and the ones a number the run keeps for its
    /// windows is written into: every block a viewer stands in, or the one block a process publishes into
    /// where it is a viewer itself, or where nothing is claimed at all and it is a run drawing its own
    /// previews (see `asked_blocks`).
    fn blocks(
        &'static self,
        own: &'static ViewerFacts,
        a_viewer: bool,
    ) -> [Option<&'static ViewerFacts>; SEAM_VIEWERS] {
        let mut blocks = [None; SEAM_VIEWERS];

        if a_viewer {
            blocks[0] = Some(own);
            return blocks;
        }

        for (slot, block) in blocks.iter_mut().zip(self.viewers.iter()) {
            *slot = (block.owner.load(Ordering::Acquire) != 0).then_some(block);
        }

        // Nothing is claimed: this is a run of one process drawing its own previews, and the block it
        // publishes into is the one every question about the run is asked of (see `the_facts`).
        if blocks.iter().all(Option::is_none) {
            blocks[0] = Some(own);
        }

        blocks
    }

    /// Claim a free block of this run for `pid`, answering its place in `viewers`.
    ///
    /// **A block is claimed once and by one process**, with a compare and swap on its owner: two viewers
    /// begun at the same moment cannot stand in the same one, and the pid a block holds is what every
    /// other process asks it by (see `facts_of`). `None` is every block held by a running process, which
    /// a viewer's start-up is the one thing that can meet: it answers false to `join_the_seam_section`
    /// and stops, rather than standing as a process with no facts to draw from.
    ///
    /// **A block nothing is standing in is taken, whether or not it was given back.** Nothing releases a
    /// block when its viewer ends — a detached window's block ends with the window, and a viewer that dies
    /// ends without telling anyone — so the second pass takes one whose owner is no longer running. That is
    /// asked of the kernel rather than recorded anywhere (see `engine_processes::is_running`), and only on
    /// a viewer's start-up path, where the syscall costs nothing an eye can see.
    ///
    /// **And it is swapped away from the pid seen in it, not from zero**, which is what keeps a block out
    /// of two processes' hands: two viewers beginning together both find this block, and the swap fails for
    /// whichever of them comes second. A pid Windows has since handed to another process answers as
    /// running, and the block stays claimed — a block lost rather than a block shared.
    fn claim_a_block(&'static self, pid: u32) -> Option<usize> {
        for (at, block) in self.viewers.iter().enumerate() {
            if block
                .owner
                .compare_exchange(0, pid, Ordering::AcqRel, Ordering::Acquire)
                .is_ok()
            {
                return Some(at);
            }
        }

        for (at, block) in self.viewers.iter().enumerate() {
            let gone = block.owner.load(Ordering::Acquire);
            if gone == 0 || gone == pid || crate::app::engine_processes::is_running(gone) {
                continue;
            }

            if block
                .owner
                .compare_exchange(gone, pid, Ordering::AcqRel, Ordering::Acquire)
                .is_ok()
            {
                return Some(at);
            }
        }

        None
    }

    /// Whether a viewer begun now could stand in the run's facts: a block free, or one whose viewer has
    /// gone (see `claim_a_block`).
    ///
    /// Asked before a pin is detached into a window of its own, which is the one road on this side that
    /// needs a block the run does not already have: the window keeps the block it stands in, and the tray
    /// begins a viewer for the hovers that one leaves (see `pin_detach::detach_the_pinned_window`).
    fn a_block_can_be_had(&'static self) -> bool {
        self.viewers.iter().any(|block| {
            let owner = block.owner.load(Ordering::Acquire);
            owner == 0 || !crate::app::engine_processes::is_running(owner)
        })
    }
}

/// How many UTF-16 units of a pin's path the facts carry, which is how long a name a pin can be showing
/// there (see `publish_the_pinned_path`).
const SEAM_PIN_PATH_UNITS: usize = SEAM_TEXT_UNITS;

/// How many messages one direction's ring holds, which is how far a sender may be behind before a
/// message is dropped rather than kept.
const SEAM_RING_SLOTS: usize = 64;

/// How many UTF-16 units of text a slot carries, which is how long a path it can name (see
/// `SeamMessage`, and `send_preview` for the road a longer one takes: it stays in the process that made
/// it, where that process reads a channel of its own, and is a message the viewer never hears where it
/// does not — the tray, which is why the bound is named on the trace rather than left to be guessed).
const SEAM_TEXT_UNITS: usize = 512;

/// How many numbers a slot carries, which is as many as the widest variant writes.
const SEAM_NUMBERS: usize = 10;

/// How many boxes a standing pin's own buttons can answer a press in: the nine controls a sound's card
/// carries, and the room a caption's buttons take as one (see `pin_button_covers`).
pub(crate) const PIN_BUTTONS: usize = 10;

/// One direction's ring of fixed slots: messages claimed in one order and read in that same order.
///
/// It is a ring rather than a queue because it lives in a section two processes share, where nothing may
/// be allocated and no pointer may be handed across: the slots are the storage, and the two counters say
/// how far each end has got. **A full ring is a message lost, never a wait.** A sender is a low-level
/// hook callback as often as it is a tick, and a hook held until the other side read would stop every
/// wheel message on the desktop with it; so a claim that finds no room drops the message and traces the
/// loss (see `note_a_lost_message`). A ring this size over a pointer-paced hover is a loss that does not
/// happen.
///
/// The senders of a direction are not always one thread — a hover is sent from a hook as well as from
/// the pointer's own drag — so a slot is claimed with a compare and swap, and the payload is written
/// after the claim and before the slot is said to be filled. The reader is one thread per direction: the
/// preview loop for one, the hook loop for the other.
#[repr(C)]
struct SeamRing {
    /// Slots claimed so far, monotonic: the slot a claim takes is `claimed % SEAM_RING_SLOTS`.
    claimed: AtomicU32,
    /// Slots read so far, monotonic. `claimed - read` is how full the ring is.
    read: AtomicU32,
    /// Per slot, one more than the claim that wrote it, or zero while it holds nothing: a reader that
    /// finds the number it is waiting for knows the payload under it is whole.
    filled: [AtomicU32; SEAM_RING_SLOTS],
    /// The messages themselves, written through a claim rather than through a `&self`.
    slots: [UnsafeCell<SeamMessage>; SEAM_RING_SLOTS],
}

impl SeamRing {
    const fn new() -> Self {
        Self {
            claimed: AtomicU32::new(0),
            read: AtomicU32::new(0),
            filled: [const { AtomicU32::new(0) }; SEAM_RING_SLOTS],
            slots: [const { UnsafeCell::new(SeamMessage::EMPTY) }; SEAM_RING_SLOTS],
        }
    }

    /// Put `message` in the next free slot, or answer `false` where the ring was full (see the type).
    fn push(&self, message: SeamMessage, direction: &str) -> bool {
        let mut claimed = self.claimed.load(Ordering::Relaxed);
        loop {
            let read = self.read.load(Ordering::Acquire);
            if claimed.wrapping_sub(read) as usize >= SEAM_RING_SLOTS {
                note_a_lost_message(direction);
                return false;
            }

            match self.claimed.compare_exchange_weak(
                claimed,
                claimed.wrapping_add(1),
                Ordering::AcqRel,
                Ordering::Relaxed,
            ) {
                Ok(_) => break,
                Err(actual) => claimed = actual,
            }
        }

        let slot = claimed as usize % SEAM_RING_SLOTS;
        // Safety: the claim above is this sender's own, so no other sender writes this slot until the
        // ring has been read a whole turn past it, and the reader reads it only once `filled` names it.
        unsafe {
            self.slots[slot].get().write(message);
        }
        // The write is what makes the slot whole: nothing reads it before this names it.
        // shortcut: after 2^32 messages the claim that fills a slot wraps onto the value a reader reads as
        // empty; upgrade if one direction of a session ever carries that many.
        self.filled[slot].store(claimed.wrapping_add(1), Ordering::Release);
        true
    }

    /// Take the next message, or `None` where there is nothing to read yet.
    fn pop(&self) -> Option<SeamMessage> {
        let read = self.read.load(Ordering::Relaxed);
        let slot = read as usize % SEAM_RING_SLOTS;
        if self.filled[slot].load(Ordering::Acquire) != read.wrapping_add(1) {
            return None;
        }

        // Safety: `filled` names this claim, so the payload under it is whole, and no sender touches the
        // slot again until the ring has been read a whole turn past it.
        let message = unsafe { self.slots[slot].get().read() };
        self.read.store(read.wrapping_add(1), Ordering::Release);
        Some(message)
    }
}

/// The facts of a run that has no section: this process's own copy, which is where every test and every
/// start-up that could not open one publishes.
static LOCAL_FACTS: SeamFacts = SeamFacts::new();

/// Where a run's facts are once it has opened a section, or null where it has not.
static MAPPED_FACTS: AtomicPtr<SeamFacts> = AtomicPtr::new(std::ptr::null_mut());

/// Which block of this run's facts this process stands in, as its place in `SeamFacts::viewers`, or
/// `NO_BLOCK` while it has claimed none.
static THE_OWN_BLOCK: AtomicUsize = AtomicUsize::new(NO_BLOCK);

/// The one block a run's facts hold on their own is the first: where a block is asked for and none was
/// claimed, which is every test and every run that draws its own previews (see `THE_OWN_BLOCK`).
const NO_BLOCK: usize = usize::MAX;

/// The handle to this run's section, kept open for the life of the process.
///
/// **A view does not keep the name; the handle does.** The section is backed by the pagefile — no file
/// behind it — and an object of that kind is deleted the moment its last handle is closed, whatever views of
/// it are still mapped. A run that closed the handle would go on reading and writing the very same pages, so
/// nothing would look wrong from inside it, while the name every other process is told to open had gone: a
/// viewer cannot join, draws nothing at all, and a tray that had already made the rings the road publishes a
/// whole session of hovers into a section no one can reach (see `app::viewer` and `main`).
///
/// So it is never closed, and it needs no more than that: the last handle going is the run ending.
static THE_SECTION_HANDLE: AtomicIsize = AtomicIsize::new(0);

/// The facts this process publishes to and reads from: the block it stands in, of the section where
/// there is one and of this process's own copy otherwise.
///
/// Written once, at start-up, before the threads that publish are begun; read on every hook callback
/// and every tick, which is why it is one pointer load and one index and not a lock.
pub(crate) fn the_facts() -> &'static ViewerFacts {
    let run = the_run_facts();
    let own = THE_OWN_BLOCK.load(Ordering::Acquire);

    match run.viewers.get(own) {
        Some(block) => block,
        // Nothing claimed a block and no block is this process's: the run is one process drawing its
        // own previews, and the first block is the one it stands in.
        None => &run.viewers[0],
    }
}

/// The block the hovers of this run are served from: the serving viewer's, and this process's own where
/// it is the one drawing the previews (see `the_viewer_process`).
///
/// **It is the answer to every question about a preview under the pointer**, and the facts a `Fact`
/// names below are read and written in it: an asker inside a viewer is answered with its own block,
/// because `the_viewer_process` is that process's own pid where no viewer was noted — so a viewer that
/// has become a window of its own still draws, wheels and keys its own preview and only stops being
/// asked what the pointer is on.
fn the_serving_facts() -> &'static ViewerFacts {
    facts_of(the_viewer_process()).unwrap_or_else(the_facts)
}

/// The block of the viewer `pid` stands in, or `None` where no viewer of that pid is standing.
///
/// One pointer load and a walk of the blocks, and no lock: it is asked from the hooks, which may take
/// none (see `the_facts`).
pub(crate) fn facts_of(pid: u32) -> Option<&'static ViewerFacts> {
    the_run_facts().facts_of(pid)
}

/// The blocks of this run a question about the pin, or about a key of the pin's, is asked of: every one
/// standing, or the one block of a process that is itself a viewer.
///
/// **A pin is one window of this run wherever it stands**, so the tray asks every block before it
/// answers that no pin is up, that no button is under the press or that the foreground is no pin's. A
/// viewer asks about its own window and no other, which is why it is answered with its own block alone.
fn asked_blocks() -> [Option<&'static ViewerFacts>; SEAM_VIEWERS] {
    the_run_facts().blocks(the_facts(), this_process_is_a_viewer())
}

/// Whether this process is one of the run's viewers rather than the run itself: it joined a section a
/// tray had opened, rather than opening one (see `join_the_seam_section`).
fn this_process_is_a_viewer() -> bool {
    THE_RUN.load(Ordering::Acquire) != 0
}

/// The run's facts: the section where this process has opened or joined one, this process's own copy
/// otherwise.
///
/// Written once, at start-up, before the threads that publish are begun; read on every hook callback
/// and every tick, which is why it is one pointer load and not a lock.
fn the_run_facts() -> &'static SeamFacts {
    let mapped = MAPPED_FACTS.load(Ordering::Acquire);
    if mapped.is_null() {
        return &LOCAL_FACTS;
    }

    // Safety: the pointer is set once, from a view this process mapped and never unmapped, and the
    // section is a whole `SeamFacts` written through `SeamFacts` and nothing else.
    unsafe { &*mapped }
}

/// Claim a block of this run's facts for this process, and answer it (see `SeamFacts::claim_a_block`).
fn claim_a_block() -> Option<&'static ViewerFacts> {
    let at = the_run_facts().claim_a_block(std::process::id())?;
    THE_OWN_BLOCK.store(at, Ordering::Release);
    Some(&the_run_facts().viewers[at])
}

/// Whether a viewer begun now could stand in the run's facts (see `SeamFacts::a_block_can_be_had`).
pub(crate) fn a_block_can_be_had() -> bool {
    the_run_facts().a_block_can_be_had()
}

/// Count one press of the pin key, in every block the run has standing: what a press means is each
/// window's own, so the tray counts it for all of them and only the blocks a viewer stands in (see
/// `PIN_PRESSES`).
pub(crate) fn note_a_pin_press() {
    count_a_pin_press_in(asked_blocks());
}

/// Count one press of the pin key in `blocks` (see `note_a_pin_press`).
fn count_a_pin_press_in(blocks: [Option<&'static ViewerFacts>; SEAM_VIEWERS]) {
    for block in blocks.iter().flatten() {
        PIN_PRESSES.in_block(block).fetch_add(1, Ordering::AcqRel);
    }
}

/// Ask the pin to come down, in every block the run has standing: the pin is one window of the run's
/// wherever it stands, and only the viewer holding it acts on this (see `PIN_END_REQUESTED`).
pub(crate) fn ask_the_pin_to_come_down() {
    ask_the_pin_to_come_down_in(asked_blocks());
}

/// Ask the pin to come down in `blocks` (see `ask_the_pin_to_come_down`).
fn ask_the_pin_to_come_down_in(blocks: [Option<&'static ViewerFacts>; SEAM_VIEWERS]) {
    for block in blocks.iter().flatten() {
        PIN_END_REQUESTED
            .in_block(block)
            .store(true, Ordering::Release);
    }
}

/// Ask the pin's menu to come down, in every block the run has standing (see
/// `PIN_MENU_DISMISS_REQUESTED`).
pub(crate) fn ask_the_pin_menu_to_come_down() {
    for block in asked_blocks().iter().flatten() {
        PIN_MENU_DISMISS_REQUESTED
            .in_block(block)
            .store(true, Ordering::Release);
    }
}

/// Open the shared section this run's facts live in, and publish into that from here on.
///
/// **The section is named for this run, pid and all** (see `the_seam_name`), so a second run opens one of
/// its own rather than reading this one's hovers, and nothing is torn down when it ends: the section goes
/// with the last view of it.
///
/// Called once, at start-up, before the threads that publish are begun. It answers whether a section was
/// had: a run without one has nothing a second process could read, and is therefore a run that draws its
/// own previews (`main` is where that is decided). Nothing else here is a condition for a run to work — a
/// section that could not be made leaves `the_facts` on this process's own copy, which is what every test
/// and every run before the split used.
pub(crate) fn open_the_seam_section() -> bool {
    // The two events the messages are announced on exist before the viewer does, so neither side ever
    // makes one the other has already waited on: a named event is one object whichever process asked for
    // it first, and all that matters is that it exists before the first send signals it. The readiness
    // event is not one of them: it is named for the viewer that sets it, and is made by whichever side
    // asks for it first (see `preview_ready_event`).
    let _ = preview_channel_event();
    let _ = tray_channel_event();

    let Some(facts) = a_section_of_this_run() else {
        return false;
    };
    MAPPED_FACTS.store(facts, Ordering::Release);
    // A run with a section stands in a block of it, and this is the first block claimed there: the one
    // the tray stands in until a viewer of its own claims the next (see `claim_a_block`).
    let _ = claim_a_block();
    true
}

/// Join the section the tray app opened, and publish into that from here on.
///
/// **A viewer never makes one.** The name carries the tray's pid, which is the one thing a viewer is told
/// on its command line (see `the_seam_name`), and a section made here instead of mapped would be a second
/// copy of the facts with nothing publishing into it: two sides reading boxes that were never written.
/// Answers whether the tray's section was found; a viewer that cannot find one has nothing to read and
/// nothing to be told, and stops rather than standing as a process that draws nothing.
///
/// Called once, at the start of a viewer, before any thread of it that publishes is begun.
pub(crate) fn join_the_seam_section(tray: u32) -> bool {
    THE_RUN.store(tray, Ordering::Release);
    // The events beside the section: asked for here so that a send which finds nobody waiting yet is still
    // a signal on the object the other side will read. The readiness event is this process's own and is
    // made where it is set (see `preview_ready_event` and `note_the_preview_loop_is_up`).
    let _ = preview_channel_event();
    let _ = tray_channel_event();

    let name = the_seam_name("");
    let size = size_of::<SeamFacts>();

    // Safety: opening by name a section this run's own tray made, for read and write. Nothing is created
    // here, and a name nothing holds is answered as an error rather than as an empty section.
    let Ok(mapped) =
        (unsafe { OpenFileMappingW(FILE_MAP_ALL_ACCESS.0, false, PCWSTR(name.as_ptr())) })
    else {
        return false;
    };

    // Safety: a whole view of the section the handle names.
    let view = unsafe { MapViewOfFile(mapped, FILE_MAP_ALL_ACCESS, 0, 0, size) };
    // The handle is kept rather than closed, for the reason the creator keeps its own (see
    // `THE_SECTION_HANDLE`).
    THE_SECTION_HANDLE.store(mapped.0 as isize, Ordering::Release);
    if view.Value.is_null() {
        return false;
    }

    MAPPED_FACTS.store(view.Value.cast::<SeamFacts>(), Ordering::Release);
    // The block this viewer draws in, claimed before any thread of it that publishes is begun: a block
    // another viewer of the run already stands in is not one this one may write (see `claim_a_block`).
    if claim_a_block().is_some() {
        return true;
    }

    // Said out loud rather than only answered: a viewer that stops here draws nothing at all, and a run
    // whose blocks are all stood in has no room for the previews it keeps asking for.
    eprintln!("seam: every block of the run's facts is stood in; this viewer stops");
    false
}

/// This run's section, mapped and written with the at-rest facts, or `None` where there is none to be had.
///
/// What it holds is one `SeamFacts`, atoms and two rings, so a viewer built with this run agrees about
/// where a field starts and reads a fact in one load. The creator writes the at-rest facts into it before
/// anything is told where it is, which is what makes a box that was never published read as no box rather
/// than as one at the origin.
fn a_section_of_this_run() -> Option<*mut SeamFacts> {
    let name = the_seam_name("");
    let size = size_of::<SeamFacts>();

    // Safety: a section backed by the pagefile, named for this run and sized to one `SeamFacts`; no
    // security descriptor, which is this user's own access, and no handle of an existing file.
    let Ok(mapping) = (unsafe {
        CreateFileMappingW(
            INVALID_HANDLE_VALUE,
            None,
            PAGE_READWRITE,
            0,
            size as u32,
            PCWSTR(name.as_ptr()),
        )
    }) else {
        return None;
    };

    // Safety: a whole view of the section just made, mapped read and write for this process.
    let view = unsafe { MapViewOfFile(mapping, FILE_MAP_ALL_ACCESS, 0, 0, size) };
    // **The handle is kept rather than closed, and the name goes with it** (see `THE_SECTION_HANDLE`).
    THE_SECTION_HANDLE.store(mapping.0 as isize, Ordering::Release);
    if view.Value.is_null() {
        return None;
    }

    let facts = view.Value.cast::<SeamFacts>();
    // Safety: the view is the whole of a `SeamFacts`, freshly mapped, and nothing else has been told
    // where it is yet, so writing the at-rest facts into it cannot race a reader.
    //
    // shortcut: the section exists the moment it is made and the facts are written a call later, so a
    // reader that found the name in that window would read zeroes; nothing does, because this runs before
    // any thread of this process is begun. Upgrade to a section nothing may name until it is written if a
    // second process ever opens one while the first is still starting.
    unsafe {
        facts.write(SeamFacts::new());
    }

    Some(facts)
}

/// One published fact, named rather than kept: which field of a block it is, asked of the block the
/// hovers are served from (see `ViewerFacts` and `the_serving_facts`).
///
/// It is a function from a block to a field, so a handle cannot name a field that is not there and
/// cannot name one of another type — the compiler checks both, and there is no index to be got wrong at
/// run time. What the indirection costs is one cached load and one predicted call, which is what a hook
/// callback can afford: the atomics themselves are the atomics they were, and a handle is used exactly
/// as the static it replaced was, by dereferencing to it.
///
/// **Which block a handle answers with is the whole of what a fact is about.** A handle reads and writes
/// the block the hovers are served from, which is the asker's own block inside a viewer (see
/// `the_serving_facts`); the facts a caller needs of *any* window of the run — a pin, wherever it
/// stands — are asked of `asked_blocks` instead, and the ones a hook asks of one particular window are
/// read through `in_block`.
pub(crate) struct Fact<T: 'static> {
    home: fn(&'static ViewerFacts) -> &'static T,
}

impl<T: 'static> Fact<T> {
    const fn of(home: fn(&'static ViewerFacts) -> &'static T) -> Self {
        Self { home }
    }

    /// This field of the block `block` is, for a caller that has settled which window's fact it means —
    /// a notch's, turned in a box one viewer published, or a key's, whose caret one viewer holds (see
    /// `PublishedRegion::the_block_over` and `the_pin_of_the_window`).
    pub(crate) fn in_block(&self, block: &'static ViewerFacts) -> &'static T {
        (self.home)(block)
    }
}

impl<T: 'static> Deref for Fact<T> {
    type Target = T;

    fn deref(&self) -> &T {
        (self.home)(the_serving_facts())
    }
}

/// One fact of the run's own rather than of a window's, named rather than kept: which field of `SeamFacts`
/// it is (see `ViewerFacts` for why a field is named at all).
///
/// It is a type of its own rather than a `Fact` because what it names is not in a block: a count here is
/// kept once for the run, and a handle that could name a field of either would be a handle whose reader
/// cannot tell which window's answer it is about to be given.
pub(crate) struct RunFact<T: 'static> {
    home: fn(&'static SeamFacts) -> &'static T,
}

impl<T: 'static> RunFact<T> {
    const fn of(home: fn(&'static SeamFacts) -> &'static T) -> Self {
        Self { home }
    }
}

impl<T: 'static> Deref for RunFact<T> {
    type Target = T;

    fn deref(&self) -> &T {
        (self.home)(the_run_facts())
    }
}

#[derive(Clone)]
pub(crate) enum PreviewMessage {
    /// A preview of the file the pointer hovers, opened from the cursor it was
    /// hovered at.
    ///
    /// The region that comes with it is what the `Avoid` setting measured off the
    /// hovered item — its name, and the columns beside it at `Avoid Details` — which
    /// the placement is kept off at either way of avoiding, and whether that region is
    /// a column of the item's view, which decides *how* a placement is kept off it: a
    /// column is stepped off to one of its sides, an item's own text in whichever
    /// direction asks the least. See `AvoidRegion` and `avoiding_text`. It is `None`
    /// when the setting is off, when the view reported no text for the item, or when
    /// the walk found no item at the cursor.
    Show(PathBuf, i32, i32, Option<AvoidRegion>),
    /// A preview of the focused item, whose box comes with it, and the region the
    /// `Avoid` setting keeps it off — the item's own text, or the name alone, or the
    /// column the name sits in, depending on the way the setting is on. That region is
    /// where its preview is placed from as well as what it is kept off, the way a
    /// hovered item's is, for an item that draws its text as a row of its view: the
    /// last field says whether it does, read off the item's own text by the hook (see
    /// `explorer_hook::ItemText`). See `compute_keyboard_layout`.
    ShowKeyboard(PathBuf, i32, i32, i32, i32, Option<ScreenRegion>, bool),
    Hide,
    Refresh,
    /// A preview type was switched on or off. Only a preview whose own kind is
    /// now off is rebuilt, and it is rebuilt from the hover it came from, so it
    /// goes away on the spot rather than at the next pointer move.
    RefreshTypes,
    /// A switch that ends the engines behind a kind of preview has been turned off, in the
    /// tray's `Preview Types` submenu or by the engine choice moving to LibreOffice.
    ///
    /// The processes are the viewer's: it started them, they are its children, and only it can
    /// end them. The tray asks rather than reaching into a process that has no engine in it (see
    /// `office_render::stop_engines`).
    StopEngines,
    /// The page the browser engine is drawing has been switched off — the `Render HTML` row, or
    /// the `Text` kind that page rides — and comes down with the switch it was asked for by.
    ///
    /// Only the viewer can take it down: the page is a window of its engine, and whether it is a
    /// hover's page or a pin's is a question about what is on its screen (see `hide_html_page`).
    HideHtmlPage,
    /// A disk or memory budget for what previews leave behind has been made smaller, and what is
    /// over the new one goes now rather than at the next render that happens to pass through a
    /// cache.
    ///
    /// The folders are the machine's, but what may not be given up is not: a page a hover is
    /// waiting for is held by the viewer, and a trim that ran in the tray would see no page held
    /// at all and give up the very one on its way (see `document_cache::trim_now`).
    TrimCaches,
    /// Nothing is hovered and no pin is up: a player that outlived the hover it was started for
    /// is ended, which is a question only the process that spawned it can answer — the pid is
    /// that process's own record (see `ffplay_window::kill_stray_video_process`).
    KillStrayPlayer,
    /// The pin key was pressed while a preview was on screen: what is up becomes a
    /// window of its own — captioned, movable, always on top — and the hover machinery
    /// stays quiet behind it until it is closed. The rect is the box the preview is
    /// already in, so nothing about the picture moves when it is pinned.
    ///
    /// The same message installs the file a pin that is up is shown next, and `arriving` says
    /// which of the two it is: the take-up out of a hover is the pin *arriving*, and it comes
    /// with no pin standing because that is what arriving is, while a file stepping onto a pin
    /// that is up is one the pin can have ended before the install landed — the watchdog ends a
    /// pin from a thread of its own (see `pin_lifecycle::the_pin_takes_the_file_up`).
    Pin {
        path: PathBuf,
        rect: ScreenRegion,
        arriving: bool,
    },
    /// The pinned window was given another box — maximized, restored, or resized by an
    /// edge — and the media in it is laid out again for that box. It is the same
    /// question a hover asks, with the box already answered.
    PinBox(ScreenRegion),
    /// The file the user picked while a preview was pinned, which the pin is asked to show
    /// instead of the one it has: the window keeps its place and its size, and the media in it
    /// becomes the new file's, fitted to the box the pin already has (see `Pin Mode → Update
    /// Preview`). The Explorer hook sends it — for a click, for a key, and — where the setting
    /// asks for it — for the pointer settling on another file (see `update_pinned_preview`).
    PinUpdate(PathBuf),
    /// A press landed on the listing behind the pin, before the file it picked is known: the
    /// click activates the listing, which puts it over the pin for as long as the shell takes to
    /// answer for the file, so the pin comes forward here rather than after that answer. `window`
    /// is the frame of the listing the press was made in, and it travels because the activation
    /// that lands after this loop reads the press is *that* window being put over the pin: the
    /// answer to it is asked of the window, not of whatever is in front by then (see
    /// `PinFrontOwed`). The answer is a [`PinUpdate`](Self::PinUpdate); a press that names no file
    /// is put back behind the listing by [`PinNoPick`](Self::PinNoPick) (see
    /// `press_pinned_preview`).
    PinPress {
        window: isize,
    },
    /// A press that was read as a pick named no file — a folder, a file no kind previews, the gap
    /// between the rows, the toolbar above them — and the pin goes back behind the window the press
    /// was made in, which is the frame `window` names: the foreground is still the window the user
    /// came from where the click has not been dispatched yet. What the press above raised is undone
    /// here; a pick that lands after it still raises the pin again (see
    /// `pin_press_was_not_a_pick`).
    PinNoPick {
        window: isize,
    },
    /// Whether the pin key is watched was changed in the tray, or previews themselves
    /// were turned off. A pin on screen is ended here rather than left as a window
    /// nothing would ever take down again.
    PinChanged,
    /// One of the tray's two `Volume` rows was moved, and the one it is: the film's half or
    /// the sound's. A pin already playing at a level of its own is told the new one here,
    /// on the thread its player belongs to, rather than being left at what it was taken up
    /// at (see `publish_tray_volume`).
    TrayVolume {
        audio: bool,
    },
    /// The tray's `Text → Font Size` row was picked, and the size it names. A pin already showing a
    /// page is put to that size on the thread the page is drawn on — the reader's own size on it
    /// given up, the way closing the pin gives it up — while a hover of a page is rebuilt from its
    /// own load at the new size and is owed nothing here (see `publish_text_font_size` and
    /// `redraw_pinned_page`).
    TrayTextSize {
        percent: u32,
    },
    /// Every probed geometry is given up: the engine choice moved, and a geometry read by FFprobe is
    /// not the geometry a media-engine preview wants, so the next hover probes again. Given up on the
    /// loop because the cache is that side's, and a row moved in the tray is not one that reaches a
    /// preview already up — the box it is being placed by has to stand until the next one begins (see
    /// `forget_video_geometry`).
    ForgetVideoGeometry,
    /// Only the extraction of `path` is kept running, every other one being dropped; `None` drops all
    /// of them. Sent by the tray's `Subtitles` switch, to drop an extraction running for a film no
    /// subtitle will be drawn from (see `keep_extraction_for`).
    KeepExtraction {
        path: Option<PathBuf>,
    },
    /// The render tier is done with a document: a page is waiting in the cache,
    /// or there is no page. The generation is the hover that asked for it, so a
    /// render landing after the pointer has moved on is ignored — the page is
    /// still cached for the next hover either way.
    OfficeRenderReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// A video's probe is done: the geometry is waiting in the cache, or the answer is
    /// that there is none. The generation is the hover that was waiting on it, so a probe
    /// landing after the pointer has moved on is ignored — the answer is held for the next
    /// hover either way (see `video_probe`).
    VideoProbed {
        path: PathBuf,
        generation: u64,
    },
    /// A measure that reads a file is done: the box it answers with is held, or the answer is
    /// that the reader has none for the file — which is not a wait that can be answered, so it
    /// is the one answer a wait comes down on (see `measured_off_the_tick`).
    ///
    /// What was waiting on it is the hover that is on screen, and that is the whole of what
    /// this answer has to be matched to: a box is measured per version of the file, and a hover
    /// of a file whose version has changed is measured again rather than answered with the box
    /// of the version before it.
    MeasureProbed {
        path: PathBuf,
        size: Option<(u32, u32)>,
    },
    /// The planner is done with a question the pin asked — a walk along the folder, or the
    /// name of the program a file would open with. Neither question can be asked on the
    /// thread that draws the pin, so both are asked elsewhere and answered here (see
    /// `PinPlanner`).
    ///
    /// A walk is answered whatever it found: a folder that could not be read and a folder
    /// with nothing in it to step onto are both answers, not the absence of one, so the pin's
    /// arc comes down when they land rather than waiting out a bound (see `answer_pin_job`).
    PinAnswered(PinPlanned),
    /// The band under a pinned sound turned over, so the walk behind it is asked again: how wide
    /// that walk is depends on the band, and nothing else re-reads the mode now that it has changed
    /// (see `pin_playlist::toggle` and `pin_walk::ask_the_walk_again`).
    PinWalkAgain,
    /// The tray's `Restore Pin Mode` row was clicked: the pin that is up is to be brought back over
    /// everything, whatever state it is in — put away as its bubble, minimized into the taskbar, or
    /// left standing behind the program the user moved on to (see `restore_pinned_preview`).
    ///
    /// It is a message rather than a call because the pin is the preview thread's: the window it
    /// stands in, the rank it is owed and the caret it is given are all that thread's to ask, and a
    /// tray answering a menu row is not on it.
    PinRestore,
    /// The small subtitle files a film's own tracks were copied into are ready, sent the
    /// moment the extraction's one pass finishes (see `subtitle_files`).
    ///
    /// Only a pinned window acts on it: a pin has its film on screen already, so what the
    /// copy landing changes is what its *next* relaunch is drawn with — the copy is preloaded
    /// for that rather than the picture being changed hands for a subtitle (see
    /// `reload_pinned_subtitles`). A hover is deliberately not begun again under the pointer,
    /// and this message is not for one (see the note on the first hover in
    /// `video_launch::subtitle_filter`).
    VideoSubtitlesReady(PathBuf),
    /// The ImageMagick engine is done with a file: the picture it developed is in hand, or
    /// there is none — a file it cannot read is remembered as one it will not draw. The
    /// generation is the hover that was waiting on it, so a conversion landing after the
    /// pointer has moved on is ignored; the picture itself is held for the hover that asked
    /// either way (see `magick_render_is_due`).
    MagickReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// The PeaZip engine is done with a file: the archive's table of contents is in hand, read
    /// into the listing cache, or there is none — a file it cannot open is remembered as one it
    /// will not list. The generation is the hover that was waiting on it, so a listing landing
    /// after the pointer has moved on is ignored; the listing itself is held under the file's own
    /// key either way, so the next hover of it is a read rather than a launch (see
    /// `peazip_render_is_due`).
    PeazipReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
}

/// The pid this run's seam is named for: this process's own for the tray app, the tray's for a viewer it
/// began. Zero until a run says otherwise, which is what every test and every single-process start is.
static THE_RUN: AtomicU32 = AtomicU32::new(0);

/// The pid this run's seam is named for (see `THE_RUN`).
fn the_run_pid() -> u32 {
    let run = THE_RUN.load(Ordering::Acquire);
    if run != 0 {
        return run;
    }

    // Safety: asks the kernel for this process's own id; nothing is dereferenced.
    unsafe { GetCurrentProcessId() }
}

/// The name a piece of this run's seam is known by, as the UTF-16 the Win32 calls take.
///
/// **The run's own pid is in it, and that is the whole of why two runs never meet.** A kernel object's
/// name belongs to the session rather than to the process, so a fixed one would be one section and one
/// pair of events for every run at once: a second run would open the first one's facts and publish its
/// own hovers into them, and a run a developer left behind would take the place of the user's. A pid is
/// the one thing a run has and no other run has, and the `Local\` prefix keeps even that within the
/// session the run was begun in.
///
/// A viewer is the exception that proves it: it is a process of its own whose facts are the tray's, so it
/// is named for the tray's pid and joins what the tray opened rather than opening anything.
///
/// Nothing outlives the run that made it: a section and an event live exactly as long as the last
/// handle to them, so a run that ends takes its names with it.
fn the_seam_name(suffix: &str) -> Vec<u16> {
    let pid = the_run_pid();
    format!("Local\\rust-hover-preview-seam-{pid}{suffix}")
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect()
}

/// The event the preview channel's sends are announced on, so that a thread waiting for one
/// sleeps on the channel as well as on the message queue.
///
/// The wait the loop does was a wait on the queue alone, polled in eight-millisecond slices
/// because nothing would otherwise wake it: a thread with nothing on screen spent the whole of
/// `IDLE_WAIT_MS` waking to ask an empty channel whether it had anything, several times a
/// second, for the whole life of the process. Signalling this on the send is what lets the
/// wait block on the channel too — so a hover is answered when it is sent rather than at the
/// end of a slice, and an idle loop costs one kernel wait per tick instead of one per slice.
pub(crate) fn preview_channel_event() -> HANDLE {
    // Kept as the number a handle is rather than as the handle, for the reason
    // `NOACTIVATE_WAKE` is: a handle is a raw pointer, and this is shared across threads.
    static EVENT: Lazy<isize> = Lazy::new(|| {
        let name = the_seam_name("-preview");
        // Safety: an event with a name of this run's own and no security descriptor, manual-reset,
        // which is what a wait that may be woken several times before it is read wants.
        unsafe { CreateEventW(None, true, false, PCWSTR(name.as_ptr())) }
            .map(|handle| handle.0 as isize)
            .unwrap_or_default()
    });

    HANDLE(*EVENT as *mut core::ffi::c_void)
}

/// Wake the preview loop where it stands, without a message: what a wheel notch over a pinned
/// picture, or a pin button press, asks for.
///
/// A notch or a press is counted on a hook's own thread — an atomic written there and nothing else —
/// and read by the loop on its next tick, so a loop waiting for a pin is a loop that will not see it
/// until its own wait runs out: `STATIC_PIN_WAIT_MS` rounded up to the system timer's own 15.6 ms
/// tick, which measured 62.2 ms over forty waits, worst 63.7 — four frames of a 144 Hz display, and
/// the whole of what makes a zoom start unevenly. The same wait signalled where the notch is counted
/// came back in 5.4 ms with the signal sent five in. A middle click shorter than that wait would
/// otherwise be lost outright.
///
/// It is the channel's own event rather than one of its own, so nothing new is waited on: the loop is
/// already armed on it, and a signal to an event nobody is waiting on is read by the next wait — it is
/// manual-reset, and reset after the wait rather than at the send (see `wait_preview_channel`).
pub(crate) fn wake_preview_loop() {
    // Safety: an event created once for the process, signalled: nothing here is dereferenced.
    unsafe {
        let _ = SetEvent(preview_channel_event());
    }
}

/// The event the viewer `viewer` sets once its preview loop is running, and the one the tray app waits
/// on before it makes the rings the road (see `app::viewer`).
///
/// **It is named for the viewer and not for the run**, because a run may have more than one of them
/// standing: a viewer whose pin has become a window of its own turns its own loop for the rest of the
/// session, so one manual-reset event of the run's own could not say which of the two answered (see
/// `app::detach`). Whoever asks for the name first makes the object — a name that exists already answers
/// with the object it names, signal and all — and the caller owns the handle: it is closed by the side
/// that asked only to wait on it, and kept for the life of the process by the viewer that sets it, an
/// event whose last handle closes being a name nothing holds (see `note_the_preview_loop_is_up`).
///
/// **It is what keeps a viewer that dies at its start-up from costing the user every preview.** The tray
/// commits to the rings on the strength of a process it has just begun, and a process is begun whether or
/// not it ever stands: without an answer here, a run whose viewer could not join the section would publish
/// every hover of the session into a ring nobody reads, which is a run with no previews at all. So a viewer
/// sets this once its section is joined and its loop is turning, and a tray that does not hear it in time
/// draws the run's previews itself (see `viewer::begin_the_viewer`).
pub(crate) fn preview_ready_event(viewer: u32) -> HANDLE {
    let name = the_seam_name(&format!("-up-{viewer}"));
    // Safety: an event with a name of this run's own and no security descriptor, manual-reset, which is
    // what a signal set before the tray's wait is armed wants.
    unsafe { CreateEventW(None, true, false, PCWSTR(name.as_ptr())) }.unwrap_or_default()
}

/// Say that this process's preview loop is running: the section is joined, the loop is turning, and every
/// message the tray publishes from here on is one this process reads (see `preview_ready_event`).
pub(crate) fn note_the_preview_loop_is_up() {
    // The handle is kept for the life of this process rather than closed after the signal: a viewer that
    // sets its event and drops it leaves a name nothing holds, and a tray that asked for the name
    // afterwards would make a fresh one and read its own silence as a viewer that never came up (see
    // `preview_ready_event`).
    static READY: Lazy<isize> = Lazy::new(|| preview_ready_event(std::process::id()).0 as isize);

    // Safety: an event of this run's own, signalled: nothing here is dereferenced.
    unsafe {
        let _ = SetEvent(HANDLE(*READY as *mut core::ffi::c_void));
    }

    // And one line from the side the tray is waiting on, on the trace a run asks for: a run whose previews
    // are missing is told here whether the process that was to draw them ever stood. Only a viewer's rings
    // carry messages in a run that draws in another process, so this is said by a viewer and by nothing
    // else (see `main`).
    if the_rings_carry_messages() && std::env::var_os("RHP_SEAM_TRACE").is_some() {
        eprintln!("seam: the viewer is up, its preview loop running");
    }
}

/// Whether this process's messages cross the rings rather than the channels its own threads share.
///
/// Set once, at start-up, by whichever side is one of two — the tray app once its viewer is standing, the
/// viewer as it begins — and never moved again: there is no road back, and a run that has no viewer takes
/// the channels from the first moment and keeps them (see `main`). It is where the split is decided, and
/// it is one atomic load for every send and every wait.
static THE_RINGS_CARRY_MESSAGES: AtomicBool = AtomicBool::new(false);

/// Make the rings the road the messages take, in both directions, from here on (see
/// `THE_RINGS_CARRY_MESSAGES`).
pub(crate) fn carry_messages_in_the_rings() {
    THE_RINGS_CARRY_MESSAGES.store(true, Ordering::Release);
}

fn the_rings_carry_messages() -> bool {
    THE_RINGS_CARRY_MESSAGES.load(Ordering::Acquire)
}

/// Whether this run's messages cross the rings at all, which is the road it took at start-up (see
/// `THE_RINGS_CARRY_MESSAGES`). A run that draws its own previews never sends one anywhere, so there is no
/// viewer for it to owe one to.
pub(crate) fn the_rings_are_the_road() -> bool {
    the_rings_carry_messages()
}

/// Whether the preview ring is this process's to read and to write, right now: the rings are the road,
/// and this process is still in the room the ring belongs to.
///
/// It is the second half that moves. A viewer whose pin has become a window of its own leaves the room —
/// every hover the ring carries from then on belongs to the viewer that took its place — and the messages
/// of its own that used to cross go back to being the channel its own threads share (see `app::detach`).
fn the_preview_ring_is_ours() -> bool {
    the_rings_carry_messages() && !crate::app::detach::is_detached()
}

/// Whether this process has left the room the rings belong to: a viewer that has become a window of its
/// own, whose own channel is its road from then on (see `the_preview_ring_is_ours` and
/// `wait_preview_channel`).
fn the_preview_ring_was_left() -> bool {
    the_rings_carry_messages() && crate::app::detach::is_detached()
}

/// Whether this message is one only a viewer drawing something can answer: a hover, the keyboard's hover,
/// or a pin taken up from either.
fn a_show(message: &PreviewMessage) -> bool {
    matches!(
        message,
        PreviewMessage::Show(..) | PreviewMessage::ShowKeyboard(..) | PreviewMessage::Pin { .. }
    )
}

/// Send a message to the preview loop and wake it, so that a wait in progress ends at once.
pub(crate) fn send_preview(message: PreviewMessage) {
    if the_preview_ring_is_ours() {
        // A viewer that has gone is begun again before the show that would have been drawn by nobody is
        // put in the ring (see `app::viewer`).
        if a_show(&message) {
            crate::app::viewer::make_sure_a_viewer_is_standing();
        }

        // A message the ring has no room for is one not carried (see `SeamRing`), and one a slot cannot
        // carry at all was never offered to it. Neither is waited on, and neither is dropped on the floor:
        // both fall to this process's own channel below, where a loop here reads one.
        if let Some(slot) = SeamMessage::of_preview(&message) {
            let carried = the_run_facts().preview_ring.push(slot, "preview");

            // One line per hover and notice put in the ring, on the trace a run asks for: a run whose
            // previews are missing is told here whether the tray ever sent one (see
            // `next_preview_message`). Hovers are a pointer's pace and not a tick's, which is what a line
            // here costs.
            if carried && std::env::var_os("RHP_SEAM_TRACE").is_some() {
                eprintln!("seam: the ring took a message from the tray");
            }

            // **A full ring is not a message to wait on**, because this is called from the Explorer
            // hook's callback and a wait there is Explorer not answering (see `SeamRing`). What a sender
            // can do instead is say it to a loop of its own, where it has one: a viewer's own answer
            // keeps its place in the drain after what the ring already holds, and a tray's, which has no
            // loop, is lost with the trace `push` has already written.
            if !carried {
                send_preview_here(message);
            }

            // Safety: an event of this run's own, signalled: nothing here is dereferenced. Signalled
            // either way: a reader woken by a ring with no room in it drains what it has, which is what
            // makes room for the next.
            unsafe {
                let _ = SetEvent(preview_channel_event());
            }
            return;
        }

        // **A message no slot can carry is one only a process can act on, and it is this one.** The one
        // such message is the pin planner's answer (`PreviewMessage::PinAnswered`): it names a playlist
        // and the spawner behind it, which a fixed slot has no room for. The planner is a thread of the
        // process whose pin it is, and the loop that acts on the answer is that same process's, so its own
        // channel is the road — read beside the ring by `next_preview_message`. Dropped instead, every
        // step the pin's own **Next**/**Previous** took would answer nothing: the arc would turn out
        // `PIN_JOB_GIVEUP` and the pin would never be shown the file it stepped onto.
    }

    // A message this process reads is a message kept; one nothing here reads is a loss, and the one road
    // that can lead here is a name longer than a slot carries sent by a process with no loop of its own —
    // the tray's, where the hover or the pick it names is a preview the viewer never hears (see
    // `with_text` and `SeamMessage::of_preview`).
    if !send_preview_here(message) {
        note_a_name_a_slot_could_not_carry();
    }
}

/// Send a message to this process's own preview loop, on the channel its threads share, and wake it.
///
/// It is the whole of what a run that draws its own previews ever does with a message, and the one road
/// a message no ring slot can carry has (see `send_preview` and `next_preview_message`). The answer is
/// whether a loop of this process will read it, which is whether this process has one.
fn send_preview_here(message: PreviewMessage) -> bool {
    let Ok(sender) = PREVIEW_SENDER.lock() else {
        return false;
    };
    let Some(tx) = sender.as_ref() else {
        return false;
    };

    if tx.send(message).is_err() {
        return false;
    }

    // Signalled after the send rather than before, so an armed wait can only be woken by a message
    // that is already in the channel to be read.
    unsafe {
        let _ = SetEvent(preview_channel_event());
    }
    true
}

pub(crate) static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

/// The next preview message waiting, or `None`: off the ring where the two sides are two processes, off
/// this process's channel otherwise — and off that same channel where the ring holds nothing, which is
/// where a message no slot can carry is sent (see `send_preview`).
///
/// It is asked of the same `Receiver` on both roads, so a caller never learns which one answered it: the
/// channel's own reader is the loop, and the three places that read it — the loop's drain, its wait, and
/// the carry a drag is made of — are the whole of what a message is read out of (see
/// `pin_drag::Carry::drain`).
pub(crate) fn next_preview_message(rx: &Receiver<PreviewMessage>) -> Option<PreviewMessage> {
    if the_preview_ring_is_ours() {
        let message = the_run_facts()
            .preview_ring
            .pop()
            .and_then(|slot| slot.as_preview());

        // One line per message the ring hands over, on the trace a run asks for: a hover that was sent and
        // never taken is the difference this line is read against `send_preview`'s. It is a message's pace
        // and not a tick's — the wait calls this on every turn of an idle loop, and a message that is not
        // there costs one `is_some`.
        if message.is_some() && std::env::var_os("RHP_SEAM_TRACE").is_some() {
            eprintln!("seam: the loop took a message off the preview ring");
        }

        if message.is_some() {
            return message;
        }

        // And this process's own channel, which is where a message no slot can carry is sent (see
        // `send_preview`): the pin's planner answers a press in this process and the loop that acts on
        // that answer is this process's, so it is the one road the answer has. The ring is asked first,
        // and a slot a reader could not make a message of is not one: what both roads carry is acted on
        // in the order the loop drains them.
        return rx.try_recv().ok();
    }

    rx.try_recv().ok()
}

/// Wait for a preview message or for window input, whichever comes first.
///
/// This is what makes a pinned window follow the hand at the pointer's own
/// pace rather than at the pace of the repaint: waiting on the message queue instead of on the
/// channel wakes the moment the pointer moves, and costs nothing extra when idle.
///
/// The wait is on two things at once: the message queue, and the channel's own event. The event
/// is what the slicing this used to do was standing in for. A wait that could not see the channel
/// had to wake on a timer to ask it, so `wait_ms` was cut into eight-millisecond slices and each
/// one was a kernel wait followed by a poll of an empty channel — which for a loop with nothing
/// on screen, waiting `IDLE_WAIT_MS`, was a hundred kernel transitions a second for the whole
/// life of the process, doing nothing. A send now signals the event, so the wait ends where the
/// message lands, and an idle loop costs one wait per tick.
///
/// The event is manual-reset and is reset after the wait rather than at the send, which is what
/// makes the arm-then-send race safe: a message that arrived between the poll above and the wait
/// below has already signalled it, so the wait returns at once instead of sleeping through a
/// message that is in the queue.
pub(crate) fn wait_preview_channel(
    rx: &Receiver<PreviewMessage>,
    wait_ms: u64,
) -> Option<PreviewMessage> {
    // A message sent just before the wait is answered without waiting at all.
    if let Some(message) = next_preview_message(rx) {
        return Some(message);
    }

    let event = preview_channel_event();
    let handles = [event];
    if the_preview_ring_was_left() {
        // **A viewer that has become a window of its own is not in the room any more.** The event is the
        // one the viewer that took its place is waiting on, and this wait resets it after the wait — a
        // wake eaten out from under the viewer the hover was meant for. So this process waits on its own
        // message queue alone, and its own channel is still polled below, on whatever the wait hands back.
        // Safety: a wait on this thread's own message queue for a bound; nothing here is dereferenced.
        unsafe {
            let _ =
                MsgWaitForMultipleObjectsEx(None, wait_ms as u32, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        }
    } else {
        // Safety: an event of this run's own, waited on for a bound and reset after the wait, which is what
        // makes the arm-then-send race safe; nothing here is dereferenced.
        unsafe {
            let _ = MsgWaitForMultipleObjectsEx(
                Some(&handles),
                wait_ms as u32,
                QS_ALLINPUT,
                MWMO_INPUTAVAILABLE,
            );
            let _ = ResetEvent(event);
        }
    }

    // Woken by input, by a message, or by the timeout. Either way the channel is only ever
    // polled, never blocked on, so the queue is never left unpumped while a drag is going.
    if let Some(message) = next_preview_message(rx) {
        return Some(message);
    }

    // Input alone is a reason to go round the loop now rather than to sit out the rest of
    // the tick: the queue holds the move the wait woke for.
    let mut msg = MSG::default();
    unsafe {
        if PeekMessageW(&mut msg, None, 0, 0, PM_NOREMOVE).as_bool() {
            return None;
        }
    }

    None
}

/// Wheel notches handed to a text preview. Signed: up is positive. Swapped to
/// zero by the preview thread, so a notch is counted once.
pub(crate) static TEXT_SCROLL_DELTA: Fact<AtomicI32> = Fact::of(|facts| &facts.text_scroll_delta);

/// Wheel notches handed to a pinned picture, in the same units and on the same terms as
/// the text preview's own count above: signed, swapped to zero by the reader, and added
/// to only where the pointer is inside the box the pin published.
pub(crate) static PIN_ZOOM_DELTA: Fact<AtomicI32> = Fact::of(|facts| &facts.pin_zoom_delta);

/// Wheel notches handed to a pin's volume control, on the same terms as the counts
/// above: signed (up is louder), swapped to zero by the reader, and added to only
/// where the pointer is over the volume popup or the button that opens it.
pub(crate) static PIN_VOLUME_DELTA: Fact<AtomicI32> = Fact::of(|facts| &facts.pin_volume_delta);

/// Wheel notches handed to a pinned page's own text size, on the same terms again: a notch turned
/// over a page with `Ctrl` held is how large that page is drawn rather than where it is (see
/// `text_size_after_notches`).
pub(crate) static PIN_TEXT_SIZE_DELTA: Fact<AtomicI32> =
    Fact::of(|facts| &facts.pin_text_size_delta);

/// Wheel notches handed to the playlist band under a pinned sound's card, on the same terms: signed
/// (up is toward the first row), swapped to zero by the reader.
pub(crate) static PIN_PLAYLIST_DELTA: Fact<AtomicI32> = Fact::of(|facts| &facts.pin_playlist_delta);

/// Whether the preview on screen is the surface the pointer is standing on, published by the
/// preview thread once a tick.
///
/// It is what makes a pin *scrollable the moment it is raised*: a notch over a preview used to be
/// this app's only while the pin was the window the keyboard was in — Always On Top, or the window
/// the user had clicked into — so a pin raised over the listing they were working in took no notch
/// and had to be clicked once before it would scroll. The pointer is over the pin the whole time;
/// the keyboard is in the listing behind it, which is the point of the raise (see
/// `Pin Mode → Update Preview → Move To Top`).
///
/// It is published rather than asked here because a `WH_MOUSE_LL` callback may not walk the
/// desktop to find out what is in front — the same reason the pin's key owners are published by
/// the same tick (see `key_input::publish_pin_key_owner`). One tick of staleness is all it can be
/// wrong by, and a notch is turned by a hand that is not moving.
pub(crate) static PREVIEW_UNDER_POINTER: Fact<AtomicU32> =
    Fact::of(|facts| &facts.preview_under_pointer);

/// Publish what the pointer is on, from the tick that is allowed to ask the desktop (see
/// `PREVIEW_UNDER_POINTER`).
pub(crate) fn publish_preview_under_the_pointer(under: bool) {
    PREVIEW_UNDER_POINTER.store(u32::from(under), Ordering::Release);
}

/// Wheel notches a text preview has been given since this was last called, in
/// wheel units (120 to a notch).
pub(crate) fn take_text_scroll_delta() -> i32 {
    TEXT_SCROLL_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pinned page has been given for its own text size since this was last called, in
/// wheel units (120 to a notch). Signed, up being larger.
pub(crate) fn take_text_size_delta() -> i32 {
    PIN_TEXT_SIZE_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pinned picture has been given since this was last called, in wheel
/// units (120 to a notch). Signed, up being positive — which is a zoom in.
pub(crate) fn take_pin_zoom_delta() -> i32 {
    PIN_ZOOM_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pin's playlist band has been given since this was last called, in wheel units
/// (120 to a notch). Signed, up being toward the band's first row.
pub(crate) fn take_pin_playlist_delta() -> i32 {
    PIN_PLAYLIST_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pin's volume control has been given since this was last called, in
/// wheel units (120 to a notch). Signed, up being louder.
pub(crate) fn take_pin_volume_delta() -> i32 {
    PIN_VOLUME_DELTA.swap(0, Ordering::AcqRel)
}

/// Whether a volume notch is waiting to be read, so the preview loop can keep its frame cadence
/// for it rather than let the level land on its own slow wait (see `wait_before_the_next_tick`).
pub(crate) fn pin_volume_delta_pending() -> bool {
    PIN_VOLUME_DELTA.load(Ordering::Acquire) != 0
}

/// The two window handles a standing pin answers keys for, as raw pointers, or zero.
///
/// Published by the preview loop on every tick rather than asked for from inside the hook, which
/// may not take a lock and may not walk the desktop: two atomics and a compare are all a
/// `WH_KEYBOARD_LL` callback may afford, because every keystroke on the machine waits for it to
/// return and Windows removes a hook that is slow.
pub(crate) static PIN_KEYS_OWNER: Fact<AtomicU64> = Fact::of(|facts| &facts.pin_keys_owner);
pub(crate) static PIN_KEYS_VIDEO: Fact<AtomicU64> = Fact::of(|facts| &facts.pin_keys_video);

/// Whether the standing pin is `Pin Mode → Always On Top`, published beside the two handles for
/// the same reason: a window that is above everything is the window the pointer is over, so the
/// notch a hand turns on it is the pin's whether or not the keyboard happens to be in it (see
/// `pin_owns_foreground`).
pub(crate) static PIN_KEYS_TOPMOST: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_keys_topmost);

/// Tell the hook which windows a standing pin is answering for, or that there is no such pin.
///
/// Called by the preview loop on the tick it settles the pin's chrome, which is the same tick that
/// already knows whether a pin is up and which player it is showing. Passing zero for the video
/// handle is the normal case — a pin showing an image has no player to lose keys to.
pub(crate) fn publish_pin_key_owner(owner: u64, video: u64) {
    PIN_KEYS_OWNER.store(owner, Ordering::Release);
    PIN_KEYS_VIDEO.store(video, Ordering::Release);
}

/// The two windows the pin of `block` answers keys for, as it last published them, or two zeros where
/// that viewer has no pin up (see `PIN_KEYS_OWNER`).
///
/// **A key of the pin's belongs to whichever window of the run holds it**, so the hook that takes one
/// asks every block before it decides the caret is nobody's (see `asked_blocks`).
pub(crate) fn the_windows_of_the_pin(block: &'static ViewerFacts) -> (u64, u64) {
    (
        PIN_KEYS_OWNER.in_block(block).load(Ordering::Acquire),
        PIN_KEYS_VIDEO.in_block(block).load(Ordering::Acquire),
    )
}

/// Whether that pin stands above every other window (see `PIN_KEYS_TOPMOST`).
pub(crate) fn the_pin_stands_over_everything(block: &'static ViewerFacts) -> bool {
    PIN_KEYS_TOPMOST.in_block(block).load(Ordering::Acquire)
}

/// The block whose pin answers for the window `foreground`, or `None` where no pin of this run's does.
///
/// **A key of the pin's belongs to whichever window of the run holds it**, and there is exactly one such
/// answer: the block is asked of every one the run has standing, and the key is counted in the one it
/// answers with (see `key_input::pin_owns_caret` for what a window of a pin's is).
pub(crate) fn the_pin_of_the_window(foreground: u64) -> Option<&'static ViewerFacts> {
    if foreground == 0 {
        return None;
    }

    asked_blocks().iter().flatten().copied().find(|block| {
        let (owner, video) = the_windows_of_the_pin(block);
        owner != 0 && (foreground == owner || (video != 0 && foreground == video))
    })
}

/// Whether a find bar is up over that pin's page (see `PIN_FIND_BAR_UP`).
pub(crate) fn the_find_bar_is_up(block: &'static ViewerFacts) -> bool {
    PIN_FIND_BAR_UP.in_block(block).load(Ordering::Acquire)
}

/// Whether the pointer stands on the surface that `block`'s viewer drew (see `PREVIEW_UNDER_POINTER`).
///
/// **It is one block's answer and not the run's**, which is the whole of what a detached window owes:
/// the window the pointer is standing on is the window whose published answer says so, and the viewer
/// that took the ring over is a viewer the pointer may be nowhere near.
pub(crate) fn the_pointer_is_on(block: &'static ViewerFacts) -> bool {
    PREVIEW_UNDER_POINTER
        .in_block(block)
        .load(Ordering::Acquire)
        != 0
}

/// Tell the hooks whether the standing pin is above every other window.
///
/// It is published rather than asked for where it is read for the reason the two handles above
/// are: a low-level callback may not take the configuration's lock, and every notch on the
/// machine waits for the answer. No pin up is the same answer as a pin that is not on top —
/// a hover preview stands in front of everything and needs no exemption (see
/// `pin_owns_foreground`).
pub(crate) fn publish_pin_always_on_top(always_on_top: bool) {
    PIN_KEYS_TOPMOST.store(always_on_top, Ordering::Release);
}

/// Whether the standing pin is shown the file the listing picks (`Pin Mode → Update Preview`),
/// published beside the two handles above for the same reason: the press road asks it from inside a
/// `WH_MOUSE_LL` callback, which may take no lock, and every press on the machine waits for the answer.
///
/// **It is what tells a press the pin owes its place from one the pick road is going to answer.** A
/// press in the listing behind a pin that follows that listing names the file the pin is about to be
/// shown, and the answer to it is the raise the pick makes — so the press must not put the pin behind
/// the listing first, which is the Explorer window flashing over the pin for the tick or ten the pick
/// takes to resolve and the pin takes to come back (see `a_press_puts_us_behind`). A pin that follows
/// nothing is a pin a press in the listing *is* the answer for, and it goes behind as it always has
/// (see `note_a_press`).
///
/// No pin up is the same answer as a pin that follows nothing: a hover is no window to be kept over a
/// listing it does not follow.
pub(crate) static PIN_FOLLOWS_THE_LISTING: Fact<AtomicBool> =
    Fact::of(|facts| &facts.pin_follows_the_listing);

/// Tell the press road whether the standing pin is shown the file the listing picks (see
/// `PIN_FOLLOWS_THE_LISTING`).
pub(crate) fn publish_the_pin_follows_the_listing(follows: bool) {
    PIN_FOLLOWS_THE_LISTING.store(follows, Ordering::Release);
}

/// Whether the standing pin is shown the file the listing picks, as the preview loop last published
/// it (see `PIN_FOLLOWS_THE_LISTING`).
pub(crate) fn the_pin_follows_the_listing() -> bool {
    PIN_FOLLOWS_THE_LISTING.load(Ordering::Acquire)
}

/// The pin's own key commands, one count per command of `PIN_COMMAND_COUNT`, in the order that table
/// lists them (see `take_pin_key_presses`).
static PIN_KEY_PRESSES: Fact<[AtomicU32; PIN_COMMAND_COUNT]> =
    Fact::of(|facts| &facts.pin_key_presses);

/// The Win chords a standing pin answers, in the order the numbers the hook counts them under name them:
/// the pin's own minimize, and its window carried to the display on that side (see
/// `key_input::pin_win_chord` and `preview_window::hook_pin_win_chord`).
static PIN_WIN_CHORDS: Fact<[AtomicU32; PIN_WIN_CHORD_COUNT]> =
    Fact::of(|facts| &facts.pin_win_chords);

/// Take this pin's Win chords, one number per chord and per press since the last call, in order.
///
/// The mirror of `take_pin_key_presses` for the chords that are the Shell's rather than the pin's, drained
/// on the same tick: a chord travels the path a press of the pin's own caption button does, and the loop
/// acts on it (see `ask_pin`).
pub(crate) fn take_pin_win_chords() -> Vec<u8> {
    PIN_WIN_CHORDS
        .iter()
        .enumerate()
        .flat_map(|(chord, count)| {
            std::iter::repeat_n(chord as u8, count.swap(0, Ordering::AcqRel) as usize)
        })
        .collect()
}

/// Count one Win chord the hook took for `block`'s pin, which the loop drains (see `take_pin_win_chords`).
///
/// Counted in the block whose window held the caret and in no other, for the reason a key command is (see
/// `note_a_pin_key_command`): the chord is about the pin of that window, and a viewer whose pin never held
/// the caret has nothing to carry.
pub(crate) fn note_a_pin_win_chord(block: &'static ViewerFacts, chord: u8) {
    if let Some(count) = PIN_WIN_CHORDS.in_block(block).get(chord as usize) {
        count.fetch_add(1, Ordering::AcqRel);
    }
}

/// Take the pin's own key commands, one count per command and per press since the last call.
pub(crate) fn take_pin_key_presses() -> Vec<u8> {
    PIN_KEY_PRESSES
        .iter()
        .enumerate()
        .flat_map(|(command, count)| {
            std::iter::repeat_n(command as u8, count.swap(0, Ordering::AcqRel) as usize)
        })
        .collect()
}

/// Count one press of a command of the pin's key, which the loop drains (see `take_pin_key_presses`).
///
/// The count is an atomic added to on a low-level hook's own thread, where nothing may be allocated and
/// nothing may wait, and it is per command because the order the presses came in is what the loop acts
/// on while the presses themselves are not messages.
///
/// It is counted in the block of the window whose caret took the key, and in no other: a command is
/// swallowed for one window, and a viewer whose pin never held the caret has nothing to act on.
pub(crate) fn note_a_pin_key_command(block: &'static ViewerFacts, command: u8) {
    if let Some(count) = PIN_KEY_PRESSES.in_block(block).get(command as usize) {
        count.fetch_add(1, Ordering::AcqRel);
    }
}

/// How many commands this hook has counted since the process started, which is the one thing a
/// reader of the two roads a pin's key can travel can tell apart: a key this hook counted never
/// became a message, and a key the pin's own window answered never came here (see `pin_owns_caret`).
///
/// It is a `static` of this module's rather than a field of `SeamFacts` because a test is the only
/// thing that reads it, and a section's layout is no place for a field only a test build has.
#[cfg(test)]
pub(crate) static PIN_KEY_COUNTED: AtomicU64 = AtomicU64::new(0);

/// The number of commands this hook has counted (see `PIN_KEY_COUNTED`).
#[cfg(test)]
pub(crate) fn pin_key_counted() -> u64 {
    PIN_KEY_COUNTED.load(Ordering::Relaxed)
}

/// Whether a find bar is up over the standing pin's page, published beside the two handles and for
/// the same reason: `text_find::is_open` takes the media lock, and a `WH_KEYBOARD_LL` callback may
/// take none. It is also the one question that changes what this hook may do at all — a bar that is
/// up takes every key the pin would have answered, by passing it on.
pub(crate) static PIN_FIND_BAR_UP: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_find_bar_up);

/// Tell the hook whether a find bar is up over the pin's page.
///
/// **A swallowed key never becomes a character.** `TranslateMessage` runs after every low-level hook
/// has returned, so the `WM_CHAR` a text field types from exists only for a key-down the hook let
/// through. That is why a bar is not another reader of this hook's commands: while it is up the
/// pin's keys are the field's, and the field's road is a message (see `text_find::key_edit_with`).
pub(crate) fn publish_find_bar_up(open: bool) {
    PIN_FIND_BAR_UP.store(open, Ordering::Release);
}

/// Which of this app's own surfaces the pointer is on, published by the preview loop once a tick.
///
/// **It is published rather than asked because the surface is not the hook's to find.** One of the
/// three is a layered window drawn by the process on the other side of the seam, and the others are a
/// player's window and the engine's, which that process owns as well: `WindowFromPoint` in the tray
/// would name a handle the tray can hold but not draw a preview in, and the four questions the hook
/// asks of it — is the pointer on a preview at all, is it on a picture, on a page, on a film — are all
/// questions about what is *shown* rather than about what the desktop holds. One tick of staleness is
/// what a surface that appears under a hand can be wrong by, and the hand that put it there is not
/// also moving the wheel (see `PIN_KEYS_OWNER` for the same road and the same reason).
static PREVIEW_HOVER: Fact<AtomicU32> = Fact::of(|facts| &facts.preview_hover);

/// Whether the surface the pointer stands on is a picture, a film, or a page the engine draws — the
/// three of them, as the preview loop last saw the pointer (see `PREVIEW_HOVER`).
pub(crate) fn cursor_preview_hover() -> PreviewCursorHover {
    let published = PREVIEW_HOVER.load(Ordering::Acquire);

    PreviewCursorHover {
        image: published & 1 != 0,
        video: published & 2 != 0,
        engine: published & 4 != 0,
    }
}

/// Publish which surface the pointer is on, from the tick that is allowed to ask the desktop (see
/// `PREVIEW_HOVER`).
pub(crate) fn publish_preview_cursor_hover(hover: PreviewCursorHover) {
    let published =
        u32::from(hover.image) | u32::from(hover.video) << 1 | u32::from(hover.engine) << 2;

    PREVIEW_HOVER.store(published, Ordering::Release);
}

/// The box the preview surface on screen stands in, published by the preview loop once a tick: what the
/// hook asks instead of walking the desktop for a window it does not own (see `preview_screen_rect`).
pub(crate) static PREVIEW_SCREEN_RECT: PublishedRegion =
    PublishedRegion::of(|facts| &facts.preview_rect);

/// The box the preview surface on screen stands in, or nothing where none is (see
/// `PREVIEW_SCREEN_RECT`).
///
/// It is the engine's own box where a document is what is shown, because that is where a document is
/// drawn: a page that runs is held through the rectangle the engine put its window in, and the hook
/// cannot ask the engine for it (see `preview_pointer_hold`).
pub(crate) fn preview_screen_rect() -> Option<(i32, i32, i32, i32)> {
    PREVIEW_SCREEN_RECT.get()
}

/// Whether the page the engine is drawing is one that **runs** — a document with script of its own in
/// it rather than a picture of one — published beside the box above.
///
/// It is published rather than derived because the two things it is made of are the preview side's: the
/// path the engine is showing, and whether that path is a page at all (see `html_is_engine_drawn`).
/// The rule the hook reads it by is the one the preview side reads for itself: a page that runs is held
/// through its own rectangle, and a document that cannot be touched is dismissed by the pointer as a
/// picture is (see `preview_pointer_hold`).
pub(crate) static ENGINE_PAGE_RUNS: Fact<AtomicBool> = Fact::of(|facts| &facts.engine_page_runs);

/// Whether a drag that began on the page the engine drew is still down.
///
/// The hook keeps it — a press inside the engine's rectangle arms it, and a tick with no button down at
/// all stands it down, which is the only reading that says the drag is over — and the preview side
/// reads it, which is why it crosses. A drag cannot be seen from where the pointer has got to: the page
/// is under the hand only where the drag began, and everything after that is the page being orbited or
/// panned (see `preview_pointer_hold`).
pub(crate) static ENGINE_PAGE_DRAG: Fact<AtomicBool> = Fact::of(|facts| &facts.engine_page_drag);

/// Whether the page the engine is drawing is one that runs (see `ENGINE_PAGE_RUNS`), as the viewer
/// serving the hovers published it: a detached window's page is a pin the user is working in and not a
/// preview the hook's press rule is about.
pub(crate) fn the_engine_page_runs() -> bool {
    ENGINE_PAGE_RUNS.load(Ordering::Acquire)
}

/// The window the engine draws a document in, while it has one: written where the engine's own handle
/// moves, and back to `0` where the window is destroyed.
///
/// **It crosses because the keyboard is the tray's to answer for.** A click into a page that runs makes
/// the engine's window the foreground, and the app's key hook stands down for it — but the window is the
/// viewer process's, so a handle kept beside the engine answers `0` on the side that asks (see
/// `window_state::engine_owns_the_keyboard`). The handle rather than the hook's own verdict is what is
/// published because that verdict is made of the foreground and of the window's children, which only
/// the side that can ask Win32 for them can make (see `window_state::is_our_own_window`).
///
/// No staleness is accepted here, and none is needed: the writer is the one place the handle changes,
/// where the window is made and where it is destroyed.
pub(crate) static ENGINE_HWND: Fact<AtomicIsize> = Fact::of(|facts| &facts.engine_hwnd);

/// Whether the engine is showing a document at all, published where the engine's own answer moves.
///
/// It is the tray's question and the viewer's fact: the stalled-preview sweep ends the engines a loop
/// that stopped ticking left behind, and a browser drawing a document is the preview itself while it is
/// up rather than an engine being kept (see `probe_trace::end_engines_of_a_stalled_preview`). One bool
/// rather than the shown page's path, because the sweep asks only whether a page is on screen — a name
/// would be a buffer and a torn read for an identity nothing on this side asks for.
pub(crate) static ENGINE_SHOWING: Fact<AtomicBool> = Fact::of(|facts| &facts.engine_showing);

/// Whether a text preview in full mode is on screen, which is the one thing a pointer can rest on to
/// select from and scroll.
///
/// The hold the hook leaves the pointer in is answered from the two boxes below, and both are the
/// preview side's to publish: the window they are measured from is not the hook's, and a pointer the
/// hook dismissed is a preview the hook's own `hide_preview` takes down before its tick asks again.
pub(crate) static TEXT_PREVIEW_HOLDING: Fact<AtomicBool> =
    Fact::of(|facts| &facts.text_preview_holding);

/// Whether what is on screen is the spinner a page is being rendered behind.
///
/// It holds the pointer through the item the wait is for rather than through a box of its own: the
/// spinner is placed at the hand and follows it, so a box of its own would be one the pointer could
/// never leave (see `preview_pointer_hold` and `POINTER_ITEM_BOX`).
pub(crate) static WAITING_PREVIEW_HOLDING: Fact<AtomicBool> =
    Fact::of(|facts| &facts.waiting_preview_holding);

/// The two boxes a preview holds the pointer through, published whole: the journey from the point a
/// hover was opened from to the preview, and the preview itself.
///
/// **They are packed words and not a list**, because the reader is the hook's own tick and the writer
/// is a preview loop on the other side of the seam: a list is a length, a pointer and a heap the other
/// process does not share. Two boxes is all any caller publishes — a text preview's journey and its
/// box, or a pinned window's own box alone — so a slot per box is the whole of it, and a rectangle is
/// packed into each slot whole, which is what makes a torn box impossible to read.
pub(crate) static HOLD_REGION_JOURNEY: PublishedRegion =
    PublishedRegion::of(|facts| &facts.hold_region_journey);
pub(crate) static HOLD_REGION_PREVIEW: PublishedRegion =
    PublishedRegion::of(|facts| &facts.hold_region_preview);

/// The box the view drew the item under the pointer in: the hook's own answer about where the pointer
/// is, published with every look at that item and withdrawn with the window.
///
/// It is what holds a reveal to the item a hover was resolved from, what tells the hook a pointer that
/// has crossed to another row has moved at all, and what a spinner holds the pointer through. Both
/// sides read it and both sides write it — the item is found on the hook's thread and a window going
/// down withdraws it on the preview side — which is why it is one packed word here rather than the
/// locked `Option` it was while both sides were one process: a hook callback may take no lock, and a
/// box read half from one publish and half from the next names a place no item is drawn in.
pub(crate) static POINTER_ITEM_BOX: PublishedRegion =
    PublishedRegion::of(|facts| &facts.pointer_item_box);

/// The left mouse button's state and how many presses it has taken, as the one place in this app that
/// reads the buttons published it (see `pin_drag::left_button_down`).
///
/// `GetAsyncKeyState`'s low-order bit says a key has been pressed since the previous call and is spent
/// by any call for that key, whichever thread makes it: the hook and the pin's own press handling poll
/// at different rates, so a click was answered by whichever of them read it first. The count is a
/// count rather than a bit for the opposite reason — the preview side polls slower, and a bit
/// published and published again reads as no press at all.
pub(crate) static PIN_MEDIA_LEFT_DOWN: Fact<AtomicBool> =
    Fact::of(|facts| &facts.pin_media_left_down);
pub(crate) static PIN_MEDIA_LEFT_PRESSES: Fact<AtomicU64> =
    Fact::of(|facts| &facts.pin_media_left_presses);

/// The boxes a standing pin's own buttons answer a press in, in screen coordinates, published whole.
///
/// The hook reads a press on this app's own window as a press on the listing the window stands over, so
/// a press on a caption's Minimize or on a sound card's own controls would arrive with the file drawn
/// under it (see `pin_button_covers`). **The boxes rather than a mask of which buttons the pin has**:
/// the question asked is about a point, and the only other road is to publish the pin's whole chrome
/// geometry — its content box, band rows, display scale and display options — so that the tray can lay
/// the buttons out again, which is more facts to cross *and* an answer that can disagree with the one
/// the window procedure gives the same press. The geometry is published instead of the verdict, so the
/// intersection is a comparison the hook can afford inside its own callback.
pub(crate) fn publish_pin_buttons(boxes: [Option<(i32, i32, i32, i32)>; PIN_BUTTONS]) {
    let published = the_serving_facts().pin_buttons.as_slice();
    for (cell, region) in published.iter().zip(boxes) {
        cell.store(pack_a_box(region), Ordering::Release);
    }
}

/// Whether a screen point is on one of a standing pin's own buttons (see `publish_pin_buttons`).
///
/// Half-open on every edge, which is how a window hit-test reads a rectangle and how the button's own
/// press is read: a point on a box's right or bottom edge is outside it.
///
/// **Every block is asked**, because the pin whose button it is is one window of the run's wherever it
/// stands: a press on a detached window's caption button is answered for the same way. Half-open on
/// every edge, which is how a window hit-test reads a rectangle and how the button's own press is read.
pub(crate) fn pin_button_covers(point: (i32, i32)) -> bool {
    let (x, y) = point;

    asked_blocks().iter().flatten().any(|block| {
        block.pin_buttons.iter().any(|cell| {
            unpack_a_box(cell.load(Ordering::Acquire)).is_some_and(|(left, top, right, bottom)| {
                x >= left && x < right && y >= top && y < bottom
            })
        })
    })
}

/// The wall clock of the preview loop's last tick, in milliseconds since the epoch: a number both
/// processes can read, which is the whole of why it is a wall clock rather than an `Instant` (see
/// `preview_stall_ms`).
static PREVIEW_ALIVE_UNIX_MS: Fact<AtomicU64> = Fact::of(|facts| &facts.preview_alive_unix_ms);

/// How long the preview loop has been quiet, in milliseconds: the age of its last tick, growing for as
/// long as the loop is inside work that has not come back.
///
/// **A wall clock, because the two ends are two processes.** An `Instant` is this process's own and
/// cannot be read in another, so what crosses is the milliseconds since the epoch of the loop's last
/// tick, and the age is the difference between that number and the one the reader's own clock gives.
/// A loop that has never ticked answers with no quiet at all: a preview that has not been begun is not
/// a preview that has stopped. Nothing here waits on the loop or looks for it — it is only ever a
/// number read.
///
/// It is the block of the viewer the hovers are served from, and no other's: a detached window whose
/// loop has stopped is a pin of the user's standing still, not a preview the sweep owes an answer for.
///
/// shortcut: a wall clock stepped backwards reads as no quiet, upgrade if a run must survive the clock
/// being set back under it.
pub(crate) fn preview_stall_ms() -> u64 {
    let alive = PREVIEW_ALIVE_UNIX_MS.load(Ordering::Relaxed);
    if alive == 0 {
        return 0;
    }

    let now = unix_millis();
    now.saturating_sub(alive)
}

/// Note that the preview loop has run a tick, whether the tick did anything or not: a loop waiting on
/// the channel for a hover is a loop that is working (see `preview_stall_ms`).
pub(crate) fn note_the_preview_ticked() {
    PREVIEW_ALIVE_UNIX_MS.store(unix_millis(), Ordering::Relaxed);
}

/// Milliseconds since the epoch, which is the only clock two processes can both read.
fn unix_millis() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|age| age.as_millis() as u64)
        .unwrap_or_default()
}

/// What the preview side tells the tray side, on a channel of its own.
///
/// Three things, and all three are notices rather than questions: a setting the pin's own chrome can
/// be asked to turn, a shortcut pressed at a preview, and a line for the hook's click trace. None of
/// them is answered, which is why they travel one way and why the sender never waits for one to be
/// read — a preview thread that blocked on the tray would be a picture that stopped while a menu was
/// open.
pub(crate) enum TrayMessage {
    /// The pin's **Always On Top** was turned over from the pin itself — its caption button, or the
    /// key bound to the same command. The setting is a row of the tray's with a file behind it, so the
    /// flip is asked for here and made there; the pin is told again by the message the flip sends
    /// (see `PreviewMessage::PinChanged`), which is the whole of what makes the answer arrive once the
    /// file says what it says rather than once the ask was made.
    TogglePinAlwaysOnTop,
    /// The preview shortcut was pressed over Explorer, and previews are switched off or on. It is
    /// counted by the low-level keyboard hook and was acted on by the preview loop only because that
    /// loop is where it happened to be drained; the switch is the tray's own, and the loop's business
    /// is not to hold a setting.
    TogglePreviewEnabled,
    /// One line for the hook's click trace, already built and not yet written (see
    /// `note_a_pin_click`). It is a `String` because the side that writes it also decides the file, the
    /// gate and the words around it, and because a line built on the asking side is a line built
    /// whether or not the run asked for any.
    PinClick(String),
    /// A viewer's pin has become a window of its own, and there is nobody drawing the hovers until this
    /// run has a viewer again (see `app::detach`). It carries the pid of the window that left, which is
    /// the one this run keeps beside the viewer that serves the hovers from then on: that process is
    /// still standing on the screen, and it is standing there as nobody's pin.
    ViewerDetached(u32),
}

/// One message as the bytes a ring carries: which message it is, the text it names, and the numbers its
/// variant carries.
///
/// It has a fixed shape because a slot in a shared section does: nothing in it is a pointer, so a message
/// the tray writes reads the same in the viewer, and nothing in it is allocated, so a message is written
/// on a hook's own thread without a lock and without a heap. A variant whose payload is not a fixed
/// number of bytes is not carried at all (see `of_preview`).
///
/// **The text is the path a message names, or the line a notice carries**, and it is one field for the
/// reason a slot is one shape: a line of a click trace is not a path and is short, and a second buffer
/// would be a second length to keep in step with the first.
#[repr(C)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct SeamMessage {
    /// Which variant this is, as one of the `TAG_*` constants below.
    tag: u8,
    /// `text_len` units of UTF-16: the path a variant names, or the line a notice carries.
    text: [u16; SEAM_TEXT_UNITS],
    text_len: u32,
    /// What the variant carries besides its text, in the order the variant writes its fields.
    numbers: [i64; SEAM_NUMBERS],
}

impl SeamMessage {
    /// No message: what a slot holds before anything is written in it.
    const EMPTY: Self = Self {
        tag: 0,
        text: [0; SEAM_TEXT_UNITS],
        text_len: 0,
        numbers: [0; SEAM_NUMBERS],
    };

    // What each variant is tagged with. A tag is a number rather than a `repr` enum's discriminant because a
    // tag read out of a section is a byte this build may not know, and a byte that is not a variant of a
    // fieldless enum is not a value to hand the compiler.
    const SHOW: u8 = 1;
    const SHOW_KEYBOARD: u8 = 2;
    const HIDE: u8 = 3;
    const REFRESH: u8 = 4;
    const REFRESH_TYPES: u8 = 5;
    const PIN: u8 = 6;
    const PIN_BOX: u8 = 7;
    const PIN_UPDATE: u8 = 8;
    const PIN_PRESS: u8 = 9;
    const PIN_NO_PICK: u8 = 10;
    const PIN_CHANGED: u8 = 11;
    const TRAY_VOLUME: u8 = 12;
    const TRAY_TEXT_SIZE: u8 = 13;
    const FORGET_VIDEO_GEOMETRY: u8 = 14;
    const KEEP_EXTRACTION: u8 = 15;
    const OFFICE_RENDER_READY: u8 = 16;
    const VIDEO_PROBED: u8 = 17;
    const MEASURE_PROBED: u8 = 18;
    const VIDEO_SUBTITLES_READY: u8 = 19;
    const MAGICK_READY: u8 = 20;
    const PEAZIP_READY: u8 = 21;
    const TOGGLE_PIN_ALWAYS_ON_TOP: u8 = 22;
    const TOGGLE_PREVIEW_ENABLED: u8 = 23;
    const PIN_CLICK: u8 = 24;
    const STOP_ENGINES: u8 = 25;
    const HIDE_HTML_PAGE: u8 = 26;
    const TRIM_CACHES: u8 = 27;
    const KILL_STRAY_PLAYER: u8 = 28;
    const VIEWER_DETACHED: u8 = 29;
    const PIN_WALK_AGAIN: u8 = 30;
    const PIN_RESTORE: u8 = 31;

    /// A message tagged `tag` carrying `numbers` in that order, and no text yet.
    fn of(tag: u8, numbers: &[i64]) -> Self {
        let mut carried = [0; SEAM_NUMBERS];
        // A variant wider than a slot would lose its last numbers here, which is what the round trip of
        // every variant is a test of.
        for (into, from) in carried.iter_mut().zip(numbers) {
            *into = *from;
        }

        Self {
            tag,
            numbers: carried,
            ..Self::EMPTY
        }
    }

    /// Fill in the text the message carries, or answer `None` where it does not fit a slot.
    ///
    /// shortcut: a path longer than `SEAM_TEXT_UNITS` units is a message lost rather than a path
    /// truncated into another file's; upgrade if a user ever meets a path that long.
    fn with_text(mut self, text: impl Iterator<Item = u16>) -> Option<Self> {
        for unit in text {
            let at = self.text_len as usize;
            if at == SEAM_TEXT_UNITS {
                return None;
            }

            self.text[at] = unit;
            self.text_len += 1;
        }

        Some(self)
    }

    /// The text as the path it names, or `None` where it names none.
    fn path(&self) -> Option<PathBuf> {
        if self.text_len == 0 {
            return None;
        }

        Some(PathBuf::from(OsString::from_wide(
            &self.text[..self.text_len as usize],
        )))
    }

    /// The text as the line a notice carries.
    fn line(&self) -> String {
        String::from_utf16_lossy(&self.text[..self.text_len as usize])
    }

    /// A preview message as the slot it travels in, or `None` for one that is not carried at all.
    ///
    /// **A planner's answer is the one that is not carried.** What it holds is a whole playlist and the
    /// name of a program, and neither is a fixed number of bytes; nothing of it crosses either, because
    /// the question is asked of the planner by the preview side and answered on it (see `PinPlanned`).
    fn of_preview(message: &PreviewMessage) -> Option<Self> {
        match message {
            PreviewMessage::Show(path, x, y, region) => {
                let column = a_flag(region.is_some_and(|region| region.column));
                let (left, top, right, bottom) = a_box(region.map(|region| region.region));
                Self::of(
                    Self::SHOW,
                    &[
                        *x as i64,
                        *y as i64,
                        a_flag(region.is_some()),
                        left,
                        top,
                        right,
                        bottom,
                        column,
                    ],
                )
                .with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::ShowKeyboard(path, x, y, width, height, region, row) => {
                let (left, top, right, bottom) = a_box(*region);
                Self::of(
                    Self::SHOW_KEYBOARD,
                    &[
                        *x as i64,
                        *y as i64,
                        *width as i64,
                        *height as i64,
                        a_flag(region.is_some()),
                        left,
                        top,
                        right,
                        bottom,
                        a_flag(*row),
                    ],
                )
                .with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::Hide => Some(Self::of(Self::HIDE, &[])),
            PreviewMessage::Refresh => Some(Self::of(Self::REFRESH, &[])),
            PreviewMessage::RefreshTypes => Some(Self::of(Self::REFRESH_TYPES, &[])),
            PreviewMessage::StopEngines => Some(Self::of(Self::STOP_ENGINES, &[])),
            PreviewMessage::HideHtmlPage => Some(Self::of(Self::HIDE_HTML_PAGE, &[])),
            PreviewMessage::TrimCaches => Some(Self::of(Self::TRIM_CACHES, &[])),
            PreviewMessage::KillStrayPlayer => Some(Self::of(Self::KILL_STRAY_PLAYER, &[])),
            PreviewMessage::Pin {
                path,
                rect,
                arriving,
            } => {
                let (left, top, right, bottom) = a_box(Some(*rect));
                Self::of(Self::PIN, &[left, top, right, bottom, a_flag(*arriving)])
                    .with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::PinBox(rect) => {
                let (left, top, right, bottom) = a_box(Some(*rect));
                Some(Self::of(Self::PIN_BOX, &[left, top, right, bottom]))
            }
            PreviewMessage::PinUpdate(path) => {
                Self::of(Self::PIN_UPDATE, &[]).with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::PinPress { window } => {
                Some(Self::of(Self::PIN_PRESS, &[*window as i64]))
            }
            PreviewMessage::PinNoPick { window } => {
                Some(Self::of(Self::PIN_NO_PICK, &[*window as i64]))
            }
            PreviewMessage::PinChanged => Some(Self::of(Self::PIN_CHANGED, &[])),
            PreviewMessage::TrayVolume { audio } => {
                Some(Self::of(Self::TRAY_VOLUME, &[a_flag(*audio)]))
            }
            PreviewMessage::TrayTextSize { percent } => {
                Some(Self::of(Self::TRAY_TEXT_SIZE, &[*percent as i64]))
            }
            PreviewMessage::ForgetVideoGeometry => Some(Self::of(Self::FORGET_VIDEO_GEOMETRY, &[])),
            PreviewMessage::KeepExtraction { path } => {
                let numbers = [a_flag(path.is_some())];
                match path {
                    Some(path) => Self::of(Self::KEEP_EXTRACTION, &numbers)
                        .with_text(path.as_os_str().encode_wide()),
                    None => Some(Self::of(Self::KEEP_EXTRACTION, &numbers)),
                }
            }
            PreviewMessage::OfficeRenderReady {
                path,
                generation,
                ok,
            } => Self::of(
                Self::OFFICE_RENDER_READY,
                &[*generation as i64, a_flag(*ok)],
            )
            .with_text(path.as_os_str().encode_wide()),
            PreviewMessage::VideoProbed { path, generation } => {
                Self::of(Self::VIDEO_PROBED, &[*generation as i64])
                    .with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::MeasureProbed { path, size } => {
                let (width, height) =
                    size.map_or((0, 0), |(width, height)| (width as i64, height as i64));
                Self::of(
                    Self::MEASURE_PROBED,
                    &[a_flag(size.is_some()), width, height],
                )
                .with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::PinAnswered(_) => None,
            PreviewMessage::PinWalkAgain => Some(Self::of(Self::PIN_WALK_AGAIN, &[])),
            PreviewMessage::PinRestore => Some(Self::of(Self::PIN_RESTORE, &[])),
            PreviewMessage::VideoSubtitlesReady(path) => {
                Self::of(Self::VIDEO_SUBTITLES_READY, &[]).with_text(path.as_os_str().encode_wide())
            }
            PreviewMessage::MagickReady {
                path,
                generation,
                ok,
            } => Self::of(Self::MAGICK_READY, &[*generation as i64, a_flag(*ok)])
                .with_text(path.as_os_str().encode_wide()),
            PreviewMessage::PeazipReady {
                path,
                generation,
                ok,
            } => Self::of(Self::PEAZIP_READY, &[*generation as i64, a_flag(*ok)])
                .with_text(path.as_os_str().encode_wide()),
        }
    }

    /// A tray message as the slot it travels in.
    fn of_tray(message: &TrayMessage) -> Option<Self> {
        match message {
            TrayMessage::TogglePinAlwaysOnTop => {
                Some(Self::of(Self::TOGGLE_PIN_ALWAYS_ON_TOP, &[]))
            }
            TrayMessage::TogglePreviewEnabled => Some(Self::of(Self::TOGGLE_PREVIEW_ENABLED, &[])),
            TrayMessage::PinClick(line) => {
                Self::of(Self::PIN_CLICK, &[]).with_text(line.encode_utf16())
            }
            TrayMessage::ViewerDetached(pid) => {
                Some(Self::of(Self::VIEWER_DETACHED, &[i64::from(*pid)]))
            }
        }
    }

    /// The preview message this slot carries, or `None` where it carries none of them.
    ///
    /// The `_` arm is the tag of a message the other way, or a byte this build does not know. Losing that
    /// message is better than reading another one out of it: the tag is the only thing a slot says about
    /// itself, and a slot is not a thing to trust a variant out of.
    fn as_preview(self) -> Option<PreviewMessage> {
        let numbers = &self.numbers;
        Some(match self.tag {
            Self::SHOW => {
                let region = (numbers[2] != 0).then(|| {
                    let region = box_of(numbers, 3);
                    if numbers[7] != 0 {
                        AvoidRegion::column(region)
                    } else {
                        AvoidRegion::text(region)
                    }
                });
                PreviewMessage::Show(self.path()?, numbers[0] as i32, numbers[1] as i32, region)
            }
            Self::SHOW_KEYBOARD => PreviewMessage::ShowKeyboard(
                self.path()?,
                numbers[0] as i32,
                numbers[1] as i32,
                numbers[2] as i32,
                numbers[3] as i32,
                (numbers[4] != 0).then(|| box_of(numbers, 5)),
                numbers[9] != 0,
            ),
            Self::HIDE => PreviewMessage::Hide,
            Self::REFRESH => PreviewMessage::Refresh,
            Self::REFRESH_TYPES => PreviewMessage::RefreshTypes,
            Self::STOP_ENGINES => PreviewMessage::StopEngines,
            Self::HIDE_HTML_PAGE => PreviewMessage::HideHtmlPage,
            Self::TRIM_CACHES => PreviewMessage::TrimCaches,
            Self::KILL_STRAY_PLAYER => PreviewMessage::KillStrayPlayer,
            Self::PIN => PreviewMessage::Pin {
                path: self.path()?,
                rect: box_of(numbers, 0),
                arriving: numbers[4] != 0,
            },
            Self::PIN_BOX => PreviewMessage::PinBox(box_of(numbers, 0)),
            Self::PIN_UPDATE => PreviewMessage::PinUpdate(self.path()?),
            Self::PIN_PRESS => PreviewMessage::PinPress {
                window: numbers[0] as isize,
            },
            Self::PIN_NO_PICK => PreviewMessage::PinNoPick {
                window: numbers[0] as isize,
            },
            Self::PIN_CHANGED => PreviewMessage::PinChanged,
            Self::PIN_WALK_AGAIN => PreviewMessage::PinWalkAgain,
            Self::PIN_RESTORE => PreviewMessage::PinRestore,
            Self::TRAY_VOLUME => PreviewMessage::TrayVolume {
                audio: numbers[0] != 0,
            },
            Self::TRAY_TEXT_SIZE => PreviewMessage::TrayTextSize {
                percent: numbers[0] as u32,
            },
            Self::FORGET_VIDEO_GEOMETRY => PreviewMessage::ForgetVideoGeometry,
            Self::KEEP_EXTRACTION => PreviewMessage::KeepExtraction {
                path: if numbers[0] == 0 {
                    None
                } else {
                    Some(self.path()?)
                },
            },
            Self::OFFICE_RENDER_READY => PreviewMessage::OfficeRenderReady {
                path: self.path()?,
                generation: numbers[0] as u64,
                ok: numbers[1] != 0,
            },
            Self::VIDEO_PROBED => PreviewMessage::VideoProbed {
                path: self.path()?,
                generation: numbers[0] as u64,
            },
            Self::MEASURE_PROBED => PreviewMessage::MeasureProbed {
                path: self.path()?,
                size: (numbers[0] != 0).then(|| (numbers[1] as u32, numbers[2] as u32)),
            },
            Self::VIDEO_SUBTITLES_READY => PreviewMessage::VideoSubtitlesReady(self.path()?),
            Self::MAGICK_READY => PreviewMessage::MagickReady {
                path: self.path()?,
                generation: numbers[0] as u64,
                ok: numbers[1] != 0,
            },
            Self::PEAZIP_READY => PreviewMessage::PeazipReady {
                path: self.path()?,
                generation: numbers[0] as u64,
                ok: numbers[1] != 0,
            },
            _ => return None,
        })
    }

    /// The tray message this slot carries, or `None` where it carries none of them (see `as_preview`).
    fn as_tray(self) -> Option<TrayMessage> {
        Some(match self.tag {
            Self::TOGGLE_PIN_ALWAYS_ON_TOP => TrayMessage::TogglePinAlwaysOnTop,
            Self::TOGGLE_PREVIEW_ENABLED => TrayMessage::TogglePreviewEnabled,
            Self::PIN_CLICK => TrayMessage::PinClick(self.line()),
            Self::VIEWER_DETACHED => TrayMessage::ViewerDetached(self.numbers[0] as u32),
            _ => return None,
        })
    }
}

/// One for a yes and zero for a no, which is how a slot carries a flag.
const fn a_flag(flag: bool) -> i64 {
    if flag {
        1
    } else {
        0
    }
}

/// A box as the four numbers a slot carries, or four zeros where there is none.
fn a_box(region: Option<ScreenRegion>) -> (i64, i64, i64, i64) {
    region.map_or((0, 0, 0, 0), |(left, top, right, bottom)| {
        (
            i64::from(left),
            i64::from(top),
            i64::from(right),
            i64::from(bottom),
        )
    })
}

/// The box four of a slot's numbers are, read from `at`.
fn box_of(numbers: &[i64], at: usize) -> ScreenRegion {
    let edge = |at: usize| numbers[at] as i32;
    (edge(at), edge(at + 1), edge(at + 2), edge(at + 3))
}

/// Say that a message did not fit the ring it was sent on, on the trace a run asks for when what is under
/// the seam is being measured.
///
/// The line is built only where a run asked for it, and it is written on whatever thread the sender was:
/// a loss is the exceptional case already, and a lost message reported is a session whose rings are too
/// small rather than one where a sender should have waited.
fn note_a_lost_message(direction: &str) {
    if std::env::var_os("RHP_SEAM_TRACE").is_none() {
        return;
    }

    eprintln!("seam: the {direction} ring is full, message dropped");
}

/// Say that a name longer than a slot carries had no channel of this process to fall back to, on the trace
/// a run asks for when what is under the seam is being measured.
///
/// A slot carries `SEAM_TEXT_UNITS` units and a longer name is a message `SeamMessage::of_preview` refuses
/// (see `with_text`), which is the one way a message reaches the fall-through in `send_preview` without a
/// slot: a hover of a path past that bound is a preview the viewer never hears, and the process that can
/// only lose it this way is the tray, which has no loop of its own to hand it to.
fn note_a_name_a_slot_could_not_carry() {
    if std::env::var_os("RHP_SEAM_TRACE").is_none() {
        return;
    }

    eprintln!("seam: a name longer than a slot carries, and this process reads no channel: message dropped");
}

/// The channel the two, and the one event its sends are announced on.
///
/// Both ends are made here rather than handed in by whoever starts first: the tray side's reader is a
/// loop that begins with the hook and the preview thread's writer begins with the window, and neither
/// is a place to be told about the other's start-up. It is a queue rather than an atomic because a
/// trace line does not fit one and because the order the asks are made in is the order they must be
/// answered in; nothing here is on a hook's own callback, which is the only place a lock is forbidden.
struct TrayChannel {
    sender: Mutex<Sender<TrayMessage>>,
    receiver: Mutex<Receiver<TrayMessage>>,
}

static TRAY_CHANNEL: Lazy<TrayChannel> = Lazy::new(|| {
    let (sender, receiver) = std::sync::mpsc::channel();
    TrayChannel {
        sender: Mutex::new(sender),
        receiver: Mutex::new(receiver),
    }
});

/// The event a tray message's send is announced on, so a thread waiting for one sleeps on the channel
/// as well as on its own clock.
///
/// It is an event of its own rather than the preview channel's: the two waits are on different threads
/// — the preview loop waits for a hover, the hook for a pin — and an event one of them resets after
/// its wait would be an event the other could be woken by and find nothing on the channel it read.
pub(crate) fn tray_channel_event() -> HANDLE {
    static EVENT: Lazy<isize> = Lazy::new(|| {
        let name = the_seam_name("-tray");
        // Safety: an event with a name of this run's own and no security descriptor, manual-reset,
        // which is what a wait that may be woken several times before it is read wants.
        unsafe { CreateEventW(None, true, false, PCWSTR(name.as_ptr())) }
            .map(|handle| handle.0 as isize)
            .unwrap_or_default()
    });

    HANDLE(*EVENT as *mut core::ffi::c_void)
}

/// Tell the tray side something, and wake it where it stands.
///
/// The send is signalled after the message is in the channel rather than before, so a wait that is
/// already armed can only be woken by something there is to read.
pub(crate) fn send_tray(message: TrayMessage) {
    if the_rings_carry_messages() {
        // A notice the ring has no room for is dropped rather than waited on: a preview thread that
        // blocked on the tray would be a picture that stopped while a menu was open.
        if let Some(slot) = SeamMessage::of_tray(&message) {
            the_run_facts().tray_ring.push(slot, "tray");
        }
        // Safety: an event of this run's own, signalled: nothing here is dereferenced.
        unsafe {
            let _ = SetEvent(tray_channel_event());
        }
        return;
    }

    if let Ok(sender) = TRAY_CHANNEL.sender.lock() {
        let _ = sender.send(message);
        unsafe {
            let _ = SetEvent(tray_channel_event());
        }
    }
}

/// Everything the preview side has told the tray side since this was last called, in order.
pub(crate) fn take_tray_messages() -> Vec<TrayMessage> {
    if the_rings_carry_messages() {
        let ring = &the_run_facts().tray_ring;
        return std::iter::from_fn(|| ring.pop())
            .filter_map(|slot| slot.as_tray())
            .collect();
    }

    let Ok(receiver) = TRAY_CHANNEL.receiver.lock() else {
        return Vec::new();
    };

    std::iter::from_fn(|| receiver.try_recv().ok()).collect()
}

/// Wait for a tray message, or for the same time out, whichever comes first.
///
/// This is the hook loop's sleep, and it is a wait on the channel's event rather than a plain sleep for
/// the reason the preview loop's own wait is: a message is answered where it lands rather than at the
/// end of whatever a loop had decided to sit out, and the interval is still exactly the interval when
/// nothing arrives. There is no message queue to wake for as well — the hook's thread pumps no window
/// messages, its hook callbacks arriving on the two threads that pump one — so this is one kernel wait
/// and nothing else.
pub(crate) fn wait_for_the_tray(wait_ms: u64) {
    let event = tray_channel_event();
    unsafe {
        let _ = WaitForSingleObject(event, wait_ms as u32);
        let _ = ResetEvent(event);
    }
}

/// Write one line about what a click in the listing did to a pinned window, from the preview side.
///
/// The line belongs to the hook's own trace file and is written by the loop that owns it, so what
/// travels is the line and not a call: the gate is read here, on the side that pays for building it,
/// and a run that asked for no trace sends nothing at all.
pub(crate) fn note_a_pin_click(line: String) {
    if std::env::var_os("RHP_HOOK_TRACE").is_some() || std::env::var_os("RHP_PIN_TRACE").is_some() {
        send_tray(TrayMessage::PinClick(line));
    }
}

/// A pin was asked to come down, by the tray side (previews were switched off, or the trigger key is
/// holding them back) or by the resumption of the machine from sleep. What a pin *is* belongs to the
/// preview side, so this is a request rather than a take-down: the loop ends it on its next tick, through
/// the same path its own close button takes (see `pin_lifecycle`).
pub(crate) static PIN_END_REQUESTED: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_end_requested);

/// The pin's menu was asked to come down, by the hook, because a press landed on a window that is not the
/// pin's own. The pin window is `WS_NOACTIVATE` and holds no capture, so a press on another window never
/// reaches its procedure and the hook's poll is the only side that sees it. Only the menu goes: the pin is
/// left standing (see `pin_menu::PIN_MENU_DISMISS_REQUESTED`).
pub(crate) static PIN_MENU_DISMISS_REQUESTED: Fact<AtomicBool> =
    Fact::of(|facts| &facts.pin_menu_dismiss_requested);

/// Whether a press was made on the pin's own menu, asked of by the hook once a tick and cleared on read:
/// the menu's window is the preview side's, and what a press on it means is decided there (see
/// `pin_menu::take_pin_menu_pressed`).
pub(crate) static PIN_MENU_PRESSED: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_menu_pressed);

/// Whether a pin has been closed since this was last asked, cleared on read. The hook reads it once per
/// tick to know that what is under the pointer is a hover it has not answered yet: the file it was on
/// when the pin went up is not a file the pointer has left and come back to, and treating it as one would
/// hold the next preview back for the re-hover delay (see `pin_window::PIN_RESUMED`).
pub(crate) static PIN_RESUMED: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_resumed);

/// Whether the pin that came down did so because its file was handed off to another program rather than
/// closed by the user. The hook reads it beside `PIN_RESUMED`, and what it answers is the opposite: a pin
/// that was closed is a pointer on something new, while a pin that was handed off is a hand that has left
/// for another program, and a preview raised over it is a card the click did not mean (see
/// `pin_window::PIN_HANDED_OFF`).
pub(crate) static PIN_HANDED_OFF: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_handed_off);

/// Whether a preview is pinned: a window of its own with a caption, which stays until it is closed,
/// collapsed into its bubble included.
///
/// It is the fact the hook's whole tick is arranged around — while it is true the hover machinery is
/// quiet, the stray-player sweep is held, and the trigger key means nothing — and it used to be asked
/// of the preview thread through a lock, which is a hook that stops answering Explorer behind a preview
/// thread that has stopped turning. This is the published copy written beside the state it copies (see
/// `pin_window::PIN_UP`, where it is written).
pub(crate) static PIN_UP: Fact<AtomicBool> = Fact::of(|facts| &facts.pin_up);

/// Whether a preview is pinned, answered without a lock and without waiting for one (see `PIN_UP`).
///
/// **Every block is asked**, because a pin is one window of the run's wherever it stands: a pin
/// detached into a window of its own still holds the hovers back, and a viewer that took its place in
/// the ring is a viewer with no pin on screen.
pub(crate) fn the_pin_is_up() -> bool {
    the_pin_block().is_some()
}

/// The block a pin is up in, or nothing where no pin is up: the block whose `pin_up` is set (see
/// `PIN_UP`).
///
/// **A fact about a pin is asked of this block and of no other.** A viewer that is not the pin's — one
/// that left the ring with the window it detached, one begun for the hovers after it — publishes facts
/// of its own that a pin is nothing to do with, and reading the first of those that answers is a fact
/// about one window answered by another (see `pin_is_on_its_way_to` and `pinned_path`). A window of its
/// own publishes no pin of the tray's at all (see `app::detach`), which is what leaves this answering
/// nothing for it.
fn the_pin_block() -> Option<&'static ViewerFacts> {
    asked_blocks()
        .iter()
        .flatten()
        .copied()
        .find(|block| PIN_UP.in_block(block).load(Ordering::Acquire))
}

/// Take a flag one viewer set about its own pin, from whichever block of the run set it: a pin is one
/// window of the run's wherever it stands, and a flag is said once (see `PIN_RESUMED`).
pub(crate) fn take_a_pin_flag(flag: &Fact<AtomicBool>) -> bool {
    asked_blocks()
        .iter()
        .flatten()
        .any(|block| flag.in_block(block).swap(false, Ordering::AcqRel))
}

/// The window of `window` that a block published, where one did (see `PIN_BUBBLE_HWND`).
fn a_window_of_the_pin(window: &Fact<AtomicIsize>) -> isize {
    asked_blocks()
        .iter()
        .flatten()
        .map(|block| window.in_block(block).load(Ordering::Acquire))
        .find(|handle| *handle != 0)
        .unwrap_or(0)
}

/// The window a collapsed pin leaves on screen: the circle the whole of a pin is hidden behind, which
/// carries the pin's own menu in the same window.
///
/// Published by the preview thread and read by the hook, for the same reason `PIN_UP` is: a press the
/// hook reads on the tick after a row of that menu was pressed landed on the panel, and the window under
/// the pointer is how that press is told from one on the file the panel covers (see
/// `window_state::bubble_window_is_at`).
pub(crate) static PIN_BUBBLE_HWND: Fact<AtomicIsize> = Fact::of(|facts| &facts.pin_bubble_hwnd);

/// The window a standing pin's own menu is drawn in, published beside the bubble and read by the same
/// press (see `PIN_BUBBLE_HWND`).
pub(crate) static PIN_MENU_HWND: Fact<AtomicIsize> = Fact::of(|facts| &facts.pin_menu_hwnd);

/// The window a collapsed pin of this run is hidden behind, or zero where no pin is collapsed (see
/// `PIN_BUBBLE_HWND`).
pub(crate) fn the_bubble_window() -> isize {
    a_window_of_the_pin(&PIN_BUBBLE_HWND)
}

/// The window a standing pin's own menu of this run is drawn in, or zero where none is (see
/// `PIN_MENU_HWND`).
pub(crate) fn the_menu_window() -> isize {
    a_window_of_the_pin(&PIN_MENU_HWND)
}

/// The file a pin has been asked for and is not showing yet, as a fingerprint, or `0` for none.
///
/// **The file itself cannot be published here.** A `PathBuf` behind a lock is an allocation and a wait
/// on a path the hook asks every time it offers a pick, and the hook may have neither; so what crosses
/// is the number a hash of the path comes to, and the hook hashes the path it is holding and compares
/// it. What the answer is for is one press: a press on the file the pin is already on its way to is not
/// another file, and read as one it restarts the swap the pin is waiting on (see
/// `update_pinned_preview`).
///
/// shortcut: two spellings of one file — a case difference, a short name — hash apart, and the cost is a
/// press answered as a new pick, which is what it was before this was published at all. Upgrade to a
/// published path if a user ever meets it.
pub(crate) static PIN_FILE_IN_FLIGHT: Fact<AtomicU64> = Fact::of(|facts| &facts.pin_file_in_flight);

/// Note the file a pin is now on its way to, or `None` where nothing is in flight any more.
pub(crate) fn publish_the_pin_file_in_flight(path: Option<&Path>) {
    PIN_FILE_IN_FLIGHT.store(path.map_or(0, fingerprint_of), Ordering::Release);
}

/// Whether `path` is the file the pin is already on its way to (see `PIN_FILE_IN_FLIGHT`).
///
/// **Asked of the pin's own block and of no other** (see `the_pin_block`): the flag is the pin's, and a
/// block that holds one with no pin up — a window that left the ring mid-load — would otherwise answer
/// for a pin that has nothing in flight at all, which is every pick of that file refused in silence.
pub(crate) fn pin_is_on_its_way_to(path: &Path) -> bool {
    let wanted = fingerprint_of(path);

    the_pin_block()
        .is_some_and(|block| PIN_FILE_IN_FLIGHT.in_block(block).load(Ordering::Acquire) == wanted)
}

/// The window a pin is to be put behind, published by the hook and taken by the viewer that holds the
/// pin: zero for nothing asked (see `ask_the_pin_to_go_behind`).
///
/// It crosses because the two halves of the decision are on different sides: the hook is what watches the
/// foreground every tick and knows the user has come back to the listing, and the window is the pin's, so
/// the bringing-forward-or-behind is the viewer's own (see `pin_lifecycle::lower_pinned_preview`).
pub(crate) static PIN_BEHIND_ASK: Fact<AtomicIsize> = Fact::of(|facts| &facts.pin_behind_ask);

/// Ask the pin to put itself behind `window`, from any thread: the hook's answer to the user coming
/// back to the listing the pin is standing over, which is a question only the window that holds the pin
/// can answer (see `pin_lifecycle::lower_pinned_preview`).
///
/// **Every block is asked**, the way a press and an end are: a pin is one window of the run's wherever
/// it stands, a window of its own included, and the viewer that holds it is the one that acts on this (see
/// `note_a_pin_press`). A viewer with no pin spends it and puts nothing anywhere (see the preview loop).
pub(crate) fn ask_the_pin_to_go_behind(window: isize) {
    ask_the_pin_to_go_behind_in(asked_blocks(), window);
}

/// Ask the pin to go behind `window` in `blocks` (see `ask_the_pin_to_go_behind`).
fn ask_the_pin_to_go_behind_in(
    blocks: [Option<&'static ViewerFacts>; SEAM_VIEWERS],
    window: isize,
) {
    for block in blocks.iter().flatten() {
        PIN_BEHIND_ASK
            .in_block(block)
            .store(window, Ordering::Release);
    }
}

/// The window the pin was last asked to go behind, taken, or nothing where nothing was asked (see
/// `ask_the_pin_to_go_behind`).
///
/// Draining is the pin's own block's, which is this process's where this process holds the pin: two
/// windows of a run standing would otherwise take each other's asks, and the pin that was asked would
/// never hear it.
pub(crate) fn take_the_pin_behind_ask() -> Option<isize> {
    let window = PIN_BEHIND_ASK.swap(0, Ordering::AcqRel);

    (window != 0).then_some(window)
}

/// Publish the file a pin is showing, or `None` where no pin is up: the one fact here that is a name
/// rather than a number.
///
/// **A path is what the side that asks what is pinned cannot do without.** The update road a pin stands
/// for — a click on the listing swaps the file in the window — is a comparison of the picked file against
/// the pinned one, and a file stepped onto a standing pin never passes through an install. So the name is
/// written where it changes, in one fixed buffer of UTF-16 units, with a count that says how much of the
/// buffer it is and a count of the writes so that a reader can tell a name it read whole from one it read
/// while it was being written.
///
/// shortcut: a path longer than the buffer is published as no path, which leaves the asking side with
/// nothing rather than with another file's name — the update road is then off for that pin; upgrade if a
/// user ever meets a path that long.
pub(crate) fn publish_the_pinned_path(path: Option<&Path>) {
    // A window of its own publishes none of the tray's: the file was unpublished when the pin was
    // detached, and one swapped into the window afterwards does not publish it again (see
    // `app::detach::is_detached`).
    if crate::app::detach::is_detached() {
        return;
    }

    let facts = the_serving_facts();
    // Odd for the length of the write: a reader that sees the count move reads the name again on its next
    // tick rather than acting on the two halves of two names.
    facts.pin_path_writing.fetch_add(1, Ordering::AcqRel);

    let mut len = 0;
    let fits =
        path.is_some_and(|path| path.as_os_str().encode_wide().count() <= SEAM_PIN_PATH_UNITS);
    if let Some(path) = path.filter(|_| fits) {
        for unit in path.as_os_str().encode_wide() {
            facts.pin_path[len].store(u32::from(unit), Ordering::Relaxed);
            len += 1;
        }
    }

    facts.pin_path_len.store(len as u32, Ordering::Release);
    facts.pin_path_writing.fetch_add(1, Ordering::Release);
}

/// The file a pin is showing, as the side that owns the pin published it, or `None` where no pin is up or
/// the name was being written when it was asked for (see `publish_the_pinned_path`).
///
/// **The pin's own block is asked and no other** (see `the_pin_block`): a name is a fact of the window
/// that holds the pin, and a block that is not that window's holds a name of some other window's — the
/// window that took the pin at its detach published no name of its own to be read here (see
/// `app::detach`).
pub(crate) fn pinned_path() -> Option<PathBuf> {
    the_pin_block().and_then(a_pinned_path_of)
}

/// The name the block `facts` holds, or `None` where it holds none or was being written (see
/// `publish_the_pinned_path`).
fn a_pinned_path_of(facts: &'static ViewerFacts) -> Option<PathBuf> {
    let before = facts.pin_path_writing.load(Ordering::Acquire);
    if before & 1 == 1 {
        return None;
    }

    let len = facts.pin_path_len.load(Ordering::Acquire) as usize;
    if len == 0 || len > SEAM_PIN_PATH_UNITS {
        return None;
    }

    let units: Vec<u16> = (0..len)
        .map(|at| facts.pin_path[at].load(Ordering::Relaxed) as u16)
        .collect();
    if facts.pin_path_writing.load(Ordering::Acquire) != before {
        return None;
    }

    Some(PathBuf::from(OsString::from_wide(&units)))
}

/// The number a path is published as, asked the same way on both sides so that the two answers are the
/// same answer (see `PIN_FILE_IN_FLIGHT`).
///
/// A path that hashes to nothing is published as nothing in flight: the press that costs is a pick
/// offered twice, and the other way round — nothing in flight read as a file — is a pick dropped.
fn fingerprint_of(path: &Path) -> u64 {
    use std::hash::{Hash, Hasher};

    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    path.hash(&mut hasher);
    hasher.finish().max(1)
}

/// A box on the screen published whole, and read without a lock: the one fact here that does not fit an
/// atomic on its own.
///
/// Four edges are 128 bits and an atomic is 64, so the box is packed — four signed 16-bit fields in one
/// word — rather than published edge by edge. The reader is why: the wheel hook runs inside a
/// system-wide callback that every wheel message on the desktop waits for, and a box read half from one
/// publish and half from the next could claim a notch over a listing no preview covers, or refuse one
/// over a preview that does. The coordinate space is Windows' own, which is a 16-bit one in each axis —
/// the virtual screen is at most 32768 pixels across and its origin no further out than that — so a box
/// that somehow did not fit is published as no box at all, which costs a notch and never claims one.
/// the notch that was refused, which is never a notch claimed twice.
///
/// It is a handle over that word rather than the word itself, for the reason a `Fact` is (see
/// `ViewerFacts`): the packing below is this type's, and where the packed word lives is a block's.
pub(crate) struct PublishedRegion(fn(&'static ViewerFacts) -> &'static AtomicU64);

/// What a `PublishedRegion` holds where it names no box. It decodes to `(-1, -1, -1, -1)`, an empty box
/// no point is inside, so even a box of its own that happened to pack to this answers the same way.
const NO_REGION: u64 = u64::MAX;

impl PublishedRegion {
    const fn of(home: fn(&'static ViewerFacts) -> &'static AtomicU64) -> Self {
        Self(home)
    }

    /// This box's word, in the block the hovers are served from (see `the_serving_facts`).
    fn cell(&self) -> &'static AtomicU64 {
        (self.0)(the_serving_facts())
    }

    /// Publish `region` whole, or withdraw the box where there is none.
    pub(crate) fn publish(&self, region: Option<(i32, i32, i32, i32)>) {
        self.cell().store(pack_a_box(region), Ordering::Release);
    }

    /// The box as it was last published, or `None` where there is none.
    pub(crate) fn get(&self) -> Option<(i32, i32, i32, i32)> {
        unpack_a_box(self.cell().load(Ordering::Acquire))
    }

    /// This box's word, in the block `block` is (see `Fact::in_block`).
    pub(crate) fn in_block(&self, block: &'static ViewerFacts) -> &'static AtomicU64 {
        (self.0)(block)
    }

    /// The viewer of this run whose copy of this box stands over `point`, if one does: the window a notch
    /// turned at `point` was turned over, and the block it is counted into (see `wheel_input`).
    ///
    /// **The window and the count are one answer**, which is why the block is answered rather than the
    /// box: a notch over a detached pin is that window's, and a fresh viewer standing beside it publishes
    /// no box over the point at all.
    pub(crate) fn the_block_over(&self, point: (i32, i32)) -> Option<&'static ViewerFacts> {
        let (x, y) = point;

        asked_blocks().iter().flatten().copied().find(|block| {
            unpack_a_box((self.0)(block).load(Ordering::Acquire)).is_some_and(
                |(left, top, right, bottom)| x >= left && x < right && y >= top && y < bottom,
            )
        })
    }
}

/// `region` as the one word `PublishedRegion` keeps it in, or `NO_REGION` where there is none to pack:
/// a box with an edge outside the 16-bit screen this is packed in is published as no box at all.
fn pack_a_box(region: Option<(i32, i32, i32, i32)>) -> u64 {
    region
        .and_then(|(left, top, right, bottom)| {
            let edge = |value: i32| i16::try_from(value).ok().map(|value| value as u16 as u64);
            Some(edge(left)? | edge(top)? << 16 | edge(right)? << 32 | edge(bottom)? << 48)
        })
        .unwrap_or(NO_REGION)
}

/// The box `packed` holds, or `None` where it holds none.
fn unpack_a_box(packed: u64) -> Option<(i32, i32, i32, i32)> {
    if packed == NO_REGION {
        return None;
    }

    let edge = |shift: u32| ((packed >> shift) as u16 as i16) as i32;
    Some((edge(0), edge(16), edge(32), edge(48)))
}

/// The box a wheel notch over a pin's volume control belongs to: the open popup's panel and the button
/// it hangs from, or just the button while it is closed.
pub(crate) static PIN_VOLUME_REGION: PublishedRegion =
    PublishedRegion::of(|facts| &facts.pin_volume_region);

/// The box a wheel notch over a scrollable text preview belongs to: the journey from where the hover was
/// opened to the preview itself, or the pinned window's box where what is pinned is a page.
pub(crate) static TEXT_SCROLL_REGION: PublishedRegion =
    PublishedRegion::of(|facts| &facts.text_scroll_region);

/// The same box for a page that fits the one it is drawn in: there is no scroll there for Explorer to
/// lose, so a `Ctrl` and a wheel over one is the page's own size rather than its place (see
/// `text_size_after_notches`).
pub(crate) static TEXT_SIZE_REGION: PublishedRegion =
    PublishedRegion::of(|facts| &facts.text_size_region);

/// The box of a pinned picture, which a notch zooms.
pub(crate) static PIN_ZOOM_REGION: PublishedRegion =
    PublishedRegion::of(|facts| &facts.pin_zoom_region);

/// The band a pinned model is drawn in, which a notch is handed to the model's own camera for.
pub(crate) static MODELS_BAND: PublishedRegion = PublishedRegion::of(|facts| &facts.models_band);

/// The band a pinned sound's card grows under it, which a notch steps its rows by.
pub(crate) static PIN_PLAYLIST_BAND: PublishedRegion =
    PublishedRegion::of(|facts| &facts.pin_playlist_band);

/// The box the pin's window stands in, published once a tick and read by the hook's own press reads: a
/// press inside it is the pin's own whatever is under the point, and a pin dragged under a hand that did
/// not move is a box that changed (see `preview_window_rect`).
pub(crate) static PIN_WINDOW_BOX: PublishedRegion =
    PublishedRegion::of(|facts| &facts.pin_window_box);

/// The box the pin's window stands in, or nothing where no pin is up: the box a press is measured against
/// (see `PIN_WINDOW_BOX`).
///
/// It is published rather than read off the desktop because the window is not this process's: a handle
/// belongs to the process that made it, and what the hook needs is where that window is, not its handle.
/// **Every block is asked**, because a pin dragged under a hand is a box of whichever window holds it.
pub(crate) fn preview_window_rect() -> Option<(i32, i32, i32, i32)> {
    asked_blocks()
        .iter()
        .flatten()
        .find_map(|block| unpack_a_box(PIN_WINDOW_BOX.in_block(block).load(Ordering::Acquire)))
}

/// Notches turned over that band and not yet handed to the model: how many, and the screen point the last
/// one was turned at, packed as two 16-bit fields.
///
/// **A value rather than a message, and the one place here where that matters**: a notch is counted inside
/// `wheel_hook_proc`, the system-wide mouse callback every mouse message on the desktop waits for, so what
/// is written there is an atomic and nothing else — a channel send would be the first lock and the first
/// allocation on that path. The two are one fact in two atomics, and only the count is exact: the point is
/// the last one turned and is asked of the engine only to clamp the notch into its own client, so a point
/// paired with a notch turned a hair earlier answers the same way.
pub(crate) static MODELS_NOTCHES: Fact<AtomicI32> = Fact::of(|facts| &facts.models_notches);
pub(crate) static MODELS_NOTCH_POINT: Fact<AtomicU64> = Fact::of(|facts| &facts.models_notch_point);

/// Hand a notch turned over a pinned model's band to the model — f3d's wheel is its camera's zoom, so the
/// notch is the engine's and nothing of this app's (see `MODELS_BAND` and `f3d_player`).
///
/// The count is added and the loop woken; what the engine's window is and where its client is are asked on
/// the loop's own thread, where the answer may be looked up (see `take_the_models_notch`). It is counted
/// in the block whose band the notch was turned over, and in no other: the model it is for is that
/// window's, and a viewer whose tick never publishes a band has no notch to hand anybody.
pub(crate) fn post_a_notch_to_the_model(block: &'static ViewerFacts, x: i32, y: i32, notches: i32) {
    let (Ok(x), Ok(y)) = (i16::try_from(x), i16::try_from(y)) else {
        // A point that is not inside a screen this app can name is not a notch over anything of its own:
        // the box it was turned in was read off the same desktop, so this is a display that moved between
        // the two readings rather than a notch to post anywhere.
        return;
    };

    MODELS_NOTCH_POINT
        .in_block(block)
        .store(x as u16 as u64 | (y as u16 as u64) << 16, Ordering::Release);
    MODELS_NOTCHES
        .in_block(block)
        .fetch_add(notches, Ordering::AcqRel);
    wake_preview_loop();
}

/// The notches waiting to be handed to the model, and the point the last of them was turned at (see
/// `MODELS_NOTCHES`).
pub(crate) fn take_the_models_notch() -> Option<((i32, i32), i32)> {
    let notches = MODELS_NOTCHES.swap(0, Ordering::AcqRel);
    if notches == 0 {
        return None;
    }

    let point = MODELS_NOTCH_POINT.load(Ordering::Acquire);
    let edge = |shift: u32| ((point >> shift) as u16 as i16) as i32;
    Some(((edge(0), edge(16)), notches))
}

/// Pin-key presses seen since this was last drained, one per press of the key or of the mouse button it
/// is bound to.
///
/// Both low-level hooks count it — the keyboard hook for keys, and the mouse hook for `mouse3` through
/// `mouse5`, which never reach the keyboard — and the preview loop drains it once a tick. It is here
/// rather than beside the counter it is added to because its reader is the other side.
///
/// **It is one of the two facts the tray writes into every block a viewer stands in** (see
/// `note_a_pin_press`), because the press is the run's and what it means is each window's own: the pin
/// it raises is the hover's, and the bubble it brings back is the pin's that has one.
pub(crate) static PIN_PRESSES: Fact<AtomicU32> = Fact::of(|facts| &facts.pin_presses);

/// Preview-shortcut presses seen over Explorer since this was last drained.
///
/// The run's own rather than a block's: the setting a press turns is one setting, and a count in every
/// block would be a preview loop reading it in each of the windows standing, which is finding the flip
/// made twice (see `take_preview_toggle_presses`).
pub(crate) static TOGGLE_PRESSES: RunFact<AtomicU32> = RunFact::of(|facts| &facts.toggle_presses);

/// Wheel notches turned where no preview claimed them, counted by the run's one wheel hook.
///
/// **The hook and the loop that reads this are in two processes.** The wheel watcher is the viewer's, a
/// hook being a process's own, and the loop that re-resolves the item under a parked pointer after a
/// scroll is the tray's (see `shell::wheel_input`): a count kept beside the hook is one no other process
/// can see, which is a scroll that refreshes nothing. So it is the run's own count, written by whichever
/// process holds the hook.
pub(crate) static WHEEL_TICKS: RunFact<AtomicU64> = RunFact::of(|facts| &facts.wheel_ticks);

/// Count one wheel notch for the run (see `WHEEL_TICKS`).
pub(crate) fn count_a_wheel_notch() {
    WHEEL_TICKS.fetch_add(1, Ordering::Relaxed);
}

/// Wheel notches the run's hook has seen so far (see `WHEEL_TICKS`).
pub(crate) fn wheel_ticks_seen() -> u64 {
    WHEEL_TICKS.load(Ordering::Relaxed)
}

/// Pin-key presses since this was last called (see `PIN_PRESSES`).
pub(crate) fn take_presses() -> u32 {
    PIN_PRESSES.swap(0, Ordering::AcqRel)
}

/// Preview-shortcut presses since this was last called (see `TOGGLE_PRESSES`).
pub(crate) fn take_preview_toggle_presses() -> u32 {
    TOGGLE_PRESSES.swap(0, Ordering::AcqRel)
}

/// The pin key's main keys that are held, as a bit per chord slot.
///
/// It is what the preview side reads as "the hand is still on the press that took this pin up": the
/// activation Explorer dispatches for a middle press lands after this app's raise as often as before it,
/// so the band a raise gives back is held for the hand rather than for a duration (see
/// `pin_lifecycle::A_PRESS_BAND_IS_HELD_FOR`). The run's rather than a window's, because a hand is one
/// hand and the hook that counts an edge reads it back twice a message (see `the_pin_key_down`).
pub(crate) static PIN_MAIN_DOWN: RunFact<AtomicU32> = RunFact::of(|facts| &facts.pin_main_down);

/// Whether a key of the pin's own is down right now — a mouse button it is bound to included (see
/// `PIN_MAIN_DOWN`).
pub(crate) fn the_pin_key_is_down() -> bool {
    PIN_MAIN_DOWN.load(Ordering::Acquire) != 0
}

/// The keys a standing pin answers, as the commands they mean.
///
/// The table is here because the two sides have to agree about it and neither owns it: the hook reads it
/// to know which keys are the pin's before it may swallow one — `TranslateMessage` runs after every
/// low-level hook has returned, so a swallowed key never becomes a character — and the pin's own window
/// procedure reads it to answer the keys that reach it (see
/// `preview_window::pinned_key_message_command`). A key listed on one side and not the other is a key one
/// of them acts on and the other one throws away.
pub(crate) const PIN_KEY_COMMANDS: [(i32, u8); 13] = [
    (0x25, 0),  // VK_LEFT   -> Previous
    (0x27, 1),  // VK_RIGHT  -> Next
    (0x1B, 2),  // VK_ESCAPE -> Close
    (0x20, 3),  // VK_SPACE  -> TogglePlayback
    (0x54, 4),  // VK_T      -> NextSubtitle
    (0x26, 5),  // VK_UP     -> CaretUp
    (0x21, 6),  // VK_PRIOR  -> PageBack
    (0x28, 7),  // VK_DOWN   -> CaretDown
    (0x22, 8),  // VK_NEXT   -> PageOn
    (0x24, 9),  // VK_HOME   -> PageTop
    (0x23, 10), // VK_END    -> PageBottom
    (0x72, 11), // VK_F3     -> FindNext
    (0x0D, 12), // VK_RETURN -> PlayRow
];

/// How many commands `PIN_KEY_COMMANDS` names, which is how many counters are kept for them.
///
/// Named rather than written as the `13` of an array length so that a command added to the table above is
/// a command this count has to be told about, rather than one whose presses are counted into a counter
/// that was never reserved for it.
pub(crate) const PIN_COMMAND_COUNT: usize = 13;

/// How many Win chords the pin answers, which is how many counters are kept for them: the pin's own
/// minimize, and its window carried to the display on either side, in that order (see
/// `take_pin_win_chords` and `key_input::pin_win_chord`).
pub(crate) const PIN_WIN_CHORD_COUNT: usize = 3;

/// Which command of `PIN_KEY_COMMANDS` a virtual key is listed under, or `None` for a key this app does
/// not answer.
pub(crate) fn pin_key_command(vk: i32) -> Option<u8> {
    PIN_KEY_COMMANDS
        .iter()
        .find(|(key, _)| *key == vk)
        .map(|(_, command)| *command)
}

/// Whether a message about a key a standing pin answers is one the pin's own window would act on.
///
/// The class of the message is part of the question rather than an accident of which arm of a window
/// procedure it arrived on: a pinned window answers a key pressed plainly and answers no chord at all
/// (`WM_SYSKEYDOWN` is swallowed there so that `Alt+F4` cannot close it and `Alt+Tab` cannot leave it),
/// so **the hook must not count a system key either**. It did, and the difference was a chord acted on for
/// a caret this app is in: `Ctrl`+Left walked the pin's folder and `Ctrl`+Space held its film, each of them
/// for a chord the pin's own window had just thrown away.
pub(crate) fn pin_key_message_acts(message: u32) -> bool {
    message == windows::Win32::UI::WindowsAndMessaging::WM_KEYDOWN
}

/// Whether a window belongs to a viewer this tray started: the hover's window, the window a pin stands
/// in, the host a film is drawn in, a bubble or the menu it carries, a taskbar stand-in, a page's source.
///
/// **It is a question about the process the window was made in**, not a list of the windows we made: a
/// list would be published by the side that makes windows, and that side makes many of them, so a list is
/// a leak every time one is made. The process a window belongs to is the one thing about it that does not
/// change as the preview side changes what it draws, and the registry below is the tray's own record of
/// which process that is — the viewer it began, or itself where it draws its own previews.
///
/// **A run has more than one of them once a pin has been detached.** The window that left is still a
/// window of this app's standing over the desktop, so it is recorded beside the one the hovers are drawn
/// in (see `DETACHED_VIEWER_PIDS`), and a press on it is judged the way it was judged before the pin
/// became its own.
///
/// It is asked of a press the hook reads, so that a click on a pinned window is read out of the listing
/// that window is standing on rather than taken as a click on something that is not a listing at all (see
/// `click_is_over_a_listing`).
pub(crate) fn window_is_a_viewers(window: isize) -> bool {
    if window == 0 {
        return false;
    }

    let mut pid = 0u32;
    // Safety: asks the window's own thread for the process it was made in; nothing is dereferenced.
    unsafe {
        GetWindowThreadProcessId(HWND(window as *mut core::ffi::c_void), Some(&mut pid));
    }

    pid != 0 && (pid == the_viewer_process() || is_a_window_of_its_own(pid))
}

/// The process this run's previews are drawn in, or `0` where none is recorded.
///
/// It is the viewer the tray began. **Nothing recorded is this process's own**, which is what the call this
/// replaced answered: a run that draws its own previews has no second process to name, and that is every
/// fallback run and every test.
///
/// Written before the threads that ask are begun, and only again when a viewer that died is begun in its
/// place; read on every press the hook reads, which is why it is one atomic and not a lock.
static THE_VIEWER: AtomicU32 = AtomicU32::new(0);

/// Note the process this run's previews are drawn in (see `THE_VIEWER`).
pub(crate) fn note_the_viewer_process(pid: u32) {
    THE_VIEWER.store(pid, Ordering::Release);
}

/// How many windows of this run's own may be standing beside the one the hovers are drawn in.
///
/// A fixed number of slots rather than a list that grows, because the reader takes no lock: the question
/// is asked of a press on the hook's own path, and a lock that reader waited on would be a hook that stops
/// answering Explorer behind a preview thread (`THE_VIEWER` is an atomic for the same reason).
///
/// shortcut: a run that detaches more windows than this stops recognising the oldest of them, which reads
/// a press on that window as a press on a window that is not this app's rather than as a press over the
/// listing it stands on; upgrade to a set that grows if a user ever meets it.
const DETACHED_VIEWERS: usize = 16;

/// The processes of this run that are windows of their own: every viewer whose pin was detached, and
/// which draws its own window and nothing else from then on (see `window_is_a_viewers` and `app::detach`).
static DETACHED_VIEWER_PIDS: [AtomicU32; DETACHED_VIEWERS] =
    [const { AtomicU32::new(0) }; DETACHED_VIEWERS];

/// Which slot the next window of this run's own is written in: the record is a ring, and the oldest of a
/// run that detaches more windows than there are slots is the one that is lost.
static NEXT_DETACHED_VIEWER: AtomicUsize = AtomicUsize::new(0);

/// Note a window of this run's own, standing beside the one the hovers are drawn in (see
/// `DETACHED_VIEWER_PIDS`).
pub(crate) fn note_a_detached_viewer(pid: u32) {
    if pid == 0 {
        return;
    }

    let at = NEXT_DETACHED_VIEWER.fetch_add(1, Ordering::Relaxed) % DETACHED_VIEWERS;
    DETACHED_VIEWER_PIDS[at].store(pid, Ordering::Release);
}

/// Whether the process this run's previews are drawn in has become a window of its own: the viewer the
/// hovers are served by has left, and one is owed (see `app::detach` and `app::viewer`).
pub(crate) fn the_viewer_has_left() -> bool {
    let noted = THE_VIEWER.load(Ordering::Acquire);
    noted != 0 && is_a_window_of_its_own(noted)
}

/// Whether `pid` is a process of this run that has a window of its own standing (see
/// `DETACHED_VIEWER_PIDS`).
fn is_a_window_of_its_own(pid: u32) -> bool {
    DETACHED_VIEWER_PIDS
        .iter()
        .any(|slot| slot.load(Ordering::Acquire) == pid)
}

fn the_viewer_process() -> u32 {
    let noted = THE_VIEWER.load(Ordering::Acquire);
    if noted != 0 {
        return noted;
    }

    // Safety: asks the kernel for this process's own id; nothing is dereferenced.
    unsafe { GetCurrentProcessId() }
}

// What the two sides ask each other for by name, in one doorway rather than across the modules that
// happen to define them. A symbol here is a crossing: the tray side calls the preview side through this
// list, and the preview side calls the tray side through the rest of this module. What makes the step
// after this one a change of one module is that every one of these is named from here, so a body that
// becomes a message, a published fact or a round trip has no call site to follow it.
pub(crate) use crate::shell::explorer_hook::{view_sort_of, SortKey, ViewSort};
pub(crate) use crate::shell::tray::FONT_SIZE_CHOICES;
pub(crate) use crate::ui::preview_window::{
    hide_preview, monitor_dpi_from_point, note_engine_page_drag, pin_press_was_not_a_pick,
    pointer_item_box, pointer_item_holds, press_pinned_preview, preview_pointer_hold,
    publish_pin_media_press, publish_pointer_item_box, refresh_pin, refresh_preview,
    refresh_preview_types, refresh_render_html, request_pin_end, request_pin_menu_dismiss,
    show_preview, show_preview_keyboard, take_pin_handed_off, take_pin_menu_pressed,
    take_pin_resumed, trim_subtitle_cache, update_pinned_preview,
};

#[cfg(test)]
mod tests {
    use super::*;
    use std::process::Stdio;
    use std::time::{Duration, Instant};
    use windows::Win32::Foundation::WAIT_OBJECT_0;

    /// **The one section a test binary can make is named for this process** (see `a_section_of_this_run`),
    /// so the two tests that make one take turns rather than resetting the other's facts under it.
    static THE_ONE_SECTION: Mutex<()> = Mutex::new(());

    fn the_one_section() -> std::sync::MutexGuard<'static, ()> {
        THE_ONE_SECTION
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    /// The variable that makes a run of the transport test the second process rather than the first, and
    /// the pid of the run it draws for in it.
    const THE_SECOND_PROCESS: &str = "RHP_SEAM_SECOND_PROCESS";

    /// Which of a tray's two viewers a run of the block test is: the two stand at once, and each publishes
    /// numbers of its own (see `two_viewers_stand_in_blocks_of_their_own`).
    const WHICH_VIEWER: &str = "RHP_SEAM_BLOCK";

    /// How long the first process waits on anything the second one owes it.
    const AN_ANSWER: Duration = Duration::from_secs(5);

    /// How long the second process's own waits are given: the bound a wake that never came sits out, and
    /// the whole of what a broken seam can cost this test.
    const A_WAKE: u32 = 2000;

    /// How long the first process paces before it signals, so that the wait it is signalling is one that is
    /// already armed: a wake answered before that wait exists proves the ring, not the wake. The second
    /// process arms it within microseconds of publishing the fact above, and the whole test is kept short —
    /// the window tests beside it are the suite's own flake and a long block here is what tips them over.
    const AN_ARMED_WAIT: Duration = Duration::from_millis(30);

    /// How often the first process looks at a fact the second one publishes.
    const A_LOOK: Duration = Duration::from_millis(1);

    /// Every preview message a slot carries, with something in every field of each.
    fn carried_preview_messages() -> Vec<PreviewMessage> {
        let path = PathBuf::from("listing/film.mkv");

        vec![
            PreviewMessage::Show(path.clone(), 10, 20, Some(AvoidRegion::text((1, 2, 3, 4)))),
            PreviewMessage::Show(
                path.clone(),
                10,
                20,
                Some(AvoidRegion::column((-1, -2, -3, -4))),
            ),
            PreviewMessage::Show(path.clone(), 10, 20, None),
            PreviewMessage::ShowKeyboard(path.clone(), 1, 2, 3, 4, Some((5, 6, 7, 8)), true),
            PreviewMessage::ShowKeyboard(path.clone(), 1, 2, 3, 4, None, false),
            PreviewMessage::Hide,
            PreviewMessage::Refresh,
            PreviewMessage::RefreshTypes,
            PreviewMessage::StopEngines,
            PreviewMessage::HideHtmlPage,
            PreviewMessage::TrimCaches,
            PreviewMessage::KillStrayPlayer,
            PreviewMessage::Pin {
                path: path.clone(),
                rect: (9, 10, 11, 12),
                arriving: true,
            },
            PreviewMessage::PinBox((13, 14, 15, 16)),
            PreviewMessage::PinUpdate(path.clone()),
            PreviewMessage::PinPress { window: 4242 },
            PreviewMessage::PinNoPick { window: -4242 },
            PreviewMessage::PinChanged,
            PreviewMessage::TrayVolume { audio: true },
            PreviewMessage::TrayTextSize { percent: 175 },
            PreviewMessage::ForgetVideoGeometry,
            PreviewMessage::KeepExtraction {
                path: Some(path.clone()),
            },
            PreviewMessage::KeepExtraction { path: None },
            PreviewMessage::OfficeRenderReady {
                path: path.clone(),
                generation: 7,
                ok: false,
            },
            PreviewMessage::VideoProbed {
                path: path.clone(),
                generation: 8,
            },
            PreviewMessage::MeasureProbed {
                path: path.clone(),
                size: Some((1920, 1080)),
            },
            PreviewMessage::MeasureProbed {
                path: path.clone(),
                size: None,
            },
            PreviewMessage::VideoSubtitlesReady(path.clone()),
            PreviewMessage::MagickReady {
                path: path.clone(),
                generation: 9,
                ok: true,
            },
            PreviewMessage::PeazipReady {
                path: path.clone(),
                generation: 10,
                ok: true,
            },
            PreviewMessage::PinRestore,
        ]
    }

    /// What a second process of this binary said, both streams: what it held is a panic, and the harness
    /// prints that to its own output rather than to the one a panic would have gone to.
    fn what_it_said(ended: &std::process::Output) -> String {
        format!(
            "{}{}",
            String::from_utf8_lossy(&ended.stdout),
            String::from_utf8_lossy(&ended.stderr)
        )
    }

    /// A message through a slot and back is the message it went in as, compared as a slot again: what a
    /// conversion drops shows up here as a difference rather than as a preview drawn from less.
    fn a_preview_round_trip(message: PreviewMessage) {
        let carried = SeamMessage::of_preview(&message).expect("a message that fits a slot");
        let back = carried.as_preview().expect("a message back");
        assert_eq!(SeamMessage::of_preview(&back), Some(carried));
    }

    #[test]
    fn every_carried_preview_message_survives_a_slot() {
        for message in carried_preview_messages() {
            a_preview_round_trip(message);
        }
    }

    #[test]
    fn every_tray_message_survives_a_slot() {
        for message in [
            TrayMessage::TogglePinAlwaysOnTop,
            TrayMessage::TogglePreviewEnabled,
            TrayMessage::PinClick("pin: the press named no file".to_owned()),
            TrayMessage::ViewerDetached(4242),
        ] {
            let carried = SeamMessage::of_tray(&message).expect("a notice that fits a slot");
            let back = carried.as_tray().expect("a notice back");
            assert_eq!(SeamMessage::of_tray(&back), Some(carried));
        }
    }

    #[test]
    fn a_region_and_the_column_it_is_come_back_as_they_went() {
        let message = PreviewMessage::Show(
            PathBuf::from("listing/note.txt"),
            10,
            20,
            Some(AvoidRegion::column((1, 2, 3, 4))),
        );
        let back = SeamMessage::of_preview(&message)
            .expect("a message that fits a slot")
            .as_preview()
            .expect("a message back");

        let PreviewMessage::Show(path, x, y, Some(region)) = back else {
            panic!("a preview of a hovered file, which came with a region");
        };
        assert_eq!(path, PathBuf::from("listing/note.txt"));
        assert_eq!((x, y), (10, 20));
        assert_eq!(region.region, (1, 2, 3, 4));
        assert!(region.column);
    }

    #[test]
    fn a_path_that_does_not_fit_a_slot_is_a_message_that_is_not_carried() {
        let long = PathBuf::from("a".repeat(SEAM_TEXT_UNITS + 1));

        assert!(SeamMessage::of_preview(&PreviewMessage::PinUpdate(long)).is_none());
    }

    /// **The facts of a pin are its own block's, and a name is published whole or not at all.**
    ///
    /// A name published by a block with no pin up is no pin's file, and a file in flight there is no
    /// pin's file in flight: a window that left the ring mid-load holds both, and reading either as the
    /// pin's is every pick of that file refused in silence (see `the_pin_block` and
    /// `pin_is_on_its_way_to`). The facts are this process's own copy in a test (see `LOCAL_FACTS`), so
    /// what is written here is read back here — and put at rest at the end, since the copy is one for the
    /// whole binary.
    #[test]
    fn the_facts_of_a_pin_are_its_own_blocks_and_a_name_is_whole_or_nothing() {
        let facts = the_run_facts();
        let path = Path::new("listing/a note.txt");
        let other = Path::new("listing/on its way.png");
        let at_rest = |block: &'static ViewerFacts| {
            block.pin_up.store(false, Ordering::Release);
            block.pin_file_in_flight.store(0, Ordering::Release);
            block.pin_path_len.store(0, Ordering::Release);
        };

        facts.viewers[0]
            .owner
            .store(std::process::id(), Ordering::Release);
        facts.viewers[1].owner.store(4_242, Ordering::Release);
        at_rest(&facts.viewers[0]);
        at_rest(&facts.viewers[1]);
        publish_the_pinned_path(None);
        assert_eq!(pinned_path(), None);

        publish_the_pinned_path(Some(path));
        assert_eq!(
            pinned_path(),
            None,
            "a name published by a block with no pin up is no pin's file"
        );

        PIN_UP.store(true, Ordering::Release);
        assert_eq!(
            pinned_path().as_deref(),
            Some(path),
            "and the pin's own block is the one that names the pin's file"
        );

        // A name longer than the facts can hold is no name rather than half of one.
        let long = PathBuf::from("a".repeat(SEAM_PIN_PATH_UNITS + 1));
        publish_the_pinned_path(Some(long.as_path()));
        assert_eq!(pinned_path(), None);

        // **A window that left the ring mid-load**: a block standing of its own with a file in flight and
        // no pin up. The file it names is no pin's, and a pick of it is not dropped on its account.
        facts.viewers[1]
            .pin_file_in_flight
            .store(fingerprint_of(other), Ordering::Release);
        assert!(
            !pin_is_on_its_way_to(other),
            "a file in flight in a block with no pin is no pin's file in flight"
        );

        // The pin's own block is the one asked, and its own flag is the one that answers.
        facts.viewers[0]
            .pin_file_in_flight
            .store(fingerprint_of(other), Ordering::Release);
        assert!(
            pin_is_on_its_way_to(other),
            "the pin's own flag is the one the update road is refused by"
        );

        // And a pin up with nothing in flight answers nothing: a flag is not a file on its way.
        facts.viewers[0]
            .pin_file_in_flight
            .store(0, Ordering::Release);
        assert!(!pin_is_on_its_way_to(other));

        publish_the_pinned_path(None);
        at_rest(&facts.viewers[0]);
        at_rest(&facts.viewers[1]);
        facts.viewers[0].owner.store(0, Ordering::Release);
        facts.viewers[1].owner.store(0, Ordering::Release);
    }

    #[test]
    fn the_ask_to_go_behind_reaches_every_viewer_and_is_taken_once() {
        let facts = the_run_facts();
        let stood_in = facts.blocks(&facts.viewers[0], false);

        ask_the_pin_to_go_behind_in(stood_in, 0x1234);
        assert!(
            stood_in
                .iter()
                .flatten()
                .all(|block| PIN_BEHIND_ASK.in_block(block).load(Ordering::Acquire) == 0x1234),
            "the window the user came back to is written into every window of the run's"
        );

        assert_eq!(
            take_the_pin_behind_ask(),
            Some(0x1234),
            "and the pin's own viewer takes it"
        );
        assert_eq!(
            take_the_pin_behind_ask(),
            None,
            "once: a second tick has nothing to lower"
        );
    }

    #[test]
    fn a_viewer_that_is_named_is_the_process_a_window_is_judged_by() {
        note_the_viewer_process(4242);
        assert_eq!(the_viewer_process(), 4242);
        assert!(!window_is_a_viewers(0), "no window belongs to anybody");

        // And back to nothing recorded, which is this process's own previews and what every other test
        // in this binary reads.
        note_the_viewer_process(0);
        assert_eq!(the_viewer_process(), std::process::id());
    }

    #[test]
    fn a_ring_keeps_its_order_and_drops_what_it_cannot_hold() {
        let ring = SeamRing::new();
        assert_eq!(ring.pop(), None);

        for press in 0..SEAM_RING_SLOTS as i64 {
            assert!(ring.push(SeamMessage::of(SeamMessage::HIDE, &[press]), "test"));
        }
        assert!(
            !ring.push(SeamMessage::of(SeamMessage::HIDE, &[]), "test"),
            "a full ring keeps nothing"
        );

        for press in 0..SEAM_RING_SLOTS as i64 {
            let message = ring.pop().expect("a message that was pushed");
            assert_eq!(message.numbers[0], press);
        }
        assert_eq!(ring.pop(), None);
    }

    #[test]
    fn a_section_carries_the_facts_and_a_ring_of_the_messages() {
        let _the_section_is_mine = the_one_section();

        // Mapped and not published: every other test in this binary reads and writes the facts as it finds
        // them, and moving them under it is not this test's to do.
        let Some(facts) = a_section_of_this_run() else {
            panic!("a section every run can open");
        };
        // Safety: a view of this run's own section, written with the at-rest facts and told to nobody.
        let facts = unsafe { &*facts };

        assert!(
            !facts.viewers[0].pin_up.load(Ordering::Acquire),
            "a section starts at rest"
        );
        assert_eq!(
            facts.viewers[0].pin_zoom_region.load(Ordering::Acquire),
            NO_REGION,
            "a box that was never published is no box, not the one at the origin"
        );

        // The ring is the one part of the section nothing else here touches, and a message through it is
        // the mapped bytes being written and read again.
        assert!(facts
            .preview_ring
            .push(SeamMessage::of(SeamMessage::HIDE, &[7]), "test"));
        assert_eq!(
            facts.preview_ring.pop().map(|message| message.numbers[0]),
            Some(7)
        );
    }

    /// The pin's own walk is answered on the channel of the process that asked, and it can be answered on
    /// nothing else: a step names a playlist and the spawner behind it, and a fixed slot holds neither. Put
    /// through the rings instead, every **Next**/**Previous** press would answer nothing — the arc would
    /// turn out `PIN_JOB_GIVEUP` and the pin would never be shown the file it stepped onto (see
    /// `send_preview` and `next_preview_message`).
    #[test]
    fn the_pin_planners_answer_has_a_road_a_slot_does_not() {
        // It sends through the process's own channel and wakes this run's preview event, which are the two
        // things every other test in this binary that sends a preview also touches: taking the section's
        // turn is what keeps it from sharing a sender with one of them (see `the_one_section`).
        let _the_section_is_mine = the_one_section();

        let answer = PreviewMessage::PinAnswered(PinPlanned::OpenWith {
            path: PathBuf::from("listing/a note.txt"),
            name: "Notepad".to_string(),
        });

        // A slot cannot carry it, which is the whole of why it must not be offered to the ring.
        assert!(
            SeamMessage::of_preview(&answer).is_none(),
            "the pin's answer is not a slot's to carry"
        );

        // And the channel is, on the road a loop reads it by.
        let (tx, rx) = std::sync::mpsc::channel();
        let previous = PREVIEW_SENDER
            .lock()
            .ok()
            .and_then(|mut held| held.replace(tx));
        send_preview_here(answer);
        if let Ok(mut held) = PREVIEW_SENDER.lock() {
            *held = previous;
        }

        let Some(PreviewMessage::PinAnswered(PinPlanned::OpenWith { path, name })) =
            rx.try_recv().ok()
        else {
            panic!("the answer a process made is one that process can read");
        };
        assert_eq!(path, PathBuf::from("listing/a note.txt"), "the name, whole");
        assert_eq!(name, "Notepad", "and the answer it carried");
    }

    /// **A press the tray's hook answers is a press the process holding the pin hears**, which `pinned()`
    /// cannot answer from the tray: the copy it reads is written by the process that installed the pin, and
    /// in a run of two processes that is never this one. Asked of that copy alone, every press answered in
    /// a listing was answered with nothing — no band held across the listing's own activation, and no pin
    /// put back behind the window the press was made in (see `press_pinned_preview` and
    /// `pin_press_was_not_a_pick`).
    #[test]
    fn a_press_in_the_listing_reaches_a_pin_the_other_process_is_holding() {
        // It replaces the process's own preview sender and publishes a pin-up fact that the whole binary
        // reads: the section's turn is what keeps it from doing both under another test (see
        // `the_one_section`).
        let _the_section_is_mine = the_one_section();

        let (tx, rx) = std::sync::mpsc::channel();
        let previous = PREVIEW_SENDER
            .lock()
            .ok()
            .and_then(|mut held| held.replace(tx));

        // The published fact and nothing else: a pin is up for the run, and the copy this process draws
        // by is untouched, which is the state the tray is in.
        let was_up = PIN_UP.load(Ordering::Acquire);
        PIN_UP.store(true, Ordering::Release);
        assert!(the_pin_is_up(), "a pin the run publishes is a pin up");

        crate::ui::preview_window::press_pinned_preview(7);
        crate::ui::preview_window::pin_press_was_not_a_pick(9);

        PIN_UP.store(was_up, Ordering::Release);
        if let Ok(mut held) = PREVIEW_SENDER.lock() {
            *held = previous;
        }

        let Some(PreviewMessage::PinPress { window }) = rx.try_recv().ok() else {
            panic!("the pick's press is asked of the pin holding it");
        };
        assert_eq!(window, 7, "and it names the window the press was made in");

        let Some(PreviewMessage::PinNoPick { window }) = rx.try_recv().ok() else {
            panic!("and a press that named no file is asked of it too");
        };
        assert_eq!(window, 9, "naming the listing the pin is put behind");
    }

    #[test]
    fn every_block_a_run_has_is_claimed_once_and_the_last_one_is_none() {
        let _the_section_is_mine = the_one_section();
        let Some(facts) = a_section_of_this_run() else {
            panic!("a section every run can open");
        };
        // Safety: a view of this run's own section, and the blocks it holds are this test's to claim.
        let facts = unsafe { &*facts };
        for block in facts.viewers.iter() {
            block.owner.store(0, Ordering::Release);
        }

        // One block a viewer, taken in the order they are free, and the pid is what one is asked by.
        let claimed: Vec<usize> = (1..=SEAM_VIEWERS as u32)
            .map(|pid| facts.claim_a_block(pid).expect("a free block of the run's"))
            .collect();
        assert_eq!(
            claimed.len(),
            SEAM_VIEWERS,
            "a block for every viewer it may have"
        );

        for pid in 1..=SEAM_VIEWERS as u32 {
            assert_eq!(
                facts.facts_of(pid).map(|block| block as *const ViewerFacts),
                Some(&facts.viewers[(pid - 1) as usize] as *const ViewerFacts),
                "the block a viewer stands in is the one its pid is asked by"
            );
        }

        // **A pool whose every block a running process holds answers none**, which is what a detach asks
        // before it flips anything (see `a_block_can_be_had`). The pids claimed above are not this process's
        // own, so whether any of them is still a process is the machine's answer rather than this test's.
        for block in facts.viewers.iter() {
            block.owner.store(std::process::id(), Ordering::Release);
        }
        assert_eq!(facts.claim_a_block(4242), None, "and none for one more");
        assert!(
            !facts.a_block_can_be_had(),
            "and no detach it would leave a viewer owing"
        );

        for block in facts.viewers.iter() {
            block.owner.store(0, Ordering::Release);
        }
    }

    /// **A block a viewer walked out of is one the next viewer walks into.** Nothing gives a block back when
    /// its viewer ends, so without the second pass of `claim_a_block` a run that lost viewers — a window
    /// detached and closed, a viewer that died — would spend its blocks and then draw nothing at all for the
    /// rest of its life, every viewer begun after that stopping at its own start-up.
    ///
    /// The child here is a viewer's own start and nothing else: it joins, takes a block and stops, which is
    /// the block a viewer that exits leaves behind. Every block of the run is then marked with that gone pid,
    /// as a run whose windows have all closed would leave them, and the viewers begun after are ones no free
    /// block is left for.
    #[test]
    fn a_run_keeps_beginning_viewers_past_the_blocks_it_has() {
        if let Some(tray) = std::env::var(THE_SECOND_PROCESS)
            .ok()
            .and_then(|pid| pid.parse().ok())
        {
            a_viewer_that_stops_once_it_has_a_block(tray);
            return;
        }

        let _the_section_is_mine = the_one_section();
        let Some(facts) = a_section_of_this_run() else {
            panic!("a section a viewer can open");
        };
        // Safety: a view of this run's own section, and the blocks it holds are this test's to mark.
        let facts = unsafe { &*facts };
        for block in facts.viewers.iter() {
            block.owner.store(0, Ordering::Release);
        }

        // The tray's own block, claimed as a run's tray claims it at its start-up.
        assert_eq!(
            facts.claim_a_block(std::process::id()),
            Some(0),
            "the tray stands in the first block of its own"
        );

        let (first, gone) = a_viewer_stops_after_its_block();
        assert_eq!(
            first, 1,
            "a viewer takes the first block the tray's leaves free"
        );
        assert!(
            !crate::app::engine_processes::is_running(gone),
            "the viewer that took it has gone"
        );
        assert_eq!(
            facts.viewers[first].owner.load(Ordering::Acquire),
            gone,
            "and its pid is still what the block says: no block is ever given back"
        );

        // Every block the run has left marked by that gone pid, a run whose windows all closed.
        for block in facts.viewers.iter().skip(1) {
            block.owner.store(gone, Ordering::Release);
        }
        assert!(facts.a_block_can_be_had(), "a block is had again");

        let (second, second_pid) = a_viewer_stops_after_its_block();
        assert_ne!(second, 0, "a viewer stands in a block the gone one held");
        assert_eq!(
            facts.viewers[second].owner.load(Ordering::Acquire),
            second_pid,
            "and the block is the one it claimed rather one it was refused"
        );

        // And the one after it: a run keeps beginning viewers against the blocks it has, not against the
        // number of viewers it has begun.
        let (third, third_pid) = a_viewer_stops_after_its_block();
        assert_ne!(third, 0, "a run keeps beginning viewers past them");
        assert_eq!(
            facts.viewers[third].owner.load(Ordering::Acquire),
            third_pid,
            "and the block is the one it claimed"
        );

        for block in facts.viewers.iter() {
            block.owner.store(0, Ordering::Release);
        }
    }

    /// Begin a viewer of this run and answer the block it stood in and its pid: the process stops as soon as
    /// it has claimed one, so the block it took is the one a viewer that has gone leaves behind (see
    /// `a_run_keeps_beginning_viewers_past_the_blocks_it_has`).
    fn a_viewer_stops_after_its_block() -> (usize, u32) {
        let child = crate::app::engine_processes::hidden_command(
            std::env::current_exe().expect("this test's own binary"),
        )
        .arg(stringify!(
            a_run_keeps_beginning_viewers_past_the_blocks_it_has
        ))
        // The block index said on the child's own stdout rather than through a fact of the section: the
        // process is gone by the time this side reads it, and what a harness captures is given up by it.
        .arg("--nocapture")
        .env(THE_SECOND_PROCESS, std::process::id().to_string())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("a second process of this test binary");
        let pid = child.id();

        let ended = child.wait_with_output().expect("the second process ended");
        let said = what_it_said(&ended);
        assert!(ended.status.success(), "the second process held: {said}");

        let at = said
            .split_whitespace()
            .skip_while(|word| *word != "block")
            .nth(1)
            .and_then(|at| at.parse().ok())
            .unwrap_or_else(|| panic!("the block it stood in was said: {said}"));
        assert_ne!(at, NO_BLOCK, "the second process claimed a block: {said}");

        (at, pid)
    }

    /// The second process: a viewer's own start and then nothing — a viewer that stops the moment it has a
    /// block, which is what leaves a block a pid that has gone still holds (see `claim_a_block`).
    fn a_viewer_that_stops_once_it_has_a_block(tray: u32) {
        let joined = join_the_seam_section(tray);
        assert!(joined, "the section the tray opened is joined");
        let at = the_block_it_stands_in().unwrap_or(NO_BLOCK);
        println!("joined 1 block {at}");
    }

    /// The place in the section's `viewers` of the block this process stands in, or nothing where it stands
    /// in none (see `THE_OWN_BLOCK`).
    fn the_block_it_stands_in() -> Option<usize> {
        let own = the_facts() as *const ViewerFacts;
        the_run_facts()
            .viewers
            .iter()
            .position(|block| std::ptr::eq(block, own))
    }

    #[test]
    fn a_ring_wraps_onto_its_first_slots() {
        let ring = SeamRing::new();

        for press in 0..(SEAM_RING_SLOTS as i64 * 3) {
            assert!(ring.push(SeamMessage::of(SeamMessage::HIDE, &[press]), "test"));
            let message = ring.pop().expect("the message just pushed");
            assert_eq!(message.numbers[0], press);
        }
    }

    /// **The seam over a real process boundary, which is the one road no other test here takes.**
    ///
    /// Every other test reads the facts from this process's own copy (see `LOCAL_FACTS`), and that is
    /// exactly why a run of two processes could lose every preview with the suite green: the section, the
    /// three events named for the tray's pid and the ring the hovers cross on are all of them a second
    /// process's business, and nothing here had ever been a second process. So this test is two. The first
    /// opens the section and the events and waits; the second — this same test binary, begun again with
    /// `THE_SECOND_PROCESS` set — joins them through the very calls a viewer makes at its own start (see
    /// `join_the_seam_section`), and answers.
    ///
    /// What it proves, each in the direction the app uses it: a fact the tray publishes is read by the
    /// viewer; a hover the tray pushes is popped by the viewer with its path and its numbers whole; the wake
    /// the tray signals ends a wait in the viewer that was already armed; and the same two in reverse — a
    /// fact and a notice the viewer publishes are read by the tray.
    #[test]
    fn a_viewer_of_another_process_draws_from_the_section_the_tray_opened() {
        if let Some(tray) = std::env::var(THE_SECOND_PROCESS)
            .ok()
            .and_then(|pid| pid.parse().ok())
        {
            a_viewer_answers(tray);
            return;
        }

        let _the_section_is_mine = the_one_section();
        let Some(facts) = a_section_of_this_run() else {
            panic!("a section a viewer can open");
        };

        // **The run's three names are opened before the second process exists**, events included: a named
        // event lives while a handle to it does, so a viewer that signalled one and exited before this side
        // had ever asked for it would leave the tray creating a fresh object nothing was ever signalled on,
        // and the wake this test is about would be waited out to the bound instead (see
        // `preview_channel_event` and `tray_channel_event`).
        let _ = preview_channel_event();
        let _ = tray_channel_event();

        // Safety: a view of this run's own section, freshly written with the at-rest facts and told to the
        // second process by its name alone; nothing else in this process writes it.
        let facts = unsafe { &*facts };

        let child = crate::app::engine_processes::hidden_command(
            std::env::current_exe().expect("this test's own binary"),
        )
        // The name of this test rather than a harness flag of its own: the second process runs this same
        // test, and `THE_SECOND_PROCESS` is what makes it the second process (see above).
        .arg(stringify!(
            a_viewer_of_another_process_draws_from_the_section_the_tray_opened
        ))
        .env(THE_SECOND_PROCESS, std::process::id().to_string())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("a second process of this test binary");

        // The second process says it is up by publishing a fact rather than by making an event of its own:
        // it is the first thing it does once it has joined, and the fact this wait is looking at. The
        // block it publishes in is the one its pid is asked by (see `facts_of`).
        let deadline = Instant::now() + AN_ANSWER;
        let mut standing = None;
        while standing.is_none() && Instant::now() < deadline {
            standing = facts.facts_of(child.id());
            std::thread::sleep(A_LOOK);
        }
        let Some(block) = standing else {
            let said = what_it_said(&child.wait_with_output().expect("the second process ended"));
            panic!("the second process stands in a block of its own: {said}");
        };

        let deadline = Instant::now() + AN_ANSWER;
        while block.preview_under_pointer.load(Ordering::Acquire) == 0 && Instant::now() < deadline
        {
            std::thread::sleep(A_LOOK);
        }
        let it_said_it_was_up = block.preview_under_pointer.load(Ordering::Acquire) != 0;

        // A fact the tray publishes, and the hover it pushes at the same point the hook's road does.
        block.pin_zoom_delta.store(7, Ordering::Release);
        let hover = PreviewMessage::Show(
            PathBuf::from("listing/a note.txt"),
            10,
            20,
            Some(AvoidRegion::column((1, 2, 3, 4))),
        );
        assert!(
            facts.preview_ring.push(
                SeamMessage::of_preview(&hover).expect("a hover that fits a slot"),
                "test"
            ),
            "the preview ring has room for one hover"
        );
        std::thread::sleep(AN_ARMED_WAIT);
        // Safety: an event of this run's own, signalled: nothing here is dereferenced. The wake the tray's
        // own `send_preview` sends, on the same event.
        unsafe {
            let _ = SetEvent(preview_channel_event());
        }

        // Safety: an event of this run's own, waited on for a bound: nothing here is dereferenced.
        let woken =
            unsafe { WaitForSingleObject(tray_channel_event(), AN_ANSWER.as_millis() as u32) };
        let notice = facts.tray_ring.pop().and_then(|slot| slot.as_tray());
        let the_fact = block.text_scroll_delta.load(Ordering::Acquire);

        let ended = child.wait_with_output().expect("the second process ended");
        let said = what_it_said(&ended);

        assert!(it_said_it_was_up, "the section is joined: {said}");
        assert!(ended.status.success(), "the second process held: {said}");
        assert_eq!(
            woken, WAIT_OBJECT_0,
            "the notice it pushed woke this wait: {said}"
        );
        assert!(
            matches!(notice, Some(TrayMessage::PinClick(..))),
            "and it is the notice it pushed: {said}"
        );
        assert_eq!(the_fact, -3, "a fact it published is read here: {said}");
    }

    /// The second process: a viewer's own start, and then the answers a viewer owes the tray.
    fn a_viewer_answers(tray: u32) {
        assert!(
            join_the_seam_section(tray),
            "the section the tray opened is joined rather than made"
        );
        let facts = the_facts();
        assert_eq!(
            facts.owner.load(Ordering::Acquire),
            std::process::id(),
            "and it stands in a block of its own, asked of by its pid"
        );
        // It is up: a fact, because a fact is the one thing both processes read without an event of their
        // own (see the wait above). Nothing here waits for the tray's signal until this is written.
        facts.preview_under_pointer.store(1, Ordering::Release);

        // Safety: an event of this run's own, named for the tray's pid, waited on for a bound: nothing
        // here is dereferenced.
        assert_eq!(
            unsafe { WaitForSingleObject(preview_channel_event(), A_WAKE) },
            WAIT_OBJECT_0,
            "the wait the tray's wake had to reach ends on that signal rather than on the bound"
        );

        // **A wake is not the tray's until the ring says so.** The event is the run's, and this is a
        // process of a test binary in which every other test that sends a preview signals the same one:
        // a wake this process did not earn is one a viewer waits out rather than reads as an empty ring
        // (see `wait_preview_channel`), and the wait above is what proves the signal reached — not what
        // proves which signal it was.
        let run = the_run_facts();
        let deadline = Instant::now() + AN_ANSWER;
        let mut popped = run.preview_ring.pop();
        while popped.is_none() && Instant::now() < deadline {
            std::thread::sleep(A_LOOK);
            popped = run.preview_ring.pop();
        }

        let Some(PreviewMessage::Show(path, x, y, Some(region))) =
            popped.and_then(|slot| slot.as_preview())
        else {
            panic!("the hover the tray pushed is in the ring");
        };
        assert_eq!(path, PathBuf::from("listing/a note.txt"), "the name, whole");
        assert_eq!((x, y), (10, 20), "and the point it was hovered at");
        assert_eq!(region.region, (1, 2, 3, 4), "and the box that came with it");
        assert!(region.column, "and that the box is a column of the view");
        assert_eq!(
            facts.pin_zoom_delta.load(Ordering::Acquire),
            7,
            "a fact the tray published is read here"
        );

        // And the other way: a fact the tray reads and a notice on the ring it drains.
        facts.text_scroll_delta.store(-3, Ordering::Release);
        assert!(
            run.tray_ring.push(
                SeamMessage::of_tray(&TrayMessage::PinClick("the second process".to_owned()))
                    .expect("a notice that fits a slot"),
                "test"
            ),
            "the tray ring has room for one notice"
        );
        // Safety: an event of this run's own, signalled: nothing here is dereferenced.
        unsafe {
            let _ = SetEvent(tray_channel_event());
        }
    }

    /// **Two viewers stand at once, and neither one's facts are the other's.**
    ///
    /// A run has two of them the moment a pin is detached: the viewer serving the hovers and the window
    /// that left the ring. Both used to publish into one record, so a notch turned over the detached pin
    /// was claimed by the hover's viewer, and the boxes a notch belongs to, the window under the pointer
    /// and the pin's own handles were whichever of the two wrote last. So this test is the tray with both
    /// of them standing: two processes of this same binary, each claiming a block of its own and
    /// publishing a number only its own block can hold, and the tray reading them by pid and writing what
    /// it owns for the run into every block a viewer stands in.
    #[test]
    fn two_viewers_stand_in_blocks_of_their_own() {
        if let Some(tray) = std::env::var(THE_SECOND_PROCESS)
            .ok()
            .and_then(|pid| pid.parse().ok())
        {
            let which = std::env::var(WHICH_VIEWER)
                .ok()
                .and_then(|which| which.parse().ok())
                .expect("a numbered viewer of the tray's two");
            two_viewers_answer(tray, which);
            return;
        }

        let _the_section_is_mine = the_one_section();
        let Some(facts) = a_section_of_this_run() else {
            panic!("a section two viewers can open");
        };
        // Safety: a view of this run's own section, and the blocks it holds are this test's to claim.
        let facts = unsafe { &*facts };

        // One section is made for the whole binary, so it is put back at rest rather than read as the test
        // beside it left it: no block stood in and nothing of that test's published.
        for block in facts.viewers.iter() {
            block.owner.store(0, Ordering::Release);
            block.pin_zoom_delta.store(0, Ordering::Release);
            block.pin_bubble_hwnd.store(0, Ordering::Release);
            block.preview_under_pointer.store(0, Ordering::Release);
            block.pin_presses.store(0, Ordering::Release);
            block.pin_end_requested.store(false, Ordering::Release);
        }

        // The run stands in a block of its own before it begins one viewer, which is what gives the two
        // that follow two blocks of their own rather than one the tray is also in (see
        // `open_the_seam_section`).
        assert!(
            facts.claim_a_block(std::process::id()).is_some(),
            "a block for the run itself"
        );

        let first = a_viewer_of_block(1);
        let second = a_viewer_of_block(2);

        // Each says it is up by publishing a number rather than by an event of its own: what this wait is
        // looking at is the fact only a viewer standing in its own block could have written.
        let deadline = Instant::now() + AN_ANSWER;
        let (one, two) = loop {
            let published = |child: &std::process::Child| {
                facts
                    .facts_of(child.id())
                    .filter(|block| block.pin_zoom_delta.load(Ordering::Acquire) != 0)
            };
            if let (Some(one), Some(two)) = (published(&first), published(&second)) {
                break (one, two);
            }
            assert!(
                Instant::now() < deadline,
                "both viewers stand in blocks of their own"
            );
            std::thread::sleep(A_LOOK);
        };

        // **No crosstalk**: each block holds the number the viewer standing in it published, whole, and
        // nothing of the other's.
        assert_eq!(
            one.pin_zoom_delta.load(Ordering::Acquire),
            -7,
            "the first's own count"
        );
        assert_eq!(
            two.pin_zoom_delta.load(Ordering::Acquire),
            -14,
            "and the second's"
        );
        assert_eq!(
            one.pin_bubble_hwnd.load(Ordering::Acquire),
            0x1000,
            "the first's own bubble"
        );
        assert_eq!(
            two.pin_bubble_hwnd.load(Ordering::Acquire),
            0x2000,
            "and the second's"
        );

        // **Hover admission is the serving viewer's answer and nobody else's** — the question the wheel
        // asks of the block whose box the notch was turned in (see `the_pointer_is_on`).
        assert!(
            the_pointer_is_on(one),
            "the pointer stands on the serving viewer's preview"
        );
        assert!(
            !the_pointer_is_on(two),
            "and on nothing the window that left the ring drew"
        );

        // **What the run writes for its windows reaches every one of them**: a press of the pin key is
        // each window's to act on, and a pin asked to come down is the pin of whichever one holds it.
        let stood_in = facts.blocks(&facts.viewers[0], false);
        count_a_pin_press_in(stood_in);
        ask_the_pin_to_come_down_in(stood_in);
        for (block, which) in [(one, "the first"), (two, "the second")] {
            assert_eq!(
                block.pin_presses.load(Ordering::Acquire),
                1,
                "the press reaches {which}"
            );
            assert!(
                block.pin_end_requested.load(Ordering::Acquire),
                "and the ask for the pin does too, at {which}"
            );
        }

        for (child, which) in [(first, "the first"), (second, "the second")] {
            let ended = child.wait_with_output().expect("the viewer ended");
            let said = what_it_said(&ended);
            assert!(ended.status.success(), "{which} viewer held: {said}");
        }
    }

    /// One of the two viewers of the block test, as the same binary begun again.
    fn a_viewer_of_block(which: u8) -> std::process::Child {
        crate::app::engine_processes::hidden_command(
            std::env::current_exe().expect("this test's own binary"),
        )
        .arg(stringify!(two_viewers_stand_in_blocks_of_their_own))
        .env(THE_SECOND_PROCESS, std::process::id().to_string())
        .env(WHICH_VIEWER, which.to_string())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("a viewer of this same test binary")
    }

    /// A viewer of the block test: its own block claimed, its own numbers published, and then the facts
    /// the tray wrote for every window standing read out of that same block.
    fn two_viewers_answer(tray: u32, which: u8) {
        assert!(
            join_the_seam_section(tray),
            "the section the tray opened is joined rather than made"
        );
        let facts = the_facts();
        assert_eq!(
            facts.owner.load(Ordering::Acquire),
            std::process::id(),
            "the block this viewer stands in is claimed for it, and asked of by its pid"
        );

        facts
            .pin_zoom_delta
            .store(-7 * i32::from(which), Ordering::Release);
        facts
            .pin_bubble_hwnd
            .store(0x1000 * isize::from(which), Ordering::Release);
        // The first of the two is the window the hovers are served from, which is what the tray's answer
        // about the thing under the pointer is made of; the second is the window that left the ring.
        facts
            .preview_under_pointer
            .store(u32::from(which == 1), Ordering::Release);

        let deadline = Instant::now() + Duration::from_millis(u64::from(A_WAKE));
        while !facts.pin_end_requested.load(Ordering::Acquire) {
            assert!(
                Instant::now() < deadline,
                "the ask the tray made for the pin reached this block"
            );
            std::thread::sleep(A_LOOK);
        }
        assert_eq!(
            facts.pin_presses.load(Ordering::Acquire),
            1,
            "and so did the press it counted for every window standing"
        );
    }
}
