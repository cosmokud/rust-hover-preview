//! Where Windows is told to start this app, and whether it still will.
//!
//! **Two mechanisms, one setting.** `run_at_startup` is answered by the machine: off, nothing at a
//! logon starts this app; on, either a value under
//! `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` holds the command line a logon runs, or —
//! where the user has also asked for `Run as administrator` — a scheduled task with a logon trigger
//! and the highest run level does, because a `Run` value is never elevated and no UAC prompt is ever
//! raised for one (see `app::task`). This module is everything this app knows about both: writing
//! them, removing them, reading them back, and deciding whether what is there is what starts the copy
//! running now.
//!
//! **The machine is the truth, and the configuration is a copy of it.** Whether this app starts
//! is a fact about the machine, so a row of the tray and `run_at_startup` in `config.ini` are both
//! answered by asking the machine — an entry the user turned off in **Settings → Apps → Startup**,
//! or one another tool removed, reads as the app that will not start rather than as the value the
//! file remembers choosing.
//!
//! **Two things about a `Run` entry are the platform's rather than this app's, and both are answered
//! here.** Windows keeps a second value per entry under `Explorer\StartupApproved\Run`, holding
//! whether the entry is *enabled*; a value that is present but disabled is an entry that never
//! runs, and asking only whether the value is there answers yes for a machine where nothing
//! starts. And the value only runs unelevated: Windows never raises a UAC prompt for a `Run` entry,
//! so a command that requires elevation — an app marked *Run this program as an administrator* —
//! is a command that does not start at all. That last fact is what the task above is for.
//!
//! **Which account's entry it is** is the other half of the same problem: this app installs per
//! user and writes under `HKEY_CURRENT_USER`, which an elevated run answers with the hive of the
//! account that was elevated to rather than of the account sitting at the desktop. When those are
//! two different accounts — an administrator's credentials answering another user's prompt — the
//! entry is written where no logon will ever read it (see `app::elevation`), so it is written into
//! the desktop account's hive through `HKEY_USERS` instead.

use std::env;
use std::fs;
use std::path::Path;

use crate::app::elevation;
use crate::app::task::{self, Refused};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_OBJECT_0};
use windows::Win32::System::Registry::{
    RegCloseKey, RegCreateKeyExW, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW,
    HKEY, HKEY_CURRENT_USER, HKEY_USERS, KEY_QUERY_VALUE, KEY_READ, KEY_SET_VALUE, REG_BINARY,
    REG_CREATE_KEY_DISPOSITION, REG_OPTION_NON_VOLATILE, REG_SAM_FLAGS, REG_SZ,
};
use windows::Win32::System::Threading::WaitForSingleObject;
use windows::Win32::UI::Shell::{
    ShellExecuteExW, SEE_MASK_NOASYNC, SEE_MASK_NOCLOSEPROCESS, SHELLEXECUTEINFOW,
};
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

/// The key a logon reads the commands to start out of, relative to the account's hive.
const STARTUP_KEY: PCWSTR = w!(r"Software\Microsoft\Windows\CurrentVersion\Run");
/// The value this app writes in it. The name is a fact of the machine rather than of a run: it is
/// what the tray's own row, the logon and anyone reading `reg query` know this app by.
const APP_NAME: PCWSTR = w!("RustHoverPreview");
/// Where Windows keeps whether a `Run` value is enabled, the same way round: the entry's own name
/// is the value's name.
const APPROVED_KEY: PCWSTR =
    w!(r"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run");
/// The first byte of a value under `APPROVED_KEY` is the switch Windows and Task Manager flip: an
/// **even** byte is an entry a logon runs, an **odd** one is an entry somebody turned off. The shapes
/// seen on a machine are `02` (and `06`, which the Shell writes another way) for on, and `03` — or
/// `01` on older builds, followed by the FILETIME the switch was flipped at — for off, so the parity
/// is the rule that holds across all of them without this having to know which build wrote which.
/// What is written here is the shape this build's own Task Manager writes: `02` and then a FILETIME
/// of zeroes.
const APPROVED_ENABLED: [u8; 12] = [2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
/// Where Windows keeps the compatibility flags set on **Properties → Compatibility**: one value per
/// executable, named for its full path, holding the flags as one space-separated run of words.
const COMPAT_LAYERS_KEY: PCWSTR =
    w!(r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers");

/// What a write to the startup entry did, so that a row of the tray can say something other than
/// a checkmark that flips back by itself.
#[derive(Clone, PartialEq, Eq, Debug, Default)]
pub(crate) enum Outcome {
    /// The machine holds what was asked for.
    #[default]
    Settled,
    /// The registry refused the write, or the read back after it differs from what was asked for.
    /// Carries what the machine actually holds now.
    NotWritten { held: bool },
    /// The entry was written into another account's hive, because this process is not running as
    /// the account the desktop belongs to. Carries that account's SID.
    AnotherAccount(String),
    /// A logon will start this app, but it needs a scheduled task rather than a `Run` value, and
    /// writing one needs the elevated token this process does not have. Nothing was written.
    ///
    /// It is the ordinary answer of an unelevated run whose machine already holds an entry that
    /// asks for elevation — an exe marked *Run this program as an administrator* in its
    /// compatibility properties, or a `run_as_administrator` a file asked for — and what the caller
    /// does about it is hand the write to an elevated copy of this app (see `the_elevated_helper`).
    NeedsElevation,
    /// The row was asked for an elevated start and the machine refused the task for a reason of its
    /// own: a policy, a broken Task Scheduler, a service that is not running.
    TaskRefused,
    /// The mark Windows keeps on the file itself — *Run this program as an administrator* — did not
    /// follow the row, in the direction it was going. `on` is the direction: a mark that could not be
    /// put on leaves a logon elevated and every other launch not, and one that could not be taken off
    /// leaves a `Run` value beside it that Windows starts nothing from.
    MarkRefused { on: bool },
    /// The ordinary `Run` value was written, but the scheduled task that starts this app **elevated**
    /// at a logon could not be taken away: a logon would start it twice, once unelevated from the
    /// value and once elevated from the task. It is what turning the `Run as administrator` row off
    /// answers when the token at hand may write a value but not delete a task.
    TaskStayed,
}

impl Outcome {
    /// Whether what was asked for is what the machine holds.
    pub(crate) fn settled(&self) -> bool {
        matches!(self, Outcome::Settled | Outcome::AnotherAccount(_))
    }
}

/// Which of the two mechanisms a machine's startup entry is, as the tray's rows are drawn from it.
#[derive(Clone, Copy, PartialEq, Eq, Debug, Default)]
pub(crate) enum Started {
    /// Nothing at a logon starts this app.
    #[default]
    Nothing,
    /// A `Run` value starts this copy, unelevated, which is the ordinary way this app has always
    /// started with Windows.
    Entry,
    /// A scheduled task starts this copy with the highest run level — the row that asks Windows for
    /// administrator rights at a logon.
    Task,
    /// One of the two is there and does not start anything this copy is party to: a `Run` value
    /// another tool switched off, a task Task Scheduler disabled, an entry naming a copy that has
    /// since moved, or a task registered for another account.
    NotRunning,
}

impl Started {
    /// Whether the machine starts this copy at a logon.
    pub(crate) fn runs(&self) -> bool {
        matches!(self, Started::Entry | Started::Task)
    }
}

/// What the machine holds about the two mechanisms, asked once: it is four registry reads and two
/// reads of a task, and every row of the `System` submenu is drawn from one answer.
#[derive(Clone, Copy, PartialEq, Eq, Debug, Default)]
pub(crate) struct Machine {
    /// The `Run` value's presence, and whether its `StartupApproved` flag lets it run.
    pub entry_present: bool,
    pub entry_enabled: bool,
    /// The entry names this copy of the app rather than another one.
    pub entry_is_this_copy: bool,
    /// The scheduled task: there, enabled, at the highest run level, and naming this copy.
    pub task_present: bool,
    pub task_enabled: bool,
    pub task_elevated: bool,
    pub task_is_this_copy: bool,
}

impl Machine {
    /// Which mechanism, if either, starts this copy at a logon.
    pub(crate) fn started(&self) -> Started {
        if self.entry_present && self.entry_enabled && self.entry_is_this_copy {
            Started::Entry
        } else if self.task_present
            && self.task_enabled
            && self.task_elevated
            && self.task_is_this_copy
        {
            Started::Task
        } else if self.entry_present || self.task_present {
            // One of the two is on the machine and starts nothing of this copy's: a value
            // switched off in Task Manager, a task somebody disabled, or one naming another
            // copy — which is what the row says rather than what it hides.
            Started::NotRunning
        } else {
            Started::Nothing
        }
    }

    /// Whether the machine would start this app at a logon, elevated: the whole of what the
    /// `Run as administrator` row is drawn from.
    pub(crate) fn elevates(&self) -> bool {
        self.started() == Started::Task
    }

    /// Whether the exe is marked *Run this program as an administrator* in its own properties.
    ///
    /// The mark is one Windows honours for a program a person starts by hand — a launch from the
    /// Start menu, a shortcut, an association raises the prompt for it — and the one a `Run` value
    /// can never satisfy. Where it is set, the elevated start is the one the user asked for, whether
    /// they asked for it here or in the file's properties (see `entry_would_ask_for_elevation`).
    pub(crate) fn compat_elevation(&self) -> bool {
        entry_would_ask_for_elevation()
    }
}

/// The command line a logon runs, spelled the way the Shell spells one: the path quoted, so that a
/// folder with a space in it is one path rather than a folder and an argument.
///
/// A bare path is what this app used to write, and it still starts a path with no spaces in it —
/// but a path that has one is half a path to `CreateProcess`. Both spellings are read back the
/// same way (see `executable_of`), so a machine carrying either is understood.
fn command_line(path: &Path) -> String {
    format!("\"{}\"", path.display())
}

/// Write the command line of the copy that is running, and enable the entry.
///
/// One write of the command, one of the enabled flag, and one read back: a write the machine
/// refused is an answer rather than nothing, because the caller is a row of the tray that would
/// otherwise say a switch is on where nothing starts.
pub(crate) fn enable_startup() -> Outcome {
    let Ok(exe_path) = env::current_exe() else {
        return Outcome::NotWritten { held: false };
    };

    if elevated_start_wanted() {
        return enable_the_task(&exe_path);
    }

    write_the_entry_and_take_the_task_away(&exe_path)
}

/// Write the ordinary `Run` value and take the elevated task away — the machine a logon starts
/// unelevated from — **without asking the file whether it wants an elevated start**.
///
/// **It is the answer the `Run as administrator` row gives when it is turned off, and the one call
/// that can be trusted to give it.** [`enable_startup`] asks [`elevated_start_wanted`], which reads
/// the mark on the file as well as the file's own setting: right for a logon, wrong for this row,
/// whose whole subject is that mark. A write that asked would be answered by a mark that had not
/// gone yet — or by one the machine would not let go of — and would register the task the row is
/// turning off, which is the state this call exists to make impossible.
///
/// A task the machine will not take away is answered rather than ignored: a logon that starts this
/// app elevated beside the value that starts it unelevated is not what the row asked for.
pub(crate) fn settle_as_unelevated() -> Outcome {
    let Ok(exe_path) = env::current_exe() else {
        return Outcome::NotWritten { held: false };
    };

    let written = write_entry(&exe_path) && enable_approved();
    let held = is_entry_enabled_here();

    if !written || !held {
        return Outcome::NotWritten { held };
    }

    // A task this token may not delete is a task that still starts the app at a logon — elevated,
    // and beside the value just written — which is the one state the row cannot be left in.
    if task::unregister().is_err() && task::facts().present {
        return Outcome::TaskStayed;
    }

    if elevation::runs_under_another_account() {
        // The entry was written into another account's hive rather than this process's own, which
        // is still the right entry — the one the logon of the desktop's account reads — and is
        // worth saying, because nothing in this account's own `Run` key will show it.
        elevation::desktop_user_sid().map_or(Outcome::Settled, Outcome::AnotherAccount)
    } else {
        Outcome::Settled
    }
}

/// The entry half of [`enable_startup`], with the task's removal as something that may be refused:
/// the ordinary answer of a process that may write a `Run` value but not delete a task.
fn write_the_entry_and_take_the_task_away(exe_path: &Path) -> Outcome {
    let written = write_entry(exe_path) && enable_approved();
    let held = is_entry_enabled_here();

    if !written || !held {
        return Outcome::NotWritten { held };
    }

    // The task is the other half: a machine that carries one from an earlier answer would start
    // this app twice at a logon — once unelevated from the registry and once elevated from the
    // task — so the task goes with the answer that does not want it. A removal the machine refuses
    // is not a failure of this write: what was asked for (the entry) is what is there.
    let _ = task::unregister();

    if elevation::runs_under_another_account() {
        // The entry was written into another account's hive rather than this process's own, which
        // is still the right entry — the one the logon of the desktop's account reads — and is
        // worth saying, because nothing in this account's own `Run` key will show it.
        elevation::desktop_user_sid().map_or(Outcome::Settled, Outcome::AnotherAccount)
    } else {
        Outcome::Settled
    }
}

/// Whether the logon that starts this app should start it **elevated**.
///
/// Two things ask for that, and the answer is either of them: the `run_as_administrator` a user set
/// on the tray's row, and the *Run this program as an administrator* mark a user set on the file in
/// Windows' own compatibility properties. The second is deliberately read as the first rather than
/// warned about: the mark is the same request, Windows honours it for every launch that is not a
/// `Run` value, and a row that said "this cannot work" over a machine where it now can would be a
/// row arguing with the mechanism this app put on it.
fn elevated_start_wanted() -> bool {
    crate::CONFIG
        .lock()
        .map(|config| config.run_as_administrator)
        .unwrap_or(false)
        || entry_would_ask_for_elevation()
}

/// Write the scheduled task that starts this copy at a logon with the highest run level, remove the
/// `Run` value the same logon would otherwise also run, and mark the file so that every other launch
/// of it asks for administrator rights too.
fn enable_the_task(exe_path: &Path) -> Outcome {
    let user = elevation::desktop_user_name();
    let Some(user) = user else {
        // A machine with no Shell to ask is one where the account cannot be named, and a task
        // registered for the wrong account is worse than no task: it is a logon entry nobody's
        // logon reads.
        return Outcome::NotWritten { held: false };
    };

    match task::register(exe_path, &user) {
        Ok(()) => {
            // The two are alternatives rather than both, so the value that would start an
            // unelevated copy beside the elevated one goes. A registry this token may not write is
            // not a reason to call the write failed: the task is what starts the app.
            delete_entry();
            delete_approved();

            let held = task::facts().starts_this_copy(exe_path);

            if !held {
                // Nothing else is written over a task the machine does not read back as this
                // copy's: a mark on the file under a startup that is not the task is a machine
                // where neither mechanism starts anything.
                return Outcome::TaskRefused;
            }

            // **And the mark, which is the half a logon never goes through.** The task starts this
            // app elevated once, at a logon; a double-click, a shortcut, an association or a
            // `Start → Run` line begins the executable itself, and the file's own compatibility
            // flags are the only thing Windows reads for one of those (see
            // `write_compat_elevation`).
            match write_compat_elevation(exe_path) {
                true => Outcome::Settled,
                false => Outcome::MarkRefused { on: true },
            }
        }
        Err(Refused::NeedsElevation) => Outcome::NeedsElevation,
        Err(Refused::NoWrite) => Outcome::TaskRefused,
    }
}

/// Remove both mechanisms: the `Run` value and its flag, and the scheduled task.
///
/// **The mark on the file is not one of them.** It belongs to `run_as_administrator`, which is the
/// row under the one this answers: an `Run at Startup` turned off is a startup turned off, not an
/// elevation turned off, and a machine left asking for elevation on every launch is what that row
/// still says. What removes it is that row, in both directions, and the reconciler behind it (see
/// `honour_the_configuration`).
pub(crate) fn disable_startup() -> Outcome {
    delete_entry();
    delete_approved();

    // A task this process may not delete is a task that still starts the app — which is the state
    // the caller is told about rather than one it is left to discover at the next logon.
    let task_gone = task::unregister().is_ok() || !task::facts().present;

    let held = is_entry_enabled_here() || !task_gone;

    match held {
        true => Outcome::NotWritten { held: true },
        false => Outcome::Settled,
    }
}

/// Whether this app is registered to start when the user logs in, **and enabled** — by either
/// mechanism, and naming this copy.
///
/// Asked of the machine rather than of the configuration, because the machine is what actually
/// starts it: the value's presence and its flags are the entry, and a task's presence and its flags
/// are the task. An entry the user turned off in Task Manager, a task Task Scheduler disabled, an
/// entry another tool removed, or one naming a copy that has since moved is answered as the app
/// that will not start.
pub(crate) fn is_startup_enabled() -> bool {
    machine().started().runs()
}

/// What the machine holds about both mechanisms, as the tray's rows are drawn from it.
pub(crate) fn machine() -> Machine {
    let exe = env::current_exe().ok();
    let facts = task::facts();

    Machine {
        entry_present: is_entry_present(),
        entry_enabled: is_entry_enabled(),
        entry_is_this_copy: entry_names_this_copy(),
        task_present: facts.present,
        task_enabled: facts.enabled,
        task_elevated: facts.elevated,
        task_is_this_copy: exe
            .as_deref()
            .is_some_and(|exe| facts.starts_this_copy(exe)),
    }
}

/// Whether the `Run` value is there and Windows is allowed to run it, which is the half of
/// `is_startup_enabled` that `enable_startup`'s read-back is made of.
fn is_entry_enabled_here() -> bool {
    is_entry_present() && is_entry_enabled()
}

/// Whether the value is there at all, whether or not Windows is allowed to run it.
pub(crate) fn is_entry_present() -> bool {
    let Some(key) = open_startup(KEY_QUERY_VALUE) else {
        return false;
    };

    // Safety: a value of the key this call opened, asked for its presence only.
    let present = unsafe { RegQueryValueExW(key, APP_NAME, None, None, None, None) }.is_ok();
    // Safety: the handle `open_startup` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(key);
    }

    present
}

/// Whether Windows is allowed to run the entry — which is not the same question as whether it is
/// there, and is the one a run has to ask to know that a logon will start anything.
///
/// **And it is the second of two facts a `Run` entry has to pass.** The first is this one; the other
/// is whether the executable is marked *Run this program as an administrator* (see
/// `entry_would_ask_for_elevation`), which Windows answers by starting nothing at all, because a
/// `Run` value is never elevated and no UAC prompt is raised for one.
pub(crate) fn is_entry_enabled() -> bool {
    let Some(key) = open_approved(KEY_QUERY_VALUE) else {
        // No flag at all is an entry nobody disabled: the flag is written when Windows or the user
        // touches the switch, and its absence is the state a fresh entry is in.
        return true;
    };

    let mut buffer = [0u8; 16];
    let mut size = buffer.len() as u32;
    // Safety: a buffer of `size` bytes is offered, and the size is what is read back into.
    let status = unsafe {
        RegQueryValueExW(
            key,
            APP_NAME,
            None,
            None,
            Some(buffer.as_mut_ptr()),
            Some(&mut size),
        )
    };
    // Safety: the handle `open_approved` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(key);
    }

    if status.is_ok() && size >= 1 {
        // An even switch is an entry a logon runs, an odd one is an entry somebody turned off (see
        // `APPROVED_ENABLED`).
        buffer[0].is_multiple_of(2)
    } else {
        // A flag that could not be read is not a disable: the entry's presence is the answer, and a
        // machine whose `StartupApproved` key this process may not read is one where Windows has not
        // written an off switch this app can see.
        true
    }
}

/// Whether the copy that is running is marked *Run this program as an administrator*.
///
/// **That mark is what a `Run` entry cannot survive.** Windows starts a `Run` value unelevated and
/// never raises a UAC prompt for one, so an executable carrying the `RUNASADMIN` compatibility flag —
/// set on **Properties → Compatibility**, or by anyone who once launched it elevated and ticked the
/// box — is an entry that does nothing at a logon: the Shell will not elevate it, and the app is not
/// started at all. Nothing this app writes can undo that mark, which is the user's own setting on the
/// file, so what is left is to say so where the startup row is read.
///
/// The flag lives in a per-user list keyed by the full path of the executable and valued as one
/// space-separated run of words, so what is looked for is `RUNASADMIN` in the entry of this copy's
/// own path.
pub(crate) fn entry_would_ask_for_elevation() -> bool {
    env::current_exe().is_ok_and(|exe| file_asks_for_elevation(&exe))
}

/// Whether one executable's own compatibility layers carry the mark.
fn file_asks_for_elevation(path: &Path) -> bool {
    compat_layers_value(path).is_some_and(|layers| {
        layers
            .split_whitespace()
            .any(|word| word.eq_ignore_ascii_case("RUNASADMIN"))
    })
}

/// The compatibility layers this machine holds for one executable, as the words they are written
/// as, or nothing where there is no entry for it.
fn compat_layers_value(path: &Path) -> Option<String> {
    // **One door for the read and the write.** The key is opened through `open_layers`, which is the
    // hive a write goes to: the desktop account's where this process was elevated into another one,
    // and this account's own otherwise. A read from `HKEY_CURRENT_USER` while the write went to
    // `HKEY_USERS\<the desktop's SID>` would be two answers about two different keys, and the mark
    // this app put on would be a mark it could never see to take off.
    let hkey = open_layers(KEY_QUERY_VALUE)?;

    let layers = read_layers(hkey, path);

    // Safety: the handle this call opened and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    layers
}

/// The layer value a key of `AppCompatFlags\Layers` holds for one executable.
///
/// The value's name is the path, spelled as the compatibility tab wrote it: the exact key is asked
/// for, and a lookup that misses is not retried in another case — one that differs only in case
/// would be a second entry for one file, which the tab does not make.
fn read_layers(hkey: HKEY, path: &Path) -> Option<String> {
    let name = layers_value_name(path);

    let mut size: u32 = 0;
    // Safety: a size is written into a local; the value's bytes are not offered yet.
    let sized = unsafe {
        RegQueryValueExW(
            hkey,
            PCWSTR(name.as_ptr()),
            None,
            None,
            None,
            Some(&mut size),
        )
    }
    .is_ok()
        && size >= 2;

    if !sized {
        return None;
    }

    let mut buffer = vec![0u8; size as usize];
    // Safety: a buffer of `size` bytes is offered, and the size is what is read back into.
    let read = unsafe {
        RegQueryValueExW(
            hkey,
            PCWSTR(name.as_ptr()),
            None,
            None,
            Some(buffer.as_mut_ptr()),
            Some(&mut size),
        )
    }
    .is_ok();

    read.then(|| decode_string(&buffer[..size.min(buffer.len() as u32) as usize]))
}

/// A layer value with the *Run this program as an administrator* word taken out of it.
///
/// **The word is one of a list, and only it goes.** A compatibility mode (`~ WIN8RTM`), a DPI fix or
/// `HIGHDPIAWARE` beside it is the user's own setting on the file and is left exactly as written.
/// `None` is a value that named nothing else and is therefore a value with no layers left in it —
/// which is removed rather than written back as the `~` separator on its own, a run of characters
/// that names no layer and is not a request Windows understands.
fn layers_without_the_mark(layers: &str) -> Option<String> {
    let words: Vec<&str> = layers.split_whitespace().collect();

    // `~` is the list's own separator rather than a layer, and it opens a value that has one. A
    // list left with layers in it keeps it — `~ WIN8RTM` is what Windows' own properties tab writes
    // — and a list with none left is no value at all rather than a separator naming nothing.
    let separator = words.first().is_some_and(|word| *word == "~");
    let kept: Vec<&str> = words
        .into_iter()
        .filter(|word| *word != "~")
        .filter(|word| !word.eq_ignore_ascii_case("RUNASADMIN"))
        .collect();

    if kept.is_empty() {
        return None;
    }

    Some(match separator {
        true => format!("~ {}", kept.join(" ")),
        false => kept.join(" "),
    })
}

/// A layer value with the *Run this program as an administrator* word in it.
///
/// **The word is added to the list rather than written over it.** A compatibility mode (`~ WIN8RTM`)
/// or a DPI fix beside it is the user's own setting on the file and is left where it is, and the `~`
/// that opens a list of layers is written once — a value made from nothing is the separator and the
/// mark, which is what Windows' own properties tab writes for one flag.
fn layers_with_the_mark(layers: Option<&str>) -> String {
    let mut words: Vec<&str> = layers
        .unwrap_or_default()
        .split_whitespace()
        .filter(|word| *word != "~")
        .filter(|word| !word.eq_ignore_ascii_case("RUNASADMIN"))
        .collect();

    words.push("RUNASADMIN");

    format!("~ {}", words.join(" "))
}

/// The name a compatibility layer value is filed under: this copy's path, UTF-16 and terminated.
fn layers_value_name(path: &Path) -> Vec<u16> {
    path.as_os_str()
        .to_string_lossy()
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect()
}

/// Take the *Run this program as an administrator* mark off this copy, which is what unchecking the
/// `Run as administrator` row does.
///
/// **The mark has to go with the row, or the row cannot turn it off.** Windows never elevates a
/// `Run` value and silently skips one whose exe asks for elevation, so a machine left carrying this
/// flag beside an ordinary startup entry is a machine where nothing starts at all — the row would
/// untick, the value would be written, and the next logon would run nothing. `RUNASADMIN` is one
/// word among the layers a program may carry (compatibility modes, DPI fixes, `HIGHDPIAWARE`), so
/// only that word is taken: the value is rewritten without it, and a value that is nothing but it
/// leaves no empty string behind, because an empty layer is a request Windows does not understand.
///
/// The key is the **desktop account's** while this process is an elevated copy of another account,
/// for the reason the `Run` entry is (see `app::elevation`): the compatibility flags are read for the
/// user who starts the program, and that user is the one at the keyboard.
///
/// Answers whether the mark is gone afterwards, which is a machine with none at all as much as one
/// whose layers no longer name it.
fn remove_compat_elevation(path: &Path) -> bool {
    // **What is asked is the machine afterwards, not the write.** A key this process may not open,
    // or a value it may not read, used to be answered as "no mark at all" — which is the one answer
    // that leaves the mark on the file while the row says it is off, and a machine that starts
    // nothing at a logon because Windows starts no `Run` value whose file asks for elevation. So the
    // mark is read first, through the same hive the write goes to, and a mark that is there and
    // cannot be reached is a mark that did not go.
    if !file_asks_for_elevation(path) {
        return true;
    }

    let Some(hkey) = open_layers(KEY_QUERY_VALUE | KEY_SET_VALUE) else {
        return false;
    };

    let Some(layers) = read_layers(hkey, path) else {
        // Safety: the handle this call opened and nothing else holds.
        unsafe {
            let _ = RegCloseKey(hkey);
        }
        return false;
    };

    let kept = layers_without_the_mark(&layers);

    let name = layers_value_name(path);

    let status = match &kept {
        // Safety: this app's own value for this copy's path in a key this call opened.
        None => unsafe { RegDeleteValueW(hkey, PCWSTR(name.as_ptr())) },
        Some(kept) => {
            let written: Vec<u16> = kept.encode_utf16().chain(std::iter::once(0)).collect();
            // Safety: the value's bytes are a UTF-16 string with its terminator.
            unsafe {
                RegSetValueExW(
                    hkey,
                    PCWSTR(name.as_ptr()),
                    0,
                    REG_SZ,
                    Some(written.align_to::<u8>().1),
                )
            }
        }
    };

    // Safety: the handle this call opened and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    status.is_ok() && !file_asks_for_elevation(path)
}

/// Put the *Run this program as an administrator* mark on this copy, which is what turning the
/// `Run as administrator` row on does.
///
/// **The task is what a logon runs, and this mark is what every other launch goes through.** A
/// double-click, a shortcut, an association and a `Start → Run` line all begin the executable itself
/// and never touch the task, and Windows raises a prompt for one of those only where the file carries
/// this flag. So the row writes both halves of itself, and a machine carrying one without the other
/// is a row that means less than it says.
///
/// **It is written after the task, never before.** A `Run` value under this mark starts nothing at
/// all — Windows never elevates one — so a mark written over a startup entry that the task was meant
/// to replace would be a machine where a logon starts nothing; the entry is gone by the time this is
/// called (see `enable_the_task`).
///
/// The key is the **desktop account's** while this process is an elevated copy of another account,
/// for the reason every other per-user write here is (see `app::elevation`).
///
/// Answers whether the mark is there afterwards.
fn write_compat_elevation(path: &Path) -> bool {
    let Some(hkey) = create_layers() else {
        return false;
    };

    let marked = layers_with_the_mark(read_layers(hkey, path).as_deref());
    let name = layers_value_name(path);
    let written: Vec<u16> = marked.encode_utf16().chain(std::iter::once(0)).collect();

    // Safety: this app's own value for this copy's path in a key this call opened, and the value's
    // bytes are a UTF-16 string with its terminator.
    let status = unsafe {
        RegSetValueExW(
            hkey,
            PCWSTR(name.as_ptr()),
            0,
            REG_SZ,
            Some(written.align_to::<u8>().1),
        )
    };

    // Safety: the handle this call opened and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    status.is_ok() && file_asks_for_elevation(path)
}

/// Write this copy's command line into the entry, creating the key where it is not there.
///
/// The key is created rather than only opened: a `Run` key that another tool removed is exactly the
/// case this exists for, and `RegOpenKeyExW` fails on it while a write to a value in a key that is
/// not there looks — from outside — like a switch that will not stay on.
fn write_entry(path: &Path) -> bool {
    // The value is a counted string, and its terminator is part of what is written: a path whose
    // bytes stop one short of it is read back carrying whatever was in the registry next to it.
    let wide: Vec<u16> = command_line(path)
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    let Some(hkey) = create_startup() else {
        return false;
    };

    // Safety: the value's bytes are a UTF-16 string with its terminator, and the key is the one
    // this call made and closes here.
    let status =
        unsafe { RegSetValueExW(hkey, APP_NAME, 0, REG_SZ, Some(wide.align_to::<u8>().1)) };
    // Safety: the handle `create_startup` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    status.is_ok()
}

/// Write the flag that says the entry runs, so that an entry the user turned off in Task Manager is
/// put back on when this app is asked to start with Windows.
fn enable_approved() -> bool {
    let Some(hkey) = create_approved() else {
        return false;
    };

    // Safety: the value's bytes are the twelve-byte shape the Shell's own switch writes, and the
    // key is closed here.
    let status = unsafe { RegSetValueExW(hkey, APP_NAME, 0, REG_BINARY, Some(&APPROVED_ENABLED)) };
    // Safety: the handle `create_approved` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    status.is_ok()
}

fn delete_entry() {
    let Some(hkey) = open_startup(KEY_SET_VALUE | KEY_QUERY_VALUE) else {
        return;
    };

    // Safety: the value is this app's own name in a key this call opened, and the key is closed
    // here.
    unsafe {
        let _ = RegDeleteValueW(hkey, APP_NAME);
        let _ = RegCloseKey(hkey);
    }
}

fn delete_approved() {
    let Some(hkey) = open_approved(KEY_SET_VALUE | KEY_QUERY_VALUE) else {
        return;
    };

    // Safety: this app's own value in a key this call opened, closed here.
    unsafe {
        let _ = RegDeleteValueW(hkey, APP_NAME);
        let _ = RegCloseKey(hkey);
    }
}

/// The `Run` key of the account whose logon will run the entry.
///
/// **`HKEY_CURRENT_USER` unless this process was elevated into another account's token.** An
/// administrator answering their own UAC prompt is the same account as the desktop's, and there is
/// nothing to redirect; one answering somebody else's prompt runs as an account the desktop does not
/// belong to, and the entry has to name the desktop's hive or nothing at the logon reads it (see
/// `app::elevation`).
///
/// A SID this process may not write under — an account whose hive is not loaded, or one whose key
/// this token may not create — falls back to `HKEY_CURRENT_USER`, which is where an ordinary run
/// writes anyway, and what that write does is read back before it is believed.
fn redirected_hive() -> Option<String> {
    if !elevation::runs_under_another_account() {
        return None;
    }

    let sid = elevation::desktop_user_sid()?;

    // The account's own key under `HKEY_USERS` is what says whether its hive can be written at all:
    // an account nobody has logged on to since the machine started has no hive there, and a key that
    // is not there is not one to write a subkey under. The handle is closed at once: what is asked
    // here is whether the key can be opened, and the write below opens it again for its own purpose.
    let reachable = open_user_key(&sid, KEY_READ)?;
    // Safety: the handle `open_user_key` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(reachable);
    }

    Some(sid)
}

/// The subkey of the hive the entry lives under: the same path, under `HKEY_CURRENT_USER` or under
/// the desktop account's own key in `HKEY_USERS`.
fn startup_subkey() -> Option<String> {
    let sid = redirected_hive()?;
    Some(format!(
        "{sid}\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    ))
}

/// The same subkey for the flag Windows keeps beside the entry.
fn approved_subkey() -> Option<String> {
    let sid = redirected_hive()?;
    Some(format!(
        "{sid}\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StartupApproved\\Run"
    ))
}

/// The subkey the compatibility layers live under: the same path, under this account's own hive or
/// under the desktop account's, which is the one a launch of the exe is judged by.
fn compat_layers_subkey() -> Option<String> {
    let sid = redirected_hive()?;
    Some(format!(
        "{sid}\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers"
    ))
}

/// The key the compatibility layers live under, opened for `access`.
fn open_layers(access: REG_SAM_FLAGS) -> Option<HKEY> {
    match compat_layers_subkey() {
        Some(subkey) => open_user_key(&subkey, access),
        None => open_own_key(COMPAT_LAYERS_KEY, access),
    }
}

/// The same key, created where it is not there: a machine that has never had a compatibility flag set
/// has no key to write one into, and a mark that could not be written for want of a key is a row that
/// does nothing on the machine it was turned on for.
fn create_layers() -> Option<HKEY> {
    let access = KEY_SET_VALUE | KEY_QUERY_VALUE;

    match compat_layers_subkey() {
        Some(subkey) => create_user_key(&subkey, access),
        None => create_own_key(COMPAT_LAYERS_KEY, access),
    }
}

/// The key the entry lives under, opened for `access`, or nothing where it cannot be.
fn open_startup(access: REG_SAM_FLAGS) -> Option<HKEY> {
    match startup_subkey() {
        Some(subkey) => open_user_key(&subkey, access),
        None => open_own_key(STARTUP_KEY, access),
    }
}

/// The same key, created where it is not there.
fn create_startup() -> Option<HKEY> {
    let access = KEY_SET_VALUE | KEY_QUERY_VALUE;

    match startup_subkey() {
        Some(subkey) => create_user_key(&subkey, access),
        None => create_own_key(STARTUP_KEY, access),
    }
}

fn open_approved(access: REG_SAM_FLAGS) -> Option<HKEY> {
    match approved_subkey() {
        Some(subkey) => open_user_key(&subkey, access),
        None => open_own_key(APPROVED_KEY, access),
    }
}

fn create_approved() -> Option<HKEY> {
    let access = KEY_SET_VALUE | KEY_QUERY_VALUE;

    match approved_subkey() {
        Some(subkey) => create_user_key(&subkey, access),
        None => create_own_key(APPROVED_KEY, access),
    }
}

/// A key of this account's own hive, opened for `access`.
fn open_own_key(subkey: PCWSTR, access: REG_SAM_FLAGS) -> Option<HKEY> {
    let mut hkey = HKEY::default();
    // Safety: a handle is written into a local, and the caller closes it.
    let opened = unsafe { RegOpenKeyExW(HKEY_CURRENT_USER, subkey, 0, access, &mut hkey) };
    opened.ok().ok()?;
    Some(hkey)
}

/// A key of this account's own hive, created where it is not there.
///
/// Creating is the whole of what a `Run` key that another tool removed needs: `RegOpenKeyExW` fails
/// on a key that is not there, and a write to a value in it would then look — from outside — like a
/// switch that will not stay on.
fn create_own_key(subkey: PCWSTR, access: REG_SAM_FLAGS) -> Option<HKEY> {
    let mut hkey = HKEY::default();
    let mut disposition = REG_CREATE_KEY_DISPOSITION::default();
    // Safety: a handle is written into a local, and the caller closes it; the disposition is an
    // optional out-parameter the API is given so that it is not asked to write nowhere.
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            subkey,
            0,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            access,
            None,
            &mut hkey,
            Some(&mut disposition),
        )
    };
    status.ok().ok()?;
    Some(hkey)
}

/// A subkey of `HKEY_USERS`, opened for `access`.
fn open_user_key(subkey: &str, access: REG_SAM_FLAGS) -> Option<HKEY> {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();

    let mut hkey = HKEY::default();
    // Safety: a handle is written into a local, and the caller closes it.
    let opened = unsafe { RegOpenKeyExW(HKEY_USERS, PCWSTR(wide.as_ptr()), 0, access, &mut hkey) };
    opened.ok().ok()?;
    Some(hkey)
}

/// A subkey of `HKEY_USERS`, created where it is not there.
fn create_user_key(subkey: &str, access: REG_SAM_FLAGS) -> Option<HKEY> {
    let wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();

    let mut hkey = HKEY::default();
    let mut disposition = REG_CREATE_KEY_DISPOSITION::default();
    // Safety: as in `create_own_key`.
    let status = unsafe {
        RegCreateKeyExW(
            HKEY_USERS,
            PCWSTR(wide.as_ptr()),
            0,
            PCWSTR::null(),
            REG_OPTION_NON_VOLATILE,
            access,
            None,
            &mut hkey,
            Some(&mut disposition),
        )
    };
    status.ok().ok()?;
    Some(hkey)
}

/// What the entry holds, as the text it is read back as.
///
/// An entry this app writes is a path, and the shape of it can be told from the type alone —
/// but the entry is the registry's, and a value is whatever is in it, so what is read is
/// the bytes rather than the type's promise of them: a `REG_SZ` carrying a number is still
/// text to be judged, not a length to be trusted.
fn registered_command() -> Option<String> {
    let hkey = open_startup(KEY_READ)?;

    // The size is asked for before the bytes are, the way the API counts a string: one
    // call for how many bytes there are, and one for the bytes themselves.
    let mut size: u32 = 0;
    let sized = unsafe { RegQueryValueExW(hkey, APP_NAME, None, None, None, Some(&mut size)) }.is_ok()
        // A string is at least its own terminator; anything shorter is a value that
        // names no executable rather than one that names it emptily.
        && size >= 2;

    let command = if sized {
        let mut buffer = vec![0u8; size as usize];
        let read = unsafe {
            RegQueryValueExW(
                hkey,
                APP_NAME,
                None,
                None,
                Some(buffer.as_mut_ptr()),
                Some(&mut size),
            )
        }
        .is_ok();

        read.then(|| decode_string(&buffer[..size.min(buffer.len() as u32) as usize]))
    } else {
        None
    };

    // Safety: the handle `open_startup` made and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    command
}

/// A counted registry string as text: UTF-16, ending at its terminator.
///
/// The terminator is where the string ends whatever the buffer behind it holds, and an odd
/// trailing byte is a byte this reads nothing out of rather than a panic.
fn decode_string(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .take_while(|unit| *unit != 0)
        .collect();

    String::from_utf16_lossy(&units)
}

/// The executable a startup entry names, out of the command line the value holds.
///
/// Windows runs the value as a command line, so what is in the registry need not be a bare
/// path: a path holding a space is quoted, or half of it would run, and arguments can
/// follow the executable. Only the executable is the app — an argument is not a path to
/// another copy of it — so the name is taken from in front of them.
fn executable_of(command: &str) -> &str {
    let trimmed = command.trim();

    match trimmed.strip_prefix('"') {
        Some(quoted) => quoted.split('"').next().unwrap_or(quoted),
        None => trimmed.split_whitespace().next().unwrap_or(trimmed),
    }
}

/// Whether two spellings are the path of one file, as the machine reads a path.
///
/// The spelling itself settles it where it can: a Windows path carries no case, and the
/// verbatim prefix names the same file as without it. Where the two read differently the
/// files themselves are asked, which is what settles a short name for a folder or a
/// junction — a spelling this app never writes, and would otherwise replace over one file
/// with itself. A path that does not resolve is asked nothing: an entry naming a copy that
/// has been deleted is not this one, and is the case this exists for.
fn same_path(registered: &str, running: &Path) -> bool {
    let left = crate::paths::plain_path(Path::new(registered.trim()));
    let right = crate::paths::plain_path(running);

    if left.eq_ignore_ascii_case(&right) {
        return true;
    }

    match (fs::canonicalize(&left), fs::canonicalize(&right)) {
        (Ok(left), Ok(right)) => {
            crate::paths::plain_path(&left).eq_ignore_ascii_case(&crate::paths::plain_path(&right))
        }
        _ => false,
    }
}

/// Whether the entry, as it stands, starts the copy of the app that is running.
fn entry_names_this_copy() -> bool {
    let (Some(command), Ok(exe_path)) = (registered_command(), env::current_exe()) else {
        return false;
    };

    same_path(executable_of(&command), &exe_path)
}

/// Put the entry back on the copy of the app that is running, and say whether it was
/// rewritten.
///
/// What the entry says and what this process are are two things that can come apart: a
/// portable copy run once from a folder of its own, an older version still installed where
/// it put itself, an app moved since — and each writes its own path in, which is what the
/// next logon starts. The entry then names a version of this app that the tray the user is
/// looking at does not belong to, or nothing at all.
///
/// Only a path is touched, and only where the entry is already there: whether this app
/// starts at all is the user's choice, and a choice to have it off is not one this
/// overrides. Where the entry names this copy it is left exactly as it is, an argument of
/// the user's own included, and where it names another it is written in this app's own
/// spelling, since that is what starts it.
pub fn repair_startup_entry() -> bool {
    // An entry that is gone is not written back here — whether this app starts is the user's choice,
    // and a file that asks for it is answered by `honour_the_configuration` below. A task that is
    // gone is the same absence, and it is `honour_the_configuration` that answers it too.
    let held = machine();

    if !matches!(held.started(), Started::NotRunning) {
        return false;
    }

    // A copy that is named and does not start is put right in place, which is the one thing this
    // repairs: a task somebody disabled in Task Scheduler is turned back on (see
    // `honour_the_configuration` for the register that does it), and a value naming another copy
    // is written in this app's own spelling, since that is what starts it.
    match (held.task_present, held.entry_present) {
        (true, _) => false,
        (false, true) => enable_startup().settled(),
        (false, false) => false,
    }
}

/// Note that this app has just been started, and make the machine's entry agree with the
/// configuration that asked for it.
///
/// **The file's `run_at_startup` is an instruction here, not only a record.** A file that says yes
/// is a user who asked for this app to start with Windows, and every way that can come apart while
/// it says so is put right at the next start: the entry a cleaner removed, the value another tool
/// switched off in Task Manager (the flag Windows keeps beside it is read as well, see
/// `is_entry_enabled`), and the entry naming a copy that has since moved or been replaced. `true`
/// is answered where the entry was written, which is what the trace says happened.
///
/// The other direction is not taken. A file that says no is the user's answer — the tray's own row
/// is what wrote it, and it removed the entry as it did — so nothing is written at all, and the file
/// is brought into step with what the machine holds either way: a `run_at_startup` that says yes
/// beside an entry that is not there is a value that lies about the machine, and it is the thing
/// that made this look like the app forgetting the choice it was given.
///
/// **Asked at a start and at every opening of the tray menu**: it is three registry reads in the
/// ordinary case, and what it buys is that the row a user is about to read and the entry a logon
/// will run are the same fact.
pub fn honour_the_configuration() -> bool {
    let (wanted, wants_elevation) = crate::CONFIG
        .lock()
        .map(|config| (config.run_at_startup, config.run_as_administrator))
        .unwrap_or((false, false));

    let before = machine();

    // **The mark on the file follows the file's answer, in both directions.** It is the half of the
    // elevated row a logon never goes through — every launch of the exe itself reads it and nothing
    // else — so a file asking for the elevated start is a machine where it has to be there, and a
    // file that stopped asking is a machine that must not go on asking: a `Run` value under that
    // mark starts nothing at all. It is only written where it is missing and only removed where it
    // is there, because this runs at every opening of the tray menu as well as at every start.
    //
    // **And it is never written over an entry that is what a logon runs.** Windows starts no `Run`
    // value whose file carries the mark, so a mark put on here while the value is still the
    // mechanism would be a logon that starts nothing — and this pass cannot replace the value with
    // the task, because that is a write an unelevated process may not make. The task's own write
    // puts the mark on once the value is gone (see `enable_the_task`); a machine left as it is starts
    // the app unelevated, which is the state it was already in.
    if let Ok(exe) = env::current_exe() {
        let marked = file_asks_for_elevation(&exe);
        let entry_is_the_mechanism = before.started() == Started::Entry;

        match (wants_elevation, marked) {
            (true, false) if !entry_is_the_mechanism => {
                let _ = write_compat_elevation(&exe);
            }
            (false, true) => {
                let _ = remove_compat_elevation(&exe);
            }
            _ => {}
        }
    }

    // What the machine holds, and what is asked of it: an entry that starts this copy is settled
    // for the ordinary answer, a task that starts this copy at the highest run level is settled for
    // the elevated one, and anything else — nothing there, a value or a task Windows has been told
    // not to run, one naming a copy that has since moved — is what the write below is for.
    let settled = !wanted
        || match (before.started(), wants_elevation) {
            // A task beside the entry is the app starting twice at a logon — once unelevated from
            // the value and once elevated from the task — so the entry is settled only on its own.
            (Started::Entry, false) => !before.task_present,
            (Started::Task, true) => true,
            // A machine carrying both starts the app twice, which is the one state nothing is
            // written over: the answer that leaves it is the one the task is for, so the two are
            // told apart by whether an elevated start is asked for.
            _ => false,
        };

    let outcome = match settled {
        true => Outcome::Settled,
        // A task that is there, names this copy and is switched off is registered again rather than
        // written from nothing: the same call, and the same reason the `Run` value's own flag is
        // written by `enable_startup` (the register carries `SetEnabled(true)`).
        false => enable_startup(),
    };

    // **A write this token could not make is not a machine that refused one.** An unelevated process
    // asking for an elevated start is answered `NeedsElevation`, and the file is left saying what the
    // user asked for: bringing it into step with the machine there would erase the choice every time
    // the row was read on a copy that cannot act on it yet, and the row itself already says the
    // machine will not do it (see the label in `menus`).
    let written = outcome.settled() && !settled;

    if !matches!(outcome, Outcome::NeedsElevation) {
        // What the machine holds, written back where the file disagrees with it — which is the one
        // direction this never argues: `true` asks for the entry (and has just been given it), and
        // anything the machine refused is a `false` in the file rather than a lie in it.
        let held = machine();
        let running = held.started().runs();
        // The elevated row is in effect where a logon starts the app as an administrator **or**
        // where the file asks for one on every launch: the two are one row with two roads, and a
        // file that says yes is not written back as a no because only one of them is in place.
        let elevated = held.elevates() || held.compat_elevation();

        if let Ok(mut config) = crate::CONFIG.lock() {
            if config.run_at_startup != running || config.run_as_administrator != elevated {
                config.run_at_startup = running;
                config.run_as_administrator = elevated;
                config.save();
            }
        }
    }

    written
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_entry_is_read_as_the_executable_it_names() {
        // The path this app writes.
        assert_eq!(
            executable_of(r"C:\apps\rust-hover-preview\rust-hover-preview.exe"),
            r"C:\apps\rust-hover-preview\rust-hover-preview.exe"
        );
        // A path holding a space is quoted, and the quotes are not part of the name.
        assert_eq!(
            executable_of(r#""C:\Program Files\Rust Hover Preview\app.exe""#),
            r"C:\Program Files\Rust Hover Preview\app.exe"
        );
        // Arguments follow the executable and are not another copy of the app.
        assert_eq!(
            executable_of(r#""C:\apps\app.exe" --silent"#),
            r"C:\apps\app.exe"
        );
        assert_eq!(
            executable_of(r"C:\apps\app.exe --silent"),
            r"C:\apps\app.exe"
        );
        // Surrounding whitespace is not part of the name either.
        assert_eq!(executable_of("  C:\\apps\\app.exe  "), r"C:\apps\app.exe");
        // A value naming nothing at all is read as nothing rather than as a path.
        assert_eq!(executable_of(""), "");
        assert_eq!(executable_of(r#""""#), "");
    }

    #[test]
    fn a_written_command_line_is_the_path_and_nothing_else() {
        // What this app writes is the path in quotes, so that a folder with a space in it is one
        // path to `CreateProcess` — and it is read back as the same one path.
        let written = command_line(Path::new(r"C:\Program Files\Rust Hover Preview\app.exe"));

        assert_eq!(written, r#""C:\Program Files\Rust Hover Preview\app.exe""#);
        assert_eq!(
            executable_of(&written),
            r"C:\Program Files\Rust Hover Preview\app.exe"
        );
    }

    /// The mark is one word of a list a program may carry, and only that word goes: a compatibility
    /// mode, a DPI fix or a `HIGHDPIAWARE` beside it is the user's own setting and is left exactly
    /// as it is — and a value that was nothing but the mark leaves no empty string behind, because an
    /// empty layer is a request Windows does not understand.
    #[test]
    fn only_the_one_word_is_taken_off_a_layer_value() {
        // A value that was nothing but the mark is no value at all: it is removed rather than left
        // holding the list's own separator, which names no layer.
        assert_eq!(layers_without_the_mark("RUNASADMIN"), None);
        assert_eq!(layers_without_the_mark("~ RUNASADMIN"), None);
        assert_eq!(layers_without_the_mark("RUNASADMIN ~"), None);

        // A layer the user set beside it is theirs, and it stays — the mark is the one word taken.
        assert_eq!(
            layers_without_the_mark("~ RUNASADMIN HIGHDPIAWARE").as_deref(),
            Some("~ HIGHDPIAWARE")
        );
        assert_eq!(
            layers_without_the_mark("~ WIN8RTM runasadmin").as_deref(),
            Some("~ WIN8RTM")
        );
        // And a value that does not name the mark is handed back as it was.
        assert_eq!(
            layers_without_the_mark("~ WIN8RTM").as_deref(),
            Some("~ WIN8RTM")
        );
        // A word that only begins with it is not it.
        assert_eq!(
            layers_without_the_mark("~ RUNASADMINISTRATOR").as_deref(),
            Some("~ RUNASADMINISTRATOR"),
            "the word is compared whole"
        );
    }

    /// The mark is put on a value the way Windows' own properties tab writes it: the list's separator,
    /// the layers the file already carried, and the mark — and a value made from nothing is the
    /// separator and the mark alone.
    #[test]
    fn the_mark_is_added_to_the_layers_a_file_already_carries() {
        assert_eq!(layers_with_the_mark(None), "~ RUNASADMIN");
        assert_eq!(layers_with_the_mark(Some("")), "~ RUNASADMIN");
        assert_eq!(layers_with_the_mark(Some("~")), "~ RUNASADMIN");
        assert_eq!(
            layers_with_the_mark(Some("~ WIN8RTM")),
            "~ WIN8RTM RUNASADMIN"
        );
        // A value that already names it is not one that names it twice.
        assert_eq!(
            layers_with_the_mark(Some("~ runasadmin HIGHDPIAWARE")),
            "~ HIGHDPIAWARE RUNASADMIN"
        );
    }

    /// What the row puts on is what the row takes off: the mark is added and removed without a layer
    /// the user set on the file being lost on the way, which is the whole of what the two writes
    /// promise.
    #[test]
    fn the_mark_goes_on_and_comes_off_without_taking_the_users_layers_with_it() {
        for layers in [
            None,
            Some(""),
            Some("~"),
            Some("~ WIN8RTM"),
            Some("HIGHDPIAWARE"),
        ] {
            let marked = layers_with_the_mark(layers);
            assert!(
                marked
                    .split_whitespace()
                    .any(|word| word.eq_ignore_ascii_case("RUNASADMIN")),
                "the mark is in the value it wrote: {marked}"
            );

            let bare = layers
                .unwrap_or_default()
                .split_whitespace()
                .filter(|word| *word != "~")
                .collect::<Vec<_>>()
                .join(" ");

            assert_eq!(
                layers_without_the_mark(&marked),
                (!bare.is_empty()).then(|| format!("~ {bare}")),
                "the layers it started with, and nothing else: {marked}"
            );
        }
    }

    #[test]
    fn the_administrator_mark_on_the_executable_is_read_from_its_own_layer_value() {
        // The compatibility flags are words, and `RUNASADMIN` among them is the mark that keeps a
        // `Run` entry from starting anything: Windows never elevates one. The word is found in any
        // position and in any case, and a layer value that does not name it is not the mark.
        let asks = |value: &str| {
            value
                .split_whitespace()
                .any(|word| word.eq_ignore_ascii_case("RUNASADMIN"))
        };

        assert!(asks("~ RUNASADMIN"));
        assert!(asks("~ WIN8RTM runasadmin"));
        assert!(!asks("~ WIN8RTM"));
        assert!(!asks(""));
        assert!(!asks("~ RUNASADMINISTRATOR"));
    }

    #[test]
    fn one_path_in_another_spelling_is_this_path() {
        let running = Path::new(r"C:\apps\rust-hover-preview\rust-hover-preview.exe");

        // The same path, written as this app writes it.
        assert!(same_path(
            r"C:\apps\rust-hover-preview\rust-hover-preview.exe",
            running
        ));
        // A Windows path carries no case.
        assert!(same_path(
            r"C:\Apps\Rust-Hover-Preview\Rust-Hover-Preview.EXE",
            running
        ));
        // The verbatim prefix names the same file as without it, and is what the Shell
        // canonicalizes to, so an entry holding one is not a different copy.
        assert!(same_path(
            r"\\?\C:\apps\rust-hover-preview\rust-hover-preview.exe",
            running
        ));
        // Whitespace around the value is not part of the path.
        assert!(same_path(
            "  C:\\apps\\rust-hover-preview\\rust-hover-preview.exe  ",
            running
        ));
    }

    #[test]
    fn another_copy_of_the_app_is_not_this_copy() {
        let running = Path::new(r"C:\apps\rust-hover-preview\rust-hover-preview.exe");

        // The stray portable copy: another folder, another version, the thing the check
        // exists for.
        assert!(!same_path(r"C:\downloads\rust-hover-preview.exe", running));
        // A path that no longer resolves at all is not this one either, and is not asked
        // about: it is exactly what has to be replaced.
        assert!(!same_path(
            r"C:\apps\rust-hover-preview-0.3.6\app.exe",
            running
        ));
        // An empty value names nothing.
        assert!(!same_path("", running));
    }

    #[test]
    fn a_counted_string_is_read_up_to_its_terminator() {
        let counted = |text: &str| -> Vec<u8> {
            text.encode_utf16()
                .chain(std::iter::once(0))
                .flat_map(u16::to_le_bytes)
                .collect()
        };

        assert_eq!(
            decode_string(&counted(r"C:\apps\app.exe")),
            r"C:\apps\app.exe"
        );
        // Whatever is in the buffer behind the terminator is not part of the string.
        assert_eq!(decode_string(&counted("app.exe")), "app.exe");
        // A string that is nothing but its terminator names nothing.
        assert_eq!(decode_string(&counted("")), "");
        // An odd trailing byte is a byte this reads nothing out of.
        assert_eq!(decode_string(&[0x41, 0x00, 0x42]), "A");
    }

    #[test]
    fn the_flag_windows_keeps_beside_an_entry_is_read_the_way_the_shell_writes_it() {
        // The shapes Windows 10 and 11 write: an entry that runs is even, one the user turned off in
        // Task Manager is odd — and the flag this app writes is the first of those.
        let on = |byte: u8| byte.is_multiple_of(2);

        assert!(on(APPROVED_ENABLED[0]));
        assert!(on(2));
        assert!(on(6));
        assert!(!on(3));
        assert!(!on(1));
    }
}

// ------------------------------------------------------------- the elevated helper

/// The switch that says this launch is an elevated copy of the tray app's that has one write to
/// make and nothing else to do: it is begun by [`raise_to_elevation`], registers or removes the
/// scheduled task, and exits. Never a tray, never a viewer (see `app::detach::startup_mode`).
pub(crate) const ELEVATE_SWITCH: &str = "--apply-elevated-start";

/// What the helper is told to do, as the word it is given after the switch.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Applied {
    /// Register the task for the copy of the app the helper itself is running from.
    TurnOn,
    /// Remove the task, and the `Run` value beside it.
    TurnOff,
}

impl Applied {
    /// The word the helper is told, and the one it reads back.
    pub(crate) fn word(self) -> &'static str {
        match self {
            Applied::TurnOn => "on",
            Applied::TurnOff => "off",
        }
    }

    /// The request a word names, or nothing where it names none.
    pub(crate) fn of(word: &str) -> Option<Self> {
        match word {
            "on" => Some(Applied::TurnOn),
            "off" => Some(Applied::TurnOff),
            _ => None,
        }
    }
}

/// Start the scheduled task now, which is how this app comes back **elevated** after the row was
/// turned on: the very record that starts it at a logon is the one that starts it here, so what a
/// user gets is what they will get at the next logon rather than a second way of starting it.
///
/// It is the elevated copy's own answer — the task raises a process whose token is the highest the
/// account has — and the caller is an unelevated one that ends itself once it has been started.
pub(crate) fn start_the_task() -> bool {
    task::run_now().is_ok()
}

/// Start the copy of this app that the ordinary `Run` entry names, through the Shell, and leave the
/// starting of it to Windows' own startup path.
///
/// It is the mirror of [`start_the_task`], and it exists for the same reason: a turn-off that left the
/// tray standing as an administrator would be a tray a user has to restart by hand before the row
/// they just unchecked means anything. What is done is what a logon does — the value is run — except
/// that the value is run *unelevated*: `ShellExecuteW` through the Shell's own token, which is the one
/// token that can start a program without raising a prompt (see `app::hand_off`, and the same road a
/// launch from the Start menu takes).
pub(crate) fn start_through_the_entry() -> bool {
    let Ok(exe) = env::current_exe() else {
        return false;
    };

    // **The desktop's token, whatever the hand-off setting says.** `Elevate Hand-Offs` is about the
    // programs this app opens for the user; a copy of this app replacing this one is the row being
    // turned off, and the token is the whole of what that is about — a restart that came back
    // elevated would leave the tray standing as the state the row has just stopped asking for.
    crate::app::hand_off::start_a_copy(&crate::paths::plain_path(&exe))
}

/// What this app knows about the machine's task, for a reader that has to wait on it.
pub(crate) use task::facts as task_facts;

/// Take the *Run this program as an administrator* mark off the copy that is running, for the row
/// that is turning the elevated start back off.
///
/// It is a write to the user's own compatibility settings rather than to the startup entry, which is
/// why it is a call of its own: a mark that was set on the file's **Properties → Compatibility** tab
/// is the same request as this app's row, and unchecking the row has to be able to answer it — or the
/// machine is left asking for elevation with nothing that starts it, where a `Run` value does nothing.
pub(crate) fn clear_compat_elevation() -> bool {
    env::current_exe().is_ok_and(|exe| remove_compat_elevation(&exe))
}

/// Put that mark on the copy that is running, for the row that is turning the elevated start on with
/// no logon start to carry it yet — and for the reconciler, which puts back a mark the file asks for
/// and the machine has lost.
pub(crate) fn set_compat_elevation() -> bool {
    env::current_exe().is_ok_and(|exe| write_compat_elevation(&exe))
}

/// Whether the copy that is running carries that mark, which the row is drawn from.
pub(crate) fn compat_elevation_is_set() -> bool {
    env::current_exe().is_ok_and(|exe| file_asks_for_elevation(&exe))
}

/// Ask Windows to run an elevated copy of this app, which makes the one write this process may not
/// make: registering or removing the scheduled task (see `app::task`).
///
/// It is `runas` through the Shell, which is the documented way a process asks for elevation, and
/// the prompt it raises is the platform's own. What is answered is the **handle of the process the
/// Shell began**, because the helper is a process of its own that says nothing back: it writes one
/// thing and exits, and that exit is the only answer it ever gives. Nothing is said here about what
/// the write did — the caller reads the machine for that (see `wait_for_the_helper`).
///
/// [`Elevated::Refused`] is answered where the person at the machine said no to the prompt, or where
/// nothing was started at all, and the caller treats both as "the row did not move".
pub(crate) fn raise_to_elevation(apply: Applied) -> Elevated {
    let Ok(exe) = env::current_exe() else {
        return Elevated::Refused;
    };

    let file = wide(&crate::paths::plain_path(&exe));
    let parameters = wide(&format!("\"{ELEVATE_SWITCH}\" {}", apply.word()));
    let verb = wide("runas");

    let mut info = SHELLEXECUTEINFOW {
        cbSize: std::mem::size_of::<SHELLEXECUTEINFOW>() as u32,
        // The process handle is the whole reason this is `ShellExecuteEx` rather than
        // `ShellExecute`: it is what the wait below is made of, and a wait that had nothing to wait
        // on would be a poll with a deadline — a tray frozen for the whole of it whenever the write
        // was refused. `NOASYNC` keeps the launch on this thread's own time rather than a
        // background one's, since this thread goes on to wait for it.
        fMask: SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC,
        lpVerb: PCWSTR(verb.as_ptr()),
        lpFile: PCWSTR(file.as_ptr()),
        lpParameters: PCWSTR(parameters.as_ptr()),
        nShow: SW_SHOWNORMAL.0,
        ..Default::default()
    };

    // Safety: the three strings are this function's own buffers living past the call, and the
    // structure is one this function owns and reads back.
    if unsafe { ShellExecuteExW(&mut info) }.is_err() {
        return Elevated::Refused;
    }

    // A Shell that began the copy without leaving a handle is a Shell that left this side nothing to
    // wait on — answered as a start without one rather than as a handle of zeroes, so that the
    // caller asks the machine instead of waiting on nothing.
    Elevated::Started((!info.hProcess.is_invalid()).then_some(info.hProcess))
}

/// What the Shell answered when it was asked for an elevated copy of this app.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Elevated {
    /// The copy was begun, with the handle of its process where the Shell left one to wait on.
    Started(Option<HANDLE>),
    /// The person at the machine answered no, or the Shell began nothing at all.
    Refused,
}

/// Wait for the elevated copy to exit — the moment its one write is behind it — and close the handle
/// the Shell left.
///
/// **It is a wait rather than a poll.** The helper is one `RegisterTaskDefinition` and an exit, and
/// a poll of the machine had nothing to stop it early: a write the helper could not make was waited
/// on for the whole deadline, which is half a minute of a tray that answered no click. What is
/// waited on now is the process itself, and it is gone the moment its one write is behind it. The
/// handle is closed here because it is this module's own and there is exactly one thing to do with
/// it. A helper still running at the deadline is answered `false` and left alone: it is a process of
/// its own, and nothing here may end it.
pub(crate) fn wait_for_the_helper(helper: HANDLE) -> bool {
    /// What a helper is given to make one COM call and leave. It is a ceiling on a machine under a
    /// scanner's thumb rather than an expectation: the wait ends with the process, not with it.
    const DEADLINE: std::time::Duration = std::time::Duration::from_secs(30);

    // Safety: a handle `ShellExecuteEx` answered with, which nothing else holds.
    let exited =
        unsafe { WaitForSingleObject(helper, DEADLINE.as_millis() as u32) } == WAIT_OBJECT_0;

    // Safety: as above, and it is closed exactly once.
    unsafe {
        let _ = CloseHandle(helper);
    }

    exited
}

/// What an elevated helper is told to do, out of its own command line.
pub(crate) fn helper_request(
    args: impl IntoIterator<Item = std::ffi::OsString>,
) -> Option<Applied> {
    let mut args = args.into_iter();

    while let Some(arg) = args.next() {
        if arg.to_str() == Some(ELEVATE_SWITCH) {
            let word = args.next()?;
            return Applied::of(word.to_str()?);
        }
    }

    None
}

/// Do the one write this launch was begun for, and say whether the machine took it.
///
/// It is a function rather than the body of an `if` in `main` so that the reading of a task's facts
/// and the writing of one are in the same place: what a helper answers is what the tray's next
/// opening of its menu will read, and nothing here reports anything to anybody.
pub(crate) fn apply_for_the_helper(apply: Applied) -> bool {
    let Ok(exe) = env::current_exe() else {
        return false;
    };

    // **The mark goes first on the way off and last on the way on.** Off, a machine still asking for
    // elevation is one `enable_startup` would answer with a task rather than the entry this row is
    // leaving behind — `elevated_start_wanted` reads the mark — so it is cleared before the write.
    // On, the write is what puts it there, after the task and never before: a `Run` value under that
    // mark starts nothing at all (see `enable_the_task`).
    if apply == Applied::TurnOff {
        let _ = remove_compat_elevation(&exe);
    }

    // **What a logon does is the file's answer, and what this launch changes is only the mechanism.**
    // A `run_at_startup` that is still on is an ordinary entry written in place of the task rather
    // than nothing at all, and only a startup nobody wants is removed whole.
    let wanted = crate::CONFIG
        .lock()
        .map(|config| config.run_at_startup)
        .unwrap_or(false);

    let outcome = match (apply, wanted) {
        (Applied::TurnOn, _) => enable_startup(),
        // Off, the answer is the row's rather than the file's: the ordinary entry is written and
        // the task taken away whatever the mark said a moment ago. The mark has just been cleared
        // above, and a write that read it back would be asking a question this call has already
        // answered.
        (Applied::TurnOff, true) => settle_as_unelevated(),
        (Applied::TurnOff, false) => disable_startup(),
    };

    // And the mark, which is what the row says of every launch that is not a logon's: the machine's
    // own answer rather than the write's, because a registry that refused it is a row half in place.
    let marked = match apply {
        Applied::TurnOn => file_asks_for_elevation(&exe),
        Applied::TurnOff => !file_asks_for_elevation(&exe),
    };

    outcome.settled() && marked
}

/// A string as the Shell's own call wants it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}
