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
use windows::Win32::System::Registry::{
    RegCloseKey, RegCreateKeyExW, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW,
    HKEY, HKEY_CURRENT_USER, HKEY_USERS, KEY_QUERY_VALUE, KEY_READ, KEY_SET_VALUE, REG_BINARY,
    REG_CREATE_KEY_DISPOSITION, REG_OPTION_NON_VOLATILE, REG_SAM_FLAGS, REG_SZ,
};

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

    let written = write_entry(&exe_path) && enable_approved();
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

/// Write the scheduled task that starts this copy at a logon with the highest run level, and remove
/// the `Run` value the same logon would otherwise also run.
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

            if held {
                Outcome::Settled
            } else {
                Outcome::TaskRefused
            }
        }
        Err(Refused::NeedsElevation) => Outcome::NeedsElevation,
        Err(Refused::NoWrite) => Outcome::TaskRefused,
    }
}

/// Remove both mechanisms: the `Run` value and its flag, and the scheduled task.
///
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
    let Ok(exe_path) = env::current_exe() else {
        return false;
    };

    let Some(value) = compat_layers_value(&exe_path) else {
        return false;
    };

    value
        .split_whitespace()
        .any(|word| word.eq_ignore_ascii_case("RUNASADMIN"))
}

/// The compatibility layers this machine holds for one executable, as the words they are written
/// as, or nothing where there is no entry for it.
fn compat_layers_value(path: &Path) -> Option<String> {
    let mut hkey = HKEY::default();
    // Safety: a handle is written into a local, and this call closes it.
    let opened = unsafe {
        RegOpenKeyExW(
            HKEY_CURRENT_USER,
            COMPAT_LAYERS_KEY,
            0,
            KEY_QUERY_VALUE,
            &mut hkey,
        )
    };

    if opened.is_err() {
        return None;
    }

    // The value's name is the path, spelled as the compatibility tab wrote it: the exact key is
    // asked for, and a lookup that misses is not retried in another case — one that differs only in
    // case would be a second entry for one file, which the tab does not make.
    let name: Vec<u16> = path
        .as_os_str()
        .to_string_lossy()
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

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

    let layers = if sized {
        let mut buffer = vec![0u8; size as usize];
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
    } else {
        None
    };

    // Safety: the handle this call opened and nothing else holds.
    unsafe {
        let _ = RegCloseKey(hkey);
    }

    layers
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

/// Turn a scheduled task Windows has been told not to run back on, and answer whether it was.
///
/// It is the other half of `is_entry_enabled` for a task: the `Disabled` switch in Task Scheduler
/// (and in **Settings → Apps → Startup**, which lists the task beside the values) is a task that is
/// there and starts nothing, and the row has to be able to say so rather than a checkmark over a
/// logon entry that does not happen. Registering it again is what puts the switch back, since that
/// write is the same one the row makes.
fn enable_the_task_that_is_there(exe: &Path) -> bool {
    let Some(user) = elevation::desktop_user_name() else {
        return false;
    };

    task::register(exe, &user).is_ok() && task::facts().starts_this_copy(exe)
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

    // What the machine holds, and what is asked of it: an entry that starts this copy is settled
    // for the ordinary answer, a task that starts this copy at the highest run level is settled for
    // the elevated one, and anything else — nothing there, a value or a task Windows has been told
    // not to run, one naming a copy that has since moved — is what the write below is for.
    let before = machine();
    let settled = !wanted
        || match (before.started(), wants_elevation) {
            (Started::Entry, false) => true,
            (Started::Task, true) => true,
            // A machine carrying both starts the app twice, which is the one state nothing is
            // written over: the answer that leaves it is the one the task is for, so the two are
            // told apart by whether an elevated start is asked for.
            _ => false,
        };

    let outcome = match settled {
        true => Outcome::Settled,
        false => enable_startup(),
    };

    // A task that is there, names this copy and is switched off is turned back on rather than
    // registered from nothing: the same register, and the same reason the `Run` value's own flag is
    // written by `enable_startup` (see `enable_the_task_that_is_there`).
    let written = outcome.settled()
        || wants_elevation
            && before.task_present
            && !before.task_enabled
            && env::current_exe()
                .ok()
                .is_some_and(|exe| enable_the_task_that_is_there(&exe));

    // What the machine holds, written back where the file disagrees with it — which is the one
    // direction this never argues: `true` asks for the entry (and has just been given it), and
    // anything the machine refused is a `false` in the file rather than a lie in it.
    let held = machine();
    let running = held.started().runs();

    if let Ok(mut config) = crate::CONFIG.lock() {
        let elevated = held.elevates();

        if config.run_at_startup != running || config.run_as_administrator != elevated {
            config.run_at_startup = running;
            config.run_as_administrator = elevated;
            config.save();
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

/// What `ShellExecuteW` answers when the person at the machine says no to the UAC prompt.
const ELEVATION_CANCELLED: usize = 5;

/// Start the scheduled task now, which is how this app comes back **elevated** after the row was
/// turned on: the very record that starts it at a logon is the one that starts it here, so what a
/// user gets is what they will get at the next logon rather than a second way of starting it.
///
/// It is the elevated copy's own answer — the task raises a process whose token is the highest the
/// account has — and the caller is an unelevated one that ends itself once it has been started.
pub(crate) fn start_the_task() -> bool {
    task::run_now().is_ok()
}

/// What this app knows about the machine's task, for a reader that has to wait on it.
pub(crate) use task::facts as task_facts;

/// Ask Windows to run an elevated copy of this app, which makes the one write this process may not
/// make: registering or removing the scheduled task (see `app::task`).
///
/// It is `runas` through the Shell, which is the documented way a process asks for elevation, and
/// the prompt it raises is the platform's own. `true` is answered where the elevated copy was
/// started — not where it succeeded, which is the helper's own answer and is reported by the row
/// the user clicks next rather than by this call.
/// `false` is answered where the person at the machine said no, or where nothing was started at
/// all, and the caller treats both as "the row did not move".
pub(crate) fn raise_to_elevation(apply: Applied) -> bool {
    let Ok(exe) = env::current_exe() else {
        return false;
    };

    let file = wide(&crate::paths::plain_path(&exe));
    let parameters = wide(&format!("\"{ELEVATE_SWITCH}\" {}", apply.word()));
    let verb = wide("runas");

    // Safety: the three strings are this function's own buffers living past the call, and the
    // answer is a value rather than a handle this side has to close.
    let started = unsafe {
        windows::Win32::UI::Shell::ShellExecuteW(
            windows::Win32::Foundation::HWND::default(),
            PCWSTR(verb.as_ptr()),
            PCWSTR(file.as_ptr()),
            PCWSTR(parameters.as_ptr()),
            PCWSTR::null(),
            windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL,
        )
    };

    // The Shell answers a refusal with a value of five or below rather than with a null handle.
    started.0 as usize > ELEVATION_CANCELLED
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

    match apply {
        Applied::TurnOn => {
            let Some(user) = elevation::desktop_user_name() else {
                return false;
            };

            if task::register(&exe, &user).is_err() {
                return false;
            }

            // The two are alternatives: a machine carrying both starts the app twice at a logon,
            // once unelevated from the registry and once elevated from the task.
            delete_entry();
            delete_approved();

            task::facts().starts_this_copy(&exe)
        }
        Applied::TurnOff => {
            delete_entry();
            delete_approved();
            task::unregister().is_ok() || !task::facts().present
        }
    }
}

/// A string as the Shell's own call wants it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}
